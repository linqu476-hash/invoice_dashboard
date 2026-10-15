"""What the Django views did around the invoicing package, as plain functions on bytes (no Django, no Streamlit).

Files live in a temporary folder only while a build runs; results come back as bytes so the UI can keep them in the session.
"""
import io
import re
import tempfile
import uuid
from dataclasses import asdict
from datetime import date
from pathlib import Path

import openpyxl

from invoicing import engine, extract, presell, report_pdf, validation

from . import store

TEMPLATE_XLSX = Path(__file__).resolve().parent.parent / "assets" / "invoice_template.xlsx"
MAX_BYTES = validation.MAX_UPLOAD_BYTES


class InvalidMaster(Exception):
    """The master workbook cannot give correct documents; .problems lists every problem with its Excel row."""

    def __init__(self, problems):
        super().__init__("; ".join(problems))
        self.problems = problems


def _check_size(name, data):
    if not data:
        raise ValueError(f"{name}: the file is empty.")
    if len(data) > MAX_BYTES:
        raise ValueError(f"{name}: the file is larger than {MAX_BYTES // (1024 * 1024)} MB.")


# ---------------------------------------------------------------- step 1: packing plan -> master
def build_master(packing_bytes, filename, sheet=None, skip_text=""):
    """Packing plan (or Pre-sell file) -> MASTER_INVOICE.xlsx. Returns a dict with master (bytes), converted (bytes | None), and the summary."""
    _check_size(filename, packing_bytes)
    skip = {x.strip() for x in re.split(r"[,\s]+", skip_text or "") if x.strip()} or None
    with tempfile.TemporaryDirectory() as tmp:
        d = Path(tmp)
        src = d / "packing.xlsx"
        src.write_bytes(packing_bytes)
        converted = None
        wb = openpyxl.load_workbook(src, read_only=True)
        try:
            is_presell = presell.find_presell(wb, sheet or None)
        finally:
            wb.close()
        if is_presell:
            converted, _ = presell.convert_file(src, d / "PACKING_PLAN_CONVERTED.xlsx")
        rows, notes = extract.extract(converted or src, sheet or None, None, skip)
        if converted:
            notes.insert(0, "Pre-sell packing plan detected: converted to the packing plan layout before the master was built.")
        out = d / "MASTER_INVOICE.xlsx"
        extract.write_master(rows, out)
        summary = extract.summarize(rows)
        return dict(master=out.read_bytes(), converted=Path(converted).read_bytes() if converted else None,
                    rows=len(rows), qty=sum(r["qty"] for r in rows), notes=notes, summary=summary,
                    amount=round(sum(s["amount"] for s in summary), 2), skipped=sorted(skip) if skip else [],
                    low=dict(getattr(extract.extract, "last_low", {}) or {}), source=filename,
                    splits=sum(1 for s in summary if "-" in s["inv"]))


def transform_presell(src_bytes, filename, sheet=None):
    """Only rename the Pre-sell columns to the packing plan names. -> dict(report, notes, modes, converted bytes | None)."""
    _check_size(filename, src_bytes)
    with tempfile.TemporaryDirectory() as tmp:
        src, dst = Path(tmp) / "presell.xlsx", Path(tmp) / "PACKING_PLAN_CONVERTED.xlsx"
        src.write_bytes(src_bytes)
        report, notes, modes = presell.transform_file(src, dst, sheet or None)
        return dict(report=report, notes=notes, modes=modes, converted=dst.read_bytes() if report else None)


# ---------------------------------------------------------------- step 2: master -> invoices
def _validation_block(rep, totals):
    return dict(rows=rep.rows, input=dict(qty=rep.qty, amount=rep.amount, sets=rep.sets),
                issues=[asdict(i) for i in rep.warnings], differences=validation.reconcile(rep, totals))


def check_lines(result):
    """'Check before sending' as plain lines: engine warnings (the same warning on several invoices = one line) + input validation.
    -> (lines, differences)"""
    grouped = {}
    for i in result["invoices"]:
        for m in i["warnings"]:
            grouped.setdefault(m.replace(i["no"] + ":", "", 1).strip(), []).append(i["no"])
    lines = [f"{', '.join(v)}: {k}" for k, v in grouped.items()]
    v = result.get("validation") or {}
    return lines + [validation.describe(i) for i in v.get("issues", [])], list(v.get("differences", []))


def _build(master_bytes, template_bytes, merge, hide_content):
    with tempfile.TemporaryDirectory() as tmp:
        d = Path(tmp)
        master = d / "master.xlsx"
        master.write_bytes(master_bytes)
        rep = validation.validate_master(master)        # stop BEFORE building when the workbook cannot give correct documents
        if rep.errors:
            raise InvalidMaster([validation.describe(i) for i in rep.errors])
        if template_bytes:
            template = d / "template.xlsx"
            template.write_bytes(template_bytes)
        else:
            template = TEMPLATE_XLSX
        out = d / "output"
        result = engine.process(master, template, None, None, out, store.consignee_table(),
                                hide_comp=(True if hide_content else None), merge=merge, marks=store.marks_table())
        result["validation"] = _validation_block(rep, result["totals"])
        files = {p.name: p.read_bytes() for p in sorted(out.iterdir()) if p.is_file()}
        return result, files, rep


def build_invoices(master_bytes, master_name, template_bytes=None, merge=True, hide_content=False):
    """Validate the master, build every invoice. Returns the 'run' dict kept in the session:
    result (engine output + validation), files {name: bytes}, report_pdf (bytes), validation_csv (text), and what is needed to build again."""
    _check_size(master_name, master_bytes)
    if template_bytes:
        _check_size("template", template_bytes)
    result, files, rep = _build(master_bytes, template_bytes, merge, hide_content)
    return _finish(dict(master_bytes=master_bytes, template_bytes=template_bytes, merge=merge, hide_content=hide_content,
                        run_id=uuid.uuid4().hex), result, files, master_name, built=None)


def rebuild(run):
    """Build the invoices of a finished run again with the consignees / shipping marks saved NOW."""
    result, files, _ = _build(run["master_bytes"], run["template_bytes"], run["merge"], run["hide_content"])
    return _finish(dict(run), result, files, run["result"]["master_name"], built=date.today().isoformat())


def _finish(run, result, files, master_name, built):
    result.update(master_name=master_name, created=run.get("result", {}).get("created") or date.today().isoformat(), rebuilt=built)
    checks, diffs = check_lines(result)
    buf = io.BytesIO()
    report_pdf.build(result, run["run_id"], buf, check_lines=checks, diff_lines=diffs)
    run.update(result=result, files=files, report_pdf=buf.getvalue(),
               validation_csv=validation.report_csv(result["validation"]["issues"]))
    return run


# ---------------------------------------------------------------- consignee checks on a finished run
def missing_consignees(result):
    """Customer numbers of the invoices of a run that have NO consignee on file NOW (the list shrinks as consignees are added)."""
    table = store.consignee_table()
    groups = {}
    for i in result["invoices"]:
        cus = str(i["cus"])
        cus = cus[:-2] if cus.endswith(".0") else cus
        if engine.find_consignee(table, i["country"], cus) is None:
            g = groups.setdefault(engine.cus_key(cus), dict(customer_no=cus, invoices=[]))
            g["invoices"].append(i["no"])
    return sorted((dict(g, invoices=", ".join(g["invoices"])) for g in groups.values()), key=lambda m: m["customer_no"])


def fixable(result):
    """Invoices that printed the yellow placeholder, but whose customer number has a consignee NOW: they only need to be built again."""
    table = store.consignee_table()
    return [i["no"] for i in result["invoices"]
            if not i.get("consignee_ok", True) and engine.find_consignee(table, i["country"], str(i["cus"])) is not None]


def gauge(lines, limit):
    return report_pdf.gauge_state(lines or 0, limit)
