"""Validation center (V2 / P2). Pure Python (openpyxl only), so it also works from the command line and in plain unit tests.

Rules:
- Validation never changes financial data. It only reports what is wrong and where.
- "error"   = the workbook cannot be built correctly; the run is stopped before anything is generated.
- "warning" = the run is built, but a person should look at it (shown on the result page and in the validation report).
"""
import csv
import io
import os
import zipfile
from dataclasses import dataclass, field, asdict
from decimal import Decimal, InvalidOperation

import openpyxl

MAX_UPLOAD_BYTES = 20 * 1024 * 1024
MASTER_REQUIRED = ["INV NO", "PO NO", "CUS NO", "COUNTRY", "DESCRIPTION", "HTS CODE", "AMOUNT", "QTY", "N.W", "G.W"]
MAX_PER_CODE = 8                       # list at most this many rows per problem, then "and N more"
AMOUNT_TOLERANCE = Decimal("0.05")     # engine puts rounding differences on the last PO line, so totals must still agree

XLSX_MAGIC = b"PK\x03\x04"
XLS_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


@dataclass
class Issue:
    level: str                         # "error" | "warning"
    code: str
    message: str
    rows: list = field(default_factory=list)


@dataclass
class Report:
    issues: list = field(default_factory=list)
    rows: int = 0
    qty: int = 0
    amount: float = 0.0
    sets: bool = False

    @property
    def errors(self):
        return [i for i in self.issues if i.level == "error"]

    @property
    def warnings(self):
        return [i for i in self.issues if i.level == "warning"]

    def add(self, level, code, message, rows=None):
        self.issues.append(Issue(level, code, message, rows or []))

    def as_dict(self):
        return dict(rows=self.rows, qty=self.qty, amount=self.amount, issues=[asdict(i) for i in self.issues])


def describe(issue):
    """One readable line, e.g. 'Rows 3, 5, 9: QTY is not a whole number.' Accepts an Issue or its dict form."""
    i = issue if isinstance(issue, dict) else asdict(issue)
    if not i["rows"]:
        return i["message"]
    shown = ", ".join(str(r) for r in i["rows"][:MAX_PER_CODE])
    more = f" and {len(i['rows']) - MAX_PER_CODE} more" if len(i["rows"]) > MAX_PER_CODE else ""
    return f"Row{'s' if len(i['rows']) > 1 else ''} {shown}{more}: {i['message']}"


# ---- file level --------------------------------------------------------------------------
def check_file(name, size, head, allowed=(".xlsx",), max_bytes=MAX_UPLOAD_BYTES):
    """Extension, size and real file type (not just the name). head = first bytes of the file. Returns a list of error texts."""
    low = (name or "").lower()
    if not low.endswith(tuple(allowed)):
        return [f"Please upload a {' / '.join(allowed)} file."]
    if size == 0:
        return ["The file is empty."]
    if size > max_bytes:
        return [f"The file is {size / 1024 / 1024:.1f} MB; the limit is {max_bytes // 1024 // 1024} MB."]
    if low.endswith(".xlsx") and not head.startswith(XLSX_MAGIC):
        return ["This is not a real .xlsx workbook (it may be a renamed CSV or a corrupted download)."]
    if low.endswith(".xls") and not head.startswith(XLS_MAGIC):
        return ["This is not a real .xls workbook (it may be a renamed CSV or a corrupted download)."]
    return []


def check_xlsx_path(path):
    """Same check for a file already on disk (zip structure must contain a workbook)."""
    p = str(path)
    try:
        with open(p, "rb") as f:
            head = f.read(8)
        errs = check_file(os.path.basename(p), os.path.getsize(p), head)
        if errs:
            return errs
        with zipfile.ZipFile(p) as z:
            if "xl/workbook.xml" not in z.namelist():
                return ["The file is a zip archive, but not an Excel workbook."]
    except zipfile.BadZipFile:
        return ["The workbook is damaged and cannot be opened."]
    except OSError as e:
        return [f"The file cannot be read: {e}"]
    return []


# ---- master invoice rules ----------------------------------------------------------------
def _num(v):
    """Decimal of a cell, or None when it is empty / not a number. Booleans and text like '12 pcs' are not numbers."""
    if v is None or isinstance(v, bool) or (isinstance(v, str) and not v.strip()):
        return None
    try:
        return Decimal(str(v).strip().replace(",", ""))
    except InvalidOperation:
        return None


def _bucket(report, store, code, level, message, row):
    if code not in store:
        store[code] = Issue(level, code, message, [])
        report.issues.append(store[code])
    store[code].rows.append(row)


def validate_master(path):
    """Check a MASTER_INVOICE.xlsx before the engine reads it. Rows are Excel row numbers (header = row 1)."""
    rep = Report()
    bad = check_xlsx_path(path)
    if bad:
        for b in bad:
            rep.add("error", "file", b)
        return rep
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        try:
            iterator = wb.active.iter_rows(values_only=True)
            header = next(iterator, None)
            if header is None:
                rep.add("error", "empty", "The first sheet is empty.")
                return rep
            hdr = [str(c).strip().upper() if c else "" for c in header]
            ix = {h: i for i, h in enumerate(hdr) if h}
            miss = [n for n in MASTER_REQUIRED if n not in ix]
            if miss:
                rep.add("error", "headers", "Missing column(s) in row 1: " + ", ".join(miss) + ". Make the master on the 'Packing plan to master' page.")
                return rep

            store, seen = {}, {}
            sums_qty, sums_amt = 0, Decimal(0)
            for n, row in enumerate(iterator, start=2):
                cell = lambda k: row[ix[k]] if ix[k] < len(row) else None
                if not cell("INV NO"):
                    continue  # the engine skips rows without an invoice number as well
                rep.rows += 1
                qty, amt = _num(cell("QTY")), _num(cell("AMOUNT"))
                if qty is None:
                    _bucket(rep, store, "qty_missing", "error", "QTY is empty or not a number.", n)
                elif qty != qty.to_integral_value():
                    _bucket(rep, store, "qty_fraction", "error", "QTY is not a whole number.", n)
                else:
                    sums_qty += int(qty)
                    if qty <= 0:
                        _bucket(rep, store, "qty_not_positive", "warning", "QTY is zero or negative.", n)
                if amt is None:
                    _bucket(rep, store, "amount_missing", "error", "AMOUNT is empty or not a number.", n)
                else:
                    sums_amt += amt
                    if amt < 0:
                        _bucket(rep, store, "amount_negative", "warning", "AMOUNT is negative.", n)
                if not str(cell("PO NO") or "").strip():
                    _bucket(rep, store, "po_blank", "warning", "PO NO is blank.", n)
                if cell("CUS NO") in (None, "") or not str(cell("CUS NO")).strip():
                    _bucket(rep, store, "cus_blank", "warning", "CUS NO is blank, so no consignee or shipping marks can be matched.", n)
                if not str(cell("COUNTRY") or "").strip():
                    _bucket(rep, store, "country_blank", "warning", "COUNTRY is blank.", n)
                if not str(cell("HTS CODE") or "").strip():
                    _bucket(rep, store, "hts_blank", "warning", "HTS CODE is blank.", n)
                if "SET" in str(cell("DESCRIPTION") or "").upper():
                    rep.sets = True
                key = tuple(str(cell(k)).strip().upper() for k in ("INV NO", "PO NO", "DESCRIPTION", "HTS CODE", "QTY", "AMOUNT"))
                if key in seen:
                    _bucket(rep, store, "duplicate", "warning", f"identical to row {seen[key]} (same invoice, PO, description, HTS code, QTY and AMOUNT).", n)
                else:
                    seen[key] = n
            if rep.rows == 0:
                rep.add("error", "no_rows", "No data rows found: every row has an empty INV NO.")
            rep.qty, rep.amount = sums_qty, float(sums_amt)
            return rep
        finally:
            wb.close()  # read-only workbook is closed on success and failure
    except Exception as e:
        rep.add("error", "unreadable", f"The workbook cannot be read: {e}")
        return rep


# ---- input vs output ---------------------------------------------------------------------
def reconcile(report, totals):
    """Compare the totals of the master with the totals of the generated invoices. Returns a list of warning texts (empty = they agree).
    The piece count is only compared when the master has no SET lines (sets are split into several invoice lines)."""
    out = []
    diff = abs(Decimal(str(round(report.amount, 2))) - Decimal(str(totals.get("amount", 0))))
    if diff > AMOUNT_TOLERANCE:
        out.append(f"Total amount differs: master {report.amount:,.2f} vs invoices {float(totals.get('amount', 0)):,.2f}.")
    if not report.sets and report.qty != totals.get("qty", 0):
        out.append(f"Total pieces differ: master {report.qty:,} vs invoices {totals.get('qty', 0):,}.")
    return out


def report_csv(rows_of_issues):
    """Validation report as CSV text (level, problem, rows) for download."""
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["level", "code", "problem", "rows"])
    for i in rows_of_issues:
        i = i if isinstance(i, dict) else asdict(i)
        w.writerow([i["level"], i["code"], i["message"], " ".join(str(r) for r in i["rows"])])
    return buf.getvalue()
