#!/usr/bin/env python3
"""
generate_invoices.py  --  MASTER_INVOICE.xlsx  ->  one commercial-invoice .xlsx per INV NO

Rules implemented (from README.MD):
  * header / "via country" line / QUANTITY/PCS column title are taken from the template
  * per invoice:  TOTAL : n CARTONS, CUSTOMER NO., then product groups
  * group key = description (gender + garment) + HTS code + composition  -> same key is merged
  * each group shows 4 lines: TOTAL / CUSTOMER NO. (once) , "<NAME>  (HTS.CODE : x)", composition,
    then the PO.NO. / CUSTOMER PO NO. / ART.NO. header (once, after the first group), then PO lines
  * file name = invoice number (SC260458, SC260458-1, SC260458-2 ...)

Usage:
  python generate_invoices.py --master MASTER_INVOICE.xlsx --template Sample.xlsx \
      --packing Packing_Plan.xlsx --out invoices --date "OCTOBER 19 , 2026"
"""
import argparse, copy, re, sys, logging
from collections import OrderedDict
from decimal import Decimal, ROUND_HALF_UP
import openpyxl
from openpyxl.styles import PatternFill, Alignment

log = logging.getLogger("inv")
YELLOW = PatternFill("solid", fgColor="FFFF00")

# ---------------------------------------------------------------- lookup tables (edit here)
COUNTRY_LABEL = {"UNITED KINGDOM": "UK"}          # text printed after "VIA THAILAND TO"
CONSIGNEES = {                                     # per destination country
    "UNITED KINGDOM": dict(
        name="ADIDAS INTERNATIONAL TRADING AG",
        address=None,        # None -> keep the address already in the template (UK)
        marks=None),
    # add GERMANY / UNITED STATES / ... here once the real addresses are known
}
PLACEHOLDER = "<<{what} - {country}: TO BE CONFIRMED>>"
# ----------------------------------------------------------------------------------------

def r2(x):
    """Weights are rounded to the nearest 0.1 KG (half-up) and shown with 2 decimals: 5.39 -> 5.40, 6.43 -> 6.40."""
    d = Decimal(str(round(float(x), 6))).quantize(Decimal("0.1"), ROUND_HALF_UP)
    return d.quantize(Decimal("0.01"))

def split_desc(desc):
    desc = " ".join(str(desc).split())
    m = re.search(r"\d+%", desc)
    if not m:
        return desc, ""
    return desc[:m.start()].strip(), desc[m.start():].strip()

def fmt_date(v):
    """date / datetime / text from the master -> 'OCTOBER 19 , 2026' (the format printed on the invoice)."""
    if hasattr(v, "strftime"):
        return v.strftime("%B %d , %Y").upper()
    return " ".join(str(v).split()).upper() if v else None

def group_rows(rows, hide_comp=False):
    """Group PO lines. Normal: description (gender + garment) + HTS + composition. hide_comp: description + HTS only,
    so e.g. UNISEX JACKET / 6202401091 becomes ONE group whatever the composition was."""
    groups = OrderedDict()
    for r in rows:
        name, comp = split_desc(r["desc"])
        key = (name, r["hts"]) if hide_comp else (name, r["hts"], comp)
        groups.setdefault(key, []).append(r)
    return groups

def hts10(v):
    s = str(v).strip()
    return s if "+" in s else s[:10]

def read_master(path):
    ws = openpyxl.load_workbook(path, data_only=True).active
    hdr = [str(c.value).strip().upper() if c.value else "" for c in ws[1]]
    ix = {h: i for i, h in enumerate(hdr)}
    need = ["INV NO", "PO NO", "CUS NO", "COUNTRY", "DESCRIPTION", "HTS CODE", "AMOUNT", "QTY", "N.W", "G.W"]
    miss = [n for n in need if n not in ix]
    if miss:
        raise ValueError(f"MASTER_INVOICE is missing columns: {miss}")
    invs = OrderedDict()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[ix["INV NO"]]:
            continue
        g = lambda k: row[ix[k]]
        inv = invs.setdefault(str(g("INV NO")).strip(), dict(rows=[], nw=None, gw=None, cus=set(), country=set(), dates=[]))
        opt = lambda k: (str(row[ix[k]]).strip() if k in ix and row[ix[k]] not in (None, "") else None)
        inv["rows"].append(dict(po=str(g("PO NO")), desc=g("DESCRIPTION"), hts=hts10(g("HTS CODE")),
                                amt=float(g("AMOUNT")), qty=int(g("QTY")),
                                cpo=opt("CUSTOMER PO NO"), art=opt("ART NO")))
        d = fmt_date(row[ix["INV DATE"]]) if "INV DATE" in ix else None
        if d and d not in inv["dates"]: inv["dates"].append(d)
        inv["cus"].add(g("CUS NO")); inv["country"].add(str(g("COUNTRY")).strip().upper())
        if g("N.W") is not None: inv["nw"] = g("N.W")
        if g("G.W") is not None: inv["gw"] = g("G.W")
    return invs

def read_packing(path):
    """PO -> article no, customer number, CRD date (optional enrichment; master has no ART.NO)."""
    if not path:
        return {}, None
    ws = openpyxl.load_workbook(path, data_only=False).worksheets[0]
    out, crd = {}, None
    for r in range(2, ws.max_row + 1):
        po = ws.cell(r, 9).value
        if po:
            out[str(po)] = dict(art=ws.cell(r, 29).value, cus=ws.cell(r, 20).value,
                                price=ws.cell(r, 38).value, qty=ws.cell(r, 37).value, adj=ws.cell(r, 41).value)
            crd = crd or ws.cell(r, 7).value
    return out, crd

def snap(ws, r):
    return dict(styles={c: copy.copy(ws.cell(r, c)._style) for c in range(1, 27)},
                merges=[(m.min_col, m.max_col) for m in ws.merged_cells.ranges if m.min_row == r == m.max_row],
                h=ws.row_dimensions[r].height)

def apply(ws, r, t):
    for c in range(1, 27):
        ws.cell(r, c).value = None
        ws.cell(r, c)._style = copy.copy(t["styles"][c])
    for a, b in t["merges"]:
        if b > a:
            ws.merge_cells(start_row=r, end_row=r, start_column=a, end_column=b)
    ws.row_dimensions[r].height = t["h"]

def allocate(total, qtys):
    """Split `total` (0.01 steps) over the PO lines in proportion to qty; the parts add up exactly to `total`."""
    cents = int((total * 100).to_integral_value())
    s = sum(qtys)
    base = [cents * q // s for q in qtys]
    order = sorted(range(len(qtys)), key=lambda i: (-(cents * qtys[i] % s), i))   # largest remainder first
    for i in order[:cents - sum(base)]:
        base[i] += 1
    return [Decimal(b) / 100 for b in base]

def add_packing_list(wb, ws, inv_no, po_rows, tot, inv, title=None):
    """Copy the invoice sheet and turn the copy into the PACKING LIST (PRICE/USD -> N.W, AMOUNT -> G.W)."""
    pl = wb.copy_worksheet(ws)
    pl.title = title or ("PL-" + inv_no)[:31]
    pl.sheet_view.showGridLines = ws.sheet_view.showGridLines
    if ws.print_area:
        pl.print_area = ws.print_area.split("!")[-1].replace("$", "")
    pl["C6"] = "PACKING LIST"
    pl["U22"] = "N.W(KGS)"
    pl["X22"] = "G.W(KGS)"
    if inv["nw"] is None or inv["gw"] is None:
        for r in po_rows:
            pl.cell(r, 21).value = pl.cell(r, 24).value = None
        log.warning("%s: no N.W./G.W. in master -> packing list weights left blank", inv_no)
    else:
        qtys = [ws.cell(r, 19).value for r in po_rows]
        nws, gws = allocate(r2(inv["nw"]), qtys), allocate(r2(inv["gw"]), qtys)
        for r, n, g in zip(po_rows, nws, gws):
            pl.cell(r, 21).value = float(n)
            pl.cell(r, 24).value = float(g)
            pl.cell(r, 21).number_format = pl.cell(r, 24).number_format = "0.00"
    pl.cell(tot, 21)._style = copy.copy(pl.cell(tot, 24)._style)
    pl.cell(tot, 21).value = f"=SUM(U{po_rows[0]}:U{tot-1})"
    pl.cell(tot, 21).number_format = pl.cell(tot, 24).number_format = "0.00"
    pl.cell(tot, 21).alignment = Alignment(horizontal="center", vertical=pl.cell(tot, 21).alignment.vertical)  # N.W total under the N.W lines
    pl.cell(tot, 24).value = f"=SUM(X{po_rows[0]}:X{tot-1})"
    return pl

def build(inv_no, inv, template, pack, date_txt, outdir, wb=None, inv_title=None, pl_title=None, consignees=None, hide_comp=False):
    """Alone: saves <outdir>/<inv_no>.xlsx.  With `wb` (combined mode): adds the INV + PL sheets to that workbook."""
    shared = wb is not None
    if shared:
        tpl = wb["__tpl"]
        ws = wb.copy_worksheet(tpl)
        ws.sheet_view.showGridLines = tpl.sheet_view.showGridLines
        if tpl.print_area:
            ws.print_area = tpl.print_area.split("!")[-1].replace("$", "")
    else:
        wb = openpyxl.load_workbook(template); ws = wb.active
    country = next(iter(inv["country"]))
    # ---- groups
    groups = group_rows(inv["rows"], hide_comp)
    n_po, n_grp = len(inv["rows"]), len(groups)
    # per group: description line + composition line; with hide_comp only the FIRST group keeps its composition line
    need_last = 26 + 2 * n_po + n_grp if hide_comp else 25 + 2 * n_po + 2 * n_grp
    shift = need_last - 47                         # template table ends at row 47

    # ---- snapshots, then re-flow the footer
    T = dict(desc=snap(ws, 25), comp=snap(ws, 26), hdr=snap(ws, 27), po=snap(ws, 28), sp=snap(ws, 29))
    foot = [m for m in ws.merged_cells.ranges if m.min_row >= 48]
    foot_h = {k: v.height for k, v in ws.row_dimensions.items() if k >= 48 and v.height}
    fm = [(m.min_row, m.max_row, m.min_col, m.max_col) for m in foot]
    for m in [m for m in list(ws.merged_cells.ranges) if m.min_row >= 25]:
        ws.unmerge_cells(str(m))
    for k in foot_h: ws.row_dimensions[k].height = None
    if shift:
        ws.move_range("A48:Z90", rows=shift)
    for a, b, c, d in fm:
        ws.merge_cells(start_row=a + shift, end_row=b + shift, start_column=c, end_column=d)
    for k, h in foot_h.items(): ws.row_dimensions[k + shift].height = h

    # ---- header block
    ws["C8"] = f"INV.NO.  {inv_no}"
    inv_date = inv["dates"][0] if inv.get("dates") else date_txt     # INV DATE (= CRD) from the master first
    if inv_date is None:
        raise ValueError(f"{inv_no}: no invoice date (INV DATE is empty in the master and the packing plan has no CRD).")
    if len(inv.get("dates", [])) > 1:
        log.warning("%s has several INV DATE values %s -> the first one is used", inv_no, inv["dates"])
    ws["C9"] = f"DATE : {inv_date}"
    label = COUNTRY_LABEL.get(country, country)
    ws["C19"] = re.sub(r"TO\s+\S+\s*$", f"TO {label}", ws["C19"].value.strip())
    cus = sorted(inv["cus"], key=str)
    if len(cus) > 1: log.warning("%s has several CUS NO %s", inv_no, cus)
    ws["C24"] = f"CUSTOMER NO. {cus[0]}"
    ws["C23"] = "TOTAL : 1 CARTONS. "
    cons = (consignees if consignees is not None else CONSIGNEES).get(country)
    if cons is not None:                           # known destination: write what is on file, blank = keep template (UK)
        if cons.get("name"):
            ws["E13"] = cons["name"]
        if cons.get("address"):
            ws["E14"] = cons["address"]
        if cons.get("marks"):
            ws["B%d" % (55 + shift)] = cons["marks"]
    if cons is None:                               # unknown destination -> never reuse the UK address
        ws["E13"] = "ADIDAS INTERNATIONAL TRADING AG"
        ws["E14"] = PLACEHOLDER.format(what="CONSIGNEE ADDRESS", country=country)
        ws["B%d" % (55 + shift)] = PLACEHOLDER.format(what="SHIPPING MARKS", country=country)
        for a in ("E14", "B%d" % (55 + shift)): ws[a].fill = YELLOW
        log.warning("%s: no consignee on file for %s -> placeholders written (yellow)", inv_no, country)

    # ---- table
    r = 25; first = True; po_rows = []
    for key, items in groups.items():
        name, hts = key[0], key[1]
        apply(ws, r, T["desc"]); ws.cell(r, 3).value = f"{name}  (HTS.CODE : {hts})"; r += 1
        if first or not hide_comp:                 # first group always shows its content, the others only when not removed
            apply(ws, r, T["comp"]); ws.cell(r, 3).value = split_desc(items[0]["desc"])[1]; r += 1
        if first:
            apply(ws, r, T["hdr"])
            ws.cell(r, 3).value, ws.cell(r, 4).value, ws.cell(r, 11).value = "PO.NO.", "CUSTOMER PO NO.", "ART.NO."
            r += 1; first = False
        for it in items:
            apply(ws, r, T["po"])
            p = pack.get(it["po"], {})
            ws.cell(r, 3).value = int(it["po"]) if it["po"].isdigit() else it["po"]
            ws.cell(r, 4).value = it.get("cpo") or p.get("cus"); ws.cell(r, 11).value = it.get("art") or p.get("art")
            ws.cell(r, 19).value = it["qty"]
            ws.cell(r, 21).value = round(it["amt"] / it["qty"], 6)
            ws.cell(r, 24).value = f"=S{r}*U{r}"
            po_rows.append(r); r += 1
            apply(ws, r, T["sp"]); r += 1
    assert r - 1 == need_last, (r - 1, need_last)

    # ---- totals / weights
    tot = 50 + shift
    ws.cell(tot, 19).value = f"=SUM(S{po_rows[0]}:S{tot-1})"
    ws.cell(tot, 24).value = f"=SUM(X{po_rows[0]}:X{tot-1})"
    # qty total sits in the same column position as the qty lines above it (centered, like S28...)
    ws.cell(tot, 19).alignment = Alignment(horizontal="center", vertical=ws.cell(tot, 19).alignment.vertical)
    if inv["nw"] is None or inv["gw"] is None:
        log.warning("%s: N.W./G.W. missing in master", inv_no)
    else:
        ws.cell(56 + shift, 14).value = f"TOTAL N.W.:                 {r2(inv['nw'])}  KGS."
        ws.cell(58 + shift, 14).value = f"TOTAL G.W.:                 {r2(inv['gw'])}  KGS."
    ws.cell(60 + shift, 14).value = "TOTAL PACKED:          1  CARTONS."
    ws.title = inv_title or inv_no[:31]
    add_packing_list(wb, ws, inv_no, po_rows, tot, inv, pl_title)
    out = None
    if not shared:
        out = f"{outdir}/{inv_no}.xlsx"
        wb.save(out)
    # ---- reconciliation with master
    exp_q = sum(x["qty"] for x in inv["rows"]); exp_a = round(sum(x["amt"] for x in inv["rows"]), 2)
    for it in inv["rows"]:
        p = pack.get(it["po"])
        if p and p["price"] is not None:
            calc = round(p["qty"] * p["price"] + (p["adj"] or 0), 2)
            if abs(calc - it["amt"]) > 0.005 or p["qty"] != it["qty"]:
                log.warning("%s PO %s differs from packing plan (master %.2f / plan %.2f)", inv_no, it["po"], it["amt"], calc)
        elif pack and not it.get("art"):
            log.warning("%s PO %s not found in packing plan (ART.NO. left blank)", inv_no, it["po"])
    return out, n_po, n_grp, exp_q, exp_a

def sheet_names(inv_no, multi_base):
    """INV / PL / INV-1 / PL-1 for one invoice number; with several invoice numbers: 'INV SC260458', 'PL SC260458-1' ..."""
    if multi_base:
        return f"INV {inv_no}"[:31], f"PL {inv_no}"[:31]
    suf = inv_no.split("-", 1)[1] if "-" in inv_no else ""
    sfx = f"-{suf}" if suf else ""
    return "INV" + sfx, "PL" + sfx

def describe(inv_no, inv, pack, hide_comp=False):
    """Plain-data preview of one invoice (groups and PO lines) for the dashboard."""
    groups = group_rows(inv["rows"], hide_comp)
    out = []
    for i, (key, items) in enumerate(groups.items()):
        name, hts = key[0], key[1]
        comp = split_desc(items[0]["desc"])[1] if (i == 0 or not hide_comp) else ""
        lines = []
        for it in items:
            p = pack.get(it["po"], {})
            lines.append(dict(po=it["po"], cus_po=it.get("cpo") or p.get("cus") or "", art=it.get("art") or p.get("art") or "", qty=it["qty"],
                              price=round(it["amt"] / it["qty"], 2), amount=round(it["amt"], 2)))
        out.append(dict(name=name, hts=hts, comp=comp, lines=lines))
    return out

class _Collect(logging.Handler):
    def __init__(self):
        super().__init__(logging.WARNING); self.items = []
    def emit(self, record):
        self.items.append(record.getMessage())

def process(master, template, packing, date_txt, outdir, consignees=None, hide_comp=False):
    """Build every invoice (+ packing list) and the combined workbook. Returns a JSON-friendly result dict."""
    import os, zipfile
    os.makedirs(outdir, exist_ok=True)
    pack, crd = read_packing(packing)
    if not date_txt:                                   # only a fallback for a master without INV DATE
        date_txt = crd.strftime("%B %d , %Y").upper() if crd else None
    invs = read_master(master)
    if not invs:
        raise ValueError("MASTER_INVOICE has no data rows.")
    h = _Collect(); log.addHandler(h)
    try:
        res, files, wb_all = [], [], openpyxl.load_workbook(template)
        wb_all.active.title = "__tpl"
        multi = len({n.split("-")[0] for n in invs}) > 1
        fam_wb = {}                                   # base invoice -> ONE workbook holding INV, PL, INV-1, PL-1 ...
        for no, inv in invs.items():
            h.items.clear()
            base = no.split("-")[0]
            if base not in fam_wb:
                fam_wb[base] = openpyxl.load_workbook(template)
                fam_wb[base].active.title = "__tpl"
            _, n_po, n_grp, q, amt = build(no, inv, template, pack, date_txt, outdir, wb=fam_wb[base],
                                           inv_title=sheet_names(no, False)[0], pl_title=sheet_names(no, False)[1],
                                           consignees=consignees, hide_comp=hide_comp)
            it, pt = sheet_names(no, multi)
            build(no, inv, template, pack, date_txt, outdir, wb=wb_all, inv_title=it, pl_title=pt,
                  consignees=consignees, hide_comp=hide_comp)
            warns = list(dict.fromkeys(h.items))          # same warning comes from both builds -> keep one
            country = next(iter(inv["country"]))
            res.append(dict(
                no=no, country=country, cus=sorted(map(str, inv["cus"]))[0], n_po=n_po, n_grp=n_grp, qty=q,
                amount=amt, nw=float(r2(inv["nw"])) if inv["nw"] is not None else None,
                gw=float(r2(inv["gw"])) if inv["gw"] is not None else None,
                consignee_ok=(consignees if consignees is not None else CONSIGNEES).get(country) is not None,
                warnings=warns, groups=describe(no, inv, pack, hide_comp)))
        for base, fwb in fam_wb.items():              # one file per invoice family: SC260460.xlsx = SC260460, -1, -2 ...
            del fwb["__tpl"]; fwb.active = 0
            fwb.save(os.path.join(outdir, f"{base}.xlsx"))
            files.append(f"{base}.xlsx")
        del wb_all["__tpl"]; wb_all.active = 0
        first = next(iter(invs)).split("-")[0]
        combined = f"{first}_INV-PL.xlsx"
        wb_all.save(os.path.join(outdir, combined))
        zip_name = f"{first}_all_invoices.zip"
        with zipfile.ZipFile(os.path.join(outdir, zip_name), "w", zipfile.ZIP_DEFLATED) as z:
            for f in files + [combined]:
                z.write(os.path.join(outdir, f), f)
    finally:
        log.removeHandler(h)
    by_country = OrderedDict()
    for r in res:
        c = by_country.setdefault(r["country"], dict(country=r["country"], invoices=0, qty=0, amount=0.0))
        c["invoices"] += 1; c["qty"] += r["qty"]; c["amount"] = round(c["amount"] + r["amount"], 2)
    all_dates = list(dict.fromkeys(d for inv in invs.values() for d in (inv["dates"] or [date_txt]) if d))
    return dict(
        date=" / ".join(all_dates), hide_comp=hide_comp, invoices=res, by_country=list(by_country.values()),
        totals=dict(invoices=len(res), po_lines=sum(r["n_po"] for r in res), qty=sum(r["qty"] for r in res),
                    amount=round(sum(r["amount"] for r in res), 2), countries=len(by_country),
                    warnings=sum(len(r["warnings"]) for r in res)),
        files=dict(combined=combined, zip=zip_name, items=files))

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--master", help="MASTER_INVOICE.xlsx (omit it to build the master from --packing first)")
    ap.add_argument("--template", required=True)
    ap.add_argument("--invoices", help="with --packing and no --master: only these invoices, e.g. SC260458,SC260460")
    ap.add_argument("--packing"); ap.add_argument("--out", default="invoices"); ap.add_argument("--date", help="fallback only, used when the master has no INV DATE column")
    ap.add_argument("--combine", help="also write ONE workbook with sheets INV, PL, INV-1, PL-1 ... (give its file name)")
    ap.add_argument("--no-content", action="store_true",
                    help="keep the content line of the first group only, remove it from the next groups; groups merge by description + HTS code")
    a = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    import os; os.makedirs(a.out, exist_ok=True)
    pack, crd = read_packing(a.packing)
    date_txt = a.date or (crd.strftime("%B %d , %Y").upper() if crd else None)   # fallback only: INV DATE of the master wins
    if not a.master:
        if not a.packing: sys.exit("Give --master, or --packing to build the master from it.")
        try:
            from . import extract as ptm          # inside the Django project
        except ImportError:
            import extract as ptm                 # run alone: python dashboard/engine.py
        only = {x.strip() for x in a.invoices.split(",")} if a.invoices else None
        rows, notes = ptm.extract(a.packing, only=only)
        a.master = ptm.write_master(rows, f"{a.out}/MASTER_INVOICE.xlsx")
        log.info("Master built from packing plan: %d PO lines -> %s", len(rows), a.master)
        for n in notes: log.warning("%s", n)
    invs = read_master(a.master)
    fam = {}
    for no, inv in invs.items():                    # one workbook per invoice family (SC260460 + -1, -2, -3 ...)
        base = no.split("-")[0]
        if base not in fam:
            fam[base] = openpyxl.load_workbook(a.template); fam[base].active.title = "__tpl"
        it, pt = sheet_names(no, False)
        _, n_po, n_grp, q, amt = build(no, inv, a.template, pack, date_txt, a.out, wb=fam[base], inv_title=it, pl_title=pt, hide_comp=a.no_content)
        log.info("OK %-12s %2d PO lines, %d groups, qty %d, amount %.2f", no, n_po, n_grp, q, amt)
    for base, fwb in fam.items():
        del fwb["__tpl"]; fwb.active = 0
        fwb.save(f"{a.out}/{base}.xlsx"); log.info("FILE %s -> %s", fwb.sheetnames, f"{a.out}/{base}.xlsx")
    if a.combine:
        wb = openpyxl.load_workbook(a.template); wb.active.title = "__tpl"
        multi = len({n.split("-")[0] for n in invs}) > 1
        for no, inv in invs.items():
            it, pt = sheet_names(no, multi)
            build(no, inv, a.template, pack, date_txt, a.out, wb=wb, inv_title=it, pl_title=pt, hide_comp=a.no_content)
        del wb["__tpl"]
        wb.active = 0
        path = f"{a.out}/{a.combine}"
        wb.save(path); log.info("COMBINED %s -> %s", wb.sheetnames, path)

if __name__ == "__main__":
    main()
