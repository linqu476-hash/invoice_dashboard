#!/usr/bin/env python3
"""
presell_to_packing.py  --  "Packing Plan Pre-sell" (.xlsx)  ->  the packing plan layout that extract.py / the dashboard reads

Differences between the two templates, and what this does about them
  * 'CUSTNOCD'          is the 'Gps Customer Number'            (renamed)
  * 'HTS CODE:'         is 'HTS CODE1:'                         (renamed)
  * 'GARMENT DETAIL' / 'WEIGHT' / 'ORDER' ... differ only in capital letters (same column)
  * 'Order Type'        does not exist in the pre-sell file      (empty column added, so every later column keeps its place)
  * 'REMARK' / 'Remark' become 'REMARK2' / 'REMARK1'
  * 'import' (country of import) has no place in the packing plan -> kept after the last packing plan column
  * formulas (carton totals, grand totals) are rewritten for the new column letters, cell styles, column widths and
    hidden columns are carried over; freeze pane and filter are set like the packing plan (J2)
Columns are matched by header text, never by letter, so a moved column does not break it.
extract.py calls this automatically when it sees a pre-sell sheet; run it alone to get the converted file.

Usage
  python presell_to_packing.py "Packing_Plan_Pre-sell.xlsx" -o packing_plan_converted.xlsx
"""
import argparse
import copy
import re
import sys

import openpyxl
from openpyxl.utils import get_column_letter as L

# the packing plan layout, exactly as in packing_plan.xlsx (row 1)
TARGET = ["Order ", "Season", "For Account Of", "Seller", "Assigned Factory", "Factory", "Customer Request Date (CRD)",
          "INV.NO.", "PO Number", "PO Line Aggregator", "PO Line Item #", "PACKING PLAN", "COO", "SI", "FWD NAME",
          "Gps Customer Number", "Shipment Method", "COUNTRY", "Final Destination", "Customer Number", "Customer PO",
          "Order Type", "Working Number", "Garment detail", "DESCRIPTION", "HTS CODE1:", "HTS CODE2:", "Gender",
          "Article Number", "Article Description", "Model Number", "Model Name", "PSDD", "PODD", "First Production Date",
          "Last Production Date", "Ordered Quantity", "Price Per Unit", "Total Merchandise Amount", "Total Tax Amount",
          "Total Adjustments", "Total Document Amount", "Customer Size Run", "Technical Notation", "Customer Size",
          "Technical Size", "Ordered Quantity", "weight", "NW./KGS.", "GW./KGS.", "CARTONS SIZE", "CARTON WEIGHT",
          "CARTON NO.", "REMARK2", "REMARK1", "REMARK3", "REMARK4"]
RENAME = {"custnocd": "gps customer number", "hts code:": "hts code1:"}      # pre-sell header -> packing plan header
RENAME_SEQ = {"remark": ["remark2", "remark1", "remark3", "remark4"]}         # same header twice -> in this order
REQUIRED = ["inv.no.", "po number", "gps customer number", "country", "description", "hts code1:",
            "ordered quantity", "price per unit"]


class PresellError(ValueError):
    pass


def _n(v):
    return " ".join(str(v).split()).lower() if v is not None else ""


def headers(ws):
    return [_n(c.value) for c in ws[1]]


def is_presell(ws):
    """A packing plan sheet that still has the pre-sell headers (CUSTNOCD / HTS CODE:) instead of the packing plan ones."""
    h = set(headers(ws))
    return "inv.no." in h and "po number" in h and ("custnocd" in h or "hts code:" in h) \
        and not ("gps customer number" in h and "hts code1:" in h)


def find_presell(wb, name=None):
    sheets = [wb[name]] if name and name in wb.sheetnames else wb.worksheets
    return next((ws for ws in sheets if is_presell(ws)), None)


def column_plan(ws):
    """-> (mapping, extra) ; mapping[new_col] = old_col or None (empty column) ; extra = [(old_col, header)] kept after the layout."""
    seen, src = {}, {}                                    # (packing plan header, occurrence) -> old column
    used, seq_i = set(), {}
    for c in ws[1]:
        if c.value is None or _n(c.value) == "":
            continue
        h = _n(c.value)
        if h in RENAME_SEQ:
            i = seq_i.get(h, 0); seq_i[h] = i + 1
            h = RENAME_SEQ[h][i] if i < len(RENAME_SEQ[h]) else h
        h = RENAME.get(h, h)
        k = seen.get(h, 0); seen[h] = k + 1
        src[(h, k)] = c.column
    mapping, occ = [], {}
    for t in TARGET:
        h = _n(t); k = occ.get(h, 0); occ[h] = k + 1
        col = src.get((h, k))
        mapping.append(col)
        if col: used.add(col)
    missing = [h for h in REQUIRED if (h, 0) not in src]
    if missing:
        raise PresellError(f"Pre-sell packing plan is missing columns: {missing}")
    extra = [(c.column, str(c.value).strip()) for c in ws[1]
             if c.value is not None and _n(c.value) != "" and c.column not in used]
    return mapping, extra


REF = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def _formula(f, colmap):
    def sub(m):
        from openpyxl.utils import column_index_from_string as ci
        old = ci(m.group(2))
        new = colmap.get(old)
        return f"{m.group(1)}{L(new) if new else m.group(2)}{m.group(3)}{m.group(4)}"
    return REF.sub(sub, f)


def convert_sheet(wb, ws):
    """Builds the converted sheet inside `wb` (so cell styles can be copied), replaces `ws` by it and returns it."""
    mapping, extra = column_plan(ws)
    out = wb.create_sheet("__presell_converted")
    cols = [(i + 1, old) for i, old in enumerate(mapping)] + [(len(TARGET) + 1 + j, old) for j, (old, _) in enumerate(extra)]
    colmap = {old: new for new, old in cols if old}              # old column -> new column (for the formulas)
    for r in range(1, ws.max_row + 1):
        prev_style = None
        for new, old in cols:
            cell = out.cell(r, new)
            if old:
                s = ws.cell(r, old)
                v = s.value
                if isinstance(v, str) and v.startswith("="):
                    v = _formula(v, colmap)
                cell.value = v
                cell._style = copy.copy(s._style)
                prev_style = s._style
            elif prev_style is not None:                         # new empty column (Order Type): same look as its neighbour
                cell._style = copy.copy(prev_style)
        if r == 1:
            for new, old in cols:
                out.cell(1, new).value = TARGET[new - 1] if new <= len(TARGET) else extra[new - len(TARGET) - 1][1]
        h = ws.row_dimensions[r]
        if h.height: out.row_dimensions[r].height = h.height
        if h.hidden: out.row_dimensions[r].hidden = True
    for new, old in cols:
        d = ws.column_dimensions[L(old)] if old else None
        o = out.column_dimensions[L(new)]
        o.width = (d.width if d is not None and d.width else 9)
        if d is not None and d.hidden: o.hidden = True
    out.freeze_panes = "J2"
    out.auto_filter.ref = f"A1:{L(len(cols))}{ws.max_row}"
    out.sheet_view.zoomScale = ws.sheet_view.zoomScale
    title, idx = ws.title, wb.index(ws)
    wb.remove(ws)
    out.title = title
    wb.move_sheet(out, offset=idx - wb.index(out))
    return out, extra


# Shipment Method ("mode") that must always get a master invoice (and so an invoice): AIR-CC, also written Air-CC / AIR CC / aircc
MUST_BUILD_MODE_LABEL = "AIR-CC"


def is_air_cc(mode):
    return "".join(ch for ch in str(mode or "").lower() if ch.isalpha()) == "aircc"


def mode_summary(ws):
    """Shipment Method (mode) read from the sheet -> [dict(mode, invoices, lines, must_build)], biggest first.
    Counts PO lines (rows with INV.NO. and PO Number) and distinct invoice numbers per mode."""
    h = headers(ws)
    if "shipment method" not in h:
        return []
    mi, ii, pi = h.index("shipment method"), h.index("inv.no.") if "inv.no." in h else None, h.index("po number") if "po number" in h else None
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        inv = row[ii] if ii is not None and ii < len(row) else None
        po = row[pi] if pi is not None and pi < len(row) else None
        if not inv or not po:
            continue
        m = " ".join(str(row[mi]).split()) if mi < len(row) and row[mi] not in (None, "") else "(empty)"
        d = out.setdefault(m, dict(mode=m, invs=set(), lines=0))
        d["invs"].add(str(inv).strip()); d["lines"] += 1
    return sorted((dict(mode=d["mode"], invoices=len(d["invs"]), lines=d["lines"], must_build=is_air_cc(d["mode"]))
                   for d in out.values()), key=lambda x: -x["lines"])


def mapping_report(ws):
    """Column-by-column list for the 'Transform column names' page.
    -> [dict(pos, letter, new, old_letter, old, status)] ; status = same | case | renamed | added | kept
    (same: identical header, case: only capital letters / spaces differ, renamed: different header text,
     added: new empty column, kept: pre-sell column with no place in the packing plan, put after the last column)."""
    mapping, extra = column_plan(ws)
    rows = []
    for i, old in enumerate(mapping):
        new = TARGET[i]
        if old is None:
            rows.append(dict(pos=i + 1, letter=L(i + 1), new=new.strip(), old_letter="", old="", status="added"))
            continue
        src = str(ws.cell(1, old).value)
        status = "same" if src.strip() == new.strip() else "case" if _n(src) == _n(new) else "renamed"
        rows.append(dict(pos=i + 1, letter=L(i + 1), new=new.strip(), old_letter=L(old), old=src.strip(), status=status))
    for j, (old, h) in enumerate(extra):
        pos = len(TARGET) + 1 + j
        rows.append(dict(pos=pos, letter=L(pos), new=h, old_letter=L(old), old=h, status="kept"))
    return rows


def transform_file(src, dst, sheet=None):
    """Pre-sell .xlsx -> packing plan layout .xlsx, with the column report.
    -> (report, notes, modes). Raises PresellError when the file has no pre-sell sheet / misses required columns.
    modes = Shipment Method (mode) summary read from the pre-sell sheet (AIR-CC always gets a master invoice).
    report == [] means the file already is in the packing plan layout (nothing written)."""
    wb = openpyxl.load_workbook(src, data_only=False)
    ws = find_presell(wb, sheet)
    if ws is None:
        return [], [], []
    report = mapping_report(ws)
    modes = mode_summary(ws)
    wb, notes = convert_workbook(wb, sheet)
    wb.save(dst)
    return report, notes, modes


def convert_workbook(wb, sheet=None):
    """-> (wb, notes). notes is [] when the workbook already is in the packing plan layout."""
    ws = find_presell(wb, sheet)
    if ws is None:
        return wb, []
    new, extra = convert_sheet(wb, ws)
    notes = [f"Pre-sell packing plan detected: converted to the packing plan layout (CUSTNOCD -> Gps Customer Number, "
             f"HTS CODE: -> HTS CODE1:, Order Type column added)."]
    if extra:
        notes.append("Kept after the last packing plan column: " + ", ".join(h for _, h in extra) + ".")
    return wb, notes


def convert_file(src, dst):
    """Pre-sell .xlsx -> packing plan layout .xlsx. Returns (dst, notes); notes == [] means nothing had to be converted."""
    wb = openpyxl.load_workbook(src, data_only=False)
    wb, notes = convert_workbook(wb)
    if notes:
        wb.save(dst)
    return (dst if notes else None), notes


def main():
    ap = argparse.ArgumentParser(description="Pre-sell packing plan -> packing plan layout")
    ap.add_argument("presell")
    ap.add_argument("-o", "--out", default="packing_plan_converted.xlsx")
    a = ap.parse_args()
    try:
        out, notes = convert_file(a.presell, a.out)
    except PresellError as e:
        sys.exit(f"ERROR: {e}")
    if not out:
        sys.exit("Nothing to convert: the file already has the packing plan headers.")
    for n in notes:
        print(n)
    print(f"-> {out}")


if __name__ == "__main__":
    main()
