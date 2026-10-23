#!/usr/bin/env python3
"""
packing_to_master.py  --  Packing plan (.xlsx)  ->  MASTER_INVOICE.xlsx

Run this BEFORE generate_invoices.py.

How the packing plan is read
  * columns are found by their header text (row 1), not by letter, so a moved column does not break it
  * a PO line   = a row with INV.NO. and PO Number
  * a carton    = the PO lines above an orange "carton total" row (CARTON NO. = 1, 2, 3 ...)
                  or, for an invoice with a single carton, above the green "1 OF 1" row
  * every carton after the first one of the same invoice is a split:
        SC260458  (carton 1),  SC260458-1  (carton 2),  SC260458-2  (carton 3) ...
  * N.W  = sum of  (Ordered Quantity x weight)  of the carton's PO lines
    G.W  = N.W + CARTON WEIGHT of that carton            (both are written on the carton's first row)
  * AMOUNT = Ordered Quantity x Price Per Unit + Total Adjustments
  * INV DATE = Customer Request Date (CRD) of the PO line

Usage
  python packing_to_master.py "Packing Plan.xlsx" -o MASTER_INVOICE.xlsx
  python packing_to_master.py "Packing Plan.xlsx" -o MASTER_INVOICE.xlsx --invoices SC260458,SC260460
  python packing_to_master.py "Packing Plan.xlsx" -o MASTER_INVOICE.xlsx --sheet "Packing Plan on 19.10. rev3"
"""
import argparse
import re
import sys
from collections import OrderedDict
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP

import openpyxl
from openpyxl.styles import Font

try:
    from . import presell                 # inside the Django project
except ImportError:
    import presell                        # run alone: python invoicing/extract.py

MASTER_HEADER = ["INV NO", "INV DATE", "PO NO", "CUSTOMER PO NO", "ART NO", "CUS NO", "COUNTRY", "DESCRIPTION", "HTS CODE",
                 "AMOUNT", "QTY", "N.W", "G.W", "MODE"]

# An invoice whose TOTAL AMOUNT (all its cartons) is below MIN_INVOICE_AMOUNT USD is not built into the master invoice,
# except for these destination countries, which are always built.
MIN_INVOICE_AMOUNT = 50.0
ALWAYS_BUILD_COUNTRIES = {"UNITED KINGDOM", "AUSTRALIA", "SOUTH AFRICA"}
# USA: an invoice whose total G.W (gross weight, all its cartons) is below MIN_USA_GW_KG is not built either.
MIN_USA_GW_KG = 50.0
USA_COUNTRIES = {"USA", "US", "U.S.A.", "UNITED STATES", "UNITED STATES OF AMERICA"}

# MODE = column "Shipment Method" of the packing plan / pre-sell. An invoice with mode AIR-CC is ALWAYS built into the master
# (the under-$50 and the USA under-50-KG checks do not apply to it) and so it is built as an invoice as well.

# header text in the packing plan (compared lower-case, trimmed)
WANT = dict(inv="inv.no.", po="po number", cus="gps customer number", country="country", desc="description",
            hts="hts code1:", price="price per unit", adj="total adjustments", unit_w="weight",
            carton_w="carton weight", carton_no="carton no.",
            cpo="customer number", art="article number",
            crd="customer request date (crd)", mode="shipment method",
            nwk="nw./kgs", gwk="gw./kgs", merch="total merchandise amount")
# other spellings of a header that are accepted (e.g. "HTS.CODE.1" instead of "HTS CODE1:")
ALIASES = dict(carton_w=("cartons weight",), hts=("hts.code.1", "hts code.1", "hts code 1", "hts code1", "hts code"))


def _norm_hdr(h):
    """lower-case and drop spaces, dots, colons, dashes, underscores so 'HTS.CODE.1' == 'hts code1:'"""
    return re.sub(r"[\s.:_/\\-]+", "", str(h).lower())


CARTON_END_TEXT = re.compile(r"^\s*\d+\s+OF\s+\d+\s*$", re.I)


class PackingError(ValueError):
    pass


def _num(v):
    if v is None or v == "":
        return None
    if isinstance(v, str) and v.startswith("="):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _date(v):
    """CRD cell -> date (a real date stays one; text like 2026-10-19 or 19/10/2026 is parsed; else None)."""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        for f in ("%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y", "%d-%m-%Y", "%B %d, %Y"):
            try:
                return datetime.strptime(v.strip(), f).date()
            except ValueError:
                pass
    return None


def _txt(v):
    return None if v is None or str(v).strip() == "" else str(v).strip()


def _money(x):
    return float(Decimal(str(round(x, 6))).quantize(Decimal("0.01"), ROUND_HALF_UP))


def find_sheet(wb, name=None):
    """The sheet to read: the one asked for, else the first one whose row 1 has INV.NO. and PO Number."""
    if name:
        if name not in wb.sheetnames:
            raise PackingError(f"Sheet '{name}' not found. Sheets: {wb.sheetnames}")
        return wb[name]
    for ws in wb.worksheets:
        head = {str(c.value).strip().lower() for c in ws[1] if c.value}
        if "inv.no." in head and "po number" in head:
            return ws
    raise PackingError("No sheet with the headers 'INV.NO.' and 'PO Number' in row 1.")


def column_map(ws):
    hdr = [str(c.value).strip().lower() if c.value is not None else "" for c in ws[1]]
    optional = {"adj", "carton_w", "carton_no", "unit_w", "cpo", "art", "crd", "mode", "nwk", "gwk", "merch"}
    cols, missing = {}, []
    norm = [_norm_hdr(h) for h in hdr]
    for key, text in WANT.items():
        found = None
        for cand in (text,) + ALIASES.get(key, ()):
            if cand in hdr:
                found = hdr.index(cand)
                break
            if _norm_hdr(cand) in norm:
                found = norm.index(_norm_hdr(cand))
                break
        if found is not None:
            cols[key] = found + 1
        elif key not in optional:
            missing.append(text)
    qtys = [i + 1 for i, h in enumerate(hdr) if h == "ordered quantity"]
    if qtys:
        cols["qty"] = qtys[0]                      # order level quantity
        cols["uqty"] = qtys[-1]                    # size level quantity (used for the weight)
    else:
        missing.append("ordered quantity")
    if missing:
        raise PackingError(f"Packing plan is missing columns: {missing}")
    return cols


def _round1(x):
    """Round half up to 1 decimal (15.25 -> 15.3, 15.24 -> 15.2); plain round() would round 15.25 down."""
    from decimal import Decimal, ROUND_HALF_UP
    return float(Decimal(str(round(x, 6))).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP))


def _orig_totals(path, sheet):
    """Carton totals (NW./KGS, GW./KGS) read straight from the uploaded file, before the pre-sell conversion
    (which may not carry them over): {invoice: [(nw, gw), ...]} in carton order. Cached values are used (the totals are formulas)."""
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[sheet] if sheet and sheet in wb.sheetnames else None
        if ws is None:
            for w in wb.worksheets:
                head = {str(x.value).strip().lower() for x in w[1] if x.value}
                if "inv.no." in head and "po number" in head:
                    ws = w
                    break
        if ws is None:
            return {}
        hdr = [_norm_hdr(x.value) if x.value is not None else "" for x in ws[1]]
        col = lambda name: hdr.index(name) + 1 if name in hdr else None
        ci, cp, cn, cg = col("invno"), col("ponumber"), col("nwkgs"), col("gwkgs")
        if not (ci and cp and cg):
            return {}
        out, cur, opened = {}, None, False
        for r in range(2, ws.max_row + 1):
            inv, po = ws.cell(r, ci).value, ws.cell(r, cp).value
            if inv and po:
                cur, opened = str(inv).strip(), True
            elif opened:
                out.setdefault(cur, []).append((_num(ws.cell(r, cn).value) if cn else None, _num(ws.cell(r, cg).value)))
                opened = False
        return out
    except Exception:
        return {}


def _po_key(v):
    try:
        return str(int(float(v)))
    except (TypeError, ValueError):
        return str(v).strip()


def _orig_merch(path, sheet):
    """AMOUNT = column "Total Merchandise Amount" of the uploaded file, read before the pre-sell conversion (which may drop it):
    {(invoice, po): [amount, ...]} in row order. Cached values are used (the column is a formula)."""
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[sheet] if sheet and sheet in wb.sheetnames else None
        if ws is None:
            for w in wb.worksheets:
                head = {str(x.value).strip().lower() for x in w[1] if x.value}
                if "inv.no." in head and "po number" in head:
                    ws = w
                    break
        if ws is None:
            return {}
        hdr = [_norm_hdr(x.value) if x.value is not None else "" for x in ws[1]]
        col = lambda name: hdr.index(name) + 1 if name in hdr else None
        ci, cp, cm = col("invno"), col("ponumber"), col("totalmerchandiseamount")
        if not (ci and cp and cm):
            return {}
        out = {}
        for r in range(2, ws.max_row + 1):
            inv, po = ws.cell(r, ci).value, ws.cell(r, cp).value
            if inv and po:
                out.setdefault((str(inv).strip(), _po_key(po)), []).append(_num(ws.cell(r, cm).value))
        return out
    except Exception:
        return {}


def _carton_weights(ct):
    """(N.W, G.W) of a carton: the values on its total row (NW./KGS, GW./KGS) when the packing plan has them,
    else the sum of the NW./KGS lines, else computed (qty x unit weight; G.W = N.W + carton weight)."""
    nw = ct.get("nw_t") or ct.get("nw_sum") or ct["nw"]
    gw = ct.get("gw_t") or (nw + (ct["cw"] or 0.0))
    return nw, gw


def _is_separator(ws, r, c):
    """A row that is not a PO line and is either filled with a solid colour (the orange / green total rows) or empty."""
    try:
        cell = ws.cell(r, c["inv"])
        if cell.fill is not None and cell.fill.fill_type == "solid":
            rgb = cell.fill.fgColor.rgb if cell.fill.fgColor.type == "rgb" else "theme"
            if rgb not in ("00000000", "FFFFFFFF"):
                return True
    except Exception:
        pass
    return not any(ws.cell(r, col).value not in (None, "") for col in (c["inv"], c["po"], c["desc"]))


def extract(path, sheet=None, only=None, skip=None, min_amount=MIN_INVOICE_AMOUNT, min_usa_gw=MIN_USA_GW_KG):
    """Returns (rows, notes). rows are master rows; notes are warnings (text).
    only = build just these invoices; skip = leave these invoices out of the master (SC260458 skips SC260458 and its -1, -2 ... cartons).
    min_amount = invoices with a total amount below this (USD) are not built, except UK / AUSTRALIA / SOUTH AFRICA (0 = check off).
    min_usa_gw = USA invoices with a total G.W (kg) below this are not built (0 = check off).
    The skipped small invoices are listed in notes and in extract.last_low."""
    wb = openpyxl.load_workbook(path, data_only=False)
    try:
        wb, notes = presell.convert_workbook(wb, sheet)     # "Packing Plan Pre-sell" template -> packing plan layout first
    except presell.PresellError as e:
        raise PackingError(str(e))
    ws = find_sheet(wb, sheet)
    c = column_map(ws)
    g = lambda r, k: ws.cell(r, c[k]).value if k in c else None
    cartons = []
    for k, label in (("art", "Article Number"), ("cpo", "Customer Number"), ("crd", "Customer Request Date (CRD)"), ("mode", "Shipment Method")):
        if k not in c:
            notes.append(f"Column '{label}' not found in the packing plan -> { {'art': 'ART NO', 'cpo': 'CUSTOMER PO NO', 'crd': 'INV DATE', 'mode': 'MODE (the AIR-CC rule is not applied)'}[k] } left blank.")
    cur = None                                      # cartons: list of dict(inv, lines, nw, cw)
    merch_orig, occ = _orig_merch(path, sheet), {}

    def close(r):
        nonlocal cur
        if cur and cur["lines"]:
            if cur["cw"] is None and not cur.get("gw_t"):
                notes.append(f"{cur['inv']}: no CARTON WEIGHT found above row {r} -> G.W. = N.W.")
            cartons.append(cur)
        cur = None

    for r in range(2, ws.max_row + 1):
        inv, po = g(r, "inv"), g(r, "po")
        if inv and po:                                       # PO line
            inv = str(inv).strip()
            if cur and cur["inv"] != inv:
                notes.append(f"Row {r}: invoice changes to {inv} without a carton total row; "
                             f"carton of {cur['inv']} closed here.")
                close(r)
            cur = cur or dict(inv=inv, lines=[], nw=0.0, cw=None, first_row=r)
            k = (inv, _po_key(po))                               # AMOUNT: Total Merchandise Amount of this PO line
            i = occ.get(k, 0)
            occ[k] = i + 1
            merch = _num(g(r, "merch"))
            if merch is None and i < len(merch_orig.get(k, [])):
                merch = merch_orig[k][i]
            q, uq = _num(g(r, "qty")), _num(g(r, "uqty"))
            price, adj = _num(g(r, "price")), _num(g(r, "adj")) or 0.0
            if q is None or price is None:
                notes.append(f"Row {r} PO {po}: quantity or price is empty -> row skipped.")
                continue
            uw = _num(g(r, "unit_w"))
            if uw is None:
                notes.append(f"Row {r} PO {po}: unit weight is empty -> counted as 0.")
                uw = 0.0
            cur["nw"] += (uq if uq is not None else q) * uw
            ln_nw = _num(g(r, "nwk"))
            if ln_nw:
                cur["nw_sum"] = cur.get("nw_sum", 0.0) + ln_nw
            cw = _num(g(r, "carton_w"))
            if cw:
                cur["cw"] = (cur["cw"] or 0.0) + cw
            h = g(r, "hts")
            cur["lines"].append(dict(
                po=int(po) if str(po).isdigit() else po, cus=g(r, "cus"),
                cpo=_txt(g(r, "cpo")), art=_txt(g(r, "art")), mode=_txt(g(r, "mode")), crd=_date(g(r, "crd")), country=str(g(r, "country") or "").strip(),
                desc=str(g(r, "desc") or "").strip(),
                hts=int(h) if isinstance(h, (int, float)) else h,
                amount=_money(merch if merch is not None else q * price + adj), qty=int(q)))
            continue
        # not a PO line: carton total (orange) or invoice total (green) rows
        no = g(r, "carton_no")
        is_carton_total = isinstance(no, (int, float)) or (isinstance(no, str) and no.strip().isdigit())
        is_invoice_total = isinstance(no, str) and bool(CARTON_END_TEXT.match(no))
        # pre-sell files have no "Carton No." numbers: their cartons are separated by an orange (filled) or empty row
        is_separator = _is_separator(ws, r, c)
        if cur and (is_carton_total or is_separator):      # N.W / G.W printed on the carton total row (NW./KGS, GW./KGS)
            tn, tg = _num(g(r, "nwk")), _num(g(r, "gwk"))
            if tn:
                cur["nw_t"] = tn
            if tg:
                cur["gw_t"] = tg
        if is_carton_total or is_invoice_total or is_separator:
            close(r)                                         # an invoice total row after cartons has nothing open
    if cur and cur["lines"]:
        notes.append(f"Last carton of {cur['inv']} has no closing total row; it was kept.")
        close(ws.max_row)

    # N.W / G.W printed on the orange carton rows of the uploaded file win over anything computed
    orig, idx = _orig_totals(path, sheet), {}
    for ct in cartons:
        n = idx.get(ct["inv"], 0)
        idx[ct["inv"]] = n + 1
        tot = orig.get(ct["inv"], [])
        if n < len(tot):
            tn, tg = tot[n]
            if tn:
                ct["nw_t"] = tn
            if tg:
                ct["gw_t"] = tg
    for inv, tot in orig.items():
        if inv in idx and len(tot) != idx[inv]:
            notes.append(f"{inv}: {len(tot)} carton total rows in the uploaded file but {idx[inv]} cartons built -> check N.W / G.W.")

    # condition check BEFORE the master is built: total amount of the whole invoice (all cartons) under min_amount -> not built
    totals, gws, low = {}, {}, {}
    for ct in cartons:
        totals[ct["inv"]] = round(totals.get(ct["inv"], 0.0) + sum(l["amount"] for l in ct["lines"]), 2)
        gws[ct["inv"]] = round(gws.get(ct["inv"], 0.0) + _round1(_carton_weights(ct)[1]), 3)
    modes = {}                                       # invoice -> mode (AIR-CC wins when any PO line of the invoice is AIR-CC)
    for ct in cartons:
        for l in ct["lines"]:
            if l["mode"] and (ct["inv"] not in modes or presell.is_air_cc(l["mode"])):
                modes[ct["inv"]] = presell.MUST_BUILD_MODE_LABEL if presell.is_air_cc(l["mode"]) else l["mode"]
    for inv, tot in totals.items():
        country = next((l["country"] for ct in cartons if ct["inv"] == inv for l in ct["lines"] if l["country"]), "")
        cu = country.strip().upper()
        if presell.is_air_cc(modes.get(inv)):          # mode AIR-CC: master invoice must be created, whatever the amount / weight
            would_skip = (min_amount and tot < min_amount and cu not in ALWAYS_BUILD_COUNTRIES) or \
                         (min_usa_gw and cu in USA_COUNTRIES and gws[inv] < min_usa_gw)
            if would_skip:
                notes.append(f"{inv} ({country or 'no country'}): mode AIR-CC -> master invoice built although total is ${tot:.2f} / "
                             f"G.W {gws[inv]:.2f} KG (below the minimum).")
            continue
        if min_amount and tot < min_amount and cu not in ALWAYS_BUILD_COUNTRIES:
            low[inv] = (tot, country)
            notes.append(f"{inv} ({country or 'no country'}): total ${tot:.2f} is below ${min_amount:.0f} -> no master invoice built "
                         f"(UK, Australia and South Africa are always built).")
        elif min_usa_gw and cu in USA_COUNTRIES and gws[inv] < min_usa_gw:
            low[inv] = (tot, country)
            notes.append(f"{inv} ({country}): total G.W {gws[inv]:.2f} KG is below {min_usa_gw:g} KG -> no master invoice built (USA).")
    extract.last_low = {k: dict(total=v[0], country=v[1]) for k, v in low.items()}

    # splits: 2nd carton of an invoice -> -1, 3rd -> -2 ...
    seen, rows = {}, []
    for ct in cartons:
        n = seen.get(ct["inv"], 0)
        seen[ct["inv"]] = n + 1
        no = ct["inv"] if n == 0 else f"{ct['inv']}-{n}"
        if only and ct["inv"] not in only and no not in only:
            continue
        if skip and (ct["inv"] in skip or no in skip):
            continue
        if ct["inv"] in low:
            continue
        nw, gw = _carton_weights(ct)
        nw, gw = _round1(nw), _round1(gw)                    # N.W / G.W: 1 digit after the decimal point (15.165 -> 15.2)
        for i, ln in enumerate(ct["lines"]):
            rows.append(dict(inv=no, mode=modes.get(ct["inv"]) or ln["mode"], date=ln["crd"], po=ln["po"], cpo=ln["cpo"], art=ln["art"], cus=ln["cus"], country=ln["country"], desc=ln["desc"],
                             hts=ln["hts"], amount=ln["amount"], qty=ln["qty"],
                             nw=nw if i == 0 else None, gw=gw if i == 0 else None))
    if only:
        found = {r["inv"].split("-")[0] for r in rows} | {r["inv"] for r in rows}
        for want in only:
            if want not in found:
                notes.append(f"Invoice {want} was not found in the packing plan.")
    if skip:
        present = {ct["inv"] for ct in cartons}
        for want in skip:
            if want not in present and not any(want == f"{i}-{n}" for i in present for n in range(1, 50)):
                notes.append(f"Invoice {want} to skip was not found in the packing plan.")
    if not rows:
        raise PackingError("No PO lines left to build: every invoice was skipped or is below the minimum amount / USA weight." if (skip or low)
                           else "No PO lines found in the packing plan.")
    return rows, notes


def write_master(rows, out):
    """Same layout as MASTER_INVOICE.xlsx: header, one block per invoice, a blank row between blocks."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(MASTER_HEADER)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    prev = None
    for r in rows:
        if prev is not None and r["inv"] != prev:
            ws.append([])
        prev = r["inv"]
        ws.append([r["inv"], r["date"], r["po"], r["cpo"], r["art"], r["cus"], r["country"], r["desc"], r["hts"],
                   r["amount"], r["qty"], r["nw"], r["gw"], r.get("mode")])
    for row in ws.iter_rows(min_row=2):
        row[11].number_format = row[12].number_format = "0.0"
        row[9].number_format = "0.00"
        row[1].number_format = "dd-mmm-yyyy"
    for col, w in zip("ABCDEFGHIJKLMN", (13, 13, 13, 16, 10, 11, 16, 62, 15, 10, 7, 8, 8, 13)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    wb.save(out)
    return out


def summarize(rows):
    inv = OrderedDict()
    for r in rows:
        d = inv.setdefault(r["inv"], dict(inv=r["inv"], country=r["country"], mode=r.get("mode"), po=0, qty=0, amount=0.0, nw=None, gw=None))
        d["po"] += 1
        d["qty"] += r["qty"]
        d["amount"] = round(d["amount"] + r["amount"], 2)
        if r["nw"] is not None:
            d["nw"], d["gw"] = r["nw"], r["gw"]
    return list(inv.values())


def main():
    ap = argparse.ArgumentParser(description="Packing plan -> MASTER_INVOICE.xlsx")
    ap.add_argument("packing")
    ap.add_argument("-o", "--out", default="MASTER_INVOICE.xlsx")
    ap.add_argument("--sheet", help="sheet name (default: first sheet that has INV.NO. and PO Number)")
    ap.add_argument("--invoices", help="only these invoices, comma separated, e.g. SC260458,SC260460")
    ap.add_argument("--min-amount", type=float, default=MIN_INVOICE_AMOUNT,
                    help="invoices below this total amount (USD) are not built, except UK/AUSTRALIA/SOUTH AFRICA (default 50, 0 = off)")
    ap.add_argument("--min-usa-gw", type=float, default=MIN_USA_GW_KG,
                    help="USA invoices below this total G.W in KG are not built (default 50, 0 = off)")
    ap.add_argument("--skip", help="leave these invoices out of the master, comma separated, e.g. SC260458")
    a = ap.parse_args()
    only = {x.strip() for x in a.invoices.split(",")} if a.invoices else None
    skip = {x.strip() for x in a.skip.split(",")} if a.skip else None
    try:
        rows, notes = extract(a.packing, a.sheet, only, skip, a.min_amount, a.min_usa_gw)
    except PackingError as e:
        sys.exit(f"ERROR: {e}")
    write_master(rows, a.out)
    for s in summarize(rows):
        print(f"{s['inv']:<13}{s['country']:<16}{s['po']:>4} PO  {s['qty']:>4} pcs  {s['amount']:>10.2f} USD  "
              f"N.W {s['nw']}  G.W {s['gw']}")
    print(f"-> {a.out}: {len(rows)} PO lines, {sum(r['qty'] for r in rows)} pcs")
    for n in notes:
        print("WARNING", n)


if __name__ == "__main__":
    main()
