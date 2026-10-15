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
from openpyxl.styles import PatternFill, Alignment, Border, Side

log = logging.getLogger("inv")
YELLOW = PatternFill("solid", fgColor="FFFF00")

# ---------------------------------------------------------------- lookup tables (edit here)
COUNTRY_LABEL = {"UNITED KINGDOM": "UK"}          # text printed after "VIA THAILAND TO"
CONSIGNEES = {                                     # per destination country
    "UNITED KINGDOM": dict(
        name=None,
        address=None,        # None -> keep the address already in the template (UK)
        marks=None),
    # add GERMANY / UNITED STATES / ... here once the real addresses are known
}
# Destination countries that keep the content (composition) line in EVERY group.
# All other countries: content only on the first group, next groups = description + HTS code only.
KEEP_CONTENT_COUNTRIES = {"UNITED KINGDOM", "SOUTH AFRICA", "AUSTRALIA"}
# Maximum printed lines of ONE invoice, counted from the "TOTAL : n CARTONS." row down to the last PO line.
# Only rows that hold a value are counted (the blank spacer rows between PO lines are not).
LINE_LIMIT_KEEP = 28      # UK / SOUTH AFRICA / AUSTRALIA
LINE_LIMIT_OTHER = 34     # every other destination country
REMOVED_NAMES = {"ADIDAS INTERNATIONAL TRADING AG"}     # not printed for an entry WITHOUT address of its own (the UK template entry)
def consignee_name(cons):
    """Name printed in E13. A consignee with its own address always prints its name (e.g. 506010: adidas International Trading AG, Platz 1b ...).
    The plain name ADIDAS INTERNATIONAL TRADING AG is blanked only for an entry without address (UK: the template keeps its own lines)."""
    n = " ".join(str(cons.get("name") or "").split()) if cons else ""
    if n.upper() in REMOVED_NAMES and not (cons or {}).get("address"):
        return None
    return n or None
PLACEHOLDER = "<<{what} - {country}: TO BE CONFIRMED>>"
def cus_key(v):
    """Customer number as a lookup key: '859015', 859015, 859015.0 and '000035' / '35' are the same customer."""
    s = str(v).strip()
    s = s[:-2] if s.endswith(".0") else s
    return s.lstrip("0") or s
COUNTRY_ALIASES = {                                # other spellings / codes of the same destination country -> one name
    "UK": "UNITED KINGDOM", "GB": "UNITED KINGDOM", "GREAT BRITAIN": "UNITED KINGDOM", "ENGLAND": "UNITED KINGDOM",
    "USA": "UNITED STATES", "US": "UNITED STATES", "UNITED STATES OF AMERICA": "UNITED STATES", "AMERICA": "UNITED STATES",
    "KOREA": "SOUTH KOREA", "KOREA SOUTH": "SOUTH KOREA", "KOREA REPUBLIC OF": "SOUTH KOREA", "REPUBLIC OF KOREA": "SOUTH KOREA", "KR": "SOUTH KOREA",
    "CZECH REPUBLIC": "CZECHIA", "CZ": "CZECHIA", "RUSSIAN FEDERATION": "RUSSIA", "RU": "RUSSIA", "VIET NAM": "VIETNAM", "VN": "VIETNAM",
    "TURKIYE": "TURKEY", "TR": "TURKEY", "UAE": "UNITED ARAB EMIRATES", "AE": "UNITED ARAB EMIRATES", "HOLLAND": "NETHERLANDS", "NL": "NETHERLANDS",
    "PEOPLES REPUBLIC OF CHINA": "CHINA", "PRC": "CHINA", "P R CHINA": "CHINA", "CN": "CHINA",
    "TAIWAN PROVINCE OF CHINA": "TAIWAN", "CHINESE TAIPEI": "TAIWAN", "TW": "TAIWAN",
    "HONG KONG SAR": "HONG KONG", "HONGKONG": "HONG KONG", "HK": "HONG KONG", "MACAO": "MACAU",
    "DK": "DENMARK", "DE": "GERMANY", "DEUTSCHLAND": "GERMANY", "FR": "FRANCE", "IT": "ITALY", "ES": "SPAIN", "BE": "BELGIUM", "SE": "SWEDEN",
    "NO": "NORWAY", "FI": "FINLAND", "AT": "AUSTRIA", "CH": "SWITZERLAND", "PL": "POLAND", "PT": "PORTUGAL", "IE": "IRELAND", "CA": "CANADA",
    "AU": "AUSTRALIA", "NZ": "NEW ZEALAND", "JP": "JAPAN", "SG": "SINGAPORE", "MY": "MALAYSIA", "TH": "THAILAND", "ID": "INDONESIA",
    "IN": "INDIA", "ZA": "SOUTH AFRICA", "MX": "MEXICO", "BR": "BRAZIL",
}
def country_key(v):
    """Country as a lookup key: capitals, no punctuation, 'THE ' dropped, other spellings / ISO codes mapped (USA = UNITED STATES = US)."""
    s = re.sub(r"[^A-Z0-9 ]+", " ", str(v or "").upper().replace("&", " AND "))
    s = " ".join(s.split())
    if s.startswith("THE "):
        s = s[4:]
    return COUNTRY_ALIASES.get(s, s)
def find_consignee(consignees, country, cus=None):
    """Consignee for the customer number (CUS NO) of an invoice. The invoice's customer number is compared with the customer number
    saved on each consignee (000035 = 35 = 35.0); the COUNTRY does not matter (the consignee's own country, e.g. SWITZERLAND, can differ
    from the destination COUNTRY of the invoice, e.g. DENMARK). The table is keyed by cus_key(customer number).
    A consignee of another customer number is never used, and neither is an entry without customer number (no guessing).
    Only UNITED KINGDOM keeps its template default: an entry without customer number, else the built-in template address (CONSIGNEES above).
    Anything else -> None (-> yellow placeholder, and the customer number appears in the Consignee missing list)."""
    table = consignees if consignees is not None else CONSIGNEES
    k = cus_key(cus) if cus not in (None, "") else ""
    if k:
        c = table.get(k)
        if c is not None:
            return c
    if country_key(country) == "UNITED KINGDOM":                            # UK never needs an entry: it keeps the consignee address of the Excel template
        c = table.get("UNITED KINGDOM")
        return c if c is not None else CONSIGNEES.get("UNITED KINGDOM")
    return None
def fill_consignee(ws, shift, country, cus, consignees, marks, inv_no, warn=True):
    """CONSIGNEE block (E13 / E14, per destination country and customer number) and SHIPPING MARKS (per CUSTOMER NO.).
    Shipping marks: the customer number's marks first, else the marks saved with the country, else (UK template) the template text;
    a customer without marks gets a yellow placeholder."""
    cons = find_consignee(consignees, country, cus)
    cell = "B%d" % (55 + shift)
    cm = (marks or {}).get(cus_key(cus))
    if cons is not None:                           # known destination: write what is on file, blank = keep template (UK)
        ws["E13"] = consignee_name(cons)           # blank when there is no name (ADIDAS INTERNATIONAL TRADING AG removed)
        if cons.get("address"):
            ws["E14"] = cons["address"]
    else:                                          # unknown destination -> never reuse the UK address
        ws["E13"] = None
        ws["E14"] = PLACEHOLDER.format(what="CONSIGNEE ADDRESS", country=country)
        ws["E14"].fill = YELLOW
        if warn: log.warning("%s: no consignee on file for %s -> placeholder written (yellow)", inv_no, country)
    if cm:
        ws[cell] = cm
    elif cons is not None and cons.get("marks"):
        ws[cell] = cons["marks"]
    elif cons is None or cons.get("address"):      # not the UK template: never print the template marks for another customer
        ws[cell] = PLACEHOLDER.format(what="SHIPPING MARKS", country=f"{country}, customer {cus}")
        ws[cell].fill = YELLOW
        if warn: log.warning("%s: no shipping marks on file for customer %s -> placeholder written (yellow)", inv_no, cus)
# ----------------------------------------------------------------------------------------

CUT_FROM, CUT_ROWS = 15, 2                          # the 2 empty rows 15-16 under the CONSIGNEE block are cut out of every sheet
_REF = re.compile(r'"[^"]*"|\'[^\']*\'|(?<![A-Za-z0-9_.])(\$?[A-Z]{1,3}\$?)(\d+)(?![A-Za-z0-9_(])')
def _shift_formula(f, first, n):
    """Cell references to rows below the cut move up by n (text in "..." and sheet names in '...' are left alone)."""
    def one(m):
        if m.group(1) is None:
            return m.group(0)
        row = int(m.group(2))
        return m.group(1) + str(row - n if row >= first + n else row)
    return _REF.sub(one, f)
def cut_rows(wb, first=CUT_FROM, n=CUT_ROWS):
    """Delete rows first .. first+n-1 from every sheet of the finished workbook: cells + styles, merged cells, row heights and every
    formula (also the ones that point to another sheet, e.g. PL -> INV) move up with the rows below."""
    for ws in wb.worksheets:
        merges = [(m.min_row, m.max_row, m.min_col, m.max_col) for m in ws.merged_cells.ranges]
        heights = {k: v.height for k, v in ws.row_dimensions.items() if v.height is not None and k >= first}
        for m in [m for m in list(ws.merged_cells.ranges) if m.max_row >= first]:
            ws.unmerge_cells(str(m))
        ws.delete_rows(first, n)
        for k in heights: ws.row_dimensions[k].height = None
        for k, h in heights.items():
            if k >= first + n: ws.row_dimensions[k - n].height = h
        for a, b, c, d in merges:
            if b < first:                                                     # above the cut: was not touched
                continue
            if a >= first and b < first + n:                                  # only inside the cut rows -> gone
                continue
            a2 = a - n if a >= first + n else min(a, first)
            b2 = b - n if b >= first + n else first - 1
            if (a2, c) != (b2, d):
                ws.merge_cells(start_row=a2, end_row=b2, start_column=c, end_column=d)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    c.value = _shift_formula(c.value, first, n)

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

def hide_for(country, override=None):
    """True = remove the content line from the next groups. Decided by destination country unless `override` is True/False."""
    if override is not None:
        return bool(override)
    return str(country).strip().upper() not in KEEP_CONTENT_COUNTRIES

def norm(t):
    """Comparison form of a text: spaces removed, upper case ('...RECYCLED) WOVEN' == '...RECYCLED)WOVEN')."""
    return re.sub(r"\s+", "", str(t or "")).upper()

def fabric_type(desc):
    """'KNITTED' / 'WOVEN' / '' -- read from the whole description. Woven and knitted are never combined."""
    u = " ".join(str(desc or "").upper().split())
    m = re.search(r"\d+%", u)
    for text in ((u[m.start():], u) if m else (u,)):      # the content text first, then the whole description
        k, w = text.find("KNIT"), text.find("WOVEN")
        if k >= 0 or w >= 0:
            return "WOVEN" if k < 0 or (w >= 0 and w < k) else "KNITTED"
    return ""

def desc_suffix(items, show_comp):
    """A group WITHOUT its own content line gets KNITTED / WOVEN at the end of the description line:
    'UNISEX PANTS  (HTS.CODE : 6204631890) WOVEN'."""
    return "" if show_comp else fabric_type(items[0]["desc"])

def desc_text(name, hts, suffix):
    return f"{name}  (HTS.CODE : {hts})" + (f" {suffix}" if suffix else "")

def group_rows(rows, hide_comp=False):
    """Group PO lines. Normal: description (gender + garment) + HTS + composition. hide_comp: description + HTS + fabric
    (KNITTED / WOVEN), so e.g. UNISEX JACKET / 6202401091 becomes ONE group whatever the % composition was, but a WOVEN
    and a KNITTED line are NEVER merged.
    Texts are compared ignoring spaces and case, so a content line typed with a stray space still joins its group;
    the group shows the text of its first PO line."""
    groups, seen = OrderedDict(), {}
    for r in rows:
        name, comp = split_desc(r["desc"])
        fab = fabric_type(r["desc"])
        nk = (norm(name), norm(r["hts"]), fab) if hide_comp else (norm(name), norm(r["hts"]), norm(comp))
        key = seen.setdefault(nk, (name, r["hts"], fab) if hide_comp else (name, r["hts"], comp))
        groups.setdefault(key, []).append(r)
    return groups

def comp_flags(groups, hide_comp):
    """One bool per group: print its content (composition) line?  Normal: always.  hide_comp: only the first group;
    the next groups carry KNITTED / WOVEN at the end of their description line instead (see desc_suffix)."""
    return [(not hide_comp) or n == 0 for n in range(len(groups))]

def count_lines(rows, hide_comp):
    """Printed lines of an invoice = TOTAL CARTONS row + CUSTOMER NO. row + (description [+ content]) per group
    + the PO.NO./CUSTOMER PO NO./ART.NO. header + one line per PO line.  Blank spacer rows are not counted."""
    groups = group_rows(rows, hide_comp)
    return 2 + len(groups) + sum(comp_flags(groups, hide_comp)) + 1 + len(rows)

def line_limit(country):
    return LINE_LIMIT_KEEP if str(country).strip().upper() in KEEP_CONTENT_COUNTRIES else LINE_LIMIT_OTHER

def hts10(v):
    s = str(v).strip()
    return s if "+" in s else s[:10]

SET_PCS = 2        # one BABIES SETS (or any single-HTS "SETS") line = 2 pieces

def expand_set(r):
    """SET lines -> invoice lines. Returns (rows, sets).
    * single HTS ("BABIES SETS")      : qty x SET_PCS pieces, price per piece, amount unchanged
    * two HTS ("UNISEX SETS(T-SHIRT +SHORTS)", HTS a+b) : split into one line per garment (UNISEX T-SHIRT / UNISEX SHORTS),
      qty = number of sets for each, amount split in cents (the parts add up exactly to the original amount)."""
    name, comp = split_desc(r["desc"])
    if not re.search(r"\bSETS?\b", name.upper()):
        return [r], 0
    parts = [x.strip()[:10] for x in str(r["hts"]).split("+")]
    if len(parts) == 1:
        d = dict(r); d["qty"] = r["qty"] * SET_PCS
        return [d], r["qty"]
    m = re.search(r"\(([^)]*\+[^)]*)\)", name)
    items = [x.strip() for x in m.group(1).split("+")] if m else []
    if len(items) != len(parts):
        log.warning("PO %s: cannot split set '%s' (%d items / %d HTS) -> left as one line", r["po"], name, len(items), len(parts))
        return [r], 0
    prefix = re.sub(r"\s*SETS?\s*$", "", name[:m.start()].strip(), flags=re.I)
    cents, k, used, out = int(round(r["amt"] * 100)), len(parts), 0, []
    for i, (item, h) in enumerate(zip(items, parts)):
        c = cents // k if i < k - 1 else cents - used
        used += c
        d = dict(r); d.update(desc=f"{prefix} {item}  {comp}".strip(), hts=h, amt=c / 100)
        out.append(d)
    return out, r["qty"]

def carton_text(inv):
    """'TOTAL : 1 CARTONS. ' ; with sets: 'TOTAL : 1 CARTONS. = 6 SETS = 12 PCS'."""
    n = inv.get("cartons", 1)
    if inv.get("sets"):
        return f"TOTAL : {n} CARTONS. = {inv['sets']} SETS = {sum(x['qty'] for x in inv['rows'])} PCS"
    return f"TOTAL : {n} CARTONS. "

def widen_carton_row(ws, r):
    """Longer carton text: merge C:J instead of C:G so the text is not cut off."""
    for rng in (f"C{r}:G{r}", f"I{r}:J{r}"):
        if rng in {str(m) for m in ws.merged_cells.ranges}:
            ws.unmerge_cells(rng)
    ws.merge_cells(start_row=r, end_row=r, start_column=3, end_column=10)

def widen_text_row(ws, r):
    """Description / content rows: merge C:R (up to the QUANTITY column) instead of C:J, so a long line
    ('MEN'S LONG SLEEVE T-SHIRT  (HTS.CODE : 6109902000) KNITTED') is shown in full on ONE line."""
    have = {str(m) for m in ws.merged_cells.ranges}
    for rng in (f"C{r}:J{r}", f"K{r}:N{r}", f"O{r}:R{r}"):
        if rng in have:
            ws.unmerge_cells(rng)
    ws.merge_cells(start_row=r, end_row=r, start_column=3, end_column=18)

def balance_prices(inv_no, rows):
    """Unit prices are rounded to 2 decimals (half up). The rounding difference is added to ANOTHER PO line
    (the last line with qty 1, else a line where it divides evenly by the qty) so the invoice total stays the same."""
    if not rows:
        return
    cents = lambda x: int((Decimal(str(round(x, 6))) * 100).to_integral_value(ROUND_HALF_UP))
    orig = [cents(r["amt"]) for r in rows]
    price = [(2 * c + r["qty"]) // (2 * r["qty"]) for c, r in zip(orig, rows)]      # unit price in cents, half up
    diff = sum(orig) - sum(p * r["qty"] for p, r in zip(price, rows))                 # cents still to give back
    if diff:
        ones = [i for i, r in enumerate(rows) if r["qty"] == 1]
        even = [i for i, r in enumerate(rows) if diff % r["qty"] == 0]
        pick = ones[-1] if ones else (even[-1] if even else None)
        if pick is not None:
            price[pick] += diff // rows[pick]["qty"]
            diff = 0
    for r, p, c in zip(rows, price, orig):
        r["amt"] = p * r["qty"] / 100
    if diff:                                           # no line can take it with a 2-decimal price -> last line amount
        rows[-1]["amt"] = round(rows[-1]["amt"] + diff / 100, 2)
        log.warning("%s: rounding difference %.2f put on the last PO line (price keeps more than 2 decimals)", inv_no, diff / 100)

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
        inv = invs.setdefault(str(g("INV NO")).strip(), dict(rows=[], raw=[], sets=0, nw=None, gw=None, cus=set(), country=set(), dates=[], mode=set(),
                                                                     cartons=1, names=[str(g("INV NO")).strip()]))
        opt = lambda k: (str(row[ix[k]]).strip() if k in ix and row[ix[k]] not in (None, "") else None)
        raw = dict(po=str(g("PO NO")), desc=g("DESCRIPTION"), hts=hts10(g("HTS CODE")),
                   amt=float(g("AMOUNT")), qty=int(g("QTY")),
                   cpo=opt("CUSTOMER PO NO"), art=opt("ART NO"))
        new_rows, n_sets = expand_set(raw)             # BABIES SETS / UNISEX SETS(T-SHIRT +SHORTS) -> invoice lines
        inv["raw"].append(raw); inv["rows"].extend(dict(x) for x in new_rows); inv["sets"] += n_sets
        d = fmt_date(row[ix["INV DATE"]]) if "INV DATE" in ix else None
        if d and d not in inv["dates"]: inv["dates"].append(d)
        if opt("MODE"): inv["mode"].add(opt("MODE").upper())
        inv["cus"].add(g("CUS NO")); inv["country"].add(str(g("COUNTRY")).strip().upper())
        if g("N.W") is not None: inv["nw"] = g("N.W")
        if g("G.W") is not None: inv["gw"] = g("G.W")
    for no, inv in invs.items():
        balance_prices(no, inv["rows"])
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

def close_gap(ws, row):
    """Packing list GRAND TOTAL row: thin top line + double bottom line on EVERY cell from QTY (S) to the end of the
    G.W block (Y), spacer column W included, so the line is unbroken and nothing is missing."""
    line = Side(style="thin"); dbl = Side(style="double")
    for c in range(19, 26):                           # S T U V W X Y
        cell = ws.cell(row, c)
        b = copy.copy(cell.border)
        cell.border = Border(left=b.left, right=b.right, top=line, bottom=dbl)

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
    for r in po_rows:                                  # quantities are LINKED to the invoice sheet (change INV -> PL follows)
        pl.cell(r, 19).value = f"='{ws.title}'!S{r}"
    pl.cell(tot, 19).value = f"=SUM(S{po_rows[0]}:S{tot-1})"   # GRAND TOTAL: qty | N.W | G.W  (all formulas)
    pl.cell(tot, 21)._style = copy.copy(pl.cell(tot, 24)._style)
    pl.cell(tot, 21).value = f"=SUM(U{po_rows[0]}:U{tot-1})"
    pl.cell(tot, 21).number_format = pl.cell(tot, 24).number_format = "0.00"
    pl.cell(tot, 21).alignment = Alignment(horizontal="center", vertical=pl.cell(tot, 21).alignment.vertical)  # N.W total under the N.W lines
    close_gap(pl, tot)
    pl.cell(tot, 24).value = f"=SUM(X{po_rows[0]}:X{tot-1})"
    pl.cell(tot, 24).alignment = Alignment(horizontal="right", vertical=pl.cell(tot, 24).alignment.vertical)
    # GRAND TOTAL row of the packing list: qty | N.W | G.W, and the TOTAL N.W./G.W. lines below are linked to it
    shift = tot - 50
    if inv["nw"] is not None and inv["gw"] is not None:
        pl.cell(56 + shift, 14).value = f'="TOTAL N.W.:                 "&TEXT(U{tot},"0.00")&"  KGS."'
        pl.cell(58 + shift, 14).value = f'="TOTAL G.W.:                 "&TEXT(X{tot},"0.00")&"  KGS."'
    return pl

def build(inv_no, inv, template, pack, date_txt, outdir, wb=None, inv_title=None, pl_title=None, consignees=None, hide_comp=None, marks=None):
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
    hide_comp = hide_for(country, hide_comp)       # UK / SOUTH AFRICA / AUSTRALIA keep content, other countries remove it
    # ---- groups
    groups = group_rows(inv["rows"], hide_comp)
    n_po, n_grp = len(inv["rows"]), len(groups)
    flags = comp_flags(groups, hide_comp)
    lines, limit = count_lines(inv["rows"], hide_comp), line_limit(country)
    if lines > limit:
        log.warning("%s: %d printed lines, more than the %d allowed for %s (one carton cannot be split)", inv_no, lines, limit, country)
    # per group: description line + composition line (when flagged); + header once + PO line and spacer row per PO
    need_last = 24 + n_grp + sum(flags) + 1 + 2 * n_po
    inv["lines"] = lines
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
    if len(cus) > 1:
        log.warning("%s has several CUS NO %s", inv_no, cus)
        found = [find_consignee(consignees, country, c) for c in cus]          # every customer number of the invoice is compared with the consignees
        if any(f is None for f in found) or len({(f or {}).get("address") for f in found}) > 1:
            log.warning("%s: its CUS NO %s do not all have the same consignee on file -> check the CONSIGNEE block (customer %s is used)", inv_no, cus, cus[0])
    ws["C24"] = f"CUSTOMER NO. {cus[0]}"
    ws["C23"] = carton_text(inv)
    if inv.get("sets"): widen_carton_row(ws, 23)
    fill_consignee(ws, shift, country, cus[0], consignees, marks, inv_no)

    # ---- table
    r = 25; first = True; po_rows = []
    for (key, items), show_comp in zip(groups.items(), flags):
        name, hts = key[0], key[1]
        apply(ws, r, T["desc"]); ws.cell(r, 3).value = desc_text(name, hts, desc_suffix(items, show_comp)); widen_text_row(ws, r); r += 1
        if show_comp:                              # first group (and a fabric change) always shows its content, the others only when not removed
            apply(ws, r, T["comp"]); ws.cell(r, 3).value = split_desc(items[0]["desc"])[1]; widen_text_row(ws, r); r += 1
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
    ws.cell(60 + shift, 14).value = f"TOTAL PACKED:          {inv.get('cartons', 1)}  CARTONS."
    ws.title = inv_title or inv_no[:31]
    add_packing_list(wb, ws, inv_no, po_rows, tot, inv, pl_title)
    out = None
    if not shared:
        out = f"{outdir}/{inv_no}.xlsx"
        if "__tpl" in wb.sheetnames: del wb["__tpl"]
        cut_rows(wb)
        wb.save(out)
    # ---- reconciliation with master
    exp_q = sum(x["qty"] for x in inv["rows"]); exp_a = round(sum(x["amt"] for x in inv["rows"]), 2)
    for it in inv.get("raw", inv["rows"]):         # compare the ORIGINAL master lines (before sets are expanded)
        p = pack.get(it["po"])
        if p and p["price"] is not None:
            calc = round(p["qty"] * p["price"] + (p["adj"] or 0), 2)
            if abs(calc - it["amt"]) > 0.005 or p["qty"] != it["qty"]:
                log.warning("%s PO %s differs from packing plan (master %.2f / plan %.2f)", inv_no, it["po"], it["amt"], calc)
        elif pack and not it.get("art"):
            log.warning("%s PO %s not found in packing plan (ART.NO. left blank)", inv_no, it["po"])
    return out, n_po, n_grp, exp_q, exp_a

def build_bkk(wb, base, fam, pack, date_txt, title="INV.BKK", consignees=None, hide_comp=None, marks=None):
    """INV.BKK = ONE summary commercial invoice for a whole invoice family (SC260460, SC260460-1, SC260460-2 ...).
    One block per carton (= per invoice) with its own TOTAL row, then GRAND TOTAL; footer = total N.W. / G.W. / cartons.
    `fam` = list of (invoice_no, inv) in order. The sheet is added to `wb` (which must hold the "__tpl" sheet)."""
    tpl = wb["__tpl"]
    ws = wb.copy_worksheet(tpl)
    ws.sheet_view.showGridLines = tpl.sheet_view.showGridLines
    if tpl.print_area:
        ws.print_area = tpl.print_area.split("!")[-1].replace("$", "")
    first_inv = fam[0][1]
    country = next(iter(first_inv["country"]))
    countries = {c for _, i in fam for c in i["country"]}
    if len(countries) > 1:
        log.warning("%s: several destination countries %s in one INV.BKK -> %s is used", base, sorted(countries), country)

    # ---- plan every row first (so the footer can be shifted once)
    plan, prev_cus, first_group = [], None, True
    for n, (no, inv) in enumerate(fam):
        hide = hide_for(next(iter(inv["country"])), hide_comp)
        cus = sorted(map(str, inv["cus"]))[0]
        plan.append(("carton", None))
        if cus != prev_cus:
            plan.append(("cus", cus)); prev_cus = cus
        grps = group_rows(inv["rows"], hide)
        for (key, items), show_comp in zip(grps.items(), comp_flags(grps, hide)):
            plan.append(("desc", (key[0], key[1], desc_suffix(items, show_comp))))
            if show_comp:
                plan.append(("comp", split_desc(items[0]["desc"])[1]))
            if first_group:
                plan.append(("hdr", None)); first_group = False
            for it in items:
                plan.append(("po", it)); plan.append(("sp", None))
        plan.append(("total", None))
        if n < len(fam) - 1:
            plan.append(("blank", None))
    need_last = 22 + len(plan)
    shift = need_last - 47

    # ---- snapshots, then re-flow the footer (same method as build())
    T = dict(carton=snap(ws, 23), cus=snap(ws, 24), desc=snap(ws, 25), comp=snap(ws, 26), hdr=snap(ws, 27),
             po=snap(ws, 28), sp=snap(ws, 29), total=snap(ws, 50))
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

    # ---- header block (invoice number = the family number)
    ws["C8"] = f"INV.NO.  {base}"
    dates = list(dict.fromkeys(d for _, i in fam for d in i.get("dates", [])))
    inv_date = dates[0] if dates else date_txt
    if inv_date is None:
        raise ValueError(f"{base}: no invoice date (INV DATE is empty in the master and the packing plan has no CRD).")
    if len(dates) > 1:
        log.warning("%s INV.BKK: several INV DATE values %s -> the first one is used", base, dates)
    ws["C9"] = f"DATE : {inv_date}"
    label = COUNTRY_LABEL.get(country, country)
    ws["C19"] = re.sub(r"TO\s+\S+\s*$", f"TO {label}", ws["C19"].value.strip())
    fill_consignee(ws, shift, country, sorted(map(str, first_inv["cus"]))[0], consignees, marks, base, warn=False)

    # ---- table
    centre = lambda c: Alignment(horizontal="center", vertical=c.alignment.vertical)
    r, n_inv, sec_start, sub_rows = 23, 0, None, []
    cartons = iter(fam)
    cur_inv = None
    for kind, val in plan:
        if kind == "carton":
            cur_inv = next(cartons)[1]
            if r != 23: apply(ws, r, T["carton"])     # row 23 / 24 already exist in the template (header block)
            ws.cell(r, 3).value = carton_text(cur_inv)
            if cur_inv.get("sets"): widen_carton_row(ws, r)
            sec_start = r + 1; r += 1
        elif kind == "cus":
            if r != 24: apply(ws, r, T["cus"])
            ws.cell(r, 3).value = f"CUSTOMER NO. {val}"; r += 1
        elif kind == "desc":
            apply(ws, r, T["desc"]); ws.cell(r, 3).value = desc_text(*val); widen_text_row(ws, r); r += 1
        elif kind == "comp":
            apply(ws, r, T["comp"]); ws.cell(r, 3).value = val; widen_text_row(ws, r); r += 1
        elif kind == "hdr":
            apply(ws, r, T["hdr"])
            ws.cell(r, 3).value, ws.cell(r, 4).value, ws.cell(r, 11).value = "PO.NO.", "CUSTOMER PO NO.", "ART.NO."
            r += 1
        elif kind == "po":
            it = val; p = pack.get(it["po"], {})
            apply(ws, r, T["po"])
            ws.cell(r, 3).value = int(it["po"]) if it["po"].isdigit() else it["po"]
            ws.cell(r, 4).value = it.get("cpo") or p.get("cus"); ws.cell(r, 11).value = it.get("art") or p.get("art")
            ws.cell(r, 19).value = it["qty"]
            ws.cell(r, 21).value = round(it["amt"] / it["qty"], 6)
            ws.cell(r, 24).value = f"=S{r}*U{r}"
            r += 1
        elif kind == "sp":
            apply(ws, r, T["sp"]); r += 1
        elif kind == "total":                         # TOTAL row of this carton
            apply(ws, r, T["total"])
            ws.cell(r, 11).value = "TOTAL:"
            ws.cell(r, 19).value = f"=SUM(S{sec_start}:S{r-1})"; ws.cell(r, 19).alignment = centre(ws.cell(r, 19))
            ws.cell(r, 24).value = f"=SUM(X{sec_start}:X{r-1})"
            sub_rows.append(r); r += 1
        elif kind == "blank":
            apply(ws, r, T["sp"]); r += 1
    assert r - 1 == need_last, (r - 1, need_last)

    # ---- GRAND TOTAL = sum of the carton TOTAL rows
    tot = 50 + shift
    ws.cell(tot, 19).value = "=" + "+".join(f"S{x}" for x in sub_rows)
    ws.cell(tot, 24).value = "=" + "+".join(f"X{x}" for x in sub_rows)
    ws.cell(tot, 19).alignment = centre(ws.cell(tot, 19))

    # ---- footer: total weights and cartons of the whole family
    if any(i["nw"] is None or i["gw"] is None for _, i in fam):
        log.warning("%s INV.BKK: N.W./G.W. missing in master for some cartons -> total weights left blank", base)
    else:
        ws.cell(56 + shift, 14).value = f"TOTAL N.W.:                 {sum(r2(i['nw']) for _, i in fam)}  KGS."
        ws.cell(58 + shift, 14).value = f"TOTAL G.W.:                 {sum(r2(i['gw']) for _, i in fam)}  KGS."
    ws.cell(60 + shift, 14).value = f"TOTAL PACKED:          {sum(i.get('cartons', 1) for _, i in fam)}  CARTONS."
    ws.title = title[:31]
    return ws

def family_map(invs):
    """base invoice number -> [(invoice_no, inv), ...] in master order (SC260460, SC260460-1, SC260460-2 ...)."""
    fams = OrderedDict()
    for no, inv in invs.items():
        fams.setdefault(no.split("-")[0], []).append((no, inv))
    return fams

def _merge_key(inv):
    """Cartons may only share an invoice when country, CUSTOMER NO., INV DATE and fabric agree.
    A carton that holds BOTH woven and knitted lines, or any SET line (BABIES SETS, UNISEX SETS T-SHIRT+SHORTS),
    is never merged with another carton."""
    if inv.get("sets") or any(re.search(r"\bSETS?\b", split_desc(x["desc"])[0].upper()) for x in inv["raw"]):
        return None                                        # BABIES SETS / UNISEX SETS(T-SHIRT +SHORTS): never merged
    fabrics = frozenset(fabric_type(x["desc"]) for x in inv["rows"])
    if len(fabrics) > 1:
        return None
    return (frozenset(inv["country"]), frozenset(map(str, inv["cus"])), tuple(inv["dates"]), fabrics, frozenset(inv.get("mode", ())))

def _join(a, b):
    """Add carton b to the (merged) invoice a."""
    a["rows"] += [dict(x) for x in b["rows"]]; a["raw"] += b["raw"]; a["sets"] += b["sets"]
    a["cus"] |= b["cus"]; a["country"] |= b["country"]; a.setdefault("mode", set()).update(b.get("mode", ()))
    a["dates"] += [d for d in b["dates"] if d not in a["dates"]]
    a["cartons"] += b["cartons"]; a["names"] += b["names"]
    if a["nw"] is None or b["nw"] is None or a["gw"] is None or b["gw"] is None:
        a["nw"] = a["gw"] = None
    else:                                                  # weights are rounded per carton (0.1 KG), then added
        a["nw"] = float(r2(a["nw"]) + r2(b["nw"])); a["gw"] = float(r2(a["gw"]) + r2(b["gw"]))

def merge_cartons(invs, hide_comp=None):
    """Master invoice (one INV NO per carton: SC260458, SC260458-1 ...) -> invoices with several cartons each.
    Per invoice family, every carton goes into the first invoice where
      * country, CUSTOMER NO., INV DATE and fabric (KNITTED / WOVEN) are the same -- woven and knitted never share an invoice,
      * and the printed lines after combining (count_lines: TOTAL CARTONS row ... last PO line, blank rows not counted)
        stay within the limit: 28 for UK / SOUTH AFRICA / AUSTRALIA, 34 for every other country.
    Same description + HTS code merge into one group, the carton counts add up (TOTAL : 2 CARTONS.).
    Families where nothing could be merged keep their master numbers; otherwise the invoices are renumbered
    SC260458, SC260458-1 ... in order."""
    out = OrderedDict()
    for base, fam in family_map(invs).items():
        bins = []
        for no, inv in fam:
            country = next(iter(inv["country"])); key = _merge_key(inv)
            hide, limit = hide_for(country, hide_comp), line_limit(country)
            for b in bins:
                if key is not None and b["key"] == key and count_lines(b["inv"]["rows"] + inv["rows"], hide) <= limit:
                    _join(b["inv"], inv); break
            else:
                first = dict(inv); first.update(rows=[dict(x) for x in inv["rows"]], raw=list(inv["raw"]), cus=set(inv["cus"]),
                                                country=set(inv["country"]), dates=list(inv["dates"]), names=list(inv["names"]))
                bins.append(dict(key=key, inv=first, orig=no))
        merged = any(b["inv"]["cartons"] > 1 for b in bins)
        for i, b in enumerate(bins):
            out[(base if i == 0 else f"{base}-{i}") if merged else b["orig"]] = b["inv"]
    return out

def sheet_names(inv_no, multi_base):
    """INV / PL / INV-1 / PL-1 for one invoice number; with several invoice numbers: 'INV SC260458', 'PL SC260458-1' ..."""
    if multi_base:
        return f"INV {inv_no}"[:31], f"PL {inv_no}"[:31]
    suf = inv_no.split("-", 1)[1] if "-" in inv_no else ""
    sfx = f"-{suf}" if suf else ""
    return "INV" + sfx, "PL" + sfx

def describe(inv_no, inv, pack, hide_comp=None):
    """Plain-data preview of one invoice (groups and PO lines) for the dashboard."""
    hide_comp = hide_for(next(iter(inv["country"])), hide_comp)
    groups = group_rows(inv["rows"], hide_comp)
    out = []
    for (key, items), show_comp in zip(groups.items(), comp_flags(groups, hide_comp)):
        name, hts = key[0], key[1]
        comp = split_desc(items[0]["desc"])[1] if show_comp else ""
        lines = []
        for it in items:
            p = pack.get(it["po"], {})
            lines.append(dict(po=it["po"], cus_po=it.get("cpo") or p.get("cus") or "", art=it.get("art") or p.get("art") or "", qty=it["qty"],
                              price=round(it["amt"] / it["qty"], 2), amount=round(it["amt"], 2)))
        out.append(dict(name=(name + " " + desc_suffix(items, show_comp)).strip(), hts=hts, comp=comp, lines=lines))
    return out

class _Collect(logging.Handler):
    def __init__(self):
        super().__init__(logging.WARNING); self.items = []
    def emit(self, record):
        self.items.append(record.getMessage())

def process(master, template, packing, date_txt, outdir, consignees=None, hide_comp=None, merge=True, marks=None):
    """Build every invoice (+ packing list) and the combined workbook. Returns a JSON-friendly result dict."""
    import os, zipfile
    os.makedirs(outdir, exist_ok=True)
    pack, crd = read_packing(packing)
    if not date_txt:                                   # only a fallback for a master without INV DATE
        date_txt = crd.strftime("%B %d , %Y").upper() if crd else None
    invs = read_master(master)
    if not invs:
        raise ValueError("MASTER_INVOICE has no data rows.")
    n_master = len(invs)
    if merge:
        invs = merge_cartons(invs, hide_comp)
    h = _Collect(); log.addHandler(h)
    try:
        res, files, wb_all = [], [], openpyxl.load_workbook(template)
        wb_all.active.title = "__tpl"
        multi = len({n.split("-")[0] for n in invs}) > 1
        fam_wb = {}                                   # base invoice -> ONE workbook holding INV.BKK, INV, PL, INV-1, PL-1 ...
        fams = family_map(invs)
        for no, inv in invs.items():
            h.items.clear()
            base = no.split("-")[0]
            if base not in fam_wb:
                fam_wb[base] = openpyxl.load_workbook(template)
                fam_wb[base].active.title = "__tpl"
                build_bkk(fam_wb[base], base, fams[base], pack, date_txt, "INV.BKK", consignees, hide_comp, marks)
                build_bkk(wb_all, base, fams[base], pack, date_txt, f"INV.BKK {base}" if multi else "INV.BKK", consignees, hide_comp, marks)
            _, n_po, n_grp, q, amt = build(no, inv, template, pack, date_txt, outdir, wb=fam_wb[base],
                                           inv_title=sheet_names(no, False)[0], pl_title=sheet_names(no, False)[1],
                                           consignees=consignees, hide_comp=hide_comp, marks=marks)
            it, pt = sheet_names(no, multi)
            build(no, inv, template, pack, date_txt, outdir, wb=wb_all, inv_title=it, pl_title=pt,
                  consignees=consignees, hide_comp=hide_comp, marks=marks)
            warns = list(dict.fromkeys(h.items))          # same warning comes from both builds -> keep one
            country = next(iter(inv["country"]))
            res.append(dict(
                no=no, country=country, cus=sorted(map(str, inv["cus"]))[0], n_po=n_po, n_grp=n_grp, qty=q,
                cartons=inv["cartons"], merged_from=inv["names"] if inv["cartons"] > 1 else [],
                lines=inv.get("lines"), limit=line_limit(country),
                amount=amt, nw=float(r2(inv["nw"])) if inv["nw"] is not None else None,
                gw=float(r2(inv["gw"])) if inv["gw"] is not None else None,
                consignee_ok=find_consignee(consignees, country, sorted(map(str, inv["cus"]))[0]) is not None,
                warnings=warns, groups=describe(no, inv, pack, hide_comp)))
        for base, fwb in fam_wb.items():              # one file per invoice family: SC260460.xlsx = SC260460, -1, -2 ...
            del fwb["__tpl"]; fwb.active = 0
            cut_rows(fwb)
            fwb.save(os.path.join(outdir, f"{base}.xlsx"))
            files.append(f"{base}.xlsx")
        del wb_all["__tpl"]; wb_all.active = 0
        cut_rows(wb_all)
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
        date=" / ".join(all_dates), hide_comp=hide_comp, merge=merge, master_invoices=n_master, invoices=res, by_country=list(by_country.values()),
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
    ap.add_argument("--no-merge", action="store_true",
                    help="do NOT merge cartons: one invoice per master INV NO (default: merge up to 28 lines for UK/SOUTH AFRICA/AUSTRALIA, 34 for other countries)")
    ap.add_argument("--no-content", action="store_true",
                    help="force for ALL countries: keep the content line of the first group only, remove it from the next groups; groups merge by description + HTS code")
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
            import extract as ptm                 # run alone: python invoicing/engine.py
        only = {x.strip() for x in a.invoices.split(",")} if a.invoices else None
        rows, notes = ptm.extract(a.packing, only=only)
        a.master = ptm.write_master(rows, f"{a.out}/MASTER_INVOICE.xlsx")
        log.info("Master built from packing plan: %d PO lines -> %s", len(rows), a.master)
        for n in notes: log.warning("%s", n)
    invs = read_master(a.master)
    if not a.no_merge:
        n0 = len(invs); invs = merge_cartons(invs, True if a.no_content else None)
        log.info("Merge: %d master invoices -> %d invoices", n0, len(invs))
    fam = {}
    fams = family_map(invs)
    for no, inv in invs.items():                    # one workbook per invoice family (SC260460 + -1, -2, -3 ...)
        base = no.split("-")[0]
        if base not in fam:
            fam[base] = openpyxl.load_workbook(a.template); fam[base].active.title = "__tpl"
            build_bkk(fam[base], base, fams[base], pack, date_txt, "INV.BKK", hide_comp=(True if a.no_content else None))
        it, pt = sheet_names(no, False)
        _, n_po, n_grp, q, amt = build(no, inv, a.template, pack, date_txt, a.out, wb=fam[base], inv_title=it, pl_title=pt, hide_comp=(True if a.no_content else None))
        log.info("OK %-12s %2d PO lines, %d groups, qty %d, amount %.2f", no, n_po, n_grp, q, amt)
    for base, fwb in fam.items():
        del fwb["__tpl"]; fwb.active = 0
        cut_rows(fwb)
        fwb.save(f"{a.out}/{base}.xlsx"); log.info("FILE %s -> %s", fwb.sheetnames, f"{a.out}/{base}.xlsx")
    if a.combine:
        wb = openpyxl.load_workbook(a.template); wb.active.title = "__tpl"
        multi = len({n.split("-")[0] for n in invs}) > 1
        done = set()
        for no, inv in invs.items():
            base = no.split("-")[0]
            if base not in done:
                done.add(base)
                build_bkk(wb, base, fams[base], pack, date_txt, f"INV.BKK {base}" if multi else "INV.BKK", hide_comp=(True if a.no_content else None))
            it, pt = sheet_names(no, multi)
            build(no, inv, a.template, pack, date_txt, a.out, wb=wb, inv_title=it, pl_title=pt, hide_comp=(True if a.no_content else None))
        del wb["__tpl"]
        wb.active = 0
        cut_rows(wb)
        path = f"{a.out}/{a.combine}"
        wb.save(path); log.info("COMBINED %s -> %s", wb.sheetnames, path)

if __name__ == "__main__":
    main()
