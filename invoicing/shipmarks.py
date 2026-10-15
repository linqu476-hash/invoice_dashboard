#!/usr/bin/env python3
"""
shipmarks.py  --  CONSIGNEE Excel (.xls / .xlsx, ONE SHEET PER CUSTOMER NO.)  ->  shipping marks per customer number

Layout of every sheet (column A):
    A1  CONSIGNEE :
    A2  ( CUSTOMER NO. 859015 )            <- customer number (the sheet name is used when A2 has none)
    A3  adidas Taiwan Ltd                  \\
    A4  13/F, No.2, Sec.3, Min-Quan E. Rd.  |  these rows are the SHIPPING MARKS, printed on the invoice exactly as written
    A5  Taipei City 104                     |  (name, address, country, ATTN, TEL, FAX)
    ...                                    /
Empty rows and stray numbers are ignored. Reading .xls needs the xlrd package (pip install xlrd); .xlsx needs nothing extra.

Second file type: SHIPPING INSTRUCTION (e.g. SI-400447-DE-all-sample-20230330.xls)
    sheet "overall information":  sold to: 400447  ->  customer number
    block "consignee", column "on customs documents": name / street and street number / ZIP code and city / country
    those 4 rows are saved as the shipping marks of that customer number.
The file type is recognised by the sheet name "overall information"; anything else is read as the CONSIGNEE Excel above.

The same shipping instruction can also be loaded on the CONSIGNEES page (read_si_consignee): name + address of the column
'on customs documents' are saved for the destination country written in its 'country' row.

Usage:  python invoicing/shipmarks.py "CONSIGNEE__19-OCT-2026.xls"      (prints what would be loaded)
"""
import re
import sys
from pathlib import Path

import openpyxl

CUS_RE = re.compile(r"CUSTOMER\s*NO\.?\s*:?\s*(\d+)", re.I)
HEADER_RE = re.compile(r"^CONSIGNEE\s*:?$", re.I)


class MarksError(ValueError):
    pass


def _clean(v):
    """Cell -> text. Non-breaking spaces become spaces, line breaks inside a cell are kept, trailing spaces are removed."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).replace("\xa0", " ").replace("\r\n", "\n").replace("\r", "\n")
    return "\n".join(line.rstrip() for line in s.split("\n")).strip()


def _sheets(path, filename=None):
    """-> [(sheet name, grid)] ; grid = list of rows, each a list of cell values"""
    ext = Path(filename or str(path)).suffix.lower()
    if ext == ".xls":
        try:
            import xlrd
        except ImportError:
            raise MarksError("Reading an .xls file needs the xlrd package: run  pip install xlrd  (or save the file as .xlsx and upload that).")
        try:
            wb = xlrd.open_workbook(str(path))
        except Exception as e:
            raise MarksError(f"Cannot read the .xls file: {e}")
        return [(sh.name, [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]) for sh in wb.sheets()]
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        raise MarksError(f"Cannot read the Excel file: {e}")
    return [(ws.title, [list(row) for row in ws.iter_rows(values_only=True)]) for ws in wb.worksheets]


def parse_sheet(name, cells):
    """-> dict(customer_no, marks, note, sheet) or None when the sheet holds no customer number / no address rows."""
    vals = [_clean(v) for v in cells]
    cus, note, start = None, "", 0
    for i, v in enumerate(vals[:6]):
        m = CUS_RE.search(v)
        if m:
            cus, start = m.group(1), i + 1
            rest = v[m.end():].strip()
            note = (rest[:-1] if rest.endswith(")") else rest).strip()        # e.g. "(Import & re-import via HK)"
            break
    if cus is None:
        if re.fullmatch(r"\d{3,}", name.strip()):                             # no "CUSTOMER NO." row: the sheet name is the number
            cus = name.strip()
        else:
            return None
    lines = [v for v in vals[start:] if v and not v.isdigit() and not HEADER_RE.match(v)]
    if not lines:
        return None
    return dict(customer_no=cus, marks="\n".join(lines), note=note, sheet=name, kind="consignee")


SI_SHEET = "overall information"
SI_COLUMN = "on customs documents"
SI_FIELDS = ("name", "street and street number", "zip code and city", "country")


def _norm(v):
    return " ".join(_clean(v).lower().split())


def _si_customer(grid, filename=None):
    """Customer number of a shipping instruction: the value next to 'sold to' (else taken from the file name SI-400447-...)."""
    cus = None
    for r, row in enumerate(grid):                                         # sold to: 400447
        for c, v in enumerate(row):
            if _norm(v).startswith("sold to"):
                cus = next((re.sub(r"\D", "", _clean(x)) for x in row[c + 1:] if _clean(x)), None) or None
                break
        if cus:
            break
    if not cus and filename:                                               # SI-400447-DE-all-sample-20230330.xls
        m = re.search(r"SI[-_ ]*(\d{4,})", Path(filename).name, re.I)
        cus = m.group(1) if m else None
    return cus


def _si_customs_rows(grid):
    """[(label, text)] of the 'consignee' block in the column 'on customs documents' (labels: name / street and street number /
    ZIP code and city / country, in sheet order), or [] when the block or the column is missing. Found by label text, not cell address."""
    cell = lambda r, c: _clean(grid[r][c]) if r < len(grid) and c < len(grid[r]) else ""
    r_cons = next((r for r, row in enumerate(grid) if any(_norm(v) == "consignee" for v in row)), None)
    if r_cons is None:
        return []
    r_end = next((r for r in range(r_cons + 1, len(grid)) if any(_norm(v).startswith("notify party") for v in grid[r])), len(grid))
    head = next(((r, c) for r in range(r_cons, r_end) for c, v in enumerate(grid[r]) if SI_COLUMN in _norm(v)), None)
    if head is None:
        return []
    col, rows = head[1], []
    for r in range(head[0] + 1, r_end):                                    # label in the first columns, value in the 'customs' column
        label = next((_norm(grid[r][c]) for c in range(min(2, len(grid[r]))) if _norm(grid[r][c])), "")
        if label in SI_FIELDS and cell(r, col):
            rows.append((label, cell(r, col)))
    return rows


def parse_si(name, grid, filename=None):
    """Shipping instruction, sheet 'overall information' -> dict(customer_no, marks, note, sheet, kind) or None.
    Customer number = value next to 'sold to'; marks = the rows name / street and street number / ZIP code and city / country
    of the 'consignee' block, column 'on customs documents'."""
    cus = _si_customer(grid, filename)
    if not cus:
        return None
    rows = _si_customs_rows(grid)
    if not rows:
        return None
    return dict(customer_no=cus, marks="\n".join(v for _, v in rows), note="from shipping instruction", sheet=name, kind="shipping-instruction")


def read_si_consignee(path, filename=None):
    """Shipping instruction -> dict(country, name, address, customer_no, sheet) for the CONSIGNEES page.
    Sheet 'overall information', block 'consignee', column 'on customs documents':
        name                      -> name     (several lines are joined with a space: 'adidas International Trading AG c/o adidas AG')
        street / ZIP+city / country -> address (one line per row, the country is the last line)
        country                   -> country  (the key of the consignee; saved in capitals)"""
    sheets = _sheets(path, filename)
    si = next(((n, g) for n, g in sheets if _norm(n) == SI_SHEET), None)
    if si is None:
        raise MarksError("This is not a Shipping Instruction file: the sheet 'overall information' was not found. "
                         "Sheets in the file: " + ", ".join(n for n, _ in sheets) + ".")
    rows = _si_customs_rows(si[1])
    if not rows:
        raise MarksError("Shipping instruction: no consignee address found. The sheet 'overall information' needs the 'consignee' block "
                         "with the column 'on customs documents' (name, street and street number, ZIP code and city, country).")
    f = {}
    for label, text in rows:
        f.setdefault(label, text)
    country = " ".join(f.get("country", "").split())
    if not country:
        raise MarksError("Shipping instruction: the row 'country' under 'on customs documents' is empty, so the consignee cannot be saved for a country.")
    lines = [f[k] for k in ("street and street number", "zip code and city", "country") if f.get(k)]
    if len(lines) < 2:
        raise MarksError("Shipping instruction: the address under 'on customs documents' is empty (street and street number / ZIP code and city).")
    return dict(country=country, name=" ".join(f.get("name", "").split())[:120], address="\n".join(lines),
                customer_no=_si_customer(si[1], filename) or "", sheet=si[0])


def read_marks(path, filename=None):
    """-> (items, skipped). items = one dict per customer number (a repeated number: the last sheet wins);
    skipped = names of the sheets without a customer number or without address rows.
    A workbook with a sheet 'overall information' is a shipping instruction (customs consignee address, one customer number);
    every other workbook is the CONSIGNEE Excel (one sheet per customer number)."""
    sheets = _sheets(path, filename)
    si = next(((n, g) for n, g in sheets if _norm(n) == SI_SHEET), None)
    if si is not None:
        d = parse_si(si[0], si[1], filename)
        if d is None:
            raise MarksError("Shipping instruction: no consignee address found. The sheet 'overall information' needs 'sold to' (customer number) "
                             "and the consignee block with the column 'on customs documents' (name, street, ZIP code and city, country).")
        return [d], []
    items, skipped = {}, []
    for name, grid in sheets:
        d = parse_sheet(name, [row[0] if row else None for row in grid])
        if d is None:
            skipped.append(name)
            continue
        items[d["customer_no"].lstrip("0") or d["customer_no"]] = d
    if not items:
        raise MarksError("No customer number found. Each sheet needs a row like '( CUSTOMER NO. 859015 )' "
                         "(or the customer number as sheet name) followed by the address rows. "
                         "A shipping instruction file needs the sheet 'overall information'.")
    return list(items.values()), skipped


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    try:
        items, skipped = read_marks(sys.argv[1], sys.argv[1])
    except MarksError as e:
        sys.exit(f"ERROR: {e}")
    for d in items:
        print(f"--- CUSTOMER NO. {d['customer_no']} [{d['kind']}] {d['note']}\n{d['marks']}")
    if skipped:
        print("Skipped sheets:", ", ".join(skipped))


if __name__ == "__main__":
    main()
