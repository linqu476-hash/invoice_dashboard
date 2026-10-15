"""CSV for the All runs export (V3). Pure Python, no Django.

A run's source file name comes from an upload, so it is untrusted text. Excel and LibreOffice run a cell that starts with = + - @
(or a tab / carriage return) as a FORMULA ("CSV injection"). safe_cell() puts a single quote in front of such text so it stays text.
"""
import csv
import io
import numbers

FORMULA_STARTS = ("=", "+", "-", "@", "\t", "\r")
COLUMNS = [("created", "Date"), ("master_name", "Source file"), ("status", "Status"), ("invoices", "Invoices"), ("qty", "Pieces"),
           ("amount", "Amount (USD)"), ("warnings", "To check"), ("rebuild_count", "Rebuilds"), ("job_id", "Run id")]


def safe_cell(value):
    """Numbers stay numbers; text that a spreadsheet would run as a formula gets a leading quote."""
    if isinstance(value, numbers.Number) and not isinstance(value, bool):      # int, float, Decimal (a negative amount is a number, not a formula)
        return value
    s = "" if value is None else str(value)
    return "'" + s if s.startswith(FORMULA_STARTS) else s


def runs_csv(rows):
    """rows: iterable of dicts with the COLUMNS keys. Returns CSV text (UTF-8 BOM is added by the view so Excel reads accents)."""
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow([label for _, label in COLUMNS])
    n = 0
    for r in rows:
        w.writerow([safe_cell(r.get(key, "")) for key, _ in COLUMNS])
        n += 1
    return buf.getvalue(), n
