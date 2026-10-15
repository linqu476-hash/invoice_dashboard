"""One-page-per-section PDF report of a finished run (V3). Pure Python + reportlab, no Django.

    build(result, job_id, path_or_buffer, check_lines=(), diff_lines=())

`result` is the run's result.json. The report is a SUMMARY for people to read and file: totals, how cartons merged, the
"Check before sending" list, countries and one line per invoice with its line gauge. The invoices and packing lists themselves
stay Excel workbooks (they come from the template); this file does not replace them.
"""
from datetime import datetime
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

BRAND = colors.HexColor("#2748D8")
INK = colors.HexColor("#14202E")
MUTED = colors.HexColor("#6F7C8B")
LINE = colors.HexColor("#E1E6EC")
OK, NEAR, OVER = colors.HexColor("#17794F"), colors.HexColor("#8F5B00"), colors.HexColor("#B42F26")


def _t(value):
    """Text that is safe inside a Paragraph and printable with the built-in Helvetica font (Latin-1). Anything else becomes '?'."""
    s = "" if value is None else str(value)
    return escape(s.encode("latin-1", "replace").decode("latin-1"))


def _money(v):
    return f"{float(v or 0):,.2f}"


def gauge_state(lines, limit):
    """'ok' / 'near' (85% or more of the limit) / 'over' - the same rule as the result page."""
    if not limit:
        return "ok"
    if lines > limit:
        return "over"
    return "near" if lines >= 0.85 * limit else "ok"


def _styles():
    base = getSampleStyleSheet()
    return dict(
        title=ParagraphStyle("t", parent=base["Title"], fontName="Helvetica-Bold", fontSize=20, leading=24, alignment=0, textColor=INK, spaceAfter=2),
        meta=ParagraphStyle("m", parent=base["Normal"], fontSize=9, leading=12, textColor=MUTED),
        h2=ParagraphStyle("h", parent=base["Heading2"], fontName="Helvetica-Bold", fontSize=12.5, leading=16, textColor=INK, spaceBefore=14, spaceAfter=6),
        body=ParagraphStyle("b", parent=base["Normal"], fontSize=9.5, leading=13, textColor=INK),
        small=ParagraphStyle("s", parent=base["Normal"], fontSize=8.5, leading=11, textColor=INK),
    )


def _table(data, widths, align_right=(), header=True):
    t = Table(data, colWidths=widths, repeatRows=1 if header else 0)
    style = [("FONT", (0, 0), (-1, -1), "Helvetica", 8.5), ("TEXTCOLOR", (0, 0), (-1, -1), INK), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
             ("LINEBELOW", (0, 0), (-1, -1), 0.4, LINE), ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]
    if header:
        style += [("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8.5), ("TEXTCOLOR", (0, 0), (-1, 0), MUTED), ("LINEBELOW", (0, 0), (-1, 0), 0.8, INK)]
    for c in align_right:
        style.append(("ALIGN", (c, 0), (c, -1), "RIGHT"))
    t.setStyle(TableStyle(style))
    return t


def _footer(job_id):
    def draw(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(MUTED)
        canvas.drawString(18 * mm, 10 * mm, f"Trax Apparel · invoice clearance · run {job_id[:8]}")
        canvas.drawRightString(A4[0] - 18 * mm, 10 * mm, f"Page {doc.page}")
        canvas.restoreState()
    return draw


def build(result, job_id, target, check_lines=(), diff_lines=()):
    """Write the PDF to `target` (a path or a file-like object). Returns the number of invoice rows written."""
    st = _styles()
    r = result
    t = r.get("totals", {})
    story = [Paragraph(_t(r.get("master_name", "Invoice run")), st["title"]),
             Paragraph(f"Run report · invoice date {_t(r.get('date', ''))} · built {_t(r.get('created', ''))} · generated {datetime.now():%Y-%m-%d %H:%M}", st["meta"]),
             Spacer(1, 8)]

    # totals strip
    cells = [("Invoices", f"{t.get('invoices', 0)}"), ("PO lines", f"{t.get('po_lines', 0)}"), ("Pieces", f"{t.get('qty', 0):,}"),
             ("Total amount (USD)", _money(t.get("amount", 0))), ("Countries", f"{t.get('countries', 0)}")]
    strip = Table([[Paragraph(f"<font size=8 color='#6F7C8B'>{_t(k)}</font><br/><font size=15><b>{_t(v)}</b></font>", st["body"]) for k, v in cells]],
                  colWidths=[(A4[0] - 36 * mm) / len(cells)] * len(cells))
    strip.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.6, LINE), ("INNERGRID", (0, 0), (-1, -1), 0.4, LINE), ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8)]))
    story.append(strip)

    if r.get("master_invoices"):
        if r.get("merge", True):
            story += [Spacer(1, 6), Paragraph(f"<b>{r['master_invoices']} cartons became {t.get('invoices', 0)} invoice{'s' if t.get('invoices', 0) != 1 else ''}.</b> "
                                              "Same description and HTS code were combined; woven and knitted are never combined.", st["body"])]
        else:
            story += [Spacer(1, 6), Paragraph("<b>Merging was off.</b> Each carton of the master is its own invoice.", st["body"])]

    # check before sending
    checks = list(check_lines) + list(diff_lines)
    story.append(Paragraph("Check before sending", st["h2"]))
    if checks:
        story += [Paragraph("&bull; " + _t(c), st["small"]) for c in checks]
    else:
        story.append(Paragraph("Nothing to check.", st["body"]))

    # countries
    countries = r.get("by_country", [])
    if countries:
        rows = [["Country", "Invoices", "Pieces", "Amount (USD)"]] + [[_t(c["country"]), c["invoices"], f"{c['qty']:,}", _money(c["amount"])] for c in countries]
        story.append(KeepTogether([Paragraph("By country", st["h2"]), _table(rows, [70 * mm, 25 * mm, 30 * mm, 45 * mm], align_right=(1, 2, 3))]))

    # invoices
    invoices = r.get("invoices", [])
    rows = [["Invoice", "Country", "Customer", "PO", "Pieces", "Amount (USD)", "Lines", "Gauge"]]
    marks = []
    for n, i in enumerate(invoices, start=1):
        lim, ln = i.get("limit") or 0, i.get("lines") or 0
        state = gauge_state(ln, lim)
        rows.append([_t(i["no"]), _t(i.get("country", "")), _t(i.get("cus", "")), i.get("n_po", 0), f"{i.get('qty', 0):,}", _money(i.get("amount", 0)),
                     f"{ln} / {lim}" if lim else str(ln), state.upper()])
        marks.append(("TEXTCOLOR", (7, n), (7, n), {"ok": OK, "near": NEAR, "over": OVER}[state]))
    story.append(Paragraph("Invoices", st["h2"]))
    if invoices:
        tbl = _table(rows, [24 * mm, 36 * mm, 21 * mm, 12 * mm, 18 * mm, 28 * mm, 18 * mm, 15 * mm], align_right=(3, 4, 5, 6))
        tbl.setStyle(TableStyle(marks + [("FONT", (7, 1), (7, -1), "Helvetica-Bold", 8)]))
        story.append(tbl)
    else:
        story.append(Paragraph("This run has no invoices.", st["body"]))

    doc = SimpleDocTemplate(target, pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm, topMargin=16 * mm, bottomMargin=18 * mm,
                            title=f"Run report {r.get('master_name', '')}", author="Trax Apparel invoice clearance")
    doc.build(story, onFirstPage=_footer(job_id), onLaterPages=_footer(job_id))
    return len(invoices)
