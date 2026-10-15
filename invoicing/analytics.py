"""Analytics over finished runs (V3). Pure Python, no Django: feed it the result.json of each run.

    summarize(runs)   ->  totals, per month, per country, top customers, "Check before sending" kinds

`runs` is a list of dicts:  dict(job_id=..., created="2026-10-15", review=True/False, result=<result.json dict>)
Everything here only READS the results. Nothing is changed, nothing is written.
"""
from collections import OrderedDict
from datetime import date, timedelta

from .engine import country_key, cus_key

# Range choices of the analytics page: key -> days (None = everything)
RANGES = OrderedDict([("30", 30), ("90", 90), ("365", 365), ("all", None)])
RANGE_LABELS = {"30": "Last 30 days", "90": "Last 90 days", "365": "Last 12 months", "all": "All time"}
DEFAULT_RANGE = "90"

# Order matters: the first matching phrase wins. Wording comes from invoicing/engine.py.
WARNING_KINDS = [
    ("no consignee on file", "Consignee missing"),
    ("no shipping marks on file", "Shipping marks missing"),
    ("printed lines, more than the", "Over the line limit"),
]
OTHER_KIND = "Other"


def warning_kind(message):
    """Group an engine warning into one of a few readable kinds (the invoice number and country differ per warning)."""
    m = (message or "").lower()
    for phrase, kind in WARNING_KINDS:
        if phrase in m:
            return kind
    return OTHER_KIND


def range_start(key, today=None):
    """First day (date) included by a range key, or None for 'all' / an unknown key's default handled by the caller."""
    days = RANGES.get(key)
    if days is None:
        return None
    return (today or date.today()) - timedelta(days=days)


def _pct(part, whole):
    return round(100 * part / whole) if whole else 0


def summarize(runs, top=10):
    """Aggregate finished runs. Returns a dict that is safe to hand to a template (no objects, only numbers and strings)."""
    runs = list(runs)
    months, countries, customers, kinds = OrderedDict(), {}, {}, {}
    tot = dict(runs=0, review=0, clean=0, invoices=0, po_lines=0, qty=0, amount=0.0, warnings=0, cartons=0, merged_invoices=0, merged_away=0)

    for run in sorted(runs, key=lambda x: x["created"]):
        r = run["result"]
        t = r.get("totals", {})
        tot["runs"] += 1
        tot["review"] += 1 if run.get("review") else 0
        tot["invoices"] += t.get("invoices", 0)
        tot["po_lines"] += t.get("po_lines", 0)
        tot["qty"] += t.get("qty", 0)
        tot["amount"] += t.get("amount", 0.0)
        tot["warnings"] += t.get("warnings", 0)
        cartons = r.get("master_invoices") or 0
        if cartons and r.get("merge", True):
            tot["cartons"] += cartons
            tot["merged_invoices"] += t.get("invoices", 0)          # invoices of the runs that merged, to compare with their cartons
            tot["merged_away"] += max(0, cartons - t.get("invoices", 0))

        m = months.setdefault(str(run["created"])[:7], dict(month=str(run["created"])[:7], runs=0, invoices=0, qty=0, amount=0.0))
        m["runs"] += 1
        m["invoices"] += t.get("invoices", 0)
        m["qty"] += t.get("qty", 0)
        m["amount"] += t.get("amount", 0.0)

        for inv in r.get("invoices", []):
            ck = country_key(inv.get("country", "")) or "(no country)"
            c = countries.setdefault(ck, dict(country=ck, invoices=0, qty=0, amount=0.0))
            c["invoices"] += 1
            c["qty"] += inv.get("qty", 0)
            c["amount"] += inv.get("amount", 0.0)

            key = cus_key(inv.get("cus", "")) or "(blank)"
            u = customers.setdefault(key, dict(cus=key, countries=set(), invoices=0, qty=0, amount=0.0))
            u["countries"].add(ck)
            u["invoices"] += 1
            u["qty"] += inv.get("qty", 0)
            u["amount"] += inv.get("amount", 0.0)

            for w in inv.get("warnings", []):
                k = warning_kind(w)
                kinds[k] = kinds.get(k, 0) + 1

    tot["clean"] = tot["runs"] - tot["review"]
    tot["amount"] = round(tot["amount"], 2)
    tot["review_pct"] = _pct(tot["review"], tot["runs"])
    tot["avg_invoice"] = round(tot["amount"] / tot["invoices"], 2) if tot["invoices"] else 0.0

    month_rows = [dict(m, amount=round(m["amount"], 2)) for m in months.values()]
    top_month = max([m["amount"] for m in month_rows] + [1])
    for m in month_rows:
        m["pct"] = max(2, round(100 * m["amount"] / top_month)) if m["amount"] else 0

    country_rows = sorted((dict(c, amount=round(c["amount"], 2)) for c in countries.values()), key=lambda c: (-c["amount"], c["country"]))
    top_country = max([c["amount"] for c in country_rows] + [1])
    for c in country_rows:
        c["pct"] = max(2, round(100 * c["amount"] / top_country)) if c["amount"] else 0
        c["share"] = _pct(c["amount"], tot["amount"])

    customer_rows = sorted((dict(u, amount=round(u["amount"], 2), countries=", ".join(sorted(u["countries"]))) for u in customers.values()),
                           key=lambda u: (-u["amount"], u["cus"]))[:top]
    top_customer = max([u["amount"] for u in customer_rows] + [1])
    for u in customer_rows:
        u["pct"] = max(2, round(100 * u["amount"] / top_customer)) if u["amount"] else 0

    kind_rows = sorted((dict(kind=k, count=n) for k, n in kinds.items()), key=lambda k: (-k["count"], k["kind"]))
    top_kind = max([k["count"] for k in kind_rows] + [1])
    for k in kind_rows:
        k["pct"] = max(2, round(100 * k["count"] / top_kind))

    return dict(totals=tot, months=month_rows, countries=country_rows, customers=customer_rows, warning_kinds=kind_rows)
