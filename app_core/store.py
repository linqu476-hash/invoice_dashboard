"""Consignees and shipping marks, stored in one small SQLite file (replaces the Django models).

No Streamlit and no Django in here, so it can be tested on its own. The rules are the ones of the Django app:
  * a consignee is found by customer number only (000035 = 35 = 35.0); the country is information only
  * UNITED KINGDOM may have no customer number (it keeps the template address)
  * the same customer number is REPLACED, never duplicated
"""
import json
import os
import sqlite3
import tempfile
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path

from invoicing import engine, shipmarks

ROOT = Path(__file__).resolve().parent.parent
SEED_FILE = ROOT / "data" / "seed.json"

SCHEMA = """
CREATE TABLE IF NOT EXISTS consignee (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    country TEXT NOT NULL DEFAULT '',
    customer_no TEXT NOT NULL DEFAULT '',
    name TEXT NOT NULL DEFAULT '',
    address TEXT NOT NULL DEFAULT '',
    marks TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS shippingmark (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    customer_no TEXT NOT NULL UNIQUE,
    marks TEXT NOT NULL DEFAULT '',
    note TEXT NOT NULL DEFAULT '',
    source TEXT NOT NULL DEFAULT '',
    updated TEXT NOT NULL DEFAULT ''
);
"""


def db_path():
    p = os.environ.get("INVOICE_DB_PATH")
    return Path(p) if p else ROOT / "data" / "invoice_dashboard.sqlite3"


@contextmanager
def _conn():
    path = db_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    c = sqlite3.connect(path, timeout=15)
    c.row_factory = sqlite3.Row
    try:
        c.executescript(SCHEMA)
        yield c
        c.commit()
    except Exception:
        c.rollback()
        raise
    finally:
        c.close()


def _now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _rows(sql, args=()):
    with _conn() as c:
        return [dict(r) for r in c.execute(sql, args)]


# ---------------------------------------------------------------- consignees
def consignees():
    return _rows("SELECT * FROM consignee ORDER BY country, customer_no, id")


def consignee_table():
    """{customer key: dict(name, address, marks)} for engine.process (same shape as the Django _consignees())."""
    out = {}
    for c in _rows("SELECT * FROM consignee ORDER BY id"):
        if c["customer_no"].strip():
            key = engine.cus_key(c["customer_no"])
        elif engine.country_key(c["country"]) == "UNITED KINGDOM":
            key = "UNITED KINGDOM"
        else:
            continue                      # no customer number = never matches an invoice
        out[key] = dict(name=c["name"], address=c["address"], marks=c["marks"])
    return out


def marks_table():
    """{customer key: marks text} for engine.process."""
    return {engine.cus_key(m["customer_no"]): m["marks"] for m in _rows("SELECT * FROM shippingmark") if m["marks"].strip()}


def save_consignee(customer_no, name, address, pk=None):
    """Add (pk None) or edit one consignee by hand. Raises ValueError with the same rules as the Django form."""
    cus, name, address = (customer_no or "").strip(), (name or "").strip(), (address or "").strip()
    with _conn() as c:
        old = c.execute("SELECT * FROM consignee WHERE id=?", (pk,)).fetchone() if pk else None
        uk = engine.country_key(old["country"] if old else "") == "UNITED KINGDOM"
        if not cus and not uk:
            raise ValueError("Customer number is required: an invoice only gets this consignee when its CUS NO is the same as this customer number.")
        if cus:
            key = engine.cus_key(cus)
            for r in c.execute("SELECT id, customer_no FROM consignee"):
                if r["id"] != pk and r["customer_no"].strip() and engine.cus_key(r["customer_no"]) == key:
                    raise ValueError(f"Customer number {r['customer_no']} already has a consignee. Edit that entry instead.")
        if not uk and not address:
            raise ValueError("Address is required. Shipping marks are loaded per customer number on the Shipping marks page.")
        if old:
            c.execute("UPDATE consignee SET customer_no=?, name=?, address=? WHERE id=?", (cus, name, address, pk))
            return old["id"]
        return c.execute("INSERT INTO consignee(country, customer_no, name, address) VALUES('', ?, ?, ?)", (cus, name, address)).lastrowid


def save_si_consignee(item):
    """Save one consignee read from a Shipping Instruction -> (status, row). Found by customer number only;
    the same customer number is REPLACED (its shipping marks stay), whatever country the address has."""
    cus = (item.get("customer_no") or "").strip()
    key = engine.cus_key(cus) if cus else ""
    if not key:
        raise shipmarks.MarksError("Shipping instruction: no customer number found next to 'sold to' (or in the file name), so the consignee "
                                   "cannot be matched to an invoice.")
    country = (item["country"] or "").strip().upper()
    with _conn() as c:
        rows = [dict(r) for r in c.execute("SELECT * FROM consignee ORDER BY id")]
        cur = next((r for r in rows if r["customer_no"].strip() and engine.cus_key(r["customer_no"]) == key), None)
        if cur is None:               # older entry without customer number, same data -> it gets the number
            cur = next((r for r in rows if not r["customer_no"].strip() and engine.country_key(r["country"]) != "UNITED KINGDOM"
                        and engine.country_key(r["country"]) == engine.country_key(country)
                        and r["name"] == item["name"] and r["address"] == item["address"]), None)
        if cur is None:
            c.execute("INSERT INTO consignee(country, customer_no, name, address) VALUES(?,?,?,?)", (country, cus, item["name"], item["address"]))
            status = "added"
        else:
            same = cur["customer_no"].strip() == cus and cur["name"] == item["name"] and cur["address"] == item["address"]
            status = "unchanged" if same else "replaced"
            c.execute("UPDATE consignee SET country=?, customer_no=?, name=?, address=? WHERE id=?", (country, cus, item["name"], item["address"], cur["id"]))
    return status, dict(country=country, customer_no=cus, name=item["name"])


def delete_consignee(pk):
    with _conn() as c:
        c.execute("DELETE FROM consignee WHERE id=?", (pk,))


# ---------------------------------------------------------------- shipping marks
def marks():
    return _rows("SELECT * FROM shippingmark ORDER BY customer_no")


def save_mark(customer_no, marks_text, note="", source="", pk=None):
    """Add or edit the shipping marks of one customer number. Adding a number that exists replaces its marks (as the Django page did)."""
    cus = (customer_no or "").strip()
    if not cus:
        raise ValueError("Customer number is required.")
    if not (marks_text or "").strip():
        raise ValueError("Shipping marks are required (one line per row).")
    key = engine.cus_key(cus)
    with _conn() as c:
        rows = [dict(r) for r in c.execute("SELECT id, customer_no FROM shippingmark")]
        same = next((r for r in rows if engine.cus_key(r["customer_no"]) == key), None)
        if pk and same and same["id"] != pk:
            raise ValueError(f"Customer number {same['customer_no']} already has shipping marks.")
        target = pk or (same["id"] if same else None)
        if target:
            c.execute("UPDATE shippingmark SET customer_no=?, marks=?, note=?, source=?, updated=? WHERE id=?",
                      (cus, marks_text, note or "", source or "", _now(), target))
        else:
            c.execute("INSERT INTO shippingmark(customer_no, marks, note, source, updated) VALUES(?,?,?,?,?)",
                      (cus, marks_text, note or "", source or "", _now()))


def save_marks_items(items, source):
    """Items read from the CONSIGNEE Excel (shipmarks.read_marks). A customer number already saved is REPLACED.
    -> list of dict(no, status)."""
    done = []
    with _conn() as c:
        existing = {engine.cus_key(m["customer_no"]): dict(m) for m in c.execute("SELECT * FROM shippingmark")}
        for d in items:
            m = existing.get(engine.cus_key(d["customer_no"]))
            if m is None:
                status = "added"
                c.execute("INSERT INTO shippingmark(customer_no, marks, note, source, updated) VALUES(?,?,?,?,?)",
                          (d["customer_no"], d["marks"], d.get("note", ""), source[:200], _now()))
            else:
                status = "unchanged" if (m["marks"] == d["marks"] and m["note"] == d.get("note", "")) else "replaced"
                c.execute("UPDATE shippingmark SET customer_no=?, marks=?, note=?, source=?, updated=? WHERE id=?",
                          (d["customer_no"], d["marks"], d.get("note", ""), source[:200], _now(), m["id"]))
            done.append(dict(no=d["customer_no"], status=status))
    return done


def delete_mark(pk):
    with _conn() as c:
        c.execute("DELETE FROM shippingmark WHERE id=?", (pk,))


def conflicts(kind):
    """Customer numbers that are the same number written differently (000035 / 35): lists of the spellings. kind: 'consignee' | 'shippingmark'."""
    assert kind in ("consignee", "shippingmark")
    groups = {}
    for o in _rows(f"SELECT customer_no FROM {kind}"):
        k = engine.cus_key(o["customer_no"])
        if k:
            groups.setdefault(k, []).append(o["customer_no"])
    return [sorted(v) for v in groups.values() if len(v) > 1]


# ---------------------------------------------------------------- reading Excel uploads
def read_si_bytes(filename, data):
    """Shipping Instruction workbook (bytes) -> consignee item dict. Raises shipmarks.MarksError for a file that is not usable."""
    with tempfile.TemporaryDirectory() as tmp:
        p = Path(tmp) / ("si" + Path(filename).suffix.lower())
        p.write_bytes(data)
        return shipmarks.read_si_consignee(p, filename)


def read_marks_bytes(filename, data):
    with tempfile.TemporaryDirectory() as tmp:
        p = Path(tmp) / ("consignee" + Path(filename).suffix.lower())
        p.write_bytes(data)
        return shipmarks.read_marks(p, filename)


# ---------------------------------------------------------------- backup / restore / migration
def export_data():
    return dict(version=1, exported=_now(), consignees=[{k: v for k, v in c.items() if k != "id"} for c in consignees()],
                shippingmarks=[{k: v for k, v in m.items() if k != "id"} for m in marks()])


def import_data(data, replace=False):
    """Load an export (see export_data). replace=True empties both tables first, else rows are merged by customer number.
    -> dict(consignees=n, shippingmarks=n)"""
    if not isinstance(data, dict) or not ("consignees" in data or "shippingmarks" in data):
        raise ValueError("This is not an export of this app (expected the keys 'consignees' / 'shippingmarks').")
    n_c = n_m = 0
    with _conn() as c:
        if replace:
            c.execute("DELETE FROM consignee")
            c.execute("DELETE FROM shippingmark")
        for r in data.get("consignees", []):
            cus = str(r.get("customer_no", "") or "").strip()
            args = (str(r.get("country", "") or "").strip().upper(), cus, str(r.get("name", "") or ""), str(r.get("address", "") or ""), str(r.get("marks", "") or ""))
            cur = None
            if cus:
                cur = next((x for x in c.execute("SELECT id, customer_no FROM consignee") if x["customer_no"].strip() and engine.cus_key(x["customer_no"]) == engine.cus_key(cus)), None)
            elif engine.country_key(args[0]) == "UNITED KINGDOM":
                cur = next((x for x in c.execute("SELECT id, customer_no, country FROM consignee") if not x["customer_no"].strip()
                            and engine.country_key(x["country"]) == "UNITED KINGDOM"), None)
            if cur:
                c.execute("UPDATE consignee SET country=?, customer_no=?, name=?, address=?, marks=? WHERE id=?", args + (cur["id"],))
            else:
                c.execute("INSERT INTO consignee(country, customer_no, name, address, marks) VALUES(?,?,?,?,?)", args)
            n_c += 1
        for m in data.get("shippingmarks", []):
            cus = str(m.get("customer_no", "") or "").strip()
            if not cus or not str(m.get("marks", "") or "").strip():
                continue
            cur = next((x for x in c.execute("SELECT id, customer_no FROM shippingmark") if engine.cus_key(x["customer_no"]) == engine.cus_key(cus)), None)
            args = (cus, str(m["marks"]), str(m.get("note", "") or ""), str(m.get("source", "") or ""), str(m.get("updated", "") or _now()))
            if cur:
                c.execute("UPDATE shippingmark SET customer_no=?, marks=?, note=?, source=?, updated=? WHERE id=?", args + (cur["id"],))
            else:
                c.execute("INSERT INTO shippingmark(customer_no, marks, note, source, updated) VALUES(?,?,?,?,?)", args)
            n_m += 1
    return dict(consignees=n_c, shippingmarks=n_m)


def import_django_db(path):
    """Read consignees and shipping marks out of the old Django db.sqlite3 (tables dashboard_consignee / dashboard_shippingmark)."""
    src = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    src.row_factory = sqlite3.Row
    try:
        have = {r[0] for r in src.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        if not {"dashboard_consignee", "dashboard_shippingmark"} & have:
            raise ValueError("This is not the Django db.sqlite3 of the invoice dashboard (tables dashboard_consignee / dashboard_shippingmark not found).")
        data = dict(consignees=[], shippingmarks=[])
        if "dashboard_consignee" in have:
            data["consignees"] = [dict(country=r["country"], customer_no=r["customer_no"], name=r["name"], address=r["address"], marks=r["marks"])
                                  for r in src.execute("SELECT * FROM dashboard_consignee ORDER BY id")]
        if "dashboard_shippingmark" in have:
            data["shippingmarks"] = [dict(customer_no=r["customer_no"], marks=r["marks"], note=r["note"], source=r["source"], updated=str(r["updated"] or ""))
                                     for r in src.execute("SELECT * FROM dashboard_shippingmark ORDER BY id")]
    finally:
        src.close()
    return import_data(data)


def seed_if_empty():
    """On a host whose disk is wiped at every restart: load data/seed.json (an export of this app) when the tables are empty.
    Keep that file out of any public repository (it holds customer addresses)."""
    if not SEED_FILE.exists():
        return None
    with _conn() as c:
        empty = c.execute("SELECT (SELECT COUNT(*) FROM consignee) + (SELECT COUNT(*) FROM shippingmark)").fetchone()[0] == 0
    if not empty:
        return None
    return import_data(json.loads(SEED_FILE.read_text(encoding="utf-8")))
