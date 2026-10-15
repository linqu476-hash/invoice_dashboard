"""Tests of the service layer with SYNTHETIC data only (no customer data).   python -m unittest discover -s tests -v"""
import datetime
import io
import json
import os
import sqlite3
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from invoicing import extract, shipmarks  # noqa: E402

_TMP = tempfile.mkdtemp()
os.environ["INVOICE_DB_PATH"] = str(Path(_TMP) / "test.sqlite3")
from app_core import runner, store, uploads  # noqa: E402


def fresh_db():
    p = Path(os.environ["INVOICE_DB_PATH"])
    if p.exists():
        p.unlink()


def make_master(country="UNITED KINGDOM", cus="999001", date=datetime.datetime(2026, 10, 19)):
    rows = [dict(inv="SC900001", date=date, po=4500000001 + i, cpo="AA1", art=f"KU{i}", cus=cus, country=country,
                 desc="UNISEX SHORTS 100%POLYESTER(100%RECYCLED) KNITTED", hts=6104630000, amount=20.0 + i, qty=2,
                 nw=1.3 if i == 0 else None, gw=1.85 if i == 0 else None, mode="Air Express") for i in range(2)]
    p = Path(_TMP) / "m.xlsx"
    extract.write_master(rows, p)
    return p.read_bytes()


def make_si(cus, name="Acme AG", street="Platz 1", city="1234 Zug", country="Switzerland"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "overall information"
    for row in (["sold to", cus], [], ["consignee", None, "on invoice", "on customs documents"], ["name", None, "x", name],
                ["street and street number", None, "x", street], ["ZIP code and city", None, "x", city], ["country", None, "x", country], ["notify party"]):
        ws.append(row)
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


class StoreTests(unittest.TestCase):
    def setUp(self):
        fresh_db()

    def test_manual_rules(self):
        with self.assertRaises(ValueError):
            store.save_consignee("", "n", "addr")            # customer number required
        with self.assertRaises(ValueError):
            store.save_consignee("35", "n", "")              # address required
        store.save_consignee("35", "Acme", "1 Road")
        with self.assertRaises(ValueError):
            store.save_consignee("000035", "x", "y")         # 000035 = 35
        self.assertIn("35", store.consignee_table())

    def test_si_added_unchanged_replaced(self):
        item = store.read_si_bytes("SI-400447.xlsx", make_si("400447"))
        self.assertEqual(store.save_si_consignee(item)[0], "added")
        self.assertEqual(store.save_si_consignee(item)[0], "unchanged")
        self.assertEqual(store.save_si_consignee(dict(item, name="NEW"))[0], "replaced")
        self.assertEqual(len(store.consignees()), 1)
        self.assertEqual(store.consignees()[0]["country"], "SWITZERLAND")

    def test_si_without_customer_number_is_refused(self):
        with self.assertRaises(shipmarks.MarksError):
            store.save_si_consignee(dict(country="X", name="n", address="a", customer_no=""))

    def test_marks_replace_and_conflicts(self):
        store.save_mark("859015", "ACME\nBERLIN")
        store.save_mark("0859015", "ACME2")                   # same number: replaced, not duplicated
        self.assertEqual(len(store.marks()), 1)
        self.assertEqual(store.marks_table()["859015"], "ACME2")
        with self.assertRaises(ValueError):
            store.save_mark("", "x")

    def test_backup_roundtrip(self):
        store.save_consignee("35", "Acme", "1 Road")
        store.save_mark("35", "MARKS")
        data = json.loads(json.dumps(store.export_data()))
        store.import_data(data, replace=True)
        self.assertEqual((len(store.consignees()), len(store.marks())), (1, 1))
        store.import_data(data)                               # merge: no duplicates
        self.assertEqual((len(store.consignees()), len(store.marks())), (1, 1))
        with self.assertRaises(ValueError):
            store.import_data({"nope": 1})

    def test_import_django_db(self):
        p = Path(_TMP) / "old.sqlite3"
        p.unlink(missing_ok=True)
        c = sqlite3.connect(p)
        c.executescript("CREATE TABLE dashboard_consignee(id INTEGER PRIMARY KEY, country TEXT, customer_no TEXT, name TEXT, address TEXT, marks TEXT);"
                        "CREATE TABLE dashboard_shippingmark(id INTEGER PRIMARY KEY, customer_no TEXT, marks TEXT, note TEXT, source TEXT, updated TEXT);"
                        "INSERT INTO dashboard_consignee VALUES(1,'DENMARK','400447','N','A','');"
                        "INSERT INTO dashboard_shippingmark VALUES(1,'859015','M','','f.xls','2026-10-15');")
        c.commit()
        c.close()
        self.assertEqual(store.import_django_db(p), dict(consignees=1, shippingmarks=1))
        with self.assertRaises(ValueError):
            c = sqlite3.connect(Path(_TMP) / "other.sqlite3")
            c.execute("CREATE TABLE x(a)")
            c.commit()
            c.close()
            store.import_django_db(Path(_TMP) / "other.sqlite3")


class UploadTests(unittest.TestCase):
    def setUp(self):
        fresh_db()

    def test_zip_and_plain_and_bad_files(self):
        z = io.BytesIO()
        with zipfile.ZipFile(z, "w") as zf:
            zf.writestr("a/SI-111111.xlsx", make_si("111111"))
            zf.writestr("__MACOSX/._SI-111111.xlsx", b"x")
            zf.writestr("readme.txt", b"hi")
        res = uploads.save_si_uploads([("pack.zip", z.getvalue()), ("SI-222222.xlsx", make_si("222222")),
                                       ("notes.pdf", b"x"), ("bad.xlsx", b"not excel"), ("empty.zip", b"nope")])
        self.assertEqual((res["added"], res["failed"], res["total"]), (2, 3, 5))
        self.assertEqual(len(store.consignees()), 2)


class PipelineTests(unittest.TestCase):
    def setUp(self):
        fresh_db()

    def test_build_missing_then_fixable_then_rebuild(self):
        run = runner.build_invoices(make_master(country="GERMANY", cus="999001"), "m.xlsx")
        r = run["result"]
        self.assertEqual(r["totals"]["invoices"], 1)
        self.assertTrue(run["report_pdf"].startswith(b"%PDF"))
        self.assertIn(r["files"]["zip"], run["files"])
        self.assertEqual([m["customer_no"] for m in runner.missing_consignees(r)], ["999001"])
        self.assertEqual(runner.fixable(r), [])
        store.save_consignee("999001", "Test Co", "1 Street\nBerlin")
        store.save_mark("999001", "TEST MARKS")
        self.assertEqual(runner.fixable(r), ["SC900001"])
        run2 = runner.rebuild(run)
        self.assertTrue(run2["result"]["invoices"][0]["consignee_ok"])
        self.assertEqual(runner.missing_consignees(run2["result"]), [])

    def test_invalid_master_lists_problems_and_builds_nothing(self):
        wb = openpyxl.Workbook()
        wb.active.append(["INV NO", "PO NO"])                 # required columns missing
        b = io.BytesIO()
        wb.save(b)
        with self.assertRaises(runner.InvalidMaster) as cm:
            runner.build_invoices(b.getvalue(), "bad.xlsx")
        self.assertTrue(cm.exception.problems)

    def test_not_an_excel_file(self):
        with self.assertRaises(runner.InvalidMaster):
            runner.build_invoices(b"a,b,c\n1,2,3\n", "renamed.xlsx")

    def test_size_limit(self):
        with self.assertRaises(ValueError):
            runner.build_invoices(b"x" * (runner.MAX_BYTES + 1), "huge.xlsx")


class FastMergeTests(unittest.TestCase):
    """engine._merge / _unmerge / _copy_sheet only skip work: the sheet must be the same as with the openpyxl methods
    (this also catches an openpyxl upgrade that changes the internals they rely on)."""

    @staticmethod
    def _state(ws):
        cells = []
        for (r, c), x in sorted(ws._cells.items()):
            b = x.border
            cells.append((r, c, type(x).__name__, x.value, b.left.style, b.right.style, b.top.style, b.bottom.style, x.protection.locked))
        return cells, sorted(str(m) for m in ws.merged_cells.ranges)

    def test_same_sheet_as_openpyxl(self):
        from invoicing import engine
        slow, fast = (openpyxl.load_workbook(engine_template()) for _ in range(2))
        s_ws = slow.copy_worksheet(slow.active)
        f_ws = engine._copy_sheet(fast, fast.active)
        self.assertEqual(self._state(s_ws), self._state(f_ws))
        for ws, fast_mode in ((s_ws, False), (f_ws, True)):
            old = [m for m in ws.merged_cells.ranges if m.min_row >= 25]
            if fast_mode:
                engine._unmerge(ws, old)
                engine._merge(ws, 30, 30, 3, 18)
                engine._merge(ws, 31, 33, 3, 10)
            else:
                for m in old:
                    ws.unmerge_cells(str(m))
                ws.merge_cells(start_row=30, end_row=30, start_column=3, end_column=18)
                ws.merge_cells(start_row=31, end_row=33, start_column=3, end_column=10)
        self.assertEqual(self._state(s_ws), self._state(f_ws))


def engine_template():
    from app_core.runner import TEMPLATE_XLSX
    return TEMPLATE_XLSX


if __name__ == "__main__":
    unittest.main()
