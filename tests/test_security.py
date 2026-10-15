"""Tests of the three security layers with SYNTHETIC data only.   python -m unittest discover -s tests -v"""
import io
import json
import os
import sqlite3
import sys
import tempfile
import time
import unittest
import zipfile
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
_TMP = tempfile.mkdtemp()
os.environ["INVOICE_DB_PATH"] = str(Path(_TMP) / "sec.sqlite3")
os.environ["INVOICE_LOCKOUT_FILE"] = str(Path(_TMP) / "lock.json")
os.environ["INVOICE_AUDIT_LOG"] = str(Path(_TMP) / "audit.log")
from app_core import authlib, security, store, uploads  # noqa: E402


def xlsx_bytes():
    b = io.BytesIO()
    openpyxl.Workbook().save(b)
    return b.getvalue()


def zip_of(files, compression=zipfile.ZIP_DEFLATED):
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w", compression) as z:
        for n, d in files.items():
            z.writestr(n, d)
    return b.getvalue()


class Layer1SignIn(unittest.TestCase):
    def setUp(self):
        authlib.record_success()

    def test_hash_round_trip(self):
        h = authlib.hash_password("correct horse battery", n=2 ** 10)
        self.assertTrue(h.startswith("scrypt$"))
        self.assertNotIn("correct", h)
        self.assertTrue(authlib.verify_password("correct horse battery", h))
        self.assertFalse(authlib.verify_password("wrong", h))
        self.assertFalse(authlib.verify_password("x", "scrypt$broken"))

    def test_plain_password_still_works(self):
        self.assertTrue(authlib.verify_password("abc", "abc"))
        self.assertFalse(authlib.verify_password("abd", "abc"))

    def test_lockout_after_repeated_failures_and_clears_on_success(self):
        now = time.time()
        for _ in range(authlib.MAX_FAILS - 1):
            self.assertEqual(authlib.record_failure(now), 0)
        self.assertEqual(authlib.seconds_locked(now), 0)
        lock = authlib.record_failure(now)
        self.assertEqual(lock, authlib.BASE_LOCK_SECONDS)
        self.assertGreater(authlib.seconds_locked(now), 0)           # a page refresh does not reset it: it is stored in a file
        self.assertEqual(authlib.seconds_locked(now + lock + 1), 0)
        self.assertEqual(authlib.record_failure(now + lock + 1), lock * 2)   # keeps doubling
        authlib.record_success()
        self.assertEqual(authlib.seconds_locked(), 0)

    def test_old_failures_are_forgotten(self):
        now = time.time()
        for _ in range(authlib.MAX_FAILS - 1):
            authlib.record_failure(now)
        self.assertEqual(authlib.record_failure(now + authlib.FAIL_WINDOW_SECONDS + 5), 0)


class Layer2Uploads(unittest.TestCase):
    def test_real_xlsx_passes(self):
        self.assertTrue(security.check_upload("a.xlsx", xlsx_bytes()))

    def test_renamed_file_is_rejected(self):
        with self.assertRaises(security.UploadRejected):
            security.check_upload("a.xlsx", b"a,b,c\n1,2,3\n")
        with self.assertRaises(security.UploadRejected):
            security.check_upload("a.xls", xlsx_bytes(), kinds=("xls",))

    def test_empty_and_wrong_type(self):
        with self.assertRaises(security.UploadRejected):
            security.check_upload("a.xlsx", b"")
        with self.assertRaises(security.UploadRejected):
            security.check_upload("a.exe", b"MZ....")

    def test_zip_bomb_is_rejected(self):
        bomb = zip_of({"[Content_Types].xml": "x", "big.xml": b"0" * (60 * 1024 * 1024)})
        with self.assertRaises(security.UploadRejected):
            security.check_upload("bomb.xlsx", bomb)

    def test_path_traversal_member_is_rejected(self):
        evil = zip_of({"[Content_Types].xml": "x", "../../etc/passwd": "x"})
        with self.assertRaises(security.UploadRejected):
            security.check_upload("evil.xlsx", evil)

    def test_a_plain_zip_is_not_a_workbook(self):
        with self.assertRaises(security.UploadRejected):
            security.check_upload("a.xlsx", zip_of({"readme.txt": "hi"}))

    def test_si_zip_reports_bad_members_but_keeps_good_ones(self):
        z = zip_of({"good.xlsx": xlsx_bytes(), "fake.xlsx": b"not excel"})
        res = {r[0].split(" / ")[-1]: (r[2], r[3]) for r in uploads.iter_si_uploads([("batch.zip", z)])}
        self.assertIsNotNone(res["good.xlsx"][0])
        self.assertIsNone(res["fake.xlsx"][0])
        self.assertIn("content", res["fake.xlsx"][1])
        # a zip with an unsafe path is refused as a whole, with the reason
        out = list(uploads.iter_si_uploads([("evil.zip", zip_of({"ok.xlsx": xlsx_bytes(), "../up.xlsx": xlsx_bytes()}))]))
        self.assertEqual(len(out), 1)
        self.assertIsNone(out[0][2])
        self.assertIn("unsafe path", out[0][3])

    def test_sqlite_and_json(self):
        p = Path(_TMP) / "x.sqlite3"
        sqlite3.connect(p).close()
        sqlite3.connect(p).execute("create table t(a)").connection.commit()
        self.assertTrue(security.check_upload("x.sqlite3", p.read_bytes(), kinds=("sqlite",)))
        with self.assertRaises(security.UploadRejected):
            security.check_upload("x.sqlite3", b"hello", kinds=("sqlite",))
        self.assertEqual(security.load_json_backup(json.dumps({"consignees": []}).encode()), {"consignees": []})
        for bad in (b"[1,2]", b"{nope", b"\xff\xfe"):
            with self.assertRaises(security.UploadRejected):
                security.load_json_backup(bad)

    def test_formula_text_is_neutralized(self):
        self.assertEqual(security.neutralize_formula("=HYPERLINK(1)"), "'=HYPERLINK(1)")
        self.assertEqual(security.neutralize_formula("12 Main St"), "12 Main St")
        self.assertEqual(security.neutralize_formula("+1 line", strict=True), "'+1 line")
        self.assertEqual(security.neutralize_formula("+1 line"), "+1 line")      # openpyxl stores it as text, so it is left alone

    def test_stored_text_cannot_become_a_formula(self):
        p = Path(os.environ["INVOICE_DB_PATH"])
        p.unlink(missing_ok=True)
        store.save_consignee("777001", "=cmd|' /C calc'!A0", "=1+1\nZug")
        row = store.consignees()[0]
        self.assertFalse(row["name"].startswith("="))
        self.assertFalse(row["address"].startswith("="))
        store.import_data({"shippingmarks": [{"customer_no": "777001", "marks": "=EVIL()"}]})
        self.assertFalse(store.marks()[0]["marks"].startswith("="))


class Layer3Data(unittest.TestCase):
    def test_audit_line_has_no_newlines_and_is_private(self):
        log = Path(os.environ["INVOICE_AUDIT_LOG"])
        log.unlink(missing_ok=True)
        security.audit("login_failed", note="a\nb", locked_seconds=0)
        lines = log.read_text().splitlines()
        self.assertEqual(len(lines), 1)
        self.assertIn("login_failed", lines[0])
        if os.name == "posix":
            self.assertEqual(log.stat().st_mode & 0o777, 0o600)

    def test_database_file_is_private(self):
        p = Path(os.environ["INVOICE_DB_PATH"])
        p.unlink(missing_ok=True)
        store.consignees()
        if os.name == "posix":
            self.assertEqual(p.stat().st_mode & 0o777, 0o600)

    def test_repo_ships_no_secrets(self):
        root = Path(__file__).resolve().parent.parent
        ignore = (root / ".gitignore").read_text()
        for must in (".streamlit/secrets.toml", "data/*.sqlite3", "data/seed.json"):
            self.assertIn(must, ignore)
        dockerignore = (root / ".dockerignore").read_text()
        self.assertIn(".streamlit/secrets.toml", dockerignore)


if __name__ == "__main__":
    unittest.main()
