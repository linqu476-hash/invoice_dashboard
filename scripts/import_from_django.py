#!/usr/bin/env python3
"""Copy the consignees and shipping marks of the old Django db.sqlite3 into this app.

    python scripts/import_from_django.py ../invoice_dashboard/db.sqlite3                  # into the app database (INVOICE_DB_PATH or data/)
    python scripts/import_from_django.py ../invoice_dashboard/db.sqlite3 --seed data/seed.json   # write a seed / backup file instead
"""
import argparse
import json
import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from app_core import store  # noqa: E402

ap = argparse.ArgumentParser()
ap.add_argument("django_db")
ap.add_argument("--seed", metavar="FILE", help="write the data to this JSON file instead of the database (keep it out of public repositories)")
a = ap.parse_args()
if a.seed:
    import os, tempfile
    os.environ["INVOICE_DB_PATH"] = str(Path(tempfile.mkdtemp()) / "tmp.sqlite3")
    store.import_django_db(a.django_db)
    Path(a.seed).write_text(json.dumps(store.export_data(), ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"written: {a.seed}")
else:
    print(store.import_django_db(a.django_db), "->", store.db_path())
