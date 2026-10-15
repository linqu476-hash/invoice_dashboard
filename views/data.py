"""Backup, restore and migration of the consignees / shipping marks."""
import json

import streamlit as st

from app_core import store

st.title("Backup and data")
st.caption("Consignees and shipping marks are the only data this app keeps. Uploaded workbooks and built invoices are held in your browser "
           "session only and are gone when you sign out or close the tab.")

st.info(f"Database file: `{store.db_path()}`")
st.warning("**Hosts without a persistent disk** (for example Streamlit Community Cloud) wipe local files when the app restarts or is redeployed, and "
           "the consignees and marks saved here go with them. Either keep the database on a persistent disk (set `INVOICE_DB_PATH`), "
           "or download a backup below after every change and restore it when needed, or commit a private `data/seed.json` (see README).")

n_c, n_m = len(store.consignees()), len(store.marks())
c1, c2 = st.columns(2)
c1.metric("Consignees", n_c)
c2.metric("Shipping marks", n_m)

st.subheader("Download a backup")
st.download_button("⬇ Backup (JSON)", json.dumps(store.export_data(), ensure_ascii=False, indent=2).encode("utf-8"),
                   file_name="invoice_dashboard_backup.json", mime="application/json", type="primary")
st.caption("Contains customer addresses: keep it private. The same file can be saved as `data/seed.json` to refill an empty database at startup.")

st.subheader("Restore a backup")
bk = st.file_uploader("Backup file (.json)", type=["json"], key="restore_file")
replace = st.checkbox("Replace everything (delete the current consignees and marks first)", key="restore_replace")
if st.button("Restore", disabled=bk is None):
    try:
        res = store.import_data(json.loads(bk.getvalue().decode("utf-8")), replace=replace)
    except (ValueError, UnicodeDecodeError) as e:
        st.error(f"Could not restore: {e}")
    else:
        st.success(f"Restored {res['consignees']} consignees and {res['shippingmarks']} shipping marks.")

st.subheader("Import from the Django version")
st.caption("Upload the old `db.sqlite3`: its consignees and shipping marks are copied here (merged by customer number). Nothing else is read from it.")
old = st.file_uploader("db.sqlite3", type=["sqlite3", "sqlite", "db"], key="django_db")
if st.button("Import", disabled=old is None):
    import tempfile
    from pathlib import Path
    try:
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "old.sqlite3"
            p.write_bytes(old.getvalue())
            res = store.import_django_db(p)
    except Exception as e:
        st.error(f"Could not import: {e}")
    else:
        st.success(f"Imported {res['consignees']} consignees and {res['shippingmarks']} shipping marks.")
