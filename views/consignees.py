"""Consignees per customer number: list, add / edit / delete, Shipping Instruction upload (many files or a zip)."""
import pandas as pd
import streamlit as st

from app_core import runner, store, uploads
from app_core.ui import apply_reset, flash, reset_choice, show_flash

st.title("Consignees")
st.caption("One entry per **customer number** (e.g. 400447). An invoice gets a consignee only when its CUS NO is the same as the customer number "
           "of the entry (000035 = 35 = 35.0); the country of the address does not matter. **United Kingdom** keeps the consignee address of the "
           "Excel template when it has no entry of its own. Without a match the invoice shows the yellow placeholder.")

show_flash()
run = st.session_state.get("run")
missing = runner.missing_consignees(run["result"]) if run else []
fixable = runner.fixable(run["result"]) if run else []

if run and (missing or fixable):
    with st.container(border=True):
        st.markdown(f"**From the last run ({run['result']['master_name']})**")
        if missing:
            st.warning("Customer numbers without a consignee: " + "; ".join(f"{m['customer_no']} ({m['invoices']})" for m in missing))
        if fixable:
            st.info(f"{len(fixable)} invoice(s) printed the placeholder but have a consignee now.")
            if st.button("Update invoices", type="primary", key="cons_rebuild"):
                with st.spinner("Building the invoices again (this can take a minute) ..."):
                    try:
                        st.session_state["run"] = runner.rebuild(run)
                    except Exception as e:
                        st.error(f"Could not build the invoices again: {e}")
                    else:
                        flash("Invoices rebuilt. Open **Build invoices** to download them.")
                        st.rerun()

for group in store.conflicts("consignee"):
    st.error("Duplicate customer number written differently: " + " / ".join(group) + ". Only one can match an invoice. Delete one of them.")

tab_list, tab_edit, tab_upload = st.tabs(["Saved consignees", "Add / edit", "Upload Shipping Instructions"])
items = store.consignees()

# ------------------------------------------------------------------ list
with tab_list:
    if items:
        st.dataframe(pd.DataFrame([{"Customer no.": c["customer_no"], "Name": c["name"], "Address": c["address"].replace("\n", ", "),
                                    "Country": c["country"]} for c in items]), hide_index=True)
    else:
        st.info("No consignees yet. Upload Shipping Instructions, add one by hand, or restore a backup on **Backup and data**.")

# ------------------------------------------------------------------ add / edit / delete
with tab_edit:
    labels = {"new": "➕ New consignee"}
    for c in items:
        labels[c["id"]] = f"{c['customer_no'] or '(no customer number)'} - {c['name'] or c['address'][:30]}"
    apply_reset("cons_choice")
    choice = st.selectbox("Entry", list(labels), format_func=labels.get, key="cons_choice")
    cur = next((c for c in items if c["id"] == choice), None)
    prefill = ""
    if cur is None and missing:
        pick = st.selectbox("Fill in a customer number missing from the last run", ["-"] + [m["customer_no"] for m in missing], key="cons_missing_pick")
        prefill = "" if pick == "-" else pick
    kid = f"{choice}_{prefill}"
    with st.form(f"cons_form_{kid}"):
        cus = st.text_input("Customer number", value=cur["customer_no"] if cur else prefill, placeholder="400447")
        name = st.text_input("Name", value=cur["name"] if cur else "", help="Printed above the address.")
        addr = st.text_area("Address", value=cur["address"] if cur else "", height=160, help="Printed under CONSIGNEE, one line per row.")
        saved = st.form_submit_button("Save", type="primary")
    if saved:
        try:
            store.save_consignee(cus, name, addr, pk=cur["id"] if cur else None)
        except ValueError as e:
            st.error(str(e))
        else:
            flash("Saved.")
            if cur is None:
                reset_choice("cons_choice")
            st.rerun()
    if cur is not None:
        st.divider()
        sure = st.checkbox(f"Yes, delete {cur['customer_no'] or 'this entry'}", key=f"cons_del_ok_{cur['id']}")
        if st.button("Delete consignee", disabled=not sure, key=f"cons_del_{cur['id']}"):
            store.delete_consignee(cur["id"])
            reset_choice("cons_choice")
            flash("Deleted.")
            st.rerun()

# ------------------------------------------------------------------ Shipping Instruction upload
with tab_upload:
    st.caption("Many files at once (Ctrl / Shift + click) or one **.zip** (up to 500 files). Sheet `overall information`, block `consignee`, column "
               "`on customs documents`: name, street, ZIP code and city, country. The number next to `sold to` is the customer number. "
               "The same customer number is **replaced** (its shipping marks stay). A file without a customer number is refused.")
    files = st.file_uploader("Shipping Instruction files (.xls, .xlsx or .zip)", type=["xls", "xlsx", "zip"], accept_multiple_files=True, key="si_files")
    if st.button("Read and save consignees", type="primary", disabled=not files):
        with st.spinner(f"Reading {len(files)} upload(s) ..."):
            res = uploads.save_si_uploads([(f.name, f.getvalue()) for f in files])
        st.session_state["si_result"] = res
        st.rerun()
    res = st.session_state.get("si_result")
    if res:
        st.success(f"{res['total']} file(s): {res['added']} added, {res['replaced']} replaced, {res['unchanged']} unchanged, {res['failed']} failed.")
        st.dataframe(pd.DataFrame([{"File": r["file"], "Result": r["status"], "Customer no.": r["customer_no"], "Country": r["country"],
                                    "Name": r["name"], "Problem": r["error"]} for r in res["rows"]]), hide_index=True)
