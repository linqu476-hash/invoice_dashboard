"""Shipping marks per customer number: CONSIGNEE Excel upload, add / edit / delete."""
import pandas as pd
import streamlit as st

from app_core import store
from app_core.ui import apply_reset, flash, reset_choice, show_flash
from invoicing import shipmarks

st.title("Shipping marks")
st.caption("The marks of a customer number are printed on every invoice whose CUS NO matches (INV, PL and INV.BKK). A customer number without "
           "marks gets a yellow placeholder and a warning.")

show_flash()
for group in store.conflicts("shippingmark"):
    st.error("Duplicate customer number written differently: " + " / ".join(group) + ". Only one can match an invoice. Delete one of them.")

tab_list, tab_edit, tab_upload = st.tabs(["Saved marks", "Add / edit", "Upload CONSIGNEE Excel"])
items = store.marks()

with tab_list:
    if items:
        st.dataframe(pd.DataFrame([{"Customer no.": m["customer_no"], "Marks": m["marks"].replace("\n", ", "), "Note": m["note"],
                                    "Source": m["source"], "Updated": m["updated"][:16].replace("T", " ")} for m in items]), hide_index=True)
    else:
        st.info("No shipping marks yet. Upload the CONSIGNEE Excel or add them by hand.")

with tab_edit:
    labels = {"new": "➕ New shipping marks"}
    for m in items:
        labels[m["id"]] = m["customer_no"]
    apply_reset("mark_choice")
    choice = st.selectbox("Entry", list(labels), format_func=labels.get, key="mark_choice")
    cur = next((m for m in items if m["id"] == choice), None)
    with st.form(f"mark_form_{choice}"):
        cus = st.text_input("Customer number", value=cur["customer_no"] if cur else "", placeholder="859015")
        text = st.text_area("Shipping marks", value=cur["marks"] if cur else "", height=160, help="Printed under SHIPPING MARKS, one line per row.")
        note = st.text_input("Note", value=cur["note"] if cur else "")
        saved = st.form_submit_button("Save", type="primary")
    if saved:
        try:
            store.save_mark(cus, text, note, source=cur["source"] if cur else "", pk=cur["id"] if cur else None)
        except ValueError as e:
            st.error(str(e))
        else:
            flash("Saved.")
            reset_choice("mark_choice")
            st.rerun()
    if cur is not None:
        st.divider()
        sure = st.checkbox(f"Yes, delete the marks of {cur['customer_no']}", key=f"mark_del_ok_{cur['id']}")
        if st.button("Delete shipping marks", disabled=not sure, key=f"mark_del_{cur['id']}"):
            store.delete_mark(cur["id"])
            reset_choice("mark_choice")
            flash("Deleted.")
            st.rerun()

with tab_upload:
    st.caption("The CONSIGNEE Excel (`.xls` or `.xlsx`, one sheet per customer number, e.g. `CONSIGNEE__19-OCT-2026.xls`): the rows under "
               "`( CUSTOMER NO. 859015 )` become the marks of that customer number. A customer number already saved is **replaced**. "
               "A single Shipping Instruction file is also accepted here (its customs consignee address becomes the marks).")
    up = st.file_uploader("CONSIGNEE Excel", type=["xls", "xlsx"], key="marks_file")
    if st.button("Read and save marks", type="primary", disabled=up is None):
        try:
            found, skipped = store.read_marks_bytes(up.name, up.getvalue())
            done = store.save_marks_items(found, up.name)
        except shipmarks.MarksError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Could not read the consignee file: {e}")
        else:
            count = lambda s: sum(1 for x in done if x["status"] == s)
            st.success(f"{up.name}: {count('added')} added, {count('replaced')} replaced, {count('unchanged')} unchanged.")
            if skipped:
                st.warning("Sheets skipped (no customer number or no address rows): " + ", ".join(skipped))
            st.dataframe(pd.DataFrame([{"Customer no.": x["no"], "Result": x["status"]} for x in done]), hide_index=True)
