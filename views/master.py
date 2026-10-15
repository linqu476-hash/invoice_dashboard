"""Step 1: packing plan (or Pre-sell file) -> MASTER_INVOICE.xlsx, and the Pre-sell column-name transform."""
import pandas as pd
import streamlit as st

from app_core import runner
from invoicing import extract

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.title("Packing plan to master")
tab_master, tab_transform = st.tabs(["Build master invoice", "Transform column names"])

# ------------------------------------------------------------------ build the master
with tab_master:
    st.caption("Upload the packing plan. A **Packing Plan Pre-sell** file is converted to the packing plan layout automatically before the master is built.")
    packing = st.file_uploader("Packing plan (.xlsx)", type=["xlsx"], key="master_packing")
    c1, c2 = st.columns(2)
    sheet = c1.text_input("Sheet name (optional)", help="Default: the first sheet that has the INV.NO. and PO Number headers.")
    skip = c2.text_input("Invoice skip build master invoice (optional)", placeholder="SC260458",
                         help="Also skips its -1, -2 cartons. SC260458-1 skips just that carton. Separate several with commas.")
    with st.expander("Rules applied"):
        st.markdown(f"- An invoice with a total amount below **${extract.MIN_INVOICE_AMOUNT:,.0f}** is not built, except "
                    f"{', '.join(sorted(extract.ALWAYS_BUILD_COUNTRIES))}.\n"
                    f"- A USA invoice with a total G.W below **{extract.MIN_USA_GW_KG:,.0f} KG** is not built.\n"
                    "- Mode **AIR-CC** is always built (the two checks above do not apply).\n"
                    "- Change the limits at the top of `invoicing/extract.py`.")
    if st.button("Build master invoice", type="primary", disabled=packing is None):
        st.session_state.pop("master_run", None)
        try:
            with st.spinner("Reading the packing plan ..."):
                st.session_state["master_run"] = runner.build_master(packing.getvalue(), packing.name, sheet.strip() or None, skip)
        except Exception as e:
            st.error(f"Could not build the master invoice: {e}")

    m = st.session_state.get("master_run")
    if m:
        st.divider()
        cols = st.columns(4)
        for col, (label, value) in zip(cols, [("PO lines", f"{m['rows']:,}"), ("Pieces", f"{m['qty']:,}"),
                                              ("Amount (USD)", f"{m['amount']:,.2f}"), ("Invoices / cartons", f"{len(m['summary'])} ({m['splits']} split)")]):
            col.metric(label, value)
        if m["converted"]:
            st.info("Pre-sell packing plan detected: converted to the packing plan layout before the master was built.")
        if m["skipped"]:
            st.caption("Skipped on request: " + ", ".join(m["skipped"]))
        d1, d2, d3 = st.columns(3)
        d1.download_button("⬇ MASTER_INVOICE.xlsx", m["master"], file_name="MASTER_INVOICE.xlsx", mime=XLSX, type="primary", key="dl_master")
        if m["converted"]:
            d2.download_button("⬇ Converted packing plan", m["converted"], file_name="PACKING_PLAN_CONVERTED.xlsx", mime=XLSX, key="dl_converted")
        if d3.button("Build the invoices from this master →"):
            st.session_state["master_handoff"] = dict(data=m["master"], source=f"MASTER_INVOICE ({m['source']})")
            st.session_state.pop("run", None)
            st.switch_page("views/build.py")
        notes = [n for n in m["notes"] if not n.startswith("Pre-sell packing plan detected")]
        if notes:
            with st.expander(f"To check ({len(notes)})", expanded=True):
                for n in notes:
                    st.write(f"- {n}")
        if m["low"]:
            with st.expander(f"Skipped as too small ({len(m['low'])})"):
                st.dataframe(pd.DataFrame([{"Invoice": k, "Total (USD)": v.get("total"), "Country": v.get("country")} for k, v in m["low"].items()]),
                             hide_index=True)
        st.markdown("##### Invoices in the master")
        st.dataframe(pd.DataFrame(m["summary"]).rename(columns={"inv": "Invoice", "country": "Country", "mode": "Mode", "po": "PO lines",
                                                                 "qty": "Pieces", "amount": "Amount (USD)", "nw": "N.W", "gw": "G.W"}),
                     hide_index=True, column_config={"Amount (USD)": st.column_config.NumberColumn(format="%.2f")})

# ------------------------------------------------------------------ transform column names only
with tab_transform:
    st.caption("Only rename the columns of a **Pre-sell** file to the packing plan names (CUSTNOCD → Gps Customer Number, HTS CODE: → HTS CODE1:, "
               "an empty Order Type column added). No master is built.")
    pre = st.file_uploader("Pre-sell packing plan (.xlsx)", type=["xlsx"], key="transform_file")
    tsheet = st.text_input("Sheet name (optional)", key="transform_sheet")
    if st.button("Transform column names", disabled=pre is None):
        st.session_state.pop("transform_run", None)
        try:
            with st.spinner("Converting ..."):
                st.session_state["transform_run"] = dict(source=pre.name, **runner.transform_presell(pre.getvalue(), pre.name, tsheet.strip() or None))
        except Exception as e:
            st.error(f"Could not transform the column names: {e}")
    tr = st.session_state.get("transform_run")
    if tr:
        if not tr["report"]:
            st.info("This file already has the packing plan column names. Nothing to transform.")
        else:
            st.success(f"{tr['source']}: {len(tr['report'])} columns checked.")
            st.download_button("⬇ PACKING_PLAN_CONVERTED.xlsx", tr["converted"], file_name="PACKING_PLAN_CONVERTED.xlsx", mime=XLSX, type="primary", key="dl_tr")
            air = next((x for x in tr["modes"] if x["must_build"]), None)
            if air:
                st.info(f"Mode AIR-CC found: {air['invoices']} invoice(s), {air['lines']} lines. These are always built into the master.")
            for n in tr["notes"]:
                st.write(f"- {n}")
            if tr["modes"]:
                st.markdown("##### Shipment methods found")
                st.dataframe(pd.DataFrame(tr["modes"]).rename(columns={"mode": "Mode", "invoices": "Invoices", "lines": "Lines", "must_build": "Always built"}), hide_index=True)
            st.markdown("##### Column report")
            st.dataframe(pd.DataFrame(tr["report"]), hide_index=True)
