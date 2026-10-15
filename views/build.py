"""Master invoice -> commercial invoices + packing lists, with the result view."""
import pandas as pd
import streamlit as st

from app_core import runner
from invoicing import engine

st.title("Build invoices")
st.caption("Upload MASTER_INVOICE.xlsx (make it on **Packing plan to master**). Each carton after the first one of an invoice becomes "
           "SC260458-1, -2, ... The invoice date is the CRD of the packing plan (master column INV DATE).")


# ------------------------------------------------------------------ result view
def _rebuild(run):
    with st.spinner("Building the invoices again with the consignees saved now (this can take a minute) ..."):
        try:
            st.session_state["run"] = runner.rebuild(run)
        except Exception as e:
            st.error(f"Could not build the invoices again: {e}")
            return
    st.rerun()


def show_run(run):
    r = run["result"]
    t = r["totals"]
    st.divider()
    st.subheader(f"Result: {r['master_name']}")
    if r.get("rebuilt"):
        st.caption(f"Built again on {r['rebuilt']} with the consignees saved at that time.")

    cols = st.columns(6)
    for col, (label, value) in zip(cols, [("Invoices", f"{t['invoices']:,}"), ("PO lines", f"{t['po_lines']:,}"), ("Pieces", f"{t['qty']:,}"),
                                          ("Amount (USD)", f"{t['amount']:,.2f}"), ("Countries", t["countries"]), ("Warnings", t["warnings"])]):
        col.metric(label, value)
    if r.get("merge", True):
        saved = r["master_invoices"] - t["invoices"]
        st.write(f"**{r['master_invoices']} cartons became {t['invoices']} invoice{'s' if t['invoices'] != 1 else ''}**"
                 + (f" (merging saved {saved})." if saved > 0 else "."))
    else:
        st.write("Merging was off: each carton of the master is its own invoice.")
    st.caption(f"Invoice date: {r['date']}")

    # downloads
    st.markdown("#### Downloads")
    d1, d2, d3, d4 = st.columns(4)
    zip_name, comb_name = r["files"]["zip"], r["files"]["combined"]
    d1.download_button("⬇ All invoices (zip)", run["files"][zip_name], file_name=zip_name, mime="application/zip",
                       key=f"dl_zip_{run['run_id']}", type="primary")
    d2.download_button("⬇ Combined workbook", run["files"][comb_name], file_name=comb_name,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_comb_{run['run_id']}")
    d3.download_button("⬇ Run report (PDF)", run["report_pdf"], file_name=f"run_report_{run['run_id'][:8]}.pdf", mime="application/pdf",
                       key=f"dl_pdf_{run['run_id']}")
    d4.download_button("⬇ Validation report (CSV)", ("\ufeff" + run["validation_csv"]).encode("utf-8"), file_name="validation.csv",
                       mime="text/csv", key=f"dl_val_{run['run_id']}")
    with st.expander(f"One workbook per invoice ({len(r['files']['items'])})"):
        for name in r["files"]["items"]:
            st.download_button(f"⬇ {name}", run["files"][name], file_name=name,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_{run['run_id']}_{name}")

    # check before sending
    checks, diffs = runner.check_lines(r)
    missing = runner.missing_consignees(r)
    fixable = runner.fixable(r)
    if checks or diffs or missing:
        st.markdown("#### Check before sending")
        for line in checks:
            st.warning(line)
        for line in diffs:
            st.warning(line)
        if missing:
            st.warning("**Consignee missing** (the yellow placeholder is printed): " + "; ".join(f"{m['customer_no']} ({m['invoices']})" for m in missing))
            st.page_link("views/consignees.py", label="Add the consignees", icon="🏢")
        if fixable:
            st.info(f"A consignee is saved now for {len(fixable)} invoice(s) that printed the placeholder: {', '.join(fixable[:8])}"
                    + (" ..." if len(fixable) > 8 else ""))
            if st.button("Update invoices", key="rebuild_fix", type="primary"):
                _rebuild(run)
    else:
        st.success("Nothing to check: no warnings.")

    # invoice list
    st.markdown("#### Invoices")
    f1, f2 = st.columns([1, 2])
    view = f1.selectbox("Show", ["All", "Needs checking", "Merged only"], key="inv_filter")
    q = f2.text_input("Search invoice, country or customer number", key="inv_search").strip().lower()
    rows = []
    for i in r["invoices"]:
        g = runner.gauge(i.get("lines"), i["limit"])
        needs = bool(i["warnings"]) or g != "ok" or not i.get("consignee_ok", True)
        if view == "Needs checking" and not needs:
            continue
        if view == "Merged only" and i["cartons"] <= 1:
            continue
        if q and q not in f"{i['no']} {i['country']} {i['cus']}".lower():
            continue
        rows.append({"Invoice": i["no"], "Country": i["country"], "Customer": str(i["cus"]), "Cartons": i["cartons"],
                     "PO lines": i["n_po"], "Groups": i["n_grp"], "Pieces": i["qty"], "Amount (USD)": i["amount"],
                     "N.W": i["nw"], "G.W": i["gw"], "Lines": f"{i.get('lines') or 0} / {i['limit']}",
                     "Gauge": {"ok": "OK", "near": "NEAR LIMIT", "over": "OVER LIMIT"}[g],
                     "Consignee": "yes" if i.get("consignee_ok", True) else "MISSING", "Warnings": len(i["warnings"])})
    st.caption(f"Line limit: {engine.LINE_LIMIT_KEEP} printed lines for UK / South Africa / Australia, {engine.LINE_LIMIT_OTHER} for every other country.")
    if rows:
        st.dataframe(pd.DataFrame(rows), hide_index=True, column_config={"Amount (USD)": st.column_config.NumberColumn(format="%.2f")})
    else:
        st.info("No invoice matches.")

    # one invoice
    st.markdown("#### Invoice detail")
    by_no = {i["no"]: i for i in r["invoices"]}
    pick = st.selectbox("Invoice", list(by_no), key="inv_detail")
    inv = by_no[pick]
    if inv["merged_from"]:
        st.caption("Merged from master invoices: " + ", ".join(inv["merged_from"]))
    for w in inv["warnings"]:
        st.warning(w)
    for grp in inv["groups"]:
        st.markdown(f"**{grp['name']}**  ·  HTS {grp['hts']}" + (f"  \n{grp['comp']}" if grp["comp"] else ""))
        st.dataframe(pd.DataFrame([{"PO": ln["po"], "Customer PO": ln["cus_po"], "Article": ln["art"], "Qty": ln["qty"],
                                    "Price": ln["price"], "Amount": ln["amount"]} for ln in grp["lines"]]), hide_index=True)

    # by country
    st.markdown("#### By country")
    bc = pd.DataFrame(r["by_country"]).rename(columns={"country": "Country", "invoices": "Invoices", "qty": "Pieces", "amount": "Amount (USD)"})
    c1, c2 = st.columns([1, 1])
    c1.dataframe(bc, hide_index=True, column_config={"Amount (USD)": st.column_config.NumberColumn(format="%.2f")})
    c2.bar_chart(bc.set_index("Country")["Amount (USD)"])


# ------------------------------------------------------------------ upload form
handoff = st.session_state.get("master_handoff")
up = st.file_uploader("MASTER_INVOICE.xlsx", type=["xlsx"], key="build_master")
if up is None and handoff:
    a, b = st.columns([4, 1])
    a.info(f"Using the master built from **{handoff['source']}** on the Packing plan to master page.")
    if b.button("Forget it"):
        st.session_state.pop("master_handoff", None)
        st.rerun()

with st.expander("Options"):
    merge = st.checkbox("Merge cartons into one invoice", value=True,
                        help="Cartons are merged while the printed lines still fit (28 lines UK / South Africa / Australia, 34 elsewhere). "
                             "Woven and knitted are never combined; invoices with SET lines are never merged.")
    hide_content = st.checkbox("Remove content line from the next groups",
                               help="The first group keeps its composition line, the next groups lose it and groups are description + HTS code only.")
    tpl = st.file_uploader("Invoice template (optional, default: assets/invoice_template.xlsx)", type=["xlsx"], key="build_template")

if st.button("Build invoices", type="primary", disabled=(up is None and not handoff)):
    st.session_state.pop("run", None)                       # a failed build must not leave the previous result on screen
    data, name = (up.getvalue(), up.name) if up else (handoff["data"], handoff["source"])
    try:
        with st.spinner("Checking the master and building the invoices (this can take a minute) ..."):
            st.session_state["run"] = runner.build_invoices(data, name, tpl.getvalue() if tpl else None, merge, hide_content)
    except runner.InvalidMaster as e:                       # nothing was generated: every problem with its row
        st.error("Could not build the invoices: the master workbook has problems. Fix them and upload it again.")
        for p in e.problems:
            st.write(f"- {p}")
    except Exception as e:                                  # bad workbook, missing columns, ...
        st.error(f"Could not build the invoices: {e}")

if st.session_state.get("run"):
    show_run(st.session_state["run"])
