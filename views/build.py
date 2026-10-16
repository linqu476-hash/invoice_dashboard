"""Master invoice -> commercial invoices + packing lists, with the result view."""
import time
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

from app_core import runner
from invoicing import engine

st.title("Build invoices")
st.caption("Upload MASTER_INVOICE.xlsx (make it on **Packing plan to master**). Each carton after the first one of an invoice becomes "
           "SC260458-1, -2, ... The invoice date is the CRD of the packing plan (master column INV DATE).")


# ------------------------------------------------------------------ live progress while a build runs
def _with_progress(label, work):
    """Run work(on_progress) inside a status box that shows, while it runs: the start time, a progress bar with the current step
    (invoice n of N, saving files ...), the elapsed time and a rough time left. Returns what work returns; errors are re-raised."""
    t0 = time.perf_counter()
    started = _local(datetime.now(timezone.utc).isoformat())
    with st.status(label, expanded=True) as box:
        st.write(f"**Started:** {started}")
        bar = st.progress(0.0, text="Starting ...")
        info = st.empty()

        def on_progress(fraction, text):
            elapsed = time.perf_counter() - t0
            bar.progress(max(0.0, min(float(fraction), 1.0)), text=f"{text}  ({fraction * 100:.0f}%)")
            left = f" · about {_duration(elapsed * (1 - fraction) / fraction)} left" if fraction >= 0.15 else ""
            info.caption(f"Elapsed {_duration(elapsed)}{left}")

        try:
            out = work(on_progress)
        except Exception:
            box.update(label="Build stopped", state="error", expanded=True)
            raise
        box.update(label=f"Finished in {_duration(time.perf_counter() - t0)}", state="complete", expanded=False)
    return out


# ------------------------------------------------------------------ result view
def _local(iso):
    """A stored UTC time as 'YYYY-MM-DD HH:MM:SS <zone>' in the browser's time zone (UTC when the browser's zone is not known)."""
    dt = datetime.fromisoformat(iso)
    try:
        tz = getattr(st.context, "timezone", None)
        if tz:
            dt = dt.astimezone(ZoneInfo(tz))
    except Exception:
        pass
    return f"{dt:%Y-%m-%d %H:%M:%S} {dt.tzname() or ''}".strip()


def _duration(seconds):
    m, s = divmod(float(seconds), 60)
    return f"{int(m)} min {s:04.1f} s" if m >= 1 else f"{s:.1f} s"


def _rebuild(run):
    try:
        st.session_state["run"] = _with_progress("Building the invoices again with the consignees saved now ...",
                                                 lambda on_progress: runner.rebuild(run, progress=on_progress))
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

    cols = st.columns(3)
    for col, (label, value) in zip(cols, [("Invoices", f"{t['invoices']:,}"),
                                          ("Amount (USD)", f"{t['amount']:,.2f}"),
                                          ("Warnings", f"{t['warnings']:,}")]):
        col.metric(label, value)
    # Keep the default results view compact; detailed readiness metrics are available on demand.
    with st.expander("Readiness summary", expanded=False):
        sc = runner.status_counts(r)
        s1, s2 = st.columns(2)
        s1.metric("Ready to send", f"{sc['ok']:,} of {sc['total']:,}", help="No warning, line count inside the limit, consignee on file.")
        s2.metric("Pending (need checking)", f"{sc['pending']:,}", help="Use the invoice filter to show invoices that need checking.")
        st.progress(sc["ok"] / sc["total"] if sc["total"] else 0.0, text=f"{sc['ok']} ready, {sc['pending']} pending")
        if sc["pending"]:
            why = [f"{n} {label}" for n, label in ((sc["missing_consignee"], "consignee missing"), (sc["line_limit"], "near / over the line limit"),
                                                  (sc["warnings"], "with warnings")) if n]
            st.caption("Pending because of: " + ", ".join(why) + " (one invoice can be in more than one group).")
    if r.get("merge", True):
        saved = r["master_invoices"] - t["invoices"]
        st.write(f"**{r['master_invoices']} cartons became {t['invoices']} invoice{'s' if t['invoices'] != 1 else ''}**"
                 + (f" (merging saved {saved})." if saved > 0 else "."))
    else:
        st.write("Merging was off: each carton of the master is its own invoice.")
    st.caption(f"Invoice date: {r['date']}")
    # downloads
    st.markdown("#### Downloads")
    zip_name, comb_name = r["files"]["zip"], r["files"].get("combined")
    download_cols = st.columns(4 if comb_name else 3)
    download_cols[0].download_button("All invoices (ZIP)", run["files"][zip_name], file_name=zip_name, mime="application/zip",
                                     key=f"dl_zip_{run['run_id']}", type="primary")
    offset = 1
    if comb_name and comb_name in run["files"]:
        download_cols[1].download_button("Combined workbook", runner.output_file_bytes(run, comb_name), file_name=comb_name,
                                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_comb_{run['run_id']}")
        offset = 2
    download_cols[offset].download_button("Run report (PDF)", run["report_pdf"], file_name=f"run_report_{run['run_id'][:8]}.pdf", mime="application/pdf",
                                          key=f"dl_pdf_{run['run_id']}")
    download_cols[offset + 1].download_button("Validation report (CSV)", ("\ufeff" + run["validation_csv"]).encode("utf-8"), file_name="validation.csv",
                                              mime="text/csv", key=f"dl_val_{run['run_id']}")
    if not comb_name:
        st.caption("Combined workbook was skipped to reduce build time. Enable it under Options if you need it.")
    with st.expander(f"One workbook per invoice ({len(r['files']['items'])})"):
        item_names = r["files"]["items"]
        if item_names:
            selected_file = st.selectbox("Choose workbook", item_names, key=f"family_file_{run['run_id']}")
            st.download_button("Download selected workbook", runner.output_file_bytes(run, selected_file), file_name=selected_file,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_family_{run['run_id']}")

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
    # Filter compact invoice records first; only format and render the current page.
    filtered = []
    for i in r["invoices"]:
        if view == "Needs checking" and not runner.needs_check(i):
            continue
        if view == "Merged only" and i["cartons"] <= 1:
            continue
        if q and q not in f"{i['no']} {i['country']} {i['cus']}".lower():
            continue
        filtered.append(i)

    page_size = st.selectbox("Rows per page", [50, 100, 250, 500], index=1, key="inv_page_size")
    page_count = max(1, (len(filtered) + page_size - 1) // page_size)
    filter_signature = (view, q, page_size)
    if st.session_state.get("_inv_filter_signature") != filter_signature:
        st.session_state["_inv_filter_signature"] = filter_signature
        st.session_state["inv_table_page"] = 1
    page_number = st.selectbox("Page", list(range(1, page_count + 1)), key="inv_table_page")
    start = (page_number - 1) * page_size
    visible = filtered[start:start + page_size]
    st.caption(f"Showing {len(visible):,} of {len(filtered):,} matching invoices · page {page_number} of {page_count}.")

    rows = []
    for i in visible:
        g = runner.gauge(i.get("lines"), i["limit"])
        rows.append({"Invoice": i["no"], "Country": i["country"], "Customer": str(i["cus"]), "Cartons": i["cartons"],
                     "PO lines": i["n_po"], "Groups": i["n_grp"], "Pieces": i["qty"], "Amount (USD)": i["amount"],
                     "N.W": i["nw"], "G.W": i["gw"], "Lines": f"{i.get('lines') or 0} / {i['limit']}",
                     "Gauge": {"ok": "OK", "near": "NEAR LIMIT", "over": "OVER LIMIT"}[g],
                     "Consignee": "yes" if i.get("consignee_ok", True) else "MISSING", "Warnings": len(i["warnings"])})
    if rows:
        st.dataframe(pd.DataFrame(rows), hide_index=True, column_config={"Amount (USD)": st.column_config.NumberColumn(format="%.2f")})
    else:
        st.info("No invoice matches.")

    # Detailed line-by-line tables were removed from the default UI in v11.
    # The authoritative invoice workbook remains downloadable; this keeps reruns
    # lighter for builds with many groups and PO lines.



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
    combine = st.checkbox("Generate combined workbook (slower; optional)", value=False,
                          help="Adds a single workbook containing every invoice and packing list. The ZIP always contains the per-invoice family workbooks.")
    tpl = st.file_uploader("Invoice template (optional, default: assets/invoice_template.xlsx)", type=["xlsx"], key="build_template")

if st.button("Build invoices", type="primary", disabled=(up is None and not handoff)):
    st.session_state.pop("run", None)                       # a failed build must not leave the previous result on screen
    data, name = (up.getvalue(), up.name) if up else (handoff["data"], handoff["source"])
    try:
        tpl_bytes = tpl.getvalue() if tpl else None
        st.session_state["run"] = _with_progress("Building the invoices ...",
                                                 lambda on_progress: runner.build_invoices(data, name, tpl_bytes, merge, hide_content, progress=on_progress, combine=combine))
    except runner.InvalidMaster as e:                       # nothing was generated: every problem with its row
        st.error("Could not build the invoices: the master workbook has problems. Fix them and upload it again.")
        for p in e.problems:
            st.write(f"- {p}")
    except Exception as e:                                  # bad workbook, missing columns, ...
        st.error(f"Could not build the invoices: {e}")

if st.session_state.get("run"):
    show_run(st.session_state["run"])
