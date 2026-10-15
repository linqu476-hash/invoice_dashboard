"""Dashboard (home): what the app holds, the 4-step workflow, and the result of the last build in this session."""
import pandas as pd
import streamlit as st

from app_core import jobs_ui, runner, store, theme

consignees, marks = store.consignees(), store.marks()
run = st.session_state.get("run")
dupes = len(store.conflicts("consignee")) + len(store.conflicts("shippingmark"))

theme.hero("Invoice dashboard", "Packing plan → master invoice → commercial invoices and packing lists, in one place.",
           pills=("🔒 Password protected", "🛡️ Upload checks", "📝 Audit log"))

# ------------------------------------------------------------------ key numbers
c1, c2, c3, c4 = st.columns(4)
with c1:
    theme.kpi("Consignees", f"{len(consignees):,}", "saved by customer number", "blue")
with c2:
    theme.kpi("Shipping marks", f"{len(marks):,}", "saved by customer number", "teal")
if run:
    t = run["result"]["totals"]
    with c3:
        theme.kpi("Last build: invoices", f"{t['invoices']:,}", f"{t['qty']:,} pieces", "navy")
    with c4:
        theme.kpi("Last build: amount", f"${t['amount']:,.0f}", "USD", "green")
else:
    with c3:
        theme.kpi("Last build", "none yet", "in this session", "navy")
    with c4:
        theme.kpi("Data check", "OK" if not dupes else f"{dupes} to fix", "duplicate customer numbers" if dupes else "no duplicates",
                  "green" if not dupes else "amber")

# ------------------------------------------------------------------ builds in progress and failed builds (kept when the user changes page)
jobs_ui.show_failures("home")
jobs_ui.live_panel()

# ------------------------------------------------------------------ workflow
st.markdown("#### How it works")
steps = [("1", "Packing plan to master", "Upload the packing plan and get MASTER_INVOICE.xlsx.", "views/master.py", "📦"),
         ("2", "Build invoices", "Upload the master: invoices and packing lists are built and checked.", "views/build.py", "🧾"),
         ("3", "Consignees", "Keep the consignee address of every customer number.", "views/consignees.py", "🏢"),
         ("4", "Shipping marks", "Keep the shipping marks printed on each invoice.", "views/marks.py", "🏷️")]
for col, (n, title, text, page, icon) in zip(st.columns(4), steps):
    with col:
        theme.step(n, title, text)
        st.page_link(page, label=f"Open {title.lower()}" if len(title) < 16 else "Open", icon=icon)

# ------------------------------------------------------------------ last build
st.markdown("#### Last build")
if not run:
    theme.empty("Nothing built in this session yet", "Upload a master invoice on the Build invoices page. Uploaded files and results are kept in this "
                "browser session only and are dropped when you sign out.")
    st.page_link("views/build.py", label="Go to Build invoices", icon="🧾")
else:
    r = run["result"]
    sc = runner.status_counts(r)
    left, right = st.columns([1, 1])
    with left:
        st.markdown(f"**{r['master_name']}**")
        s1, s2 = st.columns(2)
        s1.metric("Ready to send", f"{sc['ok']:,} of {sc['total']:,}")
        s2.metric("Need checking", f"{sc['pending']:,}")
        st.progress(sc["ok"] / sc["total"] if sc["total"] else 0.0, text=f"{sc['ok']} ready, {sc['pending']} pending")
        st.page_link("views/build.py", label="Open the result and downloads", icon="🧾")
    with right:
        st.caption("Amount by country (USD)")
        df = pd.DataFrame(r["by_country"])
        if not df.empty:
            st.bar_chart(df.set_index("country")["amount"].sort_values(ascending=False).head(8), color="#2563EB", height=240)

if dupes:
    st.warning(f"{dupes} customer number(s) are saved twice, written differently (for example 000035 and 35). Only one can match an invoice.")
    st.page_link("views/data.py", label="Open Backup and data", icon="💾")

# ------------------------------------------------------------------ build history of this session
rows = jobs_ui.history_rows()
if rows:
    st.markdown("#### Recent builds")
    st.dataframe(pd.DataFrame(rows), hide_index=True)
    st.caption("Builds run in the background: you can open any page while one is running. The list is kept in this browser session only.")
