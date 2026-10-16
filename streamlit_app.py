"""Invoice dashboard, Streamlit version 11.   Run:  streamlit run streamlit_app.py

Packing plan -> MASTER_INVOICE.xlsx -> commercial invoices + packing lists. The invoice logic is the unchanged `invoicing/` package of the
Django version; this file only provides the pages, the password gate and the navigation.
Builds run in background jobs (app_core/jobs.py) so they survive page changes. The look is in app_core/theme.py."""
import streamlit as st

st.set_page_config(page_title="Invoice dashboard", page_icon="🧾", layout="wide", initial_sidebar_state="expanded")

from app_core import auth, jobs_ui, store, theme  # noqa: E402  (after set_page_config)

theme.inject()
auth.require_login()


@st.cache_resource
def _seed_once():
    """Hosts with a wiped disk: load data/seed.json into an empty database, once per server start."""
    return store.seed_if_empty()


_seed_once()

pages = {
    "Overview": [st.Page("views/home.py", title="Dashboard", icon="🏠", default=True)],
    "Invoicing": [st.Page("views/master.py", title="Packing plan to master", icon="📦"),
                  st.Page("views/build.py", title="Build invoices", icon="🧾")],
    "Master data": [st.Page("views/consignees.py", title="Consignees", icon="🏢"),
                    st.Page("views/marks.py", title="Shipping marks", icon="🏷️")],
    "System": [st.Page("views/data.py", title="Backup and data", icon="💾")],
}
page = st.navigation(pages)
auth.sign_out_button()
theme.version_label()
jobs_ui.sync()            # take over builds that finished while the user was on another page (result, failure, toast)
jobs_ui.sidebar_status()  # running build + failed-build link, on every page
page.run()
