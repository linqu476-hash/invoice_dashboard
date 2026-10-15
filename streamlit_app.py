"""Invoice dashboard, Streamlit version.   Run:  streamlit run streamlit_app.py

Packing plan -> MASTER_INVOICE.xlsx -> commercial invoices + packing lists. The invoice logic is the unchanged `invoicing/` package of the
Django version; this file only provides the pages, the password gate and the navigation."""
import streamlit as st

st.set_page_config(page_title="Invoice dashboard", page_icon="🧾", layout="wide")

from app_core import auth, store  # noqa: E402  (after set_page_config)

auth.require_login()


@st.cache_resource
def _seed_once():
    """Hosts with a wiped disk: load data/seed.json into an empty database, once per server start."""
    return store.seed_if_empty()


_seed_once()

pages = [
    st.Page("views/build.py", title="Build invoices", icon="🧾", default=True),
    st.Page("views/master.py", title="Packing plan to master", icon="📦"),
    st.Page("views/consignees.py", title="Consignees", icon="🏢"),
    st.Page("views/marks.py", title="Shipping marks", icon="🏷️"),
    st.Page("views/data.py", title="Backup and data", icon="💾"),
]
page = st.navigation(pages)
auth.sign_out_button()
page.run()
