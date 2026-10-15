"""Password gate. The app handles customer addresses and invoice data, so it refuses to run without a password.

The password comes from the host's secrets (APP_PASSWORD in .streamlit/secrets.toml or the Secrets box of the host) or an
environment variable of the same name. This is ONE shared password: there are no per-user accounts, roles or audit trail
(the Django version had them). For per-user sign-in use Streamlit's st.login (OIDC) or put the app behind your company's
SSO / reverse proxy."""
import hmac
import os
import time

import streamlit as st

MAX_FAILS_BEFORE_DELAY = 3


def _expected():
    try:
        value = st.secrets.get("APP_PASSWORD")
    except Exception:                      # no secrets file at all
        value = None
    return str(value) if value else os.environ.get("APP_PASSWORD")


def require_login():
    """Stop the script with a sign-in form until the password is right."""
    expected = _expected()
    if not expected:
        st.error("**No password is configured.** Set `APP_PASSWORD` in `.streamlit/secrets.toml` (local) or in the host's Secrets "
                 "(see `.streamlit/secrets.toml.example`). The app will not start without one.")
        st.stop()
    if st.session_state.get("auth_ok"):
        return
    st.title("🧾 Invoice dashboard")
    with st.form("login"):
        pw = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Sign in")
    if submitted:
        if hmac.compare_digest(pw.encode("utf-8"), expected.encode("utf-8")):
            st.session_state["auth_ok"] = True
            st.session_state["fails"] = 0
            st.rerun()
        fails = st.session_state.get("fails", 0) + 1
        st.session_state["fails"] = fails
        if fails > MAX_FAILS_BEFORE_DELAY:
            time.sleep(min(2 * (fails - MAX_FAILS_BEFORE_DELAY), 10))     # slows guessing within one session (not per IP)
        st.error("Wrong password.")
    st.stop()


def sign_out_button():
    if st.sidebar.button("Sign out"):
        st.session_state.clear()           # also drops the uploaded files and built invoices held in this session
        st.rerun()
