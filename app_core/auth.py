"""Password gate (security layer 1). The app handles customer addresses and invoice data, so it refuses to run without a password.

Secrets (in .streamlit/secrets.toml, the host's Secrets box, or environment variables of the same name):
  APP_PASSWORD_HASH   recommended: a scrypt hash made with  python scripts/make_password_hash.py  (the password itself is never stored)
  APP_PASSWORD        plain password, still accepted when no hash is set
  IDLE_TIMEOUT_MIN    sign out after this many minutes without activity (default 30)
  MAX_SESSION_HOURS   sign out after this many hours in any case (default 12)

Protection: constant-time check, lockout after repeated wrong passwords that survives a refresh and applies to every session
(app_core/authlib.py), idle and absolute session timeouts, and an audit line for every sign-in event.
This is still ONE shared password: for per-user accounts use Streamlit's st.login (OIDC) or put the app behind your company's SSO.
"""
import os
import time

import streamlit as st

from . import authlib, security, theme


def _setting(name, default=None):
    try:
        value = st.secrets.get(name)
    except Exception:                      # no secrets file at all
        value = None
    return str(value) if value else os.environ.get(name, default)


def _expected():
    return _setting("APP_PASSWORD_HASH") or _setting("APP_PASSWORD")


def _number(name, default):
    try:
        return float(_setting(name, default))
    except (TypeError, ValueError):
        return float(default)


def _expired():
    """Reason text when the signed-in session ran out, else None."""
    now = time.time()
    idle, absolute = _number("IDLE_TIMEOUT_MIN", 30) * 60, _number("MAX_SESSION_HOURS", 12) * 3600
    if now - st.session_state.get("auth_last", now) > idle:
        return "idle"
    if now - st.session_state.get("auth_start", now) > absolute:
        return "max_age"
    return None


def require_login():
    """Stop the script with a sign-in form until the password is right."""
    expected = _expected()
    if not expected:
        st.error("**No password is configured.** Set `APP_PASSWORD_HASH` (or `APP_PASSWORD`) in `.streamlit/secrets.toml` (local) or in the "
                 "host's Secrets (see `.streamlit/secrets.toml.example`). The app will not start without one.")
        st.stop()
    if st.session_state.get("auth_ok"):
        reason = _expired()
        if reason is None:
            st.session_state["auth_last"] = time.time()
            return
        security.audit("session_expired", reason=reason)
        st.session_state.clear()           # drops the uploaded files and built invoices held in this session
        st.session_state["auth_notice"] = "You were signed out because the session was idle too long." if reason == "idle" \
            else "You were signed out because the session reached its maximum length."
    theme.login_header()
    _, mid, _ = st.columns([1, 1.3, 1])          # a centred sign-in card
    with mid:
        _login_form(expected)
    st.stop()


def _login_form(expected):
    if st.session_state.get("auth_notice"):
        st.info(st.session_state.pop("auth_notice"))
    left = authlib.seconds_locked()
    if left:
        st.error(f"Too many wrong passwords. Sign-in is locked for {left} more seconds, then try again.")
        return
    with st.form("login"):
        pw = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Sign in", use_container_width=True)
    if submitted:
        if authlib.seconds_locked():
            st.rerun()
        if authlib.verify_password(pw, expected):
            authlib.record_success()
            security.audit("login_ok")
            st.session_state.update(auth_ok=True, auth_start=time.time(), auth_last=time.time())
            st.rerun()
        lock = authlib.record_failure()
        security.audit("login_failed", locked_seconds=lock)
        if lock:
            st.error(f"Too many wrong passwords. Sign-in is locked for {lock} seconds.")
        else:
            st.error("Wrong password.")


def sign_out_button():
    if st.sidebar.button("Sign out"):
        security.audit("logout")
        st.session_state.clear()           # also drops the uploaded files and built invoices held in this session
        st.rerun()
