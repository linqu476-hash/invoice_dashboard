"""Small Streamlit helpers shared by the pages."""
import streamlit as st


def flash(message, kind="success"):
    """Show a message on the NEXT run (a message written right before st.rerun() would vanish)."""
    st.session_state["_flash"] = (kind, message)


def show_flash():
    item = st.session_state.pop("_flash", None)
    if item:
        getattr(st, item[0])(item[1])


def reset_choice(key, value="new"):
    """Ask for a selectbox to go back to `value` on the next run. Apply it with apply_reset(key) BEFORE the selectbox is created
    (Streamlit refuses to change a widget's state after the widget exists in the run)."""
    st.session_state[f"_reset_{key}"] = value


def apply_reset(key):
    flag = f"_reset_{key}"
    if flag in st.session_state:
        st.session_state[key] = st.session_state.pop(flag)


def local_time(iso):
    """A stored UTC time as 'YYYY-MM-DD HH:MM:SS <zone>' in the browser's time zone (UTC when the browser's zone is not known)."""
    from datetime import datetime
    from zoneinfo import ZoneInfo
    dt = datetime.fromisoformat(iso)
    try:
        tz = getattr(st.context, "timezone", None)
        if tz:
            dt = dt.astimezone(ZoneInfo(tz))
    except Exception:
        pass
    return f"{dt:%Y-%m-%d %H:%M:%S} {dt.tzname() or ''}".strip()


def duration(seconds):
    m, s = divmod(float(seconds), 60)
    return f"{int(m)} min {s:04.1f} s" if m >= 1 else f"{s:.1f} s"
