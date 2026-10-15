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
