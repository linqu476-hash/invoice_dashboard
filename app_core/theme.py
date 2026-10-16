"""Version 3 look (version 2 palette + new typography): colours, CSS and small HTML building blocks. Static markup only.

Every piece of text that comes from data goes through html.escape, so nothing a user typed (customer names, file names)
can inject markup into the page. The CSS is a fixed string.
"""
import html

import streamlit as st

from . import __version__

NAVY, BLUE, SKY, TEAL = "#0B1F3A", "#2563EB", "#38BDF8", "#0EA5A4"
AMBER, GREEN, RED, INK, MUTED, LINE = "#F59E0B", "#16A34A", "#DC2626", "#0F172A", "#5B6B82", "#E3E8F0"
TONES = {"blue": BLUE, "teal": TEAL, "amber": AMBER, "green": GREEN, "red": RED, "navy": NAVY}

CSS = f"""
<style>
/* Minimal system-font styling: no remote assets, gradients, animations, or heavy shadows. */
.stApp, .stApp button, .stApp input, .stApp textarea, .stApp select {{
    font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}}
.stApp {{ font-size: 15px; line-height: 1.45; }}
.block-container {{ padding-top: 1.25rem; padding-bottom: 2rem; max-width: 1280px; }}
h1, h2, h3, h4 {{ color: {NAVY}; }}
h1 {{ font-size: 1.8rem; font-weight: 700; }}
h4, .stApp h4 {{ font-size: 1.1rem; font-weight: 650; border-left: 3px solid {BLUE}; padding-left: .55rem; }}
[data-testid="stMetric"] {{ border: 1px solid {LINE}; border-radius: 6px; padding: 10px 12px; }}
[data-testid="stMetricLabel"] p {{ color: {MUTED}; font-size: .78rem; }}
[data-testid="stMetricValue"] {{ color: {NAVY}; font-weight: 700; }}
.stButton > button, .stDownloadButton > button, [data-testid="stFormSubmitButton"] > button {{
    border-radius: 5px; font-weight: 600; transition: none;
}}
.stButton > button:hover, .stDownloadButton > button:hover {{ transform: none; box-shadow: none; }}
[data-testid="stFileUploaderDropzone"] {{ border: 1px dashed #AAB4C2; border-radius: 6px; }}
[data-testid="stExpander"] {{ border: 1px solid {LINE}; border-radius: 6px; }}
[data-testid="stForm"] {{ border: 1px solid {LINE}; border-radius: 6px; padding: 1rem; }}
[data-testid="stDataFrame"] {{ border: 1px solid {LINE}; border-radius: 5px; }}
[data-testid="stAlert"] {{ border-radius: 5px; }}
[data-testid="stSidebar"] [data-testid="stSidebarNav"]::before {{
    content: "Invoice dashboard"; display: block; color: #fff; font-weight: 700; font-size: 1rem; padding: .8rem 1rem .3rem;
}}
[data-testid="stHeader"] {{ background: transparent; }}
.hero {{ background: #F1F5F9; color: {INK}; border: 1px solid {LINE}; border-radius: 6px; padding: 18px 20px; margin-bottom: 1rem; }}
.hero h1 {{ color: {NAVY}; margin: 0 0 .25rem 0; font-size: 1.7rem; }}
.hero p {{ color: {MUTED}; margin: 0 0 .5rem 0; }}
.pill {{ display: inline-block; background: #fff; border: 1px solid {LINE}; color: {INK}; border-radius: 4px; padding: 2px 8px; margin: 0 5px 5px 0; font-size: .78rem; }}
.pg-head {{ display: flex; align-items: center; gap: 10px; margin: 0 0 .3rem 0; }}
.pg-icon {{ width: 36px; height: 36px; border: 1px solid {LINE}; border-radius: 5px; display: flex; align-items: center; justify-content: center; font-size: 1.2rem; }}
.pg-head h1 {{ margin: 0; font-size: 1.65rem; }}
.kpi {{ background: #fff; border: 1px solid {LINE}; border-radius: 6px; padding: 12px 14px; border-left: 3px solid var(--tone); }}
.kpi .k-label {{ color: {MUTED}; font-size: .74rem; font-weight: 650; text-transform: uppercase; }}
.kpi .k-value {{ color: {NAVY}; font-size: 1.65rem; font-weight: 750; line-height: 1.2; margin-top: 2px; }}
.kpi .k-hint {{ color: {MUTED}; font-size: .82rem; margin-top: 2px; }}
.step {{ background: #fff; border: 1px solid {LINE}; border-radius: 6px; padding: 14px; min-height: 110px; }}
.step .s-num {{ display: inline-flex; width: 25px; height: 25px; border-radius: 4px; background: {BLUE}; color: #fff; align-items: center; justify-content: center; font-weight: 700; font-size: .82rem; margin-bottom: 7px; }}
.step .s-title {{ color: {NAVY}; font-weight: 700; font-size: 1rem; }}
.step .s-text {{ color: {MUTED}; font-size: .86rem; margin-top: 4px; }}
.empty {{ background: #fff; border: 1px dashed #AAB4C2; border-radius: 6px; padding: 22px; text-align: center; color: {MUTED}; }}
.empty b {{ color: {NAVY}; font-size: 1rem; }}
.login-brand {{ text-align: center; margin: 2rem 0 1rem 0; }}
.login-brand .logo {{ width: 48px; height: 48px; margin: 0 auto 8px auto; border: 1px solid {LINE}; border-radius: 6px; display: flex; align-items: center; justify-content: center; font-size: 1.7rem; }}
.login-brand h1 {{ margin: 0; font-size: 1.65rem; }}
.login-brand p {{ color: {MUTED}; margin: .2rem 0 0 0; }}
</style>
"""


def inject():
    """Add the stylesheet. Call once per run, right after st.set_page_config."""
    st.markdown(CSS, unsafe_allow_html=True)


def _e(text):
    return html.escape(str(text), quote=True)


def page_header_html(title, icon=""):
    badge = f'<div class="pg-icon">{_e(icon)}</div>' if icon else ""
    return f'<div class="pg-head">{badge}<h1>{_e(title)}</h1></div>'


def page_header(title, icon=""):
    """Replaces st.title: icon badge + title. The page's own st.caption follows it."""
    st.markdown(page_header_html(title, icon), unsafe_allow_html=True)


def hero(title, subtitle, pills=()):
    chips = "".join(f'<span class="pill">{_e(p)}</span>' for p in pills)
    st.markdown(f'<div class="hero"><h1>{_e(title)}</h1><p>{_e(subtitle)}</p>{chips}</div>', unsafe_allow_html=True)


def kpi_html(label, value, hint="", tone="blue"):
    color = TONES.get(tone, BLUE)
    hint_html = f'<div class="k-hint">{_e(hint)}</div>' if hint else ""
    return (f'<div class="kpi" style="--tone:{color}"><div class="k-label">{_e(label)}</div>'
            f'<div class="k-value">{_e(value)}</div>{hint_html}</div>')


def kpi(label, value, hint="", tone="blue"):
    st.markdown(kpi_html(label, value, hint, tone), unsafe_allow_html=True)


def step(number, title, text):
    st.markdown(f'<div class="step"><div class="s-num">{_e(number)}</div><div class="s-title">{_e(title)}</div>'
                f'<div class="s-text">{_e(text)}</div></div>', unsafe_allow_html=True)


def empty(title, text):
    st.markdown(f'<div class="empty"><b>{_e(title)}</b><br>{_e(text)}</div>', unsafe_allow_html=True)


def login_header():
    st.markdown('<div class="login-brand"><div class="logo">🧾</div><h1>Invoice dashboard</h1>'
                '<p>Sign in to continue</p></div>', unsafe_allow_html=True)


def version_label():
    st.sidebar.caption(f"Version {__version__}")
