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

# Typography (v3). The font files are in static/fonts and declared in .streamlit/config.toml, so the stylesheet only names them.
FONT_BODY = "'Inter', system-ui, -apple-system, 'Segoe UI', Roboto, sans-serif"
FONT_HEAD = "'Plus Jakarta Sans', 'Inter', system-ui, sans-serif"

CSS = f"""
<style>
.stApp, .stApp button, .stApp input, .stApp textarea, .stApp select {{ font-family: {FONT_BODY}; font-feature-settings: "cv11", "ss01"; }}
.stApp {{ font-size: 15px; line-height: 1.55; -webkit-font-smoothing: antialiased; text-rendering: optimizeLegibility; }}
h1, h2, h3, h4, h5, .hero h1, .pg-head h1, .login-brand h1, .kpi .k-value, .step .s-title {{ font-family: {FONT_HEAD}; }}
h2, h3 {{ letter-spacing: -0.015em; font-weight: 700; color: {NAVY}; }}
[data-testid="stMetricValue"], .kpi .k-value, [data-testid="stDataFrame"] {{ font-variant-numeric: tabular-nums; }}
code, pre, kbd {{ font-family: 'JetBrains Mono', ui-monospace, Menlo, Consolas, monospace; font-size: .88em; }}
[data-testid="stCaptionContainer"] {{ font-size: .86rem; letter-spacing: .005em; }}
[data-testid="stSidebar"] {{ font-size: .95rem; }}
[data-testid="stSidebarNav"] a span {{ font-weight: 500; }}
[data-testid="stSidebarNavSeparator"], [data-testid="stSidebarNav"] [data-testid="stNavSectionHeader"] {{ letter-spacing: .08em; text-transform: uppercase; font-size: .7rem; font-weight: 700; }}
.block-container {{ padding-top: 2.2rem; padding-bottom: 3rem; max-width: 1280px; }}
h1 {{ font-weight: 800; letter-spacing: -0.025em; line-height: 1.15; color: {NAVY}; }}
h4, .stApp h4 {{ color: {NAVY}; font-weight: 700; font-size: 1.25rem; letter-spacing: -0.01em; border-left: 4px solid {BLUE}; padding: .1rem 0 .1rem .7rem; margin-top: 1.6rem; }}
[data-testid="stMetric"] {{ background: #fff; border: 1px solid {LINE}; border-radius: 14px; padding: 14px 18px;
    box-shadow: 0 1px 2px rgba(15,23,42,.05), 0 6px 16px rgba(15,23,42,.05); }}
[data-testid="stMetricLabel"] p {{ color: {MUTED}; font-weight: 700; font-size: .74rem; text-transform: uppercase; letter-spacing: .05em; }}
[data-testid="stMetricValue"] {{ color: {NAVY}; font-weight: 800; }}
.stButton > button, .stDownloadButton > button, [data-testid="stFormSubmitButton"] > button {{
    border-radius: 10px; font-weight: 600; border: 1px solid {LINE}; transition: transform .08s ease, box-shadow .15s ease; }}
.stButton > button:hover, .stDownloadButton > button:hover {{ transform: translateY(-1px); box-shadow: 0 6px 14px rgba(37,99,235,.18); }}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"], [data-testid="stFormSubmitButton"] > button {{
    background: linear-gradient(135deg, {BLUE}, #1D4ED8); border: 0; color: #fff; }}
.stTabs [data-baseweb="tab-list"] {{ gap: 6px; border-bottom: 1px solid {LINE}; }}
.stTabs [data-baseweb="tab"] {{ padding: 8px 18px; font-weight: 600; border-radius: 10px 10px 0 0; }}
[data-testid="stFileUploaderDropzone"] {{ border: 2px dashed #9DB4DD; border-radius: 14px; background: #F8FAFF; }}
[data-testid="stExpander"] {{ border: 1px solid {LINE}; border-radius: 12px; background: #fff; }}
[data-testid="stForm"] {{ border: 1px solid {LINE}; border-radius: 16px; background: #fff; padding: 1.2rem;
    box-shadow: 0 10px 30px rgba(15,23,42,.08); }}
[data-testid="stDataFrame"] {{ border: 1px solid {LINE}; border-radius: 12px; overflow: hidden; }}
[data-testid="stAlert"] {{ border-radius: 12px; }}
[data-testid="stSidebar"] [data-testid="stSidebarNav"]::before {{ content: "🧾  Invoice dashboard"; display: block; color: #fff;
    font-weight: 800; font-size: 1.15rem; padding: 1.1rem 1rem .4rem; }}
[data-testid="stHeader"] {{ background: transparent; }}

.hero {{ background: linear-gradient(120deg, {NAVY} 0%, #123A73 55%, {BLUE} 100%); color: #fff; border-radius: 20px;
    padding: 28px 32px; margin-bottom: 1.2rem; box-shadow: 0 14px 34px rgba(11,31,58,.28); }}
.hero h1 {{ color: #fff; margin: 0 0 .3rem 0; font-size: 2rem; }}
.hero p {{ color: #CFE0FF; margin: 0 0 .9rem 0; font-size: 1.02rem; }}
.pill {{ display: inline-block; background: rgba(255,255,255,.14); border: 1px solid rgba(255,255,255,.28); color: #fff;
    border-radius: 999px; padding: 3px 12px; margin: 0 6px 6px 0; font-size: .78rem; font-weight: 600; }}
.pg-head {{ display: flex; align-items: center; gap: 14px; margin: 0 0 .3rem 0; }}
.pg-icon {{ width: 46px; height: 46px; border-radius: 13px; display: flex; align-items: center; justify-content: center;
    font-size: 1.5rem; background: linear-gradient(135deg, #E7EEFF, #D3E1FF); }}
.pg-head h1 {{ margin: 0; font-size: 1.8rem; }}
.kpi {{ background: #fff; border: 1px solid {LINE}; border-radius: 16px; padding: 16px 18px; border-top: 4px solid var(--tone);
    box-shadow: 0 1px 2px rgba(15,23,42,.05), 0 8px 20px rgba(15,23,42,.06); }}
.kpi .k-label {{ color: {MUTED}; font-size: .74rem; font-weight: 700; text-transform: uppercase; letter-spacing: .05em; }}
.kpi .k-value {{ color: {NAVY}; font-size: 1.9rem; font-weight: 800; line-height: 1.2; margin-top: 2px; }}
.kpi .k-hint {{ color: {MUTED}; font-size: .82rem; margin-top: 2px; }}
.step {{ background: #fff; border: 1px solid {LINE}; border-radius: 16px; padding: 16px 18px; min-height: 128px;
    box-shadow: 0 1px 2px rgba(15,23,42,.05), 0 8px 20px rgba(15,23,42,.05); }}
.step .s-num {{ display: inline-flex; width: 28px; height: 28px; border-radius: 50%; background: {BLUE}; color: #fff;
    align-items: center; justify-content: center; font-weight: 800; font-size: .85rem; margin-bottom: 8px; }}
.step .s-title {{ color: {NAVY}; font-weight: 800; font-size: 1.02rem; }}
.step .s-text {{ color: {MUTED}; font-size: .86rem; margin-top: 4px; }}
.empty {{ background: #fff; border: 2px dashed #B9C8E6; border-radius: 16px; padding: 26px; text-align: center; color: {MUTED}; }}
.empty b {{ color: {NAVY}; font-size: 1.05rem; }}
.login-brand {{ text-align: center; margin: 2.5rem 0 1rem 0; }}
.login-brand .logo {{ width: 64px; height: 64px; margin: 0 auto 10px auto; border-radius: 18px; display: flex; align-items: center;
    justify-content: center; font-size: 2rem; background: linear-gradient(135deg, {NAVY}, {BLUE}); box-shadow: 0 10px 24px rgba(37,99,235,.35); }}
.login-brand h1 {{ margin: 0; font-size: 1.8rem; }}
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
