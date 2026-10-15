"""Streamlit side of the background jobs (version 4): start a job, show it live on any page, report how it ended.

Nothing here runs the build itself (app_core/jobs.py does, in a worker thread); every function only reads the job state, so it is safe
to call on every page and on every re-run. A page that shows a running job polls it with a fragment (st.fragment run_every) and
asks for one full re-run when the job ends, so the result / the failure appears without a click.
"""
import uuid

import streamlit as st

from . import jobs, ui

POLL_SECONDS = 1.0
ICON = {jobs.QUEUED: "⏳", jobs.RUNNING: "⚙️", jobs.DONE: "✅", jobs.FAILED: "❌", jobs.CANCELLED: "🚫"}


def owner():
    """The id of this browser session. Jobs are always asked for with it, so a session never sees another one's jobs."""
    return st.session_state.setdefault("_job_owner", uuid.uuid4().hex)


def start(kind, label, filename, work):
    """Queue work(progress) -> run. Returns True when it was queued, False (with a message) when this session already has a build."""
    try:
        jobs.manager().submit(owner(), kind, label, filename, work)
    except jobs.JobBusy:
        st.warning("A build is already running. Wait for it to finish, or cancel it first.")
        return False
    return True


def is_busy():
    return jobs.manager().active(owner()) is not None


def sync():
    """Call once per full run, after the sign-in. Takes over what finished while the user was on another page:
    a good result becomes st.session_state['run'] (what the Dashboard and the Build page show) and every end gets one toast."""
    for view, run in jobs.manager().collect(owner()):
        if run is not None:
            st.session_state["run"] = run
            st.toast(f"Build finished: {view.filename}", icon="✅")
        elif view.status == jobs.FAILED:
            st.toast(f"Build failed: {view.filename}. {view.error}", icon="❌")
        elif view.status == jobs.CANCELLED:
            st.toast(f"Build cancelled: {view.filename}", icon="🚫")


def _live(render):
    """Run render(active_view | None). While a job is active it re-runs itself every second (only this block, not the page);
    when the job has ended it asks for ONE full re-run so sync() can take over the result."""
    was_active = jobs.manager().active(owner()) is not None

    def block():
        cur = jobs.manager().active(owner())
        if was_active and cur is None:
            st.rerun()
        render(cur)

    st.fragment(block, run_every=POLL_SECONDS if was_active else None)()


# ------------------------------------------------------------------ the running build (Build page, Dashboard)
def _progress_panel(cur):
    if cur is None:
        return
    with st.container(border=True):
        st.markdown(f"{ICON[cur.status]} **{cur.label}**  ·  {cur.filename}")
        if cur.started:
            st.caption(f"Started {ui.local_time(cur.started)}")
        pct = cur.fraction * 100
        st.progress(cur.fraction, text=f"{cur.text}  ({pct:.0f}%)" if cur.status == jobs.RUNNING else cur.text)
        left = f" · about {ui.duration(cur.elapsed * (1 - cur.fraction) / cur.fraction)} left" if cur.fraction >= 0.15 and cur.status == jobs.RUNNING else ""
        st.caption(f"Elapsed {ui.duration(cur.elapsed)}{left}  ·  it keeps running if you open another page.")
        if cur.cancelling:
            st.caption("Cancelling ... it stops at the next step.")
        elif st.button("Cancel this build", key=f"cancel_{cur.id}"):
            jobs.manager().cancel(owner(), cur.id)


def live_panel():
    """The running build with a progress bar. Shows nothing when no build is running."""
    _live(_progress_panel)


# ------------------------------------------------------------------ how builds ended
def _failure_box(v, detail, key):
    with st.container(border=True):
        what = "cancelled" if v.status == jobs.CANCELLED else "failed"
        st.error(f"**Build {what}:** {v.filename}" + (f" ({ui.local_time(v.finished)})" if v.finished else "") + (f"\n\n{v.error}" if v.error else ""))
        if v.problems:
            st.write("Fix these in the master workbook and upload it again:")
            for p in (v.problems if detail else v.problems[:3]):
                st.write(f"- {p}")
            if not detail and len(v.problems) > 3:
                st.caption(f"... and {len(v.problems) - 3} more. Open Build invoices for the full list.")
        if st.button("Dismiss", key=f"dismiss_{key}_{v.id}"):
            jobs.manager().dismiss(owner(), v.id)
            st.rerun()


def show_failures(key, detail=False):
    """Every build of this session that failed or was cancelled and not dismissed yet (newest first)."""
    for v in jobs.manager().jobs(owner()):
        if v.status in (jobs.FAILED, jobs.CANCELLED) and not v.dismissed:
            _failure_box(v, detail, key)


def history_rows():
    """Recent builds of this session as rows for a table."""
    rows = []
    for v in jobs.manager().jobs(owner()):
        rows.append({"": ICON[v.status], "Status": v.status.capitalize(), "File": v.filename, "Type": "Rebuild" if v.kind == "rebuild" else "Build",
                     "Started": ui.local_time(v.started) if v.started else "waiting", "Duration": ui.duration(v.elapsed) if v.started else "",
                     "Result": v.error or (f"{v.fraction * 100:.0f}%" if v.active else "")})
    return rows


# ------------------------------------------------------------------ sidebar badge (every page)
def _badge(cur):
    failed = sum(1 for v in jobs.manager().jobs(owner()) if v.status == jobs.FAILED and not v.dismissed)
    if cur is not None:
        st.caption(f"{ICON[cur.status]} **{cur.label}**  ·  {cur.fraction * 100:.0f}%")
        st.progress(cur.fraction)
    if failed:
        st.page_link("views/home.py", label=f"{failed} build{'s' if failed != 1 else ''} failed. See the Dashboard", icon="⚠️")


def sidebar_status():
    """Small status under the navigation on every page: the running build with its progress, and a link when a build failed."""
    with st.sidebar:
        _live(_badge)
