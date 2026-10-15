"""Executes every page script against a FAKE `streamlit` module, with real built data in the session.

What this catches: wrong keys, DataFrame shaping, typos in how the pages use app_core / invoicing, exceptions in the idle state.
What it does NOT catch: misuse of Streamlit's own API or layout problems. Always click through the real app once.

    python tests/smoke_pages.py
"""
import os
import runpy
import sys
import tempfile
import types
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
os.environ["INVOICE_DB_PATH"] = str(Path(tempfile.mkdtemp()) / "smoke.sqlite3")


class State(dict):
    __getattr__ = dict.get

    def __setattr__(self, k, v):
        self[k] = v


class Dummy:
    """Context manager / container: every method returns another Dummy; falsy so buttons are 'not clicked'."""
    def __enter__(self): return self
    def __exit__(self, *a): return False
    def __bool__(self): return False
    def __call__(self, *a, **k): return Dummy()
    def __getattr__(self, name): return Dummy()
    def __iter__(self): return iter(())


class FakeStreamlit(types.ModuleType):
    def __init__(self):
        super().__init__("streamlit")
        self.session_state = State()
        self.frames = []                      # every DataFrame a page tried to show
        self.messages = []                    # st.error / st.warning texts
        self.column_config = Dummy()

    # layout
    def columns(self, spec, **k): return [Dummy() for _ in range(spec if isinstance(spec, int) else len(spec))]
    def tabs(self, labels): return [Dummy() for _ in labels]
    # inputs (idle: nothing uploaded, nothing clicked, defaults returned)
    def file_uploader(self, *a, accept_multiple_files=False, **k): return [] if accept_multiple_files else None
    def text_input(self, label, value="", **k): return value
    def text_area(self, label, value="", **k): return value
    def checkbox(self, label, value=False, **k): return value
    def selectbox(self, label, options, **k):
        opts = list(options)
        return opts[0] if opts else None
    def button(self, *a, **k): return False
    def download_button(self, *a, **k): return False
    def form_submit_button(self, *a, **k): return False
    # output
    def dataframe(self, df, **k): self.frames.append(df)
    def error(self, msg, *a, **k): self.messages.append(("error", str(msg)))
    def warning(self, msg, *a, **k): self.messages.append(("warning", str(msg)))
    def rerun(self): raise RuntimeError("st.rerun called in smoke run")
    def switch_page(self, p): raise RuntimeError("st.switch_page called")
    def cache_resource(self, f=None, **k): return f if f else (lambda g: g)
    def fragment(self, f=None, **k): return f if f else (lambda g: g)     # run the block once, like the first run of a real fragment
    def __getattr__(self, name):
        if name.startswith("__"):
            raise AttributeError(name)
        return Dummy()


def run_page(name, state=None):
    fake = FakeStreamlit()
    fake.session_state.update(state or {})
    sys.modules["streamlit"] = fake
    for m in [m for m in sys.modules if m.startswith(("app_core.ui", "app_core.jobs_ui"))]:
        del sys.modules[m]
    runpy.run_path(str(ROOT / "views" / name), run_name="__page__")
    return fake


if __name__ == "__main__":
    from app_core import runner, store

    # synthetic-free: uses the shipped template + a master built from whatever packing plan is given on the command line
    plan = sys.argv[1] if len(sys.argv) > 1 else None
    pages = ["home.py", "build.py", "master.py", "consignees.py", "marks.py", "data.py"]
    for p in pages:
        f = run_page(p)
        print(f"idle  {p:15s} ok  (frames={len(f.frames)})")
    # version 4: pages while a build is running, after a failure and after a bad master (synthetic, no customer data)
    import threading
    import time

    from app_core import jobs

    mgr, gate = jobs.manager(), threading.Event()
    def bad(progress):
        raise runner.InvalidMaster([f"Row {i}: QTY is missing" for i in range(5)])
    mgr.submit("o1", "build", "Building the invoices", "bad.xlsx", bad)
    while mgr.active("o1"):
        time.sleep(0.02)
    mgr.submit("o2", "build", "Building the invoices", "slow.xlsx", lambda p: (p(0.4, "Invoice 3 of 7"), gate.wait(5))[1])
    time.sleep(0.2)
    for owner, what in [("o1", "failed"), ("o2", "running")]:
        for p in ["home.py", "build.py"]:
            f = run_page(p, {"_job_owner": owner})
            print(f"jobs  {p:15s} ok  ({what})")
    gate.set()
    if plan:
        m = runner.build_master(Path(plan).read_bytes(), Path(plan).name)
        run = runner.build_invoices(m["master"], "MASTER_INVOICE.xlsx")
        store.save_consignee("12345", "Acme", "1 Road\nTown")
        store.save_mark("12345", "ACME\nTOWN")
        for p, st in [("home.py", dict(run=run)), ("build.py", dict(run=run)), ("master.py", dict(master_run=m)), ("consignees.py", dict(run=run)), ("marks.py", {})]:
            f = run_page(p, st)
            print(f"data  {p:15s} ok  (frames={len(f.frames)}, rows={sum(len(x) for x in f.frames)}, notes={len(f.messages)})")
