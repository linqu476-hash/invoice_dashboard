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
    def __getattr__(self, name):
        if name.startswith("__"):
            raise AttributeError(name)
        return Dummy()


def run_page(name, state=None):
    fake = FakeStreamlit()
    fake.session_state.update(state or {})
    sys.modules["streamlit"] = fake
    for m in [m for m in sys.modules if m.startswith("app_core.ui")]:
        del sys.modules[m]
    runpy.run_path(str(ROOT / "views" / name), run_name="__page__")
    return fake


if __name__ == "__main__":
    from app_core import runner, store

    # synthetic-free: uses the shipped template + a master built from whatever packing plan is given on the command line
    plan = sys.argv[1] if len(sys.argv) > 1 else None
    pages = ["build.py", "master.py", "consignees.py", "marks.py", "data.py"]
    for p in pages:
        f = run_page(p)
        print(f"idle  {p:15s} ok  (frames={len(f.frames)})")
    if plan:
        m = runner.build_master(Path(plan).read_bytes(), Path(plan).name)
        run = runner.build_invoices(m["master"], "MASTER_INVOICE.xlsx")
        store.save_consignee("12345", "Acme", "1 Road\nTown")
        store.save_mark("12345", "ACME\nTOWN")
        for p, st in [("build.py", dict(run=run)), ("master.py", dict(master_run=m)), ("consignees.py", dict(run=run)), ("marks.py", {})]:
            f = run_page(p, st)
            print(f"data  {p:15s} ok  (frames={len(f.frames)}, rows={sum(len(x) for x in f.frames)}, notes={len(f.messages)})")
