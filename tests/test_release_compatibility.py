"""Release compatibility gates for the v30 package."""
from pathlib import Path
import inspect

import app_core
from app_core import runner

ROOT = Path(__file__).resolve().parents[1]

def test_version_file_and_sidebar_version_are_in_sync():
    assert (ROOT / "VERSION").read_text(encoding="utf-8").strip() == app_core.__version__ == "30.0.0"
    theme = (ROOT / "app_core" / "theme.py").read_text(encoding="utf-8")
    assert "from . import __version__" in theme
    assert "Version {__version__}" in theme

def test_converted_packing_plan_handoff_is_present():
    master_view = (ROOT / "views" / "master.py").read_text(encoding="utf-8")
    assert "build_master_from_converted" in master_view
    assert "tr[\"converted\"]" in master_view
    assert "build_invoices_from_converted_master" in master_view
    assert "def build_master(" in inspect.getsource(runner.build_master)

def test_runner_preserves_optional_combined_workbook_argument():
    assert inspect.signature(runner.build_invoices).parameters["combine"].default is False
    assert inspect.signature(runner._build).parameters["combine"].default is False
