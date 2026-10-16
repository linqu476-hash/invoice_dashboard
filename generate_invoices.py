#!/usr/bin/env python3
"""Step 2 from the command line: python generate_invoices.py --packing PLAN.xlsx --template assets/invoice_template.xlsx --out out --combine all.xlsx
Same code the dashboard uses (dashboard/engine.py), also available in the UI on the Dashboard page."""
import runpy, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / "dashboard"))
runpy.run_module("engine", run_name="__main__")
