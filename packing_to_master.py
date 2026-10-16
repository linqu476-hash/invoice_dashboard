#!/usr/bin/env python3
"""Step 1 from the command line: python packing_to_master.py "Packing Plan.xlsx" -o MASTER_INVOICE.xlsx
Same code the dashboard uses (dashboard/extract.py), also available in the UI under 'Packing plan -> Master'."""
import runpy, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / "dashboard"))
runpy.run_module("extract", run_name="__main__")
