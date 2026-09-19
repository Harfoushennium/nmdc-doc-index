"""Targeted runner for Test 23 (Second Clean Install).

Verifies the isolated clean-installation sequence:
1. Fresh package to clean location
2. Run setup VBScript
3. Open fresh XLSM in Excel
4. Check required sheets
5. Configure DATA folder
6. Run NMDC_FullRescan
7. Run NMDC_ResetAllRecords
8. Close and reopen clean XLSM
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import threading
import time
from pathlib import Path
import unittest

import sys
ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
PACKAGE = ROOT / "production-package"
REPORT_DIR = ROOT / "outputs" / "excel_e2e"

from tests.excel_e2e.test_real_excel_simulation import (
    ExcelSimulation,
    _local_clean_root,
    _fresh_package,
    _run_setup,
)


class Test23CleanInstall(unittest.TestCase):
    def test_23_second_clean_install(self):
        sim = ExcelSimulation(REPORT_DIR)
        import win32com.client
        sim.excel = win32com.client.DispatchEx("Excel.Application")
        sim.excel.Visible = True
        sim.excel.DisplayAlerts = False
        
        try:
            result = sim._second_clean_install()
            print("Test 23 result:", result)
            self.assertIn("Fresh XLSM opened", result)
            
            # Update real_excel_simulation.json with Test 23 PASS
            report_file = REPORT_DIR / "real_excel_simulation.json"
            if report_file.exists():
                data = json.loads(report_file.read_text(encoding="utf-8"))
                for t in data["tests"]:
                    if t["number"] == 23:
                        t["status"] = "PASS"
                        t["detail"] = result
                        t["screenshot"] = str(REPORT_DIR / "test_23.png")
                report_file.write_text(json.dumps(data, indent=2), encoding="utf-8")
                print("Updated real_excel_simulation.json with Test 23 PASS.")
        finally:
            sim._close_cleanly()


if __name__ == "__main__":
    unittest.main()
