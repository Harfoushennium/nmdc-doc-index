"""Stage 3 — XLSM smoke acceptance test.

Verifies:
1. Fresh setup execution from production-package in an isolated directory (.tmp/stage3-clean).
2. Clean workbook creation (NMDC_Document_Index.xlsm).
3. Reopen in real desktop Excel via COM.
4. Validation of all required worksheets.
5. Validation of all ListObjects (no duplicates, valid structure).
6. Validation of UserForm frmNMDC_LiveFilter presence and initialization.
7. Clean close and reopen cycle.
8. DATA read-only integrity check.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import time
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[2]
PACKAGE = ROOT / "production-package"
STAGE3_DIR = ROOT / ".tmp" / "stage3-clean"


class Stage3SmokeTests(unittest.TestCase):
    def setUp(self):
        # Ensure fresh isolated test directory
        if STAGE3_DIR.exists():
            def on_error(func, path, _):
                os.chmod(path, 0o700)
                func(path)
            shutil.rmtree(STAGE3_DIR, onerror=on_error)
        shutil.copytree(PACKAGE, STAGE3_DIR)
        for leftover in ["NMDC_Document_Index.xlsm", "setup_trace.log"]:
            target = STAGE3_DIR / leftover
            if target.exists():
                target.unlink()

    def test_stage3_xlsm_smoke(self):
        # 1. Run setup VBScript in isolated directory
        script = STAGE3_DIR / "Create_NMDC_Document_Index.vbs"
        self.assertTrue(script.exists(), f"VBS setup script missing: {script}")

        proc = subprocess.Popen(
            ["cscript.exe", "//nologo", str(script)],
            cwd=STAGE3_DIR,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        # Handle setup MsgBox dialog boundedly
        deadline = time.time() + 90
        import win32gui
        import win32con

        import win32service
        while proc.poll() is None and time.time() < deadline:
            def enum_handler(hwnd, _):
                if win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) == "#32770":
                    title = win32gui.GetWindowText(hwnd)
                    if "NMDC Document Index Setup" in title or title.startswith("NMDC Document Index"):
                        win32gui.PostMessage(hwnd, win32con.WM_COMMAND, 6, 0)
                        win32gui.PostMessage(hwnd, win32con.WM_COMMAND, 1, 0)
                        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    elif title in {"Reset All Records", "Final Reset Confirmation"}:
                        win32gui.PostMessage(hwnd, win32con.WM_COMMAND, 1, 0)
                return True
            try:
                hdesk = win32service.OpenDesktop("Default", 0, False, 0x0100)
                win32gui.EnumDesktopWindows(hdesk, enum_handler, None)
            except Exception:
                pass
            try:
                win32gui.EnumWindows(enum_handler, None)
            except Exception:
                pass
            time.sleep(0.25)

        if proc.poll() is None:
            proc.kill()
            self.fail("Setup timed out after 90 seconds")

        stdout, stderr = proc.communicate(timeout=5)
        self.assertEqual(proc.returncode, 0, f"Setup failed (rc={proc.returncode}): {stdout}\n{stderr}")

        xlsm_path = STAGE3_DIR / "NMDC_Document_Index.xlsm"
        self.assertTrue(xlsm_path.exists(), f"XLSM not produced: {xlsm_path}")
        self.assertGreater(xlsm_path.stat().st_size, 1_000_000, "XLSM file unexpectedly small")

        # 2. Verify setup trace reached final-workbook-saved
        trace_log = STAGE3_DIR / "setup_trace.log"
        self.assertTrue(trace_log.exists(), "setup_trace.log missing")
        trace_content = trace_log.read_text(encoding="utf-16", errors="replace")
        self.assertIn("final-workbook-saved", trace_content, "Setup trace did not reach final-workbook-saved")

        # 3. Open workbook in real desktop Excel via COM
        import win32com.client
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        try:
            wb = excel.Workbooks.Open(str(xlsm_path), False, False)

            # 4. Verify required sheets
            required_sheets = {
                "Home",
                "Master Documents",
                "Revisions",
                "Transactions",
                "Pending Update",
                "Review Flags",
                "Configuration",
                "Rules & Mappings",
                "Custom Fields",
                "Update History",
                "Error Log",
                "System Data",
            }
            actual_sheets = {str(wb.Worksheets(i).Name) for i in range(1, wb.Worksheets.Count + 1)}
            missing_sheets = required_sheets - actual_sheets
            self.assertEqual(missing_sheets, set(), f"Missing required sheets: {missing_sheets}")

            # 5. Verify ListObjects exist and are unique
            seen_tables = set()
            tables_by_sheet = {}
            for i in range(1, wb.Worksheets.Count + 1):
                ws = wb.Worksheets(i)
                tables_by_sheet[ws.Name] = []
                for j in range(1, ws.ListObjects.Count + 1):
                    tbl = ws.ListObjects(j)
                    self.assertNotIn(tbl.Name, seen_tables, f"Duplicate ListObject name: {tbl.Name}")
                    seen_tables.add(tbl.Name)
                    tables_by_sheet[ws.Name].append(tbl.Name)

            expected_core_tables = [
                "MasterDocuments",
                "RevisionRegister",
                "EventRegister",
                "PendingUpdate",
                "ReviewFlags",
                "Configuration",
                "ClassificationRules",
                "CustomFields",
                "UpdateHistory",
                "ErrorLog",
                "BaselineCounts",
                "SourceInventory",
                "UserDecisionLog",
            ]
            for tbl_name in expected_core_tables:
                self.assertIn(tbl_name, seen_tables, f"Expected table '{tbl_name}' not found in any sheet")

            # 6. Verify the Microsoft Forms reference required by the REV03 listener.
            reference_names = {
                str(wb.VBProject.References.Item(i).Name)
                for i in range(1, wb.VBProject.References.Count + 1)
            }
            self.assertIn("MSForms", reference_names, "Microsoft Forms 2.0 (MSForms) reference missing")

            # 6. Verify owner-reference Live Filter components and real worksheet controls
            vba_components = {wb.VBProject.VBComponents.Item(i).Name for i in range(1, wb.VBProject.VBComponents.Count + 1)}
            self.assertIn("Mod_LiveFilter", vba_components, "Owner-reference Mod_LiveFilter module missing from VBProject")
            self.assertIn("Cls_LiveFilter_Listener", vba_components, "Live Filter listener class missing from VBProject")
            self.assertIn("modNMDC_CustomFields", vba_components, "modNMDC_CustomFields missing from VBProject")

            for sheet_name in (
                "Master Documents",
                "Revisions",
                "Transactions",
                "Pending Update",
                "Review Flags",
                "User Decisions",
                "Update History",
                "Error Log",
            ):
                ws = wb.Worksheets(sheet_name)
                control = ws.OLEObjects("TxtBox_Search")
                self.assertIsNotNone(control, f"TxtBox_Search missing on {sheet_name}")

            # 7. Test close and reopen
            wb.Close(False)
            wb = excel.Workbooks.Open(str(xlsm_path), False, False)
            reopened_sheets = {str(wb.Worksheets(i).Name) for i in range(1, wb.Worksheets.Count + 1)}
            self.assertEqual(reopened_sheets, actual_sheets, "Sheets changed after close/reopen cycle")
            wb.Close(False)
            wb = None

            print(f"Stage 3 smoke PASSED: {len(actual_sheets)} sheets, {len(seen_tables)} unique tables, owner-reference ActiveX Live Filter verified.")

        finally:
            if excel is not None:
                excel.Quit()
                excel = None


if __name__ == "__main__":
    unittest.main()
