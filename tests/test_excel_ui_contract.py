from __future__ import annotations

import json
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class ExcelUIContractTests(unittest.TestCase):
    def test_contract_contains_required_sheets_and_actions(self):
        contract = json.loads((ROOT / "excel" / "contracts" / "ui_contract.json").read_text(encoding="utf-8"))
        required_sheets = {
            "Home",
            "Master Documents",
            "Revisions",
            "Transactions",
            "Pending Update",
            "Review Flags",
            "User Decisions",
            "Configuration",
            "Rules & Mappings",
            "Update History",
            "Error Log",
            "Help",
            "System Data",
        }
        self.assertEqual(set(contract["workbook_sheets"]), required_sheets)
        required_buttons = {
            "Update Changed Files",
            "Full Rescan / Rebuild All",
            "Review Pending Update",
            "Approve Update",
            "Hold Update",
            "Reject Update",
            "Select Data Folder",
            "Refresh Dashboard",
            "Review Flags",
            "Flag Wrong Data",
            "Report Requirement / Problem",
            "View Log",
        }
        self.assertEqual(set(contract["buttons"]), required_buttons)
        self.assertEqual(contract["normal_user_interface"], "Excel only")
        self.assertIn("user-editable in Excel", contract["dynamic_rules_rule"])

    def test_exchange_file_contract_is_stable(self):
        contract = json.loads((ROOT / "excel" / "contracts" / "ui_contract.json").read_text(encoding="utf-8"))
        self.assertEqual(
            contract["exchange_files"],
            [
                "dashboard.csv",
                "master_documents.csv",
                "revisions.csv",
                "events.csv",
                "pending_update.csv",
                "flags.csv",
                "history.csv",
                "errors.csv",
            ],
        )

    def test_vba_launcher_hides_console_and_waits_for_result(self):
        text = (ROOT / "excel" / "vba" / "modNMDC_Engine.bas").read_text(encoding="utf-8")
        self.assertIn('shell.Run(cmd, 0, True)', text)
        self.assertNotIn("cmd.exe", text.lower())
        self.assertNotIn("powershell", text.lower())
        self.assertIn("ENGINE_MISSING", text)
        self.assertIn("Error Log", text)
        self.assertIn('"EXCEL:" & code', text)

    def test_vba_actions_expose_required_button_macros(self):
        text = (ROOT / "excel" / "vba" / "modNMDC_Actions.bas").read_text(encoding="utf-8")
        macros = [
            "NMDC_UpdateChangedFiles",
            "NMDC_FullRescan",
            "NMDC_ReviewPendingUpdate",
            "NMDC_ApproveUpdate",
            "NMDC_HoldUpdate",
            "NMDC_RejectUpdate",
            "NMDC_SelectDataFolder",
            "NMDC_RefreshDashboard",
            "NMDC_ReviewFlags",
            "NMDC_FlagWrongData",
            "NMDC_ReportRequirement",
            "NMDC_ViewLog",
        ]
        for macro in macros:
            self.assertIn(f"Public Sub {macro}", text)
        self.assertIn("The approved master index was not changed", text)

    def test_error_refresh_preserves_excel_local_entries(self):
        text = (ROOT / "excel" / "vba" / "modNMDC_Refresh.bas").read_text(encoding="utf-8")
        self.assertIn("NMDC_LoadErrorCsvPreserveLocal", text)
        self.assertIn('Left$(CStr(ws.Cells(r, 3).Value), 6) = "EXCEL:"', text)
        self.assertNotIn('NMDC_LoadCsvToSheet NMDC_ExchangePath() & "\\errors.csv", "Error Log"', text)

    def test_workbook_spec_states_excel_only_and_staged_approval(self):
        text = (ROOT / "excel" / "WORKBOOK_UI_SPEC.md").read_text(encoding="utf-8")
        self.assertIn("Excel only", text)
        self.assertIn("A staged update never replaces the approved master automatically", text)
        self.assertIn("Approve Update", text)
        self.assertIn("Flag Wrong Data", text)
        self.assertIn("Report Requirement / Problem", text)
        self.assertIn("Rules & Mappings", text)
        self.assertIn("Refreshing engine-exported errors must not erase locally captured Excel errors", text)


if __name__ == "__main__":
    unittest.main()
