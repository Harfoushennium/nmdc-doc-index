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
            "Configuration",
            "Rules & Mappings",
            "Report Requirement / Problem",
            "View Log",
            "Help",
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
        self.assertEqual(
            contract["exchange_columns"]["flags.csv"],
            [
                "Flag Level",
                "Flag Code",
                "Plain-English Problem",
                "Recommended User Action",
                "Project No.",
                "Document No.",
                "Revision",
                "Source File",
                "Source Sheet",
                "Source Row",
                "Source Cell",
                "User Decision",
                "User Comment",
                "Resolution Status",
                "Event Key",
            ],
        )
        self.assertIn("exact pending run ID", contract["approval_binding_rule"])
        self.assertIn("do not open or announce", contract["stale_review_rule"])

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
            "NMDC_OpenConfiguration",
            "NMDC_OpenRulesMappings",
            "NMDC_ReportRequirement",
            "NMDC_ViewLog",
            "NMDC_OpenHelp",
        ]
        for macro in macros:
            self.assertIn(f"Public Sub {macro}", text)
        self.assertIn("The approved master index was not changed", text)
        self.assertIn('approvalArgs = "--run-id " & NMDC_Quote(reviewedRunId)', text)
        self.assertIn('approvalArgs = approvalArgs & " --allow-conflicts"', text)
        self.assertIn('NMDC_RunEngine("approve", approvalArgs)', text)
        self.assertIn("duplicate-key conflicts can never be overridden", text)
        self.assertIn("The pending update changed after your last review", text)
        self.assertIn('currentStatus <> "STAGED" And currentStatus <> "REVIEW_REQUIRED" And currentStatus <> "HOLD"', text)
        self.assertIn("The proposed update was staged, but Excel could not load it for review", text)
        self.assertIn('NMDC_GoToSheet "Error Log"', text)
        self.assertNotIn("Or Not NMDC_RefreshExchangeData", text)

    def test_csv_refresh_preserves_identifiers_as_text(self):
        text = (ROOT / "excel" / "vba" / "modNMDC_Refresh.bas").read_text(encoding="utf-8")
        self.assertIn(".TextFileColumnDataTypes = NMDC_TextColumnTypes(csvPath)", text)
        self.assertIn("dataTypes(index) = xlTextFormat", text)
        self.assertIn("revisions such as 00", text)

    def test_error_refresh_preserves_excel_local_entries(self):
        text = (ROOT / "excel" / "vba" / "modNMDC_Refresh.bas").read_text(encoding="utf-8")
        self.assertIn("NMDC_LoadErrorCsvPreserveLocal", text)
        self.assertIn('Left$(CStr(ws.Cells(r, 3).Value), 6) = "EXCEL:"', text)
        self.assertNotIn('NMDC_LoadCsvToSheet NMDC_ExchangePath() & "\\errors.csv", "Error Log"', text)
        self.assertIn("ReDim rowValues(1 To 10)", text)

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
