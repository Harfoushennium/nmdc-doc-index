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
            "Source Selection",
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
            "Source Selection",
        }
        self.assertEqual(set(contract["buttons"]), required_buttons)
        self.assertEqual(contract["normal_user_interface"], "Excel only")
        self.assertIn("user-editable in Excel", contract["dynamic_rules_rule"])
        self.assertIn("checkbox", contract["source_selection_rule"].lower())
        self.assertIn("dropdown", contract["user_input_rule"].lower())

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
                "source_selection.csv",
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
        self.assertEqual(
            contract["exchange_columns"]["source_selection.csv"],
            [
                "Include in Index?",
                "Source File",
                "Source Family",
                "Current Status",
                "Owner Note",
                "Selection Reason",
                "Last Processed Run",
            ],
        )
        self.assertIn("exact pending run ID", contract["approval_binding_rule"])
        self.assertIn("do not open or announce", contract["stale_review_rule"])

    def test_vba_launcher_hides_console_keeps_excel_responsive_and_waits_for_result(self):
        text = (ROOT / "excel" / "vba" / "modNMDC_Engine.bas").read_text(encoding="utf-8")
        self.assertIn('shell.Run NMDC_Quote(launcherPath), 0, False', text)
        self.assertIn("Do While Not fso.FileExists(completionPath)", text)
        self.assertIn("NMDC_ShowEngineProgress commandName", text)
        self.assertIn("DoEvents", text)
        self.assertIn('exitCode = CLng(Val(Trim$(NMDC_ReadTextFile(completionPath))))', text)
        self.assertIn("Engine message:", text)
        self.assertNotIn("shell.Exec(cmd)", text)
        self.assertNotIn("cmd.exe", text.lower())
        self.assertNotIn("powershell", text.lower())
        self.assertIn("ENGINE_MISSING", text)
        self.assertIn("Error Log", text)
        self.assertIn('"EXCEL:" & code', text)
        self.assertIn('" --config-dir " & NMDC_Quote(NMDC_ConfigPath())', text)
        self.assertIn('NMDC_RuntimePath() & "\\" & EXCHANGE_RELATIVE_PATH', text)
        self.assertIn('NMDC_ConfigValue("Configuration Folder")', text)
        self.assertIn("Scripting.FileSystemObject", text)
        self.assertIn("NMDC_PathDiagnostics", text)
        self.assertNotIn("Dir$(enginePath)", text)

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
        self.assertNotIn("Dir$(dataFolder", text)
        self.assertIn("NMDC_FolderExists(dataFolder)", text)

    def test_csv_refresh_targets_exact_named_tables_and_preserves_layout(self):
        text = (ROOT / "excel" / "vba" / "modNMDC_Refresh.bas").read_text(encoding="utf-8")
        expected = {
            "master_documents.csv": ("Master Documents", "MasterDocuments"),
            "revisions.csv": ("Revisions", "RevisionRegister"),
            "events.csv": ("Transactions", "EventRegister"),
            "pending_update.csv": ("Pending Update", "PendingUpdate"),
            "flags.csv": ("Review Flags", "ReviewFlags"),
            "history.csv": ("Update History", "UpdateHistory"),
        }
        for csv_name, (sheet, table) in expected.items():
            self.assertIn(f'"\\{csv_name}", "{sheet}", "{table}"', text)
        self.assertIn('ws.ListObjects(tableName)', text)
        self.assertIn("table.Resize targetRange", text)
        self.assertIn("If tableRows < 1 Then tableRows = 1", text)
        self.assertNotIn("ws.Cells.ClearContents", text)
        self.assertNotIn("temp.Cells.ClearContents", text)
        self.assertNotIn('Destination:=ws.Range("A1")', text)
        self.assertIn("NMDC_ActivateDocumentLinks table", text)
        self.assertIn("Hyperlinks.Add Anchor:=targetCell", text)
        self.assertIn("Application.AutoCorrect.AutoFillFormulasInLists = False", text)
        self.assertNotIn('formulas(rowIndex, 1) = "=HYPERLINK(', text)

    def test_csv_refresh_keeps_identifiers_text_but_types_dates_and_numbers(self):
        refresh = (ROOT / "excel" / "vba" / "modNMDC_Refresh.bas").read_text(encoding="utf-8")
        parser = (ROOT / "excel" / "vba" / "modNMDC_Csv.bas").read_text(encoding="utf-8")
        self.assertIn("NMDC_ParseCsvFile", refresh)
        self.assertIn("NMDC_FillCsvArrays", parser)
        self.assertIn('column.DataBodyRange.NumberFormat = "dd-mmm-yyyy"', refresh)
        self.assertIn('column.DataBodyRange.NumberFormat = "dd-mmm-yyyy hh:mm"', refresh)
        self.assertIn('column.DataBodyRange.NumberFormat = "#,##0"', refresh)
        self.assertIn('column.DataBodyRange.NumberFormat = "@"', refresh)
        self.assertIn("NMDC_TryParseDate", refresh)
        self.assertNotIn("QueryTables.Add", refresh)

    def test_configuration_and_error_log_write_through_named_tables(self):
        text = (ROOT / "excel" / "vba" / "modNMDC_Engine.bas").read_text(encoding="utf-8")
        self.assertIn('ws.ListObjects("Configuration")', text)
        self.assertIn('table.ListColumns("Setting")', text)
        self.assertIn('table.ListColumns("Current value")', text)
        self.assertIn('ws.ListObjects("ErrorLog")', text)
        self.assertIn('NMDC_SetTableValue table, row, "Date/Time", Now', text)
        self.assertNotIn('ws.Cells(nextRow, 1).Value = Now', text)

    def test_rules_are_validated_saved_and_logged_through_named_tables(self):
        actions = (ROOT / "excel" / "vba" / "modNMDC_Actions.bas").read_text(encoding="utf-8")
        rules = (ROOT / "excel" / "vba" / "modNMDC_Rules.bas").read_text(encoding="utf-8")
        self.assertIn("If Not NMDC_PrepareRulesForStage() Then Exit Sub", actions)
        self.assertIn('Set table = ws.ListObjects("ClassificationRules")', rules)
        self.assertIn("Duplicate Rule_ID", rules)
        self.assertIn('NMDC_SetConfigValue "Configuration Version"', rules)
        self.assertIn("fso.FolderExists(folderPath)", rules)
        self.assertNotIn("Dir$(folderPath", rules)
        self.assertIn('Set table = ws.ListObjects("UserDecisionLog")', rules)
        self.assertIn('NMDC_SetDecisionValue table, row, "Decision Type", "CONFIGURATION CHANGE"', rules)

    def test_table_aware_wrong_data_capture_uses_actual_listobject_header(self):
        text = (ROOT / "excel" / "vba" / "modNMDC_TableActions.bas").read_text(encoding="utf-8")
        self.assertIn("Public Sub NMDC_FlagWrongDataFromTable", text)
        self.assertIn("Public Sub NMDC_ReportRequirementFromTable", text)
        self.assertIn("Application.InputBox", text)
        self.assertIn('Type:=8', text)
        self.assertIn("For Each table In selectedCell.Worksheet.ListObjects", text)
        self.assertIn("table.HeaderRowRange.Cells", text)
        self.assertIn("table.ListColumns(headerName).Index", text)
        self.assertIn("NMDC_TableHeaderForCell(table, selectedCell)", text)
        self.assertNotIn("ws.Rows(1).Find", text)

    def test_windows_setup_preflights_package_writes_paths_and_guarantees_tables(self):
        text = (ROOT / "packaging" / "Create_NMDC_Document_Index.vbs").read_text(encoding="utf-8")
        for required in [
            "engine\\nmdc_index_engine.exe",
            "classification_rules.csv",
            "project_identity_overrides.csv",
            "source_exclusions.csv",
            "modNMDC_Engine.bas",
            "modNMDC_Csv.bas",
            "modNMDC_Refresh.bas",
            "modNMDC_Actions.bas",
            "modNMDC_TableActions.bas",
            "modNMDC_Admin.bas",
            "modNMDC_Rules.bas",
            "modNMDC_Startup.bas",
            "modNMDC_Performance.bas",
            "modNMDC_OwnerUX.bas",
            "modNMDC_Checkboxes.bas",
        ]:
            self.assertIn(required, text)
        self.assertIn('SetWorkbookConfig workbook, "Engine Executable Path", "engine\\nmdc_index_engine.exe"', text)
        self.assertIn('SetWorkbookConfig workbook, "Runtime Folder", "runtime"', text)
        self.assertIn('SetWorkbookConfig workbook, "Configuration Folder", "config"', text)
        self.assertIn('SetWorkbookConfig workbook, "Parser Version", "cycle3-extractor-v2"', text)
        self.assertNotIn("%LOCALAPPDATA%", text)
        self.assertIn('runtimeFolder = fso.BuildPath(packageRoot, "runtime")', text)
        self.assertIn("EnsureFolderTree runtimeFolder", text)
        self.assertIn("EnsureNamedTables workbook", text)
        expected_tables = [
            "MasterDocuments",
            "RevisionRegister",
            "EventRegister",
            "PendingUpdate",
            "ReviewFlags",
            "SourceSelection",
            "UserDecisionLog",
            "Configuration",
            "ClassificationRules",
            "UpdateHistory",
            "ErrorLog",
            "BaselineCounts",
            "SourceInventory",
        ]
        for table_name in expected_tables:
            self.assertIn(f'"{table_name}"', text)
        self.assertIn("If table.ListRows.Count = 0 Then table.ListRows.Add", text)
        self.assertIn('ws.ListObjects("Configuration")', text)
        self.assertIn('"NMDC_FlagWrongDataFromTable"', text)
        self.assertIn('"NMDC_ReportRequirementFromTable"', text)
        self.assertIn('"Include in Index?"', text)

    def test_source_selection_uses_native_in_cell_checkboxes_but_multi_choice_inputs_stay_dropdowns(self):
        checkboxes = (ROOT / "excel" / "vba" / "modNMDC_Checkboxes.bas").read_text(encoding="utf-8")
        setup = (ROOT / "packaging" / "Create_NMDC_Document_Index.vbs").read_text(encoding="utf-8")
        self.assertIn("CellControl.SetCheckbox", checkboxes)
        self.assertNotIn("CheckBoxes.Add", checkboxes)
        self.assertIn("NMDC_SaveSourceSelections", checkboxes)
        self.assertIn("save-source-selections", checkboxes)
        self.assertIn("NMDC_CheckAllSources", checkboxes)
        self.assertIn("ACKNOWLEDGED,NO ACTION REQUIRED", setup)
        self.assertIn("OPEN,ACKNOWLEDGED,RESOLVED,DEFERRED", setup)

    def test_error_refresh_preserves_excel_local_entries(self):
        text = (ROOT / "excel" / "vba" / "modNMDC_Refresh.bas").read_text(encoding="utf-8")
        self.assertIn("NMDC_LoadErrorCsvPreserveLocal", text)
        self.assertIn('table.ListColumns("Action").Index', text)
        self.assertIn('Left$(CStr(row.Range.Cells(1, actionColumn).Value), 6) = "EXCEL:"', text)
        self.assertIn('NMDC_LoadCsvToTable(csvPath, "Error Log", "ErrorLog")', text)

    def test_dashboard_refresh_does_not_create_or_clear_support_sheets(self):
        text = (ROOT / "excel" / "vba" / "modNMDC_Refresh.bas").read_text(encoding="utf-8")
        self.assertIn("NMDC_ParseCsvFile(csvPath, headers, data, dataCount, columnCount)", text)
        self.assertNotIn("Worksheets.Add", text)
        self.assertNotIn("QueryTables.Add", text)
        self.assertNotIn('Set temp = ThisWorkbook.Worksheets("System Data")', text)
        self.assertNotIn("temp.Cells.ClearContents", text)
        self.assertNotIn("temp.Delete", text)

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
