import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


class OwnerLastWorkingRegressionTests(unittest.TestCase):
    def _read(self, relative: str) -> str:
        return (ROOT / relative).read_text(encoding="utf-8-sig")

    def test_setup_does_not_read_zero_line_thisworkbook_module(self):
        setup = self._read("packaging/Create_NMDC_Document_Index.vbs")
        self.assertIn("If thisModule.CountOfLines > 0 Then", setup)
        self.assertIn("sourceText = thisModule.Lines(1, thisModule.CountOfLines)", setup)
        self.assertNotIn(
            'If InStr(1, thisModule.Lines(1, thisModule.CountOfLines), "Workbook_Open", 1) = 0 Then',
            setup,
        )

    def test_setup_embeds_all_owner_feature_modules_in_fresh_workbook(self):
        setup = self._read("packaging/Create_NMDC_Document_Index.vbs")
        for module in (
            "modNMDC_LiveFilter.bas",
            "Cls_LiveFilter_Listener.cls",
            "modNMDC_CustomFields.bas",
            "modNMDC_CustomFieldsSetup.bas",
        ):
            self.assertIn(f'RequireFile fso.BuildPath(modulesFolder, "{module}")', setup)
            self.assertIn(f'ImportModule workbook, fso.BuildPath(modulesFolder, "{module}")', setup)
        self.assertIn('RequireFile fso.BuildPath(modulesFolder, "Cls_LiveFilter_Listener.cls")', setup)
        self.assertIn('ImportModule workbook, fso.BuildPath(modulesFolder, "Cls_LiveFilter_Listener.cls")', setup)

        assembler = self._read("tests/excel_e2e/assemble_package.py")
        self.assertIn('glob("*.cls")', assembler)

    def test_setup_initializes_new_features_before_save(self):
        setup = self._read("packaging/Create_NMDC_Document_Index.vbs")
        self.assertIn("ConfigureOwnerEvents workbook", setup)
        self.assertIn("NMDC_EnsureCustomFieldsStructure", setup)
        self.assertIn("NMDC_CustomFieldsInitialize", setup)
        self.assertIn("NMDC_LiveFilterInitialize", setup)
        self.assertIn("NMDC_ApplyOwnerUX", setup)

    def test_setup_reacquires_saved_workbook_and_reports_named_table_context(self):
        setup = self._read("packaging/Create_NMDC_Document_Index.vbs")
        self.assertIn("excel.Workbooks(fso.GetFileName(outputWorkbook))", setup)
        self.assertIn('setupStage = "EnsureNamedTables"', setup)
        self.assertIn('setupObject = sheetName & "!" & tableName', setup)
        self.assertIn('TraceStep "setup-error stage=" & setupStage', setup)

    def test_last_working_recovery_controls_are_never_dropped(self):
        setup = self._read("packaging/Create_NMDC_Document_Index.vbs")
        admin = self._read("excel/vba/modNMDC_Admin.bas")
        expected = {
            "Undo Last Approval": "NMDC_UndoLastApproval",
            "Reset All Records": "NMDC_ResetAllRecords",
            "Help": "NMDC_OpenHelp",
        }
        for label, macro in expected.items():
            self.assertIn(f'"{label}"', setup)
            self.assertIn(f'"{macro}"', setup)
        self.assertIn("Public Sub NMDC_ResetAllRecords()", admin)
        self.assertIn("Public Sub NMDC_UndoLastApproval()", admin)

    def test_source_selection_is_integrated_without_dropping_backend_compatibility(self):
        checkboxes = self._read("excel/vba/modNMDC_Checkboxes.bas")
        self.assertIn('ThisWorkbook.Worksheets("Pending Update")', checkboxes)
        self.assertIn('SOURCE_PANEL_TABLE As String = "SourceSelection"', checkboxes)
        self.assertIn('legacyWs.Visible = xlSheetVeryHidden', checkboxes)
        self.assertIn('homeWs.Shapes("NMDC_Action_19").Delete', checkboxes)
        self.assertIn("Public Sub NMDC_SourceCheckboxClicked()", checkboxes)

    def test_remaining_event_installer_is_zero_line_safe_and_live_filter_matches_owner_reference(self):
        live = self._read("excel/vba/modNMDC_LiveFilter.bas")
        listener = self._read("excel/vba/Cls_LiveFilter_Listener.cls")
        custom = self._read("excel/vba/modNMDC_CustomFieldsSetup.bas")
        self.assertIn('ClassType:="Forms.TextBox.1"', live)
        self.assertIn('Application.OnKey "^+F", "Ask_User_For_Target_Column"', live)
        self.assertIn('SEARCH_BOX_NAME As String = "TxtBox_Search"', live)
        self.assertIn("Private Sub SearchBox_Change()", listener)
        self.assertIn("Run_Live_Filter(SearchBox.Value)", listener)
        self.assertIn("If codeModule.CountOfLines > 0 Then", custom)
        self.assertIn("sourceText = codeModule.Lines(1, codeModule.CountOfLines)", custom)

    def test_review_flags_parser_fix_is_actionable_not_a_dead_end(self):
        admin = self._read("excel/vba/modNMDC_Admin.bas")
        setup = self._read("packaging/Create_NMDC_Document_Index.vbs")
        review = self._read("nmdc_profiler/review_decisions.py")
        self.assertIn("Public Sub NMDC_RequestParserMappingFix()", admin)
        self.assertIn("Public Sub NMDC_RetryAfterParserMappingFix()", admin)
        self.assertIn("PARSER_FIX_REQUEST_LATEST.md", admin)
        self.assertIn("Public Sub NMDC_ReportSelectedParserFixes()", admin)
        self.assertIn('"Report Selected Parser Fix"', setup)
        self.assertIn('"Retry After Fix"', setup)
        self.assertIn("parser_mapping_fix_requests", review)

    def test_packaged_engine_uses_shared_parser_cache_version(self):
        engine = self._read("nmdc_index_engine.py")
        update = self._read("nmdc_profiler/update_engine.py")
        self.assertIn('DEFAULT_PARSER_VERSION = "cycle3-extractor-v2"', update)
        self.assertIn("default=DEFAULT_PARSER_VERSION", engine)
        self.assertNotIn('default="cycle3-extractor-v1"', engine)

    def test_setup_registers_msforms_before_importing_rev03_listener(self):
        setup = self._read("packaging/Create_NMDC_Document_Index.vbs")
        ensure_pos = setup.index("EnsureMSFormsReference workbook")
        listener_pos = setup.index('ImportModule workbook, fso.BuildPath(modulesFolder, "Cls_LiveFilter_Listener.cls")')
        self.assertGreaterEqual(ensure_pos, 0)
        self.assertGreater(listener_pos, ensure_pos)
        self.assertIn('{0D452EE1-E08F-101A-852E-02608C4D0BB4}', setup)
        self.assertIn("VBComponents.Add(vbext_ct_MSForm)", setup)
        self.assertIn('TraceStep "msforms-reference-ready"', setup)

    def test_sheet_activate_event_does_not_reference_nonexistent_target_argument(self):
        setup = self._read("packaging/Create_NMDC_Document_Index.vbs")
        activate_block = setup.split('Private Sub Workbook_SheetActivate(ByVal Sh As Object)', 1)[1].split('End Sub', 1)[0]
        self.assertIn("NMDC_LiveFilterSheetActivate Sh", activate_block)
        self.assertNotIn("NMDC_LiveFilterSelectionChange Sh, Target", activate_block)

    def test_rev03_activex_creation_is_deferred_out_of_hidden_setup(self):
        live = self._read("excel/vba/modNMDC_LiveFilter.bas")
        setup = self._read("packaging/Create_NMDC_Document_Index.vbs")
        init = live.split("Public Sub NMDC_LiveFilterInitialize()", 1)[1].split(
            "Private Sub NMDC_EnsureReferenceLiveFilterControls", 1
        )[0]
        self.assertIn("If Not Application.Visible Then Exit Sub", init)
        self.assertNotIn("For Each item In sheetNames", init)
        self.assertIn("NMDC_EnsureReferenceLiveFilterControls Sh", live)
        self.assertIn("excel.Visible = False", setup)
        self.assertIn("NMDC_LiveFilterInitialize", setup)

    def test_sheet_activate_does_not_reactivate_same_sheet(self):
        live = self._read("excel/vba/modNMDC_LiveFilter.bas")
        block = live.split("Public Sub NMDC_LiveFilterSheetActivate", 1)[1].split("End Sub", 1)[0]
        self.assertIn("NMDC_EnsureReferenceLiveFilterControls Sh", block)
        self.assertNotIn("Sh.Activate", block)

    def test_runtime_and_parser_reports_are_package_local(self):
        setup = self._read("packaging/Create_NMDC_Document_Index.vbs")
        engine = self._read("nmdc_index_engine.py")
        review = self._read("nmdc_profiler/review_decisions.py")
        bridge = self._read("nmdc_profiler/excel_bridge.py")
        self.assertIn('runtimeFolder = fso.BuildPath(packageRoot, "runtime")', setup)
        self.assertNotIn("%LOCALAPPDATA%", setup)
        self.assertIn('SetWorkbookConfig workbook, "Runtime Folder", "runtime"', setup)
        self.assertIn('SetWorkbookConfig workbook, "Parser Version", "cycle3-extractor-v2"', setup)
        self.assertIn('review.add_argument("--request-dir", type=Path, default=None)', engine)
        self.assertIn('PARSER_FIX_REQUEST_LATEST.md', review)
        self.assertIn('"Select?",', bridge)

    def test_review_flags_use_checkbox_selection_and_visible_parser_report(self):
        admin = self._read("excel/vba/modNMDC_Admin.bas")
        setup = self._read("packaging/Create_NMDC_Document_Index.vbs")
        refresh = self._read("excel/vba/modNMDC_Refresh.bas")
        layout = self._read("nmdc_profiler/ui_layout.py")
        self.assertIn('"Select?",', layout)
        self.assertIn('table.ListColumns("Select?").DataBodyRange', refresh)
        self.assertIn("CellControl.SetCheckbox", refresh)
        self.assertIn("Public Sub NMDC_ReportSelectedParserFixes()", admin)
        self.assertIn("PARSER_FIX_REQUEST_LATEST.md", admin)
        self.assertIn("NMDC_RevealFile requestPath", admin)
        self.assertIn('"Report Selected Parser Fix"', setup)
        self.assertIn('"Select All"', setup)
        self.assertIn('"Clear Selection"', setup)

    def test_native_m365_checkboxes_are_required_not_silent_fallbacks(self):
        source = self._read("excel/vba/modNMDC_Checkboxes.bas")
        refresh = self._read("excel/vba/modNMDC_Refresh.bas")
        requirements = self._read("USER_PRODUCT_REQUIREMENTS.md")
        acceptance = self._read("docs/RELEASE_ACCEPTANCE_TEST_PLAN.md")

        self.assertIn("CellControl.SetCheckbox", source)
        self.assertIn("NATIVE_CHECKBOX_REQUIRED", source)
        self.assertIn("does not fall back to TRUE/FALSE text", source)
        self.assertNotIn("TRUE/FALSE source choices remain usable", source)

        self.assertIn("CellControl.SetCheckbox", refresh)
        self.assertIn("REVIEW_CHECKBOX_REQUIRED", refresh)
        self.assertIn("does not fall back to TRUE/FALSE text", refresh)
        self.assertNotIn("TRUE/FALSE selection values remain usable", refresh)

        self.assertIn("must not silently downgrade", requirements)
        self.assertIn("release-blocking UI error", acceptance)

    def test_packaged_readme_matches_package_local_storage_contract(self):
        readme = self._read("README.md")
        self.assertNotIn("%LOCALAPPDATA%", readme)
        self.assertIn("PARSER_FIX_REQUEST_LATEST.md", readme)
        self.assertIn("directly beside `NMDC_Document_Index.xlsm`", readme)
        self.assertIn("does not configure AppData", readme)
        self.assertIn("must not silently downgrade", readme)

    def test_production_package_readme_and_payload_match_rev03_contract(self):
        readme = self._read("packaging/README.md")
        workflow = self._read(".github/workflows/production-package.yml")
        assembler = self._read("tests/excel_e2e/assemble_package.py")
        self.assertNotIn("%LOCALAPPDATA%", readme)
        self.assertIn("TxtBox_Search", readme)
        self.assertIn("Cls_LiveFilter_Listener", readme)
        self.assertIn("PARSER_FIX_REQUEST_LATEST.md", readme)
        self.assertIn("directly beside `NMDC_Document_Index.xlsm`", readme)
        self.assertIn("does not silently downgrade", readme)
        self.assertIn('Remove-Item "$package\\vba\\frmNMDC_LiveFilter.frm"', workflow)
        self.assertIn('Remove-Item "$package\\vba\\frmNMDC_LiveFilter.frx"', workflow)
        self.assertIn('frmnmdc_livefilter.frm', assembler)
        self.assertIn('frmnmdc_livefilter.frx', assembler)

    def test_production_workflow_has_one_clean_artifact_upload(self):
        workflow = self._read(".github/workflows/production-package.yml")
        self.assertEqual(workflow.count("actions/upload-artifact@v4"), 1)
        self.assertNotIn("} | Copy-Item -Destination", workflow)
        self.assertEqual(workflow.count("Compress-Archive -Path"), 1)

    def test_setup_keeps_fast_scan_and_last_working_recovery_controls_together(self):
        setup = self._read("packaging/Create_NMDC_Document_Index.vbs")
        self.assertIn('"NMDC_UpdateChangedFilesFast"', setup)
        self.assertIn('"NMDC_FullRescanFast"', setup)
        self.assertIn('"NMDC_ResetAllRecords"', setup)
        self.assertIn('"NMDC_UndoLastApproval"', setup)


if __name__ == "__main__":
    unittest.main()
