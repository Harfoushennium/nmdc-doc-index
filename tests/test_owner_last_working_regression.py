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
            "modNMDC_CustomFields.bas",
            "modNMDC_CustomFieldsSetup.bas",
        ):
            self.assertIn(f'RequireFile fso.BuildPath(modulesFolder, "{module}")', setup)
            self.assertIn(f'ImportModule workbook, fso.BuildPath(modulesFolder, "{module}")', setup)

    def test_setup_initializes_new_features_before_save(self):
        setup = self._read("packaging/Create_NMDC_Document_Index.vbs")
        self.assertIn("ConfigureOwnerEvents workbook", setup)
        self.assertIn("NMDC_EnsureCustomFieldsStructure", setup)
        self.assertIn("NMDC_CustomFieldsInitialize", setup)
        self.assertIn("NMDC_LiveFilterInitialize", setup)
        self.assertIn("NMDC_ApplyOwnerUX", setup)

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

    def test_remaining_event_installer_is_zero_line_safe_and_live_filter_needs_no_code_injection(self):
        live = self._read("excel/vba/modNMDC_LiveFilter.bas")
        custom = self._read("excel/vba/modNMDC_CustomFieldsSetup.bas")
        self.assertNotIn("TxtBox_Search_Change", live)
        self.assertNotIn("codeModule.AddFromString", live)
        self.assertIn("Application.OnKey", live)
        self.assertIn("If codeModule.CountOfLines > 0 Then", custom)
        self.assertIn("sourceText = codeModule.Lines(1, codeModule.CountOfLines)", custom)

    def test_setup_keeps_fast_scan_and_last_working_recovery_controls_together(self):
        setup = self._read("packaging/Create_NMDC_Document_Index.vbs")
        self.assertIn('"NMDC_UpdateChangedFilesFast"', setup)
        self.assertIn('"NMDC_FullRescanFast"', setup)
        self.assertIn('"NMDC_ResetAllRecords"', setup)
        self.assertIn('"NMDC_UndoLastApproval"', setup)


if __name__ == "__main__":
    unittest.main()
