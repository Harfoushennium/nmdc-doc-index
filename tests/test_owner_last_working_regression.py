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

    def test_last_working_home_controls_are_never_dropped(self):
        setup = self._read("packaging/Create_NMDC_Document_Index.vbs")
        admin = self._read("excel/vba/modNMDC_Admin.bas")
        expected = {
            "Undo Last Approval": "NMDC_UndoLastApproval",
            "Reset All Records": "NMDC_ResetAllRecords",
            "Help": "NMDC_OpenHelp",
            "Source Selection": "NMDC_OpenSourceSelection",
        }
        for label, macro in expected.items():
            self.assertIn(f'"{label}"', setup)
            self.assertIn(f'"{macro}"', setup)
        self.assertIn("Public Sub NMDC_ResetAllRecords()", admin)
        self.assertIn("Public Sub NMDC_UndoLastApproval()", admin)

    def test_owner_event_installers_are_zero_line_safe(self):
        for relative in (
            "excel/vba/modNMDC_Checkboxes.bas",
            "excel/vba/modNMDC_CustomFieldsSetup.bas",
        ):
            text = self._read(relative)
            self.assertIn("If codeModule.CountOfLines > 0 Then", text)
            self.assertIn("sourceText = codeModule.Lines(1, codeModule.CountOfLines)", text)

    def test_setup_keeps_fast_scan_and_last_working_recovery_controls_together(self):
        setup = self._read("packaging/Create_NMDC_Document_Index.vbs")
        self.assertIn('"NMDC_UpdateChangedFilesFast"', setup)
        self.assertIn('"NMDC_FullRescanFast"', setup)
        self.assertIn('"NMDC_ResetAllRecords"', setup)
        self.assertIn('"NMDC_UndoLastApproval"', setup)


if __name__ == "__main__":
    unittest.main()
