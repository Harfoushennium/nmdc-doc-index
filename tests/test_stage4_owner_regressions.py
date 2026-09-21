from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[1]


class Stage4OwnerRegressionTests(unittest.TestCase):
    def read(self, path: str) -> str:
        return (ROOT / path).read_text(encoding="utf-8")

    def test_setup_reacquires_saved_xlsm_before_named_table_access(self):
        text = self.read("packaging/Create_NMDC_Document_Index.vbs")
        self.assertIn("Set workbook = ReacquireSavedWorkbook(excel, outputWorkbook)", text)
        self.assertIn('TraceStep "workbook-reacquired-after-saveas"', text)
        self.assertIn('TraceStep "ensure-table-start " & sheetName & "!" & tableName', text)

    def test_custom_fields_banner_never_merges_populated_cells(self):
        text = self.read("excel/vba/modNMDC_CustomFields.bas")
        block = 'ws.Range("A4:N4").UnMerge\n    ws.Range("A4:N4").ClearContents\n    ws.Range("A4:N4").Merge'
        self.assertIn(block, text)

    def test_plain_rules_editor_does_not_offer_regex_or_fuzzy(self):
        text = self.read("excel/vba/modNMDC_Startup.bas")
        self.assertIn('"CONTAINS,EXACT,STARTS_WITH,ENDS_WITH"', text)
        self.assertNotIn('"CONTAINS,EXACT,FUZZY,REGEX"', text)

    def test_visible_search_bar_is_owner_reference_activex_input(self):
        setup = self.read("packaging/Create_NMDC_Document_Index.vbs")
        live = self.read("excel/vba/modNMDC_LiveFilter.bas")
        listener = self.read("excel/vba/Cls_LiveFilter_Listener.cls")
        self.assertIn('ImportModule workbook, fso.BuildPath(modulesFolder, "Cls_LiveFilter_Listener.cls")', setup)
        self.assertIn('ClassType:="Forms.TextBox.1"', live)
        self.assertIn('Application.OnKey "^+F", "Ask_User_For_Target_Column"', live)
        self.assertIn("Private Sub SearchBox_Change()", listener)
        self.assertIn("Call Mod_LiveFilter.Run_Live_Filter(SearchBox.Value)", listener)

    def test_pending_update_guidance_matches_below_table_design(self):
        owner = self.read("excel/vba/modNMDC_OwnerUX.bas")
        check = self.read("excel/vba/modNMDC_Checkboxes.bas")
        perf = self.read("excel/vba/modNMDC_Performance.bas")
        combined = owner + "\n" + check + "\n" + perf
        self.assertIn("source-selection section below", combined)
        self.assertNotIn("source panel at the right", combined)
        self.assertNotIn("source panel on the right", combined)
        self.assertNotIn("Esc/Enter stops typing mode", combined)


if __name__ == "__main__":
    unittest.main()
