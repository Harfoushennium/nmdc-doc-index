import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class OwnerLiveFilterCustomFieldsTests(unittest.TestCase):
    def _read(self, relative: str) -> str:
        return (ROOT / relative).read_text(encoding="utf-8-sig")

    def test_live_filter_uses_true_per_keystroke_activex_change_event(self):
        live = self._read("excel/vba/modNMDC_LiveFilter.bas")

        self.assertIn('NMDC_SEARCH_BOX As String = "TxtBox_Search"', live)
        self.assertIn('ClassType:="Forms.TextBox.1"', live)
        self.assertIn('Private Sub TxtBox_Search_Change()', live)
        self.assertIn('NMDC_LiveFilterTextChanged Me', live)
        self.assertIn('Application.OnKey "^+F"', live)
        self.assertIn("NMDC_LiveFilterChooseColumn", live)
        self.assertIn('Criteria1:="=*" & NMDC_LiveFilterEscapeWildcards(cleanText) & "*"', live)
        self.assertIn("vbBinaryCompare", live)

        # The rejected helper-column/all-columns implementation must not return.
        self.assertNotIn("__NMDC_LiveFilter", live)
        self.assertNotIn('"ALL COLUMNS"', live)
        self.assertNotIn("NMDC_LiveFilterRowText", live)
        self.assertNotIn("Split(Replace(cleanText", live)

    def test_live_filter_requires_one_explicit_target_column(self):
        live = self._read("excel/vba/modNMDC_LiveFilter.bas")

        self.assertIn('"Click the HEADER of the column you want to search."', live)
        self.assertIn("targetColumn.DataBodyRange", live)
        self.assertIn("NMDC_LiveFilterSetTarget ws, targetColumn", live)
        self.assertIn('button.TextFrame.Characters.Text = "SELECT COLUMN"', live)
        self.assertIn('button.TextFrame.Characters.Text = "COLUMN: " & targetName', live)

        for sheet_name in (
            "MASTER DOCUMENTS",
            "REVISIONS",
            "TRANSACTIONS",
            "PENDING UPDATE",
            "REVIEW FLAGS",
            "USER DECISIONS",
            "UPDATE HISTORY",
            "ERROR LOG",
        ):
            self.assertIn(sheet_name, live)

        table_mapping = live.split("Private Function NMDC_LiveFilterTableForSheet", 1)[1]
        self.assertNotIn('Case "SOURCE SELECTION"', table_mapping)

    def test_live_filter_reset_matches_owner_reference_behavior(self):
        live = self._read("excel/vba/modNMDC_LiveFilter.bas")
        self.assertIn("Public Sub NMDC_LiveFilterClear()", live)
        self.assertIn("table.AutoFilter.ShowAllData", live)
        self.assertIn('box.Object.Value = ""', live)

    def test_custom_fields_workspace_and_keyword_dictionary_are_user_driven(self):
        setup = self._read("excel/vba/modNMDC_CustomFieldsSetup.bas")
        custom = self._read("excel/vba/modNMDC_CustomFields.bas")

        self.assertIn('ws.Name = "Custom Fields"', setup)
        self.assertIn('"CustomFields"', setup)
        self.assertIn('"KeywordMappings"', setup)
        self.assertIn(
            'Array("Enabled?", "Field Name", "Search In", "Match Behavior", "Separator", "Notes")',
            setup,
        )
        self.assertIn(
            'Array("Enabled?", "Field Name", "Keyword / Pattern", "Result", "Match Type", "Priority", "Notes")',
            setup,
        )

        self.assertIn("NMDC_AddCustomField", custom)
        self.assertIn("NMDC_AddKeywordMapping", custom)
        self.assertIn("NMDC_ApplyCustomFieldsToMaster", custom)
        self.assertIn('"ALL UNIQUE"', custom)
        self.assertIn('"FIRST"', custom)
        self.assertIn('"CONTAINS,ALL TERMS,EXACT,WILDCARD"', custom)
        self.assertIn('Case "ALL TERMS"', custom)
        self.assertIn('Case "EXACT"', custom)
        self.assertIn('Case "WILDCARD"', custom)
        self.assertIn('Replace(Trim$(expressionText), "+", " ")', custom)
        self.assertIn('If Left$(cleanToken, 1) = "-"', custom)
        self.assertIn('starCount = Len(pattern) - Len(Replace(pattern, "*", ""))', custom)
        self.assertIn('Replace(pattern, "?", "")', custom)

    def test_custom_fields_protect_core_master_document_fields(self):
        custom = self._read("excel/vba/modNMDC_CustomFields.bas")
        protected = (
            "FLAG LEVEL",
            "PROJECT NO.",
            "SOURCE FAMILY",
            "DISCIPLINE",
            "CATEGORY",
            "SUBCATEGORY",
            "DOCUMENT NO.",
            "DOCUMENT TITLE",
            "COMPANY DOCUMENT NO.",
            "LATEST REVISION",
            "LATEST EVENT DATE",
            "LATEST EVENT STATUS",
            "DOCUMENT LINK",
            "SOURCE FILE",
            "SOURCE SHEET",
            "SOURCE ROW",
            "SOURCE CELL",
            "GLOBAL DOCUMENT KEY",
        )
        for field in protected:
            self.assertIn(f'"{field}"', custom)
        self.assertIn('Left$(UCase$(Trim$(fieldName)), 7) = "__NMDC_"', custom)

    def test_custom_field_choices_use_checkboxes_and_multi_choice_uses_dropdowns(self):
        custom = self._read("excel/vba/modNMDC_CustomFields.bas")
        self.assertIn("CheckBoxes.Add", custom)
        self.assertIn("NMDC_CustomCheckboxClicked", custom)
        self.assertIn('targetCell.NumberFormat = ";;;"', custom)
        self.assertIn('"Match Behavior", "FIRST,ALL UNIQUE"', custom)
        self.assertIn('"Match Type", "CONTAINS,ALL TERMS,EXACT,WILDCARD"', custom)

    def test_custom_field_configuration_is_backed_up_outside_the_workbook(self):
        custom = self._read("excel/vba/modNMDC_CustomFields.bas")
        self.assertIn('NMDC_CUSTOM_FIELDS_FILE As String = "user_custom_fields.csv"', custom)
        self.assertIn('NMDC_KEYWORD_MAPPINGS_FILE As String = "user_keyword_mappings.csv"', custom)
        self.assertIn("NMDC_RuntimePath()", custom)
        self.assertIn("NMDC_SaveCustomFieldConfiguration", custom)
        self.assertIn("NMDC_LoadCustomFieldConfigurationIfEmpty", custom)

    def test_custom_fields_restore_when_core_refresh_has_removed_derived_columns(self):
        setup = self._read("excel/vba/modNMDC_CustomFieldsSetup.bas")
        self.assertIn("NMDC_EnsureCustomFieldsCurrent", setup)
        self.assertIn("Workbook_SheetSelectionChange", setup)
        self.assertIn('Application.Run "NMDC_ApplyCustomFieldsToMaster", False', setup)

    def test_source_checkbox_module_wakes_persisted_owner_enhancements(self):
        checkboxes = self._read("excel/vba/modNMDC_Checkboxes.bas")
        self.assertIn("NMDC_EnsureOwnerEnhancements", checkboxes)
        self.assertIn('Application.Run "NMDC_LiveFilterWake"', checkboxes)
        self.assertIn('Application.Run "NMDC_CustomFieldsInitialize"', checkboxes)

    def test_production_workflow_packages_all_vba_standard_modules(self):
        workflow = self._read(".github/workflows/production-package.yml")
        self.assertIn('Copy-Item .\\excel\\vba\\*.bas "$package\\vba\\"', workflow)


if __name__ == "__main__":
    unittest.main()
