import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class OwnerLiveFilterCustomFieldsTests(unittest.TestCase):
    def _read(self, relative: str) -> str:
        return (ROOT / relative).read_text(encoding="utf-8-sig")

    def test_live_filter_is_table_safe_and_supports_owner_search_syntax(self):
        live = self._read("excel/vba/modNMDC_LiveFilter.bas")

        self.assertIn('Private Const NMDC_LIVE_HELPER As String = "__NMDC_LiveFilter"', live)
        self.assertIn('Application.OnKey "^+F"', live)
        self.assertIn('NMDC_LIVE_ALL_COLUMNS As String = "ALL COLUMNS"', live)
        self.assertIn('Split(Replace(cleanText, "+", " "), " ")', live)
        self.assertIn('If Left$(token, 1) = "-"', live)
        self.assertIn("vbBinaryCompare", live)
        self.assertIn("table.Range.AutoFilter Field:=helper.Index, Criteria1:=True", live)
        self.assertIn("helper.Range.EntireColumn.Hidden = True", live)

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

        # Source Selection deliberately uses row 3 for checkbox status and is a
        # short owner-choice list, so it should not get the heavy live-filter UX.
        table_mapping = live.split("Private Function NMDC_LiveFilterTableForSheet", 1)[1]
        self.assertNotIn('Case "SOURCE SELECTION"', table_mapping)

    def test_live_filter_avoids_activex_dependency_from_owner_concept(self):
        live = self._read("excel/vba/modNMDC_LiveFilter.bas")
        self.assertNotIn("OLEObjects", live)
        self.assertNotIn("MSForms.TextBox", live)
        self.assertNotIn("TxtBox_Search", live)
        self.assertIn('ws.Range("B3:D3").Merge', live)
        self.assertIn('ws.Range("F3:H3").Merge', live)

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

    def test_existing_source_checkbox_module_bootstraps_optional_owner_enhancements(self):
        checkboxes = self._read("excel/vba/modNMDC_Checkboxes.bas")
        self.assertIn("NMDC_EnsureOwnerEnhancements", checkboxes)
        self.assertIn("modNMDC_LiveFilter.bas", checkboxes)
        self.assertIn("modNMDC_CustomFields.bas", checkboxes)
        self.assertIn("modNMDC_CustomFieldsSetup.bas", checkboxes)
        self.assertIn("NMDC_InstallLiveFilterWorkbookEvents", checkboxes)
        self.assertIn("Workbook_SheetChange", checkboxes)
        self.assertIn("Workbook_SheetActivate", checkboxes)

    def test_production_workflow_packages_all_new_vba_modules(self):
        workflow = self._read(".github/workflows/production-package.yml")
        self.assertIn('Copy-Item .\\excel\\vba\\*.bas "$package\\vba\\"', workflow)


if __name__ == "__main__":
    unittest.main()
