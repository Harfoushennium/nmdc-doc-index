from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from nmdc_profiler.extractor import SheetModel
from nmdc_profiler.full_extractor import (
    FullWorkItem,
    _case_id,
    build_full_worklist,
    run_full_extraction,
)


class FullExtractorTests(unittest.TestCase):
    def test_01_worklist_uses_selected_workbooks_only(self):
        inventory = [
            {"relative_path": "DATA/TECH/1000.xlsx", "selected_excluded_status": "SELECTED", "inferred_project_numbers": "1000", "source_family": "TECH"},
            {"relative_path": "DATA/TECH/2000.xlsx", "selected_excluded_status": "SUPERSEDED", "inferred_project_numbers": "2000", "source_family": "TECH"},
        ]
        discovery = [
            {"workbook_path": "DATA/TECH/1000.xlsx", "worksheet_name": "Documents", "project_number": "1000", "source_family": "TECH", "proposed_action": "INCLUDE"},
            {"workbook_path": "DATA/TECH/2000.xlsx", "worksheet_name": "Documents", "project_number": "2000", "source_family": "TECH", "proposed_action": "INCLUDE"},
        ]
        items, selected, review = build_full_worklist(inventory, discovery)
        self.assertEqual({"DATA/TECH/1000.xlsx"}, set(selected))
        self.assertEqual(1, len(items))
        self.assertEqual("DATA/TECH/1000.xlsx", items[0].source_file)
        self.assertEqual([], review)

    def test_02_section_rows_are_deduplicated_to_one_worksheet(self):
        inventory = [{"relative_path": "DATA/METHODS/2891.xlsx", "selected_excluded_status": "SELECTED", "inferred_project_numbers": "2891", "source_family": "METHODS"}]
        discovery = [
            {"workbook_path": "DATA/METHODS/2891.xlsx", "worksheet_name": "Incomming DOC and DRG", "original_section": "DOCUMENTS", "project_number": "2891", "source_family": "METHODS", "proposed_action": "INCLUDE"},
            {"workbook_path": "DATA/METHODS/2891.xlsx", "worksheet_name": "Incomming DOC and DRG", "original_section": "DRAWINGS", "project_number": "2891", "source_family": "METHODS", "proposed_action": "INCLUDE"},
        ]
        items, _, review = build_full_worklist(inventory, discovery)
        self.assertEqual(1, len(items))
        self.assertEqual([], review)

    def test_03_selected_workbook_without_include_sheet_is_reviewed(self):
        inventory = [{"relative_path": "DATA/TECH/3000.xlsx", "selected_excluded_status": "SELECTED", "inferred_project_numbers": "3000", "source_family": "TECH"}]
        discovery = [{"workbook_path": "DATA/TECH/3000.xlsx", "worksheet_name": "Deleted Documents", "project_number": "3000", "source_family": "TECH", "proposed_action": "EXCLUDE"}]
        items, selected, review = build_full_worklist(inventory, discovery)
        self.assertEqual([], items)
        self.assertEqual(1, len(selected))
        self.assertEqual(1, len(review))
        self.assertEqual("NO_INCLUDED_WORKSHEET_DISCOVERED", review[0]["Reason"])

    def test_04_worklist_uses_inventory_project_as_fallback(self):
        inventory = [{"relative_path": "DATA/TECH/4000.xlsx", "selected_excluded_status": "SELECTED", "inferred_project_numbers": "4000", "source_family": "TECH"}]
        discovery = [{"workbook_path": "DATA/TECH/4000.xlsx", "worksheet_name": "Documents", "project_number": "", "source_family": "TECH", "proposed_action": "INCLUDE"}]
        items, _, _ = build_full_worklist(inventory, discovery)
        self.assertEqual("4000", items[0].project_number)

    def test_05_case_id_is_deterministic(self):
        self.assertEqual(_case_id("DATA/TECH/1.xlsx", "Documents"), _case_id("DATA/TECH/1.xlsx", "Documents"))
        self.assertNotEqual(_case_id("DATA/TECH/1.xlsx", "Documents"), _case_id("DATA/TECH/2.xlsx", "Documents"))

    def test_06_run_full_extraction_uses_validated_project_for_global_key(self):
        item = FullWorkItem("FULL-X", "DATA/TECH/5000.xlsx", "Documents", "5000", "TECH")
        selected = {"DATA/TECH/5000.xlsx": {"relative_path": "DATA/TECH/5000.xlsx"}}
        model = SheetModel("DATA/TECH/5000.xlsx", "TECH", "", "Documents", 1, 1)
        sample_record = {
            "Case ID": "FULL-X",
            "Project No.": "WRONG",
            "Source Family": "TECH",
            "Document No.": "5000-NN-RP-001",
            "Global_Document_Key": "old",
            "Source_Document_Key": "SDOC-1",
            "Revision_Key": "REV-1",
            "Event_Key": "EVT-1",
            "Document_Row_Flag": 1,
            "Revision_Row_Flag": 1,
            "Parsing Status": "INCLUDE",
            "Document Link": "",
            "Source File": "DATA/TECH/5000.xlsx",
            "Source Sheet": "Documents",
            "Source Row": 1,
        }
        recon = {"status": "INCLUDE", "event_records": 1, "rows_with_identity": 1, "distinct_documents": 1, "distinct_revisions": 1, "warnings": []}
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            source = root / item.source_file
            source.parent.mkdir(parents=True)
            source.write_bytes(b"placeholder")
            with patch("nmdc_profiler.full_extractor.resolve_sheet_name", return_value="Documents"), patch("nmdc_profiler.full_extractor.read_sheet_model", return_value=model), patch("nmdc_profiler.full_extractor.extract_model", return_value=([dict(sample_record)], dict(recon))):
                records, reconciliation, review = run_full_extraction(root, [item], selected, [], [])
        self.assertEqual([], review)
        self.assertEqual("5000", records[0]["Project No."])
        self.assertNotEqual("old", records[0]["Global_Document_Key"])
        self.assertEqual(0, reconciliation["summary"]["event_key_duplicates"])

    def test_07_missing_source_is_visible_in_review_queue(self):
        item = FullWorkItem("FULL-X", "DATA/TECH/6000.xlsx", "Documents", "6000", "TECH")
        selected = {"DATA/TECH/6000.xlsx": {"relative_path": "DATA/TECH/6000.xlsx"}}
        with tempfile.TemporaryDirectory() as td:
            records, reconciliation, review = run_full_extraction(Path(td), [item], selected, [], [])
        self.assertEqual([], records)
        self.assertEqual(1, len(review))
        self.assertEqual("SOURCE_FILE_NOT_FOUND", review[0]["Reason"])
        self.assertEqual(1, reconciliation["summary"]["worksheets_review_required"])

    def test_08_nonselected_source_cannot_enter_worklist(self):
        inventory = [{"relative_path": "DATA/TECH/7000.xlsx", "selected_excluded_status": "EXCLUDED", "inferred_project_numbers": "7000", "source_family": "TECH"}]
        discovery = [{"workbook_path": "DATA/TECH/7000.xlsx", "worksheet_name": "Documents", "project_number": "7000", "source_family": "TECH", "proposed_action": "INCLUDE"}]
        items, selected, review = build_full_worklist(inventory, discovery)
        self.assertEqual([], items)
        self.assertEqual({}, selected)
        self.assertEqual([], review)


if __name__ == "__main__":
    unittest.main()
