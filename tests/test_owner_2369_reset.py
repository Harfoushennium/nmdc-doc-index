from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from nmdc_profiler.ooxml import profile_workbook
from nmdc_profiler.runtime_admin import _remove_path_with_retries


ROOT = Path(__file__).resolve().parents[1]


class Owner2369AndResetRegressionTests(unittest.TestCase):
    def test_real_2369_workbook_contains_correct_project_and_document_evidence(self):
        data_root = ROOT / "DATA"
        source = data_root / "METHODS" / "1 Completed Project  Deliverables" / "2369 - NMGL Delivarables.xlsx"
        profile = profile_workbook(source, data_root, data_root)
        self.assertEqual(profile["readability_status"], "READABLE")
        self.assertIn("2369", profile["candidate_internal_project_numbers"])
        sample_docs = {
            document
            for sheet in profile["sheets"]
            for document in sheet.get("sample_document_numbers", [])
        }
        self.assertIn("2369-PP-OF-003", sample_docs)
        self.assertIn("2369-AP-0001", sample_docs)

    def test_excel_uses_native_hyperlinks_not_listobject_calculated_formulas(self):
        text = (ROOT / "excel" / "vba" / "modNMDC_Refresh.bas").read_text(encoding="utf-8")
        self.assertIn("Hyperlinks.Add Anchor:=targetCell", text)
        self.assertIn("Application.AutoCorrect.AutoFillFormulasInLists = False", text)
        self.assertNotIn('formulas(rowIndex, 1) = "=HYPERLINK(', text)
        self.assertNotIn("NMDC_AssignRowSpecificTableFormulas", text)

    def test_reset_retries_a_transient_runtime_lock(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "staging"
            target.mkdir()
            calls = {"count": 0}

            def flaky_rmtree(path, onerror=None):
                calls["count"] += 1
                if calls["count"] == 1:
                    raise PermissionError("simulated OneDrive lock")
                Path(path).rmdir()

            with patch("nmdc_profiler.runtime_admin.shutil.rmtree", side_effect=flaky_rmtree), \
                 patch("nmdc_profiler.runtime_admin.time.sleep", return_value=None):
                _remove_path_with_retries(target, attempts=3)

            self.assertEqual(calls["count"], 2)
            self.assertFalse(target.exists())

    def test_setup_normalizes_merged_ui_ranges_without_formatting_all_excel_columns(self):
        setup = (ROOT / "packaging" / "Create_NMDC_Document_Index.vbs").read_text(encoding="utf-8")
        self.assertIn("NormalizeMergedUiRanges workbook", setup)
        self.assertIn("Sub NormalizeMergedUiRanges", setup)
        self.assertIn('ws.Range("A1:L33").Font.Name = "Aptos"', setup)
        self.assertNotIn('ws.Cells.Font.Name = "Aptos"', setup)


if __name__ == "__main__":
    unittest.main()
