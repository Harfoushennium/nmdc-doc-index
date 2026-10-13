import csv
import json
import tempfile
import unittest
from pathlib import Path

from nmdc_profiler.source_selection_view import export_source_selection

ROOT = Path(__file__).resolve().parents[1]


class OwnerPendingSourcePanelTests(unittest.TestCase):
    def _read(self, relative: str) -> str:
        return (ROOT / relative).read_text(encoding="utf-8-sig")

    def test_source_selection_exchange_includes_project_number(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            state = root / "runtime"
            stage = state / "staging" / "R1"
            stage.mkdir(parents=True)
            (state / "staging" / "latest.json").write_text(
                json.dumps({"run_id": "R1"}), encoding="utf-8"
            )
            (stage / "manifest.json").write_text(
                json.dumps(
                    {
                        "files": [
                            {
                                "relative_path": "METHODS/2642 HAIL GHASHA.xlsx",
                                "selection_status": "SELECTED",
                                "selection_exclusion_reason": "",
                                "last_processed_run": "R1",
                            }
                        ]
                    }
                ),
                encoding="utf-8",
            )
            (stage / "summary.json").write_text("{}", encoding="utf-8")
            (stage / "flags.json").write_text("[]", encoding="utf-8")
            (stage / "record_changes.json").write_text("{}", encoding="utf-8")
            (stage / "records.jsonl").write_text(
                json.dumps(
                    {
                        "Source File": "METHODS/2642 HAIL GHASHA.xlsx",
                        "Project No.": "2642",
                    }
                )
                + "\n",
                encoding="utf-8",
            )
            config = root / "config"
            config.mkdir()
            (config / "source_exclusions.csv").write_text(
                "Relative_Path,Enabled,Reason\n", encoding="utf-8-sig"
            )
            exchange = root / "exchange"

            target = export_source_selection(state, exchange, config)
            with target.open(newline="", encoding="utf-8-sig") as handle:
                rows = list(csv.DictReader(handle))

            self.assertEqual("2642", rows[0]["Project No."])
            self.assertEqual("METHODS/2642 HAIL GHASHA.xlsx", rows[0]["Source File"])

    def test_file_review_is_independent_of_optional_details(self):
        checkboxes = self._read("excel/vba/modNMDC_Checkboxes.bas")
        setup = self._read("packaging/Create_NMDC_Document_Index.vbs")
        self.assertIn('Set ws = ThisWorkbook.Worksheets("File Review")', checkboxes)
        self.assertIn('"File Review", "FileReview", 5', setup)
        self.assertIn('"Accept Update?", "Project No.", "Source File"', setup)
        self.assertNotIn('pendingTable.Range.Rows.Count', checkboxes)
        self.assertIn('"View Change Details"', checkboxes)
        self.assertIn('"Approve Update"', checkboxes)
        self.assertIn('legacyWs.Visible = xlSheetVeryHidden', checkboxes)

    def test_source_selection_uses_modern_native_excel_checkbox(self):
        checkboxes = self._read("excel/vba/modNMDC_Checkboxes.bas")
        self.assertIn("includeColumn.DataBodyRange.CellControl.SetCheckbox", checkboxes)
        self.assertNotIn("CheckBoxes.Add", checkboxes)
        self.assertNotIn("checkBox.LinkedCell", checkboxes)
        self.assertIn("Public Sub NMDC_SourceCheckboxClicked()", checkboxes)
        self.assertIn("native in-cell checkboxes", checkboxes.lower())

    def test_pending_panel_can_reinclude_previously_excluded_sources(self):
        source_view = self._read("nmdc_profiler/source_selection_view.py")
        self.assertIn("list(staged.get(\"records\", []) or [])", source_view)
        self.assertIn("list(approved.get(\"records\", []) or [])", source_view)
        self.assertIn('"FALSE" if excluded else "TRUE"', source_view)

    def test_pending_guidance_is_merge_safe(self):
        checkboxes = self._read("excel/vba/modNMDC_Checkboxes.bas")
        owner = self._read("excel/vba/modNMDC_OwnerUX.bas")
        self.assertIn('ws.Range("A4:H4").ClearContents', checkboxes)
        self.assertNotIn("DataBodyRange.Merge", checkboxes)
        self.assertIn("NMDC_SafeMergeAndSet", owner)
        self.assertNotIn('ws.Range("A4:I4").Value = _', checkboxes)

    def test_async_polling_is_workbook_qualified_not_sharepoint_url(self):
        performance = self._read("excel/vba/modNMDC_Performance.bas")
        self.assertIn('NMDC_AsyncQualifiedMacro("NMDC_PollEngineAsync")', performance)
        self.assertIn("Procedure:=mAsyncPollProcedure", performance)
        self.assertIn("Application.Run NMDC_AsyncQualifiedMacro(callbackName)", performance)
        self.assertIn("ThisWorkbook.Name", performance)
        self.assertNotIn("ThisWorkbook.FullName", performance)


if __name__ == "__main__":
    unittest.main()
