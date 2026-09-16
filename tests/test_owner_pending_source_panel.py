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

    def test_source_selection_is_moved_onto_pending_update(self):
        checkboxes = self._read("excel/vba/modNMDC_Checkboxes.bas")
        self.assertIn('Set ws = ThisWorkbook.Worksheets("Pending Update")', checkboxes)
        self.assertIn('Set sourceRange = ws.Range("L5:S6")', checkboxes)
        self.assertIn('"Include in Index?", "Project No.", "Source File"', checkboxes)
        self.assertIn('legacyWs.Visible = xlSheetVeryHidden', checkboxes)
        self.assertIn('homeWs.Shapes("NMDC_Action_19").Delete', checkboxes)
        self.assertIn('button.OnAction = "NMDC_SaveSourceSelections"', checkboxes)

    def test_source_checkboxes_do_not_depend_on_per_checkbox_macro(self):
        checkboxes = self._read("excel/vba/modNMDC_Checkboxes.bas")
        self.assertIn("checkBox.LinkedCell", checkboxes)
        self.assertIn('checkBox.OnAction = ""', checkboxes)
        self.assertIn("Public Sub NMDC_SourceCheckboxClicked()", checkboxes)
        build_section = checkboxes.split("Private Sub NMDC_BuildPendingSourceCheckboxes", 1)[1].split(
            "Public Sub NMDC_SaveSourceSelections", 1
        )[0]
        self.assertNotIn('checkBox.OnAction = "NMDC_SourceCheckboxClicked"', build_section)

    def test_pending_panel_can_reinclude_previously_excluded_sources(self):
        source_view = self._read("nmdc_profiler/source_selection_view.py")
        self.assertIn("list(staged.get(\"records\", []) or [])", source_view)
        self.assertIn("list(approved.get(\"records\", []) or [])", source_view)
        self.assertIn('"FALSE" if excluded else "TRUE"', source_view)


if __name__ == "__main__":
    unittest.main()
