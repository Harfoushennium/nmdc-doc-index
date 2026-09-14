from __future__ import annotations

import csv
import json
import tempfile
import unittest
from pathlib import Path

from nmdc_profiler.layout_compat import _extended_document_header
from nmdc_profiler.review_decisions import apply_review_decisions
from nmdc_profiler.runtime_admin import reset_runtime_state, undo_last_approval


ROOT = Path(__file__).resolve().parents[1]


class OwnerFollowupTests(unittest.TestCase):
    def test_refresh_no_longer_creates_or_deletes_temporary_worksheets(self):
        text = (ROOT / "excel" / "vba" / "modNMDC_Refresh.bas").read_text(encoding="utf-8")
        self.assertIn("NMDC_ParseCsvFile", text)
        self.assertNotIn("QueryTables.Add", text)
        self.assertNotIn("Worksheets.Add", text)
        self.assertNotIn("temp.Delete", text)
        self.assertNotIn("NMDC_DeleteTemporarySheet", text)

    def test_csv_parser_handles_quotes_commas_and_embedded_line_breaks_in_memory(self):
        text = (ROOT / "excel" / "vba" / "modNMDC_Csv.bas").read_text(encoding="utf-8")
        self.assertIn("Public Function NMDC_ParseCsvFile", text)
        self.assertIn('If nextCh = Chr$(34) Then', text)
        self.assertIn('ElseIf ch = vbCr Or ch = vbLf Then', text)
        self.assertIn('fieldText = fieldText & " "', text)
        self.assertNotIn("QueryTable", text)

    def test_empty_csv_tables_keep_structure_without_dereferencing_nothing(self):
        refresh = (ROOT / "excel" / "vba" / "modNMDC_Refresh.bas").read_text(encoding="utf-8")
        self.assertIn("If table.ListRows.Count = 0 Then table.ListRows.Add", refresh)
        self.assertIn("If Not table.DataBodyRange Is Nothing Then table.DataBodyRange.ClearContents", refresh)
        self.assertIn("If table.DataBodyRange Is Nothing Then table.ListRows.Add", refresh)
        self.assertNotIn("Else\n        table.DataBodyRange.ClearContents", refresh)

    def test_source_file_hyperlinks_and_review_dropdowns_are_applied(self):
        refresh = (ROOT / "excel" / "vba" / "modNMDC_Refresh.bas").read_text(encoding="utf-8")
        setup = (ROOT / "packaging" / "Create_NMDC_Document_Index.vbs").read_text(encoding="utf-8")
        self.assertIn("NMDC_ActivateSourceLinks table", refresh)
        self.assertIn('NMDC_ActivateSourceColumn table, "Source File"', refresh)
        self.assertIn("Set column = table.ListColumns(columnName)", refresh)
        self.assertIn('ScreenTip:="Open source workbook"', refresh)
        self.assertIn("ACKNOWLEDGED,NO ACTION REQUIRED,NEEDS SOURCE CORRECTION,NEEDS PARSER/MAPPING FIX,HOLD FOR REVIEW", setup)
        self.assertIn("OPEN,ACKNOWLEDGED,RESOLVED,DEFERRED", setup)
        self.assertIn("YES,NO", setup)

    def test_home_has_reset_undo_and_save_review_decision_controls(self):
        setup = (ROOT / "packaging" / "Create_NMDC_Document_Index.vbs").read_text(encoding="utf-8")
        admin = (ROOT / "excel" / "vba" / "modNMDC_Admin.bas").read_text(encoding="utf-8")
        for label in ("Reset All Records", "Undo Last Approval", "Save Review Decisions"):
            self.assertIn(label, setup)
        for macro in ("NMDC_ResetAllRecords", "NMDC_UndoLastApproval", "NMDC_SaveReviewDecisions"):
            self.assertIn(f"Public Sub {macro}", admin)
        self.assertIn("StyleHomeDashboard workbook", setup)
        self.assertIn("ConfigureReviewFlags workbook", setup)

    def test_reset_removes_records_but_preserves_config_and_history(self):
        with tempfile.TemporaryDirectory() as tmp:
            state = Path(tmp)
            for name in ("approved", "staging", "cache", "user_flags", "profile"):
                (state / name).mkdir(parents=True)
                (state / name / "x.txt").write_text("x", encoding="utf-8")
            (state / "config").mkdir()
            (state / "config" / "keep.txt").write_text("keep", encoding="utf-8")
            result = reset_runtime_state(state)
            self.assertEqual(result["decision"], "RESET")
            for name in ("approved", "staging", "cache", "user_flags", "profile"):
                self.assertFalse((state / name).exists())
            self.assertTrue((state / "config" / "keep.txt").exists())
            self.assertTrue((state / "logs" / "history.jsonl").exists())

    def test_undo_restores_previous_approved_pointer(self):
        with tempfile.TemporaryDirectory() as tmp:
            state = Path(tmp)
            approved = state / "approved"
            versions = approved / "versions"
            for run_id, stamp in (("RUN-A", "2026-09-01T10:00:00+00:00"), ("RUN-B", "2026-09-02T10:00:00+00:00")):
                folder = versions / run_id
                folder.mkdir(parents=True)
                (folder / "approval.json").write_text(
                    json.dumps({"run_id": run_id, "approved_at": stamp}), encoding="utf-8"
                )
            approved.mkdir(parents=True, exist_ok=True)
            (approved / "current.json").write_text(
                json.dumps({"run_id": "RUN-B", "version_path": "versions/RUN-B"}), encoding="utf-8"
            )
            result = undo_last_approval(state)
            self.assertEqual(result["undone_run_id"], "RUN-B")
            self.assertEqual(result["restored_run_id"], "RUN-A")
            current = json.loads((approved / "current.json").read_text(encoding="utf-8"))
            self.assertEqual(current["run_id"], "RUN-A")

    def test_undo_refuses_when_no_previous_approved_version_exists(self):
        with tempfile.TemporaryDirectory() as tmp:
            state = Path(tmp)
            approved = state / "approved"
            version = approved / "versions" / "RUN-A"
            version.mkdir(parents=True)
            (version / "approval.json").write_text(
                json.dumps({"run_id": "RUN-A", "approved_at": "2026-09-01T10:00:00+00:00"}), encoding="utf-8"
            )
            (approved / "current.json").write_text(
                json.dumps({"run_id": "RUN-A", "version_path": "versions/RUN-A"}), encoding="utf-8"
            )
            with self.assertRaisesRegex(ValueError, "no previous approved version"):
                undo_last_approval(state)
            current = json.loads((approved / "current.json").read_text(encoding="utf-8"))
            self.assertEqual(current["run_id"], "RUN-A")

    def test_review_decisions_persist_to_latest_stage(self):
        with tempfile.TemporaryDirectory() as tmp:
            state = Path(tmp)
            stage = state / "staging" / "RUN-1"
            stage.mkdir(parents=True)
            (state / "staging" / "latest.json").write_text(json.dumps({"run_id": "RUN-1"}), encoding="utf-8")
            flags = [{
                "code": "UNRECOGNIZED_LAYOUT",
                "source": "METHODS/test.xlsx",
                "source_sheet": "Setup Plans",
                "event_key": "",
                "project_no": "1000",
                "document_no": "",
                "revision": "",
                "resolution_status": "OPEN",
            }]
            (stage / "flags.json").write_text(json.dumps(flags), encoding="utf-8")
            decisions = state / "decisions.csv"
            with decisions.open("w", newline="", encoding="utf-8-sig") as handle:
                writer = csv.DictWriter(handle, fieldnames=[
                    "Flag Code", "Source File", "Source Sheet", "Event Key", "Project No.",
                    "Document No.", "Revision", "User Decision", "User Comment", "Resolution Status",
                ])
                writer.writeheader()
                writer.writerow({
                    "Flag Code": "UNRECOGNIZED_LAYOUT",
                    "Source File": "METHODS/test.xlsx",
                    "Source Sheet": "Setup Plans",
                    "Event Key": "",
                    "Project No.": "1000",
                    "Document No.": "",
                    "Revision": "",
                    "User Decision": "NEEDS PARSER/MAPPING FIX",
                    "User Comment": "Same format as another project",
                    "Resolution Status": "ACKNOWLEDGED",
                })
            result = apply_review_decisions(state, decisions)
            self.assertEqual(result["updated_flags"], 1)
            saved = json.loads((stage / "flags.json").read_text(encoding="utf-8"))[0]
            self.assertEqual(saved["user_decision"], "NEEDS PARSER/MAPPING FIX")
            self.assertEqual(saved["user_comment"], "Same format as another project")
            self.assertEqual(saved["resolution_status"], "ACKNOWLEDGED")

    def test_layout_compatibility_accepts_common_engineering_identifier_headers(self):
        for header in (
            "Procedure No.",
            "Method Statement Number",
            "Setup Plan No",
            "Anchor Pattern Number",
            "Cut List No.",
        ):
            self.assertTrue(_extended_document_header(header), header)


if __name__ == "__main__":
    unittest.main()
