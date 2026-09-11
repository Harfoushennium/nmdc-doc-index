from __future__ import annotations

import csv
import json
import tempfile
import unittest
from pathlib import Path

from nmdc_profiler.excel_bridge import create_support_request, export_excel_exchange


def write_json(path: Path, payload) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")


def write_jsonl(path: Path, rows) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        for row in rows:
            f.write(json.dumps(row, sort_keys=True) + "\n")


def read_csv(path: Path):
    with path.open(encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


class ExcelBridgeTests(unittest.TestCase):
    def test_empty_runtime_exports_user_friendly_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            state = Path(tmp) / "runtime"
            out = Path(tmp) / "exchange"
            result = export_excel_exchange(state, out)
            self.assertEqual(result["documents"], 0)
            self.assertEqual(result["events"], 0)
            self.assertEqual(read_csv(out / "dashboard.csv")[0]["Approved Status"], "NO APPROVED INDEX")
            for name in (
                "master_documents.csv",
                "revisions.csv",
                "events.csv",
                "pending_update.csv",
                "flags.csv",
                "history.csv",
                "errors.csv",
            ):
                self.assertTrue((out / name).exists(), name)

    def test_versioned_approved_state_builds_document_revision_and_event_views(self):
        with tempfile.TemporaryDirectory() as tmp:
            state = Path(tmp) / "runtime"
            run_id = "RUN-001"
            approved = state / "approved"
            version = approved / "versions" / run_id
            write_json(approved / "current.json", {"run_id": run_id, "version_path": f"versions/{run_id}"})
            write_json(version / "manifest.json", {"run_id": run_id, "data_dir": "C:/NMDC/DATA"})
            write_json(version / "approval.json", {"run_id": run_id, "decision": "APPROVED"})
            records = [
                {
                    "Project No.": "2369",
                    "Source Family": "METHODS",
                    "Discipline": "OFFSHORE INSTALLATION",
                    "Category": "PROCEDURE",
                    "Subcategory": "INSTALLATION PROCEDURE",
                    "Document No.": "2369-PP-OF-003",
                    "Document Title": "Installation Procedure",
                    "Company Document No.": "C-001",
                    "Revision": "A1",
                    "Event Type": "Issue",
                    "Event Date": "2026-01-01",
                    "Event Status": "IFA",
                    "Document Link": "file:///doc.pdf",
                    "Source File": "METHODS/2369.xlsx",
                    "Source Sheet": "Procedures",
                    "Source Row": 10,
                    "Global_Document_Key": "GDOC-1",
                    "Source_Document_Key": "SDOC-1",
                    "Revision_Key": "REV-1",
                    "Event_Key": "EVT-1",
                    "Document_Row_Flag": 1,
                    "Revision_Row_Flag": 1,
                    "Is_Latest_Revision": 0,
                    "Is_Latest_Event": 0,
                    "Parsing Status": "INCLUDE",
                },
                {
                    "Project No.": "2369",
                    "Source Family": "METHODS",
                    "Discipline": "OFFSHORE INSTALLATION",
                    "Category": "PROCEDURE",
                    "Subcategory": "INSTALLATION PROCEDURE",
                    "Document No.": "2369-PP-OF-003",
                    "Document Title": "Installation Procedure",
                    "Company Document No.": "C-001",
                    "Revision": "1",
                    "Event Type": "Issue",
                    "Event Date": "2026-02-01",
                    "Event Status": "IFC",
                    "Document Link": "file:///doc.pdf",
                    "Source File": "METHODS/2369.xlsx",
                    "Source Sheet": "Procedures",
                    "Source Row": 11,
                    "Global_Document_Key": "GDOC-1",
                    "Source_Document_Key": "SDOC-1",
                    "Revision_Key": "REV-2",
                    "Event_Key": "EVT-2",
                    "Document_Row_Flag": 0,
                    "Revision_Row_Flag": 1,
                    "Is_Latest_Revision": 1,
                    "Is_Latest_Event": 1,
                    "Parsing Status": "INCLUDE",
                },
            ]
            write_jsonl(version / "records.jsonl", records)
            out = Path(tmp) / "exchange"
            result = export_excel_exchange(state, out)
            self.assertEqual(result["documents"], 1)
            self.assertEqual(result["revisions"], 2)
            self.assertEqual(result["events"], 2)
            doc = read_csv(out / "master_documents.csv")[0]
            self.assertEqual(doc["Latest Revision"], "1")
            self.assertEqual(doc["Latest Event Status"], "IFC")
            dashboard = read_csv(out / "dashboard.csv")[0]
            self.assertEqual(dashboard["Approved Run ID"], run_id)
            self.assertEqual(dashboard["Approved Documents"], "1")

    def test_staged_changes_and_flags_are_exported_for_user_review(self):
        with tempfile.TemporaryDirectory() as tmp:
            state = Path(tmp) / "runtime"
            approved_run = "RUN-APPROVED"
            approved = state / "approved"
            version = approved / "versions" / approved_run
            write_json(approved / "current.json", {"run_id": approved_run})
            write_json(version / "manifest.json", {"run_id": approved_run})
            old = {
                "Project No.": "2705",
                "Document No.": "2705-PP-001",
                "Revision": "A1",
                "Event Type": "Issue",
                "Source File": "TECH/2705.xlsx",
                "Event_Key": "EVT-1",
                "Revision_Key": "REV-1",
                "Global_Document_Key": "GDOC-1",
                "Document_Row_Flag": 1,
                "Revision_Row_Flag": 1,
                "Parsing Status": "INCLUDE",
            }
            write_jsonl(version / "records.jsonl", [old])

            staged_run = "RUN-STAGED"
            stage = state / "staging" / staged_run
            write_json(state / "staging" / "latest.json", {"run_id": staged_run})
            new = dict(old)
            new["Revision"] = "1"
            write_jsonl(stage / "records.jsonl", [new])
            write_json(
                stage / "record_changes.json",
                {"added": [], "modified": ["EVENT:EVT-1"], "removed": [], "unchanged": [], "counts": {"added": 0, "modified": 1, "removed": 0, "unchanged": 0}},
            )
            write_json(stage / "summary.json", {"run_id": staged_run, "status": "REVIEW_REQUIRED", "record_counts": {"added": 0, "modified": 1, "removed": 0, "unchanged": 0}})
            write_json(
                stage / "flags.json",
                [
                    {"level": "REVIEW", "code": "SOURCE_CHANGED", "message": "Source workbook content changed.", "recommended_action": "Review staged differences.", "source": "TECH/2705.xlsx"},
                    {"level": "CONFLICT", "code": "SOURCE_REMOVED", "message": "Another approved source is missing.", "recommended_action": "Confirm removal.", "source": "TECH/old.xlsx"},
                ],
            )

            out = Path(tmp) / "exchange"
            result = export_excel_exchange(state, out)
            self.assertEqual(result["pending_rows"], 1)
            pending = read_csv(out / "pending_update.csv")[0]
            self.assertEqual(pending["Change Type"], "MODIFIED")
            self.assertEqual(pending["Document No."], "2705-PP-001")
            flags = read_csv(out / "flags.csv")
            self.assertEqual({row["Flag Level"] for row in flags}, {"REVIEW", "CONFLICT"})
            dashboard = read_csv(out / "dashboard.csv")[0]
            self.assertEqual(dashboard["Review Flags"], "1")
            self.assertEqual(dashboard["Conflict Flags"], "1")

    def test_support_request_captures_user_message_and_runtime_ids(self):
        with tempfile.TemporaryDirectory() as tmp:
            state = Path(tmp) / "runtime"
            write_json(state / "approved" / "current.json", {"run_id": "APP-1"})
            version = state / "approved" / "versions" / "APP-1"
            write_json(version / "manifest.json", {"run_id": "APP-1"})
            write_jsonl(version / "records.jsonl", [])
            write_json(state / "staging" / "latest.json", {"run_id": "STAGE-2"})
            (state / "staging" / "STAGE-2").mkdir(parents=True, exist_ok=True)
            path = create_support_request(
                state,
                message="Revision is not being read correctly.",
                source_file="TECH/example.xlsx",
                worksheet="Documents",
                source_row="17",
                document_no="P-001",
                revision="A1",
                current_value="",
                expected_value="A1",
                user_name="Owner",
            )
            payload = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(payload["approved_run_id"], "APP-1")
            self.assertEqual(payload["pending_run_id"], "STAGE-2")
            self.assertEqual(payload["document_no"], "P-001")
            self.assertIn("Revision", payload["message"])


if __name__ == "__main__":
    unittest.main()
