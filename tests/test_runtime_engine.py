from __future__ import annotations

import csv
import json
import tempfile
import unittest
from pathlib import Path

from nmdc_index_engine import main
from nmdc_profiler.excel_bridge import export_excel_exchange, resolve_document_target
from nmdc_profiler.extractor import read_sheet_model
from nmdc_profiler.full_extractor import resolve_sheet_name
from nmdc_profiler.runtime_engine import record_user_flag, resolution_route
from nmdc_profiler.review_decisions import apply_review_decisions
from nmdc_profiler.update_engine import approve_stage, scan_sources, stage_update


def read_csv(path: Path):
    with path.open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


class RuntimeEngineTests(unittest.TestCase):
    def test_selected_data_folder_root_preserves_source_family(self):
        source = (
            Path(__file__).resolve().parents[1]
            / "DATA"
            / "METHODS"
            / "1 Completed Project  Deliverables"
            / "2369 - NMGL Delivarables.xlsx"
        )
        sheet_name = resolve_sheet_name(source, "Installation Procedures")
        model = read_sheet_model(source, source.parents[2], sheet_name)
        self.assertEqual(model.source_family, "METHODS")

    def test_document_link_is_resolved_from_source_workbook_folder(self):
        with tempfile.TemporaryDirectory() as tmp:
            data = Path(tmp) / "DATA"
            expected = data / "DRAWINGS" / "P-001.pdf"
            actual = resolve_document_target(
                "..\\DRAWINGS\\P-001.pdf",
                "TECH/register.xlsx",
                data,
            )
            self.assertEqual(Path(actual), expected.resolve())
            self.assertEqual(resolve_document_target("https://example.com/P-001.pdf", "TECH/register.xlsx", data), "https://example.com/P-001.pdf")

    def test_user_flag_is_exported_to_excel_review_queue(self):
        with tempfile.TemporaryDirectory() as tmp:
            state = Path(tmp) / "runtime"
            exchange = Path(tmp) / "exchange"
            result = record_user_flag(
                state,
                message="The revision is wrong.",
                project_no="2705",
                document_no="2705-PP-001",
                revision="A1",
                source_file="TECH/2705.xlsx",
                worksheet="Documents",
                source_row="17",
                source_cell="B17",
                event_identity="EVT-1",
                flag_code="",
                current_field="Revision",
                current_value="A1",
                expected_value="1",
                user_name="Owner",
            )
            self.assertEqual(result["code"], "USER_FLAGGED_WRONG_DATA")
            export_excel_exchange(state, exchange)
            flags = read_csv(exchange / "flags.csv")
            self.assertEqual(len(flags), 1)
            self.assertEqual(flags[0]["Document No."], "2705-PP-001")
            self.assertEqual(flags[0]["User Comment"], "The revision is wrong.")

    def test_runtime_resolution_routes_are_deterministic(self):
        self.assertEqual(resolution_route("Project No.", ""), "OWNER_OVERRIDE_REVIEW")
        self.assertEqual(resolution_route("Subcategory", ""), "CONFIGURATION_CHANGE_REVIEW")
        self.assertEqual(resolution_route("Revision", "PARSER_ERROR"), "CODE_CHANGE_REQUIRED")
        self.assertEqual(resolution_route("Revision", ""), "REVIEW_REQUIRED")

    def test_temporary_excel_lock_files_are_not_scanned(self):
        with tempfile.TemporaryDirectory() as tmp:
            data = Path(tmp)
            (data / "~$locked.xlsx").write_bytes(b"temporary")
            (data / "register.xlsx").write_bytes(b"source")
            entries, flags = scan_sources(data)
            self.assertEqual([row["relative_path"] for row in entries], ["register.xlsx"])
            self.assertEqual(flags, [])

    def test_stage_collects_post_processor_review_flags(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            data = root / "DATA"
            state = root / "runtime"
            data.mkdir()
            (data / "register.xlsx").write_bytes(b"placeholder")

            def processor(_path: Path, rel: str):
                return [{"Event_Key": "E1", "Source File": rel, "Document No.": "D1"}]

            result = stage_update(
                data_dir=data,
                state_dir=state,
                processor=processor,
                post_process_flags=lambda: [
                    {
                        "level": "REVIEW",
                        "code": "UNRECOGNIZED_LAYOUT",
                        "message": "Review one unusual worksheet.",
                        "source": "register.xlsx",
                    }
                ],
            )
            flags = json.loads((state / "staging" / result["run_id"] / "flags.json").read_text(encoding="utf-8"))
            self.assertTrue(any(flag["code"] == "UNRECOGNIZED_LAYOUT" for flag in flags))
            self.assertEqual(result["review_flags"], 2)  # SOURCE_NEW plus layout review

            approve_stage(state, result["run_id"])
            reused = stage_update(
                data_dir=data,
                state_dir=state,
                processor=lambda *_: self.fail("unchanged source should use cached records"),
                post_process_flags=lambda: [],
            )
            reused_flags = json.loads(
                (state / "staging" / reused["run_id"] / "flags.json").read_text(encoding="utf-8")
            )
            self.assertTrue(any(flag["code"] == "UNRECOGNIZED_LAYOUT" for flag in reused_flags))

    def test_parser_mapping_review_decision_creates_fix_request_handoff(self):
        with tempfile.TemporaryDirectory() as tmp:
            state = Path(tmp) / "runtime"
            stage = state / "staging" / "RUN-1"
            stage.mkdir(parents=True)
            (state / "staging" / "latest.json").write_text(
                json.dumps({"run_id": "RUN-1"}), encoding="utf-8"
            )
            (stage / "flags.json").write_text(
                json.dumps(
                    [
                        {
                            "level": "REVIEW",
                            "code": "UNRECOGNIZED_LAYOUT",
                            "source": "METHODS/register.xlsx",
                            "source_sheet": "Deliverables",
                            "project_no": "2171",
                            "document_no": "",
                            "revision": "",
                            "event_key": "",
                            "resolution_status": "OPEN",
                        }
                    ]
                ),
                encoding="utf-8",
            )
            decisions = Path(tmp) / "review_decisions.csv"
            decisions.write_text(
                "Flag Code,Source File,Worksheet Name,Event Key,Project No.,Document No.,Revision,User Decision,User Comment,Resolution Status\n"
                "UNRECOGNIZED_LAYOUT,METHODS/register.xlsx,Deliverables,,2171,,,NEEDS PARSER/MAPPING FIX,"
                "Document number is in column C and data starts at row 7,OPEN\n",
                encoding="utf-8-sig",
            )

            result = apply_review_decisions(state, decisions, request_dir=Path(tmp))

            self.assertEqual(result["parser_mapping_fix_requests"], 1)
            self.assertEqual(result["request_sequence"], "0001")
            handoff = Path(tmp) / "PARSER_FIX_REPORTS" / "0001" / "PARSER_FIX_REQUEST.md"
            payload = Path(tmp) / "PARSER_FIX_REPORTS" / "0001" / "PARSER_FIX_REQUEST.json"
            self.assertTrue(handoff.exists())
            self.assertTrue(payload.exists())
            content = handoff.read_text(encoding="utf-8")
            self.assertIn("Worksheet Name", content)
            self.assertIn("METHODS/register.xlsx", content)
            self.assertIn("Deliverables", content)
            self.assertIn("Document number is in column C", content)

            second = apply_review_decisions(state, decisions, request_dir=Path(tmp))
            self.assertEqual(second["request_sequence"], "0002")
            self.assertTrue((Path(tmp) / "PARSER_FIX_REPORTS" / "0002" / "PARSER_FIX_REQUEST.md").exists())

    def test_cli_exports_empty_excel_exchange(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            result = main(
                [
                    "export-excel",
                    "--state-dir",
                    str(root / "runtime"),
                    "--exchange-dir",
                    str(root / "exchange"),
                    "--config-dir",
                    str(root / "config"),
                ]
            )
            self.assertEqual(result, 0)
            self.assertTrue((root / "exchange" / "dashboard.csv").exists())


if __name__ == "__main__":
    unittest.main()
