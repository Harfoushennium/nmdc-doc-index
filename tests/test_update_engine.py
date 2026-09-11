from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from nmdc_profiler.update_engine import (
    approve_stage,
    compare_records,
    hold_stage,
    reject_stage,
    stage_update,
)


class UpdateEngineTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.data = self.root / "external-data-folder"
        self.state = self.root / "runtime-state"
        self.data.mkdir()
        self.calls: list[str] = []

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def processor(self, path: Path, relative_path: str):
        self.calls.append(relative_path)
        payload = path.read_bytes().decode("utf-8")
        return [
            {
                "Project No.": "1000",
                "Document No.": relative_path,
                "Revision": "0",
                "Event_Key": f"{relative_path}-EVENT",
                "Event Status": payload,
            }
        ]

    def _write(self, relative_path: str, content: str) -> Path:
        path = self.data / relative_path
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(content.encode("utf-8"))
        return path

    def _approve_baseline(self) -> None:
        stage_update(data_dir=self.data, state_dir=self.state, processor=self.processor)
        approve_stage(self.state)
        self.calls.clear()

    def test_initial_stage_detects_new_files_and_approval_creates_baseline(self) -> None:
        self._write("A.xlsx", "A1")
        self._write("sub/B.xlsm", "B1")
        summary = stage_update(data_dir=self.data, state_dir=self.state, processor=self.processor)
        self.assertEqual(summary["source_counts"]["NEW"], 2)
        self.assertEqual(summary["staged_records"], 2)
        self.assertEqual(sorted(self.calls), ["A.xlsx", "sub/B.xlsm"])
        self.assertFalse((self.state / "approved" / "manifest.json").exists())
        approve_stage(self.state)
        self.assertTrue((self.state / "approved" / "manifest.json").exists())
        self.assertTrue((self.state / "approved" / "records.jsonl").exists())

    def test_unchanged_sources_reuse_cache_without_processor_call(self) -> None:
        self._write("A.xlsx", "A1")
        self._write("B.xlsx", "B1")
        self._approve_baseline()
        summary = stage_update(data_dir=self.data, state_dir=self.state, processor=self.processor)
        self.assertEqual(self.calls, [])
        self.assertEqual(summary["source_counts"]["UNCHANGED"], 2)
        self.assertEqual(summary["record_counts"]["unchanged"], 2)

    def test_only_changed_source_is_reprocessed(self) -> None:
        self._write("A.xlsx", "A1")
        self._write("B.xlsx", "B1")
        self._approve_baseline()
        self._write("A.xlsx", "A2")
        summary = stage_update(data_dir=self.data, state_dir=self.state, processor=self.processor)
        self.assertEqual(self.calls, ["A.xlsx"])
        self.assertEqual(summary["source_counts"]["CHANGED"], 1)
        self.assertEqual(summary["source_counts"]["UNCHANGED"], 1)
        self.assertEqual(summary["record_counts"]["modified"], 1)

    def test_removed_source_is_staged_as_conflict_and_cannot_silently_approve(self) -> None:
        self._write("A.xlsx", "A1")
        self._write("B.xlsx", "B1")
        self._approve_baseline()
        (self.data / "B.xlsx").unlink()
        summary = stage_update(data_dir=self.data, state_dir=self.state, processor=self.processor)
        self.assertEqual(summary["source_counts"]["REMOVED"], 1)
        self.assertEqual(summary["blocking_flags"], 1)
        self.assertEqual(summary["record_counts"]["removed"], 1)
        with self.assertRaises(ValueError):
            approve_stage(self.state)

    def test_full_rescan_reprocesses_all_present_sources(self) -> None:
        self._write("A.xlsx", "A1")
        self._write("B.xlsx", "B1")
        self._approve_baseline()
        summary = stage_update(
            data_dir=self.data,
            state_dir=self.state,
            processor=self.processor,
            full_rescan=True,
        )
        self.assertEqual(sorted(self.calls), ["A.xlsx", "B.xlsx"])
        self.assertEqual(summary["mode"], "FULL_RESCAN")

    def test_configuration_change_reprocesses_unchanged_sources(self) -> None:
        self._write("A.xlsx", "A1")
        config = self.root / "rules.csv"
        config.write_text("v1", encoding="utf-8")
        stage_update(data_dir=self.data, state_dir=self.state, processor=self.processor, config_paths=[config])
        approve_stage(self.state)
        self.calls.clear()
        config.write_text("v2", encoding="utf-8")
        stage_update(data_dir=self.data, state_dir=self.state, processor=self.processor, config_paths=[config])
        self.assertEqual(self.calls, ["A.xlsx"])

    def test_parser_version_change_reprocesses_unchanged_sources(self) -> None:
        self._write("A.xlsx", "A1")
        stage_update(data_dir=self.data, state_dir=self.state, processor=self.processor, parser_version="v1")
        approve_stage(self.state)
        self.calls.clear()
        stage_update(data_dir=self.data, state_dir=self.state, processor=self.processor, parser_version="v2")
        self.assertEqual(self.calls, ["A.xlsx"])

    def test_hold_and_reject_do_not_replace_approved_dataset(self) -> None:
        self._write("A.xlsx", "A1")
        self._approve_baseline()
        approved_before = (self.state / "approved" / "records.jsonl").read_bytes()
        self._write("A.xlsx", "A2")
        stage_update(data_dir=self.data, state_dir=self.state, processor=self.processor)
        hold_stage(self.state, note="owner wants to inspect")
        self.assertEqual((self.state / "approved" / "records.jsonl").read_bytes(), approved_before)
        reject_stage(self.state, note="wrong source version")
        self.assertEqual((self.state / "approved" / "records.jsonl").read_bytes(), approved_before)

    def test_processor_failure_keeps_previous_approved_rows_and_flags_conflict(self) -> None:
        self._write("A.xlsx", "A1")
        self._approve_baseline()
        self._write("A.xlsx", "A2")

        def broken_processor(path: Path, relative_path: str):
            raise RuntimeError("synthetic parser failure")

        summary = stage_update(data_dir=self.data, state_dir=self.state, processor=broken_processor)
        self.assertEqual(summary["blocking_flags"], 1)
        self.assertEqual(summary["staged_records"], 1)
        latest = json.loads((self.state / "staging" / "latest.json").read_text(encoding="utf-8"))["run_id"]
        flags = json.loads((self.state / "staging" / latest / "flags.json").read_text(encoding="utf-8"))
        self.assertTrue(any(row["code"] == "PARSER_ERROR" for row in flags))

    def test_compare_records_reports_modified_record_with_same_event_key(self) -> None:
        approved = [{"Event_Key": "E1", "Document No.": "D1", "Event Status": "A"}]
        staged = [{"Event_Key": "E1", "Document No.": "D1", "Event Status": "B"}]
        changes, flags = compare_records(approved, staged)
        self.assertEqual(changes["counts"]["modified"], 1)
        self.assertEqual(flags, [])

    def test_arbitrary_external_data_folder_is_supported(self) -> None:
        external = self.root / "somewhere" / "outside" / "project-registers"
        external.mkdir(parents=True)
        (external / "P.xlsx").write_bytes(b"P1")
        calls: list[str] = []

        def proc(path: Path, rel: str):
            calls.append(rel)
            return [{"Event_Key": "P-E1", "Document No.": "P-001"}]

        summary = stage_update(data_dir=external, state_dir=self.state, processor=proc)
        self.assertEqual(summary["source_counts"]["NEW"], 1)
        self.assertEqual(calls, ["P.xlsx"])

    def test_log_records_staging_and_user_decision(self) -> None:
        self._write("A.xlsx", "A1")
        stage_update(data_dir=self.data, state_dir=self.state, processor=self.processor)
        approve_stage(self.state)
        lines = (self.state / "logs" / "history.jsonl").read_text(encoding="utf-8").splitlines()
        events = [json.loads(line)["event"] for line in lines]
        self.assertEqual(events, ["STAGED", "APPROVED"])


if __name__ == "__main__":
    unittest.main()
