import csv
import json
import tempfile
import unittest
from pathlib import Path

from nmdc_profiler.file_review import FIELDS, affected_files, export_file_review, save_file_review
from nmdc_profiler.update_engine import approve_stage, stage_update


class FileReviewTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.data = self.root / "data"
        self.data.mkdir()
        self.state = self.root / "runtime"
        self.calls = []

    def write(self, source, value):
        (self.data / source).write_text(value)

    def process(self, path, source):
        self.calls.append(source)
        return [{"Event_Key": source, "Source File": source, "Event Status": path.read_text()}]

    def stage(self, **kwargs):
        return stage_update(data_dir=self.data, state_dir=self.state, processor=self.process, **kwargs)["run_id"]

    def baseline(self):
        self.write("A.xlsx", "old A")
        self.write("B.xlsx", "old B")
        approve_stage(self.state, self.stage())

    def rows(self):
        pointer = json.loads((self.state / "approved/current.json").read_text())
        path = self.state / "approved/versions" / pointer["run_id"] / "records.jsonl"
        return {r["Source File"]: r["Event Status"] for r in map(json.loads, path.read_text().splitlines())}

    def save(self, run, deferred=()):
        rows = affected_files(self.state, run)
        for row in rows:
            row["Accept Update?"] = row["Source File"] not in deferred
            row["Owner Note"] = 'Keep old, "review later"\nsecond line'
        path = self.root / "choices.csv"
        with path.open("w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.DictWriter(handle, fieldnames=FIELDS)
            writer.writeheader()
            writer.writerows(rows)
        return save_file_review(self.state, run, path)

    def test_mixed_approval_retains_old_rows_and_reoffers_deferred_change(self):
        self.baseline()
        self.write("A.xlsx", "new A")
        self.write("B.xlsx", "new B")
        run = self.stage()
        pointer = (self.state / "approved/current.json").read_bytes()
        self.save(run, ["A.xlsx"])
        self.assertEqual((self.state / "approved/current.json").read_bytes(), pointer)
        approve_stage(self.state, run, require_file_review=True)
        self.assertEqual(self.rows(), {"A.xlsx": "old A", "B.xlsx": "new B"})
        next_run = self.stage()
        self.assertEqual([r["Source File"] for r in affected_files(self.state, next_run)], ["A.xlsx"])

    def test_all_deferred_preserves_approved_rows(self):
        self.baseline()
        self.write("A.xlsx", "new A")
        run = self.stage()
        self.save(run, ["A.xlsx"])
        approve_stage(self.state, run, require_file_review=True)
        self.assertEqual(self.rows(), {"A.xlsx": "old A", "B.xlsx": "old B"})

    def test_deferred_new_file_adds_no_records(self):
        self.baseline()
        self.write("C.xlsx", "new C")
        run = self.stage()
        self.save(run, ["C.xlsx"])
        approve_stage(self.state, run, require_file_review=True)
        self.assertNotIn("C.xlsx", self.rows())
        self.assertEqual(affected_files(self.state, self.stage())[0]["Source File"], "C.xlsx")

    def test_deferred_removed_file_retains_data_and_conflict_still_blocks(self):
        self.baseline()
        (self.data / "A.xlsx").unlink()
        run = self.stage()
        self.save(run, ["A.xlsx"])
        with self.assertRaises(ValueError):
            approve_stage(self.state, run, require_file_review=True)
        approve_stage(self.state, run, require_file_review=True, allow_conflicts=True)
        self.assertEqual(self.rows()["A.xlsx"], "old A")

    def test_unsaved_choices_block_and_empty_affected_list_can_be_saved(self):
        self.baseline()
        run = self.stage()
        self.assertEqual(affected_files(self.state, run), [])
        with self.assertRaisesRegex(ValueError, "Save File Choices"):
            approve_stage(self.state, run, require_file_review=True)
        self.save(run)
        approve_stage(self.state, run, require_file_review=True)

    def test_replacement_run_rejects_old_choices_and_retains_review(self):
        self.baseline()
        self.write("A.xlsx", "new A")
        run = self.stage()
        self.save(run, ["A.xlsx"])
        path = self.state / "staging" / run / "file_review.json"
        before = path.read_bytes()
        self.stage()
        with self.assertRaisesRegex(ValueError, "replaced"):
            approve_stage(self.state, run, require_file_review=True)
        self.assertEqual(path.read_bytes(), before)

    def test_proposal_tamper_invalidates_saved_choices(self):
        run = self.stage()
        self.save(run)
        path = self.state / "staging" / run / "records.jsonl"
        path.write_text('\n')
        with self.assertRaisesRegex(ValueError, "stale"):
            approve_stage(self.state, run, require_file_review=True)

    def test_notes_and_choice_export_survive_reload(self):
        self.write("A.xlsx", "A")
        run = self.stage()
        review = self.save(run, ["A.xlsx"])
        export_file_review(self.state, self.root, run)
        with (self.root / "file_review.csv").open(encoding="utf-8-sig", newline="") as handle:
            row = next(csv.DictReader(handle))
        self.assertEqual(row["Accept Update?"], "False")
        self.assertEqual(row["Owner Note"], review["choices"]["A.xlsx"]["note"])

    def test_partial_parser_upgrade_reprocesses_deferred_file_next_time(self):
        self.baseline()
        run = self.stage(parser_version="new-parser")
        # Reprocessing with unchanged output alone is not a visible file update.
        self.write("A.xlsx", "changed under new parser")
        run = self.stage(parser_version="new-parser")
        self.save(run, ["A.xlsx"])
        approve_stage(self.state, run, require_file_review=True)
        self.calls.clear()
        self.stage(parser_version="new-parser")
        self.assertIn("A.xlsx", self.calls)

    def test_mixed_approval_duplicate_keys_blocks_before_pointer_change(self):
        def process(path, source):
            return [{"Event_Key": path.read_text(), "Source File": source}]
        self.write("A.xlsx", "key-A")
        self.write("B.xlsx", "key-B")
        baseline = stage_update(data_dir=self.data, state_dir=self.state, processor=process)
        approve_stage(self.state, baseline["run_id"])
        pointer = (self.state / "approved/current.json").read_bytes()
        self.write("A.xlsx", "key-C")
        self.write("B.xlsx", "key-A")
        run = stage_update(data_dir=self.data, state_dir=self.state, processor=process)["run_id"]
        self.save(run, ["A.xlsx"])
        with self.assertRaisesRegex(ValueError, "duplicate record keys"):
            approve_stage(self.state, run, require_file_review=True)
        self.assertEqual((self.state / "approved/current.json").read_bytes(), pointer)

    def test_missing_file_choice_rejected_without_creating_saved_review(self):
        self.write("A.xlsx", "A")
        run = self.stage()
        choices = self.root / "choices.csv"
        choices.write_text("Source File,Accept Update?,Owner Note\n")
        with self.assertRaisesRegex(ValueError, "every affected file"):
            save_file_review(self.state, run, choices)
        self.assertFalse((self.state / "staging" / run / "file_review.json").exists())

    def test_completed_approval_export_keeps_reviewed_file_choices_and_notes(self):
        self.baseline()
        self.write("A.xlsx", "new A")
        run = self.stage()
        saved = self.save(run, ["A.xlsx"])
        approve_stage(self.state, run, require_file_review=True)
        export_file_review(self.state, self.root, run)
        with (self.root / "file_review.csv").open(encoding="utf-8-sig", newline="") as handle:
            rows = list(csv.DictReader(handle))
        self.assertEqual([r["Source File"] for r in rows], ["A.xlsx"])
        self.assertEqual(rows[0]["Accept Update?"].upper(), "FALSE")
        self.assertEqual(rows[0]["Owner Note"], saved["choices"]["A.xlsx"]["note"])

    def test_partial_config_upgrade_reprocesses_deferred_file_next_time(self):
        config = self.root / "rules.csv"
        config.write_text("first rules")
        self.write("A.xlsx", "A")
        self.write("B.xlsx", "B")
        run = self.stage(config_paths=[config])
        approve_stage(self.state, run)
        config.write_text("changed rules")
        self.write("A.xlsx", "changed A")
        run = self.stage(config_paths=[config])
        self.save(run, ["A.xlsx"])
        approve_stage(self.state, run, require_file_review=True)
        self.calls.clear()
        self.stage(config_paths=[config])
        self.assertEqual(self.calls, ["A.xlsx"])


if __name__ == "__main__":
    unittest.main()
