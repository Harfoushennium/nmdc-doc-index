from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from nmdc_profiler.layout_compat import install_layout_compatibility
from nmdc_profiler.runtime_engine import build_runtime_catalog, stage_runtime_update

ROOT = Path(__file__).resolve().parents[1]

EXPECTED_POPULATED = {
    ("METHODS/2035 - UMM Shaif Delivarables.xlsx", "setup plans & anchor patterns"),
    ("TECH/2745-PP-GE-001-MDR Rev_2.xlsx", "anchor pattern"),
    ("TECH/2745-PP-GE-001-MDR Rev_2.xlsx", "dp"),
    ("TECH/7279-OS Document register.xlsx", "naval marine"),
}
BARGE_SOURCE = "METHODS/2. Barges Sketch Deliverables.xlsx"


def canon(value: str) -> str:
    return " ".join((value or "").split()).casefold()


class ReleaseZeroFindingsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        install_layout_compatibility()

    def test_current_real_data_runtime_has_no_extraction_review_flags(self):
        with tempfile.TemporaryDirectory() as tmp:
            catalog = build_runtime_catalog(
                ROOT / "DATA",
                ROOT / "config" / "classification_rules.csv",
                ROOT / "config" / "project_identity_overrides.csv",
                Path(tmp) / "profile",
            )
            processor_flags = []
            total_records = 0
            populated_seen = set()
            barge_documents = set()
            for source_file in sorted(catalog.work_items, key=str.casefold):
                rows = list(catalog.process(ROOT / "DATA" / source_file, source_file))
                total_records += len(rows)
                for row in rows:
                    key = (source_file, canon(str(row.get("Source Sheet", ""))))
                    if key in EXPECTED_POPULATED:
                        populated_seen.add(key)
                    if source_file == BARGE_SOURCE:
                        barge_documents.add(str(row.get("Document No.", "")))
                processor_flags.extend(catalog.drain_processor_flags())

            actionable_initial = [
                flag for flag in catalog.initial_flags
                if str(flag.get("level", "")).upper() in {"REVIEW", "CONFLICT", "ERROR"}
            ]
            self.assertEqual(
                actionable_initial,
                [],
                "Current known DATA must not require source-selection/workbook review: " + repr(actionable_initial),
            )
            self.assertEqual(
                processor_flags,
                [],
                "Current known DATA must not produce parser/layout Review Flags: " + repr(processor_flags),
            )
            self.assertEqual(populated_seen, EXPECTED_POPULATED)
            self.assertIn("BRG-D4200-001", barge_documents)
            self.assertEqual(catalog.selection_statuses.get(BARGE_SOURCE), "SELECTED")
            self.assertGreater(total_records, 39000)

    def test_clean_first_full_stage_has_zero_review_and_conflict_flags(self):
        with tempfile.TemporaryDirectory() as tmp:
            state_dir = Path(tmp) / "state"
            summary = stage_runtime_update(
                data_dir=ROOT / "DATA",
                state_dir=state_dir,
                config_dir=ROOT / "config",
                full_rescan=True,
            )
            run_id = str(summary["run_id"])
            stage_dir = state_dir / "staging" / run_id
            flags = json.loads((stage_dir / "flags.json").read_text(encoding="utf-8"))
            self.assertEqual(flags, [], "Clean current DATA must stage with an empty Review Flags dataset: " + repr(flags))
            self.assertEqual(int(summary.get("review_flags", -1)), 0)
            self.assertEqual(int(summary.get("blocking_flags", -1)), 0)
            self.assertEqual(summary.get("status"), "STAGED")
            self.assertGreater(int(summary.get("staged_records", 0)), 39000)
            records = [json.loads(line) for line in (stage_dir / "records.jsonl").read_text(encoding="utf-8").splitlines() if line]
            self.assertTrue(
                any(
                    str(row.get("Source File", "")) == BARGE_SOURCE
                    and str(row.get("Document No.", "")) == "BRG-D4200-001"
                    for row in records
                ),
                "Standalone barge sketch register must be extracted in the clean staged dataset",
            )


if __name__ == "__main__":
    unittest.main()
