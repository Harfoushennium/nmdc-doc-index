from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from nmdc_profiler.layout_compat import install_layout_compatibility
from nmdc_profiler.runtime_engine import build_runtime_catalog

ROOT = Path(__file__).resolve().parents[1]

EXPECTED_POPULATED = {
    ("METHODS/2035 - UMM Shaif Delivarables.xlsx", "setup plans & anchor patterns"),
    ("TECH/2745-PP-GE-001-MDR Rev_2.xlsx", "anchor pattern"),
    ("TECH/2745-PP-GE-001-MDR Rev_2.xlsx", "dp"),
    ("TECH/7279-OS Document register.xlsx", "naval marine"),
}


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
            for source_file in sorted(catalog.work_items, key=str.casefold):
                rows = list(catalog.process(ROOT / "DATA" / source_file, source_file))
                total_records += len(rows)
                for row in rows:
                    key = (source_file, canon(str(row.get("Source Sheet", ""))))
                    if key in EXPECTED_POPULATED:
                        populated_seen.add(key)
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
            self.assertGreater(total_records, 39000)


if __name__ == "__main__":
    unittest.main()
