from __future__ import annotations

import csv
import tempfile
import unittest
from pathlib import Path

from tools.build_real_data_review import build_review_exchange


ROOT = Path(__file__).resolve().parents[1]


class RealDataReviewPackageTests(unittest.TestCase):
    def test_exchange_is_derived_from_real_cycle3_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            result = build_review_exchange(ROOT, Path(tmp) / "exchange")
            self.assertEqual(result["documents"], 7498)
            self.assertEqual(result["revisions"], 10663)
            self.assertEqual(result["events"], 39344)
            self.assertEqual(result["flags"], 15)

            with (Path(tmp) / "exchange" / "dashboard.csv").open(encoding="utf-8-sig", newline="") as handle:
                dashboard = next(csv.DictReader(handle))
            self.assertEqual(dashboard["Approved Status"], "REAL-DATA BASELINE — REVIEW ONLY")
            self.assertEqual(dashboard["Approved Documents"], "7498")
            self.assertEqual(dashboard["Approved Revisions"], "10663")
            self.assertEqual(dashboard["Approved Transactions"], "39344")
            self.assertEqual(dashboard["Conflict Flags"], "0")

            with (Path(tmp) / "exchange" / "flags.csv").open(encoding="utf-8-sig", newline="") as handle:
                flags = list(csv.DictReader(handle))
            self.assertEqual(len(flags), 15)
            self.assertEqual({row["Flag Level"] for row in flags}, {"REVIEW"})
            self.assertTrue(all(row["Resolution Status"] == "OPEN" for row in flags))

    def test_review_builder_does_not_write_to_data(self) -> None:
        source = ROOT / "DATA" / "TECH" / "2705 -DOCUMENT REGISTER.xlsx"
        before = source.read_bytes()
        with tempfile.TemporaryDirectory() as tmp:
            build_review_exchange(ROOT, Path(tmp) / "exchange")
        self.assertEqual(source.read_bytes(), before)


if __name__ == "__main__":
    unittest.main()
