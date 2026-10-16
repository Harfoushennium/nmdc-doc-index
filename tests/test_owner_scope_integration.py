from __future__ import annotations

import csv
import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from nmdc_profiler.excel_bridge import load_approved_state, load_latest_stage
from tests.test_profiler import make_xlsx

ROOT = Path(__file__).resolve().parents[1]


def make_register(path: Path, project: str) -> None:
    """Write one representative merged-header METHODS register."""
    make_xlsx(path, [{
        "name": "Installation Procedures",
        "rows": [
            ["DESCRIPTION", "NMDC DOCUMENT NUMBER", "COMPANY DOCUMENT NUMBER", "REV.", "NMDC", None, "CLIENT", None],
            [None, None, None, None, "SUBMISSION STATUS", None, "APPROVAL STATUS", None],
            [None, None, None, None, "Methods TO PMT", None, "FROM CLIENT", None],
            [None, None, None, None, "REF", "DATE", "DATE", "CODE"],
            ["MOORING ANALYSIS REPORT", f"{project}-PP-OF-001", f"CLIENT-{project}-001", "A1", "T-1", "2025-02-11"],
            [None, None, None, None, None, None, "2025-02-12", "1 - Approved"],
        ],
        "merges": ["A1:A4", "B1:B4", "C1:C4", "D1:D4", "E1:F1", "E2:F2", "E3:F3", "G1:H1", "G2:H2", "G3:H3", "A5:A6", "B5:B6", "C5:C6", "D5:D6"],
    }])


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class OwnerScopeIntegrationTests(unittest.TestCase):
    def test_excluded_source_is_absent_from_real_stage_until_approval(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            data = root / "DATA"
            source_a = data / "METHODS" / "4001 Register.xlsx"
            source_b = data / "METHODS" / "4002 Register.xlsx"
            make_register(source_a, "4001")
            make_register(source_b, "4002")
            source_hashes = {source_a: sha256(source_a), source_b: sha256(source_b)}

            state, exchange, config = root / "state", root / "exchange", root / "config"
            config.mkdir()
            for name in ("classification_rules.csv", "project_identity_overrides.csv"):
                (config / name).write_bytes((ROOT / "config" / name).read_bytes())

            def cli(*args: str) -> None:
                result = subprocess.run(
                    [sys.executable, str(ROOT / "nmdc_index_engine.py"), *args],
                    cwd=ROOT,
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(result.returncode, 0, f"stdout: {result.stdout}\nstderr: {result.stderr}")

            common = ("--state-dir", str(state), "--exchange-dir", str(exchange), "--config-dir", str(config))
            cli("stage", *common, "--mode", "full", "--data-dir", str(data))
            baseline_stage = load_latest_stage(state)
            self.assertTrue(any(row.get("Source File") == "METHODS/4001 Register.xlsx" for row in baseline_stage["records"]))
            self.assertTrue(any(row.get("Source File") == "METHODS/4002 Register.xlsx" for row in baseline_stage["records"]))
            first_run = baseline_stage["run_id"]
            cli("approve", *common, "--run-id", first_run)
            approved_before = load_approved_state(state)
            pointer_before = (state / "approved" / "current.json").read_bytes()
            records_before = (state / "approved" / "versions" / approved_before["run_id"] / "records.jsonl").read_bytes()

            decisions = root / "source-decisions.csv"
            with decisions.open("w", newline="", encoding="utf-8-sig") as handle:
                writer = csv.DictWriter(handle, fieldnames=("Source File", "Include in Index?", "Owner Note"))
                writer.writeheader()
                writer.writerow({"Source File": "METHODS/4001 Register.xlsx", "Include in Index?": "Yes", "Owner Note": "Retain for comparison"})
                writer.writerow({"Source File": "METHODS/4002 Register.xlsx", "Include in Index?": "No", "Owner Note": "Outside approved review scope"})
            cli("save-source-selections", *common, "--selections-file", str(decisions))
            with (config / "source_exclusions.csv").open(encoding="utf-8-sig", newline="") as handle:
                selection_rows = {row["Relative_Path"]: row for row in csv.DictReader(handle)}
            self.assertEqual(selection_rows["METHODS/4001 Register.xlsx"]["Owner_Note"], "Retain for comparison")
            cli("stage", *common, "--mode", "full", "--data-dir", str(data))

            staged = load_latest_stage(state)
            staged_sources = {row.get("Source File") for row in staged["records"]}
            self.assertIn("METHODS/4001 Register.xlsx", staged_sources)
            self.assertNotIn("METHODS/4002 Register.xlsx", staged_sources)
            self.assertEqual((state / "approved" / "current.json").read_bytes(), pointer_before)
            self.assertEqual((state / "approved" / "versions" / approved_before["run_id"] / "records.jsonl").read_bytes(), records_before)
            self.assertEqual(load_approved_state(state)["run_id"], approved_before["run_id"])
            self.assertEqual({path: sha256(path) for path in source_hashes}, source_hashes)


if __name__ == "__main__":
    unittest.main()
