from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from nmdc_profiler.full_extractor import FullWorkItem
from nmdc_profiler.project_identity import apply_project_identity_overrides, load_project_identity_overrides


class ProjectIdentityOverrideTests(unittest.TestCase):
    def test_01_loads_only_approved_overrides(self):
        text = (
            "Source File,Approved Project No.,Decision Status,Owner Decision,Owner Note,Approved Date,Recheck On Content Change\n"
            "DATA/TECH/2705.xlsx,2705,APPROVED,USE_FILE_PROJECT,ok,2026-09-11,YES\n"
            "DATA/TECH/9999.xlsx,9999,PENDING,USE_FILE_PROJECT,pending,2026-09-11,YES\n"
        )
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "project_identity_overrides.csv"
            path.write_text(text, encoding="utf-8")
            overrides = load_project_identity_overrides(path)
        self.assertEqual({"DATA/TECH/2705.xlsx"}, set(overrides))
        self.assertEqual("2705", overrides["DATA/TECH/2705.xlsx"]["Approved Project No."])

    def test_02_owner_override_replaces_automatic_project_and_is_audited(self):
        item = FullWorkItem("FULL-X", "DATA/TECH/2705.xlsx", "Documents", "2824", "TECH")
        overrides = {
            "DATA/TECH/2705.xlsx": {
                "Approved Project No.": "2705",
                "Owner Decision": "USE_FILE_PROJECT",
                "Owner Note": "internal project number is wrong",
                "Approved Date": "2026-09-11",
            }
        }
        adjusted, audit = apply_project_identity_overrides([item], overrides)
        self.assertEqual("2705", adjusted[0].project_number)
        self.assertEqual(1, len(audit))
        self.assertEqual("2824", audit[0]["Previous Project No."])
        self.assertEqual("2705", audit[0]["Approved Project No."])
        self.assertEqual("OWNER_PROJECT_OVERRIDE_APPLIED", audit[0]["Flag"])

    def test_03_unlisted_source_is_unchanged(self):
        item = FullWorkItem("FULL-X", "DATA/TECH/5000.xlsx", "Documents", "5000", "TECH")
        adjusted, audit = apply_project_identity_overrides([item], {})
        self.assertEqual(item, adjusted[0])
        self.assertEqual([], audit)


if __name__ == "__main__":
    unittest.main()
