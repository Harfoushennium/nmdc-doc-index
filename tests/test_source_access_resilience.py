from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from nmdc_profiler.source_access import _is_transient_access_error, _retry_access


ROOT = Path(__file__).resolve().parents[1]


class SourceAccessResilienceTests(unittest.TestCase):
    def test_retry_access_recovers_from_transient_permission_error(self):
        attempts = {"count": 0}

        def operation():
            attempts["count"] += 1
            if attempts["count"] < 3:
                raise PermissionError(13, "Permission denied")
            return "ok"

        self.assertEqual(_retry_access(operation, attempts=3), "ok")
        self.assertEqual(attempts["count"], 3)

    def test_permission_denied_is_classified_as_transient_access_error(self):
        self.assertTrue(_is_transient_access_error(PermissionError(13, "Permission denied")))

    def test_packaged_engine_installs_resilient_access_before_stage(self):
        text = (ROOT / "nmdc_index_engine.py").read_text(encoding="utf-8")
        self.assertIn("from nmdc_profiler.source_access import install_resilient_source_access", text)
        stage_block = text.split('if args.command == "stage":', 1)[1].split('elif args.command == "export-excel":', 1)[0]
        self.assertIn("install_resilient_source_access()", stage_block)
        self.assertIn("install_layout_compatibility()", stage_block)
        self.assertLess(stage_block.index("install_resilient_source_access()"), stage_block.index("install_layout_compatibility()"))

    def test_source_access_module_preserves_approved_index_when_lock_persists(self):
        text = (ROOT / "nmdc_profiler" / "source_access.py").read_text(encoding="utf-8")
        self.assertIn('"SOURCE_HASH_ERROR"', text)
        self.assertIn("this file was not treated as removed", text)
        self.assertIn("Close this source workbook if it is open", text)
        self.assertIn("OneDrive sync", text)

    def test_real_2369_source_remains_readable_in_repository(self):
        source = ROOT / "DATA" / "METHODS" / "1 Completed Project  Deliverables" / "2369 - NMGL Delivarables.xlsx"
        self.assertTrue(source.exists(), source)
        self.assertGreater(source.stat().st_size, 0)
        with source.open("rb") as handle:
            self.assertEqual(handle.read(2), b"PK")


if __name__ == "__main__":
    unittest.main()
