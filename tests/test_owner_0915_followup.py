from __future__ import annotations

import unittest
from pathlib import Path

from nmdc_profiler import extractor, full_extractor
from nmdc_profiler.extractor import SentinelCase
from nmdc_profiler.layout_compat import _safe_empty_register, install_layout_compatibility
from nmdc_profiler.rules import load_rules


ROOT = Path(__file__).resolve().parents[1]


class Owner0915FollowupRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        install_layout_compatibility()
        cls.rules = load_rules(ROOT / "config" / "classification_rules.csv")

    def test_known_empty_register_sheets_do_not_need_owner_review(self):
        cases = [
            ("DATA/METHODS/E-2964-  ADNOC Delivarables.xlsx", "Installation Procedures"),
            ("DATA/METHODS/E-2964-  ADNOC Delivarables.xlsx", "Setup Plans & Anchor Patterns"),
            ("DATA/METHODS/E-3012 - ADNOC Delivarables.xlsx", "Installation Procedures"),
            ("DATA/METHODS/E-3012 - ADNOC Delivarables.xlsx", "Setup Plans & Anchor Patterns"),
            ("DATA/METHODS/E-3088- ADNOC Delivarables.xlsx", "Installation Procedures"),
            ("DATA/METHODS/E-3088- ADNOC Delivarables.xlsx", "Setup Plans & Anchor Patterns"),
            ("DATA/TECH/2631-DOCUMENT REGISTER.xlsx", "cut-list"),
            ("DATA/TECH/2631-DOCUMENT REGISTER.xlsx", "DRAWINGS"),
            ("DATA/TECH/2705 -DOCUMENT REGISTER.xlsx", "Cut-lists"),
            ("DATA/TECH/2705 -DOCUMENT REGISTER.xlsx", "Documents - Naval Marine"),
            ("DATA/TECH/2705 -DOCUMENT REGISTER.xlsx", "Sketch"),
        ]
        for source_file, worksheet in cases:
            review = {
                "Scope": "WORKSHEET",
                "Source File": source_file,
                "Worksheet": worksheet,
                "Warnings": "LAYOUT_FIRST_DATA_ROW_NOT_FOUND",
            }
            with self.subTest(source=source_file, worksheet=worksheet):
                self.assertTrue(_safe_empty_register(ROOT, review, extractor, full_extractor))

    def test_populated_owner_review_sheets_are_extracted_with_known_header_aliases(self):
        cases = [
            ("DATA/TECH/2745-PP-GE-001-MDR Rev_2.xlsx", "Anchor Pattern"),
            ("DATA/TECH/2745-PP-GE-001-MDR Rev_2.xlsx", "DP"),
            ("DATA/TECH/7279-OS Document register.xlsx", "NAVAL MARINE"),
        ]
        for index, (source_file, worksheet) in enumerate(cases, start=1):
            source = ROOT / source_file
            actual_sheet = full_extractor.resolve_sheet_name(source, worksheet)
            model = extractor.read_sheet_model(source, ROOT, actual_sheet)
            records, reconciliation = extractor.extract_model(
                SentinelCase(f"OWNER-0915-{index}", source_file, actual_sheet, "INCLUDE"),
                model,
                self.rules,
            )
            with self.subTest(source=source_file, worksheet=worksheet):
                self.assertEqual(reconciliation["status"], "INCLUDE")
                self.assertGreater(reconciliation["event_records"], 0)
                self.assertTrue(records)


if __name__ == "__main__":
    unittest.main()
