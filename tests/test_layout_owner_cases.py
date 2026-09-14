from __future__ import annotations

import unittest
from pathlib import Path

from nmdc_profiler.extractor import discover_layout, read_sheet_model
from nmdc_profiler.layout_compat import _extended_document_header, install_layout_compatibility


ROOT = Path(__file__).resolve().parents[1]


class OwnerLayoutCasesTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        install_layout_compatibility()

    def test_2035_setup_plan_layout_is_recognized_by_runtime_compatibility(self):
        source = ROOT / "DATA" / "METHODS" / "2035 - UMM Shaif Delivarables.xlsx"
        self.assertTrue(source.exists(), source)
        model = read_sheet_model(source, ROOT, "Setup Plans & Anchor Patterns")
        layout, warnings = discover_layout(model)
        self.assertIsNotNone(
            layout,
            f"2035 Setup Plans should recognize NMDC ENERGY NUMBER as the document identifier; warnings={warnings}",
        )
        self.assertEqual(layout.document_col, 3)
        self.assertEqual(layout.data_start, 9)

    def test_nmdc_energy_number_is_a_supported_engineering_header(self):
        self.assertTrue(_extended_document_header("NMDC\nENERGY NUMBER"))

    def test_layout_review_flag_explains_missing_header(self):
        from nmdc_profiler import runtime_engine

        flag = runtime_engine._flag_from_review(
            {
                "Reason": "Unable to establish safe worksheet layout",
                "Warnings": "LAYOUT_DOCUMENT_HEADER_NOT_FOUND",
                "Source File": "METHODS/example.xlsx",
                "Worksheet": "Setup Plans & Anchor Patterns",
                "Project No.": "1000",
            }
        )
        self.assertEqual(flag["code"], "UNRECOGNIZED_LAYOUT_HEADER")
        self.assertIn("document-number/identifier header", flag["message"])
        self.assertIn("NEEDS PARSER/MAPPING FIX", flag["recommended_action"])

    def test_layout_review_flag_explains_missing_data_row(self):
        from nmdc_profiler import runtime_engine

        flag = runtime_engine._flag_from_review(
            {
                "Reason": "Unable to establish safe worksheet layout",
                "Warnings": "LAYOUT_FIRST_DATA_ROW_NOT_FOUND",
                "Source File": "METHODS/example.xlsx",
                "Worksheet": "Installation Procedures",
                "Project No.": "1000",
            }
        )
        self.assertEqual(flag["code"], "UNRECOGNIZED_LAYOUT_DATA")
        self.assertIn("no safe first data row", flag["message"])
        self.assertIn("example document number", flag["recommended_action"])


if __name__ == "__main__":
    unittest.main()
