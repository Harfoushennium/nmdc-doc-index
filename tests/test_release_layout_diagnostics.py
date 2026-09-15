from __future__ import annotations

import json
import unittest
from pathlib import Path

from nmdc_profiler.extractor import discover_layout, read_sheet_model, _looks_identifier
from nmdc_profiler.layout_compat import install_layout_compatibility
from nmdc_profiler.core import norm_text

ROOT = Path(__file__).resolve().parents[1]

CASES = [
    ("DATA/METHODS/2035 - UMM Shaif Delivarables.xlsx", "Setup Plans & Anchor Patterns"),
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
    ("DATA/TECH/2745-PP-GE-001-MDR Rev_2.xlsx", "Anchor Pattern"),
    ("DATA/TECH/2745-PP-GE-001-MDR Rev_2.xlsx", "DP"),
    ("DATA/TECH/7279-OS Document register.xlsx", "NAVAL MARINE"),
]


def row_summary(model, row: int) -> str:
    values = []
    for col in range(1, min(model.max_col, 40) + 1):
        value = str(model.value(row, col) or "").strip()
        if value:
            values.append(f"C{col}={value[:120]}")
    return " | ".join(values[:12])


class ReleaseLayoutDiagnostics(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        install_layout_compatibility()

    def test_print_known_review_layouts(self):
        payload = []
        for source_rel, sheet in CASES:
            source = ROOT / source_rel
            self.assertTrue(source.exists(), source)
            model = read_sheet_model(source, ROOT, sheet)
            layout, warnings = discover_layout(model)
            nonempty_rows = sorted({row for row, _col in model.cells})
            first_rows = [
                {"row": row, "text": row_summary(model, row)}
                for row in nonempty_rows[:20]
            ]
            identifier_samples = []
            for (row, col), value in sorted(model.cells.items()):
                text = str(value or "").strip()
                if _looks_identifier(text):
                    identifier_samples.append({"row": row, "col": col, "value": text[:120]})
                if len(identifier_samples) >= 12:
                    break
            payload.append(
                {
                    "source": source_rel,
                    "sheet": sheet,
                    "max_row": model.max_row,
                    "max_col": model.max_col,
                    "layout": None if layout is None else {
                        "header_start": layout.header_start,
                        "header_end": layout.header_end,
                        "data_start": layout.data_start,
                        "document_col": layout.document_col,
                        "company_document_col": layout.company_document_col,
                        "title_col": layout.title_col,
                        "revision_col": layout.revision_col,
                    },
                    "warnings": warnings,
                    "first_nonempty_rows": first_rows,
                    "identifier_samples": identifier_samples,
                }
            )
        print("NMDC_LAYOUT_DIAGNOSTICS=" + json.dumps(payload, ensure_ascii=False, sort_keys=True))
        self.assertEqual(len(payload), len(CASES))


if __name__ == "__main__":
    unittest.main()
