import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from nmdc_profiler.extractor import (
    SentinelCase,
    SheetModel,
    _formula_hyperlink_target,
    _normalize_date,
    discover_layout,
    extract_model,
    run_sentinels,
)
from nmdc_profiler.rules import load_rules


def add_merge(model, start_row, start_col, end_row, end_col):
    anchor = (start_row, start_col)
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            model.merge_anchor[(row, col)] = anchor


def methods_model():
    cells = {
        (1, 1): "DESCRIPTION", (1, 2): "NMDC DOCUMENT NUMBER", (1, 3): "COMPANY DOCUMENT NUMBER", (1, 4): "REV.",
        (1, 5): "NMDC", (2, 5): "SUBMISSION STATUS", (3, 5): "Methods TO PMT", (4, 5): "REF", (4, 6): "DATE",
        (1, 7): "CLIENT", (2, 7): "APPROVAL STATUS", (3, 7): "FROM CLIENT", (4, 7): "DATE", (4, 8): "CODE",
        (5, 1): "MOORING ANALYSIS REPORT", (5, 2): "9999-PP-OF-001", (5, 3): "CLIENT-9999-001", (5, 4): "A1", (5, 5): "T-1", (5, 6): "45287",
        (6, 7): "45290", (6, 8): "1 - Approved",
        (7, 4): "1", (7, 5): "T-2", (7, 6): "11-Feb-2025",
    }
    model = SheetModel("DATA/METHODS/9999 Deliverables.xlsx", "METHODS", "2026-01-01T00:00:00Z", "Installation Procedures", 8, 8, cells=cells)
    for col in (1, 2, 3, 4):
        add_merge(model, 1, col, 4, col)
    for start in ((1, 5, 1, 6), (2, 5, 2, 6), (3, 5, 3, 6), (1, 7, 1, 8), (2, 7, 2, 8), (3, 7, 3, 8)):
        add_merge(model, *start)
    for col in (1, 2, 3):
        add_merge(model, 5, col, 7, col)
    add_merge(model, 5, 4, 6, 4)
    model.hyperlinks[(5, 2)] = "..\\docs\\9999-PP-OF-001.pdf"
    return model


def tech_model():
    cells = {
        (1, 1): "TITLE/DESCRIPTION", (1, 2): "NMDC DOCUMENT NUMBER", (1, 3): "REVISION",
        (1, 4): "TRANSMITTAL", (2, 4): "REF", (2, 5): "DATE",
        (3, 1): "INSTALLATION ENGINEERING ANALYSIS REPORT", (3, 2): "2820-NN-RP-001", (3, 3): "A", (3, 4): "T-1", (3, 5): "45287",
        (4, 1): "OFFSHORE INSTALLATION PROCEDURE", (4, 2): "2820-NN-RP-002", (4, 3): "0", (4, 4): "T-2", (4, 5): "45288",
    }
    model = SheetModel("DATA/TECH/2820-DOCUMENT REGISTER-NEW 30-04-2026.xlsx", "TECH", "2026-09-08T09:55:51Z", "Documents - Pipeline & Cable", 4, 5, cells=cells)
    add_merge(model, 1, 1, 2, 1); add_merge(model, 1, 2, 2, 2); add_merge(model, 1, 3, 2, 3); add_merge(model, 1, 4, 1, 5)
    return model


def reversed_header_model():
    cells = {
        (1, 1): "#", (1, 2): "Document No.", (1, 3): "Document Title", (1, 4): "Revision",
        (1, 5): "Issue Date (Planned)", (1, 6): "Outgoing Ref No.", (1, 7): "Issue Date (Actual)",
        (2, 5): "Construction and Installation Procedure",
        (3, 1): "1", (3, 2): "2171-2172-PP-OF-012", (3, 3): "ANCHOR HANDLING PROCEDURE FOR FLOATING BARGE",
        (3, 4): "A1", (3, 5): "44438", (3, 6): "T-553/21", (3, 7): "44553",
    }
    model = SheetModel(
        "DATA/METHODS/2171-2172 -Document Deliverables LATEST.xlsx",
        "METHODS",
        "2026-09-08T05:23:16Z",
        "2171-2172",
        3,
        7,
        cells=cells,
    )
    add_merge(model, 2, 5, 2, 7)
    return model


class ExtractorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.rules = load_rules(Path(__file__).resolve().parents[1] / "config" / "classification_rules.csv")

    def test_01_exact_merge_inheritance_is_bounded(self):
        model = methods_model()
        self.assertEqual("9999-PP-OF-001", model.value(6, 2))
        self.assertEqual("9999-PP-OF-001", model.value(7, 2))
        self.assertEqual("", model.value(8, 2))

    def test_02_revision_merge_inherits_only_inside_range(self):
        model = methods_model()
        self.assertEqual("A1", model.value(6, 4))
        self.assertEqual("1", model.value(7, 4))
        self.assertEqual("", model.value(8, 4))

    def test_03_layout_discovers_multilevel_headers(self):
        layout, warnings = discover_layout(methods_model())
        self.assertIsNotNone(layout)
        self.assertEqual(5, layout.data_start)
        self.assertEqual(2, layout.document_col)
        self.assertEqual(4, layout.revision_col)
        self.assertGreaterEqual(len(layout.event_groups), 2)
        self.assertNotIn("LAYOUT_DOCUMENT_HEADER_NOT_FOUND", warnings)

    def test_04_one_document_revision_history_expands_to_events(self):
        case = SentinelCase("SYN", "x", "Installation Procedures", "INCLUDE")
        records, recon = extract_model(case, methods_model(), self.rules)
        self.assertEqual(3, recon["rows_with_identity"])
        self.assertEqual(1, recon["distinct_documents"])
        self.assertEqual(2, recon["distinct_revisions"])
        self.assertGreaterEqual(len(records), 3)
        self.assertEqual(1, sum(int(r["Document_Row_Flag"]) for r in records))
        self.assertEqual(2, sum(int(r["Revision_Row_Flag"]) for r in records))

    def test_05_multiple_transaction_groups_survive(self):
        case = SentinelCase("SYN", "x", "Installation Procedures", "INCLUDE")
        records, _ = extract_model(case, methods_model(), self.rules)
        types = {r["Event Type"] for r in records}
        self.assertTrue(any("PMT" in t for t in types))
        self.assertTrue(any("CLIENT" in t.upper() or "FROM CLIENT" in t.upper() for t in types))

    def test_06_serial_date_converts_only_when_date_field_is_used(self):
        normalized, warning = _normalize_date("45287")
        self.assertRegex(normalized, r"^\d{4}-\d{2}-\d{2}$")
        self.assertEqual("", warning)
        self.assertEqual(("ABC", "DATE_TEXT_PRESERVED"), _normalize_date("ABC"))

    def test_07_native_hyperlink_inherits_with_merged_document_cell(self):
        model = methods_model()
        self.assertEqual("..\\docs\\9999-PP-OF-001.pdf", model.hyperlink(7, 2))
        case = SentinelCase("SYN", "x", "Installation Procedures", "INCLUDE")
        records, _ = extract_model(case, model, self.rules)
        self.assertTrue(all(r["Document Link"] == "..\\docs\\9999-PP-OF-001.pdf" for r in records))

    def test_08_formula_hyperlink_target(self):
        self.assertEqual("../docs/a.pdf", _formula_hyperlink_target('HYPERLINK("../docs/a.pdf","open")'))
        self.assertEqual("", _formula_hyperlink_target("SUM(A1:A2)"))

    def test_09_keys_are_deterministic_and_latest_flags_follow_source_order(self):
        case = SentinelCase("SYN", "x", "Installation Procedures", "INCLUDE")
        first, _ = extract_model(case, methods_model(), self.rules)
        second, _ = extract_model(case, methods_model(), self.rules)
        self.assertEqual([r["Event_Key"] for r in first], [r["Event_Key"] for r in second])
        latest = [r for r in first if r["Is_Latest_Revision"]]
        self.assertTrue(latest)
        self.assertTrue(all(r["Revision"] == "1" for r in latest))
        self.assertEqual(2, sum(int(r["Is_Latest_Event"]) for r in first))

    def test_10_per_document_title_refinement_does_not_leak(self):
        case = SentinelCase("TECH", "x", "Documents - Pipeline & Cable", "INCLUDE")
        records, recon = extract_model(case, tech_model(), self.rules)
        self.assertEqual("INCLUDE", recon["status"])
        by_doc = {}
        for r in records:
            by_doc.setdefault(r["Document No."], r["Subcategory"])
        self.assertEqual("ANALYSIS REPORT", by_doc["2820-NN-RP-001"])
        self.assertEqual("PROCEDURE", by_doc["2820-NN-RP-002"])

    def test_11_unknown_layout_is_visible_not_guessed(self):
        model = SheetModel("DATA/TECH/9999.xlsx", "TECH", "", "Mystery", 3, 3, cells={(1, 1): "hello", (2, 1): "world"})
        case = SentinelCase("UNKNOWN", "x", "Mystery", "INCLUDE")
        records, recon = extract_model(case, model, self.rules)
        self.assertEqual([], records)
        self.assertEqual("REVIEW_REQUIRED", recon["status"])
        self.assertIn("LAYOUT_DOCUMENT_HEADER_NOT_FOUND", recon["warnings"])

    def test_12_scoped_3291_client_is_excluded(self):
        model = SheetModel("DATA/TECH/3291 DOCUMENT REGISTER Latest.xlsx", "TECH", "", "CLIENT", 1, 1, cells={(1, 1): "anything"})
        case = SentinelCase("CLIENT", "x", "CLIENT", "EXCLUDE")
        records, recon = extract_model(case, model, self.rules)
        self.assertEqual([], records)
        self.assertEqual("EXCLUDED", recon["status"])

    def test_13_client_is_not_globally_excluded(self):
        model = SheetModel("DATA/TECH/9999 DOCUMENT REGISTER.xlsx", "TECH", "", "CLIENT", 1, 1, cells={(1, 1): "anything"})
        case = SentinelCase("CLIENT", "x", "CLIENT", "INCLUDE")
        records, recon = extract_model(case, model, self.rules)
        self.assertEqual([], records)
        self.assertEqual("REVIEW_REQUIRED", recon["status"])

    def test_14_reversed_multilevel_header_consolidates_one_transaction(self):
        layout, warnings = discover_layout(reversed_header_model())
        self.assertIsNotNone(layout)
        self.assertEqual((("Construction and Installation Procedure", (5, 6, 7)),), layout.event_groups)
        self.assertIn(1, layout.metadata_cols)
        self.assertEqual([], warnings)

    def test_15_reversed_header_uses_actual_date_and_preserves_all_fields(self):
        case = SentinelCase("REVHDR", "x", "2171-2172", "INCLUDE")
        records, recon = extract_model(case, reversed_header_model(), self.rules)
        self.assertEqual("INCLUDE", recon["status"])
        self.assertEqual(1, len(records))
        row = records[0]
        self.assertEqual("Construction and Installation Procedure", row["Event Type"])
        self.assertEqual("2021-12-23", row["Event Date"])
        self.assertEqual("T-553/21", row["Event Reference"])
        self.assertNotEqual("#", row["Event Type"])
        self.assertIn('"Issue Date (Planned)":"2021-08-30"', row["Event Values JSON"])
        self.assertIn('"Issue Date (Actual)":"2021-12-23"', row["Event Values JSON"])

    def test_16_non_date_text_in_date_column_is_preserved_but_not_promoted(self):
        model = reversed_header_model()
        model.cells[(3, 7)] = "PMT ISSUED"
        case = SentinelCase("REVHDR", "x", "2171-2172", "INCLUDE")
        records, _ = extract_model(case, model, self.rules)
        self.assertEqual("2021-08-30", records[0]["Event Date"])
        self.assertIn("DATE_TEXT_PRESERVED", records[0]["Warnings"])
        self.assertIn('"Issue Date (Actual)":"PMT ISSUED"', records[0]["Event Values JSON"])

    def test_17_run_sentinels_sorts_by_canonical_event_key(self):
        case = SentinelCase("REVHDR", "dummy.xlsx", "2171-2172", "INCLUDE")
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            (root / "dummy.xlsx").write_bytes(b"placeholder")
            with patch("nmdc_profiler.extractor.read_sheet_model", return_value=reversed_header_model()):
                records, reconciliation = run_sentinels(root, [case], self.rules)
        self.assertEqual(1, len(records))
        self.assertTrue(records[0]["Event_Key"].startswith("EVT-"))
        self.assertEqual("INCLUDE", reconciliation[0]["status"])


if __name__ == "__main__":
    unittest.main()
