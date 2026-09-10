import unittest
from pathlib import Path

import nmdc_profiler as profiler


class ClassificationV2Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.rules_path = Path(__file__).resolve().parents[1] / "config" / "classification_rules.csv"
        cls.rules = profiler.load_rules(cls.rules_path)

    def classify(self, family, **evidence):
        return profiler.apply_classification(self.rules, family, evidence)

    def assert_classification(self, result, discipline, category, subcategory):
        self.assertEqual("INCLUDE", result["status"])
        self.assertEqual(discipline, result["discipline"])
        self.assertEqual(category, result["category"])
        self.assertEqual(subcategory, result["subcategory"])

    def test_01_methods_installation_procedures_trailing_space(self):
        result = self.classify("METHODS", WORKSHEET="Installation Procedures ")
        self.assert_classification(
            result, "OFFSHORE INSTALLATION", "PROCEDURE", "INSTALLATION PROCEDURE"
        )
        self.assertIn("M001", result["rule_ids"])

    def test_02_methods_installation_content_fallback(self):
        result = self.classify(
            "METHODS",
            WORKSHEET="2171-2172",
            TITLE="Document No. | Construction and Installation Procedure | Issue Date",
        )
        self.assert_classification(
            result, "OFFSHORE INSTALLATION", "PROCEDURE", "INSTALLATION PROCEDURE"
        )
        self.assertIn("M002", result["rule_ids"])

    def test_03_cut_lists_trailing_space(self):
        result = self.classify("TECH", WORKSHEET="Cut-lists ")
        self.assert_classification(
            result, "GENERAL / MULTIDISCIPLINE", "DRAWING", "CUT LIST"
        )
        self.assertIn("T030", result["rule_ids"])

    def test_04_sketch_trailing_space(self):
        result = self.classify("TECH", WORKSHEET="Sketch ")
        self.assert_classification(
            result, "GENERAL / MULTIDISCIPLINE", "SKETCH", "ENGINEERING SKETCH"
        )
        self.assertIn("T031", result["rule_ids"])

    def test_05_bare_pipeline_alias(self):
        result = self.classify("TECH", WORKSHEET="Pipeline")
        self.assert_classification(
            result, "PIPELINE & CABLE", "DOCUMENT", "GENERAL TECHNICAL DOCUMENT"
        )
        self.assertIn("T001", result["rule_ids"])

    def test_06_bare_naval_marine_alias(self):
        result = self.classify("TECH", WORKSHEET="Naval Marine")
        self.assert_classification(
            result, "NAVAL & MARINE", "DOCUMENT", "GENERAL TECHNICAL DOCUMENT"
        )
        self.assertIn("T010", result["rule_ids"])

    def test_07_uppercase_naval_marine_alias(self):
        result = self.classify("TECH", WORKSHEET="NAVAL MARINE")
        self.assert_classification(
            result, "NAVAL & MARINE", "DOCUMENT", "GENERAL TECHNICAL DOCUMENT"
        )

    def test_08_generic_documents_safe_fallback(self):
        result = self.classify("TECH", WORKSHEET="Documents ")
        self.assert_classification(
            result,
            "GENERAL / MULTIDISCIPLINE",
            "DOCUMENT",
            "GENERAL TECHNICAL DOCUMENT",
        )
        self.assertIn("T002", result["rule_ids"])

    def test_09_generic_documents_title_refinement_preserves_discipline(self):
        result = self.classify(
            "TECH", WORKSHEET="Documents ", TITLE="Mooring Analysis Report for Offshore Barge"
        )
        self.assert_classification(
            result, "GENERAL / MULTIDISCIPLINE", "DOCUMENT", "ANALYSIS REPORT"
        )
        self.assertIn("T002", result["rule_ids"])
        self.assertIn("T070", result["rule_ids"])

    def test_10_3291_client_is_scoped_exclusion(self):
        result = self.classify(
            "TECH",
            FILE="DATA/TECH/3291 DOCUMENT REGISTER Latest.xlsx",
            WORKSHEET="CLIENT",
        )
        self.assertEqual("EXCLUDED", result["status"])
        self.assertEqual("EXCLUDED", result["discipline"])
        self.assertIn("X017", result["rule_ids"])
        self.assertIn("alternate client duplicate view", result["notes"].lower())

    def test_11_unrelated_client_is_not_globally_excluded(self):
        result = self.classify(
            "TECH", FILE="DATA/TECH/4000 DOCUMENT REGISTER.xlsx", WORKSHEET="CLIENT"
        )
        self.assertEqual("UNCLASSIFIED", result["status"])
        self.assertNotIn("X017", result["rule_ids"])

    def test_12_unknown_worksheet_remains_review_required(self):
        result = self.classify("TECH", WORKSHEET="Unknown Future Register")
        self.assertEqual("UNCLASSIFIED", result["status"])
        self.assertEqual("REVIEW_REQUIRED", result["discipline"])
        self.assertEqual([], result["rule_ids"])

    def test_13_original_worksheet_text_is_preserved(self):
        workbooks = [
            {
                "source_family": "TECH",
                "relative_path": "DATA/TECH/4000 DOCUMENT REGISTER.xlsx",
                "readability_status": "READABLE",
                "inferred_project_numbers": ["4000"],
                "project_mismatch_findings": [],
                "sheets": [
                    {
                        "sheet_name": "Cut-lists ",
                        "sections": [],
                        "representative_header_values": ["CUT-LIST DRAWINGS"],
                        "sample_document_numbers": ["4000-NN-0001-CTL"],
                        "sample_titles": ["CUT-LIST DRAWINGS"],
                    }
                ],
            }
        ]
        rows = profiler.classification_rows(workbooks, self.rules)
        self.assertEqual(1, len(rows))
        self.assertEqual("Cut-lists ", rows[0]["worksheet_name"])
        self.assertEqual("cut lists", rows[0]["normalized_worksheet"])
        self.assertEqual("CUT LIST", rows[0]["subcategory"])

    def test_14_v1_sentinel_mappings_still_pass(self):
        cases = [
            (
                "METHODS",
                {"WORKSHEET": "Setup Plans & Anchor Patterns", "TITLE": "Anchor Pattern"},
                ("MARINE OPERATIONS", "DRAWING", "ANCHOR PATTERN"),
            ),
            (
                "TECH",
                {"WORKSHEET": "TN-NA"},
                ("NAVAL & MARINE", "DOCUMENT", "TECHNICAL NOTE"),
            ),
            (
                "TECH",
                {"WORKSHEET": "Drawings"},
                ("GENERAL / MULTIDISCIPLINE", "DRAWING", "ENGINEERING DRAWING"),
            ),
            (
                "TECH",
                {"WORKSHEET": "COMMISSION  - List of OTP"},
                ("COMMISSIONING", "PROCEDURE", "OPERATIONAL TEST PROCEDURE"),
            ),
            (
                "TECH",
                {"WORKSHEET": "Specification", "TITLE": "Pipeline Specification"},
                ("PIPELINE & CABLE", "DOCUMENT", "SPECIFICATION"),
            ),
        ]
        for family, evidence, expected in cases:
            with self.subTest(family=family, evidence=evidence):
                result = profiler.apply_classification(self.rules, family, evidence)
                self.assert_classification(result, *expected)

    def test_15_path_qualifier_column_loaded(self):
        x017 = next(r for r in self.rules if r.rule_id == "X017")
        self.assertTrue(x017.path_qualifier)
        self.assertEqual("WORKSHEET", x017.scope)

    def test_16_path_qualifier_is_case_insensitive(self):
        result = self.classify(
            "TECH",
            FILE="data/tech/3291 document register latest.xlsx",
            WORKSHEET="client",
        )
        self.assertEqual("EXCLUDED", result["status"])
        self.assertIn("X017", result["rule_ids"])

    def test_17_file_exclusions_still_work(self):
        result = profiler.file_exclusion(
            self.rules, "METHODS", "DATA/METHODS/1. Delivarables FORMAT.xlsx"
        )
        self.assertIsNotNone(result)
        self.assertEqual("EXCLUDED", result["status"])
        self.assertIn("X002", result["rule_ids"])

    def test_18_generic_documents_can_refine_to_procedure_without_inventing_discipline(self):
        result = self.classify(
            "TECH",
            WORKSHEET="Documents",
            TITLE="Offshore Construction Procedure",
        )
        self.assert_classification(
            result, "GENERAL / MULTIDISCIPLINE", "DOCUMENT", "PROCEDURE"
        )
        self.assertIn("T074", result["rule_ids"])

    def test_19_methods_sketch_not_refined_to_anchor_pattern(self):
        result = self.classify(
            "METHODS",
            WORKSHEET="Sketches",
            TITLE="OFFSHORE ANCHOR PATTERN AND DP DRAWINGS | OFFSHORE Sketches",
        )
        self.assert_classification(
            result, "OFFSHORE INSTALLATION", "SKETCH", "ENGINEERING SKETCH"
        )
        self.assertIn("M020", result["rule_ids"])
        self.assertNotIn("M011", result["rule_ids"])
        self.assertNotIn("M012", result["rule_ids"])
        self.assertNotIn("M013", result["rule_ids"])

    def test_20_incoming_document_not_refined_by_methods_drawing_title(self):
        result = self.classify(
            "METHODS",
            WORKSHEET="Incomming DOC and DRG",
            SECTION="DOCUMENTS",
            TITLE="OFFSHORE ANCHOR PATTERN AND DP DRAWINGS",
        )
        self.assert_classification(
            result, "EXTERNAL / INPUT", "DOCUMENT", "INCOMING TECHNICAL DOCUMENT"
        )
        self.assertIn("M031", result["rule_ids"])
        self.assertNotIn("M011", result["rule_ids"])

    def test_21_installation_title_fallback_does_not_overwrite_sketch(self):
        result = self.classify(
            "METHODS",
            WORKSHEET="Sketches",
            TITLE="Construction and Installation Procedure reference",
        )
        self.assert_classification(
            result, "OFFSHORE INSTALLATION", "SKETCH", "ENGINEERING SKETCH"
        )
        self.assertIn("M020", result["rule_ids"])
        self.assertNotIn("M002", result["rule_ids"])

    def test_22_tech_drawing_title_does_not_become_document(self):
        result = self.classify(
            "TECH",
            WORKSHEET="Drawings",
            TITLE="Mooring Analysis Report drawing",
        )
        self.assert_classification(
            result, "GENERAL / MULTIDISCIPLINE", "DRAWING", "ENGINEERING DRAWING"
        )
        self.assertIn("T020", result["rule_ids"])
        self.assertNotIn("T070", result["rule_ids"])
        self.assertNotIn("T073", result["rule_ids"])

    def test_23_tech_sketch_title_does_not_become_procedure(self):
        result = self.classify(
            "TECH",
            WORKSHEET="Sketch",
            TITLE="Installation Procedure reference",
        )
        self.assert_classification(
            result, "GENERAL / MULTIDISCIPLINE", "SKETCH", "ENGINEERING SKETCH"
        )
        self.assertIn("T031", result["rule_ids"])
        self.assertNotIn("T074", result["rule_ids"])

    def test_24_technical_note_doc_number_does_not_override_drawing(self):
        result = self.classify(
            "TECH",
            WORKSHEET="Drawings",
            DOC_NUMBER="3000-TN-PL-001",
        )
        self.assert_classification(
            result, "GENERAL / MULTIDISCIPLINE", "DRAWING", "ENGINEERING DRAWING"
        )
        self.assertNotIn("T043", result["rule_ids"])

    def test_25_context_guard_columns_loaded(self):
        m011 = next(r for r in self.rules if r.rule_id == "M011")
        t070 = next(r for r in self.rules if r.rule_id == "T070")
        self.assertEqual("MARINE OPERATIONS", m011.requires_discipline)
        self.assertEqual("DRAWING", m011.requires_category)
        self.assertEqual("DOCUMENT", t070.requires_category)


if __name__ == "__main__":
    unittest.main()
