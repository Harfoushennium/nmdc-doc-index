"""Real-data regressions for the two owner-reported worksheet layouts.

These tests intentionally use the same OOXML reader/extractor path as the runtime;
they do not mutate the source workbooks under DATA/.
"""
from pathlib import Path
import unittest

from nmdc_profiler.extractor import SentinelCase, extract_model, read_sheet_model
from nmdc_profiler.rules import apply_classification, load_rules


ROOT = Path(__file__).resolve().parents[1]
RULES = load_rules(ROOT / "config" / "classification_rules.csv")


class LayoutRegressionTests(unittest.TestCase):
  def test_methods_2171_2172_is_extracted_and_classified_without_review_flag(self):
    rel = "DATA/METHODS/2171-2172 -Document Deliverables LATEST.xlsx"
    model = read_sheet_model(ROOT / rel, ROOT, "2171-2172")
    records, reconciliation = extract_model(
        SentinelCase("REG-METHODS-2171", rel, "2171-2172", "INCLUDE"), model, RULES
    )
    self.assertEqual(reconciliation["status"], "INCLUDE")
    self.assertGreater(reconciliation["event_records"], 0)
    self.assertTrue(all("REVIEW" not in str(w).upper() and "UNRECOGNIZED" not in str(w).upper() for w in reconciliation["warnings"]))
    self.assertTrue(records)
    self.assertEqual({row["Discipline"] for row in records}, {"OFFSHORE INSTALLATION"})
    self.assertEqual({row["Category"] for row in records}, {"PROCEDURE"})
    self.assertEqual({row["Subcategory"] for row in records}, {"INSTALLATION PROCEDURE"})
    self.assertTrue(all(row["Parsing Status"] == "INCLUDE" for row in records))
    self.assertTrue(all("REVIEW" not in str(row["Warnings"]).upper() for row in records))


  def test_tech_3291_client_is_scoped_exclusion_not_unrecognized_layout(self):
    rel = "DATA/TECH/3291 DOCUMENT REGISTER Latest.xlsx"
    classification = apply_classification(
        RULES, "TECH", {"FILE": rel, "WORKSHEET": "CLIENT"}
    )
    self.assertEqual(classification["status"], "EXCLUDED")
    self.assertIn("X017", classification["rule_ids"])
    model = read_sheet_model(ROOT / rel, ROOT, "CLIENT")
    records, reconciliation = extract_model(
        SentinelCase("REG-TECH-3291", rel, "CLIENT", "EXCLUDE"), model, RULES
    )
    self.assertEqual(records, [])
    self.assertEqual(reconciliation["status"], "EXCLUDED")
    self.assertEqual(reconciliation["warnings"], [])
    self.assertNotIn("UNRECOGNIZED", str(reconciliation).upper())
