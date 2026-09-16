import csv
import json
import tempfile
import time
import unittest
from pathlib import Path

from nmdc_profiler.rules import apply_classification, load_rules
from nmdc_profiler.source_cache import prepare_local_source_cache
from nmdc_profiler.source_selection import (
    read_source_exclusions,
    set_source_selection,
    set_source_selections_from_file,
)
from nmdc_profiler.source_selection_view import export_source_selection
from nmdc_profiler.ui_layout import (
    DOCUMENT_FIELDS,
    EVENT_FIELDS,
    FLAG_FIELDS,
    PENDING_FIELDS,
    REVISION_FIELDS,
)


ROOT = Path(__file__).resolve().parents[1]


class Owner0916UsabilityPerformanceTests(unittest.TestCase):
    def test_shipped_rules_are_plain_text_only(self):
        with (ROOT / "config" / "classification_rules.csv").open(
            newline="", encoding="utf-8-sig"
        ) as handle:
            rows = list(csv.DictReader(handle))
        self.assertTrue(rows)
        match_types = {row["Match_Type"].strip().upper() for row in rows}
        self.assertNotIn("REGEX", match_types)
        self.assertTrue(match_types <= {"CONTAINS", "EXACT", "STARTS_WITH", "ENDS_WITH"})

    def test_plain_rules_keep_representative_classification_semantics(self):
        rules = load_rules(ROOT / "config" / "classification_rules.csv")
        cases = [
            ("METHODS", {"WORKSHEET": "Installation Procedures"}, "INSTALLATION PROCEDURE"),
            ("METHODS", {"WORKSHEET": "Setup Plans & Anchor Patterns"}, "METHOD DRAWING"),
            ("TECH", {"WORKSHEET": "Naval & Marine"}, "GENERAL TECHNICAL DOCUMENT"),
            ("TECH", {"WORKSHEET": "Cut-lists"}, "CUT LIST"),
            ("TECH", {"WORKSHEET": "DP"}, "DP SETUP PLAN"),
        ]
        for family, evidence, expected_subcategory in cases:
            with self.subTest(family=family, evidence=evidence):
                result = apply_classification(rules, family, evidence)
                self.assertEqual("INCLUDE", result["status"])
                self.assertEqual(expected_subcategory, result["subcategory"])

    def test_local_source_cache_reuses_unchanged_file_without_copying_again(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            data = root / "DATA"
            cache = root / "runtime" / "source_cache"
            source = data / "TECH" / "4000 DOCUMENT REGISTER.xlsx"
            source.parent.mkdir(parents=True)
            source.write_bytes(b"first-version")

            files_root, stats1 = prepare_local_source_cache(data, cache)
            cached = files_root / "TECH" / source.name
            self.assertEqual(b"first-version", cached.read_bytes())
            self.assertEqual(1, stats1["cache_copied"])
            self.assertEqual(0, stats1["cache_reused"])

            cached_mtime = cached.stat().st_mtime_ns
            files_root2, stats2 = prepare_local_source_cache(data, cache)
            self.assertEqual(files_root, files_root2)
            self.assertEqual(0, stats2["cache_copied"])
            self.assertEqual(1, stats2["cache_reused"])
            self.assertEqual(cached_mtime, cached.stat().st_mtime_ns)

            time.sleep(0.01)
            source.write_bytes(b"second-version-longer")
            _files_root3, stats3 = prepare_local_source_cache(data, cache)
            self.assertEqual(1, stats3["cache_copied"])
            self.assertEqual(b"second-version-longer", cached.read_bytes())

    def test_source_selection_can_exclude_and_restore_without_editing_source(self):
        with tempfile.TemporaryDirectory() as temp:
            config = Path(temp)
            source = "TECH/4000 DOCUMENT REGISTER.xlsx"
            excluded = set_source_selection(
                config, source, include=False, reason="Duplicate register"
            )
            self.assertEqual("EXCLUDE", excluded["decision"])
            rows = read_source_exclusions(config / "source_exclusions.csv")
            self.assertIn(source.casefold(), rows)
            self.assertEqual("Duplicate register", rows[source.casefold()]["Reason"])

            restored = set_source_selection(config, source, include=True)
            self.assertEqual("INCLUDE", restored["decision"])
            self.assertEqual({}, read_source_exclusions(config / "source_exclusions.csv"))

    def test_batch_checkbox_source_selection_is_atomic_and_owner_friendly(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            config = root / "config"
            config.mkdir()
            decisions = root / "source_selection_decisions.csv"
            with decisions.open("w", newline="", encoding="utf-8-sig") as handle:
                writer = csv.DictWriter(
                    handle,
                    fieldnames=["Source File", "Include in Index?", "Owner Note"],
                )
                writer.writeheader()
                writer.writerow(
                    {
                        "Source File": "TECH/4000 DOCUMENT REGISTER.xlsx",
                        "Include in Index?": "FALSE",
                        "Owner Note": "Duplicate register",
                    }
                )
                writer.writerow(
                    {
                        "Source File": "METHODS/5000 Deliverables.xlsx",
                        "Include in Index?": "TRUE",
                        "Owner Note": "",
                    }
                )

            result = set_source_selections_from_file(config, decisions)
            self.assertEqual("BATCH_SOURCE_SELECTION_SAVED", result["decision"])
            self.assertEqual(2, result["processed_sources"])
            self.assertEqual(1, result["unchecked_excluded"])
            exclusions = read_source_exclusions(config / "source_exclusions.csv")
            self.assertIn("tech/4000 document register.xlsx", exclusions)
            self.assertNotIn("methods/5000 deliverables.xlsx", exclusions)
            self.assertEqual("Duplicate register", exclusions["tech/4000 document register.xlsx"]["Reason"])

    def test_source_selection_view_is_checkbox_friendly(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            state = root / "runtime"
            stage = state / "staging" / "R1"
            stage.mkdir(parents=True)
            (state / "staging" / "latest.json").write_text(
                json.dumps({"run_id": "R1"}), encoding="utf-8"
            )
            (stage / "manifest.json").write_text(
                json.dumps(
                    {
                        "files": [
                            {
                                "relative_path": "TECH/4000 DOCUMENT REGISTER.xlsx",
                                "selection_status": "SELECTED",
                                "selection_exclusion_reason": "Only confirmed source",
                                "last_processed_run": "R1",
                            }
                        ]
                    }
                ),
                encoding="utf-8",
            )
            (stage / "summary.json").write_text("{}", encoding="utf-8")
            (stage / "flags.json").write_text("[]", encoding="utf-8")
            (stage / "record_changes.json").write_text("{}", encoding="utf-8")
            (stage / "records.jsonl").write_text("", encoding="utf-8")
            config = root / "config"
            config.mkdir()
            set_source_selection(
                config,
                "TECH/4000 DOCUMENT REGISTER.xlsx",
                include=False,
                reason="Owner scope decision",
            )
            exchange = root / "exchange"
            target = export_source_selection(state, exchange, config)
            with target.open(newline="", encoding="utf-8-sig") as handle:
                rows = list(csv.DictReader(handle))
            self.assertEqual(1, len(rows))
            self.assertEqual("FALSE", rows[0]["Include in Index?"])
            self.assertEqual("TECH/4000 DOCUMENT REGISTER.xlsx", rows[0]["Source File"])
            self.assertEqual("Owner scope decision", rows[0]["Owner Note"])
            self.assertEqual("Only confirmed source", rows[0]["Selection Reason"])

    def test_review_first_column_orders_put_business_fields_before_technical_keys(self):
        self.assertEqual(["Project No.", "Document No.", "Document Title"], DOCUMENT_FIELDS[:3])
        self.assertEqual(["Project No.", "Document No.", "Revision"], REVISION_FIELDS[:3])
        self.assertEqual(["Project No.", "Document No.", "Revision", "Event Type"], EVENT_FIELDS[:4])
        self.assertEqual(["Change Type", "Project No.", "Document No."], PENDING_FIELDS[:3])
        self.assertEqual(
            ["Flag Level", "Plain-English Problem", "Recommended User Action", "User Decision"],
            FLAG_FIELDS[:4],
        )
        self.assertEqual("Record Identity", PENDING_FIELDS[-1])
        self.assertEqual("Event Key", FLAG_FIELDS[-1])

    def test_excel_package_uses_local_runtime_responsive_scan_guidance_and_native_checkboxes(self):
        setup = (ROOT / "packaging" / "Create_NMDC_Document_Index.vbs").read_text(
            encoding="utf-8-sig"
        )
        performance = (ROOT / "excel" / "vba" / "modNMDC_Performance.bas").read_text(
            encoding="utf-8-sig"
        )
        owner_ux = (ROOT / "excel" / "vba" / "modNMDC_OwnerUX.bas").read_text(
            encoding="utf-8-sig"
        )
        checkboxes = (ROOT / "excel" / "vba" / "modNMDC_Checkboxes.bas").read_text(
            encoding="utf-8-sig"
        )
        engine = (ROOT / "nmdc_index_engine.py").read_text(encoding="utf-8")

        self.assertIn("%LOCALAPPDATA%", setup)
        self.assertIn("modNMDC_Performance.bas", setup)
        self.assertIn("modNMDC_OwnerUX.bas", setup)
        self.assertIn("modNMDC_Checkboxes.bas", setup)
        self.assertIn('"Include in Index?"', setup)
        self.assertIn("NMDC_UpdateChangedFilesFast", setup)
        self.assertIn("NMDC_FullRescanFast", setup)
        self.assertIn("NMDC_ReviewPendingUpdateFast", setup)
        self.assertIn("NMDC_RefreshDashboardFast", setup)
        self.assertIn("NMDC_FastStartup", setup)

        self.assertIn("Application.OnTime", performance)
        self.assertIn("scanning in background", performance)
        self.assertIn("NMDC_RefreshReviewDataFast", performance)
        self.assertIn("NMDC_AsyncQualifiedMacro", performance)
        fast_refresh = performance.split("Public Function NMDC_RefreshReviewDataFast", 1)[1].split(
            "Public Function NMDC_RefreshDashboardOnlyFast", 1
        )[0]
        self.assertNotIn("master_documents.csv", fast_refresh)
        self.assertNotIn("revisions.csv", fast_refresh)
        self.assertNotIn("events.csv", fast_refresh)

        self.assertIn("QUICK WORKFLOW - WHAT TO DO", owner_ux)
        self.assertIn("native Excel checkboxes", owner_ux)
        self.assertIn("Save Source Choices & Restage", owner_ux)
        self.assertIn("CONTAINS,EXACT,STARTS_WITH,ENDS_WITH", owner_ux)
        self.assertNotIn("CONTAINS,EXACT,FUZZY,REGEX", owner_ux)

        self.assertIn("CellControl.SetCheckbox", checkboxes)
        self.assertNotIn("CheckBoxes.Add", checkboxes)
        self.assertIn("NMDC_SourceCheckboxClicked", checkboxes)
        self.assertIn("NMDC_SaveSourceSelections", checkboxes)
        self.assertIn("NMDC_CheckAllSources", checkboxes)
        self.assertIn("NMDC_UncheckAllSources", checkboxes)
        self.assertIn("save-source-selections", checkboxes)
        self.assertIn("save-source-selections", engine)

    def test_checkbox_policy_keeps_multi_choice_inputs_as_dropdowns(self):
        """Checkboxes are for true binary owner choices, not multi-option decisions."""
        setup = (ROOT / "packaging" / "Create_NMDC_Document_Index.vbs").read_text(
            encoding="utf-8-sig"
        )
        self.assertIn("ACKNOWLEDGED,NO ACTION REQUIRED,NEEDS SOURCE CORRECTION", setup)
        self.assertIn("OPEN,ACKNOWLEDGED,RESOLVED,DEFERRED", setup)

    def test_production_package_includes_source_selection_config(self):
        workflow = (ROOT / ".github" / "workflows" / "production-package.yml").read_text(
            encoding="utf-8"
        )
        self.assertIn("source_exclusions.csv", workflow)
        self.assertIn("source_selection.csv", workflow)


if __name__ == "__main__":
    unittest.main()
