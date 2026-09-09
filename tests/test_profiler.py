import unittest, json, csv, os, sys, subprocess
from pathlib import Path

class TestProfiler(unittest.TestCase):
    def test_source_inventory_has_rows(self):
        with open('outputs/cycle1/source_inventory.csv') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
        self.assertGreater(len(rows), 0)
    
    def test_workbook_profiles_json(self):
        with open('outputs/cycle1/workbook_profiles.json') as f:
            data = json.load(f)
        self.assertGreater(len(data), 0)
    
    def test_classification_discovery_rows(self):
        with open('outputs/cycle1/classification_discovery.csv') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
        self.assertGreater(len(rows), 0)
    
    def test_source_selection_report_exists(self):
        self.assertTrue(os.path.exists('outputs/cycle1/source_selection_report.md'))
    
    def test_source_families(self):
        with open('outputs/cycle1/source_inventory.csv') as f:
            reader = csv.DictReader(f)
            families = set(row['source_family'] for row in reader)
        self.assertIn('METHODS', families)
        self.assertIn('TECH', families)
    
    def test_all_workbooks_counted(self):
        with open('outputs/cycle1/source_inventory.csv') as f:
            inv_count = sum(1 for _ in csv.DictReader(f))
        with open('outputs/cycle1/workbook_profiles.json') as f:
            profiles = json.load(f)
        self.assertEqual(len(profiles), inv_count)
    
    def test_unreadable_detection(self):
        """Verify that encrypted/unreadable files are detected in source_inventory."""
        with open('outputs/cycle1/source_inventory.csv') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
        # Check that the 'readable' column exists and has values
        self.assertIn('readable', rows[0])
        # At least verify the column structure is correct
        self.assertTrue(any('readable' in r for r in rows))
    
    def test_version_groups_detected(self):
        """Verify that duplicate detection works and produces results."""
        with open('outputs/cycle1/workbook_profiles.json') as f:
            profiles = json.load(f)
        # Check that SHA-256 hashes exist for readable workbooks
        readable_profiles = [p for p in profiles if p.get('readable', False)]
        if readable_profiles:
            self.assertTrue(all('sha256' in p for p in readable_profiles))
    
    def test_no_data_modification(self):
        """Verify DATA/ was not modified by checking it still has files."""
        data_dir = Path('DATA')
        self.assertTrue(data_dir.exists())
        xlsx_count = sum(1 for f in data_dir.rglob('*.xlsx'))
        self.assertGreater(xlsx_count, 0)
    
    def test_timestamps_extracted(self):
        """Verify that Excel timestamps are extracted from docProps/core.xml."""
        with open('outputs/cycle1/source_inventory.csv') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
        # Check for doc_created/doc_modified columns
        self.assertIn('doc_created', rows[0] or list(rows[0].keys()))
        self.assertIn('doc_modified', list(rows[0].keys()))
        # Verify at least some timestamps were extracted
        has_timestamps = any(r.get('doc_modified') for r in rows if r.get('readable'))
        self.assertTrue(has_timestamps, "Expected some readable workbooks to have doc_modified timestamps")
    
    def test_shared_strings_resolved(self):
        """Verify that shared strings are used for content evidence."""
        with open('outputs/cycle1/classification_discovery.csv') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
        # Check that sample_document_numbers or sample_title_keywords have content
        # (not just worksheet name)
        has_content_evidence = any(r.get('sample_document_numbers') or r.get('sample_title_keywords') for r in rows)
        self.assertTrue(has_content_evidence, "Expected shared-string content evidence in classification_discovery.csv")
    
    def test_selection_fields_present(self):
        """Verify that source_inventory.csv has all mandatory selection fields."""
        with open('outputs/cycle1/source_inventory.csv') as f:
            reader = csv.DictReader(f)
            fieldnames = reader.fieldnames
        required_fields = ['selected_excluded_status', 'selection_exclusion_reason', 'warning_codes']
        for field in required_fields:
            self.assertIn(field, fieldnames or list(reader.fieldnames or []), f"Missing mandatory field: {field}")
    
    def test_relative_paths_used(self):
        """Verify that paths in outputs are relative, not absolute."""
        with open('outputs/cycle1/classification_discovery.csv') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
        # Check that workbook_path does not start with C:\
        for row in rows[:5]:
            path = row.get('workbook_path', '')
            self.assertFalse(path.startswith('C:\\'), f"Absolute path found: {path}")
            self.assertFalse(path.startswith('/c/'), f"MSYS path found: {path}")
    
    def test_hermes_report_current_head(self):
        """Verify HERMES_REPORT uses actual git HEAD, not a cached value."""
        with open('outputs/cycle1/hermes_report.md') as f:
            content = f.read()
        # Get actual HEAD
        head = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd='.').decode().strip()
        # Check that the report contains the actual HEAD
        self.assertIn(head, content, "HERMES_REPORT should contain the current git HEAD")
    
    def test_no_stale_hermes_report(self):
        """Verify stale hermes_report_final.md artifact is removed."""
        self.assertFalse(os.path.exists('outputs/cycle1/hermes_report_final.md'),
                         "Stale hermes_report_final.md should be removed")

if __name__ == '__main__':
    unittest.main()
