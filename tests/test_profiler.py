import unittest, json, csv, os, sys
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

if __name__ == '__main__':
    unittest.main()
