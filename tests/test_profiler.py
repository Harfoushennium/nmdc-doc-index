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
    
    def test_duplicate_detection(self):
        with open('outputs/cycle1/source_inventory.csv') as f:
            reader = csv.DictReader(f)
            hashes = [row['sha256'] for row in reader if row['sha256']]
        self.assertEqual(len(hashes), len(set(hashes)))

if __name__ == '__main__':
    unittest.main()
