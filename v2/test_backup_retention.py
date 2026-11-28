import json
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import Mock, patch
from zipfile import ZipFile

from v2.recovery import apply_retention, create_bundle, restore_to_new_workspace


class BackupRetentionTests(unittest.TestCase):
    def test_newest_recovery_points_survive_clock_rollback(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory).resolve()
            (root / 'config').mkdir()
            rules = root / 'config' / 'rules.csv'
            clock = Mock(wraps=time)
            clock.time.side_effect = [1700000100, 1700000000, 1700000001]
            with patch('v2.recovery.time', clock):
                rules.write_text('first')
                first = create_bundle(root, 'first', keep=2)
                rules.write_text('second')
                second = create_bundle(root, 'second', keep=2)
                rules.write_text('third')
                third = create_bundle(root, 'third', keep=2)
            self.assertEqual(apply_retention(root / 'backups', keep=2), [third['path'], second['path']])
            self.assertFalse(Path(first['path']).exists())
            restore_to_new_workspace(third['path'], root / 'restored')
            self.assertEqual((root / 'restored/config/rules.csv').read_text(), 'third')

    def test_invalid_retention_metadata_is_preserved_and_reported(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory).resolve()
            backups = root / 'backups'
            backups.mkdir()
            malformed = backups / 'old.zip'
            with ZipFile(malformed, 'w') as archive:
                archive.writestr('recovery-manifest.json', json.dumps({
                    'version': 1, 'files': {}, 'created_at': 'invalid',
                }))
            result = create_bundle(root, 'new', keep=1)
            self.assertTrue(Path(result['path']).is_file())
            self.assertTrue(malformed.is_file())
            self.assertEqual([item['file'] for item in result['retention_skipped']], ['old.zip'])
