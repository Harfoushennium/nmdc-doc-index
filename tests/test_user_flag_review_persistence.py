import csv
import json
import tempfile
import unittest
from pathlib import Path

from nmdc_profiler.excel_bridge import build_user_flag_rows
from nmdc_profiler.review_decisions import apply_review_decisions


class UserFlagReviewPersistenceTests(unittest.TestCase):
    def fixture(self, root):
        stage = root / 'staging' / 'RUN-1'
        stage.mkdir(parents=True)
        (stage.parent / 'latest.json').write_text(json.dumps({'run_id': 'RUN-1'}))
        (stage / 'flags.json').write_text('[]\n')
        flags = root / 'user_flags' / 'flags.jsonl'
        flags.parent.mkdir()
        flags.write_text(json.dumps({'code': 'USER_FLAGGED_WRONG_DATA', 'source': 'METHODS/test.xlsx',
                                     'resolution_route': 'REVIEW_REQUIRED', 'resolution_status': 'OPEN',
                                     'user_comment': 'Original test note'}) + '\n')
        return flags

    def decisions(self, root, **values):
        path = root / 'decisions.csv'
        row = {'Flag Code': 'USER_FLAGGED_WRONG_DATA', 'Source File': 'METHODS/test.xlsx', **values}
        with path.open('w', newline='', encoding='utf-8') as handle:
            writer = csv.DictWriter(handle, fieldnames=list(row))
            writer.writeheader()
            writer.writerow(row)
        return path

    def test_user_flag_decision_survives_export_reload(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            self.fixture(root)
            result = apply_review_decisions(root, self.decisions(root, **{
                'User Decision': 'NO ACTION REQUIRED', 'User Comment': 'Acceptance test only',
                'Resolution Status': 'RESOLVED'}), expected_run_id='RUN-1')
            self.assertEqual(result['updated_flags'], 1)
            row = build_user_flag_rows(root)[0]
            self.assertEqual(row['User Decision'], 'NO ACTION REQUIRED')
            self.assertEqual(row['User Comment'], 'Acceptance test only')
            self.assertEqual(row['Resolution Status'], 'RESOLVED')

    def test_explicit_clear_does_not_restore_legacy_route(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            self.fixture(root)
            result = apply_review_decisions(root, self.decisions(root, **{
                'Clear User Decision': 'TRUE', 'Clear User Comment': 'TRUE',
                'Clear Resolution Status': 'TRUE'}))
            self.assertEqual(result['updated_flags'], 1)
            row = build_user_flag_rows(root)[0]
            self.assertEqual([row[k] for k in ('User Decision', 'User Comment', 'Resolution Status')], ['', '', ''])

    def test_unmatched_decision_fails_without_rewriting_audit(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            flags = self.fixture(root)
            stage_flags = root / 'staging' / 'RUN-1' / 'flags.json'
            before = (flags.read_bytes(), stage_flags.read_bytes())
            with self.assertRaisesRegex(ValueError, 'no longer match'):
                apply_review_decisions(root, self.decisions(root, **{
                    'Source File': 'missing.xlsx', 'User Decision': 'ACKNOWLEDGED'}))
            self.assertEqual((flags.read_bytes(), stage_flags.read_bytes()), before)
            self.assertFalse((root / 'logs' / 'history.jsonl').exists())
