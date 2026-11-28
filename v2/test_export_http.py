"""Export consent must be invalidated by a different request or configuration."""
import json
import tempfile
import threading
import unittest
from http.client import HTTPConnection
from pathlib import Path
from urllib.parse import urlencode

from nmdc_profiler.update_engine import stage_update
from v2.launcher import App, server_for


class ExportHTTPAcceptance(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        root = Path(self.tmp.name)
        self.app = App(root / 'workspace')
        self.addCleanup(self.app.lock.close)
        self.addCleanup(self.app.invalidate)
        data = root / 'DATA'; data.mkdir(); (data / 'a.xlsx').write_text('fixture')
        stage_update(data_dir=data, state_dir=self.app.state,
                     processor=lambda p, r: [{'Event_Key': 'E1', 'Document No.': '0001',
                         'Document Link': 'javascript:bad', 'Source File': r,
                         'Source Sheet': 'Register', 'Source Row': 7, 'Source Cell': 'C7'}])
        p = self.app.proposal(); self.app.save_choices(p); self.app.approve(p)
        self.server = server_for(self.app)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.addCleanup(self.stop)

    def stop(self):
        self.server.shutdown(); self.thread.join(); self.server.server_close()

    def get(self, route, query):
        connection = HTTPConnection('127.0.0.1', self.server.server_port)
        try:
            connection.request('GET', route + '?' + urlencode(query), headers={'Authorization': 'Bearer ' + self.app.token})
            response = connection.getresponse()
            return response.status, response.read()
        finally:
            connection.close()

    def test_preview_pins_query_columns_format_config_and_snapshot(self):
        q = {'scope': 'events', 'format': 'json'}
        status, body = self.get('/api/export-plan', q); self.assertEqual(status, 200)
        plan = json.loads(body)
        self.assertIn('not been checked',plan['availability'])
        self.assertEqual(plan['exceptions'][0]['worksheet'], 'Register')
        self.assertEqual(plan['exceptions'][0]['cell'], 'C7')
        q['preview'] = plan['preview']
        status, body = self.get('/api/export', q)
        self.assertEqual(status, 200); self.assertEqual(json.loads(body)[0]['Document No.'], '0001')
        for change in ({'q': 'different'}, {'columns': 'Event_Key'}, {'format': 'csv'}):
            self.assertEqual(self.get('/api/export', dict(q, **change))[0], 400)
        (self.app.config / 'custom_fields.json').write_text('{"fields":[],"mappings":[]}')
        self.assertEqual(self.get('/api/export', q)[0], 400)
        q.pop('preview'); status, body = self.get('/api/export-plan', q)
        q['preview'] = json.loads(body)['preview']
        self.assertEqual(self.get('/api/export', q)[0], 200)
        self.assertEqual(self.get('/api/export', {'scope': 'events', 'format': 'json'})[0], 400)


if __name__ == '__main__':
    unittest.main()
