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
                         'Document Link': 'javascript:bad', 'Event Date':'1899-12-31', 'Source File': r,
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
        self.assertEqual(plan['date_exceptions'][0]['value'],'1899-12-31')
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

    def test_xlsx_export_requires_server_side_literal_dates_confirmation(self):
        import io, zipfile
        q = {'scope': 'events', 'format': 'xlsx'}
        status, body = self.get('/api/export-plan', q)
        self.assertEqual(status, 200)
        plan = json.loads(body)
        self.assertEqual(len(plan['date_exceptions']), 1)
        base = dict(q, preview=plan['preview'], run=plan['run'], literal_links='confirmed')
        # Without literal_dates=confirmed, server must reject with 400.
        st_missing, err_body = self.get('/api/export', base)
        self.assertEqual(st_missing, 400, err_body)
        # Wrong run with literal_dates=confirmed must also reject with 400.
        st_wrong_run, _ = self.get('/api/export', dict(base, run='stale-run', literal_dates='confirmed'))
        self.assertEqual(st_wrong_run, 400)
        # With literal_dates=confirmed and matching run/preview, server returns valid XLSX.
        st_ok, data = self.get('/api/export', dict(base, literal_dates='confirmed'))
        self.assertEqual(st_ok, 200, data)
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            self.assertIsNone(z.testzip())
            self.assertIn('xl/tables/table1.xml', z.namelist())


if __name__ == '__main__':
    unittest.main()
