"""Two independent HTTP sessions must act on exactly the displayed review."""
import copy
import json
import threading
import unittest
from http.client import HTTPConnection

from v2.launcher import server_for
from v2.test_release_review import ReviewAcceptance


class ReviewHTTPAcceptance(ReviewAcceptance):
    def setUp(self):
        super().setUp()
        self.server = server_for(self.app)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.addCleanup(self.stop_server)

    def stop_server(self):
        self.server.shutdown()
        self.thread.join()
        self.server.server_close()

    def request(self, path, body=None):
        client = HTTPConnection('127.0.0.1', self.server.server_port)
        headers = {'Authorization': 'Bearer ' + self.app.token,
                   'Origin': 'http://127.0.0.1:' + str(self.server.server_port),
                   'Content-Type': 'application/json'}
        try:
            client.request('GET' if body is None else 'POST', path,
                           None if body is None else json.dumps(body), headers)
            response = client.getresponse()
            return response.status, json.loads(response.read())
        finally:
            client.close()

    def pending(self):
        self.write('A.xlsx', 'original')
        self.stage()
        self.save(self.app.proposal())
        return self.request('/api/review')[1]

    def test_two_sessions_stale_file_and_flag_approval(self):
        old = self.pending()
        second = copy.deepcopy(old)
        second['files'][0]['Owner Note'] = 'second session'
        self.assertEqual(self.request('/api/save-choices', second)[0], 200)
        self.assertEqual(self.request('/api/approve', old)[0], 400)
        self.assertEqual(self.app.approved()['run_id'], '')
        current = self.request('/api/review')[1]
        self.assertEqual(self.request('/api/flag-save', dict(current, changes=[]))[0], 200)
        self.assertEqual(self.request('/api/approve', current)[0], 400)
        self.assertEqual(self.request('/api/approve', {'run': current['run']})[0], 400)
        self.assertEqual(self.app.approved()['run_id'], '')
        current = self.request('/api/review')[1]
        self.assertEqual(self.request('/api/approve', current)[0], 200)
        self.assertEqual(self.values(), {'A.xlsx': 'original'})

    def test_stale_hold_and_reject_leave_review_and_approved_unchanged(self):
        old = self.pending()
        self.save(self.app.proposal())
        for decision in ('HOLD', 'REJECT'):
            with self.subTest(decision=decision):
                status, result = self.request('/api/proposal-decision', dict(old, decision=decision))
                self.assertEqual(status, 400, result)
                self.assertEqual(self.app.proposal()['run'], old['run'])
                self.assertEqual(self.app.approved()['run_id'], '')

    def test_stale_discard_cannot_erase_newer_flag_review(self):
        old = self.pending()
        self.app.edit_flags(dict(old, changes=[]), save=True)
        self.assertEqual(self.request('/api/discard-draft', dict(old, confirm=True))[0], 400)
        self.assertEqual(self.app.proposal()['run'], old['run'])


if __name__ == '__main__':
    unittest.main()
