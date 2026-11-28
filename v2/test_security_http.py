"""Loopback HTTP rejects foreign origins, hosts, credentials and malformed input."""
import json
import tempfile
import threading
import unittest
from http.client import HTTPConnection
from pathlib import Path

from v2.launcher import App, server_for


class SecurityHTTPAcceptance(unittest.TestCase):
    def test_host_origin_token_and_input_guards(self):
        with tempfile.TemporaryDirectory() as directory:
            app=App(Path(directory));server=server_for(app)
            thread=threading.Thread(target=server.serve_forever,daemon=True);thread.start()
            origin='http://127.0.0.1:'+str(server.server_port)
            def request(method,path,body=None,**changes):
                headers={'Authorization':'Bearer '+app.token,'Origin':origin,'Content-Type':'application/json'}
                headers.update(changes)
                connection=HTTPConnection('127.0.0.1',server.server_port)
                try:
                    connection.request(method,path,body=body,headers=headers)
                    response=connection.getresponse();data=response.read()
                    return response.status,data
                finally:connection.close()
            try:
                self.assertEqual(server.server_address[0],'127.0.0.1')
                for header in ({'Authorization':'Bearer foreign'},{'Authorization':''},
                               {'Origin':'https://foreign.invalid'},{'Host':'foreign.invalid'}):
                    self.assertEqual(request('GET','/api/status',**header)[0],403)
                self.assertEqual(request('POST','/api/backup','{}',Origin='')[0],403)
                for body in ('[]','null','"text"','{bad'):
                    self.assertEqual(request('POST','/api/backup',body)[0],400)
                self.assertEqual(request('GET','/../../settings.json')[0],404)
                self.assertEqual(request('POST','/api/preferences',json.dumps({'theme':'<script>'}))[0],400)
                self.assertNotIn('theme',app.settings)
                self.assertEqual(request('GET','/api/index?scope=unknown')[0],400)
            finally:
                server.shutdown();thread.join();server.server_close();app.invalidate();app.lock.close()


if __name__=='__main__':unittest.main()
