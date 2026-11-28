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
                self.assertEqual(request('POST','/api/preferences',json.dumps({'theme':'light','density':'giant'}))[0],400)
                self.assertNotIn('theme',app.settings)
                self.assertNotIn('density',app.settings)
                self.assertEqual(request('POST','/api/preferences',json.dumps({'theme':'dark','density':'compact'}))[0],200)
                self.assertEqual(app.settings['density'],'compact')
                self.assertEqual(request('GET','/api/index?scope=unknown')[0],400)
            finally:
                server.shutdown();thread.join();server.server_close();app.invalidate();app.lock.close()

    def test_exit_flushes_http_200_before_starting_shutdown_thread(self):
        from unittest.mock import patch
        with tempfile.TemporaryDirectory() as directory:
            app=App(Path(directory));server=server_for(app)
            thread=threading.Thread(target=server.serve_forever,daemon=True);thread.start()
            origin='http://127.0.0.1:'+str(server.server_port)
            events=[]
            real_send=server.RequestHandlerClass.send
            def spy_send(handler,code,payload,*args,**kwargs):
                events.append(('send',code))
                return real_send(handler,code,payload,*args,**kwargs)
            real_thread=threading.Thread
            class SpyThread(real_thread):
                def start(self):
                    if getattr(self,'_target',None)==server.shutdown:
                        events.append(('shutdown_start',None))
                    return super().start()
            try:
                with patch.object(server.RequestHandlerClass,'send',spy_send), \
                     patch('v2.launcher.threading.Thread',SpyThread):
                    connection=HTTPConnection('127.0.0.1',server.server_port)
                    try:
                        connection.request('POST','/api/exit',body='{}',headers={
                            'Authorization':'Bearer '+app.token,'Origin':origin,'Content-Type':'application/json'})
                        response=connection.getresponse()
                        body=json.loads(response.read().decode('utf-8'))
                        self.assertEqual(response.status,200)
                        self.assertEqual(body,{'status':'closing'})
                    finally:connection.close()
                thread.join(timeout=5)
                self.assertFalse(thread.is_alive())
                self.assertIn(('send',200),events)
                self.assertIn(('shutdown_start',None),events)
                self.assertLess(events.index(('send',200)),events.index(('shutdown_start',None)))
            finally:
                if thread.is_alive():server.shutdown();thread.join()
                server.server_close();app.invalidate();app.lock.close()

    def test_benchmark_release_peak_memory_and_run_uat_trailing_separator(self):
        import os
        from v2.benchmark_release import peak_memory
        rss=peak_memory(os.getpid())
        self.assertIsInstance(rss,int)
        self.assertGreater(rss,0)
        self.assertIsNone(peak_memory(999999999))
        ps1=(Path(__file__).resolve().parent/'Run-UAT.ps1').read_text(encoding='utf-8')
        self.assertIn('TrimEnd',ps1)


if __name__=='__main__':unittest.main()
