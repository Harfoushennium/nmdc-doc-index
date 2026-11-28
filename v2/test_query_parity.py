"""Parity checks between SQLite ReadModel and App.query_page Python fallback."""
import json
import tempfile
import threading
import unittest
from http.client import HTTPConnection
from pathlib import Path
from urllib.parse import urlencode

from nmdc_profiler.update_engine import stage_update
from v2.launcher import App, server_for


class QueryParityAcceptance(unittest.TestCase):
    def _build_app(self, root: Path) -> App:
        data = root / "DATA"
        data.mkdir()
        (data / "fixture.xlsx").write_bytes(b"synthetic")
        app = App(root / "workspace")
        rows = [
            {
                "Event_Key": "E1",
                "Global_Document_Key": "D1",
                "Revision_Key": "R1",
                "Document No.": "DOC-A",
                "Document Title": "ExactToken",
                "Discipline": "Marine",
                "Project No.": "0001",
                "Revision": "Z",
                "Source File": "fixture.xlsx",
                "Source Row": 2,
            },
            {
                "Event_Key": "E2",
                "Global_Document_Key": "D2",
                "Revision_Key": "R2",
                "Document No.": "DOC-B",
                "Document Title": "ExactToken",
                "Discipline": "MARINE",
                "Project No.": "0001",
                "Revision": "A",
                "Source File": "fixture.xlsx",
                "Source Row": 3,
            },
            {
                "Event_Key": "E3",
                "Global_Document_Key": "D3",
                "Revision_Key": "R3",
                "Document No.": "DOC-C",
                "Document Title": "OtherTitle",
                "Discipline": "marine",
                "Project No.": "0001",
                "Revision": "M",
                "Source File": "fixture.xlsx",
                "Source Row": 4,
            },
        ]
        stage_update(data_dir=data, state_dir=app.state, processor=lambda _p, _rel: rows)
        proposal = app.proposal()
        app.save_choices(proposal)
        app.approve(proposal)
        return app

    def test_python_fallback_validates_sort_and_group_against_model_fields(self):
        with tempfile.TemporaryDirectory() as td:
            app = self._build_app(Path(td))
            try:
                with self.assertRaises(ValueError):
                    app.query_page("events", query="ExactToken", sort="UnknownField")
                with self.assertRaises(ValueError):
                    app.query_page("events", query='"ExactToken"', sort="UnknownField")
                with self.assertRaises(ValueError):
                    app.query_page("events", query='"ExactToken"', group="UnknownField")
                server = server_for(app)
                thread = threading.Thread(target=server.serve_forever, daemon=True)
                thread.start()
                try:
                    conn = HTTPConnection("127.0.0.1", server.server_port)
                    try:
                        qs = urlencode({"scope": "events", "q": '"ExactToken"', "sort": "UnknownField"})
                        conn.request("GET", "/api/index?" + qs, headers={"Authorization": "Bearer " + app.token})
                        resp = conn.getresponse()
                        resp.read()
                        self.assertEqual(resp.status, 400)
                    finally:
                        conn.close()
                finally:
                    server.shutdown()
                    thread.join()
                    server.server_close()
            finally:
                app.invalidate()
                app.lock.close()

    def test_python_fallback_group_labels_match_sqlite_min_position_before_sorting(self):
        with tempfile.TemporaryDirectory() as td:
            app = self._build_app(Path(td))
            try:
                for desc in (False, True):
                    sqlite_res = app.query_page("events", query="ExactToken", sort="Revision", desc=desc, group="Discipline")
                    fallback_res = app.query_page("events", query='"ExactToken"', sort="Revision", desc=desc, group="Discipline")
                    self.assertEqual(sqlite_res["count"], 2)
                    self.assertEqual(fallback_res["count"], 2)
                    self.assertEqual(
                        [r["Event_Key"] for r in fallback_res["rows"]],
                        [r["Event_Key"] for r in sqlite_res["rows"]],
                    )
                    self.assertEqual(sqlite_res["groups"], [{"value": "Marine", "count": 2}])
                    self.assertEqual(fallback_res["groups"], sqlite_res["groups"])
            finally:
                app.invalidate()
                app.lock.close()


if __name__ == "__main__":
    unittest.main()
