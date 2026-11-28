"""100k-event public API parity and repeated read-model invalidation."""
import tempfile
import unittest
from pathlib import Path

from nmdc_profiler.update_engine import stage_update
from v2.launcher import App, filter_rows


class ScaleAcceptance(unittest.TestCase):
    def test_100k_events_queries_dashboard_and_repeated_invalidation(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);data=root/'DATA';data.mkdir();(data/'scale.xlsx').write_text('synthetic')
            app=App(root/'workspace')
            try:
                def records(path,source):
                    for index in range(100000):
                        yield {'Event_Key':f'E{index:06d}','Global_Document_Key':f'D{index//10:05d}',
                               'Revision_Key':f'R{index//5:05d}','Document No.':f'{index//10:05d}',
                               'Document Title':'Alpha Beta' if index%2 else 'Alpha Gamma',
                               'Source File':source,'Project No.':'0001','Source Family':'TECH',
                               'Revision':str((index//5)%2),'Source Row':index+2}
                stage_update(data_dir=data,state_dir=app.state,processor=records)
                proposal=app.proposal();app.save_choices(proposal);app.approve(proposal)
                expected=filter_rows(app.view('events'),'Alpha -Gamma')
                for _ in range(3):
                    answer=app.query_page('events','Alpha -Gamma')
                    self.assertEqual(answer['count'],50000)
                    self.assertEqual(answer['rows'],expected[:100])
                    page=app.query_page('events','Alpha -Gamma',page=499)
                    self.assertEqual(page['rows'],expected[49900:50000])
                    self.assertEqual(app.query_page('events',column='Document No.',query='"00001"')['count'],10)
                    app.invalidate()
                metrics=app.dashboard()
                self.assertEqual(metrics['events'],100000)
                self.assertEqual(metrics['documents'],10000)
                self.assertEqual(sum(g['count'] for g in metrics['groups']['Source Family']),10000)
            finally:app.invalidate();app.lock.close()


if __name__=='__main__':unittest.main()
