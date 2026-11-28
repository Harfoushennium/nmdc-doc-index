"""Public job and recovery behaviors; synthetic data only."""
import tempfile, unittest, json
from pathlib import Path
from v2.launcher import App

class JobRecovery(unittest.TestCase):
    def test_interrupted_job_requires_fresh_scan_and_retains_state(self):
        with tempfile.TemporaryDirectory() as d:
            root=Path(d)/"workspace"
            app=App(root)
            try:
                app.settings["theme"]="dark";app.save_settings()
                # Persisted on start before any source work; emulate abrupt termination.
                app.job_path.write_text(json.dumps({"status":"running","mode":"full","phase":"Profile and classify"}),encoding="utf-8")
            finally:app.lock.close()
            restarted=App(root)
            try:
                self.assertEqual(restarted.job["status"],"interrupted")
                self.assertTrue(restarted.settings["requires_rescan"])
                self.assertEqual(restarted.settings["theme"],"dark")
                self.assertEqual(restarted.approved()["records"],[])
            finally:restarted.invalidate();restarted.lock.close()
    def test_repeated_launch_is_refused(self):
        with tempfile.TemporaryDirectory() as d:
            root=Path(d)/"workspace";app=App(root)
            try:
                with self.assertRaisesRegex(ValueError,"already open"):App(root)
            finally:app.lock.close()
    def test_unavailable_source_invalidates_pending_approval(self):
        from nmdc_profiler.update_engine import stage_update
        with tempfile.TemporaryDirectory() as d:
            root=Path(d);data=root/"DATA";data.mkdir();(data/"a.xlsx").write_text("fixture")
            app=App(root/"workspace")
            try:
                app.settings["data_dir"]=str(data);app.save_settings()
                stage_update(data_dir=data,state_dir=app.state,processor=lambda p,r:[{"Event_Key":"E1","Source File":r,"Document No.":"D1"}])
                p=app.proposal();app.save_choices(p)
                data.rename(root/"unavailable")
                with self.assertRaisesRegex(ValueError,"unavailable"):app.start(False)
                self.assertTrue(app.settings["requires_rescan"])
                with self.assertRaisesRegex(ValueError,"rescan"):app.approve({"run":p["run"]})
                self.assertEqual(app.approved()["records"],[])
            finally:app.invalidate();app.lock.close()
if __name__=="__main__":unittest.main()
