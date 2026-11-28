"""Public job and recovery behaviors; synthetic data only."""
import tempfile, unittest, json
from pathlib import Path
from v2.launcher import App

class JobRecovery(unittest.TestCase):
    def test_source_changed_after_review_cannot_be_approved(self):
        from nmdc_profiler.update_engine import stage_update
        import os
        with tempfile.TemporaryDirectory() as d:
            root=Path(d);data=root/'DATA';data.mkdir();source=data/'a.xlsx';source.write_text('before')
            app=App(root/'workspace')
            try:
                app.settings['data_dir']=str(data);app.save_settings()
                stage_update(data_dir=data,state_dir=app.state,processor=lambda p,r:[{'Event_Key':'E1','Source File':r,'Document No.':'0001'}])
                p=app.proposal();app.save_choices(p)
                stat=source.stat();source.write_text('after!');os.utime(source,ns=(stat.st_atime_ns,stat.st_mtime_ns))
                with self.assertRaisesRegex(ValueError,'changed.*rescan'):
                    app.approve({'run':p['run']})
                self.assertEqual(app.approved()['records'],[])
                self.assertTrue(app.settings['requires_rescan'])
            finally:app.invalidate();app.lock.close()
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

    def test_stage_finalize_runs_before_publish_proposal_and_latest_pointer(self):
        import time, xlsxwriter
        from unittest.mock import patch
        from nmdc_profiler import update_engine
        with tempfile.TemporaryDirectory() as d:
            root = Path(d); data = root / "DATA"; (data / "METHODS").mkdir(parents=True)
            wb_path = data / "METHODS" / "2701_Methods_Register.xlsx"
            with xlsxwriter.Workbook(str(wb_path)) as wb:
                ws = wb.add_worksheet("Procedures")
                ws.write_row(0, 0, ["Document No.", "Document Title", "Revision", "Event Date"])
                ws.write_row(1, 0, ["2701-PR-001", "Rigging Procedure", "01", "2026-01-10"])
            # 1) Direct finalize_stage hook runs before staging/latest.json and aborts publish on error.
            direct_state = root / "direct_state"
            seen = {}
            def failing_hook(stage_dir, summary):
                seen["latest_exists_in_hook"] = (direct_state / "staging" / "latest.json").exists()
                seen["summary_exists_in_hook"] = (stage_dir / "summary.json").is_file()
                raise RuntimeError("abort pre-publish")
            with self.assertRaisesRegex(RuntimeError, "abort pre-publish"):
                update_engine.stage_update(
                    data_dir=data, state_dir=direct_state,
                    processor=lambda _p, rel: [{"Event_Key": "E1", "Source File": rel}],
                    finalize_stage=failing_hook,
                )
            self.assertFalse(seen["latest_exists_in_hook"])
            self.assertTrue(seen["summary_exists_in_hook"])
            self.assertFalse((direct_state / "staging" / "latest.json").exists())

            # 2) Full App.start chains layout_compat + runtime_engine + unchanged.stage_proof pre-publish.
            app = App(root / "workspace")
            observed = {}
            real_write_atomic = update_engine._write_json_atomic
            def spy_write_atomic(path, payload):
                p = Path(path)
                if p.name == "latest.json" and p.parent.name == "staging":
                    run_id = payload["run_id"]
                    sdir = p.parent / run_id
                    manifest = json.loads((sdir / "manifest.json").read_text(encoding="utf-8"))
                    flags = json.loads((sdir / "flags.json").read_text(encoding="utf-8"))
                    observed["data_dir"] = manifest.get("data_dir")
                    observed["flags"] = flags
                    observed["has_config_hashes"] = (sdir / "stage_config_hashes.json").is_file()
                    observed["has_input_proof"] = (sdir / "input_proof.json").is_file()
                return real_write_atomic(path, payload)
            try:
                app.settings["data_dir"] = str(data); app.save_settings()
                with patch("nmdc_profiler.update_engine._write_json_atomic", side_effect=spy_write_atomic):
                    app.start(True)
                    deadline = time.time() + 15
                    while app.job["status"] == "running" and time.time() < deadline:
                        time.sleep(0.05)
                self.assertEqual(app.job["status"], "complete", app.job)
                self.assertEqual(observed.get("data_dir"), str(data.resolve()))
                self.assertEqual(observed.get("flags"), [])
                self.assertTrue(observed.get("has_config_hashes"))
                self.assertTrue(observed.get("has_input_proof"))
            finally:
                app.invalidate(); app.lock.close()

    def test_flagged_incremental_stage_blocks_out_of_band_source_exclusions_edit(self):
        import time, xlsxwriter
        with tempfile.TemporaryDirectory() as d:
            root = Path(d); data = root / "DATA"; (data / "METHODS").mkdir(parents=True)
            wb_path = data / "METHODS" / "2702_Methods_Register.xlsx"
            with xlsxwriter.Workbook(str(wb_path)) as wb:
                ws = wb.add_worksheet("Procedures")
                ws.write_row(0, 0, ["Document No.", "Document Title", "Revision", "Event Date"])
                ws.write_row(1, 0, ["2702-PR-001", "Initial Title", "01", "2026-01-10"])
            app = App(root / "workspace")
            try:
                app.settings["data_dir"] = str(data); app.save_settings()
                app.start(True)
                deadline = time.time() + 15
                while app.job["status"] == "running" and time.time() < deadline:
                    time.sleep(0.05)
                self.assertEqual(app.job["status"], "complete", app.job)
                p1 = app.proposal(); app.save_choices(p1); app.approve({"run": p1["run"]})
                # Create an incremental stage with a routine flag present (so input_proof.json is absent).
                with xlsxwriter.Workbook(str(wb_path)) as wb:
                    ws = wb.add_worksheet("Procedures")
                    ws.write_row(0, 0, ["Document No.", "Document Title", "Revision", "Event Date"])
                    ws.write_row(1, 0, ["2702-PR-001", "Updated Title", "01", "2026-01-11"])
                app.start(False)
                deadline = time.time() + 15
                while app.job["status"] == "running" and time.time() < deadline:
                    time.sleep(0.05)
                self.assertEqual(app.job["status"], "complete", app.job)
                p2 = app.proposal()
                stage_dir = app.state / "staging" / p2["run"]
                self.assertTrue((stage_dir / "stage_config_hashes.json").is_file())
                # Simulate a flagged stage without input_proof.json.
                (stage_dir / "input_proof.json").unlink(missing_ok=True)
                app.save_choices(p2)
                # Out-of-band edit to source_exclusions.csv after review.
                excl = app.config / "source_exclusions.csv"
                excl.write_text(excl.read_text(encoding="utf-8") + "METHODS/2702_Methods_Register.xlsx,EXCLUDED,Out-of-band edit\n", encoding="utf-8")
                with self.assertRaisesRegex(ValueError, "changed.*rescan"):
                    app.approve({"run": p2["run"]})
                self.assertTrue(app.settings["requires_rescan"])
                self.assertEqual(app.approved()["run_id"], p1["run"])
            finally:
                app.invalidate(); app.lock.close()

    def test_locked_source_permission_error_protects_approved_index(self):
        import time, xlsxwriter
        from unittest.mock import patch
        from nmdc_profiler import update_engine
        with tempfile.TemporaryDirectory() as d:
            root = Path(d); data = root / "DATA"; (data / "METHODS").mkdir(parents=True)
            wb_path = data / "METHODS" / "2703_Methods_Register.xlsx"
            with xlsxwriter.Workbook(str(wb_path)) as wb:
                ws = wb.add_worksheet("Procedures")
                ws.write_row(0, 0, ["Document No.", "Document Title", "Revision", "Event Date"])
                ws.write_row(1, 0, ["2703-PR-001", "Locked Test", "01", "2026-01-10"])
            app = App(root / "workspace")
            try:
                app.settings["data_dir"] = str(data); app.save_settings()
                app.start(True)
                deadline = time.time() + 15
                while app.job["status"] == "running" and time.time() < deadline:
                    time.sleep(0.05)
                self.assertEqual(app.job["status"], "complete", app.job)
                p1 = app.proposal(); app.save_choices(p1); app.approve({"run": p1["run"]})
                approved_before = list(app.approved()["records"])
                self.assertEqual(len(approved_before), 1)

                real_sha = update_engine.sha256_file
                locked_target = wb_path.resolve()
                def locked_sha(path, *args, **kwargs):
                    if Path(path).resolve() == locked_target:
                        raise PermissionError(13, "Sharing violation: file locked")
                    return real_sha(path, *args, **kwargs)
                with patch("nmdc_profiler.source_access.time.sleep", return_value=None), \
                     patch("nmdc_profiler.source_cache.time.sleep", return_value=None), \
                     patch("nmdc_profiler.update_engine.sha256_file", side_effect=locked_sha), \
                     patch("v2.unchanged.sha256_file", side_effect=locked_sha):
                    app.start(False)
                    deadline = time.time() + 15
                    while app.job["status"] == "running" and time.time() < deadline:
                        time.sleep(0.05)
                self.assertEqual(app.job["status"], "error", app.job)
                self.assertTrue(app.settings["requires_rescan"])
                self.assertEqual(app.approved()["records"], approved_before)
                latest = json.loads((app.state / "staging" / "latest.json").read_text(encoding="utf-8"))
                self.assertEqual(latest["run_id"], p1["run"])

                # Once unlocked, a fresh scan succeeds cleanly.
                app.start(False)
                deadline = time.time() + 15
                while app.job["status"] == "running" and time.time() < deadline:
                    time.sleep(0.05)
                self.assertEqual(app.job["status"], "complete", app.job)
            finally:
                app.invalidate(); app.lock.close()
if __name__=="__main__":unittest.main()
