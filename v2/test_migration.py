import tempfile
import unittest
import shutil
import csv
from pathlib import Path

from v2.recovery import apply_retention, create_bundle, inventory, restore_to_new_workspace
from v2.migration import import_package
from v2.migration_preview import preview_package
from nmdc_profiler.update_engine import stage_update, approve_stage
from nmdc_profiler.excel_bridge import load_approved_state
from v2.launcher import App
from v2.recovery import reset_workspace, undo_approval


class MigrationPublicApiRegression(unittest.TestCase):
    def test_import_requires_confirmed_preview_and_never_overwrites_target(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            legacy = root / "legacy"
            legacy.mkdir()
            (legacy / "runtime").mkdir()
            (legacy / "config").mkdir()
            target = root / "new-workspace"
            report = preview_package(legacy)
            self.assertTrue(report["blockers"])
            with self.assertRaises(ValueError):
                import_package(legacy, target, report["preview_sha256"])
            self.assertFalse(target.exists())

    def test_recovery_inventory_includes_derived_custom_field_store(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            (root / "custom-fields.json").write_text('{"fields":[]}', encoding="utf-8")
            before = inventory(root)
            bundle = create_bundle(root, "test")
            self.assertIn("custom-fields.json", before)
            self.assertEqual(bundle["files"], 1)

    def test_confirmed_import_is_side_by_side_and_reconciles_bytes(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            legacy = root / "legacy"
            legacy.mkdir()
            data = legacy / "DATA"
            data.mkdir()
            (data / "fixture.xlsx").write_bytes(b"fixture")
            app = App(root / "template")
            try:
                shutil.copytree(app.config, legacy / "config")
            finally:
                app.lock.close()
            runtime = legacy / "runtime"
            staged = stage_update(data_dir=data, state_dir=runtime,
                processor=lambda _p, rel: [{"Event_Key":"EV-1","Source File":rel,"Document No.":"0007"}])
            approve_stage(runtime, staged["run_id"])
            original = {p.relative_to(legacy).as_posix(): p.read_bytes() for p in legacy.rglob("*") if p.is_file()}
            report = preview_package(legacy)
            self.assertEqual(report["blockers"], [])
            result = import_package(legacy, root / "v2-workspace", report["preview_sha256"])
            self.assertEqual(result["approved_events"], 1)
            self.assertEqual(load_approved_state(root / "v2-workspace" / "state")["records"][0]["Event_Key"], "EV-1")
            self.assertEqual(original, {p.relative_to(legacy).as_posix(): p.read_bytes() for p in legacy.rglob("*") if p.is_file()})
            self.assertTrue((root / "v2-workspace" / "state" / "approved" / "current.json").is_file())
            with self.assertRaisesRegex(ValueError, "already exists"):
                import_package(legacy, root / "v2-workspace", report["preview_sha256"])
            config_file=legacy/"config"/"classification_rules.csv"
            config_file.write_bytes(config_file.read_bytes()+b"\n")
            with self.assertRaisesRegex(ValueError, "changed|preview"):
                import_package(legacy, root / "stale-preview-target", report["preview_sha256"])
            self.assertFalse((root/"stale-preview-target").exists())

    def test_recovery_refuses_symlink_inputs_and_retention_refuses_unknown_zip(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            (root / "state").mkdir()
            target = root / "outside"
            target.mkdir()
            try:
                (root / "state" / "approved").symlink_to(target, target_is_directory=True)
            except OSError:
                self.skipTest("Symlink creation unavailable")
            with self.assertRaisesRegex(ValueError, "Symlink|reparse"):
                inventory(root)

    def test_reset_preserves_settings_and_audit_and_creates_verified_backup(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            state = root / "state"
            (state / "approved").mkdir(parents=True)
            (state / "approved" / "records.jsonl").write_text("{}\n", encoding="utf-8")
            (state / "logs").mkdir()
            (state / "logs" / "history.jsonl").write_text('{"event":"old"}\n', encoding="utf-8")
            (state / "user_flags").mkdir()
            (state / "user_flags" / "flags.jsonl").write_text('{"id":"F1"}\n', encoding="utf-8")
            (root / "settings.json").write_text('{"theme":"dark"}', encoding="utf-8")
            result = reset_workspace(root, "RESET")
            self.assertTrue(Path(result["backup"]).is_file())
            self.assertFalse((state / "approved").exists())
            self.assertTrue((state / "logs" / "history.jsonl").is_file())
            self.assertTrue((state / "user_flags" / "flags.jsonl").is_file())
            self.assertEqual((root / "settings.json").read_text(encoding="utf-8"), '{"theme":"dark"}')

    def test_verified_restore_and_retention_keep_only_newest_verified_bundles(self):
        with tempfile.TemporaryDirectory() as td:
            root=Path(td);(root/"config").mkdir();(root/"config"/"rules.csv").write_text("a,b\n",encoding="utf-8")
            first=create_bundle(root,"first");second=create_bundle(root,"second");third=create_bundle(root,"third")
            retained=apply_retention(root/"backups",keep=2)
            self.assertEqual(len(retained),2)
            self.assertFalse(Path(first["path"]).exists())
            restored=restore_to_new_workspace(third["path"],root/"restored")
            self.assertEqual(restored["files"],1)
            self.assertEqual((root/"restored"/"config"/"rules.csv").read_text(encoding="utf-8"),"a,b\n")

    def test_undo_moves_pointer_back_without_deleting_approved_versions(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            data = root / "DATA"
            data.mkdir()
            source = data / "one.xlsx"
            source.write_bytes(b"v1")
            state = root / "state"
            one = stage_update(data_dir=data, state_dir=state, processor=lambda _p, rel: [{"Event_Key":"E1","Source File":rel}])
            approve_stage(state, one["run_id"])
            source.write_bytes(b"v2")
            two = stage_update(data_dir=data, state_dir=state, processor=lambda _p, rel: [{"Event_Key":"E2","Source File":rel}])
            approve_stage(state, two["run_id"])
            result = undo_approval(state, "UNDO")
            self.assertEqual(result["restored_run_id"], one["run_id"])
            self.assertTrue((state / "approved" / "versions" / two["run_id"] / "records.jsonl").is_file())


if __name__ == "__main__":
    unittest.main()
