"""M01 compatibility preview fixture; no owner data or actual workbook macros executed."""
import tempfile,unittest,shutil,json,csv
from pathlib import Path
import xlsxwriter
from nmdc_profiler.update_engine import stage_update,approve_stage
from v2.launcher import App
from v2.migration_preview import preview_package
from v2.migration import import_package

class MigrationPreviewAcceptance(unittest.TestCase):
    def test_preview_reports_approved_state_and_workbook_only_configuration_read_only(self):
        with tempfile.TemporaryDirectory() as d:
            root=Path(d);legacy=root/"legacy";legacy.mkdir();data=legacy/"DATA";data.mkdir()
            (data/"fixture.xlsx").write_bytes(b"synthetic")
            app=App(root/"template")
            try:shutil.copytree(app.config,legacy/"config")
            finally:app.lock.close()
            runtime=legacy/"runtime"
            staged=stage_update(data_dir=data,state_dir=runtime,processor=lambda p,r:[{"Event_Key":"E1","Source File":r,"Document No.":"0001","Source Sheet":"Register","Source Row":2}])
            approve_stage(runtime,staged["run_id"])
            with xlsxwriter.Workbook(str(legacy/"NMDC_Document_Index.xlsm")) as wb:
                ws=wb.add_worksheet("Configuration")
                ws.add_table(0,0,1,1,{"name":"Configuration","columns":[{"header":"Setting"},{"header":"Current value"}],"data":[["Data Folder",str(data)]]})
            before={p.relative_to(legacy).as_posix():p.read_bytes() for p in legacy.rglob("*") if p.is_file()}
            report=preview_package(legacy)
            self.assertEqual(report["approved_events"],1)
            self.assertEqual(report["approved_run"],staged["run_id"])
            self.assertEqual(report["workbook_tables"]["Configuration"]["rows"],1)
            self.assertEqual(report["blockers"],[])
            self.assertEqual(preview_package(legacy),report)
            self.assertEqual(before,{p.relative_to(legacy).as_posix():p.read_bytes() for p in legacy.rglob("*") if p.is_file()})
            # A pending legacy draft must be reported, not discarded or replayed.
            (data/"fixture.xlsx").write_bytes(b"changed")
            pending=stage_update(data_dir=data,state_dir=runtime,processor=lambda p,r:[{"Event_Key":"E1","Source File":r,"Document No.":"0001","Document Title":"Changed"}])
            with (runtime/"owner_review_drafts.csv").open("w",newline="",encoding="utf-8") as f:
                w=csv.writer(f);w.writerow(["Run ID","Section","Identity","Field","Value"]);w.writerow([pending["run_id"],"FILES","fixture.xlsx","Owner Note","Keep this draft"])
            report=preview_package(legacy)
            self.assertEqual(report["legacy_draft_rows"],1)
            self.assertTrue(any("Active legacy review draft" in x for x in report["blockers"]))
            self.assertEqual(report["approved_run"],staged["run_id"])
            with (runtime/"owner_review_drafts.csv").open("w",newline="",encoding="utf-8") as f:
                w=csv.writer(f);w.writerow(["Run ID","Section","Identity","Field","Value"]);w.writerow([pending["run_id"],"FILES","fixture.xlsx","Accept Update?","TRUE"]);w.writerow([pending["run_id"],"FILES","fixture.xlsx","Owner Note","Keep this draft"])
            report=preview_package(legacy)
            self.assertEqual(report["blockers"],[])
            imported=root/"imported"
            import_package(legacy,imported,report["preview_sha256"])
            translated=json.loads((imported/"state"/"staging"/pending["run_id"]/"web_draft.json").read_text(encoding="utf-8"))
            self.assertTrue(translated["dirty"])
            self.assertEqual(translated["choices"]["fixture.xlsx"],{"accept":True,"note":"Keep this draft"})
            (runtime/"approved"/"current.json").write_text(json.dumps({"run_id":"../escape"}),encoding="utf-8")
            with self.assertRaisesRegex(ValueError,"pointer"):preview_package(legacy)
if __name__=="__main__":unittest.main()
