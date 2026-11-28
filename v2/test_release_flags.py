"""F10: run-bound flag decisions and non-waivable parser failures."""
import tempfile,unittest,json
from pathlib import Path
from nmdc_profiler.update_engine import stage_update
from v2.launcher import App

class FlagAcceptance(unittest.TestCase):
    def test_report_rejects_overlap_invalid_identity_and_hold_blocks_approval(self):
        with tempfile.TemporaryDirectory() as d:
            root=Path(d);data=root/'project'/'DATA';data.mkdir(parents=True);(data/'a.xlsx').write_text('fixture')
            app=App(root/'workspace')
            try:
                app.settings['data_dir']=str(data);app.save_settings()
                stage_update(data_dir=data,state_dir=app.state,processor=lambda p,r:[{'Event_Key':'E1','Source File':r}],initial_flags=[{'level':'REVIEW','code':'SOURCE_NEW','source':'a.xlsx','message':'new'}])
                p=app.proposal();app.save_choices(p);p=app.proposal();flag=p['flags'][0]
                for folder in (data,data.parent,app.root):
                    app.settings['report_parent']=str(folder)
                    with self.assertRaisesRegex(ValueError,'overlap|outside'):
                        app.report_flags(dict(p,ids=[flag['Flag ID']],comment='check'))
                    self.assertEqual(app.proposal()['flag_version'],p['flag_version'])
                app.settings['report_parent']=str(root/'reports')
                with self.assertRaisesRegex(ValueError,'selection'):
                    app.report_flags(dict(p,ids=['unknown'],comment='check'))
                app.edit_flags(dict(p,changes=[{'id':flag['Flag ID'],'decision':'HOLD FOR REVIEW','status':'DEFERRED','comment':'wait'}]),save=True)
                with self.assertRaisesRegex(ValueError,'Unresolved flags'):app.approve({'run':p['run']})
                self.assertEqual(app.approved()['records'],[])
                current=app.proposal();app.decide_proposal(dict(current,decision='HOLD'),require_versions=True)
                self.assertEqual(app.proposal()['flags'][0]['User Comment'],'wait')
                app.decide_proposal(dict(current,decision='REJECT'),require_versions=True)
                self.assertEqual(app.approved()['records'],[])
            finally:app.invalidate();app.lock.close()
    def test_draft_save_report_clear_retry_preserves_approved_data(self):
        with tempfile.TemporaryDirectory() as d,tempfile.TemporaryDirectory() as reports:
            root=Path(d);data=root/"project"/"DATA";data.mkdir(parents=True)
            source=data/"A.xlsx";source.write_text("original")
            app=App(root/"workspace")
            try:
                app.settings["data_dir"]=str(data);app.settings["report_parent"]=reports;app.save_settings()
                def stage(fault=False):
                    return stage_update(data_dir=data,state_dir=app.state,processor=lambda p,r:[{"Event_Key":"E1","Source File":r,"Document No.":"D1","Document Title":p.read_text(),"Source Sheet":"Register","Source Row":9}],
                        post_process_flags=lambda:[{"level":"REVIEW","code":"UNRECOGNIZED_LAYOUT","source":"A.xlsx","source_sheet":"Register","source_row":9,"source_cell":"C9","message":"Unrecognized column"}] if fault else [])
                stage();p=app.proposal();app.save_choices(p);app.approve({"run":p["run"]})
                source.write_text("changed");stage(True);p=app.proposal();app.save_choices(p);p=app.proposal()
                flag=next(f for f in p["flags"] if f["Flag Code"]=="UNRECOGNIZED_LAYOUT")
                update={"run":p["run"],"flag_version":p["flag_version"],"changes":[{"id":flag["Flag ID"],"decision":"NEEDS PARSER/MAPPING FIX","comment":"Map C9","status":"OPEN"}]}
                app.edit_flags(update,save=False)
                with self.assertRaisesRegex(ValueError,"flag"):app.start(False)
                with self.assertRaisesRegex(ValueError,"flag"):app.approve({"run":p["run"]})
                app.lock.close();app=App(root/"workspace")
                p=app.proposal();self.assertTrue(p["flag_dirty"]);self.assertEqual(next(f for f in p["flags"] if f["Flag ID"]==flag["Flag ID"])["User Comment"],"Map C9")
                with self.assertRaises(ValueError):app.edit_flags(update,save=True)
                app.edit_flags({"run":p["run"],"flag_version":p["flag_version"],"changes":[]},save=True)
                p=app.proposal();self.assertFalse(p["flag_dirty"])
                for i in (1,2):
                    p=app.proposal();r=app.report_flags({"run":p["run"],"flag_version":p["flag_version"],"ids":[flag["Flag ID"]],"comment":"Map C9"})
                    self.assertEqual(r["request_sequence"],f"{i:04d}")
                    md=Path(r["fix_request_file"]).read_text();self.assertIn("Register",md);self.assertIn("C9",md)
                    payload=json.loads((Path(r["fix_request_folder"])/"PARSER_FIX_REQUEST.json").read_text());self.assertEqual(payload["flags"][0]["source_row"],9)
                p=app.proposal();app.edit_flags({"run":p["run"],"flag_version":p["flag_version"],"changes":[{"id":flag["Flag ID"],"decision":"NO ACTION REQUIRED","comment":"","status":"RESOLVED"}]},save=True)
                with self.assertRaisesRegex(ValueError,"Unresolved flags"):app.approve({"run":p["run"]})
                self.assertEqual(app.approved()["records"][0]["Document Title"],"original")
                p=app.proposal();app.edit_flags({"run":p["run"],"flag_version":p["flag_version"],"changes":[{"id":flag["Flag ID"],"decision":"","comment":"","status":"OPEN"}]},save=True)
                self.assertEqual(next(f for f in app.proposal()["flags"] if f["Flag ID"]==flag["Flag ID"])["User Comment"],"")
                stage();p=app.proposal();self.assertFalse(any(f["Flag Code"]=="UNRECOGNIZED_LAYOUT" for f in p["flags"]))
                app.save_choices(p);app.approve({"run":p["run"]})
                self.assertEqual(app.approved()["records"][0]["Document Title"],"changed")
                self.assertEqual(source.read_text(),"changed")
            finally:app.invalidate();app.lock.close()
if __name__=="__main__":unittest.main()
