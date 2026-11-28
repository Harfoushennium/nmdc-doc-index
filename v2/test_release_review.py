"""RP1 #33 behavioral acceptance through the App's public workflow."""
import tempfile,unittest
from pathlib import Path
from nmdc_profiler.update_engine import stage_update
from v2.launcher import App

class ReviewAcceptance(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        root=Path(self.tmp.name);self.data=root/"DATA";self.data.mkdir()
        self.app=App(root/"workspace");self.addCleanup(self.app.lock.close);self.addCleanup(self.app.invalidate)
        self.app.settings["data_dir"]=str(self.data);self.app.save_settings()
    def write(self,name,value): (self.data/name).write_text(value,encoding="utf-8")
    def stage(self):
        return stage_update(data_dir=self.data,state_dir=self.app.state,processor=lambda p,r:[{"Event_Key":r,"Source File":r,"Document No.":r,"Project No.":"001","Revision":"A","Event Status":p.read_text(encoding="utf-8"),"Source Sheet":"Register","Source Row":2,"Source Cell":"A2"}])
    def save(self,p):
        self.app.save_choices({"run":p["run"],"files":p["files"],"draft_version":p["draft_version"]})
    def approve(self):
        p=self.app.proposal();self.save(p);return self.app.approve({"run":p["run"]})
    def values(self):return {r["Source File"]:r["Event Status"] for r in self.app.view("events")}
    def test_dirty_blocks_scan_and_approval_and_discard_archives(self):
        self.write("A.xlsx","old");self.stage();p=self.app.proposal()
        p["files"][0]["Owner Note"]="do not lose";p["files"][0]["Accept Update?"]=False
        self.app.draft_choices(p);p=self.app.proposal()
        self.assertEqual(p["files"][0]["Owner Note"],"do not lose")
        with self.assertRaisesRegex(ValueError,"Save or explicitly discard"):self.app.start(False)
        with self.assertRaisesRegex(ValueError,"Save all"):self.app.approve({"run":p["run"]})
        self.app.discard_draft({"run":p["run"],"draft_version":p["draft_version"],"confirm":True})
        self.assertFalse(self.app.proposal()["draft_dirty"])
        self.assertTrue(any(d["choices"].get("A.xlsx",{}).get("note")=="do not lose" and "discarded" in d["file"] for d in self.app.draft_history()))
        self.assertEqual(self.values(),{})
    def test_mixed_change_remove_and_new_deferral(self):
        for name in ("A.xlsx","B.xlsx","C.xlsx"):self.write(name,"old "+name)
        self.stage();self.approve()
        self.write("A.xlsx","new A");self.write("B.xlsx","new B");(self.data/"C.xlsx").unlink();self.write("D.xlsx","new D")
        self.stage();p=self.app.proposal()
        details=self.app.file_details(p["run"],"A.xlsx")
        self.assertEqual(details["count"],1)
        self.assertEqual(details["rows"][0]["before"]["Event Status"],"old A.xlsx")
        self.assertEqual(details["rows"][0]["after"]["Event Status"],"new A")
        self.assertEqual(details["rows"][0]["fields"],["Event Status"])
        for f in p["files"]:f["Accept Update?"]=f["Source File"]=="A.xlsx"
        self.save(p);self.app.approve({"run":p["run"]})
        self.assertEqual(self.values(),{"A.xlsx":"new A","B.xlsx":"old B.xlsx","C.xlsx":"old C.xlsx"})
        self.assertFalse((self.data/"C.xlsx").exists())
        self.stage();p=self.app.proposal();self.assertEqual({f["Source File"] for f in p["files"]},{"B.xlsx","C.xlsx","D.xlsx"})
    def test_accepted_removal_requires_documented_intent(self):
        self.write("A.xlsx","A");self.write("B.xlsx","B");self.stage();self.approve()
        (self.data/"A.xlsx").unlink();self.stage();p=self.app.proposal();self.save(p)
        with self.assertRaisesRegex(ValueError,"document removal"):self.app.approve({"run":p["run"]})
        self.assertEqual(self.values(),{"A.xlsx":"A","B.xlsx":"B"})
        p=self.app.proposal();p["files"][0]["Owner Note"]="Intentional removal confirmed";self.save(p)
        self.app.approve({"run":p["run"]});self.assertEqual(self.values(),{"B.xlsx":"B"})
    def test_hold_reject_preserve_approved(self):
        self.write("A.xlsx","old");self.stage();self.approve()
        self.write("A.xlsx","new");self.stage();p=self.app.proposal()
        self.app.decide_proposal({"run":p["run"],"decision":"HOLD","note":"Await review"})
        self.assertEqual(self.values(),{"A.xlsx":"old"})
        self.app.decide_proposal({"run":p["run"],"decision":"REJECT","note":"Do not accept"})
        self.assertEqual(self.values(),{"A.xlsx":"old"});self.assertEqual(self.app.proposal()["run"],"")

    def test_stale_displayed_approval_rejected(self):
        self.write("A.xlsx","original");self.stage();p=self.app.proposal();self.save(p)
        old=self.app.proposal();new=self.app.proposal()
        new["files"][0]["Owner Note"]="Other tab changed the review";self.save(new)
        with self.assertRaisesRegex(ValueError,"Displayed review version changed"):
            self.app.approve({k:old[k] for k in ("run","draft_version","flag_version")},require_versions=True)
        self.assertEqual(self.values(),{})
        current=self.app.proposal()
        self.app.edit_flags({"run":current["run"],"flag_version":current["flag_version"],"changes":[]},save=True)
        with self.assertRaisesRegex(ValueError,"Displayed review version changed"):
            self.app.approve({k:current[k] for k in ("run","draft_version","flag_version")},require_versions=True)
        self.assertEqual(self.values(),{})
        current=self.app.proposal()
        self.app.approve({k:current[k] for k in ("run","draft_version","flag_version")},require_versions=True)
        self.assertEqual(self.values(),{"A.xlsx":"original"})

if __name__=="__main__":unittest.main()
