"""Public custom field behavior from the V1 CustomFields/KeywordMappings contract."""
import unittest,tempfile
from pathlib import Path
from v2.launcher import App
from nmdc_profiler.update_engine import stage_update
class CustomFieldAcceptance(unittest.TestCase):
    def test_custom_fields_are_derived_and_persistent(self):
        with tempfile.TemporaryDirectory() as d:
            root=Path(d);data=root/"DATA";data.mkdir();(data/"fixture.xlsx").write_text("synthetic")
            app=App(root/"workspace")
            try:
                stage_update(data_dir=data,state_dir=app.state,processor=lambda p,r:[{"Event_Key":"E","Global_Document_Key":"D","Document No.":"0001","Project No.":"P","Revision":"A","Source File":r,"Document Title":"SAFEEN 3000; DLB-123; offshore","Source Sheet":"Register","Source Row":2}])
                p=app.proposal();app.save_choices(p);app.approve({"run":p["run"]})
                fields=[{"Enabled?":True,"Field Name":"Vessels","Search In":"Document Title","Match Behavior":"ALL UNIQUE","Separator":" / ","Notes":""}]
                mappings=[
                    {"Enabled?":True,"Field Name":"Vessels","Keyword / Pattern":"SAFEEN+3000 -ONSHORE","Result":"SAFEEN-3000","Match Type":"ALL TERMS","Priority":"10","Notes":""},
                    {"Enabled?":True,"Field Name":"Vessels","Keyword / Pattern":"DLB-???","Result":"DLB-???","Match Type":"WILDCARD","Priority":"20","Notes":""},
                    {"Enabled?":True,"Field Name":"Vessels","Keyword / Pattern":"SAFEEN","Result":"SAFEEN-3000","Match Type":"CONTAINS","Priority":"30","Notes":""}]
                app.save_custom_fields({"fields":fields,"mappings":mappings})
                self.assertEqual(app.view("documents")[0]["Vessels"],"SAFEEN-3000 / DLB-123")
                self.assertNotIn("Vessels",app.approved()["records"][0])
                fields[0]["Field Name"]="Document No."
                with self.assertRaises(ValueError):app.save_custom_fields({"fields":fields,"mappings":mappings})
                self.assertEqual(app.view("documents")[0]["Vessels"],"SAFEEN-3000 / DLB-123")
                app.invalidate()
            finally:app.lock.close()
            app=App(root/"workspace")
            try:self.assertEqual(app.view("documents")[0]["Vessels"],"SAFEEN-3000 / DLB-123")
            finally:app.invalidate();app.lock.close()
    def test_classification_save_reject_and_restore(self):
        import copy
        with tempfile.TemporaryDirectory() as d:
            app=App(Path(d)/"workspace")
            try:
                original=app.rules()
                edited=copy.deepcopy(original)
                edited["rows"][0]["Notes"]="Release acceptance fixture"
                app.save_rules(edited)
                self.assertEqual(app.rules()["rows"][0]["Notes"],"Release acceptance fixture")
                invalid=copy.deepcopy(edited)
                enabled=next(r for r in invalid["rows"] if r["Enabled"]=="YES")
                enabled["Match_Type"]="INVALID"
                with self.assertRaises(ValueError):app.save_rules(invalid)
                self.assertEqual(app.rules()["rows"],edited["rows"])
                versions=app.configuration_history()
                v=next(x for x in versions if x["kind"]=="classification_rules")
                app.restore_configuration({"id":v["id"],"confirm":True})
                self.assertEqual(app.rules(),original)
                self.assertTrue(app.settings["requires_rescan"])
            finally:app.invalidate();app.lock.close()
if __name__=="__main__":unittest.main()
