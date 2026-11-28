"""Public custom field behavior from the V1 CustomFields/KeywordMappings contract."""
import unittest,tempfile
from pathlib import Path
from v2.launcher import App
from nmdc_profiler.update_engine import stage_update
from v2 import custom_fields

def field(name="Derived", search="Document Title", behavior="ALL UNIQUE", enabled=True):
    return {"Enabled?":enabled,"Field Name":name,"Search In":search,"Match Behavior":behavior,"Separator":" / ","Notes":""}
def mapping(pattern,result,priority="10",kind="CONTAINS",enabled=True,name="Derived"):
    return {"Enabled?":enabled,"Field Name":name,"Keyword / Pattern":pattern,"Result":result,"Match Type":kind,"Priority":priority,"Notes":""}

class CustomFieldAcceptance(unittest.TestCase):
    def test_enabled_and_top_level_schema_are_strict(self):
        for value in ("maybe", "enabled", 2):
            with self.subTest(value=value), self.assertRaises(ValueError):
                custom_fields.validate({"fields":[field(enabled=value)],"mappings":[]},[])
        with self.assertRaises(ValueError):custom_fields.validate({"fields":[],"mappings":[],"enabled":True},[])

    def test_vba_like_supported_wildcards_priority_first_unique_and_stable_ties(self):
        rows=[{"Document Title":"prefix ABC-123 suffix Alpha alpha BETA"}]
        payload=custom_fields.validate({"fields":[field()],"mappings":[
            mapping("ABC-???","Code ???","20","WILDCARD"),
            mapping("Alpha","first","10"),mapping("alpha","duplicate","10"),
            mapping("BETA","last","30")]},[])
        custom_fields.enrich(rows,payload)
        self.assertEqual(rows[0]["Derived"],"first / duplicate / Code 123 / last")
        rows=[{"Document Title":"prefix ABC-123 suffix Alpha"}]
        payload["fields"][0]["Match Behavior"]="FIRST"
        custom_fields.enrich(rows,payload)
        self.assertEqual(rows[0]["Derived"],"first")
        for pattern in ("a[#]", "a[!b]", "x#"):
            with self.subTest(pattern=pattern), self.assertRaises(ValueError):
                custom_fields.validate({"fields":[field()],"mappings":[mapping(pattern,"x",kind="WILDCARD")]},[])

    def test_derived_search_dependencies_require_v1_table_order(self):
        payload={"fields":[field("First"),field("Second","First")],"mappings":[mapping("alpha","A",name="First"),mapping("A","B",name="Second")]}
        rows=[{"Document Title":"alpha"}]
        custom_fields.enrich(rows,custom_fields.validate(payload,[]))
        self.assertEqual((rows[0]["First"],rows[0]["Second"]),("A","B"))
        payload["fields"]=list(reversed(payload["fields"]))
        with self.assertRaisesRegex(ValueError,"prior field"):
            custom_fields.validate(payload,[])

    def test_search_columns_match_document_view_and_malformed_api_save_is_atomic(self):
        from nmdc_profiler.excel_bridge import build_document_rows
        from nmdc_profiler.extractor import OUTPUT_FIELDS
        event={k:"" for k in OUTPUT_FIELDS};event.update({"Project No.":"P","Document No.":"0001","Document Title":"alpha","Revision":"A","Source File":"a.xlsx","Source Sheet":"S","Source Row":2})
        actual=set(build_document_rows([event])[0])
        # App.view("documents") supplies these two after aggregation and before enrichment.
        actual.update({"Original_Hyperlink_Target","Source Workbook Path"})
        self.assertTrue(set(custom_fields.DOC_COLUMNS)<=actual)
        self.assertIn("Revision",OUTPUT_FIELDS)
        with tempfile.TemporaryDirectory() as d:
            app=App(Path(d)/"workspace")
            try:
                valid={"fields":[field()],"mappings":[mapping("alpha","A")]}
                app.save_custom_fields(valid)
                before=(app.config/"custom_fields.json").read_bytes()
                invalid={"fields":[field(enabled="perhaps")],"mappings":[]}
                with self.assertRaises(ValueError):app.save_custom_fields(invalid)
                self.assertEqual((app.config/"custom_fields.json").read_bytes(),before)
            finally:app.invalidate();app.lock.close()
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
