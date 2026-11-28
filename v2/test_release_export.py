"""Export behavior acceptance. Native OOXML targets are structural evidence, not Excel UI certification."""
import io,unittest,zipfile
from xml.etree import ElementTree as ET
from v2.launcher import export_bytes, partition_plan
NS="{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL="{http://schemas.openxmlformats.org/package/2006/relationships}"
class ExportAcceptance(unittest.TestCase):
    def test_empty_workbook_has_one_valid_table_no_fabricated_record(self):
        fields=["Event_Key","Document No.","Source Workbook Path"]
        with zipfile.ZipFile(io.BytesIO(export_bytes([],fields,"xlsx"))) as z:
            self.assertIsNone(z.testzip())
            sheet=ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
            table=ET.fromstring(z.read("xl/tables/table1.xml"))
            self.assertEqual(table.attrib["name"],"NMDCIndex")
            self.assertEqual(table.attrib["ref"],"A1:C2")
            self.assertEqual([e.attrib["name"] for e in table.find(NS+"tableColumns")],fields)
            self.assertFalse(any(c.find(NS+"v") is not None for r in sheet.findall(".//"+NS+"row") if r.attrib["r"]=="2" for c in r.findall(NS+"c")))
            self.assertEqual(len(ET.fromstring(z.read("xl/workbook.xml")).find(NS+"sheets")),1)
    def test_document_and_source_links_count_and_preserve_targets(self):
        row={"Event_Key":"0001","Document Link":r"C:\Project\A B.pdf","Source Workbook Path":r"\\server\share\Register.xlsx","Original_Hyperlink_Target":r"..\A B.pdf"}
        fields=list(row)
        self.assertEqual(partition_plan([row,row],fields,link_limit=2),[(0,1),(1,2)])
        with zipfile.ZipFile(io.BytesIO(export_bytes([row],fields,"xlsx"))) as z:
            links=ET.fromstring(z.read("xl/worksheets/_rels/sheet1.xml.rels"))
            targets=[r.attrib["Target"] for r in links if r.attrib["Type"].endswith("/hyperlink")]
            self.assertEqual(len(targets),2)
            self.assertTrue(any("A%20B.pdf" in t or "A B.pdf" in t for t in targets))
            self.assertTrue(any("Register.xlsx" in t for t in targets))
            strings=z.read("xl/sharedStrings.xml").decode()
            self.assertIn(row["Original_Hyperlink_Target"],strings)
            self.assertNotIn("<f>",z.read("xl/worksheets/sheet1.xml").decode())
    def test_untrusted_scheme_is_not_active(self):
        row={"Document Link":"javascript:alert(1)","Source Workbook Path":"relative.xlsx","Event_Key":"X"}
        with zipfile.ZipFile(io.BytesIO(export_bytes([row],list(row),"xlsx"))) as z:
            self.assertNotIn("xl/worksheets/_rels/sheet1.xml.rels",z.namelist()) if "xl/worksheets/_rels/sheet1.xml.rels" not in z.namelist() else self.assertNotIn(b"/hyperlink",z.read("xl/worksheets/_rels/sheet1.xml.rels"))
            self.assertIn(b"javascript:alert(1)",z.read("xl/sharedStrings.xml"))
    def test_all_untrusted_formula_strings_remain_literal(self):
        row={"Document No.":"0001","Document Title":"{=1+1}","Event Status":"=HYPERLINK(\"https://example.invalid\",\"click\")","Revision":"-01"}
        with zipfile.ZipFile(io.BytesIO(export_bytes([row],list(row),"xlsx"))) as z:
            sheet=ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
            self.assertEqual(sheet.findall(".//"+NS+"f"),[])
            strings="".join(ET.fromstring(z.read("xl/sharedStrings.xml")).itertext())
            for value in row.values():self.assertIn(value,strings)
if __name__=="__main__":unittest.main()
