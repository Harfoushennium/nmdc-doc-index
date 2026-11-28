"""Export behavior acceptance. Native OOXML targets are structural evidence, not Excel UI certification."""
import io,unittest,zipfile
from xml.etree import ElementTree as ET
from v2.launcher import export_bytes, partition_plan
NS="{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL="{http://schemas.openxmlformats.org/package/2006/relationships}"
class ExportAcceptance(unittest.TestCase):
    def test_excel_date_epoch_preserves_pre1900_literals_and_calendar_boundaries(self):
        rows=[{'Event Date':value} for value in ('1899-12-31','1900-01-01','1900-02-28','1900-03-01','2026-11-28')]
        with zipfile.ZipFile(io.BytesIO(export_bytes(rows,['Event Date'],'xlsx'))) as archive:
            cells=ET.fromstring(archive.read('xl/worksheets/sheet1.xml')).findall('.//'+NS+'sheetData/'+NS+'row')[1:]
            first=cells[0].find(NS+'c');self.assertEqual(first.get('t'),'s')
            self.assertIn('1899-12-31',''.join(ET.fromstring(archive.read('xl/sharedStrings.xml')).itertext()))
            self.assertEqual([float(row.find(NS+'c').find(NS+'v').text) for row in cells[1:4]],[1,59,61])
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

    def test_partition_export_manifest_includes_ordered_keys_ranges_hashes_links_and_fields(self):
        import hashlib, json
        from v2.launcher import partition_export
        rows = [
            {"Event_Key": "E3", "Document Link": r"C:\Project\3.pdf", "Source Workbook Path": r"\\srv\share\3.xlsx"},
            {"Event_Key": "E1", "Document Link": r"C:\Project\1.pdf", "Source Workbook Path": r"\\srv\share\1.xlsx"},
            {"Event_Key": "E2", "Document Link": r"C:\Project\2.pdf", "Source Workbook Path": r"\\srv\share\2.xlsx"},
        ]
        fields = ["Event_Key", "Document Link", "Source Workbook Path"]
        archive = partition_export(rows, fields, link_limit=4)
        with zipfile.ZipFile(io.BytesIO(archive)) as z:
            self.assertIsNone(z.testzip())
            manifest = json.loads(z.read("manifest.json").decode("utf-8"))
            self.assertEqual(manifest["rows"], 3)
            self.assertEqual(manifest["fields"], fields)
            expected_keys_hash = hashlib.sha256(json.dumps(["E3", "E1", "E2"], ensure_ascii=False).encode("utf-8")).hexdigest()
            sorted_keys_hash = hashlib.sha256(json.dumps(["E1", "E2", "E3"], ensure_ascii=False).encode("utf-8")).hexdigest()
            self.assertEqual(manifest["keys_sha256"], expected_keys_hash)
            self.assertNotEqual(manifest["keys_sha256"], sorted_keys_hash)
            self.assertEqual(len(manifest["parts"]), 2)
            p1, p2 = manifest["parts"]
            self.assertEqual((p1["start_row"], p1["end_row"], p1["rows"], p1["native_links"]), (1, 2, 2, 4))
            self.assertEqual((p2["start_row"], p2["end_row"], p2["rows"], p2["native_links"]), (3, 3, 1, 2))
            for part in (p1, p2):
                raw = z.read(part["file"])
                self.assertEqual(part["bytes"], len(raw))
                self.assertEqual(part["sha256"], hashlib.sha256(raw).hexdigest())
if __name__=="__main__":unittest.main()
