from __future__ import annotations

import base64
import io
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"


class ReleaseExcelContractTests(unittest.TestCase):
    def test_source_hyperlinks_are_native_row_specific_not_table_formulas(self):
        refresh = (ROOT / "excel" / "vba" / "modNMDC_Refresh.bas").read_text(encoding="utf-8")
        self.assertIn("values = column.DataBodyRange.Value2", refresh)
        self.assertIn("Set targetCell = column.DataBodyRange.Cells(rowIndex, 1)", refresh)
        self.assertIn("targetCell.Value2 = sourceText", refresh)
        self.assertIn("table.Parent.Hyperlinks.Add Anchor:=targetCell, Address:=addressText, TextToDisplay:=sourceText", refresh)
        self.assertIn("Application.AutoCorrect.AutoFillFormulasInLists = False", refresh)
        self.assertNotIn('formulas(rowIndex, 1) = "=HYPERLINK(', refresh)
        self.assertNotIn("column.DataBodyRange.Formula = formulas", refresh)

    def test_document_links_are_native_row_specific_not_table_formulas(self):
        refresh = (ROOT / "excel" / "vba" / "modNMDC_Refresh.bas").read_text(encoding="utf-8")
        self.assertIn('targetCell.Value2 = "Open document"', refresh)
        self.assertIn('table.Parent.Hyperlinks.Add Anchor:=targetCell, Address:=target, TextToDisplay:="Open document"', refresh)

    def test_home_setup_never_formats_all_excel_columns(self):
        setup = (ROOT / "packaging" / "Create_NMDC_Document_Index.vbs").read_text(encoding="utf-8")
        self.assertIn('ws.Range("A1:L33").Font.Name = "Aptos"', setup)
        self.assertNotIn('ws.Cells.Font.Name = "Aptos"', setup)
        self.assertNotIn('Columns("A:XFD")', setup)
        self.assertNotIn('Range("A:XFD")', setup)

    def test_base_home_has_no_formatted_tail_beyond_visible_dashboard(self):
        parts = sorted((ROOT / "excel" / "base_chunks").glob("NMDC_Document_Index_Base.xlsx.b64.part*"))
        self.assertTrue(parts)
        encoded = "".join(part.read_text(encoding="ascii") for part in parts)
        raw = base64.b64decode(encoded)
        with zipfile.ZipFile(io.BytesIO(raw)) as archive:
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            rels = {
                rel.attrib["Id"]: rel.attrib["Target"]
                for rel in rels_root.findall(f"{{{REL_PKG}}}Relationship")
            }
            home_path = None
            for sheet in workbook.findall(f".//{{{MAIN}}}sheet"):
                if sheet.attrib.get("name") == "Home":
                    rid = sheet.attrib[f"{{{REL_DOC}}}id"]
                    target = rels[rid]
                    home_path = target.lstrip("/") if target.startswith("/") else "xl/" + target.lstrip("/")
                    break
            self.assertIsNotNone(home_path)
            root = ET.fromstring(archive.read(home_path))
            cols = root.find(f"{{{MAIN}}}cols")
            if cols is not None:
                formatted_beyond_l = [
                    col.attrib for col in cols.findall(f"{{{MAIN}}}col")
                    if int(col.attrib.get("max", "0")) > 12 and ("style" in col.attrib or int(col.attrib.get("min", "0")) > 12)
                ]
                self.assertEqual(
                    formatted_beyond_l,
                    [],
                    "Base Home worksheet must not serialize formatting metadata into M:XFD",
                )


if __name__ == "__main__":
    unittest.main()
