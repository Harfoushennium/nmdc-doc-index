import csv
import io
import json
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

import nmdc_profiler as profiler

MAIN = profiler.MAIN_NS
REL_DOC = profiler.REL_DOC_NS
REL_PKG = profiler.REL_PKG_NS
DCT = profiler.DCTERMS_NS
CP = profiler.CP_NS


def col_letter(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def make_xlsx(path: Path, sheets, modified='2026-01-01T00:00:00Z', created='2025-01-01T00:00:00Z'):
    """Create a minimal OOXML workbook.

    sheets: list of dict(name, rows=[list[str]], merges=[], native_links=[], formula_links=[])
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    shared = []
    shared_idx = {}
    def idx(v):
        if v not in shared_idx:
            shared_idx[v] = len(shared)
            shared.append(v)
        return shared_idx[v]

    workbook = ET.Element(f'{{{MAIN}}}workbook', {f'xmlns:r': REL_DOC})
    ss = ET.SubElement(workbook, f'{{{MAIN}}}sheets')
    rels = ET.Element(f'{{{REL_PKG}}}Relationships')
    sheet_xmls = {}
    sheet_rels = {}
    for i, spec in enumerate(sheets, 1):
        ET.SubElement(ss, f'{{{MAIN}}}sheet', {
            'name': spec['name'], 'sheetId': str(i), f'{{{REL_DOC}}}id': f'rId{i}',
            **({'state': spec['state']} if spec.get('state') else {})
        })
        ET.SubElement(rels, f'{{{REL_PKG}}}Relationship', {
            'Id': f'rId{i}', 'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
            'Target': f'worksheets/sheet{i}.xml'
        })
        ws = ET.Element(f'{{{MAIN}}}worksheet', {f'xmlns:r': REL_DOC})
        max_cols = max((len(r) for r in spec.get('rows', [])), default=1)
        max_rows = max(len(spec.get('rows', [])), 1)
        ET.SubElement(ws, f'{{{MAIN}}}dimension', {'ref': f'A1:{col_letter(max_cols)}{max_rows}'})
        sd = ET.SubElement(ws, f'{{{MAIN}}}sheetData')
        for rn, rowvals in enumerate(spec.get('rows', []), 1):
            row = ET.SubElement(sd, f'{{{MAIN}}}row', {'r': str(rn)})
            for cn, val in enumerate(rowvals, 1):
                if val is None or val == '':
                    continue
                c = ET.SubElement(row, f'{{{MAIN}}}c', {'r': f'{col_letter(cn)}{rn}', 't': 's'})
                ET.SubElement(c, f'{{{MAIN}}}v').text = str(idx(str(val)))
        if spec.get('formula_links'):
            rn = max_rows + 1
            row = ET.SubElement(sd, f'{{{MAIN}}}row', {'r': str(rn)})
            for j, formula in enumerate(spec['formula_links'], 1):
                c = ET.SubElement(row, f'{{{MAIN}}}c', {'r': f'{col_letter(j)}{rn}'})
                ET.SubElement(c, f'{{{MAIN}}}f').text = formula
                ET.SubElement(c, f'{{{MAIN}}}v').text = '0'
        if spec.get('merges'):
            mc = ET.SubElement(ws, f'{{{MAIN}}}mergeCells', {'count': str(len(spec['merges']))})
            for ref in spec['merges']:
                ET.SubElement(mc, f'{{{MAIN}}}mergeCell', {'ref': ref})
        if spec.get('native_links'):
            hls = ET.SubElement(ws, f'{{{MAIN}}}hyperlinks')
            rr = ET.Element(f'{{{REL_PKG}}}Relationships')
            for j, (ref, target) in enumerate(spec['native_links'], 1):
                ET.SubElement(hls, f'{{{MAIN}}}hyperlink', {'ref': ref, f'{{{REL_DOC}}}id': f'rId{j}'})
                ET.SubElement(rr, f'{{{REL_PKG}}}Relationship', {
                    'Id': f'rId{j}', 'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
                    'Target': target, 'TargetMode': 'External'
                })
            sheet_rels[i] = ET.tostring(rr, encoding='utf-8', xml_declaration=True)
        sheet_xmls[i] = ET.tostring(ws, encoding='utf-8', xml_declaration=True)

    sst = ET.Element(f'{{{MAIN}}}sst', {'count': str(len(shared)), 'uniqueCount': str(len(shared))})
    for value in shared:
        si = ET.SubElement(sst, f'{{{MAIN}}}si')
        ET.SubElement(si, f'{{{MAIN}}}t').text = value

    core = ET.Element(f'{{{CP}}}coreProperties', {
        'xmlns:dcterms': DCT,
        'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
        'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance'
    })
    ET.SubElement(core, f'{{{DCT}}}created').text = created
    ET.SubElement(core, f'{{{DCT}}}modified').text = modified

    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('xl/workbook.xml', ET.tostring(workbook, encoding='utf-8', xml_declaration=True))
        z.writestr('xl/_rels/workbook.xml.rels', ET.tostring(rels, encoding='utf-8', xml_declaration=True))
        z.writestr('xl/sharedStrings.xml', ET.tostring(sst, encoding='utf-8', xml_declaration=True))
        z.writestr('docProps/core.xml', ET.tostring(core, encoding='utf-8', xml_declaration=True))
        for i, xml in sheet_xmls.items():
            z.writestr(f'xl/worksheets/sheet{i}.xml', xml)
            if i in sheet_rels:
                z.writestr(f'xl/worksheets/_rels/sheet{i}.xml.rels', sheet_rels[i])


class ProfilerTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.data = self.root / 'DATA'
        self.out = self.root / 'outputs' / 'cycle1'
        self.rules = Path(__file__).resolve().parents[1] / 'config' / 'classification_rules.csv'

    def tearDown(self):
        self.tmp.cleanup()

    def test_01_recursive_discovery_both_families(self):
        make_xlsx(self.data/'METHODS'/'nested'/'2369 Deliverables.xlsx', [{'name':'Procedures','rows':[['PROJECT NO','2369']]}])
        make_xlsx(self.data/'TECH'/'2800 DOCUMENT REGISTER.xlsx', [{'name':'Drawings','rows':[['PROJECT NO','2800']]}])
        result = profiler.run(self.data, self.out, self.rules)
        self.assertEqual(2, len(result['workbooks']))
        self.assertEqual({'METHODS','TECH'}, {w['source_family'] for w in result['workbooks']})

    def test_02_source_family_preserved(self):
        p = self.data/'METHODS'/'2369 Deliverables.xlsx'
        make_xlsx(p, [{'name':'Procedures','rows':[['X']]}])
        wb = profiler.profile_workbook(p, self.root, self.data)
        self.assertEqual('METHODS', wb['source_family'])

    def test_03_exact_hash_duplicate_detection(self):
        p1 = self.data/'TECH'/'3291 DOCUMENT REGISTER.xlsx'
        p2 = self.data/'TECH'/'3291 DOCUMENT REGISTER revised.xlsx'
        make_xlsx(p1, [{'name':'Drawings','rows':[['PROJECT NO','3291']]}])
        shutil.copy2(p1, p2)
        wbs = [profiler.profile_workbook(p, self.root, self.data) for p in [p1,p2]]
        exact, _ = profiler.assign_selection(wbs, profiler.load_rules(self.rules))
        self.assertEqual(1, len(exact))
        self.assertEqual(2, len(exact[0]['files']))

    def test_04_modified_date_from_core_not_filesystem_mtime(self):
        p = self.data/'TECH'/'2820 DOCUMENT REGISTER.xlsx'
        make_xlsx(p, [{'name':'Drawings','rows':[['PROJECT NO','2820']]}], modified='2024-05-08T10:00:00Z')
        p.touch()
        wb = profiler.profile_workbook(p, self.root, self.data)
        self.assertEqual('2024-05-08T10:00:00Z', wb['doc_modified'])
        self.assertTrue(wb['timestamp_reliable'])
        self.assertIn('dcterms:modified', wb['timestamp_source'])

    def test_05_newest_source_selected_in_confirmed_group(self):
        p1 = self.data/'TECH'/'2820-DOCUMENT REGISTER.xlsx'
        p2 = self.data/'TECH'/'2820-DOCUMENT REGISTER-NEW.xlsx'
        make_xlsx(p1, [{'name':'Drawings','rows':[['PROJECT NO','2820']]}], modified='2025-05-08T00:00:00Z')
        make_xlsx(p2, [{'name':'Drawings','rows':[['PROJECT NO','2820']]}], modified='2026-09-08T00:00:00Z')
        wbs = [profiler.profile_workbook(p, self.root, self.data) for p in [p1,p2]]
        _, groups = profiler.assign_selection(wbs, profiler.load_rules(self.rules))
        self.assertEqual(1, len(groups))
        chosen = [w for w in wbs if w['selected_excluded_status']=='SELECTED']
        self.assertEqual('2820-DOCUMENT REGISTER-NEW.xlsx', chosen[0]['filename'])
        older = [w for w in wbs if w['selected_excluded_status']=='SUPERSEDED'][0]
        self.assertEqual(chosen[0]['relative_path'], older['selected_replacement_file'])

    def test_06_same_project_different_register_types_not_conflated(self):
        p1 = self.data/'TECH'/'2035 DOCUMENT REGISTER.xlsx'
        p2 = self.data/'TECH'/'2035 Installation Aids Register Sep 2024.xlsx'
        make_xlsx(p1, [{'name':'Drawings','rows':[['PROJECT NO','2035']]}])
        make_xlsx(p2, [{'name':'Sheet1','rows':[['PROJECT NO','2035']]}])
        wbs = [profiler.profile_workbook(p, self.root, self.data) for p in [p1,p2]]
        _, groups = profiler.assign_selection(wbs, profiler.load_rules(self.rules))
        self.assertEqual([], groups)
        self.assertEqual('SELECTED', wbs[0]['selected_excluded_status'])
        self.assertEqual('EXCLUDED', wbs[1]['selected_excluded_status'])

    def test_07_encrypted_unreadable_reporting(self):
        p = self.data/'METHODS'/'2722 Deliverable.xlsx'
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1' + b'x'*20)
        wb = profiler.profile_workbook(p, self.root, self.data)
        self.assertEqual('ENCRYPTED', wb['readability_status'])
        self.assertEqual('OLE2_CFB_ENCRYPTED', wb['unreadable_reason'])
        self.assertTrue(wb['password_retry_possible'])

    def test_08_section_classification_incoming_doc_and_drg(self):
        p = self.data/'METHODS'/'2891 Deliverables.xlsx'
        make_xlsx(p, [{'name':'Incomming DOC and DRG','rows':[
            ['NMDC PROJECT NO.','2891'], ['DRAWINGS'], ['A'], ['DOCUMENTS'], ['B']
        ]}])
        result = profiler.run(self.data, self.out, self.rules)
        rows = result['classification_rows']
        sections = {(r['original_section'], r['category'], r['subcategory']) for r in rows}
        self.assertIn(('DRAWINGS','DRAWING','INCOMING DRAWING'), sections)
        self.assertIn(('DOCUMENTS','DOCUMENT','INCOMING TECHNICAL DOCUMENT'), sections)

    def test_09_ambiguous_classification_review_required(self):
        p = self.data/'TECH'/'3000 DOCUMENT REGISTER.xlsx'
        make_xlsx(p, [{'name':'Mystery','rows':[['NMDC PROJECT NO.','3000'], ['UNKNOWN THING']]}])
        result = profiler.run(self.data, self.out, self.rules)
        row = result['classification_rows'][0]
        self.assertEqual('REVIEW_REQUIRED', row['proposed_action'])
        self.assertIn('REVIEW_REQUIRED', row['warning_code'])

    def test_10_merge_profiling(self):
        p = self.data/'TECH'/'3000 DOCUMENT REGISTER.xlsx'
        make_xlsx(p, [{'name':'Drawings','rows':[['NMDC PROJECT NO.','3000']], 'merges':['A1:B1','C2:D3']}])
        wb = profiler.profile_workbook(p, self.root, self.data)
        s = wb['sheets'][0]
        self.assertEqual(2, s['merge_count'])
        self.assertEqual(['A1:B1','C2:D3'], s['representative_merge_ranges'])

    def test_11_native_and_formula_hyperlinks(self):
        p = self.data/'TECH'/'3000 DOCUMENT REGISTER.xlsx'
        make_xlsx(p, [{'name':'Drawings','rows':[['NMDC PROJECT NO.','3000']],
                       'native_links':[('A1','https://example.com')],
                       'formula_links':['HYPERLINK("https://example.org","x")']}])
        wb = profiler.profile_workbook(p, self.root, self.data)
        s = wb['sheets'][0]
        self.assertEqual(1, s['native_hyperlink_count'])
        self.assertEqual(1, s['hyperlink_formula_count'])
        self.assertEqual('https://example.com', s['native_hyperlinks'][0]['target'])

    def test_12_project_mismatch_uses_internal_project_label_evidence(self):
        p = self.data/'TECH'/'2035 DOCUMENT REGISTER.xlsx'
        make_xlsx(p, [{'name':'Specification','rows':[['NMDC PROJECT NO.','2136'], ['Specification']]}])
        wb = profiler.profile_workbook(p, self.root, self.data)
        self.assertEqual('PROJECT_MISMATCH', wb['project_mismatch_findings'][0]['code'])
        self.assertEqual(['2035'], wb['project_mismatch_findings'][0]['path_projects'])
        self.assertEqual(['2136'], wb['project_mismatch_findings'][0]['internal_projects'])

    def test_13_explicit_file_exclusion(self):
        p = self.data/'METHODS'/'1. Delivarables FORMAT.xlsx'
        make_xlsx(p, [{'name':'Procedures','rows':[['X']]}])
        wb = profiler.profile_workbook(p, self.root, self.data)
        profiler.assign_selection([wb], profiler.load_rules(self.rules))
        self.assertEqual('EXCLUDED', wb['selected_excluded_status'])
        self.assertIn('Template', wb['selection_exclusion_reason'])

    def test_14_deterministic_output_ordering(self):
        make_xlsx(self.data/'TECH'/'4000 DOCUMENT REGISTER.xlsx', [{'name':'Drawings','rows':[['NMDC PROJECT NO.','4000']]}])
        make_xlsx(self.data/'METHODS'/'3000 Deliverables.xlsx', [{'name':'Procedures','rows':[['NMDC PROJECT NO.','3000']]}])
        profiler.run(self.data, self.out, self.rules)
        first = {p.name:p.read_bytes() for p in self.out.iterdir()}
        profiler.run(self.data, self.out, self.rules)
        second = {p.name:p.read_bytes() for p in self.out.iterdir()}
        self.assertEqual(first, second)

    def test_15_v1_mappings(self):
        rules = profiler.load_rules(self.rules)
        cases = [
            ('TECH', {'WORKSHEET':'Documents - Naval & Marine'}, ('NAVAL & MARINE','DOCUMENT','GENERAL TECHNICAL DOCUMENT')),
            ('TECH', {'WORKSHEET':'TN-NA'}, ('NAVAL & MARINE','DOCUMENT','TECHNICAL NOTE')),
            ('TECH', {'WORKSHEET':'Drawings'}, ('GENERAL / MULTIDISCIPLINE','DRAWING','ENGINEERING DRAWING')),
            ('TECH', {'WORKSHEET':'Sketches'}, ('GENERAL / MULTIDISCIPLINE','SKETCH','ENGINEERING SKETCH')),
            ('TECH', {'WORKSHEET':'DP'}, ('NAVAL & MARINE','DRAWING','DP SETUP PLAN')),
            ('TECH', {'WORKSHEET':'COMMISSION - List of OTP'}, ('COMMISSIONING','PROCEDURE','OPERATIONAL TEST PROCEDURE')),
            ('METHODS', {'WORKSHEET':'Setup Plans & Anchor Patterns'}, ('MARINE OPERATIONS','DRAWING','METHOD DRAWING')),
        ]
        for fam, evidence, expected in cases:
            c = profiler.apply_classification(rules, fam, evidence)
            self.assertEqual(expected, (c['discipline'],c['category'],c['subcategory']), (fam,evidence,c))

    def test_16_title_and_doc_number_refinement(self):
        rules = profiler.load_rules(self.rules)
        c = profiler.apply_classification(rules, 'TECH', {
            'WORKSHEET':'Documents - Pipeline & Cable',
            'DOC_NUMBER':'ABC-TN-PL-001',
            'TITLE':'Pipeline free span analysis report'
        })
        self.assertEqual('PIPELINE & CABLE', c['discipline'])
        self.assertIn('DOC_NUMBER:T043', c['match_basis'])
        self.assertTrue(any(x.startswith('TITLE:') for x in c['match_basis']))

    def test_17_inventory_mandatory_fields_and_relative_paths(self):
        make_xlsx(self.data/'TECH'/'2820 DOCUMENT REGISTER.xlsx', [{'name':'Drawings','rows':[['NMDC PROJECT NO.','2820']]}])
        profiler.run(self.data, self.out, self.rules)
        with (self.out/'source_inventory.csv').open() as f:
            row = next(csv.DictReader(f))
        required = {
            'relative_path','file_extension','doc_created','doc_modified','timestamp_source','timestamp_reliable',
            'readability_status','inferred_project_numbers','logical_register_identity','duplicate_version_group_id',
            'selected_excluded_status','selection_exclusion_reason','selected_replacement_file','warning_codes'
        }
        self.assertTrue(required.issubset(row.keys()))
        self.assertFalse(Path(row['relative_path']).is_absolute())
        self.assertNotIn(':\\', row['relative_path'])

    def test_18_sheet_state_and_relationship_mapping(self):
        p = self.data/'TECH'/'2820 DOCUMENT REGISTER.xlsx'
        make_xlsx(p, [
            {'name':'Second Visible','rows':[['NMDC PROJECT NO.','2820']]},
            {'name':'Hidden Sheet','state':'hidden','rows':[['X']]},
        ])
        wb = profiler.profile_workbook(p, self.root, self.data)
        self.assertEqual(['Second Visible','Hidden Sheet'], [s['sheet_name'] for s in wb['sheets']])
        self.assertEqual('hidden', wb['sheets'][1]['state'])

    def test_19_no_committed_report_sha_dependency(self):
        make_xlsx(self.data/'TECH'/'2820 DOCUMENT REGISTER.xlsx', [{'name':'Drawings','rows':[['NMDC PROJECT NO.','2820']]}])
        profiler.run(self.data, self.out, self.rules)
        self.assertFalse((self.out/'hermes_report.md').exists())
        self.assertFalse((self.out/'hermes_report_final.md').exists())

    def test_20_rules_are_external_and_editable(self):
        rules = profiler.load_rules(self.rules)
        self.assertGreater(len(rules), 20)
        self.assertTrue(any(r.rule_id == 'T041' and r.scope == 'WORKSHEET' for r in rules))


if __name__ == '__main__':
    unittest.main()
