"""Research: compare exact structural worksheet evidence, including conversion overhead.
No workbook contents are emitted. This is not a full extraction/approval equivalence test.
"""
import argparse, ctypes, hashlib, io, json, statistics, struct, subprocess, tempfile, time, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
from nmdc_profiler.ooxml import workbook_sheet_map
NS="{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL="{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
def reference(xml):
    r=ET.fromstring(xml); dim=r.find(".//"+NS+"dimension")
    rows=[]
    for row in r.findall(".//"+NS+"row"):
        cells=[]
        for c in row.findall(NS+"c"):
            cells.append([c.get("r",""),c.get("t",""),c.findtext(NS+"v",""),c.findtext(NS+"f",""),"".join(t.text or "" for t in c.iter(NS+"t"))])
        rows.append([row.get("r",""),cells])
    return [dim.get("ref","") if dim is not None else "",
        str(sum(r.get("hidden") in ("1","true","TRUE") for r in r.findall(".//"+NS+"row"))),
        str(sum(c.get("hidden") in ("1","true","TRUE") for c in r.findall(".//"+NS+"col"))),
        rows,[m.get("ref","") for m in r.findall(".//"+NS+"mergeCell")],
        [[h.get("ref",""),h.get(REL,""),h.get("location","")] for h in r.findall(".//"+NS+"hyperlink")]]
def decode(raw):
    f=io.BytesIO(raw)
    def field():
        n=struct.unpack("<I",f.read(4))[0];return f.read(n).decode("utf-8")
    assert field()=="NMDC-SCAN-1"
    dim,hr,hc=field(),field(),field()
    rows=[[field(),[[field() for _ in range(5)] for _ in range(int(field()))]] for _ in range(int(field()))]
    merges=[field() for _ in range(int(field()))]
    links=[[field() for _ in range(3)] for _ in range(int(field()))]
    assert not f.read()
    return [dim,hr,hc,rows,merges,links]
def native(xml,exe,root):
    src=root/"sheet.xml";out=root/"sheet.bin";src.write_bytes(xml)
    subprocess.run([str(exe),str(src),str(out)],check=True,capture_output=True,timeout=60)
    return decode(out.read_bytes())
def shared_scanner(path):
    lib=ctypes.CDLL(str(path.resolve()))
    lib.nmdc_scan_xml.argtypes=[ctypes.c_char_p,ctypes.c_size_t,ctypes.POINTER(ctypes.c_void_p),ctypes.POINTER(ctypes.c_size_t)]
    lib.nmdc_scan_xml.restype=ctypes.c_int
    lib.nmdc_scan_free.argtypes=[ctypes.c_void_p]
    lib.nmdc_scan_free.restype=None
    def scan(xml):
        out=ctypes.c_void_p();size=ctypes.c_size_t()
        rc=lib.nmdc_scan_xml(xml,len(xml),ctypes.byref(out),ctypes.byref(size))
        if rc:raise ValueError("Native shared parser rejected XML")
        try:return decode(ctypes.string_at(out,size.value))
        finally:lib.nmdc_scan_free(out)
    return scan
def main():
    p=argparse.ArgumentParser();p.add_argument("exe",type=Path);p.add_argument("--data",type=Path,default=Path("DATA"));p.add_argument("--library",type=Path,required=True);a=p.parse_args()
    exe=a.exe.resolve();data=a.data.resolve();files=sorted(p for p in data.rglob("*") if p.suffix.lower() in (".xlsx",".xlsm") and not p.name.startswith("~$"))
    before={p:hashlib.sha256(p.read_bytes()).hexdigest() for p in files}
    scan=shared_scanner(a.library)
    totals={"python":[],"native":[],"shared":[]};checked=0;unsupported=0;cells=0;xml_bytes=0;failures=0
    with tempfile.TemporaryDirectory() as d:
        root=Path(d)
        for f in files:
            if not zipfile.is_zipfile(f):unsupported+=1;continue
            with zipfile.ZipFile(f) as z:
                for sheet in workbook_sheet_map(z):
                    if sheet["xml_path"] not in z.namelist():unsupported+=1;continue
                    xml=z.read(sheet["xml_path"])
                    start=time.perf_counter();expected=reference(xml);pt=time.perf_counter()-start
                    start=time.perf_counter()
                    try:actual=native(xml,exe,root)
                    except (ValueError,AssertionError,subprocess.SubprocessError,struct.error):failures+=1;continue
                    nt=time.perf_counter()-start
                    if actual!=expected:failures+=1;continue
                    start=time.perf_counter()
                    try:direct=scan(xml)
                    except (ValueError,AssertionError,struct.error):failures+=1;continue
                    st=time.perf_counter()-start
                    if direct!=expected:failures+=1;continue
                    checked+=1;cells+=sum(len(r[1]) for r in expected[3]);xml_bytes+=len(xml)
                    totals["python"].append(pt);totals["native"].append(nt);totals["shared"].append(st)
        # Reject malformed XML and DTDs. No native output is trusted on failed parses.
        for xml in (b"<broken>",b"<!DOCTYPE x [<!ENTITY a 'b'>]><x>&a;</x>"):
            (root/"bad.xml").write_bytes(xml)
            r=subprocess.run([str(exe),str(root/"bad.xml"),str(root/"bad.bin")],capture_output=True,timeout=10)
            assert r.returncode!=0
            try:scan(xml)
            except ValueError:pass
            else:raise AssertionError("Shared scanner accepted invalid input")
    report={"scope":"real repository XML structural scan only; no data records emitted","workbooks":len(files),"verified_worksheets":checked,"unsupported":unsupported,"mismatches_or_failures":failures,"verified_cells":cells,"xml_bytes":xml_bytes,"python_total_seconds":round(sum(totals["python"]),4),"native_end_to_end_seconds":round(sum(totals["native"]),4),"native_includes_temp_file_io_process_and_decoding":True,"source_hashes_unchanged":all(hashlib.sha256(p.read_bytes()).hexdigest()==h for p,h in before.items()),"malformed_and_dtd_rejected":True}
    report["shared_end_to_end_seconds"]=round(sum(totals["shared"]),4)
    if checked and sum(totals["shared"]):report["shared_speed_ratio"]=round(sum(totals["python"])/sum(totals["shared"]),3)
    if checked and sum(totals["native"]):report["speed_ratio"]=round(sum(totals["python"])/sum(totals["native"]),3)
    Path("native-parser-metrics.json").write_text(json.dumps(report,indent=2),encoding="utf-8");print(json.dumps(report,indent=2))
    assert checked and not failures and report["source_hashes_unchanged"]
if __name__=="__main__":main()
