"""Read-only V1 migration preview; no import or active-state mutation."""
from __future__ import annotations
import csv, hashlib, json, re, posixpath
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from nmdc_profiler.ooxml import workbook_sheet_map,load_shared_strings,cell_value
from nmdc_profiler.excel_bridge import load_approved_state,load_latest_stage
from v2.recovery import digest
from v2 import custom_fields
from nmdc_profiler.extractor import OUTPUT_FIELDS

NS="{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL="{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG="{http://schemas.openxmlformats.org/package/2006/relationships}"
TABLES={"Configuration","ClassificationRules","CustomFields","KeywordMappings","FileReview","ReviewFlags"}
RUNTIME_AREAS=("approved","staging","cache","logs","user_flags")
CONFIGS=("classification_rules.csv","project_identity_overrides.csv","source_exclusions.csv")
def overlap(a,b):return a==b or a in b.parents or b in a.parents
def children(root,areas):
    for area in areas:
        path=root/area
        if path.is_symlink():raise ValueError("Symlink in migration input: "+str(path))
        if path.is_file():yield path
        elif path.is_dir():
            for p in sorted(path.rglob("*")):
                if p.is_symlink():raise ValueError("Symlink in migration input: "+str(p))
                if p.is_file():
                    if not p.resolve().is_relative_to(root):raise ValueError("Migration input escapes package")
                    yield p
def csv_rows(path):
    with path.open(encoding="utf-8-sig",newline="") as f:
        r=csv.DictReader(f)
        if not r.fieldnames or len(set(r.fieldnames))!=len(r.fieldnames):raise ValueError("Invalid CSV schema: "+path.name)
        rows=list(r)
        if any(None in row or None in row.values() for row in rows):raise ValueError("Invalid CSV row: "+path.name)
        return rows
def coord(ref):
    m=re.fullmatch(r"([A-Z]+)([1-9][0-9]*)",ref)
    if not m:raise ValueError("Invalid cell range")
    col=0
    for ch in m[1]:col=col*26+ord(ch)-64
    return int(m[2]),col
def workbook_tables(path):
    out={}; warnings=[]
    with ZipFile(path) as z:
        infos=z.infolist()
        if len(infos)!=len({i.filename for i in infos}) or sum(i.file_size for i in infos)>512*1024**2:raise ValueError("Unsafe or oversized workbook archive")
        names=set(z.namelist());shared=load_shared_strings(z)
        for sheet in workbook_sheet_map(z):
            sp=sheet["xml_path"];rp=posixpath.join(posixpath.dirname(sp),"_rels",posixpath.basename(sp)+".rels")
            if rp not in names:continue
            rels={r.get("Id"):r.get("Target") for r in ET.fromstring(z.read(rp)).findall(PKG+"Relationship") if r.get("Type","").endswith("/table") and r.get("TargetMode")!="External"}
            if not rels:continue
            tree=ET.fromstring(z.read(sp));ranges=[]
            for tp in tree.findall(".//"+NS+"tablePart"):
                t=rels.get(tp.get(REL+"id"))
                if not t:continue
                target=posixpath.normpath(t.lstrip("/") if t.startswith("/") else posixpath.join(posixpath.dirname(sp),t))
                if not target.startswith("xl/tables/") or target not in names:raise ValueError("Invalid workbook table relationship")
                table=ET.fromstring(z.read(target));name=table.get("name",table.get("displayName",""))
                if name not in TABLES:continue
                if name in out:raise ValueError("Duplicate workbook table: "+name)
                start,end=table.get("ref","").split(":");r1,c1=coord(start);r2,c2=coord(end)
                if not 0<r2-r1<=100000 or not 0<c2-c1+1<=100:raise ValueError("Unsupported legacy table bounds")
                headers=[e.get("name","") for e in table.findall(".//"+NS+"tableColumn")]
                if len(headers)!=c2-c1+1 or len(set(headers))!=len(headers):raise ValueError("Invalid legacy table headers")
                cells={}
                for cell in tree.findall(".//"+NS+"c"):
                    r,c=coord(cell.get("r",""))
                    if r1<r<=r2 and c1<=c<=c2:
                        if cell.find(NS+"f") is not None:warnings.append("Formula in legacy table "+name+" at "+cell.get("r","")+"; cached value only")
                        cells[r,c]=cell_value(cell,shared)
                rows=[{h:cells.get((r,c1+i),"") for i,h in enumerate(headers)} for r in range(r1+1,r2+1-int(table.get("totalsRowCount","0")))]
                out[name]={"sheet":sheet["name"],"columns":headers,"rows":[r for r in rows if any(r.values())]}
    return out,warnings
def preview_package(package):
    package=Path(package).resolve()
    runtime=package/"runtime";config=package/"config"
    if not runtime.is_dir() or not config.is_dir():raise ValueError("Select the legacy package containing runtime and config folders.")
    blockers=[];warnings=[];files={}
    for p in children(runtime,(*RUNTIME_AREAS,"owner_review_drafts.csv","owner_review_drafts.csv.tmp","owner_review_drafts.csv.bak","user_custom_fields.csv","user_keyword_mappings.csv")):
        files[p.relative_to(package).as_posix()]={"bytes":p.stat().st_size,"sha256":digest(p)}
    for p in children(config,CONFIGS):
        files[p.relative_to(package).as_posix()]={"bytes":p.stat().st_size,"sha256":digest(p)}
    for name in CONFIGS[:2]:
        if not (config/name).is_file():blockers.append("Missing legacy configuration: "+name)
    for name in ("owner_review_drafts.csv.tmp","owner_review_drafts.csv.bak"):
        if (runtime/name).exists():blockers.append("Interrupted legacy review draft store: "+name)
    # Validate pointer and cache paths before any engine reads.
    for rel,parent in (("approved/current.json","approved/versions"),("staging/latest.json","staging")):
        p=runtime/rel
        if p.exists():
            run=str(json.loads(p.read_text(encoding="utf-8")).get("run_id",""))
            if not re.fullmatch(r"[A-Za-z0-9_-]+",run) or not (runtime/parent/run/"records.jsonl").is_file():raise ValueError("Invalid legacy run pointer: "+rel)
    for rel in files:
        if rel.endswith("/cache_index.json"):
            cache=json.loads((package/rel).read_text(encoding="utf-8"))
            if not isinstance(cache,dict):raise ValueError("Invalid legacy cache index")
            for value in cache.values():
                if not isinstance(value,str) or not (runtime/value).resolve().is_relative_to((runtime/"cache").resolve()):raise ValueError("Unsafe legacy cache reference")
                if not (runtime/value).is_file():blockers.append("Missing extraction cache referenced by "+rel)
    approved=load_approved_state(runtime);stage=load_latest_stage(runtime)
    events=approved["records"];keys=[r.get("Event_Key") for r in events]
    if not all(keys) or len(set(keys))!=len(keys):blockers.append("Approved event identities missing or duplicated")
    drafts=csv_rows(runtime/"owner_review_drafts.csv") if (runtime/"owner_review_drafts.csv").is_file() else []
    if drafts:
        warnings.append("Legacy draft CSV retained as evidence; active draft translation requires exact identity reconciliation.")
        if any(r.get("Run ID")==stage["run_id"] for r in drafts) and stage.get("decision",{}).get("decision") not in ("APPROVED","REJECTED"):blockers.append("Active legacy review draft requires explicit translation before import")
    workbook_paths=sorted(p for p in package.glob("*.xlsm") if not p.name.startswith("~$"))
    tables={};workbook=""
    if len(workbook_paths)>1:blockers.append("Multiple XLSM workbooks; choose the exact legacy baseline before import")
    elif workbook_paths:
        p=workbook_paths[0]
        if p.is_symlink():raise ValueError("Legacy workbook symlink")
        workbook=p.name;files[p.name]={"bytes":p.stat().st_size,"sha256":digest(p)}
        tables,table_warnings=workbook_tables(p);warnings.extend(table_warnings)
        if table_warnings:blockers.append("Formula-backed legacy settings require review")
    else:warnings.append("No legacy XLSM present; workbook-only settings could not be checked")
    fields_path=runtime/"user_custom_fields.csv";maps_path=runtime/"user_keyword_mappings.csv"
    derived={"fields":csv_rows(fields_path) if fields_path.exists() else tables.get("CustomFields",{}).get("rows",[]),"mappings":csv_rows(maps_path) if maps_path.exists() else tables.get("KeywordMappings",{}).get("rows",[])}
    try:
        derived=custom_fields.validate(derived,list(OUTPUT_FIELDS)+["Global Document Key","Revision Key","Source Workbook Path"])
        for key,name in (("fields","CustomFields"),("mappings","KeywordMappings")):
            if name in tables and (fields_path if key=="fields" else maps_path).exists():
                alt=dict(derived);alt[key]=tables[name]["rows"]
                if custom_fields.validate(alt,list(OUTPUT_FIELDS)+["Global Document Key","Revision Key","Source Workbook Path"])[key]!=derived[key]:blockers.append("Workbook/runtime custom configuration differs: "+name)
    except ValueError as e:blockers.append("Custom-field incompatibility: "+str(e))
    if "ClassificationRules" in tables and (config/CONFIGS[0]).exists() and tables["ClassificationRules"]["rows"]!=csv_rows(config/CONFIGS[0]):blockers.append("Workbook and CSV classification rules differ")
    source=approved["manifest"].get("data_dir","") or stage["manifest"].get("data_dir","")
    if source and not Path(source).is_dir():warnings.append("Original DATA path is unavailable; explicit path reconciliation required")
    evidence={"package":str(package),"files":files,"approved_run":approved["run_id"],"approved_events":len(events),"approved_keys_sha256":hashlib.sha256(json.dumps(sorted(str(k) for k in keys),ensure_ascii=False).encode()).hexdigest(),"pending_run":stage["run_id"],"legacy_draft_rows":len(drafts),"workbook":workbook,"workbook_tables":{k:{"sheet":v["sheet"],"rows":len(v["rows"])} for k,v in tables.items()},"custom_field_count":len(derived["fields"]),"keyword_mapping_count":len(derived["mappings"]),"source_path":source,"warnings":warnings,"blockers":sorted(set(blockers))}
    evidence["preview_sha256"]=hashlib.sha256(json.dumps(evidence,sort_keys=True,ensure_ascii=False).encode()).hexdigest()
    return evidence
