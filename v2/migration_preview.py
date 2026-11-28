"""Read-only V1 migration preview; no import or active-state mutation."""
from __future__ import annotations
import csv, hashlib, json, re, posixpath
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from nmdc_profiler.ooxml import workbook_sheet_map,load_shared_strings,cell_value
from nmdc_profiler.excel_bridge import load_approved_state,load_latest_stage
from nmdc_profiler.file_review import affected_files
from v2.recovery import digest
from v2 import custom_fields
from nmdc_profiler.extractor import OUTPUT_FIELDS

NS="{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL="{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG="{http://schemas.openxmlformats.org/package/2006/relationships}"
TABLES={"Configuration","ClassificationRules","CustomFields","KeywordMappings","FileReview","ReviewFlags"}
MAX_INPUT_FILES=100_000
MAX_INPUT_BYTES=2*1024**3
MAX_INPUT_FILE=512*1024**2
RUNTIME_AREAS=("approved","staging","cache","logs","user_flags")
CONFIGS=("classification_rules.csv","project_identity_overrides.csv","source_exclusions.csv")
KNOWN_CONFIG=set(CONFIGS)|{"cycle2_sentinels.csv","classification_aliases.csv","custom_fields.json"}
KNOWN_RUNTIME={"approved","staging","cache","logs","user_flags","profile","owner_review_drafts.csv","owner_review_drafts.csv.tmp","owner_review_drafts.csv.bak","user_custom_fields.csv","user_keyword_mappings.csv"}
def overlap(a,b):return a==b or a in b.parents or b in a.parents
def _is_reparse(path):
    try:return bool(Path(path).lstat().st_file_attributes & 0x400)
    except (AttributeError,OSError):return False
def children(root,areas):
    for area in areas:
        path=root/area
        if path.is_symlink() or _is_reparse(path):raise ValueError("Symlink or reparse point in migration input: "+str(path))
        if path.is_file():yield path
        elif path.is_dir():
            for p in sorted(path.rglob("*")):
                if p.is_symlink() or _is_reparse(p):raise ValueError("Symlink or reparse point in migration input: "+str(p))
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
                if name not in TABLES:
                    out["UNSUPPORTED:"+name]={"sheet":sheet["name"],"columns":[e.get("name","") for e in table.findall(".//"+NS+"tableColumn")],"rows":[]}
                    continue
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
    raw=Path(package).absolute()
    if raw.is_symlink() or _is_reparse(raw):raise ValueError("Legacy package must not be a symlink or reparse point")
    package=raw.resolve()
    runtime=package/"runtime";config=package/"config"
    if runtime.is_symlink() or config.is_symlink() or _is_reparse(runtime) or _is_reparse(config):raise ValueError("Legacy runtime/config must not be symbolic links or reparse points")
    if not runtime.is_dir() or not config.is_dir():raise ValueError("Select the legacy package containing runtime and config folders.")
    blockers=[];warnings=[];files={};total_bytes=0
    def add_file(path):
        nonlocal total_bytes
        size=path.stat().st_size
        total_bytes+=size
        if size>MAX_INPUT_FILE or total_bytes>MAX_INPUT_BYTES or len(files)>=MAX_INPUT_FILES:raise ValueError("Legacy migration inventory exceeds configured bounds")
        files[path.relative_to(package).as_posix()]={"bytes":size,"sha256":digest(path)}
    for p in children(runtime,("approved","staging","cache","logs","user_flags","profile","owner_review_drafts.csv","owner_review_drafts.csv.tmp","owner_review_drafts.csv.bak","user_custom_fields.csv","user_keyword_mappings.csv")):
        add_file(p)
    for p in runtime.rglob("*"):
        if p.is_symlink() or _is_reparse(p):raise ValueError("Symlink or reparse point in migration input: "+str(p))
        if p.is_file() and p.relative_to(runtime).parts[0] not in KNOWN_RUNTIME:
            add_file(p)
            blockers.append("Unsupported legacy runtime content: "+p.relative_to(runtime).as_posix())
    for p in config.rglob("*"):
        if p.is_symlink() or _is_reparse(p):raise ValueError("Symlink or reparse point in migration input: "+str(p))
        if not p.is_file():continue
        add_file(p)
        if p.name not in KNOWN_CONFIG:blockers.append("Unsupported legacy configuration: "+p.relative_to(config).as_posix())
    history=package/"config-history"
    if history.is_symlink():raise ValueError("Symlink in migration input: "+str(history))
    if history.is_dir():
        for p in history.rglob("*"):
            if p.is_symlink() or _is_reparse(p):raise ValueError("Symlink or reparse point in migration input: "+str(p))
            if p.is_file():add_file(p)
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
    draft_translation=None
    if drafts:
        if set(drafts[0])!={"Run ID","Section","Identity","Field","Value"}:
            blockers.append("Unsupported legacy review draft CSV schema")
        if any(r.get("Section") not in ("FILES","FLAGS") for r in drafts):
            blockers.append("Unknown legacy review draft section")
        active=[r for r in drafts if r.get("Run ID")==stage["run_id"]]
        if active and stage.get("decision",{}).get("decision") not in ("APPROVED","REJECTED"):
            if any(r.get("Section")!="FILES" for r in active):
                blockers.append("Active legacy flag draft requires exact flag identity translation")
            else:
                try:
                    expected={r["Source File"] for r in affected_files(runtime,stage["run_id"])}
                    choices={}
                    for row in active:
                        source=row["Identity"];field=row["Field"]
                        if source not in expected or field not in ("Accept Update?","Owner Note") or field in choices.get(source,{}):
                            raise ValueError("Unknown, duplicate, or unsupported file draft field")
                        value=row["Value"]
                        if field=="Accept Update?":
                            if value.upper() not in ("TRUE","FALSE"):raise ValueError("Invalid file draft choice")
                            value=value.upper()=="TRUE"
                        choices.setdefault(source,{})[field]=value
                    if set(choices)!=expected or any(set(v)!={"Accept Update?","Owner Note"} for v in choices.values()):
                        raise ValueError("Legacy file draft does not contain every exact identity and field")
                    draft_translation={"run":stage["run_id"],"choices":{k:{"accept":v["Accept Update?"],"note":v["Owner Note"]} for k,v in choices.items()}}
                    warnings.append("Complete pending file-review draft translated; it remains dirty and requires owner review.")
                except (ValueError,KeyError) as e:blockers.append("Active legacy review draft cannot be translated: "+str(e))
        else:warnings.append("Inactive legacy review drafts retained as migration evidence.")
    workbook_paths=sorted(p for p in package.glob("*.xlsm") if not p.name.startswith("~$"))
    tables={};workbook="";workbook_data_dir=""
    if len(workbook_paths)>1:blockers.append("Multiple XLSM workbooks; choose the exact legacy baseline before import")
    elif workbook_paths:
        p=workbook_paths[0]
        if p.is_symlink() or _is_reparse(p):raise ValueError("Legacy workbook symlink or reparse point")
        workbook=p.name;add_file(p)
        tables,table_warnings=workbook_tables(p);warnings.extend(table_warnings)
        if table_warnings:blockers.append("Formula-backed legacy settings require review")
        unsupported=sorted(name.split(":",1)[1] for name in tables if name.startswith("UNSUPPORTED:"))
        if unsupported:blockers.append("Unsupported workbook tables: "+", ".join(unsupported))
        for name in ("FileReview","ReviewFlags"):
            if tables.get(name,{}).get("rows"):blockers.append("Workbook-only "+name+" rows require exact runtime reconciliation")
        if "Configuration" in tables:
            cols=tables["Configuration"]["columns"]
            rows=tables["Configuration"]["rows"]
            if cols != ["Setting","Current value"]:
                blockers.append("Unsupported workbook-only Configuration table schema")
            else:
                unknown=sorted({str(row.get("Setting","")) for row in rows}-{"Data Folder"})
                if unknown:blockers.append("Unsupported workbook-only settings: "+", ".join(unknown))
                workbook_source=next((str(row.get("Current value","")) for row in rows if row.get("Setting")=="Data Folder"),"")
                workbook_data_dir=workbook_source
                if sum(row.get("Setting")=="Data Folder" for row in rows)>1:blockers.append("Duplicate workbook Data Folder settings")
                approved_source=str(approved["manifest"].get("data_dir", ""))
                if workbook_source and approved_source and Path(workbook_source).resolve()!=Path(approved_source).resolve():
                    blockers.append("Workbook Data Folder differs from approved source metadata")
    else:warnings.append("No legacy XLSM present; workbook-only settings could not be checked")
    fields_path=runtime/"user_custom_fields.csv";maps_path=runtime/"user_keyword_mappings.csv";json_fields=config/"custom_fields.json"
    derived={"fields":csv_rows(fields_path) if fields_path.exists() else tables.get("CustomFields",{}).get("rows",[]),"mappings":csv_rows(maps_path) if maps_path.exists() else tables.get("KeywordMappings",{}).get("rows",[])}
    if json_fields.exists():
        try:
            config_fields=json.loads(json_fields.read_text(encoding="utf-8"))
            config_fields=custom_fields.validate(config_fields,list(OUTPUT_FIELDS)+["Global Document Key","Revision Key","Source Workbook Path"])
            if derived["fields"] or derived["mappings"]:
                checked=custom_fields.validate(derived,list(OUTPUT_FIELDS)+["Global Document Key","Revision Key","Source Workbook Path"])
                if checked!=config_fields:blockers.append("Legacy custom_fields.json differs from runtime/workbook custom configuration")
            else:derived=config_fields
        except (ValueError,TypeError,json.JSONDecodeError) as e:blockers.append("Custom-field incompatibility: "+str(e))
    try:
        derived=custom_fields.validate(derived,list(OUTPUT_FIELDS)+["Global Document Key","Revision Key","Source Workbook Path"])
        for key,name in (("fields","CustomFields"),("mappings","KeywordMappings")):
            if name in tables and (fields_path if key=="fields" else maps_path).exists():
                alt=dict(derived);alt[key]=tables[name]["rows"]
                if custom_fields.validate(alt,list(OUTPUT_FIELDS)+["Global Document Key","Revision Key","Source Workbook Path"])[key]!=derived[key]:blockers.append("Workbook/runtime custom configuration differs: "+name)
    except ValueError as e:blockers.append("Custom-field incompatibility: "+str(e))
    if "ClassificationRules" in tables and (config/CONFIGS[0]).exists() and tables["ClassificationRules"]["rows"]!=csv_rows(config/CONFIGS[0]):blockers.append("Workbook and CSV classification rules differ")
    source=approved["manifest"].get("data_dir","") or stage["manifest"].get("data_dir","") or workbook_data_dir
    if source and not Path(source).is_dir():warnings.append("Original DATA path is unavailable; explicit path reconciliation required")
    evidence={"package":str(package),"files":files,"approved_run":approved["run_id"],"approved_events":len(events),"approved_keys_sha256":hashlib.sha256(json.dumps(sorted(str(k) for k in keys),ensure_ascii=False).encode()).hexdigest(),"pending_run":stage["run_id"],"legacy_draft_rows":len(drafts),"legacy_draft_translation":draft_translation,"workbook":workbook,"workbook_tables":{k:{"sheet":v["sheet"],"rows":len(v["rows"])} for k,v in tables.items()},"custom_field_count":len(derived["fields"]),"keyword_mapping_count":len(derived["mappings"]),"custom_fields":derived,"source_path":source,"warnings":warnings,"blockers":sorted(set(blockers))}
    evidence["preview_sha256"]=hashlib.sha256(json.dumps(evidence,sort_keys=True,ensure_ascii=False).encode()).hexdigest()
    return evidence
