from __future__ import annotations
import csv,json
from collections import defaultdict
from pathlib import Path
from typing import Dict,List,Sequence


def write_csv(path:Path,rows:Sequence[Dict[str,object]],fields:Sequence[str])->None:
    path.parent.mkdir(parents=True,exist_ok=True)
    with path.open("w",newline="",encoding="utf-8") as f:
        w=csv.DictWriter(f,fieldnames=fields,extrasaction="ignore",lineterminator="\n"); w.writeheader()
        for row in rows: w.writerow({k:row.get(k,"") for k in fields})


def write_outputs(workbooks:List[Dict[str,object]],class_rows:List[Dict[str,object]],exact_groups:List[Dict[str,object]],version_groups:List[Dict[str,object]],output_dir:Path)->None:
    output_dir.mkdir(parents=True,exist_ok=True)
    inv_fields=["relative_path","filename","source_family","file_extension","file_size","sha256","sheet_count","doc_created","doc_modified","timestamp_source","timestamp_reliable","readability_status","unreadable_reason","password_retry_possible","inferred_project_numbers","logical_register_identity","duplicate_version_group_id","selected_excluded_status","selection_exclusion_reason","selected_replacement_file","warning_codes","native_hyperlink_count","hyperlink_formula_count","merged_range_count"]
    inv=[]
    for wb in sorted(workbooks,key=lambda w:str(w["relative_path"]).lower()):
        r=dict(wb); r["inferred_project_numbers"]=";".join(wb.get("inferred_project_numbers",[])); r["warning_codes"]=";".join(sorted(set(wb.get("warnings",[])))); inv.append(r)
    write_csv(output_dir/"source_inventory.csv",inv,inv_fields)
    (output_dir/"workbook_profiles.json").write_text(json.dumps(sorted(workbooks,key=lambda w:str(w["relative_path"]).lower()),indent=2,ensure_ascii=False,sort_keys=True)+"\n",encoding="utf-8")
    cls_fields=["workbook_path","project_number","source_family","worksheet_name","original_section","normalized_worksheet","sample_document_numbers","sample_title_keywords","discipline","category","subcategory","matched_rule_ids","match_basis","confidence","proposed_action","exclusion_reason","suggested_aliases","warning_code","notes"]
    write_csv(output_dir/"classification_discovery.csv",class_rows,cls_fields)
    selected=[w for w in workbooks if w.get("selected_excluded_status")=="SELECTED"]; excluded=[w for w in workbooks if w.get("selected_excluded_status")=="EXCLUDED"]; superseded=[w for w in workbooks if w.get("selected_excluded_status")=="SUPERSEDED"]; review=[w for w in workbooks if w.get("selected_excluded_status")=="REVIEW_REQUIRED"]; unreadable=[w for w in workbooks if w.get("readability_status")!="READABLE"]
    fam=defaultdict(int)
    for w in workbooks: fam[str(w["source_family"])]+=1
    lines=["# Source Selection Report — NMDC Document Index Cycle 1","","## Candidate sources","",f"- Total candidates: **{len(workbooks)}**",f"- METHODS: **{fam['METHODS']}**",f"- TECH: **{fam['TECH']}**",f"- Selected: **{len(selected)}**",f"- Excluded: **{len(excluded)}**",f"- Superseded: **{len(superseded)}**",f"- Review required at workbook-selection level: **{len(review)}**",f"- Encrypted/unreadable: **{len(unreadable)}**",f"- Classification rows requiring review: **{sum(1 for r in class_rows if r.get('proposed_action')=='REVIEW_REQUIRED')}**","","## Exact byte duplicate groups",""]
    lines += [f"- `{g['group_id']}`: "+", ".join(f"`{p}`" for p in g["files"]) for g in exact_groups] or ["- None"]
    lines += ["","## Logical version groups",""]
    if version_groups:
        for g in version_groups:
            lines.append(f"- `{g['group_id']}` selected `{g['selected']}` at `{g['selected_modified']}`; superseded: "+", ".join(f"`{p}`" for p in g["superseded"]) if g["decision"]=="SELECTED_NEWEST" else f"- `{g['group_id']}` requires review: "+", ".join(f"`{p}`" for p in g["files"]))
    else: lines.append("- None")
    lines += ["","## Exclusions",""]+[f"- `{w['relative_path']}` — {w['selection_exclusion_reason']}" for w in excluded]
    if not excluded: lines.append("- None")
    lines += ["","## Worksheet exclusions / support views",""]
    wsx=[r for r in class_rows if r.get("proposed_action")=="EXCLUDE" and r.get("worksheet_name")!="[UNREADABLE]"]
    lines += [f"- `{r['workbook_path']} :: {r['worksheet_name']}` — {r.get('exclusion_reason') or r.get('notes') or 'Excluded by rule'}" for r in wsx] or ["- None"]
    lines += ["","## Project mismatches",""]
    mm=[(w,f) for w in workbooks for f in w.get("project_mismatch_findings",[])]
    lines += [f"- `{w['relative_path']}` — path={f['path_projects']} internal={f['internal_projects']} — {f['evidence']}" for w,f in mm] or ["- None detected from explicit internal PROJECT NO evidence"]
    lines += ["","## Unknown layouts / review required",""]
    unknown=sorted({r["workbook_path"] for r in class_rows if r.get("proposed_action")=="REVIEW_REQUIRED"}); lines += [f"- `{p}`" for p in unknown] or ["- None"]
    lines += ["","## DATA integrity","","The profiler is read-only with respect to `DATA/`. It writes only to the requested output directory.",""]
    (output_dir/"source_selection_report.md").write_text("\n".join(lines),encoding="utf-8")
