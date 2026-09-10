from __future__ import annotations
import hashlib
from collections import defaultdict
from typing import Dict, List, Sequence, Tuple
from .rules import Rule, file_exclusion


def assign_selection(workbooks: List[Dict[str,object]], rules: Sequence[Rule]) -> Tuple[List[Dict[str,object]],List[Dict[str,object]]]:
    by_hash = defaultdict(list)
    for wb in workbooks: by_hash[str(wb["sha256"])].append(wb)
    exact = []
    for h,group in sorted(by_hash.items()):
        if len(group)>1:
            gid=f"EXACT-{h[:12]}"; exact.append({"group_id":gid,"sha256":h,"files":sorted(str(w["relative_path"]) for w in group)})
            for wb in group: wb["warnings"].append("EXACT_BYTE_DUPLICATE")
    logical=defaultdict(list)
    for wb in workbooks:
        wb.update({"duplicate_version_group_id":"","selected_excluded_status":"","selection_exclusion_reason":"","selected_replacement_file":""})
        if wb["readability_status"]!="READABLE":
            wb["selected_excluded_status"]="EXCLUDED"; wb["selection_exclusion_reason"]=str(wb["unreadable_reason"] or wb["readability_status"]); continue
        ex=file_exclusion(rules,str(wb["source_family"]),str(wb["relative_path"]))
        if ex:
            wb["selected_excluded_status"]="EXCLUDED"; wb["selection_exclusion_reason"]=str(ex["notes"]); wb["warnings"].append("FILE_EXCLUDED_BY_RULE"); continue
        projects=list(wb.get("inferred_project_numbers",[]))
        if not projects:
            wb["selected_excluded_status"]="REVIEW_REQUIRED"; wb["selection_exclusion_reason"]="No project identity inferred from path"; wb["warnings"].append("PROJECT_ID_NOT_INFERRED"); continue
        logical[(str(wb["source_family"]),"+".join(projects),str(wb["logical_register_identity"]))].append(wb)
    groups=[]
    for key in sorted(logical):
        group=logical[key]; gid="VER-"+hashlib.sha1("|".join(key).encode()).hexdigest()[:10]
        for wb in group: wb["duplicate_version_group_id"]=gid if len(group)>1 else ""
        if len(group)==1:
            group[0]["selected_excluded_status"]="SELECTED"; group[0]["selection_exclusion_reason"]="Only confirmed source for logical register"; continue
        reliable=[w for w in group if w.get("timestamp_reliable") and w.get("doc_modified")]
        if len(reliable)!=len(group):
            for wb in group:
                wb["selected_excluded_status"]="REVIEW_REQUIRED"; wb["selection_exclusion_reason"]="Version group has incomplete trustworthy modified timestamps"; wb["warnings"].append("VERSION_TIMESTAMP_INCOMPLETE")
            groups.append({"group_id":gid,"key":key,"decision":"REVIEW_REQUIRED","files":sorted(str(w["relative_path"]) for w in group)}); continue
        ordered=sorted(group,key=lambda w:(str(w.get("doc_modified","")),str(w["relative_path"])),reverse=True)
        newest=ordered[0]; tied=[w for w in ordered if w.get("doc_modified")==newest.get("doc_modified")]
        if len(tied)>1: newest["warnings"].append("VERSION_TIMESTAMP_TIE_DETERMINISTIC_PATH_TIEBREAK")
        newest["selected_excluded_status"]="SELECTED"; newest["selection_exclusion_reason"]=f"Newest trustworthy core modified timestamp: {newest['doc_modified']}"; newest["warnings"].append("VERSION_GROUP_SELECTED_NEWEST")
        for old in ordered[1:]:
            old["selected_excluded_status"]="SUPERSEDED"; old["warnings"].append("SUPERSEDED_VERSION")
            old["selection_exclusion_reason"]=f"Superseded by {newest['relative_path']} using core modified timestamp"; old["selected_replacement_file"]=newest["relative_path"]
        groups.append({"group_id":gid,"key":key,"decision":"SELECTED_NEWEST","selected":newest["relative_path"],"selected_modified":newest["doc_modified"],"superseded":[w["relative_path"] for w in ordered[1:]]})
    return exact,groups
