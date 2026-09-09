from __future__ import annotations
import re
from collections import defaultdict
from typing import Dict, List, Sequence
from .core import norm_text
from .rules import Rule, apply_classification, file_exclusion


def classification_rows(workbooks: Sequence[Dict[str,object]], rules: Sequence[Rule]) -> List[Dict[str,object]]:
    rows=[]
    for wb in workbooks:
        family,file_ev=str(wb["source_family"]),str(wb["relative_path"])
        if wb["readability_status"]!="READABLE":
            rows.append({"workbook_path":wb["relative_path"],"project_number":", ".join(wb.get("inferred_project_numbers",[])),"source_family":family,"worksheet_name":"[UNREADABLE]","original_section":"","normalized_worksheet":"","sample_document_numbers":"","sample_title_keywords":"","discipline":"EXCLUDED","category":"EXCLUDED","subcategory":"EXCLUDED","matched_rule_ids":"","match_basis":"","confidence":"","proposed_action":"EXCLUDE","exclusion_reason":wb["unreadable_reason"],"suggested_aliases":"","warning_code":"UNREADABLE_SOURCE","notes":"Password-assisted retry possible" if wb.get("password_retry_possible") else "Unreadable source"}); continue
        file_ex=file_exclusion(rules,family,file_ev)
        for sheet in wb.get("sheets",[]):
            sname=str(sheet.get("sheet_name","")); sections=sheet.get("sections",[])
            use_sections=family=="METHODS" and re.search(r"(?i)in+com+ing.*(?:doc|drg|drawing)",sname) and sections
            items=sections if use_sections else [{"label":"","row":"","cell":""}]
            for sec in items:
                evidence={"FILE":file_ev,"WORKSHEET":sname,"SECTION":str(sec.get("label","")),"HEADER":" | ".join(sheet.get("representative_header_values",[])),"DOC_NUMBER":" | ".join(sheet.get("sample_document_numbers",[])),"TITLE":" | ".join(sheet.get("sample_titles",[]))}
                c=apply_classification(rules,family,evidence)
                if file_ex: c=file_ex
                action="INCLUDE" if c["status"]=="INCLUDE" else "EXCLUDE" if c["status"]=="EXCLUDED" else "REVIEW_REQUIRED"
                warning="REVIEW_REQUIRED" if c["status"]=="UNCLASSIFIED" else ""
                if wb.get("project_mismatch_findings"): warning=";".join(filter(None,[warning,"PROJECT_MISMATCH"]))
                rows.append({"workbook_path":wb["relative_path"],"project_number":", ".join(wb.get("inferred_project_numbers",[])),"source_family":family,"worksheet_name":sname,"original_section":sec.get("label",""),"normalized_worksheet":norm_text(sname),"sample_document_numbers":" | ".join(sheet.get("sample_document_numbers",[])[:8]),"sample_title_keywords":" | ".join(sheet.get("sample_titles",[])[:5]),"discipline":c["discipline"],"category":c["category"],"subcategory":c["subcategory"],"matched_rule_ids":";".join(c["rule_ids"]),"match_basis":";".join(c["match_basis"]),"confidence":f"{float(c['confidence']):.2f}","proposed_action":action,"exclusion_reason":c["notes"] if action=="EXCLUDE" else "","suggested_aliases":"","warning_code":warning,"notes":c["notes"]})
    return sorted(rows,key=lambda r:(str(r["workbook_path"]).lower(),str(r["worksheet_name"]).lower(),str(r["original_section"]),str(r["matched_rule_ids"])))


def enrich_profiles_with_classification(workbooks: Sequence[Dict[str,object]], rows: Sequence[Dict[str,object]]) -> None:
    by_key=defaultdict(list)
    for row in rows: by_key[(str(row["workbook_path"]),str(row["worksheet_name"]))].append(row)
    for wb in workbooks:
        for sheet in wb.get("sheets",[]):
            matches=by_key.get((str(wb["relative_path"]),str(sheet.get("sheet_name",""))),[])
            sheet["candidate_classifications"]=[{"section":r.get("original_section",""),"action":r.get("proposed_action",""),"discipline":r.get("discipline",""),"category":r.get("category",""),"subcategory":r.get("subcategory",""),"rule_ids":r.get("matched_rule_ids",""),"warning_code":r.get("warning_code","")} for r in matches]
            actions={str(r.get("proposed_action")) for r in matches}
            sheet["worksheet_role"]="SUPPORT_ADMIN" if actions=={"EXCLUDE"} else "DOCUMENT_REGISTER" if "INCLUDE" in actions else "REVIEW_REQUIRED"
            sheet["support_admin_suspicion"]=actions=={"EXCLUDE"}
            sheet["duplicate_view_suspicion"]=norm_text(sheet.get("sheet_name","")) in {"client","client copy","alternate view"}
            sheet["warnings"]=sorted({str(r.get("warning_code")) for r in matches if r.get("warning_code")})
