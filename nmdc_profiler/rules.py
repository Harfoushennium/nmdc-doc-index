from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence
from .core import norm_text

@dataclass(frozen=True)
class Rule:
    rule_id: str; priority: int; family: str; scope: str; match_type: str; words: str
    exclude_words: str; discipline: str; category: str; subcategory: str; include: str
    min_confidence: float; stop: bool; notes: str


def load_rules(path: Path) -> List[Rule]:
    rules = []
    with path.open(newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            if row.get("Enabled","YES").strip().upper() != "YES": continue
            rules.append(Rule(row["Rule_ID"].strip(), int(row["Priority"]), row["Source_Family"].strip().upper(),
                              row["Match_Scope"].strip().upper(), row["Match_Type"].strip().upper(), row["Match_Words"],
                              row.get("Exclude_Words",""), row["Discipline"].strip(), row["Category"].strip(),
                              row["Subcategory"].strip(), row["Include"].strip().upper(), float(row.get("Min_Confidence") or 0),
                              row.get("Stop_On_Match","NO").strip().upper()=="YES", row.get("Notes","").strip()))
    ids = [r.rule_id for r in rules]
    if len(ids) != len(set(ids)): raise ValueError("classification_rules.csv contains duplicate Rule_ID values")
    return sorted(rules, key=lambda r:(r.priority,r.rule_id))


def _match(rule: Rule, text: str) -> bool:
    if rule.exclude_words and re.search(rule.exclude_words, text, flags=re.I): return False
    if rule.match_type == "EXACT": return norm_text(text) == norm_text(rule.words)
    if rule.match_type == "CONTAINS": return norm_text(rule.words) in norm_text(text)
    if rule.match_type == "REGEX": return re.search(rule.words, text) is not None
    if rule.match_type == "FUZZY":
        from difflib import SequenceMatcher
        threshold = rule.min_confidence*100 if rule.min_confidence <= 1 else rule.min_confidence
        return SequenceMatcher(None,norm_text(text),norm_text(rule.words)).ratio()*100 >= threshold
    return False


def apply_classification(rules: Sequence[Rule], family: str, evidence: Dict[str,str]) -> Dict[str,object]:
    result = {"status":"UNCLASSIFIED","discipline":"REVIEW_REQUIRED","category":"UNCLASSIFIED","subcategory":"UNCLASSIFIED",
              "rule_ids":[],"match_basis":[],"confidence":0.0,"notes":"No matching v1 rule; manual review required"}
    for scope in ["FILE","WORKSHEET","SECTION","HEADER","DOC_NUMBER","TITLE"]:
        text = evidence.get(scope,"") or ""
        if not text: continue
        for rule in rules:
            if rule.scope != scope or rule.family not in {"ANY",family} or not _match(rule,text): continue
            result["rule_ids"].append(rule.rule_id); result["match_basis"].append(f"{scope}:{rule.rule_id}")
            result["confidence"] = max(float(result["confidence"]), 1.0 if rule.match_type=="EXACT" else 0.95)
            if rule.include == "NO":
                result.update({"status":"EXCLUDED","discipline":"EXCLUDED","category":"EXCLUDED","subcategory":"EXCLUDED","notes":rule.notes})
                return result
            for key,val in (("discipline",rule.discipline),("category",rule.category),("subcategory",rule.subcategory)):
                if val and val.upper() not in {"KEEP","KEEP EXISTING"}: result[key]=val
            result["status"]="INCLUDE"; result["notes"]=rule.notes
            if rule.stop: break
    if result["status"]=="INCLUDE" and "REVIEW_REQUIRED" in {result["discipline"],result["category"],result["subcategory"]}:
        result["status"]="UNCLASSIFIED"
    return result


def file_exclusion(rules: Sequence[Rule], family: str, relative_path: str) -> Optional[Dict[str,object]]:
    c = apply_classification(rules,family,{"FILE":relative_path})
    return c if c["status"]=="EXCLUDED" else None
