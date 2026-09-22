from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence

from .core import norm_text


@dataclass(frozen=True)
class Rule:
    rule_id: str
    priority: int
    family: str
    scope: str
    match_type: str
    words: str
    exclude_words: str
    path_qualifier: str
    requires_discipline: str
    requires_category: str
    discipline: str
    category: str
    subcategory: str
    include: str
    min_confidence: float
    stop: bool
    notes: str


def load_rules(path: Path) -> List[Rule]:
    rules: List[Rule] = []
    with path.open(newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            if row.get("Enabled", "YES").strip().upper() != "YES":
                continue
            rule = Rule(
                row["Rule_ID"].strip(),
                int(row["Priority"]),
                row["Source_Family"].strip().upper(),
                row["Match_Scope"].strip().upper(),
                row["Match_Type"].strip().upper().replace(" ", "_"),
                row["Match_Words"],
                row.get("Exclude_Words", ""),
                row.get("Path_Qualifier", "").strip(),
                row.get("Requires_Discipline", "").strip(),
                row.get("Requires_Category", "").strip(),
                row["Discipline"].strip(),
                row["Category"].strip(),
                row["Subcategory"].strip(),
                row["Include"].strip().upper(),
                float(row.get("Min_Confidence") or 0),
                row.get("Stop_On_Match", "NO").strip().upper() == "YES",
                row.get("Notes", "").strip(),
            )
            # REGEX remains readable only for backward compatibility with old user
            # configuration files. The shipped rules and normal Excel editor use
            # plain-text match types instead.
            if rule.match_type == "REGEX":
                re.compile(rule.words)
            rules.append(rule)

    ids = [r.rule_id for r in rules]
    if len(ids) != len(set(ids)):
        raise ValueError("classification_rules.csv contains duplicate Rule_ID values")
    return sorted(rules, key=lambda r: (r.priority, r.rule_id))


def _canonical_evidence(scope: str, text: str) -> str:
    """Normalize presentation-only whitespace without changing stored source evidence."""
    text = "" if text is None else str(text)
    if scope in {"WORKSHEET", "SECTION"}:
        return re.sub(r"\s+", " ", text).strip()
    return text


def _split_plain_terms(value: str) -> List[str]:
    """Split optional plain-text alternatives without treating them as regex."""
    return [part.strip() for part in str(value or "").split("|") if part.strip()]


def _path_allowed(rule: Rule, evidence: Dict[str, str]) -> bool:
    if not rule.path_qualifier:
        return True
    actual = norm_text(evidence.get("FILE", "") or "")
    return all(norm_text(term) in actual for term in _split_plain_terms(rule.path_qualifier))


def _required_value_allowed(required: str, actual: object) -> bool:
    if not required:
        return True
    allowed = {norm_text(value) for value in required.split("|") if value.strip()}
    return norm_text(str(actual)) in allowed


def _context_allowed(rule: Rule, result: Dict[str, object]) -> bool:
    """Allow refining rules only in the taxonomy context declared in configuration."""
    return _required_value_allowed(rule.requires_discipline, result.get("discipline", "")) and _required_value_allowed(
        rule.requires_category, result.get("category", "")
    )


def _plain_exclusion_matches(rule: Rule, text: str) -> bool:
    if not rule.exclude_words:
        return False
    normalized = norm_text(text)
    return any(norm_text(term) in normalized for term in _split_plain_terms(rule.exclude_words))


def _match(rule: Rule, text: str) -> bool:
    if _plain_exclusion_matches(rule, text):
        return False

    actual = norm_text(text)
    expected = norm_text(rule.words)
    if rule.match_type == "EXACT":
        return actual == expected
    if rule.match_type == "CONTAINS":
        return expected in actual
    if rule.match_type == "STARTS_WITH":
        return actual.startswith(expected)
    if rule.match_type == "ENDS_WITH":
        return actual.endswith(expected)
    if rule.match_type == "FUZZY":
        from difflib import SequenceMatcher

        threshold = rule.min_confidence * 100 if rule.min_confidence <= 1 else rule.min_confidence
        return SequenceMatcher(None, actual, expected).ratio() * 100 >= threshold
    if rule.match_type == "REGEX":
        # Legacy-only compatibility. New rules should use EXACT / CONTAINS /
        # STARTS_WITH / ENDS_WITH so non-coders can understand and edit them.
        return re.search(rule.words, text, flags=re.I) is not None
    return False


def apply_classification(rules: Sequence[Rule], family: str, evidence: Dict[str, str]) -> Dict[str, object]:
    result = {
        "status": "UNCLASSIFIED",
        "discipline": "REVIEW_REQUIRED",
        "category": "UNCLASSIFIED",
        "subcategory": "UNCLASSIFIED",
        "rule_ids": [],
        "match_basis": [],
        "confidence": 0.0,
        "notes": "No matching v2 rule; manual review required",
    }

    for scope in ["FILE", "WORKSHEET", "SECTION", "HEADER", "DOC_NUMBER", "TITLE"]:
        text = _canonical_evidence(scope, evidence.get(scope, "") or "")
        if not text:
            continue
        for rule in rules:
            if rule.scope != scope or rule.family not in {"ANY", family}:
                continue
            if not _path_allowed(rule, evidence):
                continue
            if not _context_allowed(rule, result):
                continue
            if not _match(rule, text):
                continue

            result["rule_ids"].append(rule.rule_id)
            result["match_basis"].append(f"{scope}:{rule.rule_id}")
            result["confidence"] = max(
                float(result["confidence"]), 1.0 if rule.match_type == "EXACT" else 0.95
            )

            if rule.include == "NO":
                result.update(
                    {
                        "status": "EXCLUDED",
                        "discipline": "EXCLUDED",
                        "category": "EXCLUDED",
                        "subcategory": "EXCLUDED",
                        "notes": rule.notes,
                    }
                )
                return result

            for key, val in (
                ("discipline", rule.discipline),
                ("category", rule.category),
                ("subcategory", rule.subcategory),
            ):
                if val and val.upper() not in {"KEEP", "KEEP EXISTING"}:
                    result[key] = val
            result["status"] = "INCLUDE"
            result["notes"] = rule.notes
            if rule.stop:
                break

    if result["status"] == "INCLUDE" and "REVIEW_REQUIRED" in {
        result["discipline"],
        result["category"],
        result["subcategory"],
    }:
        result["status"] = "UNCLASSIFIED"
    return result


def file_exclusion(rules: Sequence[Rule], family: str, relative_path: str) -> Optional[Dict[str, object]]:
    c = apply_classification(rules, family, {"FILE": relative_path})
    return c if c["status"] == "EXCLUDED" else None
