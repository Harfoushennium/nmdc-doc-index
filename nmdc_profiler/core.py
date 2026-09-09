from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_DOC_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DCTERMS_NS = "http://purl.org/dc/terms/"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
NS = {"main": MAIN_NS, "r": REL_DOC_NS, "dcterms": DCTERMS_NS}
DOCNO_RE = re.compile(r"\b[A-Z0-9]{2,}(?:[-_/][A-Z0-9]{1,}){2,}\b", re.I)
PROJECT_LABEL_RE = re.compile(r"(?i)\b(?:nmdc\s+)?project\s*(?:no\.?|number)?\b")
PROJECT_VALUE_RE = re.compile(r"\b\d{4}\b")
CELL_REF_RE = re.compile(r"^([A-Z]+)(\d+)$")


def norm_text(value: object) -> str:
    s = "" if value is None else str(value)
    s = s.replace("&", " and ")
    s = re.sub(r"[_\-]+", " ", s)
    s = re.sub(r"[^\w\s/]", " ", s, flags=re.UNICODE)
    return re.sub(r"\s+", " ", s).strip().lower()


def rel_posix(path: Path, base: Path) -> str:
    return path.resolve().relative_to(base.resolve()).as_posix()


def get_source_family(path: Path, data_dir: Path) -> str:
    try:
        parts = path.resolve().relative_to(data_dir.resolve()).parts
    except ValueError:
        return "UNKNOWN"
    return parts[0].upper() if parts and parts[0].upper() in {"METHODS", "TECH"} else "UNKNOWN"


def infer_path_projects(filename: str) -> List[str]:
    stem = Path(filename).stem.strip()
    m = re.match(r"(?i)^(?:E\s*[-_ ]\s*)?(\d{4})(?:\s*[-/]\s*(\d{4}))?", stem)
    return [x for x in m.groups() if x] if m else []


def infer_logical_register_identity(filename: str, family: str) -> str:
    n = norm_text(Path(filename).stem)
    if "installation aids register" in n:
        return "INSTALLATION_AIDS_REGISTER"
    if re.search(r"\bmdr\b|master document register", n):
        return "MASTER_DOCUMENT_REGISTER"
    if "offshore construction engineering register" in n:
        return "OFFSHORE_CONSTRUCTION_ENGINEERING_REGISTER"
    if "document register" in n or "offshore support deliverables register" in n:
        return "DOCUMENT_REGISTER"
    if "pre requisites" in n or "prerequisite" in n:
        return "PREREQUISITE_MATRIX"
    if "format" in n or "template" in n:
        return "TEMPLATE"
    if "barge" in n and "sketch" in n:
        return "BARGE_SKETCH_REGISTER"
    if family == "METHODS" and ("deliverable" in n or "delivarable" in n):
        return "METHODS_DELIVERABLE_REGISTER"
    if "procedure status" in n or "harfoush" in n:
        return "PERSONAL_TRACKER"
    return "UNKNOWN_REGISTER"


def col_number(ref: str) -> int:
    m = CELL_REF_RE.match(ref or "")
    if not m:
        return 0
    n = 0
    for ch in m.group(1):
        n = n * 26 + ord(ch) - 64
    return n


def merge_pattern(ref: str) -> Dict[str, object]:
    try:
        a, b = ref.split(":", 1)
        ma, mb = CELL_REF_RE.match(a), CELL_REF_RE.match(b)
        if not ma or not mb:
            raise ValueError
        return {"range": ref, "row_span": int(mb.group(2)) - int(ma.group(2)) + 1,
                "column_span": col_number(b) - col_number(a) + 1}
    except Exception:
        return {"range": ref, "row_span": None, "column_span": None}
