from __future__ import annotations

import hashlib
import re
import zipfile
from pathlib import Path
from typing import Dict, List, Sequence, Tuple
import xml.etree.ElementTree as ET

from .core import (CELL_REF_RE, CP_NS, DCTERMS_NS, DOCNO_RE, MAIN_NS, PROJECT_LABEL_RE,
                   PROJECT_VALUE_RE, REL_DOC_NS, REL_PKG_NS, col_number, get_source_family,
                   infer_logical_register_identity, infer_path_projects, merge_pattern, rel_posix)


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def readability(path: Path) -> Tuple[str, str]:
    try:
        magic = path.read_bytes()[:8]
    except OSError as exc:
        return "UNREADABLE", f"OPEN_FAILED:{exc.__class__.__name__}"
    if magic == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "ENCRYPTED", "OLE2_CFB_ENCRYPTED"
    if not zipfile.is_zipfile(path):
        return "UNREADABLE", "NOT_OOXML_ZIP"
    return "READABLE", ""


def all_text(elem: ET.Element) -> str:
    return "".join(t.text or "" for t in elem.iter(f"{{{MAIN_NS}}}t"))


def load_shared_strings(z: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    return [all_text(si) for si in root.findall(f".//{{{MAIN_NS}}}si")]


def parse_core_properties(z: zipfile.ZipFile) -> Dict[str, object]:
    out = {"created": "", "modified": "", "timestamp_source": "", "timestamp_reliable": False}
    if "docProps/core.xml" not in z.namelist():
        return out
    root = ET.fromstring(z.read("docProps/core.xml"))
    created = root.find(f".//{{{DCTERMS_NS}}}created")
    if created is None:
        created = root.find(f".//{{{CP_NS}}}created")
    modified = root.find(f".//{{{DCTERMS_NS}}}modified")
    if modified is None:
        modified = root.find(f".//{{{CP_NS}}}modified")
    if created is not None and created.text:
        out["created"] = created.text.strip()
    if modified is not None and modified.text:
        out.update({"modified": modified.text.strip(), "timestamp_source": "docProps/core.xml:dcterms:modified",
                    "timestamp_reliable": True})
    elif out["created"]:
        out["timestamp_source"] = "docProps/core.xml:dcterms:created"
    return out


def workbook_sheet_map(z: zipfile.ZipFile) -> List[Dict[str, str]]:
    root = ET.fromstring(z.read("xl/workbook.xml"))
    rels: Dict[str, str] = {}
    if "xl/_rels/workbook.xml.rels" in z.namelist():
        rr = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        for rel in rr.findall(f".//{{{REL_PKG_NS}}}Relationship"):
            rels[rel.get("Id", "")] = rel.get("Target", "")
    sheets = []
    for s in root.findall(f".//{{{MAIN_NS}}}sheet"):
        rid = s.get(f"{{{REL_DOC_NS}}}id", "")
        target = rels.get(rid, "")
        xml_path = target.lstrip("/") if target.startswith("/") else "xl/" + target.lstrip("/")
        sheets.append({"name": s.get("name", ""), "sheet_id": s.get("sheetId", ""),
                       "state": s.get("state", "visible"), "relationship_id": rid,
                       "xml_path": re.sub(r"/\./", "/", xml_path)})
    return sheets


def cell_value(c: ET.Element, shared: Sequence[str]) -> str:
    typ = c.get("t", "")
    if typ == "inlineStr":
        return all_text(c)
    v = c.find(f"{{{MAIN_NS}}}v")
    raw = v.text if v is not None and v.text is not None else ""
    if typ == "s":
        try:
            return shared[int(raw)]
        except (ValueError, IndexError):
            return raw
    return raw


def parse_sheet(z: zipfile.ZipFile, sheet: Dict[str, str], shared: Sequence[str], sample_row_limit: int = 250) -> Dict[str, object]:
    path = sheet["xml_path"]
    if path not in z.namelist():
        return {"sheet_name": sheet["name"], "state": sheet["state"], "error": "MISSING_WORKSHEET_XML"}
    root = ET.fromstring(z.read(path))
    dim = root.find(f".//{{{MAIN_NS}}}dimension")
    used_range = dim.get("ref", "") if dim is not None else ""
    rows, max_row, max_col = [], 0, 0
    for row in root.findall(f".//{{{MAIN_NS}}}row"):
        rn = int(row.get("r", "0") or 0)
        max_row = max(max_row, rn)
        values = []
        for c in row.findall(f"{{{MAIN_NS}}}c"):
            ref = c.get("r", "")
            max_col = max(max_col, col_number(ref))
            value = cell_value(c, shared).strip()
            if value:
                values.append({"ref": ref, "value": value})
        if values and len(rows) < sample_row_limit:
            rows.append({"row": rn, "cells": values})
    merges = [m.get("ref", "") for m in root.findall(f".//{{{MAIN_NS}}}mergeCell")]
    rel_targets: Dict[str, str] = {}
    rel_path = path.rsplit("/", 1)[0] + "/_rels/" + path.rsplit("/", 1)[1] + ".rels"
    if rel_path in z.namelist():
        rr = ET.fromstring(z.read(rel_path))
        for rel in rr.findall(f".//{{{REL_PKG_NS}}}Relationship"):
            rel_targets[rel.get("Id", "")] = rel.get("Target", "")
    native = []
    for h in root.findall(f".//{{{MAIN_NS}}}hyperlink"):
        rid = h.get(f"{{{REL_DOC_NS}}}id", "")
        native.append({"ref": h.get("ref", ""), "target": rel_targets.get(rid, h.get("location", ""))})
    formulas = []
    for c in root.findall(f".//{{{MAIN_NS}}}c"):
        f = c.find(f"{{{MAIN_NS}}}f")
        if f is not None and f.text and "HYPERLINK(" in f.text.upper():
            formulas.append({"ref": c.get("r", ""), "formula": f.text[:300]})
    row_texts = [(row["row"], " | ".join(c["value"] for c in row["cells"])) for row in rows]
    header_terms = re.compile(r"(?i)\b(document|doc\.?\s*no|title|revision|rev\.?|project|description|status|date|drawing)\b")
    likely_headers = [rn for rn, txt in row_texts if header_terms.search(txt)][:8]
    representative_headers = [txt[:240] for rn, txt in row_texts if rn in likely_headers[:5]]
    sections, seen = [], set()
    for row in rows:
        for cell in row["cells"]:
            txt = re.sub(r"\s+", " ", cell["value"].strip().lower())
            if txt in {"drawing", "drawings", "document", "documents"}:
                label = "DRAWINGS" if txt.startswith("drawing") else "DOCUMENTS"
                key = (label, row["row"])
                if key not in seen:
                    seen.add(key); sections.append({"label": label, "row": row["row"], "cell": cell["ref"]})
    strings = [c["value"] for row in rows for c in row["cells"]]
    docnos, seen_doc = [], set()
    for txt in strings:
        for m in DOCNO_RE.findall(txt):
            k = m.upper()
            if k not in seen_doc:
                seen_doc.add(k); docnos.append(m)
                if len(docnos) >= 12: break
        if len(docnos) >= 12: break
    titles = []
    for txt in strings:
        if len(txt.strip()) >= 12 and not PROJECT_LABEL_RE.search(txt) and txt.strip() not in titles:
            titles.append(txt.strip()[:180])
        if len(titles) >= 12: break
    internal = set()
    for i, (_, txt) in enumerate(row_texts):
        if PROJECT_LABEL_RE.search(txt):
            vals = PROJECT_VALUE_RE.findall(txt)
            if not vals and i + 1 < len(row_texts): vals = PROJECT_VALUE_RE.findall(row_texts[i + 1][1])
            internal.update(vals)
    return {
        "sheet_name": sheet["name"], "state": sheet["state"], "used_range": used_range,
        "max_row": max_row, "max_column": max_col, "likely_header_rows": likely_headers,
        "representative_header_values": representative_headers, "merge_count": len(merges),
        "representative_merge_ranges": merges[:12], "representative_merge_patterns": [merge_pattern(x) for x in merges[:12]],
        "native_hyperlink_count": len(native), "native_hyperlinks": native[:20],
        "hyperlink_formula_count": len(formulas), "hyperlink_formulas": formulas[:20],
        "hidden_row_count": sum(1 for r in root.findall(f".//{{{MAIN_NS}}}row") if r.get("hidden") in {"1","true","TRUE"}),
        "hidden_column_count": sum(1 for c in root.findall(f".//{{{MAIN_NS}}}col") if c.get("hidden") in {"1","true","TRUE"}),
        "sections": sections, "sample_document_numbers": docnos, "sample_titles": titles,
        "candidate_internal_project_numbers": sorted(internal), "sample_rows": rows[:30],
    }


def profile_workbook(path: Path, root: Path, data_dir: Path) -> Dict[str, object]:
    relative, family = rel_posix(path, root), get_source_family(path, data_dir)
    status, reason = readability(path)
    out: Dict[str, object] = {
        "relative_path": relative, "filename": path.name, "source_family": family,
        "file_extension": path.suffix, "file_size": path.stat().st_size, "sha256": sha256_file(path),
        "readability_status": status, "unreadable_reason": reason, "password_retry_possible": status == "ENCRYPTED",
        "inferred_project_numbers": infer_path_projects(path.name),
        "logical_register_identity": infer_logical_register_identity(path.name, family), "warnings": [], "sheets": [],
    }
    if status != "READABLE":
        out.update({"doc_created":"","doc_modified":"","timestamp_source":"","timestamp_reliable":False,
                    "sheet_count":0,"candidate_internal_project_numbers":[],"project_mismatch_findings":[]})
        return out
    try:
        with zipfile.ZipFile(path) as z:
            core = parse_core_properties(z)
            out.update({"doc_created":core["created"], "doc_modified":core["modified"],
                        "timestamp_source":core["timestamp_source"], "timestamp_reliable":core["timestamp_reliable"]})
            shared = load_shared_strings(z); out["shared_string_count"] = len(shared)
            sheets = [parse_sheet(z, s, shared) for s in workbook_sheet_map(z)]
            out["sheets"], out["sheet_count"] = sheets, len(sheets)
            internal = sorted({p for s in sheets for p in s.get("candidate_internal_project_numbers", [])})
            out["candidate_internal_project_numbers"] = internal
            expected, actual = set(out["inferred_project_numbers"]), set(internal)
            findings = []
            if expected and actual and expected.isdisjoint(actual):
                findings.append({"code":"PROJECT_MISMATCH", "path_projects":sorted(expected), "internal_projects":sorted(actual),
                                 "evidence":"Path project identity is disjoint from explicit internal PROJECT NO evidence"})
                out["warnings"].append("PROJECT_MISMATCH")
            elif not actual:
                out["warnings"].append("PROJECT_ID_INTERNAL_NOT_FOUND")
            out["project_mismatch_findings"] = findings
            out["native_hyperlink_count"] = sum(int(s.get("native_hyperlink_count",0)) for s in sheets)
            out["hyperlink_formula_count"] = sum(int(s.get("hyperlink_formula_count",0)) for s in sheets)
            out["merged_range_count"] = sum(int(s.get("merge_count",0)) for s in sheets)
    except (zipfile.BadZipFile, KeyError, ET.ParseError, OSError) as exc:
        out.update({"readability_status":"UNREADABLE", "unreadable_reason":f"PARSE_FAILED:{exc.__class__.__name__}",
                    "sheets":[], "sheet_count":0})
        out["warnings"].append("WORKBOOK_PARSE_FAILED")
    return out
