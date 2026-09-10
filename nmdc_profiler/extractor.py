from __future__ import annotations

import csv
import hashlib
import json
import re
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET

from .core import CELL_REF_RE, DOCNO_RE, MAIN_NS, REL_DOC_NS, REL_PKG_NS, col_number, infer_path_projects, norm_text
from .ooxml import cell_value, load_shared_strings, parse_core_properties, workbook_sheet_map
from .rules import Rule, apply_classification


@dataclass
class SheetModel:
    source_file: str
    source_family: str
    source_modified: str
    sheet_name: str
    max_row: int
    max_col: int
    cells: Dict[Tuple[int, int], str] = field(default_factory=dict)
    formulas: Dict[Tuple[int, int], str] = field(default_factory=dict)
    merge_anchor: Dict[Tuple[int, int], Tuple[int, int]] = field(default_factory=dict)
    hyperlinks: Dict[Tuple[int, int], str] = field(default_factory=dict)

    def value(self, row: int, col: int) -> str:
        value = self.cells.get((row, col), "")
        if value:
            return value
        anchor = self.merge_anchor.get((row, col))
        if anchor:
            return self.cells.get(anchor, "")
        return ""

    def hyperlink(self, row: int, col: int) -> str:
        target = self.hyperlinks.get((row, col), "")
        if target:
            return target
        anchor = self.merge_anchor.get((row, col))
        return self.hyperlinks.get(anchor, "") if anchor else ""


@dataclass(frozen=True)
class Layout:
    header_start: int
    header_end: int
    data_start: int
    title_col: Optional[int]
    document_col: Optional[int]
    company_document_col: Optional[int]
    revision_col: Optional[int]
    event_groups: Tuple[Tuple[str, Tuple[int, ...]], ...]
    metadata_cols: Tuple[int, ...]


@dataclass(frozen=True)
class SentinelCase:
    case_id: str
    source_file: str
    worksheet: str
    expected_action: str


OUTPUT_FIELDS = [
    "Case ID", "Project No.", "Source Family", "Discipline", "Category", "Subcategory",
    "Original Worksheet", "Original Section", "Classification Rule ID", "Classification Confidence",
    "Document No.", "Document Title", "Company Document No.", "Revision",
    "Event Type", "Event Date", "Event Reference", "Event Status", "Event Values JSON", "Row Metadata JSON",
    "Document Link", "Original_Hyperlink_Target", "Global_Document_Key", "Source_Document_Key",
    "Revision_Key", "Event_Key", "Document_Row_Flag", "Revision_Row_Flag", "Is_Latest_Revision",
    "Is_Latest_Event", "Source File", "Source Modified Date", "Source Sheet", "Source Row", "Source Cell",
    "Parsing Status", "Warnings",
]


def _coord(ref: str) -> Tuple[int, int]:
    m = CELL_REF_RE.match(ref or "")
    if not m:
        return 0, 0
    return int(m.group(2)), col_number(ref)


def _col_letter(n: int) -> str:
    out = ""
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def _range_coords(ref: str) -> Iterable[Tuple[int, int]]:
    parts = ref.split(":", 1)
    r1, c1 = _coord(parts[0])
    r2, c2 = _coord(parts[-1])
    if not all((r1, c1, r2, c2)):
        return
    for row in range(min(r1, r2), max(r1, r2) + 1):
        for col in range(min(c1, c2), max(c1, c2) + 1):
            yield row, col


def _formula_hyperlink_target(formula: str) -> str:
    m = re.search(r'(?i)HYPERLINK\s*\(\s*"([^"]+)"', formula or "")
    return m.group(1) if m else ""


def read_sheet_model(path: Path, root: Path, sheet_name: str) -> SheetModel:
    relative = path.resolve().relative_to(root.resolve()).as_posix()
    family = relative.split("/", 2)[1].upper() if relative.startswith("DATA/") else "UNKNOWN"
    with zipfile.ZipFile(path) as z:
        core = parse_core_properties(z)
        shared = load_shared_strings(z)
        sheet_info = next((s for s in workbook_sheet_map(z) if s["name"] == sheet_name), None)
        if not sheet_info:
            raise KeyError(f"Worksheet not found: {sheet_name}")
        xml_path = sheet_info["xml_path"]
        if xml_path not in z.namelist():
            raise KeyError(f"Worksheet XML missing: {xml_path}")
        sheet_root = ET.fromstring(z.read(xml_path))
        cells: Dict[Tuple[int, int], str] = {}
        formulas: Dict[Tuple[int, int], str] = {}
        max_row = max_col = 0
        for row in sheet_root.findall(f".//{{{MAIN_NS}}}row"):
            rn = int(row.get("r", "0") or 0)
            max_row = max(max_row, rn)
            for c in row.findall(f"{{{MAIN_NS}}}c"):
                ref = c.get("r", "")
                rr, cc = _coord(ref)
                if not rr or not cc:
                    continue
                max_col = max(max_col, cc)
                value = cell_value(c, shared).strip()
                if value:
                    cells[(rr, cc)] = value
                formula = c.find(f"{{{MAIN_NS}}}f")
                if formula is not None and formula.text:
                    formulas[(rr, cc)] = formula.text
        merge_anchor: Dict[Tuple[int, int], Tuple[int, int]] = {}
        for merge in sheet_root.findall(f".//{{{MAIN_NS}}}mergeCell"):
            ref = merge.get("ref", "")
            coords = list(_range_coords(ref))
            if not coords:
                continue
            anchor = coords[0]
            for coord in coords:
                merge_anchor[coord] = anchor
        rel_targets: Dict[str, str] = {}
        rel_path = xml_path.rsplit("/", 1)[0] + "/_rels/" + xml_path.rsplit("/", 1)[1] + ".rels"
        if rel_path in z.namelist():
            rel_root = ET.fromstring(z.read(rel_path))
            for rel in rel_root.findall(f".//{{{REL_PKG_NS}}}Relationship"):
                rel_targets[rel.get("Id", "")] = rel.get("Target", "")
        hyperlinks: Dict[Tuple[int, int], str] = {}
        for h in sheet_root.findall(f".//{{{MAIN_NS}}}hyperlink"):
            rid = h.get(f"{{{REL_DOC_NS}}}id", "")
            target = rel_targets.get(rid, h.get("location", ""))
            if target:
                for coord in _range_coords(h.get("ref", "")):
                    hyperlinks[coord] = target
        for coord, formula in formulas.items():
            target = _formula_hyperlink_target(formula)
            if target and coord not in hyperlinks:
                hyperlinks[coord] = target
        return SheetModel(
            source_file=relative,
            source_family=family,
            source_modified=str(core.get("modified", "")),
            sheet_name=sheet_name,
            max_row=max_row,
            max_col=max_col,
            cells=cells,
            formulas=formulas,
            merge_anchor=merge_anchor,
            hyperlinks=hyperlinks,
        )


def _is_doc_header(text: str) -> bool:
    n = norm_text(text)
    if not n or any(x in n for x in ("company document", "client document", "adnoc document", "adma document")):
        return False
    if "document number" in n or "drawing number" in n or "sketch number" in n or "dwg no" in n:
        return True
    return n in {"document no", "doc no", "drawing no", "dwg no", "sketch no"}


def _is_company_doc_header(text: str) -> bool:
    n = norm_text(text)
    return ("document" in n or "drawing" in n) and any(x in n for x in ("company", "client", "adnoc", "adma"))


def _is_title_header(text: str) -> bool:
    n = norm_text(text)
    return n in {"description", "title", "title/description", "document title", "description/title"} or "title/description" in n


def _is_revision_header(text: str) -> bool:
    return norm_text(text) in {"rev", "revision", "rev no", "revision no"}


def _looks_identifier(value: str) -> bool:
    value = (value or "").strip()
    if not value or not re.search(r"\d", value):
        return False
    if DOCNO_RE.search(value):
        return True
    return len(value) >= 5 and any(ch in value for ch in "-_/.") and " " not in value[:4]


def _header_candidates(model: SheetModel, predicate) -> List[Tuple[int, int]]:
    out = []
    for (row, col), value in model.cells.items():
        if row <= min(model.max_row, 30) and predicate(value):
            out.append((row, col))
    return sorted(out)


def _header_path(model: SheetModel, col: int, start: int, end: int) -> Tuple[str, ...]:
    vals: List[str] = []
    for row in range(start, end + 1):
        value = re.sub(r"\s+", " ", model.value(row, col)).strip()
        if value and (not vals or norm_text(value) != norm_text(vals[-1])):
            vals.append(value)
    return tuple(vals)


def _is_static_path(path: Sequence[str]) -> bool:
    if any(str(part).strip() == "#" for part in path):
        return True
    n = norm_text(" ".join(path))
    return bool(re.search(r"\b(sr|serial)\s+no\b|\bclass\b|\bremarks?\b", n))


def _event_field_kind(text: str) -> str:
    n = norm_text(text)
    if not n:
        return ""
    if n.startswith("issue date") or n in {"date", "actual date", "planned date", "response date", "approval date"}:
        return "date"
    if n in {"ref", "ref no", "reference", "reference no", "outgoing ref", "outgoing ref no", "transmittal", "transmittal no"}:
        return "reference"
    if n in {"code", "status", "response", "approval code", "comment code"}:
        return "status"
    return ""


def _event_value_label(path: Sequence[str], col: int) -> str:
    if not path:
        return _col_letter(col)
    if _event_field_kind(path[-1]):
        return path[-1]
    if len(path) > 1 and _event_field_kind(path[0]):
        return path[0]
    return path[-1]


def _event_group_label(path: Sequence[str], col: int) -> str:
    if not path:
        return f"COLUMN {_col_letter(col)}"
    if _event_field_kind(path[-1]) and len(path) > 1:
        group = path[:-1]
    elif len(path) > 1 and _event_field_kind(path[0]):
        group = path[1:]
    else:
        group = path
    label = " > ".join(group).strip()
    return label or path[-1]


def discover_layout(model: SheetModel) -> Tuple[Optional[Layout], List[str]]:
    warnings: List[str] = []
    doc_candidates = _header_candidates(model, _is_doc_header)
    if not doc_candidates:
        return None, ["LAYOUT_DOCUMENT_HEADER_NOT_FOUND"]
    header_start, document_col = doc_candidates[0]
    band_end = min(model.max_row, header_start + 4)
    title_candidates = [(r, c) for (r, c), v in model.cells.items() if header_start <= r <= band_end and _is_title_header(v)]
    company_candidates = [(r, c) for (r, c), v in model.cells.items() if header_start <= r <= band_end and _is_company_doc_header(v)]
    revision_candidates = [(r, c) for (r, c), v in model.cells.items() if header_start <= r <= band_end and _is_revision_header(v)]
    title_col = sorted(title_candidates)[0][1] if title_candidates else None
    company_col = sorted(company_candidates)[0][1] if company_candidates else None
    revision_col = sorted(revision_candidates)[0][1] if revision_candidates else None
    if title_col is None:
        warnings.append("LAYOUT_TITLE_HEADER_NOT_FOUND")
    if revision_col is None:
        warnings.append("LAYOUT_REVISION_HEADER_NOT_FOUND")

    data_start = None
    for row in range(header_start + 1, model.max_row + 1):
        doc = model.value(row, document_col)
        company = model.value(row, company_col) if company_col else ""
        title = model.value(row, title_col) if title_col else ""
        rev = model.value(row, revision_col) if revision_col else ""
        if _looks_identifier(doc) or _looks_identifier(company):
            if title or rev or doc or company:
                data_start = row
                break
    if data_start is None:
        return None, warnings + ["LAYOUT_FIRST_DATA_ROW_NOT_FOUND"]
    header_end = data_start - 1
    identity_cols = {c for c in (title_col, document_col, company_col, revision_col) if c}
    metadata_cols: List[int] = []
    group_cols: List[Tuple[str, List[int]]] = []
    for col in range(1, model.max_col + 1):
        if col in identity_cols:
            continue
        path = _header_path(model, col, header_start, header_end)
        if _is_static_path(path):
            metadata_cols.append(col)
            continue
        if not path:
            continue
        label = _event_group_label(path, col)
        if group_cols and group_cols[-1][0] == label and group_cols[-1][1][-1] + 1 == col:
            group_cols[-1][1].append(col)
        else:
            group_cols.append((label, [col]))
    event_groups = tuple((label, tuple(cols)) for label, cols in group_cols)
    if not event_groups:
        warnings.append("LAYOUT_EVENT_GROUPS_NOT_FOUND")
    return Layout(header_start, header_end, data_start, title_col, document_col, company_col, revision_col, event_groups, tuple(metadata_cols)), warnings


def _active_section(model: SheetModel, row: int) -> str:
    if not re.search(r"(?i)in+com+ing.*(?:doc|drg|drawing)", model.sheet_name):
        return ""
    markers: List[Tuple[int, str]] = []
    for (rr, _cc), value in model.cells.items():
        n = norm_text(value)
        if n in {"document", "documents"}:
            markers.append((rr, "DOCUMENTS"))
        elif n in {"drawing", "drawings"}:
            markers.append((rr, "DRAWINGS"))
    prior = [item for item in markers if item[0] <= row]
    return max(prior, default=(0, ""), key=lambda x: x[0])[1]


def _normalize_date(value: str) -> Tuple[str, str]:
    raw = (value or "").strip()
    if not raw:
        return "", ""
    try:
        n = float(raw)
        if n.is_integer() and 1 <= n <= 80000:
            dt = datetime(1899, 12, 30) + timedelta(days=int(n))
            return dt.date().isoformat(), ""
    except ValueError:
        pass
    for fmt in ("%d-%b-%Y", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(raw, fmt).date().isoformat(), ""
        except ValueError:
            continue
    return raw, "DATE_TEXT_PRESERVED"


def _stable_key(prefix: str, *parts: object) -> str:
    payload = "|".join(norm_text(part) for part in parts)
    return f"{prefix}-" + hashlib.sha256(payload.encode("utf-8")).hexdigest()[:24]


def _event_values(model: SheetModel, row: int, cols: Sequence[int], layout: Layout) -> Tuple[Dict[str, str], str, str, str, List[str]]:
    values: Dict[str, str] = {}
    date_candidates: List[Tuple[int, int, str]] = []
    refs: List[str] = []
    statuses: List[str] = []
    warnings: List[str] = []
    for col in cols:
        raw = model.value(row, col).strip()
        if not raw:
            continue
        path = _header_path(model, col, layout.header_start, layout.header_end)
        field_label = _event_value_label(path, col)
        key = field_label if field_label not in values else f"{field_label}@{_col_letter(col)}"
        field_norm = norm_text(field_label)
        kind = _event_field_kind(field_label)
        if kind == "date":
            normalized, warning = _normalize_date(raw)
            values[key] = normalized
            if normalized and not warning and re.fullmatch(r"\d{4}-\d{2}-\d{2}", normalized):
                priority = 40 if "actual" in field_norm else 30 if any(x in field_norm for x in ("response", "approval")) else 10 if "planned" in field_norm else 20
                date_candidates.append((priority, col, normalized))
            if warning:
                warnings.append(warning)
        else:
            values[key] = raw
        if kind == "reference" or re.fullmatch(r"(?:outgoing )?(?:ref|reference)(?: no)?", field_norm):
            refs.append(raw)
        if kind == "status":
            statuses.append(raw)
    event_date = max(date_candidates, default=(0, 0, ""), key=lambda item: (item[0], item[1]))[2]
    return values, event_date, "; ".join(refs), "; ".join(statuses), sorted(set(warnings))


def _row_metadata(model: SheetModel, row: int, layout: Layout) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for col in layout.metadata_cols:
        raw = model.value(row, col).strip()
        if not raw:
            continue
        path = _header_path(model, col, layout.header_start, layout.header_end)
        key = " > ".join(path) if path else _col_letter(col)
        if key in out:
            key = f"{key}@{_col_letter(col)}"
        out[key] = raw
    return out


def _document_link(model: SheetModel, row: int, layout: Layout) -> str:
    priority = [layout.document_col, layout.company_document_col, layout.title_col]
    for col in priority:
        if col:
            target = model.hyperlink(row, col)
            if target:
                return target
    for col in range(1, model.max_col + 1):
        target = model.hyperlink(row, col)
        if target:
            return target
    return ""


def extract_model(case: SentinelCase, model: SheetModel, rules: Sequence[Rule]) -> Tuple[List[Dict[str, object]], Dict[str, object]]:
    structural = apply_classification(rules, model.source_family, {"FILE": model.source_file, "WORKSHEET": model.sheet_name})
    if structural["status"] == "EXCLUDED":
        return [], {
            "case_id": case.case_id, "source_file": model.source_file, "worksheet": model.sheet_name,
            "status": "EXCLUDED", "reason": structural["notes"], "rows_with_identity": 0,
            "event_records": 0, "distinct_documents": 0, "distinct_revisions": 0, "warnings": [],
            "hyperlink_targets_preserved": 0,
        }
    layout, layout_warnings = discover_layout(model)
    if layout is None:
        return [], {
            "case_id": case.case_id, "source_file": model.source_file, "worksheet": model.sheet_name,
            "status": "REVIEW_REQUIRED", "reason": "Unable to establish safe worksheet layout",
            "rows_with_identity": 0, "event_records": 0, "distinct_documents": 0, "distinct_revisions": 0,
            "warnings": layout_warnings, "hyperlink_targets_preserved": 0,
        }

    project = ";".join(infer_path_projects(Path(model.source_file).name))
    records: List[Dict[str, object]] = []
    source_rows_with_identity = 0
    seen_docs = set()
    seen_revisions = set()
    for row in range(layout.data_start, model.max_row + 1):
        title = model.value(row, layout.title_col).strip() if layout.title_col else ""
        document = model.value(row, layout.document_col).strip() if layout.document_col else ""
        company = model.value(row, layout.company_document_col).strip() if layout.company_document_col else ""
        revision = model.value(row, layout.revision_col).strip() if layout.revision_col else ""
        if not document and _looks_identifier(company):
            document = company
            row_warnings = ["DOCUMENT_NO_FALLBACK_COMPANY"]
        else:
            row_warnings = []
        if not document and not company and not title:
            continue
        if not (_looks_identifier(document) or _looks_identifier(company)):
            continue
        source_rows_with_identity += 1
        section = _active_section(model, row)
        evidence = {
            "FILE": model.source_file, "WORKSHEET": model.sheet_name, "SECTION": section,
            "HEADER": " | ".join(" > ".join(_header_path(model, c, layout.header_start, layout.header_end)) for c in range(1, model.max_col + 1)),
            "DOC_NUMBER": document, "TITLE": title,
        }
        classification = apply_classification(rules, model.source_family, evidence)
        if classification["status"] != "INCLUDE":
            row_warnings.append("ROW_CLASSIFICATION_REVIEW_REQUIRED")
        source_doc_key = _stable_key("SDOC", model.source_file, model.sheet_name, section, document)
        global_doc_key = _stable_key("GDOC", model.source_family, project, document)
        revision_key = _stable_key("REV", source_doc_key, revision or "[BLANK]")
        seen_docs.add(source_doc_key)
        seen_revisions.add(revision_key)
        metadata = _row_metadata(model, row, layout)
        link = _document_link(model, row, layout)
        emitted = 0
        for event_index, (event_type, cols) in enumerate(layout.event_groups, start=1):
            values, event_date, event_ref, event_status, value_warnings = _event_values(model, row, cols, layout)
            if not values:
                continue
            emitted += 1
            event_key = _stable_key("EVT", revision_key, row, event_index, event_type, json.dumps(values, sort_keys=True, ensure_ascii=False))
            warnings = sorted(set(row_warnings + layout_warnings + value_warnings))
            records.append({
                "Case ID": case.case_id, "Project No.": project, "Source Family": model.source_family,
                "Discipline": classification["discipline"], "Category": classification["category"], "Subcategory": classification["subcategory"],
                "Original Worksheet": model.sheet_name, "Original Section": section,
                "Classification Rule ID": ";".join(classification["rule_ids"]),
                "Classification Confidence": f"{float(classification['confidence']):.2f}",
                "Document No.": document, "Document Title": title, "Company Document No.": company, "Revision": revision,
                "Event Type": event_type, "Event Date": event_date, "Event Reference": event_ref, "Event Status": event_status,
                "Event Values JSON": json.dumps(values, sort_keys=True, ensure_ascii=False, separators=(",", ":")),
                "Row Metadata JSON": json.dumps(metadata, sort_keys=True, ensure_ascii=False, separators=(",", ":")),
                "Document Link": link, "Original_Hyperlink_Target": link,
                "Global_Document_Key": global_doc_key, "Source_Document_Key": source_doc_key,
                "Revision_Key": revision_key, "Event_Key": event_key,
                "Document_Row_Flag": 0, "Revision_Row_Flag": 0, "Is_Latest_Revision": 0, "Is_Latest_Event": 0,
                "Source File": model.source_file, "Source Modified Date": model.source_modified, "Source Sheet": model.sheet_name,
                "Source Row": row, "Source Cell": f"{_col_letter(layout.document_col or layout.company_document_col or layout.title_col or 1)}{row}",
                "Parsing Status": "INCLUDE" if classification["status"] == "INCLUDE" else "REVIEW_REQUIRED",
                "Warnings": ";".join(warnings), "_event_order": event_index,
            })
        if not emitted:
            event_key = _stable_key("EVT", revision_key, row, "RECORD")
            warnings = sorted(set(row_warnings + layout_warnings))
            records.append({
                "Case ID": case.case_id, "Project No.": project, "Source Family": model.source_family,
                "Discipline": classification["discipline"], "Category": classification["category"], "Subcategory": classification["subcategory"],
                "Original Worksheet": model.sheet_name, "Original Section": section,
                "Classification Rule ID": ";".join(classification["rule_ids"]),
                "Classification Confidence": f"{float(classification['confidence']):.2f}",
                "Document No.": document, "Document Title": title, "Company Document No.": company, "Revision": revision,
                "Event Type": "RECORD", "Event Date": "", "Event Reference": "", "Event Status": "",
                "Event Values JSON": "{}", "Row Metadata JSON": json.dumps(metadata, sort_keys=True, ensure_ascii=False, separators=(",", ":")),
                "Document Link": link, "Original_Hyperlink_Target": link,
                "Global_Document_Key": global_doc_key, "Source_Document_Key": source_doc_key,
                "Revision_Key": revision_key, "Event_Key": event_key,
                "Document_Row_Flag": 0, "Revision_Row_Flag": 0, "Is_Latest_Revision": 0, "Is_Latest_Event": 0,
                "Source File": model.source_file, "Source Modified Date": model.source_modified, "Source Sheet": model.sheet_name,
                "Source Row": row, "Source Cell": f"{_col_letter(layout.document_col or layout.company_document_col or layout.title_col or 1)}{row}",
                "Parsing Status": "INCLUDE" if classification["status"] == "INCLUDE" else "REVIEW_REQUIRED",
                "Warnings": ";".join(warnings), "_event_order": 0,
            })

    records.sort(key=lambda r: (int(r["Source Row"]), int(r["_event_order"]), str(r["Event_Key"])))
    first_doc: set[str] = set()
    first_rev: set[str] = set()
    last_revision_by_doc: Dict[str, str] = {}
    last_event_by_rev: Dict[str, str] = {}
    for rec in records:
        last_revision_by_doc[str(rec["Source_Document_Key"])] = str(rec["Revision_Key"])
        last_event_by_rev[str(rec["Revision_Key"])] = str(rec["Event_Key"])
    for rec in records:
        doc_key = str(rec["Source_Document_Key"])
        rev_key = str(rec["Revision_Key"])
        if doc_key not in first_doc:
            rec["Document_Row_Flag"] = 1; first_doc.add(doc_key)
        if rev_key not in first_rev:
            rec["Revision_Row_Flag"] = 1; first_rev.add(rev_key)
        rec["Is_Latest_Revision"] = int(last_revision_by_doc.get(doc_key) == rev_key)
        rec["Is_Latest_Event"] = int(last_event_by_rev.get(rev_key) == rec["Event_Key"])
        rec.pop("_event_order", None)
    warning_set = sorted({w for rec in records for w in str(rec["Warnings"]).split(";") if w})
    return records, {
        "case_id": case.case_id, "source_file": model.source_file, "worksheet": model.sheet_name,
        "status": "INCLUDE", "reason": "", "rows_with_identity": source_rows_with_identity,
        "event_records": len(records), "distinct_documents": len(seen_docs), "distinct_revisions": len(seen_revisions),
        "warnings": warning_set, "hyperlink_targets_preserved": sum(1 for rec in records if rec["Document Link"]),
        "header_start": layout.header_start, "header_end": layout.header_end, "data_start": layout.data_start,
    }


def load_sentinel_cases(path: Path) -> List[SentinelCase]:
    with path.open(newline="", encoding="utf-8-sig") as f:
        return [SentinelCase(row["Case_ID"].strip(), row["Source_File"].strip(), row["Worksheet"].strip(), row["Expected_Action"].strip().upper()) for row in csv.DictReader(f)]


def run_sentinels(root: Path, cases: Sequence[SentinelCase], rules: Sequence[Rule]) -> Tuple[List[Dict[str, object]], List[Dict[str, object]]]:
    all_records: List[Dict[str, object]] = []
    reconciliation: List[Dict[str, object]] = []
    for case in cases:
        source = root / case.source_file
        if not source.exists():
            reconciliation.append({"case_id": case.case_id, "source_file": case.source_file, "worksheet": case.worksheet, "status": "REVIEW_REQUIRED", "reason": "SOURCE_FILE_NOT_FOUND", "warnings": ["SOURCE_FILE_NOT_FOUND"], "rows_with_identity": 0, "event_records": 0, "distinct_documents": 0, "distinct_revisions": 0, "hyperlink_targets_preserved": 0})
            continue
        try:
            model = read_sheet_model(source, root, case.worksheet)
            records, recon = extract_model(case, model, rules)
        except (KeyError, OSError, zipfile.BadZipFile, ET.ParseError) as exc:
            records, recon = [], {"case_id": case.case_id, "source_file": case.source_file, "worksheet": case.worksheet, "status": "REVIEW_REQUIRED", "reason": f"EXTRACTION_FAILED:{exc.__class__.__name__}", "warnings": [f"EXTRACTION_FAILED:{exc.__class__.__name__}"], "rows_with_identity": 0, "event_records": 0, "distinct_documents": 0, "distinct_revisions": 0, "hyperlink_targets_preserved": 0}
        all_records.extend(records)
        reconciliation.append(recon)
    all_records.sort(key=lambda r: (str(r["Case ID"]), str(r["Source File"]), str(r["Source Sheet"]), int(r["Source Row"]), str(r["Event_Key"])))
    return all_records, reconciliation


def write_cycle2_outputs(records: Sequence[Mapping[str, object]], reconciliation: Sequence[Mapping[str, object]], output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    csv_path = output_dir / "sentinel_records.csv"
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_FIELDS, extrasaction="ignore", lineterminator="\n")
        writer.writeheader()
        for row in records:
            writer.writerow({key: row.get(key, "") for key in OUTPUT_FIELDS})
    recon_text = json.dumps(list(reconciliation), indent=2, ensure_ascii=False, sort_keys=True) + "\n"
    with (output_dir / "sentinel_reconciliation.json").open("w", encoding="utf-8", newline="\n") as f:
        f.write(recon_text)
    lines = ["# Cycle 2 Sentinel Extraction Report", "", f"- Sentinel cases: **{len(reconciliation)}**", f"- Event records: **{len(records)}**", f"- Distinct documents: **{len({r.get('Source_Document_Key') for r in records})}**", f"- Distinct revisions: **{len({r.get('Revision_Key') for r in records})}**", f"- Records with hyperlinks: **{sum(1 for r in records if r.get('Document Link'))}**", "", "## Case reconciliation", ""]
    for item in reconciliation:
        warnings = ", ".join(item.get("warnings", [])) or "none"
        lines.append(f"- `{item.get('case_id')}` — {item.get('status')}; rows={item.get('rows_with_identity', 0)}; events={item.get('event_records', 0)}; docs={item.get('distinct_documents', 0)}; revisions={item.get('distinct_revisions', 0)}; warnings={warnings}")
    lines += ["", "## Safety", "", "Cycle 2 is sentinel-only. `DATA/` is read-only. No full-source extraction and no final XLSX are produced.", ""]
    with (output_dir / "sentinel_report.md").open("w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines))
