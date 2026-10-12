from __future__ import annotations

import csv
import os
import uuid
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, MutableMapping, Sequence, Set


FIELDS = ["Relative_Path", "Enabled", "Reason", "Owner_Note"]


def _normalize(value: str) -> str:
    return str(value or "").replace("\\", "/").lstrip("/").casefold()


def _as_checked(value: Any) -> bool:
    return str(value or "").strip().casefold() in {"true", "yes", "1", "on", "checked", "include", "included"}


def read_source_exclusions(path: Path) -> Dict[str, Dict[str, str]]:
    path = Path(path)
    if not path.exists():
        return {}
    rows: Dict[str, Dict[str, str]] = {}
    with path.open(newline="", encoding="utf-8-sig") as handle:
        for raw in csv.DictReader(handle):
            rel = str(raw.get("Relative_Path", "") or "").replace("\\", "/").lstrip("/")
            enabled = str(raw.get("Enabled", "YES") or "YES").strip().upper()
            if not rel or enabled != "YES":
                continue
            rows[_normalize(rel)] = {
                "Relative_Path": rel,
                "Enabled": "YES",
                "Reason": str(raw.get("Reason", "") or "Owner excluded source from Pending Update").strip(),
                "Owner_Note": str(raw.get("Owner_Note", "") or "").strip(),
            }
    return rows


def read_source_selection_notes(path: Path) -> Dict[str, str]:
    """Read owner notes independently of include/exclude state.

    Older configuration files used ``Reason`` as the only note field for an
    excluded source, so retain that value as a compatibility fallback.
    """
    return {
        key: row["Owner_Note"]
        for key, row in _read_source_selection_note_rows(path).items()
        if row["Owner_Note"]
    }


def _read_source_selection_note_rows(path: Path) -> Dict[str, Dict[str, str]]:
    path = Path(path)
    if not path.exists():
        return {}
    rows: Dict[str, Dict[str, str]] = {}
    with path.open(newline="", encoding="utf-8-sig") as handle:
        for raw in csv.DictReader(handle):
            rel = str(raw.get("Relative_Path", "") or "").replace("\\", "/").lstrip("/")
            if not rel:
                continue
            note = raw.get("Owner_Note")
            if note is None:
                note = raw.get("Reason", "")
            rows[_normalize(rel)] = {"Relative_Path": rel, "Owner_Note": str(note or "").strip()}
    return rows


def apply_source_exclusions(workbooks: Sequence[MutableMapping[str, Any]], path: Path) -> Set[str]:
    exclusions = read_source_exclusions(path)
    applied: Set[str] = set()
    for workbook in workbooks:
        relative = str(workbook.get("relative_path", "") or "").replace("\\", "/").lstrip("/")
        decision = exclusions.get(_normalize(relative))
        if not decision:
            continue
        workbook["selected_excluded_status"] = "EXCLUDED"
        workbook["selection_exclusion_reason"] = "Owner excluded source: " + decision.get("Reason", "")
        warnings = list(workbook.get("warnings", []) or [])
        if "OWNER_EXCLUDED_SOURCE" not in warnings:
            warnings.append("OWNER_EXCLUDED_SOURCE")
        workbook["warnings"] = warnings
        applied.add(relative)
    return applied


def _write_rows(path: Path, rows: Iterable[Mapping[str, str]]) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(path.name + f".tmp-{uuid.uuid4().hex[:8]}")
    with temp.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=FIELDS, lineterminator="\n")
        writer.writeheader()
        for row in rows:
            writer.writerow({field: str(row.get(field, "") or "") for field in FIELDS})
    os.replace(temp, path)


def _write_selection_state(
    path: Path,
    exclusions: Mapping[str, Mapping[str, str]],
    notes: Mapping[str, str],
    note_paths: Mapping[str, str],
) -> None:
    keys = set(exclusions) | {key for key, note in notes.items() if note}
    rows = []
    for key in sorted(keys):
        exclusion = exclusions.get(key, {})
        rows.append({
            "Relative_Path": str(exclusion.get("Relative_Path", note_paths.get(key, key))).replace("\\", "/"),
            "Enabled": "YES" if exclusion else "NO",
            "Reason": str(exclusion.get("Reason", "")),
            "Owner_Note": str(notes.get(key, "")),
        })
    _write_rows(path, rows)


def set_source_selection(config_dir: Path, source_file: str, *, include: bool, reason: str = "") -> Dict[str, Any]:
    config_dir = Path(config_dir)
    path = config_dir / "source_exclusions.csv"
    current = read_source_exclusions(path)
    note_rows = _read_source_selection_note_rows(path)
    notes = {key: row["Owner_Note"] for key, row in note_rows.items() if row["Owner_Note"]}
    note_paths = {key: row["Relative_Path"] for key, row in note_rows.items()}
    normalized = _normalize(source_file)
    clean_source = str(source_file or "").replace("\\", "/").lstrip("/")
    if not clean_source:
        raise ValueError("A source file must be selected.")
    note_paths[normalized] = clean_source

    if include:
        current.pop(normalized, None)
        decision = "INCLUDE"
    else:
        current[normalized] = {
            "Relative_Path": clean_source,
            "Enabled": "YES",
            "Reason": str(reason or "Owner excluded source from Pending Update").strip(),
        }
        if reason:
            notes[normalized] = str(reason).strip()
        decision = "EXCLUDE"

    if normalized not in notes and reason:
        notes[normalized] = str(reason).strip()
    _write_selection_state(path, current, notes, note_paths)
    return {
        "decision": decision,
        "source_file": clean_source,
        "excluded_sources": len(current),
        "config_file": str(path),
    }


def set_source_selections_from_file(config_dir: Path, selections_file: Path) -> Dict[str, Any]:
    """Persist all owner source checkboxes in one atomic configuration update.

    Expected CSV columns are ``Source File``, ``Include in Index?`` and optional
    ``Owner Note``. Checked rows remove any owner exclusion; unchecked rows add
    one. The complete file is processed in one engine launch so selecting many
    sources does not start many scans or repeatedly rewrite the configuration.
    """
    selections_file = Path(selections_file)
    if not selections_file.exists():
        raise FileNotFoundError(f"Source-selection decisions file not found: {selections_file}")

    config_dir = Path(config_dir)
    target = config_dir / "source_exclusions.csv"
    current = read_source_exclusions(target)
    note_rows = _read_source_selection_note_rows(target)
    notes = {key: row["Owner_Note"] for key, row in note_rows.items() if row["Owner_Note"]}
    note_paths = {key: row["Relative_Path"] for key, row in note_rows.items()}
    processed = 0
    included = 0
    excluded = 0

    with selections_file.open(newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        required = {"Source File", "Include in Index?"}
        missing = required.difference(reader.fieldnames or [])
        if missing:
            raise ValueError("Source-selection decisions are missing required column(s): " + ", ".join(sorted(missing)))

        for raw in reader:
            clean_source = str(raw.get("Source File", "") or "").replace("\\", "/").lstrip("/").strip()
            if not clean_source:
                continue
            normalized = _normalize(clean_source)
            note_paths[normalized] = clean_source
            notes[normalized] = str(raw.get("Owner Note", "") or "").strip()
            if _as_checked(raw.get("Include in Index?", "")):
                current.pop(normalized, None)
                included += 1
            else:
                note = notes[normalized]
                current[normalized] = {
                    "Relative_Path": clean_source,
                    "Enabled": "YES",
                    "Reason": note or "Owner excluded using Source Selection checkbox",
                }
                excluded += 1
            processed += 1

    _write_selection_state(target, current, notes, note_paths)
    return {
        "decision": "BATCH_SOURCE_SELECTION_SAVED",
        "processed_sources": processed,
        "checked_included": included,
        "unchecked_excluded": excluded,
        "excluded_sources": len(current),
        "config_file": str(target),
    }
