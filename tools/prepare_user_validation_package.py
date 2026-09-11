from __future__ import annotations

import csv
import json
import shutil
from collections import defaultdict
from pathlib import Path


PILOT_PROJECTS = ("2035", "2631", "2705", "2745")

CONFLICT_FILES = (
    "DATA/TECH/2035 DOCUMENT REGISTER.xlsx",
    "DATA/TECH/2035 Offshore Construction Engineering Register (Pipeline & Cables).xlsx",
    "DATA/TECH/2705 -DOCUMENT REGISTER.xlsx",
    "DATA/TECH/2745-PP-GE-001-MDR Rev_2.xlsx",
)

CONFLICTS = (
    {
        "Source File": "DATA/TECH/2035 DOCUMENT REGISTER.xlsx",
        "Worksheet": "Specification",
        "Filename/Path Project": "2035",
        "Internal Project Evidence": "2136",
        "Document Prefix Evidence": "2136-N-SP-0314",
        "Evidence Strength": "STRONG",
        "Validation Status": "USER_VALIDATION_REQUIRED",
        "Notes": "Filename/path indicates 2035; explicit internal project and document prefix indicate 2136.",
    },
    {
        "Source File": "DATA/TECH/2035 Offshore Construction Engineering Register (Pipeline & Cables).xlsx",
        "Worksheet": "Specification",
        "Filename/Path Project": "2035",
        "Internal Project Evidence": "2136",
        "Document Prefix Evidence": "2136-N-SP-0314",
        "Evidence Strength": "STRONG",
        "Validation Status": "USER_VALIDATION_REQUIRED",
        "Notes": "Filename/path indicates 2035; explicit internal project and document prefix indicate 2136.",
    },
    {
        "Source File": "DATA/TECH/2705 -DOCUMENT REGISTER.xlsx",
        "Worksheet": "Documents - Naval Marine",
        "Filename/Path Project": "2705",
        "Internal Project Evidence": "2824",
        "Document Prefix Evidence": "",
        "Evidence Strength": "MEDIUM",
        "Validation Status": "USER_VALIDATION_REQUIRED",
        "Notes": "Workbook filename indicates 2705; explicit internal project label indicates 2824. This worksheet currently has no extracted document identities.",
    },
    {
        "Source File": "DATA/TECH/2705 -DOCUMENT REGISTER.xlsx",
        "Worksheet": "Drawings",
        "Filename/Path Project": "2705",
        "Internal Project Evidence": "2824",
        "Document Prefix Evidence": "2824-NN-0001",
        "Evidence Strength": "STRONG",
        "Validation Status": "USER_VALIDATION_REQUIRED",
        "Notes": "Workbook filename indicates 2705; internal project and document prefix indicate 2824.",
    },
    {
        "Source File": "DATA/TECH/2745-PP-GE-001-MDR Rev_2.xlsx",
        "Worksheet": "Anchor Pattern",
        "Filename/Path Project": "2745",
        "Internal Project Evidence": "8405",
        "Document Prefix Evidence": "8405-AP-0201",
        "Evidence Strength": "STRONG",
        "Validation Status": "USER_VALIDATION_REQUIRED",
        "Notes": "Filename indicates 2745; Contractor Project No. is 8405 and document numbers use 8405 prefix.",
    },
    {
        "Source File": "DATA/TECH/2745-PP-GE-001-MDR Rev_2.xlsx",
        "Worksheet": "DP",
        "Filename/Path Project": "2745",
        "Internal Project Evidence": "8405",
        "Document Prefix Evidence": "8405-DP-0201",
        "Evidence Strength": "STRONG",
        "Validation Status": "USER_VALIDATION_REQUIRED",
        "Notes": "Filename indicates 2745; Contractor Project No. is 8405 and document numbers use 8405 prefix.",
    },
    {
        "Source File": "DATA/TECH/2745-PP-GE-001-MDR Rev_2.xlsx",
        "Worksheet": "DRAWINGS",
        "Filename/Path Project": "2745",
        "Internal Project Evidence": "8405",
        "Document Prefix Evidence": "8405-NN-0001",
        "Evidence Strength": "STRONG",
        "Validation Status": "USER_VALIDATION_REQUIRED",
        "Notes": "Filename indicates 2745; internal project and document numbers indicate 8405.",
    },
    {
        "Source File": "DATA/TECH/2745-PP-GE-001-MDR Rev_2.xlsx",
        "Worksheet": "Naval & Marine doc",
        "Filename/Path Project": "2745",
        "Internal Project Evidence": "8405",
        "Document Prefix Evidence": "8405-ZZ-CL-001",
        "Evidence Strength": "STRONG",
        "Validation Status": "USER_VALIDATION_REQUIRED",
        "Notes": "Filename indicates 2745; internal project and document numbers indicate 8405.",
    },
    {
        "Source File": "DATA/TECH/2745-PP-GE-001-MDR Rev_2.xlsx",
        "Worksheet": "PIPELINE & CABLE doc",
        "Filename/Path Project": "2745",
        "Internal Project Evidence": "8405",
        "Document Prefix Evidence": "8405-NN-RP-001",
        "Evidence Strength": "STRONG",
        "Validation Status": "USER_VALIDATION_REQUIRED",
        "Notes": "Filename indicates 2745; internal project and document numbers indicate 8405.",
    },
)


def read_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with path.open(newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        return list(reader.fieldnames or []), [{k: (v or "") for k, v in row.items()} for row in reader]


def write_csv(path: Path, fields: list[str], rows: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore", lineterminator="\n")
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in fields})


def first_by_key(rows: list[dict[str, str]], key_field: str) -> list[dict[str, str]]:
    seen: set[str] = set()
    result: list[dict[str, str]] = []
    for row in rows:
        key = row.get(key_field, "")
        if not key or key in seen:
            continue
        seen.add(key)
        result.append(row)
    return result


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    review_root = root / "review"
    conflicts_dir = review_root / "project_conflicts"
    conflicts_files_dir = conflicts_dir / "files"
    pilot_dir = review_root / "pilot_4_projects"
    conflicts_files_dir.mkdir(parents=True, exist_ok=True)
    pilot_dir.mkdir(parents=True, exist_ok=True)

    # Copy exact source workbooks for human inspection. Originals under DATA are never modified.
    copied: list[dict[str, str]] = []
    for relative in CONFLICT_FILES:
        source = root / relative
        if not source.exists():
            raise FileNotFoundError(relative)
        destination = conflicts_files_dir / source.name
        shutil.copy2(source, destination)
        copied.append({
            "Original Source File": relative,
            "Review Copy": destination.relative_to(root).as_posix(),
        })

    conflict_fields = [
        "Source File", "Worksheet", "Filename/Path Project", "Internal Project Evidence",
        "Document Prefix Evidence", "Evidence Strength", "Validation Status", "Notes",
    ]
    write_csv(conflicts_dir / "conflict_manifest.csv", conflict_fields, [dict(row) for row in CONFLICTS])
    write_csv(conflicts_dir / "copied_files.csv", ["Original Source File", "Review Copy"], copied)

    conflict_readme = """# Project conflict validation package

This folder is for human validation only.

- The Excel files under `files/` are exact copies of the original source workbooks.
- The originals remain in `DATA/` and are not modified, renamed, or moved.
- `conflict_manifest.csv` records the project-number evidence that conflicts with the filename/path project.
- No project ownership has been automatically corrected. Every item remains `USER_VALIDATION_REQUIRED` until the owner confirms it.

Project identity evidence should be considered separately:
1. source filename/path project,
2. explicit project number inside the workbook,
3. document-number project prefix.

If the internal project number and document-number prefix agree against the filename, the source file may simply be misnamed; this package exists so that can be confirmed before the master index is produced.
"""
    (conflicts_dir / "README.md").write_text(conflict_readme, encoding="utf-8", newline="\n")

    fields, all_rows = read_csv(root / "outputs" / "cycle3" / "full_records.csv")
    pilot_rows = [row for row in all_rows if row.get("Project No.", "").strip() in PILOT_PROJECTS]
    pilot_rows.sort(key=lambda r: (
        r.get("Project No.", ""), r.get("Source File", "").casefold(), r.get("Source Sheet", "").casefold(),
        int(r.get("Source Row", "0") or 0), r.get("Event_Key", ""),
    ))
    write_csv(pilot_dir / "pilot_records.csv", fields, pilot_rows)

    document_fields = [
        "Project No.", "Source Family", "Discipline", "Category", "Subcategory",
        "Document No.", "Document Title", "Company Document No.", "Document Link",
        "Source File", "Source Sheet", "Source Row", "Parsing Status", "Warnings",
        "Source_Document_Key", "Global_Document_Key",
    ]
    document_rows = first_by_key(pilot_rows, "Source_Document_Key")
    write_csv(pilot_dir / "pilot_documents.csv", document_fields, document_rows)

    revision_fields = [
        "Project No.", "Source Family", "Discipline", "Category", "Subcategory",
        "Document No.", "Document Title", "Revision", "Event Date", "Event Status",
        "Source File", "Source Sheet", "Source Row", "Parsing Status", "Warnings",
        "Revision_Key", "Source_Document_Key",
    ]
    revision_rows = first_by_key(pilot_rows, "Revision_Key")
    write_csv(pilot_dir / "pilot_revisions.csv", revision_fields, revision_rows)

    by_project: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in pilot_rows:
        by_project[row.get("Project No.", "")].append(row)
    summary_rows: list[dict[str, object]] = []
    conflict_sources = {row["Source File"] for row in CONFLICTS}
    for project in PILOT_PROJECTS:
        rows = by_project.get(project, [])
        summary_rows.append({
            "Project": project,
            "Event Records": len(rows),
            "Source Documents": len({r.get("Source_Document_Key", "") for r in rows if r.get("Source_Document_Key", "")}),
            "Global Documents": len({r.get("Global_Document_Key", "") for r in rows if r.get("Global_Document_Key", "")}),
            "Revisions": len({r.get("Revision_Key", "") for r in rows if r.get("Revision_Key", "")}),
            "Source Workbooks": len({r.get("Source File", "") for r in rows if r.get("Source File", "")}),
            "Worksheets": len({(r.get("Source File", ""), r.get("Source Sheet", "")) for r in rows if r.get("Source Sheet", "")}),
            "Hyperlinks": sum(1 for r in rows if r.get("Document Link", "")),
            "Rows Requiring Review": sum(1 for r in rows if r.get("Parsing Status", "").upper() != "INCLUDE"),
            "Contains Known Project Conflict Source": "YES" if any(r.get("Source File", "") in conflict_sources for r in rows) else "NO",
        })
    summary_fields = [
        "Project", "Event Records", "Source Documents", "Global Documents", "Revisions", "Source Workbooks",
        "Worksheets", "Hyperlinks", "Rows Requiring Review", "Contains Known Project Conflict Source",
    ]
    write_csv(pilot_dir / "pilot_summary.csv", summary_fields, summary_rows)

    metadata = {
        "pilot_projects": list(PILOT_PROJECTS),
        "source_dataset": "outputs/cycle3/full_records.csv",
        "event_records": len(pilot_rows),
        "source_documents": len(document_rows),
        "revisions": len(revision_rows),
        "conflict_workbooks_copied": len(CONFLICT_FILES),
        "conflict_worksheets": len(CONFLICTS),
        "project_assignment_note": "Pilot records show the current Cycle 3 extracted Project No. assignment. Conflicting source identities are not auto-corrected; see review/project_conflicts/conflict_manifest.csv.",
    }
    (pilot_dir / "pilot_metadata.json").write_text(json.dumps(metadata, indent=2, sort_keys=True) + "\n", encoding="utf-8", newline="\n")

    pilot_readme = """# Four-project pilot validation data

Pilot projects: **2035, 2631, 2705, 2745**.

This is a validation snapshot of the current Cycle 3 extraction, not the final master index.

Files:
- `pilot_summary.csv` — project-level counts.
- `pilot_documents.csv` — one representative row per extracted source document.
- `pilot_revisions.csv` — one representative row per extracted revision.
- `pilot_records.csv` — all extracted event/transaction records for the four pilot projects.
- `pilot_metadata.json` — scope and provenance.

Important: current extracted `Project No.` values are intentionally left unchanged. Project-number conflicts are shown separately under `review/project_conflicts/` for owner validation before any correction is made.
"""
    (pilot_dir / "README.md").write_text(pilot_readme, encoding="utf-8", newline="\n")

    print(f"Prepared {len(CONFLICT_FILES)} conflict workbook copies and {len(CONFLICTS)} conflict worksheet entries")
    print(f"Prepared pilot projects {', '.join(PILOT_PROJECTS)}: {len(pilot_rows)} event rows, {len(document_rows)} documents, {len(revision_rows)} revisions")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
