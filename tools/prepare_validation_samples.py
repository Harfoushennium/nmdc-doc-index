from __future__ import annotations

import csv
from collections import defaultdict
from pathlib import Path


PROJECTS = ("2035", "2631", "2705", "2745")
DOCS_PER_PROJECT = 10
REVISIONS_PER_PROJECT = 12
EVENTS_PER_PROJECT = 20


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(newline="", encoding="utf-8-sig") as f:
        return [{k: (v or "") for k, v in row.items()} for row in csv.DictReader(f)]


def write_csv(path: Path, fields: list[str], rows: list[dict[str, str]]) -> None:
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore", lineterminator="\n")
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in fields})


def take_per_project(rows: list[dict[str, str]], limit: int) -> list[dict[str, str]]:
    grouped: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in rows:
        project = row.get("Project No.", "").strip()
        if project in PROJECTS and len(grouped[project]) < limit:
            grouped[project].append(row)
    out: list[dict[str, str]] = []
    for project in PROJECTS:
        out.extend(grouped.get(project, []))
    return out


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    pilot = root / "review" / "pilot_4_projects"

    docs = read_csv(pilot / "pilot_documents.csv")
    revisions = read_csv(pilot / "pilot_revisions.csv")
    events = read_csv(pilot / "pilot_records.csv")

    doc_fields = [
        "Project No.", "Source Family", "Discipline", "Category", "Subcategory",
        "Document No.", "Document Title", "Company Document No.", "Document Link",
        "Source File", "Source Sheet", "Source Row", "Parsing Status", "Warnings",
    ]
    revision_fields = [
        "Project No.", "Source Family", "Discipline", "Category", "Subcategory",
        "Document No.", "Document Title", "Revision", "Event Date", "Event Status",
        "Source File", "Source Sheet", "Source Row", "Parsing Status", "Warnings",
    ]
    event_fields = [
        "Project No.", "Source Family", "Discipline", "Category", "Subcategory",
        "Document No.", "Document Title", "Company Document No.", "Revision",
        "Event Type", "Event Date", "Event Reference", "Event Status",
        "Document Link", "Source File", "Source Sheet", "Source Row", "Parsing Status", "Warnings",
    ]

    sample_docs = take_per_project(docs, DOCS_PER_PROJECT)
    sample_revisions = take_per_project(revisions, REVISIONS_PER_PROJECT)
    sample_events = take_per_project(events, EVENTS_PER_PROJECT)

    write_csv(pilot / "validation_sample_documents.csv", doc_fields, sample_docs)
    write_csv(pilot / "validation_sample_revisions.csv", revision_fields, sample_revisions)
    write_csv(pilot / "validation_sample_events.csv", event_fields, sample_events)

    print(f"Validation sample: {len(sample_docs)} documents, {len(sample_revisions)} revisions, {len(sample_events)} events")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
