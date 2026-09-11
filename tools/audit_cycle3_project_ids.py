from __future__ import annotations

import csv
import re
from collections import Counter
from pathlib import Path

from nmdc_profiler.core import norm_text
from nmdc_profiler.extractor import _looks_identifier, read_sheet_model
from nmdc_profiler.full_extractor import resolve_sheet_name
from nmdc_profiler.project_identity import load_project_identity_overrides


def read_csv(path: Path):
    with path.open(newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def expected_projects(text: str) -> set[str]:
    return set(re.findall(r"\b\d{4}\b", text or ""))


def explicit_sheet_projects(model) -> set[str]:
    """Return only project numbers directly tied to a Project No./Number label.

    Avoid collecting unrelated four-digit values elsewhere on a header row.
    """
    found = set()
    label_re = re.compile(r"\b(?:project|proj)\s+(?:no|number)\b", re.I)
    value_re = re.compile(r"^\s*(\d{4})\b")
    inline_re = re.compile(r"\b(?:project|proj)\s+(?:no|number)\s*[:#-]?\s*(\d{4})\b", re.I)
    for (r, c), value in model.cells.items():
        if r > 15:
            continue
        n = norm_text(value)
        if not label_re.search(n):
            continue
        inline = inline_re.search(value)
        if inline:
            found.add(inline.group(1))
        right = model.value(r, c + 1) if c < model.max_col else ""
        m = value_re.match(right)
        if m:
            found.add(m.group(1))
    return found


def document_prefix_projects(model) -> Counter[str]:
    """Count four-digit project-like prefixes in actual document identifiers."""
    counts: Counter[str] = Counter()
    for (r, _c), value in model.cells.items():
        if r <= 3 or not _looks_identifier(value):
            continue
        m = re.match(r"^\s*(\d{4})[-_/ ]", value)
        if m:
            counts[m.group(1)] += 1
    return counts


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    inventory = read_csv(root / "outputs/cycle1/source_inventory.csv")
    discovery = read_csv(root / "outputs/cycle1/classification_discovery.csv")
    overrides = load_project_identity_overrides(root / "config/project_identity_overrides.csv")
    selected = {
        r["relative_path"]: expected_projects(r.get("inferred_project_numbers", ""))
        for r in inventory if r["selected_excluded_status"].strip().upper() == "SELECTED"
    }
    seen = set()
    strong_mismatches = []
    weak_mismatches = []
    matches = 0
    no_internal = 0
    owner_approved = 0

    for row in discovery:
        src = row["workbook_path"]
        if src not in selected or row["proposed_action"].strip().upper() != "INCLUDE":
            continue
        key = (src, " ".join(row["worksheet_name"].split()).casefold())
        if key in seen:
            continue
        seen.add(key)
        source = root / src
        sheet = resolve_sheet_name(source, row["worksheet_name"])
        model = read_sheet_model(source, root, sheet)
        internal = explicit_sheet_projects(model)
        prefixes = document_prefix_projects(model)
        expected = selected[src]

        if src in overrides:
            approved = overrides[src]["Approved Project No."]
            owner_approved += 1
            print("OWNER_APPROVED", (src, sheet, approved, sorted(internal), dict(prefixes)))
            continue

        if not internal:
            no_internal += 1
            continue
        if expected and internal.isdisjoint(expected):
            internal_support = max((prefixes.get(p, 0) for p in internal), default=0)
            expected_support = max((prefixes.get(p, 0) for p in expected), default=0)
            item = (src, sheet, sorted(expected), sorted(internal), dict(prefixes))
            # A conflicting header alone is not enough to quarantine a project.
            # Require repeated document-number evidence supporting the conflicting identity.
            if internal_support >= 3 and internal_support > expected_support:
                strong_mismatches.append(item)
            else:
                weak_mismatches.append(item)
        else:
            matches += 1

    print(f"audited included sheets: {len(seen)}")
    print(f"explicit project match: {matches}")
    print(f"owner-approved project identity: {owner_approved}")
    print(f"no explicit project label: {no_internal}")
    print(f"strong mismatches: {len(strong_mismatches)}")
    print(f"weak mismatches: {len(weak_mismatches)}")
    for item in strong_mismatches:
        print("STRONG_MISMATCH", item)
    for item in weak_mismatches:
        print("WEAK_MISMATCH", item)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
