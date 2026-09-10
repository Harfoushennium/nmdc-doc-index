from __future__ import annotations

import csv
import re
from pathlib import Path

from nmdc_profiler.core import norm_text
from nmdc_profiler.extractor import _looks_identifier, read_sheet_model
from nmdc_profiler.full_extractor import resolve_sheet_name


def read_csv(path: Path):
    with path.open(newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def expected_projects(text: str) -> set[str]:
    return set(re.findall(r"\b\d{4}\b", text or ""))


def explicit_sheet_projects(model) -> set[str]:
    found = set()
    for (r, c), value in model.cells.items():
        if r > 12:
            continue
        n = norm_text(value)
        if not re.search(r"\b(?:project|proj)\s+(?:no|number)\b", n):
            continue
        for token in re.findall(r"\b\d{4}\b", value):
            found.add(token)
        right = model.value(r, c + 1) if c < model.max_col else ""
        for token in re.findall(r"\b\d{4}\b", right):
            found.add(token)
    return found


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    inventory = read_csv(root / "outputs/cycle1/source_inventory.csv")
    discovery = read_csv(root / "outputs/cycle1/classification_discovery.csv")
    selected = {
        r["relative_path"]: expected_projects(r.get("inferred_project_numbers", ""))
        for r in inventory if r["selected_excluded_status"].strip().upper() == "SELECTED"
    }
    seen = set()
    mismatches = []
    matches = 0
    no_internal = 0
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
        expected = selected[src]
        if not internal:
            no_internal += 1
            continue
        if expected and internal.isdisjoint(expected):
            id_cells = [(r, c, v) for (r, c), v in model.cells.items() if r > 3 and _looks_identifier(v)]
            mismatches.append((src, sheet, sorted(expected), sorted(internal), len(id_cells), id_cells[:5]))
        else:
            matches += 1
    print(f"audited included sheets: {len(seen)}")
    print(f"explicit project match: {matches}")
    print(f"no explicit project label: {no_internal}")
    print(f"mismatches: {len(mismatches)}")
    for item in mismatches:
        print("MISMATCH", item)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
