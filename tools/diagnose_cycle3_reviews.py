from __future__ import annotations

import csv
from pathlib import Path

from nmdc_profiler.extractor import _is_company_doc_header, _is_doc_header, _is_revision_header, _is_title_header, _looks_identifier, read_sheet_model
from nmdc_profiler.full_extractor import resolve_sheet_name


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    review_path = root / "outputs" / "cycle3" / "review_queue.csv"
    with review_path.open(newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    for item in rows:
        if item["Scope"] != "WORKSHEET":
            continue
        source = root / item["Source File"]
        sheet = resolve_sheet_name(source, item["Worksheet"])
        model = read_sheet_model(source, root, sheet)
        print("\n===", item["Source File"], "::", repr(sheet), "===")
        print("reason:", item["Reason"], item["Warnings"], "size", model.max_row, "x", model.max_col)
        headers = []
        for (r, c), value in sorted(model.cells.items()):
            if r > min(model.max_row, 30):
                continue
            if _is_doc_header(value) or _is_company_doc_header(value) or _is_title_header(value) or _is_revision_header(value) or any(k in value.casefold() for k in ("number", " no", "title", "description", "revision", "rev.")):
                headers.append((r, c, value, _is_doc_header(value), _is_company_doc_header(value), _is_title_header(value), _is_revision_header(value)))
        print("candidate/header-like cells:")
        for h in headers[:40]:
            print("  ", h)
        print("identifier-like cells (first 30 after row 3):")
        count = 0
        for (r, c), value in sorted(model.cells.items()):
            if r <= 3:
                continue
            if _looks_identifier(value):
                print("  ", (r, c, value))
                count += 1
                if count >= 30:
                    break
        if count == 0:
            print("   [none]")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
