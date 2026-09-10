from __future__ import annotations

from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def replace_once(path: Path, old: str, new: str) -> None:
    text = path.read_text(encoding="utf-8")
    if old not in text:
        raise RuntimeError(f"Expected patch target not found in {path}")
    path.write_text(text.replace(old, new, 1), encoding="utf-8", newline="\n")


def main() -> int:
    extractor = ROOT / "nmdc_profiler" / "extractor.py"
    tests = ROOT / "tests" / "test_extractor.py"
    verifier = ROOT / "tests" / "verify_cycle2_real_data.py"

    replace_once(
        extractor,
        '''def _is_static_path(path: Sequence[str]) -> bool:\n    n = norm_text(" ".join(path))\n    return bool(re.search(r"\\b(sr|serial)\\s+no\\b|\\bclass\\b|\\bremarks?\\b", n))\n\n\ndef _event_group_label(path: Sequence[str], col: int) -> str:\n    if not path:\n        return f"COLUMN {_col_letter(col)}"\n    leaf = norm_text(path[-1])\n    leafish = bool(re.search(r"\\b(date|ref|reference|status|code|response|transmittal)\\b", leaf))\n    parent = path[:-1] if leafish and len(path) > 1 else path\n    label = " > ".join(parent).strip()\n    return label or path[-1]\n''',
        '''def _is_static_path(path: Sequence[str]) -> bool:\n    if any(str(part).strip() == "#" for part in path):\n        return True\n    n = norm_text(" ".join(path))\n    return bool(re.search(r"\\b(sr|serial)\\s+no\\b|\\bclass\\b|\\bremarks?\\b", n))\n\n\ndef _event_field_kind(text: str) -> str:\n    n = norm_text(text)\n    if not n:\n        return ""\n    if n.startswith("issue date") or n in {"date", "actual date", "planned date", "response date", "approval date"}:\n        return "date"\n    if n in {"ref", "ref no", "reference", "reference no", "outgoing ref", "outgoing ref no", "transmittal", "transmittal no"}:\n        return "reference"\n    if n in {"code", "status", "response", "approval code", "comment code"}:\n        return "status"\n    return ""\n\n\ndef _event_value_label(path: Sequence[str], col: int) -> str:\n    if not path:\n        return _col_letter(col)\n    if _event_field_kind(path[-1]):\n        return path[-1]\n    if len(path) > 1 and _event_field_kind(path[0]):\n        return path[0]\n    return path[-1]\n\n\ndef _event_group_label(path: Sequence[str], col: int) -> str:\n    if not path:\n        return f"COLUMN {_col_letter(col)}"\n    if _event_field_kind(path[-1]) and len(path) > 1:\n        group = path[:-1]\n    elif len(path) > 1 and _event_field_kind(path[0]):\n        group = path[1:]\n    else:\n        group = path\n    label = " > ".join(group).strip()\n    return label or path[-1]\n''',
    )

    replace_once(
        extractor,
        '''def _event_values(model: SheetModel, row: int, cols: Sequence[int], layout: Layout) -> Tuple[Dict[str, str], str, str, str, List[str]]:\n    values: Dict[str, str] = {}\n    dates: List[str] = []\n    refs: List[str] = []\n    statuses: List[str] = []\n    warnings: List[str] = []\n    for col in cols:\n        raw = model.value(row, col).strip()\n        if not raw:\n            continue\n        path = _header_path(model, col, layout.header_start, layout.header_end)\n        leaf = path[-1] if path else _col_letter(col)\n        key = leaf if leaf not in values else f"{leaf}@{_col_letter(col)}"\n        leaf_norm = norm_text(leaf)\n        if "date" in leaf_norm:\n            normalized, warning = _normalize_date(raw)\n            values[key] = normalized\n            if normalized:\n                dates.append(normalized)\n            if warning:\n                warnings.append(warning)\n        else:\n            values[key] = raw\n        if re.search(r"\\b(ref|reference|transmittal)\\b", leaf_norm):\n            refs.append(raw)\n        if re.search(r"\\b(status|code|response|approval)\\b", leaf_norm):\n            statuses.append(raw)\n    return values, "; ".join(dates), "; ".join(refs), "; ".join(statuses), sorted(set(warnings))\n''',
        '''def _event_values(model: SheetModel, row: int, cols: Sequence[int], layout: Layout) -> Tuple[Dict[str, str], str, str, str, List[str]]:\n    values: Dict[str, str] = {}\n    date_candidates: List[Tuple[int, int, str]] = []\n    refs: List[str] = []\n    statuses: List[str] = []\n    warnings: List[str] = []\n    for col in cols:\n        raw = model.value(row, col).strip()\n        if not raw:\n            continue\n        path = _header_path(model, col, layout.header_start, layout.header_end)\n        field_label = _event_value_label(path, col)\n        key = field_label if field_label not in values else f"{field_label}@{_col_letter(col)}"\n        field_norm = norm_text(field_label)\n        kind = _event_field_kind(field_label)\n        if kind == "date":\n            normalized, warning = _normalize_date(raw)\n            values[key] = normalized\n            if normalized and not warning and re.fullmatch(r"\\d{4}-\\d{2}-\\d{2}", normalized):\n                priority = 40 if "actual" in field_norm else 30 if any(x in field_norm for x in ("response", "approval")) else 10 if "planned" in field_norm else 20\n                date_candidates.append((priority, col, normalized))\n            if warning:\n                warnings.append(warning)\n        else:\n            values[key] = raw\n        if kind == "reference" or re.fullmatch(r"(?:outgoing )?(?:ref|reference)(?: no)?", field_norm):\n            refs.append(raw)\n        if kind == "status":\n            statuses.append(raw)\n    event_date = max(date_candidates, default=(0, 0, ""), key=lambda item: (item[0], item[1]))[2]\n    return values, event_date, "; ".join(refs), "; ".join(statuses), sorted(set(warnings))\n''',
    )

    replace_once(
        extractor,
        'str(r["Event Key"])',
        'str(r["Event_Key"])',
    )

    replace_once(
        tests,
        'import unittest\nfrom pathlib import Path\n',
        'import tempfile\nimport unittest\nfrom pathlib import Path\nfrom unittest.mock import patch\n',
    )
    replace_once(
        tests,
        '    extract_model,\n)',
        '    extract_model,\n    run_sentinels,\n)',
    )
    replace_once(
        tests,
        '\n\nclass ExtractorTests(unittest.TestCase):',
        '''\n\ndef reversed_header_model():\n    cells = {\n        (1, 1): "#", (1, 2): "Document No.", (1, 3): "Document Title", (1, 4): "Revision",\n        (1, 5): "Issue Date (Planned)", (1, 6): "Outgoing Ref No.", (1, 7): "Issue Date (Actual)",\n        (2, 5): "Construction and Installation Procedure",\n        (3, 1): "1", (3, 2): "2171-2172-PP-OF-012", (3, 3): "ANCHOR HANDLING PROCEDURE FOR FLOATING BARGE",\n        (3, 4): "A1", (3, 5): "44438", (3, 6): "T-553/21", (3, 7): "44553",\n    }\n    model = SheetModel(\n        "DATA/METHODS/2171-2172 -Document Deliverables LATEST.xlsx",\n        "METHODS",\n        "2026-09-08T05:23:16Z",\n        "2171-2172",\n        3,\n        7,\n        cells=cells,\n    )\n    add_merge(model, 2, 5, 2, 7)\n    return model\n\n\nclass ExtractorTests(unittest.TestCase):''',
    )
    replace_once(
        tests,
        '''    def test_13_client_is_not_globally_excluded(self):\n        model = SheetModel("DATA/TECH/9999 DOCUMENT REGISTER.xlsx", "TECH", "", "CLIENT", 1, 1, cells={(1, 1): "anything"})\n        case = SentinelCase("CLIENT", "x", "CLIENT", "INCLUDE")\n        records, recon = extract_model(case, model, self.rules)\n        self.assertEqual([], records)\n        self.assertEqual("REVIEW_REQUIRED", recon["status"])\n''',
        '''    def test_13_client_is_not_globally_excluded(self):\n        model = SheetModel("DATA/TECH/9999 DOCUMENT REGISTER.xlsx", "TECH", "", "CLIENT", 1, 1, cells={(1, 1): "anything"})\n        case = SentinelCase("CLIENT", "x", "CLIENT", "INCLUDE")\n        records, recon = extract_model(case, model, self.rules)\n        self.assertEqual([], records)\n        self.assertEqual("REVIEW_REQUIRED", recon["status"])\n\n    def test_14_reversed_multilevel_header_consolidates_one_transaction(self):\n        layout, warnings = discover_layout(reversed_header_model())\n        self.assertIsNotNone(layout)\n        self.assertEqual((("Construction and Installation Procedure", (5, 6, 7)),), layout.event_groups)\n        self.assertIn(1, layout.metadata_cols)\n        self.assertEqual([], warnings)\n\n    def test_15_reversed_header_uses_actual_date_and_preserves_all_fields(self):\n        case = SentinelCase("REVHDR", "x", "2171-2172", "INCLUDE")\n        records, recon = extract_model(case, reversed_header_model(), self.rules)\n        self.assertEqual("INCLUDE", recon["status"])\n        self.assertEqual(1, len(records))\n        row = records[0]\n        self.assertEqual("Construction and Installation Procedure", row["Event Type"])\n        self.assertEqual("2021-12-23", row["Event Date"])\n        self.assertEqual("T-553/21", row["Event Reference"])\n        self.assertNotEqual("#", row["Event Type"])\n        self.assertIn('"Issue Date (Planned)":"2021-08-30"', row["Event Values JSON"])\n        self.assertIn('"Issue Date (Actual)":"2021-12-23"', row["Event Values JSON"])\n\n    def test_16_non_date_text_in_date_column_is_preserved_but_not_promoted(self):\n        model = reversed_header_model()\n        model.cells[(3, 7)] = "PMT ISSUED"\n        case = SentinelCase("REVHDR", "x", "2171-2172", "INCLUDE")\n        records, _ = extract_model(case, model, self.rules)\n        self.assertEqual("2021-08-30", records[0]["Event Date"])\n        self.assertIn("DATE_TEXT_PRESERVED", records[0]["Warnings"])\n        self.assertIn('"Issue Date (Actual)":"PMT ISSUED"', records[0]["Event Values JSON"])\n\n    def test_17_run_sentinels_sorts_by_canonical_event_key(self):\n        case = SentinelCase("REVHDR", "dummy.xlsx", "2171-2172", "INCLUDE")\n        with tempfile.TemporaryDirectory() as td:\n            root = Path(td)\n            (root / "dummy.xlsx").write_bytes(b"placeholder")\n            with patch("nmdc_profiler.extractor.read_sheet_model", return_value=reversed_header_model()):\n                records, reconciliation = run_sentinels(root, [case], self.rules)\n        self.assertEqual(1, len(records))\n        self.assertTrue(records[0]["Event_Key"].startswith("EVT-"))\n        self.assertEqual("INCLUDE", reconciliation[0]["status"])\n''',
    )

    replace_once(
        verifier,
        '''    # Nonstandard Methods sheet must extract via the approved M002 exception.\n    nonstandard = [r for r in records if r["Case ID"] == "M2171_NONSTANDARD"]\n    if not nonstandard or not all("M002" in r["Classification Rule ID"].split(";") for r in nonstandard):\n        fail("2171-2172 sentinel did not use scoped M002 structural classification")\n''',
        '''    # Nonstandard Methods sheet must extract via the approved M002 exception.\n    nonstandard = [r for r in records if r["Case ID"] == "M2171_NONSTANDARD"]\n    if not nonstandard or not all("M002" in r["Classification Rule ID"].split(";") for r in nonstandard):\n        fail("2171-2172 sentinel did not use scoped M002 structural classification")\n    nonstandard_recon = by_case["M2171_NONSTANDARD"]\n    if int(nonstandard_recon["event_records"]) != int(nonstandard_recon["rows_with_identity"]):\n        fail(f"2171-2172 should produce one consolidated event record per revision row: {nonstandard_recon}")\n    bad_2171_types = [\n        r["Event Type"] for r in nonstandard\n        if r["Event Type"] == "#"\n        or r["Event Type"].startswith("Issue Date")\n        or r["Event Type"].startswith("Outgoing Ref")\n    ]\n    if bad_2171_types:\n        fail(f"2171-2172 still exposes column headers as separate event types: {bad_2171_types[:5]}")\n    first_2171 = [r for r in nonstandard if r["Document No."] == "2171-2172-PP-OF-012" and r["Revision"] == "A1"]\n    if not first_2171:\n        fail("2171-2172 known A1 sentinel revision is missing")\n    first_event = first_2171[0]\n    if first_event["Event Type"] != "Construction and Installation Procedure":\n        fail(f"2171-2172 event group was not consolidated: {first_event['Event Type']}")\n    if first_event["Event Date"] != "2021-12-23" or first_event["Event Reference"] != "T-553/21":\n        fail(f"2171-2172 A1 event semantics mismatch: {first_event}")\n    values_2171 = json.loads(first_event["Event Values JSON"])\n    if values_2171.get("Issue Date (Planned)") != "2021-08-30" or values_2171.get("Issue Date (Actual)") != "2021-12-23":\n        fail(f"2171-2172 planned/actual dates were not preserved correctly: {values_2171}")\n''',
    )

    print("Cycle 2 hardening patch applied")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
