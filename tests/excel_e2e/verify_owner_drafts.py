"""Verify owner-input behavior in a disposable, real Microsoft Excel workbook.

Run after the normal setup script creates an XLSM in an extracted package:
  uv run --with pywin32 python tests/excel_e2e/verify_owner_drafts.py --package PATH

Only the copied workbook/runtime is changed. This is COM behavior evidence,
not acceptance of physical checkbox clicking, typing, layout or dialogs.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from nmdc_profiler.ui_layout import FLAG_FIELDS, PENDING_FIELDS
from nmdc_profiler.source_selection_view import FIELDS as SOURCE_FIELDS


def read_csv(path):
    with path.open(newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        return reader.fieldnames, list(reader)


def write_csv(path, headers, rows):
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers)
        writer.writeheader()
        writer.writerows(rows)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--package", required=True, type=Path)
    args = parser.parse_args()
    package = args.package.resolve()
    workbook_path = package / "NMDC_Document_Index.xlsm"
    if not workbook_path.is_file():
        parser.error("Run the package setup script first; NMDC_Document_Index.xlsm is missing.")
    target = ROOT / "outputs" / "owner-draft-verification" / time.strftime("%Y%m%d-%H%M%S")
    shutil.copytree(package, target, ignore=shutil.ignore_patterns("runtime", "DATA", "PARSER_FIX_REPORTS"))
    runtime = target / "runtime"
    exchange = runtime / "excel_exchange"
    subprocess.run([str(target / "engine/nmdc_index_engine.exe"), "export-excel", "--state-dir", str(runtime),
                    "--exchange-dir", str(exchange), "--config-dir", str(target / "config")], check=True)
    report = {"package": str(package), "copy": str(target), "physical_acceptance": False,
              "workbook_sha256": hashlib.sha256(workbook_path.read_bytes()).hexdigest(), "checks": []}
    excel = None
    book = None

    def record(name, passed, **detail):
        report["checks"].append({"name": name, "passed": bool(passed), **detail})
        (target / "report.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
        print(json.dumps(report["checks"][-1]), flush=True)
        if not passed:
            raise AssertionError(name)

    try:
        import win32com.client
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.EnableEvents = False
        book = excel.Workbooks.Open(str(target / "NMDC_Document_Index.xlsm"), UpdateLinks=0)
        components = book.VBProject.VBComponents
        record("draft module shipped", components("modNMDC_ReviewDrafts").CodeModule.CountOfLines > 0)
        this = components("ThisWorkbook").CodeModule
        record("draft workbook event shipped", "nmdc_reviewdraftsheetchange sh, target" in this.Lines(1, this.CountOfLines).casefold())
        # The probe wrapper only turns public draft-read errors into observable results.
        wrapper = components.Add(1)
        wrapper.Name = "OwnerDraftReadProbe"
        wrapper.CodeModule.AddFromString('''Option Explicit
Public Function ReadDraftProbe(ByVal id As String) As String
    Dim found As Boolean
    On Error GoTo Handler
    ReadDraftProbe = NMDC_ReviewDraftGet(NMDC_DisplayedPendingRunId(), "SOURCE", id, "Owner Note", found)
    Exit Function
Handler:
    ReadDraftProbe = "ERROR " & CStr(Err.Number) & " " & Err.Description
End Function
''')
        prefix = "'" + book.Name.replace("'", "''") + "'!"
        data_folder = target / "fixture-data"
        data_folder.mkdir()
        for row in book.Worksheets("Configuration").ListObjects("Configuration").ListRows:
            key = str(row.Range.Cells(1, 1).Value or "").strip()
            if key in {"Runtime Folder", "Configuration Folder", "Data Folder"}:
                row.Range.Cells(1, 2).Value = {"Runtime Folder": "runtime", "Configuration Folder": "config", "Data Folder": str(data_folder)}[key]
        run_a = "owner-draft-fixture-A"
        source_rows = [dict.fromkeys(SOURCE_FIELDS, "") for _ in range(2)]
        for index, row in enumerate(source_rows):
            row.update({"Include in Index?": "TRUE", "Project No.": "9999", "Source File": f"METHODS/Fixture-{index}.xlsx",
                        "Source Family": "METHODS", "Current Status": "SELECTED", "Last Processed Run": run_a})
        flag_rows = [dict.fromkeys(FLAG_FIELDS, "") for _ in range(2)]
        for index, row in enumerate(flag_rows):
            row.update({"Select?": "FALSE", "Flag Code": "UNCLASSIFIED", "Source File": source_rows[index]["Source File"],
                        "Worksheet Name": "Register", "Event Key": f"fixture-event-{index}", "Project No.": "9999",
                        "Document No.": f"TEST-{index}", "Revision": "A"})
        pending_rows = [dict.fromkeys(PENDING_FIELDS, "") for _ in range(35)]
        for index, row in enumerate(pending_rows):
            row.update({"Change Type": "ADDED", "Project No.": "9999", "Document No.": f"TEST-{index}",
                        "Revision": "A", "Record Identity": f"fixture-record-{index}"})
        headers, dashboard = read_csv(exchange / "dashboard.csv")
        dashboard[0].update({"Pending Run ID": run_a, "Pending Status": "STAGED", "Pending Added": "3", "Review Flags": "2"})
        write_csv(exchange / "dashboard.csv", headers, dashboard)
        write_csv(exchange / "source_selection.csv", SOURCE_FIELDS, source_rows)
        write_csv(exchange / "flags.csv", FLAG_FIELDS, flag_rows)
        write_csv(exchange / "pending_update.csv", PENDING_FIELDS, pending_rows[:3])
        book.Worksheets("Home").Range("H7").Value = ""
        record("fixture full refresh", excel.Run(prefix + "NMDC_RefreshExchangeData"))
        pending = book.Worksheets("Pending Update")
        flags = book.Worksheets("Review Flags").ListObjects("ReviewFlags")

        def source_cell(field, index=1):
            return pending.ListObjects("SourceSelection").ListColumns(field).DataBodyRange.Cells(index, 1)

        def flag_cell(field, index=1):
            return flags.ListColumns(field).DataBodyRange.Cells(index, 1)

        def can_approve():
            return bool(excel.Run(prefix + "NMDC_ReviewDraftsCanApprove", str(excel.Run(prefix + "NMDC_DisplayedPendingRunId")), ""))

        record("clean proposal has no draft block", can_approve())
        multiline = 'First line\nSecond line, "quoted"'
        source_cell("Include in Index?").Value = False
        source_cell("Owner Note").Value = multiline
        flag_cell("Select?").Value = True
        flag_cell("User Decision").Value = "ACKNOWLEDGED"
        flag_cell("User Comment").Value = multiline
        flag_cell("Resolution Status").Value = "OPEN"
        record("fast refresh retains disabled-event edits", excel.Run(prefix + "NMDC_RefreshReviewDataFast"))
        record("source inputs retained", not bool(source_cell("Include in Index?").Value) and source_cell("Owner Note").Value == multiline)
        record("flag inputs retained", bool(flag_cell("Select?").Value) and flag_cell("User Decision").Value == "ACKNOWLEDGED" and
               flag_cell("User Comment").Value == multiline and flag_cell("Resolution Status").Value == "OPEN")
        record("dirty proposal is blocked", not can_approve())
        # Exercise the actual approval action. Replace only its blocking message
        # adapter so an unattended test does not need to dismiss a modal dialog.
        actions = components("modNMDC_Actions").CodeModule
        action_source = actions.Lines(1, actions.CountOfLines)
        blocked_notice = 'MsgBox draftBlockReason, vbExclamation, "NMDC Document Index"'
        assert action_source.count(blocked_notice) == 1
        dashboard_before = (exchange / "dashboard.csv").read_bytes()
        actions.DeleteLines(1, actions.CountOfLines)
        actions.AddFromString(action_source.replace(blocked_notice,
            'ThisWorkbook.Worksheets("Home").Range("Z1").Value2 = draftBlockReason'))
        try:
            excel.Run(prefix + "NMDC_ApproveUpdate")
            record("actual approval action stops before export or promotion",
                   "draft" in str(book.Worksheets("Home").Range("Z1").Value).casefold() and
                   (exchange / "dashboard.csv").read_bytes() == dashboard_before and
                   not (runtime / "approved/current.json").exists(), notification_adapter_test_double=True)
        finally:
            actions.DeleteLines(1, actions.CountOfLines)
            actions.AddFromString(action_source)
        record("full refresh retains draft", excel.Run(prefix + "NMDC_RefreshExchangeData") and source_cell("Owner Note").Value == multiline and flag_cell("User Comment").Value == multiline)
        # Reorder exchange rows: edits follow identity, not worksheet row number.
        write_csv(exchange / "source_selection.csv", SOURCE_FIELDS, list(reversed(source_rows)))
        write_csv(exchange / "flags.csv", FLAG_FIELDS, list(reversed(flag_rows)))
        record("reordered refresh", excel.Run(prefix + "NMDC_RefreshReviewDataFast"))
        record("drafts follow source and flag identities", source_cell("Owner Note", 2).Value == multiline and flag_cell("User Comment", 2).Value == multiline and not source_cell("Owner Note", 1).Value)
        book.Save()
        book.Close(SaveChanges=False)
        book = excel.Workbooks.Open(str(target / "NMDC_Document_Index.xlsm"), UpdateLinks=0)
        pending = book.Worksheets("Pending Update")
        flags = book.Worksheets("Review Flags").ListObjects("ReviewFlags")
        record("saved reopen refresh", excel.Run(prefix + "NMDC_RefreshReviewDataFast"))
        record("saved reopen retains draft", source_cell("Owner Note", 2).Value == multiline and flag_cell("User Comment", 2).Value == multiline and not can_approve())
        excel.EnableEvents = True
        source_cell("Owner Note", 2).Value = "Event-captured note"
        excel.Run(prefix + "NMDC_RebuildPendingSourcePanel")
        record("event-enabled panel refresh retains inputs", source_cell("Owner Note", 2).Value == "Event-captured note" and bool(excel.EnableEvents))
        excel.EnableEvents = False
        write_csv(exchange / "pending_update.csv", PENDING_FIELDS, pending_rows)
        record("pending table grows past old source panel", excel.Run(prefix + "NMDC_LoadCsvToTable", str(exchange / "pending_update.csv"), "Pending Update", "PendingUpdate") and pending.ListObjects("PendingUpdate").ListRows.Count == 35)
        record("growth preserves source inputs and native checkboxes", source_cell("Owner Note", 2).Value == "Event-captured note" and int(source_cell("Include in Index?", 2).CellControl.Type) == 2)
        # A different proposal may not inherit drafts from this one.
        run_b = "owner-draft-fixture-B"
        dashboard[0]["Pending Run ID"] = run_b
        write_csv(exchange / "dashboard.csv", headers, dashboard)
        record("replacement proposal refresh", excel.Run(prefix + "NMDC_RefreshReviewDataFast"))
        record("old drafts retained without cross-run application", not source_cell("Owner Note", 2).Value and not flag_cell("User Comment", 2).Value and "retained" in str(pending.Range("A4").Value).casefold())
        draft_path = runtime / "owner_review_drafts.csv"
        draft_headers, drafts = read_csv(draft_path)
        record("old draft report retains original run", any(row["Run ID"] == run_a and row["Value"] == "Event-captured note" for row in drafts))
        source_cell("Owner Note", 2).Value = "Saved before stage"
        record("capture restage choice", excel.Run(prefix + "NMDC_ReviewDraftCaptureCurrentTables"))
        excel.Run(prefix + "NMDC_ReviewDraftMarkSourceRestage", run_b)
        record("restage waiting blocks approval", not can_approve())
        # The proposal received the saved value, but the owner edited it again while staging.
        source_cell("Owner Note", 2).Value = "Edited during stage"
        record("capture later owner edit", excel.Run(prefix + "NMDC_ReviewDraftCaptureCurrentTables"))
        replacement_sources = list(reversed(source_rows))
        replacement_sources[1] = dict(replacement_sources[1], **{"Owner Note": "Saved before stage"})
        write_csv(exchange / "source_selection.csv", SOURCE_FIELDS, replacement_sources)
        dashboard[0]["Pending Run ID"] = "owner-draft-fixture-C"
        write_csv(exchange / "dashboard.csv", headers, dashboard)
        record("restage with later edit refresh", excel.Run(prefix + "NMDC_RefreshReviewDataFast"))
        _, drafts = read_csv(draft_path)
        record("later edit retained and restage remains blocked", any(row["Value"] == "Edited during stage" for row in drafts) and not can_approve())
        record("approved state untouched", not (runtime / "approved/current.json").exists())
        record("workbook contract preserved", book.Worksheets.Count == 15 and sum(sheet.ListObjects.Count for sheet in book.Worksheets) == 15)
        # Corruption must be rejected at the public store interface without a UI dialog.
        write_csv(draft_path, draft_headers, [{"Run ID": run_a, "Section": "UNKNOWN", "Identity": "fixture", "Field": "Owner Note", "Value": "bad"}])
        record("malformed draft rejected", str(excel.Run(prefix + "ReadDraftProbe", "fixture")).startswith("ERROR"))
    except Exception as exc:
        report["exception"] = {"type": type(exc).__name__, "message": str(exc)}
        (target / "report.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
        raise
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception as exc:
                report["close_error"] = str(exc)
        if excel is not None:
            try:
                excel.Quit()
            except Exception as exc:
                report["quit_error"] = str(exc)
        (target / "report.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    print("REPORT " + str(target / "report.json"), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
