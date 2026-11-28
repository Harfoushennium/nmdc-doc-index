"""Read-only diagnostic of the corpus already present in the repository. No raw records in output."""
import hashlib, json, os, tempfile, time, statistics
from pathlib import Path
from nmdc_profiler.runtime_engine import stage_runtime_update
from nmdc_profiler.source_access import install_resilient_source_access
from nmdc_profiler.layout_compat import install_layout_compatibility
from nmdc_profiler.excel_bridge import load_latest_stage, build_event_rows
from nmdc_profiler.file_review import affected_files, save_file_review
from nmdc_profiler.update_engine import approve_stage
from v2.launcher import App, ordered_fields, export_bytes, filter_rows, partition_plan, partition_export, link_summary
import csv, io, zipfile

def hashes(root):
    return {str(p.relative_to(root)):hashlib.sha256(p.read_bytes()).hexdigest()
            for p in root.rglob("*") if p.is_file() and p.suffix.lower() in (".xlsx",".xlsm",".xls") and not p.name.startswith("~$")}
def measure(f):
    t=time.perf_counter();r=f();return r,round(time.perf_counter()-t,3)
def run():
    root=Path(__file__).resolve().parents[1]/"DATA"
    report={"scope":"read-only repository corpus; scratch state only; no Excel UI or two-PC comparison","status":"NOT_RUN"}
    if not root.is_dir():return dict(report,reason="DATA unavailable")
    before=hashes(root)
    if not before:return dict(report,reason="No workbook files in checkout")
    report["workbook_count"]=len(before)
    install_resilient_source_access();install_layout_compatibility()
    try:
        with tempfile.TemporaryDirectory(prefix="nmdc-v2-benchmark-") as d:
            app=App(Path(d)/"workspace")
            try:
                result,elapsed=measure(lambda:stage_runtime_update(data_dir=root,state_dir=app.state,config_dir=app.config,full_rescan=True,reuse_raw_profiles=True))
                s=load_latest_stage(app.state);rows=s["records"]
                ids=[r.get("Event_Key") for r in rows]
                report.update({"status":"MEASURED","full_stage_seconds":elapsed,"full_stage_phases":result.get("phase_times"),"events":len(rows),
                    "historical_reference_events":40494,"matches_historical_reference":len(rows)==40494,
                    "legacy_event_view_count":len(build_event_rows(rows)),"unique_event_keys":len(set(ids)),
                    "all_keys_nonempty":all(ids),"blocking_flags":result.get("blocking_flags"),"review_flags":result.get("review_flags"),
                    "source_counts":result.get("source_counts")})
                # Disposable benchmark state, no owner production state is read or written.
                if not result.get("blocking_flags") and rows and all(ids) and len(ids)==len(set(ids)):
                    files=affected_files(app.state,result["run_id"])
                    p=Path(d)/"choices.csv"
                    with p.open("w",encoding="utf-8",newline="") as h:
                        w=csv.DictWriter(h,fieldnames=["Source File","Accept Update?","Owner Note"]);w.writeheader()
                        for f in files:w.writerow({"Source File":f["Source File"],"Accept Update?":"TRUE","Owner Note":"Disposable benchmark only"})
                    save_file_review(app.state,result["run_id"],p)
                    _,approval_time=measure(lambda:approve_stage(app.state,result["run_id"],require_file_review=True,allow_conflicts=False))
                    approved=app.view("events");report["scratch_approval_seconds"]=approval_time
                    report["approved_event_ids_match"]=set(r["Event_Key"] for r in approved)==set(ids)
                    _,querytime=measure(lambda:filter_rows(approved,str(approved[0].get("Document No.","")),"Document No."))
                    report["query_seconds"]=querytime
                    try:
                        fields=ordered_fields(approved,"events");parts=partition_plan(approved,fields)
                        data,exporttime=measure(lambda:partition_export(approved,fields) if len(parts)>1 else export_bytes(approved,fields,"xlsx"))
                        with zipfile.ZipFile(io.BytesIO(data)) as z:
                            assert z.testzip() is None
                            if len(parts)==1:assert "xl/tables/table1.xml" in z.namelist()
                            else:
                                manifest=json.loads(z.read("manifest.json"))
                                assert manifest["rows"]==len(approved) and sum(p["rows"] for p in manifest["parts"])==len(approved)
                                for p in manifest["parts"]:
                                    with zipfile.ZipFile(io.BytesIO(z.read(p["file"]))) as part:
                                        assert part.testzip() is None and "xl/tables/table1.xml" in part.namelist()
                        report.update({"xlsx_seconds":exporttime,"xlsx_bytes":len(data),"xlsx_structural":"PASS","xlsx_parts":len(parts),"xlsx_links":link_summary(approved,fields)})
                    except Exception as exc:report["xlsx_structural"]="FAIL: "+type(exc).__name__
                    # Alternate paired modes against the same approved baseline, no approval between trials.
                    trials={"legacy":[],"cached":[]}; signatures={"legacy":[],"cached":[]}
                    for mode in ("legacy","cached","cached","legacy","legacy","cached"):
                        nxt,elapsed=measure(lambda:stage_runtime_update(data_dir=root,state_dir=app.state,config_dir=app.config,full_rescan=False,reuse_raw_profiles=mode=="cached"))
                        stage=app.state/"staging"/nxt["run_id"]
                        sig={key:hashlib.sha256((stage/key).read_bytes()).hexdigest() for key in ("records.jsonl","flags.json","source_decisions.json","record_changes.json","cache_index.json")}
                        manifest=json.loads((stage/"manifest.json").read_text(encoding="utf-8"));manifest.pop("run_id",None)
                        sig["manifest"]=hashlib.sha256(json.dumps(manifest,sort_keys=True).encode()).hexdigest()
                        signatures[mode].append(sig)
                        trials[mode].append({"seconds":elapsed,"phases":nxt.get("phase_times"),"events":nxt.get("staged_records"),"source_counts":nxt.get("source_counts")})
                    parity=all(sig==signatures["legacy"][0] for mode in signatures for sig in signatures[mode])
                    report.update({"incremental_trials":trials,"legacy_incremental_median_seconds":statistics.median(t["seconds"] for t in trials["legacy"]),"cached_incremental_median_seconds":statistics.median(t["seconds"] for t in trials["cached"]),"incremental_output_parity":parity})
                    if not parity:raise AssertionError("Cached profiles changed staged output")
                else:report["scratch_approval"]="NOT_RUN: flags or record identity discrepancy"
            finally:app.lock.close()
    except Exception as exc:report.update({"status":"ERROR","error_type":type(exc).__name__})
    report["source_hashes_unchanged"]=hashes(root)==before
    if not report["source_hashes_unchanged"]:raise RuntimeError("DATA source hash changed")
    return report
if __name__=="__main__":
    report=run()
    Path("v2-corpus-benchmark.json").write_text(json.dumps(report,indent=2),encoding="utf-8")
    print(json.dumps(report,indent=2))
