"""Packaged full scan/approval/no-op acceptance on repository corpus, scratch workspace only."""
import hashlib, json, os, statistics, subprocess, sys, tempfile, time, urllib.request
from pathlib import Path
from nmdc_profiler.runtime_engine import stage_runtime_update
from nmdc_profiler.source_access import install_resilient_source_access
from nmdc_profiler.layout_compat import install_layout_compatibility
from v2.launcher import App, filter_rows, apply_filters, ordered_fields
from v2.unchanged import check,build_identity

def hashes(root):
    return {p.relative_to(root).as_posix():hashlib.sha256(p.read_bytes()).hexdigest() for p in root.rglob("*") if p.is_file() and p.suffix.lower() in (".xlsx",".xlsm") and not p.name.startswith("~$")}
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def signature(stage):
    return {key:sha(stage/key) for key in ("records.jsonl","flags.json","source_decisions.json","record_changes.json")}
def main(exe):
    data=(Path(__file__).resolve().parents[1]/"DATA").resolve()
    before=hashes(data);assert before
    install_resilient_source_access();install_layout_compatibility()
    with tempfile.TemporaryDirectory(prefix="nmdc-fast-accept-") as d:
        root=Path(d);serial=App(root/"serial")
        try:
            t=time.perf_counter()
            baseline=stage_runtime_update(data_dir=data,state_dir=serial.state,config_dir=serial.config,full_rescan=True,reuse_raw_profiles=True,verify_source_content=True,extraction_workers=1)
            serial_time=time.perf_counter()-t
            expected=signature(serial.state/"staging"/baseline["run_id"])
        finally:serial.lock.close()
        workspace=root/"parallel";app=App(workspace)
        try:app.settings["data_dir"]=str(data);app.save_settings()
        finally:app.lock.close()
        env=os.environ.copy();env.pop("PYTHONPATH",None)
        system=env.get("SYSTEMROOT",env.get("SystemRoot",env.get("WINDIR",r"C:\Windows")))
        env["PATH"]=os.pathsep.join([str(Path(system)/"System32"),system])
        proc=subprocess.Popen([str(exe),"--workspace",str(workspace),"--no-browser"],stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,env=env)
        try:
            line=proc.stdout.readline();assert "http://127.0.0.1:" in line,line
            url,token=line[line.index("http://"):].strip().split("#",1);url=url.rstrip("/")
            def api(path,body=None):
                headers={"Authorization":"Bearer "+token}
                if body is not None:headers.update({"Origin":url,"Content-Type":"application/json"})
                req=urllib.request.Request(url+path,data=None if body is None else json.dumps(body).encode(),headers=headers)
                try:
                    with urllib.request.urlopen(req,timeout=120) as res:return json.load(res)
                except urllib.error.HTTPError as exc:
                    raise AssertionError(f"API {path.split('?')[0]} returned {exc.code}: {exc.read().decode(errors='replace')}") from exc
            def wait_job(expected_status="complete"):
                end=time.monotonic()+180
                while time.monotonic()<end:
                    job=api("/api/status")["job"]
                    if job["status"]!="running":
                        assert job["status"]==expected_status,job
                        return job.get("result",job)
                    time.sleep(.1)
                raise TimeoutError("Scan did not complete")
            # Real packaged cancellation with read-only corpus, then a successful retry.
            api("/api/scan",{"full":True});api("/api/cancel",{});wait_job("cancelled")
            assert not (workspace/"state"/"approved"/"current.json").exists()
            assert not (workspace/"state"/"staging"/"latest.json").exists()
            assert api("/api/status")["settings"]["requires_rescan"]
            t=time.perf_counter();api("/api/scan",{"full":True});parallel=wait_job();parallel_time=time.perf_counter()-t
            actual=signature(workspace/"state"/"staging"/parallel["run_id"])
            assert expected==actual,"Serial/parallel stage output mismatch"
            assert parallel["phase_times"]["extraction_workers"]>1
            assert parallel["input_proof"] and not parallel.get("requires_rescan")
            proposal=api("/api/review");api("/api/save-choices",{"run":proposal["run"],"files":proposal["files"],"draft_version":proposal["draft_version"]})
            approval=api("/api/approve",{"run":proposal["run"]})
            assert approval["verified_unchanged_eligible"],approval
            # Compare indexed API results against the existing Python query semantics.
            scratch=App(root/"query-reference")
            try:
                import shutil
                shutil.copytree(workspace/"state"/"approved",scratch.state/"approved")
                indexed_tests=0;query_times=[]
                for scope in ("events","documents","revisions"):
                    reference=scratch.view(scope);fields=ordered_fields(reference,scope)
                    first=reference[0];doc=str(first.get("Document No.",""))
                    for query,col,sort,desc,filters in [
                        ('"'+doc+'"',"Document No.","",False,{}),
                        (doc[:4],"","Document No.",True,{}),
                        ("-__no_match_fixture__","", "Latest Revision" if scope=="documents" else "Revision",False,{}),
                        ("","", "Document No.",False,{"Project No.":[str(first["Project No."])]}),
                        ("__no_match_fixture__","","",False,{})]:
                        from urllib.parse import urlencode
                        params=urlencode({"scope":scope,"q":query,"column":col,"sort":sort,"desc":int(desc),"filters":json.dumps(filters)})
                        t=time.perf_counter();answer=api("/api/index?"+params);query_times.append(time.perf_counter()-t)
                        expected_rows=apply_filters(filter_rows(reference,query,col,sort,desc),filters,fields)
                        assert answer["count"]==len(expected_rows) and answer["rows"]==expected_rows[:100],(scope,col,sort,indexed_tests)
                        indexed_tests+=1
                assert api("/api/index-view",{"name":"Test view","scope":"events","columns":["Document No."],"filters":{}})["Test view"]["columns"]==["Document No."]
            finally:scratch.invalidate();scratch.lock.close()
            state=workspace/"state";ptr=(state/"approved"/"current.json").read_bytes();latest=(state/"staging"/"latest.json").read_bytes()
            times=[]
            for _ in range(3):
                t=time.perf_counter();api("/api/scan",{"full":False});out=wait_job();times.append(time.perf_counter()-t)
                assert out["status"]=="UNCHANGED" and not out["proposal_created"],out
                assert (state/"approved"/"current.json").read_bytes()==ptr and (state/"staging"/"latest.json").read_bytes()==latest
            # Modify config in scratch only; a fast-path hit is no longer valid.
            config=workspace/"config"/"source_exclusions.csv"
            old=config.read_bytes() if config.exists() else None
            config.write_bytes((old or b"")+b"\n")
            t=time.perf_counter();api("/api/scan",{"full":False});changed=wait_job()
            assert changed.get("run_id") and changed.get("status")!="UNCHANGED"
            assert (state/"approved"/"current.json").read_bytes()==ptr
            report={"serial_stage_seconds":round(serial_time,3),"packaged_parallel_job_seconds":round(parallel_time,3),
                    "workers":parallel["phase_times"]["extraction_workers"],"events":parallel["staged_records"],
                    "serial_parallel_exact_stage_parity":True,"packaged_unchanged_api_median_seconds":round(statistics.median(times),3),
                    "packaged_cancel_retry_passed":True,"no_new_proposal":True,"config_change_invalidated_gate":True,"indexed_query_parity_cases":indexed_tests,"indexed_api_median_seconds_including_cold_builds":round(statistics.median(query_times),6),"source_hashes_unchanged":hashes(data)==before}
            assert report["source_hashes_unchanged"]
            api("/api/exit",{});assert proc.wait(timeout=15)==0
            Path("v2-fast-path-metrics.json").write_text(json.dumps(report,indent=2),encoding="utf-8");print(json.dumps(report,indent=2))
        finally:
            if proc.poll() is None:proc.terminate();proc.wait(timeout=15)
if __name__=="__main__":main(Path(sys.argv[1]).resolve())
