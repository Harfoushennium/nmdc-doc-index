"""Architecture research only. Uses read-only repository DATA and isolated scratch state.
Benchmark source-independent extraction, content-verified no-op, and delta transactions.
No production engine changes and no corpus content in results.
"""
from __future__ import annotations
import csv, hashlib, json, multiprocessing as mp, os, sqlite3, statistics, tempfile, time
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path
from nmdc_profiler.runtime_engine import build_runtime_catalog, stage_runtime_update
from nmdc_profiler.source_access import install_resilient_source_access
from nmdc_profiler.layout_compat import install_layout_compatibility
from nmdc_profiler.update_engine import scan_sources, sha256_file, approve_stage, _load_approved_state
from nmdc_profiler.file_review import affected_files, save_file_review
from nmdc_profiler.excel_bridge import load_latest_stage
from v2.launcher import App
CATALOG=None
def init(catalog):
    global CATALOG
    install_resilient_source_access();install_layout_compatibility()
    CATALOG=catalog
def process(rel):
    records=list(CATALOG.process(CATALOG.data_dir/rel,rel))
    flags=list(CATALOG.drain_processor_flags())
    return rel,records,flags
def timed(fn):
    t=time.perf_counter();r=fn();return r,time.perf_counter()-t
def median(fn,n=5):
    vals=[timed(fn)[1] for _ in range(n)];return round(statistics.median(vals),6)
def digest(x):
    return hashlib.sha256(json.dumps(x,ensure_ascii=False,sort_keys=True,separators=(",",":"),default=str).encode()).hexdigest()
def source_identity(root):
    rows,flags=scan_sources(root)
    if flags:raise ValueError("Source unreadable")
    return {r["relative_path"]:[r["sha256"],r["size"],r["modified_ns"]] for r in rows}
def config_identity(root):
    return {p.name:sha256_file(p) for p in sorted(root.glob("*.csv"))}
def snapshot_identity(state):
    ptr=state/"approved"/"current.json"
    run=json.loads(ptr.read_text(encoding="utf-8"))["run_id"]
    root=state/"approved"/"versions"/run
    # Verify canonical records, manifest, cache index, approval and referenced extraction caches.
    files=[ptr]+[root/x for x in ("records.jsonl","manifest.json","cache_index.json","approval.json","file_review.json")]
    cache=json.loads((root/"cache_index.json").read_text(encoding="utf-8"))
    for rel in cache.values():
        p=(state/rel).resolve()
        if not p.is_relative_to(state.resolve()):raise ValueError("Invalid cache path")
        files.append(p)
        flags=p.with_suffix(".flags.json")
        if flags.exists():files.append(flags)
    return {str(p.relative_to(state)):sha256_file(p) for p in files}
def gate(data,config,state,baseline):
    return (source_identity(data)==baseline["source"] and config_identity(config)==baseline["config"]
            and snapshot_identity(state)==baseline["snapshot"])
def main():
    data=(Path(__file__).resolve().parents[2]/"DATA").resolve()
    before=source_identity(data)
    assert before
    report={"scope":"real repository workbook architecture research; Windows scratch state; no production change",
            "workbooks":len(before),"logical_cpus":os.cpu_count()}
    install_resilient_source_access();install_layout_compatibility()
    with tempfile.TemporaryDirectory(prefix="nmdc-architecture-") as d:
        root=Path(d);app=App(root/"workspace")
        try:
            result,elapsed=timed(lambda:stage_runtime_update(data_dir=data,state_dir=app.state,config_dir=app.config,full_rescan=True,reuse_raw_profiles=True))
            stage=load_latest_stage(app.state)
            assert result["blocking_flags"]==0 and result["review_flags"]==0
            files=affected_files(app.state,result["run_id"]);choices=root/"choices.csv"
            with choices.open("w",encoding="utf-8",newline="") as f:
                w=csv.DictWriter(f,fieldnames=["Source File","Accept Update?","Owner Note"]);w.writeheader()
                for item in files:w.writerow({"Source File":item["Source File"],"Accept Update?":"TRUE","Owner Note":"Scratch benchmark"})
            save_file_review(app.state,result["run_id"],choices)
            approve_stage(app.state,result["run_id"],require_file_review=True)
            manifest,_,rows=_load_approved_state(app.state)
            assert len(rows)==len({r["Event_Key"] for r in rows}) and rows
            report["events"]=len(rows);report["full_stage_seconds"]=round(elapsed,3)
            # Same engine/corpus, compare no-op proof to normal unchanged staging.
            # Research only: a production gate must ALSO enforce parser build/version, active proposal,
            # pending flags, approved all-file baseline, coherent source enumeration and audit policy.
            baseline={"source":before,"config":config_identity(app.config),"snapshot":snapshot_identity(app.state)}
            assert gate(data,app.config,app.state,baseline)
            report["verified_noop_gate_median_seconds"]=median(lambda:gate(data,app.config,app.state,baseline))
            report["current_incremental_median_seconds"]=median(lambda:stage_runtime_update(data_dir=data,state_dir=app.state,config_dir=app.config,full_rescan=False,reuse_raw_profiles=True),3)
            # Verify detection of changed config and changed canonical evidence on isolated workspace.
            cfg=app.config/"source_exclusions.csv";old=cfg.read_bytes() if cfg.exists() else None
            cfg.write_bytes((old or b"")+b"\n")
            assert not gate(data,app.config,app.state,baseline)
            if old is None:cfg.unlink()
            else:cfg.write_bytes(old)
            assert gate(data,app.config,app.state,baseline)
            ptr=app.state/"approved"/"current.json";old=ptr.read_bytes()
            ptr.write_bytes(old+b"\n");assert not gate(data,app.config,app.state,baseline);ptr.write_bytes(old)
            report["no_op_guard_checks"]="PASS (source/config/snapshot integrity; config and pointer change)"
            # Compare per-source extraction with exact records and processor flags.
            mirror=app.state/"source_cache"/"files"
            catalog=build_runtime_catalog(mirror,app.config/"classification_rules.csv",app.config/"project_identity_overrides.csv",app.config/"source_exclusions.csv",root/"profile")
            sources=sorted(k for k in catalog.work_items if catalog.selection_statuses.get(k)=="SELECTED")
            init(catalog)
            expected,t=timed(lambda:[process(rel) for rel in sources])
            serial_digest=digest(expected)
            report["serial_extraction_seconds"]=round(t,4);report["selected_workbooks"]=len(sources)
            report["parallel_extraction"]=[]
            for workers in (2,4):
                def parallel():
                    with ProcessPoolExecutor(max_workers=workers,mp_context=mp.get_context("spawn"),initializer=init,initargs=(catalog,)) as pool:
                        return list(pool.map(process,sources,chunksize=1))
                actual,elapsed=timed(parallel)
                parity=digest(actual)==serial_digest
                report["parallel_extraction"].append({"workers":workers,"seconds":round(elapsed,4),"exact_records_and_flags":parity,"speed_ratio":round(t/elapsed,3)})
                assert parity
            # Benchmark a disposable indexed per-source version representation.
            # Initial construction includes JSON serialization; delta is one existing source only.
            db=sqlite3.connect(root/"delta.sqlite")
            db.execute("PRAGMA foreign_keys=ON")
            db.executescript("CREATE TABLE events(event TEXT PRIMARY KEY,source TEXT NOT NULL,doc TEXT,payload TEXT);CREATE INDEX by_source ON events(source);CREATE INDEX by_doc ON events(doc);")
            def values(seq):
                return [(r["Event_Key"],r.get("Source File",""),r.get("Document No.",""),json.dumps(r,ensure_ascii=False,sort_keys=True)) for r in seq]
            def build():
                with db:db.executemany("INSERT INTO events VALUES(?,?,?,?)",values(rows))
            _,elapsed=timed(build);report["sqlite_initial_build_seconds"]=round(elapsed,4)
            source=max({r["Source File"] for r in rows},key=lambda k:sum(r["Source File"]==k for r in rows))
            changed=[dict(r,**{"Document Title":str(r.get("Document Title",""))+" [research]"}) for r in rows if r["Source File"]==source]
            originals=[r for r in rows if r["Source File"]==source]
            def delta():
                with db:
                    db.execute("DELETE FROM events WHERE source=?",(source,))
                    db.executemany("INSERT INTO events VALUES(?,?,?,?)",values(changed))
            _,elapsed=timed(delta)
            report["one_source_rows"]=len(changed);report["sqlite_one_source_transaction_seconds"]=round(elapsed,6)
            target={r["Event_Key"]:r for r in rows};target.update({r["Event_Key"]:r for r in changed})
            assert {key:json.loads(payload) for key,payload in db.execute("SELECT event,payload FROM events")}==target
            with db:
                db.execute("DELETE FROM events WHERE source=?",(source,));db.executemany("INSERT INTO events VALUES(?,?,?,?)",values(originals))
            assert {key:json.loads(payload) for key,payload in db.execute("SELECT event,payload FROM events")}=={r["Event_Key"]:r for r in rows}
            report["delta_and_rollback_payload_parity"]=True
            doc=rows[0]["Document No."]
            report["indexed_doc_query_median_seconds"]=median(lambda:db.execute("SELECT payload FROM events WHERE doc=?",(doc,)).fetchall(),30)
            assert db.execute("PRAGMA integrity_check").fetchone()[0]=="ok"
            db.close()
        finally:app.lock.close()
    report["source_hashes_unchanged"]=source_identity(data)==before
    assert report["source_hashes_unchanged"]
    Path("architecture-research-metrics.json").write_text(json.dumps(report,indent=2),encoding="utf-8")
    print(json.dumps(report,indent=2))
if __name__=="__main__":mp.freeze_support();main()
