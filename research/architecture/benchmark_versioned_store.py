"""Real-corpus immutable-version/transaction research. Scratch state only."""
import hashlib,json,statistics,tempfile,time
from collections import defaultdict
from pathlib import Path
from v2.launcher import App
from nmdc_profiler.runtime_engine import stage_runtime_update
from nmdc_profiler.source_access import install_resilient_source_access
from nmdc_profiler.layout_compat import install_layout_compatibility
from nmdc_profiler.excel_bridge import load_latest_stage
from versioned_store import Store
def measured(fn):
    t=time.perf_counter();value=fn();return value,round(time.perf_counter()-t,6)
def must_fail(fn,expected=ValueError):
    try:fn()
    except expected:return
    raise AssertionError("Unsafe operation accepted")
def main():
    source=(Path(__file__).resolve().parents[2]/"DATA").resolve()
    def hashes():return {p.relative_to(source).as_posix():hashlib.sha256(p.read_bytes()).hexdigest() for p in source.rglob("*") if p.is_file() and p.suffix.lower() in (".xlsx",".xlsm") and not p.name.startswith("~$")}
    before=hashes();assert before
    install_resilient_source_access();install_layout_compatibility()
    with tempfile.TemporaryDirectory(prefix="nmdc-versioned-") as d:
        root=Path(d);app=App(root/"legacy")
        try:
            staged=stage_runtime_update(data_dir=source,state_dir=app.state,config_dir=app.config,full_rescan=True,reuse_raw_profiles=True)
            assert staged["review_flags"]==staged["blocking_flags"]==0
            data=load_latest_stage(app.state);rows=data["records"];manifest=data["manifest"]
            # Disposable research baseline; never approve or alter user's live data.
            store=Store(root/"versions.db")
            first,import_time=measured(lambda:store.import_legacy(manifest,rows))
            assert store.rows()==rows
            meta={r["relative_path"]:r for r in manifest["files"]};assert store.metadata()==meta
            groups=defaultdict(list)
            for r in rows:groups[r["Source File"]].append(r)
            big,remove,defer=sorted(groups,key=lambda k:len(groups[k]),reverse=True)[:3]
            modified=[dict(r,**{"Document Title":str(r.get("Document Title",""))+" [research]"}) for r in groups[big]]
            new_source="__research_new__.xlsx";new=[dict(modified[0],**{"Source File":new_source,"Event_Key":"__research_new_event__"})]
            changes={big:(dict(meta[big],research_version=2),modified),remove:None,defer:None,new_source:({"relative_path":new_source},new)}
            proposal,stage_time=measured(lambda:store.stage(changes))
            must_fail(lambda:store.approve(proposal,0))
            choices={big:(True,"Accept modified"),remove:(True,"Accept removal"),defer:(False,"Defer removal"),new_source:(False,"Defer new")}
            rev=store.save(proposal,0,choices)
            must_fail(lambda:store.save(proposal,0,choices))
            must_fail(lambda:store.approve(proposal,rev,interrupt=True),RuntimeError)
            assert store.current()==first and store.rows()==rows
            second,approval_time=measured(lambda:store.approve(proposal,rev))
            target={r["Event_Key"]:r for r in rows if r["Source File"] not in (big,remove)}
            target.update({r["Event_Key"]:r for r in modified})
            assert {r["Event_Key"]:r for r in store.rows()}==target
            assert store.rows(first)==rows and store.metadata()[defer]==meta[defer]
            assert new_source not in store.metadata() and remove not in store.metadata()
            assert store.metadata()[big]["research_version"]==2
            # Attempt to insert a cross-source duplicate. It must never commit.
            duplicate=dict(modified[0],**{"Source File":new_source})
            p=store.stage({new_source:({"relative_path":new_source},[duplicate])})
            rv=store.save(p,0,{new_source:(True,"Duplicate test")})
            must_fail(lambda:store.approve(p,rv))
            assert store.current()==second
            p=store.stage({new_source:({"relative_path":new_source},new)},flags=[{"level":"CONFLICT","code":"TEST"}])
            rv=store.save(p,0,{new_source:(True,"Flag test")})
            must_fail(lambda:store.approve(p,rv))
            stale=store.stage({new_source:({"relative_path":new_source},new)})
            rv=store.save(stale,0,{new_source:(True,"Stale test")})
            store.stage({new_source:({"relative_path":new_source},new)})
            must_fail(lambda:store.approve(stale,rv))
            _,undo_time=measured(store.undo)
            assert store.current()==first and store.rows()==rows and store.rows(second)
            assert store.db.execute("PRAGMA integrity_check").fetchone()[0]=="ok"
            assert not store.db.execute("PRAGMA foreign_key_check").fetchall()
            store.db.close()
            report={"scope":"real-corpus research only; complete migration/flag resolution/physical crash not tested",
                    "events_imported":len(rows),"sources":len(meta),"all_imported_payloads_and_metadata_match":True,
                    "initial_import_seconds":import_time,"changed_rows":len(modified),"stage_delta_seconds":stage_time,
                    "approval_membership_transaction_seconds":approval_time,"undo_pointer_seconds":undo_time,
                    "accepted_change_and_removal":True,"deferred_removal_and_new":True,"old_version_preserved":True,
                    "unsaved_stale_duplicate_and_flag_guards":True,"simulated_transaction_interrupt_rolled_back":True,
                    "source_hashes_unchanged":hashes()==before}
            assert report["source_hashes_unchanged"]
            Path("versioned-store-metrics.json").write_text(json.dumps(report,indent=2),encoding="utf-8")
            print(json.dumps(report,indent=2))
        finally:app.lock.close()
if __name__=="__main__":main()
