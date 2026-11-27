"""Conservative fast unchanged gate. Existing JSONL snapshots remain authoritative."""
from __future__ import annotations
import hashlib, json, sys, os
from pathlib import Path
from nmdc_profiler.update_engine import scan_sources, sha256_file, _write_json_atomic, fingerprint_paths, _append_log, utc_now, DEFAULT_PARSER_VERSION

VERSION="verified-unchanged-v1"
def read(p):
    return json.loads(p.read_text(encoding="utf-8"))
def build_identity():
    if getattr(sys,"frozen",False):
        return sha256_file(Path(sys.executable))
    root=Path(__file__).resolve().parents[1]
    h=hashlib.sha256()
    for p in sorted([*root.glob("nmdc_profiler/**/*.py"),*root.glob("v2/*.py")]):
        h.update(p.relative_to(root).as_posix().encode());h.update(sha256_file(p).encode())
    return h.hexdigest()
def inputs(data,config,build):
    rows,flags=scan_sources(data)
    if flags:raise ValueError("Source hash errors")
    # Repeat inventory and metadata verification after hashing. A live external
    # producer can still modify files after this check; this is a scan observation,
    # not a lock of the user's source folder.
    paths={r["relative_path"]: [r["sha256"],r["size"],r["modified_ns"]] for r in rows}
    metadata={}
    for p in Path(data).rglob("*"):
        if p.is_file() and p.suffix.lower() in (".xlsx",".xlsm") and not p.name.startswith("~$"):
            stat=p.stat()
            metadata[p.relative_to(data).as_posix()]=[stat.st_size,stat.st_mtime_ns]
    if metadata!={k:v[1:] for k,v in paths.items()}:raise ValueError("Source changed during verification")
    return {"version":VERSION,"build":build,"parser":DEFAULT_PARSER_VERSION,"data":str(Path(data).resolve()),
            "source":paths,"config":{p.name:sha256_file(p) for p in sorted(Path(config).glob("*.csv"))}}
def stage_proof(state,run,before,after,config):
    if before!=after:return False
    stage=state/"staging"/run
    summary=read(stage/"summary.json")
    if summary.get("review_flags") or summary.get("blocking_flags") or read(stage/"flags.json"):return False
    manifest=read(stage/"manifest.json")
    # A metadata-only source mirror may be stale; never certify it unless actual source hashes agree.
    actual={f["relative_path"]:[f["sha256"],f["size"]] for f in manifest["files"]}
    if actual!={k:v[:2] for k,v in after["source"].items()}:return False
    if manifest.get("data_dir")!=after["data"] or manifest.get("parser_version")!=after["parser"]:return False
    if manifest.get("config_fingerprint")!=fingerprint_paths([config/"classification_rules.csv",config/"project_identity_overrides.csv"]):return False
    _write_json_atomic(stage/"input_proof.json",after)
    return True
def snapshot(state):
    ptr=state/"approved"/"current.json"; run=read(ptr)["run_id"]
    if not run or Path(run).name!=run:raise ValueError("Invalid approved run")
    root=state/"approved"/"versions"/run; stage=state/"staging"/run
    latest=state/"staging"/"latest.json"
    if read(latest).get("run_id")!=run:raise ValueError("Pending proposal")
    approval=read(root/"approval.json")
    if approval.get("decision")!="APPROVED" or approval.get("deferred_files") or approval.get("allow_conflicts"):raise ValueError("Not a complete approved baseline")
    if read(stage/"decision.json").get("decision")!="APPROVED" or read(stage/"flags.json"):raise ValueError("Pending decision or flags")
    user=state/"user_flags"/"flags.jsonl"
    if user.exists() and user.read_text(encoding="utf-8").strip():raise ValueError("User flags require standard review")
    review=read(root/"file_review.json")
    if any(c.get("accept") is not True for c in review.get("choices",{}).values()):raise ValueError("Deferred choice")
    files=[ptr,latest,stage/"decision.json",stage/"flags.json",stage/"input_proof.json"]+[root/x for x in ("records.jsonl","manifest.json","cache_index.json","approval.json","file_review.json")]
    for rel in read(root/"cache_index.json").values():
        p=(state/rel).resolve()
        if not p.is_relative_to((state/"cache").resolve()):raise ValueError("Invalid extraction cache path")
        files.extend([p,p.with_suffix(".flags.json")])
    return {"run":run,"hashes":{str(p.relative_to(state)):sha256_file(p) for p in files}}
def certify(state,data,config,build):
    path=state/"verified-unchanged.json"
    path.unlink(missing_ok=True)
    try:
        snap=snapshot(state);proof=read(state/"staging"/snap["run"]/"input_proof.json")
        if proof!=inputs(data,config,build):return False
        _write_json_atomic(path,{"inputs":proof,"snapshot":snap})
        return True
    except (OSError,ValueError,KeyError,TypeError):return False
def check(state,data,config,build):
    try:
        cert=read(state/"verified-unchanged.json")
        if cert["inputs"]!=inputs(data,config,build) or cert["snapshot"]!=snapshot(state):return None
        run=cert["snapshot"]["run"]
        result={"status":"UNCHANGED","approved_run_id":run,"sources_verified":len(cert["inputs"]["source"]),"proposal_created":False}
        _append_log(state,{"event":"VERIFIED_UNCHANGED","at":utc_now(),"run_id":run,**result})
        return result
    except (OSError,ValueError,KeyError,TypeError):return None
