"""Local verified recovery bundles. Never read/write the source DATA directory."""
from __future__ import annotations
import hashlib, json, os, shutil, time, uuid
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

AREAS=("state/approved","state/staging","state/cache","state/logs","state/user_flags","state/custom-fields.json","state/derived/custom-fields.json","config","config-history","settings.json","custom-fields.json","derived/custom-fields.json","migration-evidence","migration-manifest.json")
MAX_FILES=100_000
MAX_FILE_BYTES=512*1024**2
MAX_TOTAL_BYTES=2*1024**3
def digest(path):
    h=hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda:f.read(1024*1024),b""):h.update(chunk)
    return h.hexdigest()
def assert_workspace_paths(root):
    root=Path(root).absolute()
    for path in (root,root/"state",root/"config",root/"config-history"):
        for parent in (path,*path.parents):
            if parent.is_symlink() or _is_reparse(parent):raise ValueError("Symlink or reparse point in workspace ancestry: "+str(parent))
    return root.resolve()

def inventory(root):
    root=assert_workspace_paths(root); out={}; total=0
    for area in AREAS:
        p=root/area
        if p.is_symlink() or _is_reparse(p):raise ValueError("Symlink or reparse point in recovery area: "+str(p))
        if not p.exists():continue
        for f in ([p] if p.is_file() else p.rglob("*")):
            if f.is_symlink() or _is_reparse(f):raise ValueError("Symlink or reparse point in recovery area: "+str(f))
            if not f.is_file():continue
            if not f.resolve().is_relative_to(root):raise ValueError("Recovery path escapes workspace")
            size=f.stat().st_size
            if size>MAX_FILE_BYTES:raise ValueError("Recovery file exceeds size limit: "+str(f))
            total+=size
            if total>MAX_TOTAL_BYTES or len(out)>=MAX_FILES:raise ValueError("Recovery inventory exceeds configured bounds")
            out[f.relative_to(root).as_posix()]={"bytes":size,"sha256":digest(f)}
    return out
def _is_reparse(path):
    try:return bool(path.lstat().st_file_attributes & 0x400)
    except (AttributeError,OSError):return False
def verify_bundle(path):
    path=Path(path)
    if path.is_symlink() or _is_reparse(path) or path.stat().st_size>MAX_TOTAL_BYTES:raise ValueError("Recovery archive is unsafe or exceeds size limits")
    with ZipFile(path) as z:
        if len(z.infolist())>MAX_FILES+1 or sum(i.file_size for i in z.infolist())>MAX_TOTAL_BYTES:raise ValueError("Recovery bundle exceeds size limits")
        manifest=json.loads(z.read("recovery-manifest.json"))
        files=manifest["files"]
        if len(z.namelist())!=len(set(z.namelist())) or set(z.namelist())!=set(files)|{"recovery-manifest.json"}:raise ValueError("Recovery manifest mismatch")
        for name,meta in files.items():
            parts=name.split("/")
            if any(p in ("",".","..") for p in parts) or "\\" in name or ":" in name or not any(name==a or name.startswith(a+"/") for a in AREAS):raise ValueError("Unsafe recovery path")
            if not isinstance(meta,dict) or type(meta.get("bytes")) is not int or meta["bytes"]<0 or meta["bytes"]>MAX_FILE_BYTES:raise ValueError("Invalid recovery file size")
            h=hashlib.sha256();size=0
            with z.open(name) as f:
                for chunk in iter(lambda:f.read(1024*1024),b""):h.update(chunk);size+=len(chunk)
            if h.hexdigest()!=meta["sha256"] or size!=meta["bytes"]:raise ValueError("Recovery hash mismatch")
        return manifest
def create_bundle(root,reason,keep=None):
    root=assert_workspace_paths(root); dest=root/"backups"
    if keep is None:
        settings=root/"settings.json"
        keep=json.loads(settings.read_text(encoding="utf-8")).get("backup_retention",10) if settings.exists() else 10
    if type(keep) is not int or not 1<=keep<=100:raise ValueError("Backup retention must be an integer from 1 to 100")
    if dest.is_symlink() or _is_reparse(dest):raise ValueError("Backup destination must not be a reparse point")
    dest.mkdir(exist_ok=True)
    before=inventory(root)
    required=sum(item["bytes"] for item in before.values())
    if shutil.disk_usage(dest).free<required+16*1024**2:raise ValueError("Insufficient free space for a verified recovery bundle")
    name=time.strftime("%Y%m%d-%H%M%S")+"-"+uuid.uuid4().hex[:8]+".zip"
    tmp=dest/(name+".tmp"); final=dest/name
    try:
        with ZipFile(tmp,"w",ZIP_DEFLATED,compresslevel=3) as z:
            for name,meta in before.items():z.write(root/name,name)
            z.writestr("recovery-manifest.json",json.dumps({"version":1,"created_at":time.time(),"reason":str(reason),"files":before,"excluded_disposable":["state/source_cache","state/profile","derived indexes"]},indent=2))
        verify_bundle(tmp)
        if inventory(root)!=before:raise ValueError("Workspace changed during backup")
        os.replace(tmp,final)
        apply_retention(dest,keep=keep)
        return {"path":str(final),"files":len(before),"bytes":final.stat().st_size}
    finally:tmp.unlink(missing_ok=True)
def apply_retention(directory,keep=10):
    """Retain newest verified recovery archives; leave unknown files untouched."""
    directory=Path(directory)
    if type(keep) is not int or keep<1:raise ValueError("Retention must keep at least one verified backup")
    if directory.is_symlink() or _is_reparse(directory):raise ValueError("Backup directory must not be a reparse point")
    archives=[]
    for path in directory.glob("*.zip"):
        if path.is_symlink() or _is_reparse(path):raise ValueError("Backup archive is a reparse point")
        manifest=verify_bundle(path)
        created=manifest.get("created_at")
        if created is None:created=path.stat().st_mtime
        if not isinstance(created,(int,float)):raise ValueError("Invalid recovery timestamp")
        archives.append((float(created),path))
    archives.sort(key=lambda row:(row[0],row[1].name),reverse=True)
    for _,path in archives[keep:]:path.unlink()
    return [str(p) for _,p in archives[:keep]]
def restore_to_new_workspace(path,target):
    """Restore only to a new isolated folder, never overwrite active or legacy state."""
    target=Path(target).absolute()
    if any(_is_reparse(p) for p in (target,*target.parents)):raise ValueError("Restore destination traverses a reparse point")
    if target.exists():raise ValueError("Restore destination already exists")
    manifest=verify_bundle(path)
    scratch=target.with_name(target.name+".restore-"+uuid.uuid4().hex[:8])
    try:
        scratch.mkdir(parents=True,exist_ok=False)
        with ZipFile(path) as z:
            for name in manifest["files"]:
                dest=scratch/name;dest.parent.mkdir(parents=True,exist_ok=True)
                with z.open(name) as src,dest.open("xb") as out:shutil.copyfileobj(src,out)
        if inventory(scratch)!=manifest["files"]:raise ValueError("Restore reconciliation failed")
        if target.exists():raise ValueError("Restore destination appeared")
        scratch.rename(target)
    finally:
        if scratch.exists():shutil.rmtree(scratch)
    return {"workspace":str(target),"files":len(manifest["files"])}

def reset_workspace(root, confirmation):
    """Back up then clear only disposable approved/staged runtime views."""
    root=assert_workspace_paths(root)
    if confirmation != "RESET":raise ValueError('Type RESET to confirm clearing approved and staged runtime state.')
    bundle=create_bundle(root,"Before runtime reset")
    state=root/"state"
    for area in ("approved","staging","cache","profile","source_cache"):
        target=state/area
        if target.exists():
            if target.is_symlink() or _is_reparse(target):raise ValueError("Runtime reset area is a reparse point: "+str(target))
            shutil.rmtree(target)
    _append_recovery_history(state,{"event":"RESET","backup":bundle["path"],"at":time.time()})
    return {"decision":"RESET","backup":bundle["path"],"preserved":["settings","config","config-history","logs","audit history"]}

def undo_approval(state_dir, confirmation):
    """Safely repoint to the prior immutable approved version, retaining all versions."""
    raw_state=Path(state_dir).absolute()
    if raw_state.is_symlink() or _is_reparse(raw_state):raise ValueError("State directory is a reparse point")
    state=raw_state.resolve();root=state.parent
    if confirmation != "UNDO":raise ValueError('Type UNDO to confirm restoring the prior approved index.')
    current_path=state/"approved"/"current.json"
    if current_path.is_symlink() or _is_reparse(current_path):raise ValueError("Approved pointer is a reparse point")
    current=json.loads(current_path.read_text(encoding="utf-8")) if current_path.exists() else {}
    if not current.get("run_id"):raise ValueError("There is no approved update to undo.")
    import re
    if not re.fullmatch(r"[A-Za-z0-9_-]+",str(current["run_id"])):raise ValueError("Unsafe approved run pointer")
    versions=state/"approved"/"versions"
    if versions.is_symlink() or _is_reparse(versions):raise ValueError("Approved versions directory is a reparse point")
    for child in versions.iterdir():
        if child.is_symlink() or _is_reparse(child):raise ValueError("Approved version is a reparse point")
        if child.is_dir():
            record=json.loads((child/"approval.json").read_text(encoding="utf-8"))
            if child.name in (".","..") or record.get("run_id")!=child.name or not (child/"records.jsonl").is_file():
                raise ValueError("Approved version inventory is inconsistent")
    history=state/"logs"/"history.jsonl"
    if history.is_symlink() or _is_reparse(history):raise ValueError("Approval history is a reparse point")
    predecessors={};active=""
    if history.exists():
        for line in history.read_text(encoding="utf-8").splitlines():
            row=json.loads(line)
            if row.get("event")=="APPROVED" and row.get("decision")=="APPROVED":
                run=str(row.get("run_id",""))
                if active!=run:predecessors[run]=active
                active=run
            elif row.get("event")=="UNDO_APPROVAL":active=str(row.get("restored_run_id",""))
            elif row.get("event")=="RESET":active=""
    if active!=str(current["run_id"]):raise ValueError("Approval history does not reconcile with the current pointer")
    prior=predecessors.get(active,"")
    if not prior:raise ValueError("There is no previous approved version available to restore.")
    if not re.fullmatch(r"[A-Za-z0-9_-]+",prior):raise ValueError("Unsafe previous approved run")
    if not (versions/prior/"records.jsonl").is_file():raise ValueError("Previous approved version is incomplete")
    backup=create_bundle(root,"Before approval undo")
    _write_pointer(current_path,{"run_id":prior,"version_path":f"versions/{prior}"})
    result={"decision":"UNDO_APPROVAL","undone_run_id":str(current["run_id"]),"restored_run_id":prior,"undone_at":time.strftime("%Y-%m-%dT%H:%M:%SZ",time.gmtime())}
    try:_append_recovery_history(state,{"event":"UNDO_APPROVAL",**result})
    except OSError:
        _write_pointer(current_path,current)
        raise
    result["backup"]=backup["path"]
    return result

def _append_recovery_history(state,payload):
    path=Path(state)/"logs"/"history.jsonl";path.parent.mkdir(parents=True,exist_ok=True)
    with path.open("a",encoding="utf-8",newline="\n") as stream:
        stream.write(json.dumps(payload,sort_keys=True)+"\n")

def _write_pointer(path,payload):
    path=Path(path);temp=path.with_name(path.name+".tmp-"+uuid.uuid4().hex)
    try:
        with temp.open("x",encoding="utf-8",newline="\n") as stream:
            json.dump(payload,stream,sort_keys=True);stream.write("\n");stream.flush();os.fsync(stream.fileno())
        os.replace(temp,path)
    finally:temp.unlink(missing_ok=True)
