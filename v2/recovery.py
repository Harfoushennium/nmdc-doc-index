"""Local verified recovery bundles. Never read/write the source DATA directory."""
from __future__ import annotations
import hashlib, json, os, shutil, time, uuid
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

AREAS=("state/approved","state/staging","state/cache","state/logs","state/user_flags","config","config-history","settings.json")
def digest(path):
    h=hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda:f.read(1024*1024),b""):h.update(chunk)
    return h.hexdigest()
def inventory(root):
    root=Path(root).resolve(); out={}
    for area in AREAS:
        p=root/area
        if not p.exists():continue
        for f in ([p] if p.is_file() else p.rglob("*")):
            if f.is_symlink():raise ValueError("Symlink in recovery area: "+str(f))
            if not f.is_file():continue
            if not f.resolve().is_relative_to(root):raise ValueError("Recovery path escapes workspace")
            out[f.relative_to(root).as_posix()]={"bytes":f.stat().st_size,"sha256":digest(f)}
    return out
def verify_bundle(path):
    with ZipFile(path) as z:
        manifest=json.loads(z.read("recovery-manifest.json"))
        files=manifest["files"]
        if len(z.namelist())!=len(set(z.namelist())) or set(z.namelist())!=set(files)|{"recovery-manifest.json"}:raise ValueError("Recovery manifest mismatch")
        for name,meta in files.items():
            parts=name.split("/")
            if any(p in ("",".","..") for p in parts) or "\\" in name or ":" in name or not any(name==a or name.startswith(a+"/") for a in AREAS):raise ValueError("Unsafe recovery path")
            h=hashlib.sha256();size=0
            with z.open(name) as f:
                for chunk in iter(lambda:f.read(1024*1024),b""):h.update(chunk);size+=len(chunk)
            if h.hexdigest()!=meta["sha256"] or size!=meta["bytes"]:raise ValueError("Recovery hash mismatch")
        return manifest
def create_bundle(root,reason):
    root=Path(root).resolve(); dest=root/"backups";dest.mkdir(exist_ok=True)
    before=inventory(root)
    name=time.strftime("%Y%m%d-%H%M%S")+"-"+uuid.uuid4().hex[:8]+".zip"
    tmp=dest/(name+".tmp"); final=dest/name
    try:
        with ZipFile(tmp,"w",ZIP_DEFLATED,compresslevel=3) as z:
            for name,meta in before.items():z.write(root/name,name)
            z.writestr("recovery-manifest.json",json.dumps({"version":1,"reason":str(reason),"files":before,"excluded_disposable":["state/source_cache","state/profile","derived indexes"]},indent=2))
        verify_bundle(tmp)
        if inventory(root)!=before:raise ValueError("Workspace changed during backup")
        os.replace(tmp,final)
        return {"path":str(final),"files":len(before),"bytes":final.stat().st_size}
    finally:tmp.unlink(missing_ok=True)
def restore_to_new_workspace(path,target):
    """Restore only to a new isolated folder, never overwrite active or legacy state."""
    target=Path(target).resolve()
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
