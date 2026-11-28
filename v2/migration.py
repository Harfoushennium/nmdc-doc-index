"""Confirmed, isolated migration import for the read-only V1 package preview."""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import tempfile
from pathlib import Path

from v2.migration_preview import preview_package
from nmdc_profiler.file_review import _baseline


def _fingerprint(report):
    return report.get("preview_sha256")


def import_package(package, target, confirmed_fingerprint, *, source_root=None):
    """Import a compatible legacy package into a new workspace after exact preview confirmation.

    The legacy package is never modified. A failed/incompatible import leaves no target.
    """
    package = Path(package).resolve(strict=True)
    target = Path(target).absolute()
    if target.exists():
        raise ValueError("Migration destination already exists; select a new workspace.")
    if target == package or target in package.parents or package in target.parents:
        raise ValueError("Migration destination must be separate from the legacy package.")
    if source_root is not None:
        source_root = Path(source_root).resolve(strict=True)
        if not source_root.is_dir():
            raise ValueError("Selected source root is not a folder.")
        if target == source_root or target in source_root.parents or source_root in target.parents:
            raise ValueError("New workspace must be separate from the selected source folder.")
    for parent in (target, *target.parents):
        if parent.exists() and (parent.is_symlink() or _is_reparse(parent)):
            raise ValueError("Migration destination traverses a symlink or reparse point.")
    report = preview_package(package)
    if not confirmed_fingerprint or confirmed_fingerprint != _fingerprint(report):
        raise ValueError("Migration preview changed or was not confirmed; preview again.")
    if report["blockers"]:
        raise ValueError("Migration is blocked: " + "; ".join(report["blockers"]))

    # Re-hash before copying and after copying. The preview fingerprint alone is not
    # a substitute for checking every inventoried source byte at both boundaries.
    def source_inventory():
        result={}
        for name in report["files"]:
            path=package/name
            unsafe=path.is_symlink() or _is_reparse(path)
            for parent in path.parents:
                if parent==package.parent:break
                if parent.is_symlink() or _is_reparse(parent):unsafe=True
            if unsafe or not path.resolve().is_relative_to(package):
                raise ValueError("Legacy inventory path became unsafe; import cancelled.")
            result[name]={"bytes":path.stat().st_size,"sha256":_sha256(path)}
        return result

    expected = report["files"]
    if source_inventory() != expected:
        raise ValueError("Legacy inventory changed after preview; preview again.")
    target.parent.mkdir(parents=True, exist_ok=True)
    scratch = Path(tempfile.mkdtemp(prefix=target.name + ".migration-", dir=target.parent))
    try:
        state = scratch / "state"
        runtime = package / "runtime"
        for name in ("approved", "staging", "cache", "logs", "user_flags"):
            src = runtime / name
            if src.exists():
                _copy_tree_checked(src, state / name)
        translation=report.get("legacy_draft_translation")
        if translation:
            stage=state/"staging"/translation["run"]
            baseline=_baseline(state,stage)
            payload={"run":translation["run"],"baseline":baseline,"version":1,"dirty":True,"choices":translation["choices"]}
            (stage/"web_draft.json").write_text(json.dumps(payload,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
        config = scratch / "config"
        shutil.copytree(package / "config", config)
        history = package / "config-history"
        if history.is_dir():shutil.copytree(history, scratch / "config-history")
        # Preserve migration inputs that are evidence but not active V2 state.
        evidence = scratch / "migration-evidence"
        for rel in sorted(expected):
            path = package / rel
            if rel.startswith("config/") or rel.startswith("config-history/") or rel.startswith("runtime/") and Path(rel).name not in {
                "owner_review_drafts.csv", "owner_review_drafts.csv.tmp", "owner_review_drafts.csv.bak",
                "user_custom_fields.csv", "user_keyword_mappings.csv"}:
                continue
            dest = evidence / rel
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(path, dest)
        # Store validated custom fields in the current application's canonical schema.
        fields = report.get("custom_fields", {"fields": [], "mappings": []})
        custom_path=config / "custom_fields.json"
        if not custom_path.exists():custom_path.write_text(json.dumps(fields, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        settings = {"data_dir": str(source_root) if source_root else report.get("source_path", ""),
                    "requires_rescan": True}
        (scratch / "settings.json").write_text(json.dumps(settings, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        manifest = {"version": 1, "preview_sha256": report["preview_sha256"],
                    "source_inventory": expected, "approved_run": report["approved_run"],
                    "approved_events": report["approved_events"], "imported_at": None}
        (scratch / "migration-manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        if source_inventory() != expected:
            raise ValueError("Legacy package changed during import; destination was not published.")
        if target.exists():
            raise ValueError("Migration destination appeared during import.")
        os.replace(scratch, target)
    except Exception:
        if scratch.exists():
            shutil.rmtree(scratch)
        raise
    return {"workspace": str(target), "preview_sha256": report["preview_sha256"],
            "approved_run": report["approved_run"], "approved_events": report["approved_events"],
            "source_files_verified": len(expected), "requires_rescan": True}


def _sha256(path):
    h = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _copy_tree_checked(source, target):
    source = Path(source)
    for path in source.rglob("*"):
        if path.is_symlink() or _is_reparse(path) or not path.resolve().is_relative_to(source.resolve()):
            raise ValueError("Unsafe symlink or path in legacy runtime: " + str(path))
    shutil.copytree(source, target, symlinks=False)

def _is_reparse(path):
    try:return bool(Path(path).lstat().st_file_attributes & 0x400)
    except (AttributeError,OSError):return False
