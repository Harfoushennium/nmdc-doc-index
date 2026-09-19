from __future__ import annotations

import json
import os
import shutil
import time
import uuid
from pathlib import Path
from typing import Any, Dict, Tuple


_CACHE_INDEX = "index.json"
_FILES_DIR = "files"
_COPY_ATTEMPTS = 8
_COPY_DELAY_SECONDS = 0.25


def _read_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    try:
        with path.open(encoding="utf-8") as handle:
            return json.load(handle)
    except (OSError, ValueError, json.JSONDecodeError):
        return default


def _write_json_atomic(path: Path, payload: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(path.name + f".tmp-{uuid.uuid4().hex[:8]}")
    with temp.open("w", encoding="utf-8", newline="\n") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2, sort_keys=True)
        handle.write("\n")
    os.replace(temp, path)


def _copy_with_retries(source: Path, target: Path) -> None:
    """Copy one changed source to the local cache with short sharing-lock retries."""
    last_error: OSError | None = None
    target.parent.mkdir(parents=True, exist_ok=True)
    temp = target.with_name(target.name + f".tmp-{uuid.uuid4().hex[:8]}")
    for attempt in range(_COPY_ATTEMPTS):
        try:
            if temp.exists():
                temp.unlink()
            shutil.copy2(source, temp)
            os.replace(temp, target)
            return
        except OSError as exc:
            last_error = exc
            try:
                if temp.exists():
                    temp.unlink()
            except OSError:
                pass
            if attempt + 1 < _COPY_ATTEMPTS:
                time.sleep(_COPY_DELAY_SECONDS * (attempt + 1))
    detail = f"{last_error.__class__.__name__}: {last_error}" if last_error else "unknown source access error"
    raise PermissionError(
        f"Source workbook could not be copied to the local scan cache: {source}. "
        "Close the source workbook if it is open and allow OneDrive to finish any active sync, then try again. "
        f"Technical detail: {detail}"
    ) from last_error


def prepare_local_source_cache(data_dir: Path, cache_root: Path) -> Tuple[Path, Dict[str, int]]:
    """Build/reuse a local mirror of the source tree for production scanning.

    The original DATA folder is only opened when a workbook is new, changed or the
    corresponding cached copy is missing. Unchanged workbooks are processed from
    the local cache, which avoids repeatedly waking OneDrive during every scan.
    Relative paths are preserved exactly so extraction evidence and Excel source
    hyperlinks continue to refer to the original DATA tree.
    """
    data_dir = Path(data_dir).resolve()
    cache_root = Path(cache_root).resolve()
    files_root = cache_root / _FILES_DIR
    index_path = cache_root / _CACHE_INDEX

    if not data_dir.exists() or not data_dir.is_dir():
        raise FileNotFoundError(f"Data folder not found: {data_dir}")

    cache_root.mkdir(parents=True, exist_ok=True)
    files_root.mkdir(parents=True, exist_ok=True)
    old_index = _read_json(index_path, {})
    if not isinstance(old_index, dict):
        old_index = {}

    new_index: Dict[str, Dict[str, int]] = {}
    source_files = sorted(
        (
            path
            for path in data_dir.rglob("*")
            if path.is_file()
            and path.suffix.lower() in {".xlsx", ".xlsm"}
            and not path.name.startswith("~$")
        ),
        key=lambda path: path.relative_to(data_dir).as_posix().casefold(),
    )

    reused = 0
    copied = 0
    for source in source_files:
        rel = source.relative_to(data_dir).as_posix()
        stat = source.stat()
        size = int(stat.st_size)
        modified_ns = int(stat.st_mtime_ns)
        cached = files_root / Path(rel)
        previous = old_index.get(rel, {}) if isinstance(old_index.get(rel, {}), dict) else {}
        reusable = (
            cached.exists()
            and int(previous.get("size", -1)) == size
            and int(previous.get("modified_ns", -1)) == modified_ns
        )
        if reusable:
            reused += 1
        else:
            _copy_with_retries(source, cached)
            copied += 1
        new_index[rel] = {"size": size, "modified_ns": modified_ns}

    # Remove cached files for sources that no longer exist. The original source
    # tree remains untouched; only the local cache is cleaned.
    stale = set(old_index) - set(new_index)
    for rel in stale:
        cached = files_root / Path(rel)
        try:
            if cached.exists():
                cached.unlink()
        except OSError:
            pass

    _write_json_atomic(index_path, new_index)
    return files_root, {
        "source_files": len(source_files),
        "cache_reused": reused,
        "cache_copied": copied,
        "cache_removed": len(stale),
    }
