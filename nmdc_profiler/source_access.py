from __future__ import annotations

import errno
import shutil
import tempfile
import time
from pathlib import Path
from typing import Any, Callable, Dict, List, Tuple, TypeVar

from .core import get_source_family, infer_logical_register_identity, infer_path_projects, rel_posix

T = TypeVar("T")
_INSTALLED = False

_RETRY_ATTEMPTS = 12
_RETRY_DELAY_SECONDS = 0.35
_TRANSIENT_ERRNOS = {errno.EACCES, errno.EBUSY}
_TRANSIENT_WINERRORS = {5, 32, 33}


def _is_transient_access_error(exc: BaseException) -> bool:
    if isinstance(exc, PermissionError):
        return True
    if not isinstance(exc, OSError):
        return False
    if getattr(exc, "errno", None) in _TRANSIENT_ERRNOS:
        return True
    return getattr(exc, "winerror", None) in _TRANSIENT_WINERRORS


def _retry_access(operation: Callable[[], T], *, attempts: int = _RETRY_ATTEMPTS) -> T:
    """Retry transient Excel/OneDrive sharing violations before treating a source as unavailable."""
    last_error: BaseException | None = None
    for attempt in range(attempts):
        try:
            return operation()
        except OSError as exc:
            if not _is_transient_access_error(exc):
                raise
            last_error = exc
            if attempt + 1 >= attempts:
                raise
            time.sleep(_RETRY_DELAY_SECONDS * (attempt + 1))
    assert last_error is not None
    raise last_error


def _locked_profile(path: Path, root: Path, data_dir: Path, exc: OSError) -> Dict[str, object]:
    relative = rel_posix(path, root)
    family = get_source_family(path, data_dir)
    try:
        size = int(path.stat().st_size)
    except OSError:
        size = 0
    return {
        "relative_path": relative,
        "filename": path.name,
        "source_family": family,
        "file_extension": path.suffix,
        "file_size": size,
        "sha256": "",
        "readability_status": "UNREADABLE",
        "unreadable_reason": f"OPEN_FAILED:{exc.__class__.__name__}",
        "password_retry_possible": False,
        "inferred_project_numbers": infer_path_projects(path.name),
        "logical_register_identity": infer_logical_register_identity(path.name, family),
        "warnings": ["SOURCE_ACCESS_DENIED"],
        "sheets": [],
        "sheet_count": 0,
        "doc_created": "",
        "doc_modified": "",
        "timestamp_source": "",
        "timestamp_reliable": False,
        "candidate_internal_project_numbers": [],
        "project_mismatch_findings": [],
        "native_hyperlink_count": 0,
        "hyperlink_formula_count": 0,
        "merged_range_count": 0,
    }


def _profile_from_temporary_snapshot(original_profile, path: Path, root: Path, data_dir: Path):
    """Best-effort fallback for synchronized files that cannot be opened in-place but can be copied."""
    with tempfile.TemporaryDirectory(prefix="nmdc-source-") as temp_dir:
        snapshot = Path(temp_dir) / path.name
        _retry_access(lambda: shutil.copyfile(path, snapshot), attempts=4)
        profile = original_profile(snapshot, snapshot.parent, snapshot.parent)
        family = get_source_family(path, data_dir)
        profile["relative_path"] = rel_posix(path, root)
        profile["filename"] = path.name
        profile["source_family"] = family
        profile["inferred_project_numbers"] = infer_path_projects(path.name)
        profile["logical_register_identity"] = infer_logical_register_identity(path.name, family)
        warnings = [value for value in list(profile.get("warnings", []) or []) if value != "PROJECT_ID_INTERNAL_NOT_FOUND"]
        profile["warnings"] = warnings
        return profile


def install_resilient_source_access() -> None:
    """Install production-only resilience for transient Excel/OneDrive source locks.

    Valid source files are retried for roughly 20 seconds before being treated as
    inaccessible. If a direct open still fails, a temporary read-only snapshot is
    attempted. Persistent access failures are represented as a normal blocking
    SOURCE_HASH_ERROR during staging instead of crashing the packaged engine.
    """
    global _INSTALLED
    if _INSTALLED:
        return

    from . import runtime_engine, update_engine

    original_profile_workbook = runtime_engine.profile_workbook
    original_sha256_file = update_engine.sha256_file

    def resilient_profile_workbook(path: Path, root: Path, data_dir: Path):
        try:
            return _retry_access(lambda: original_profile_workbook(path, root, data_dir))
        except OSError as exc:
            if not _is_transient_access_error(exc):
                raise
            try:
                return _profile_from_temporary_snapshot(original_profile_workbook, path, root, data_dir)
            except OSError:
                return _locked_profile(path, root, data_dir, exc)

    def resilient_sha256_file(path: Path, chunk_size: int = 1024 * 1024) -> str:
        return _retry_access(lambda: original_sha256_file(path, chunk_size))

    def resilient_scan_sources(data_dir: Path) -> Tuple[List[Dict[str, Any]], List[Dict[str, str]]]:
        data_dir = Path(data_dir)
        entries: List[Dict[str, Any]] = []
        flags: List[Dict[str, str]] = []
        if not data_dir.exists() or not data_dir.is_dir():
            raise FileNotFoundError(f"Data folder not found: {data_dir}")

        files = sorted(
            (
                path
                for path in data_dir.rglob("*")
                if path.is_file()
                and path.suffix.lower() in {".xlsx", ".xlsm"}
                and not path.name.startswith("~$")
            ),
            key=lambda path: path.relative_to(data_dir).as_posix().casefold(),
        )
        for path in files:
            rel = path.relative_to(data_dir).as_posix()
            entry: Dict[str, Any] = {
                "relative_path": rel,
                "size": 0,
                "modified_ns": 0,
                "sha256": "",
                "scan_error": "",
            }
            try:
                stat = _retry_access(path.stat)
                entry["size"] = int(stat.st_size)
                entry["modified_ns"] = int(stat.st_mtime_ns)
                entry["sha256"] = resilient_sha256_file(path)
            except OSError as exc:
                entry["scan_error"] = f"{exc.__class__.__name__}: {exc}"
                flags.append(
                    update_engine._flag(
                        "CONFLICT",
                        "SOURCE_HASH_ERROR",
                        f"Source workbook '{rel}' is temporarily unavailable to the indexer.",
                        rel,
                        "Close this source workbook if it is open, wait for OneDrive sync to finish, then run the update again. The approved index was protected and this file was not treated as removed.",
                    )
                )
            entries.append(entry)
        return entries, flags

    runtime_engine.profile_workbook = resilient_profile_workbook
    update_engine.sha256_file = resilient_sha256_file
    update_engine.scan_sources = resilient_scan_sources
    _INSTALLED = True
