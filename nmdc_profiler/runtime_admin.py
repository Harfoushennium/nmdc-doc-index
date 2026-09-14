from __future__ import annotations

import json
import os
import shutil
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List


def _utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _read_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def _write_json_atomic(path: Path, payload: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(path.name + f".tmp-{uuid.uuid4().hex[:8]}")
    with temp.open("w", encoding="utf-8", newline="\n") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2, sort_keys=True)
        handle.write("\n")
    os.replace(temp, path)


def _append_history(state_dir: Path, payload: Dict[str, Any]) -> None:
    path = Path(state_dir) / "logs" / "history.jsonl"
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8", newline="\n") as handle:
        handle.write(json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str) + "\n")


def reset_runtime_state(state_dir: Path) -> Dict[str, Any]:
    """Delete indexed/staged runtime records while preserving source DATA, config and audit history."""
    state_dir = Path(state_dir)
    removed: List[str] = []
    for name in ("approved", "staging", "cache", "user_flags", "profile"):
        target = state_dir / name
        if target.exists():
            if target.is_dir():
                shutil.rmtree(target)
            else:
                target.unlink()
            removed.append(name)
    result = {
        "decision": "RESET",
        "reset_at": _utc_now(),
        "removed_runtime_areas": removed,
        "source_data_changed": False,
        "configuration_changed": False,
    }
    _append_history(state_dir, {"event": "RESET", **result})
    return result


def undo_last_approval(state_dir: Path) -> Dict[str, Any]:
    """Move the approved pointer to the immediately preceding approved version without deleting versions."""
    state_dir = Path(state_dir)
    approved_dir = state_dir / "approved"
    current = _read_json(approved_dir / "current.json", {})
    current_run = str(current.get("run_id", ""))
    if not current_run:
        raise ValueError("There is no approved update to undo.")

    versions_dir = approved_dir / "versions"
    versions: List[Dict[str, str]] = []
    if versions_dir.exists():
        for child in versions_dir.iterdir():
            if not child.is_dir():
                continue
            approval = _read_json(child / "approval.json", {})
            run_id = str(approval.get("run_id", child.name))
            approved_at = str(approval.get("approved_at", ""))
            if run_id:
                versions.append({"run_id": run_id, "approved_at": approved_at})
    versions.sort(key=lambda row: (row.get("approved_at", ""), row.get("run_id", "")))

    current_index = next((i for i, row in enumerate(versions) if row["run_id"] == current_run), -1)
    if current_index < 0:
        raise ValueError(f"The current approved version is not available for undo: {current_run}")
    if current_index == 0:
        raise ValueError("There is no previous approved version available to restore.")

    previous_run = versions[current_index - 1]["run_id"]
    _write_json_atomic(
        approved_dir / "current.json",
        {"run_id": previous_run, "version_path": f"versions/{previous_run}"},
    )

    result = {
        "decision": "UNDO_APPROVAL",
        "undone_run_id": current_run,
        "restored_run_id": previous_run,
        "undone_at": _utc_now(),
    }
    _append_history(state_dir, {"event": "UNDO_APPROVAL", **result})
    return result
