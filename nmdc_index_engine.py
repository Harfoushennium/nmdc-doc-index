from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Optional, Sequence

from nmdc_profiler.excel_bridge import create_support_request, export_excel_exchange
from nmdc_profiler.layout_compat import install_layout_compatibility
from nmdc_profiler.review_decisions import apply_review_decisions
from nmdc_profiler.runtime_admin import reset_runtime_state, undo_last_approval
from nmdc_profiler.runtime_engine import record_user_flag, stage_runtime_update
from nmdc_profiler.source_access import install_resilient_source_access
from nmdc_profiler.source_selection import set_source_selection, set_source_selections_from_file
from nmdc_profiler.source_selection_view import export_source_selection
from nmdc_profiler.ui_layout import install_review_first_column_order
from nmdc_profiler.update_engine import DEFAULT_PARSER_VERSION, approve_stage, hold_stage, reject_stage


def _common(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--state-dir", type=Path, required=True)
    parser.add_argument("--exchange-dir", type=Path, required=True)
    parser.add_argument("--config-dir", type=Path, default=Path("config"))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="NMDC Document Index packaged engine")
    subparsers = parser.add_subparsers(dest="command", required=True)

    stage = subparsers.add_parser("stage", help="Create a staged update without changing approved data")
    _common(stage)
    stage.add_argument("--mode", choices=("incremental", "full"), default="incremental")
    stage.add_argument("--data-dir", type=Path, required=True)
    stage.add_argument("--parser-version", default=DEFAULT_PARSER_VERSION)

    export = subparsers.add_parser("export-excel", help="Export approved and staged views for Excel")
    _common(export)

    approve = subparsers.add_parser("approve", help="Approve one exact staged run")
    _common(approve)
    approve.add_argument("--run-id", required=True)
    approve.add_argument("--allow-conflicts", action="store_true")

    for name in ("hold", "reject"):
        decision = subparsers.add_parser(name, help=f"{name.title()} one exact staged run")
        _common(decision)
        decision.add_argument("--run-id", required=True)
        decision.add_argument("--note", default="")

    reset = subparsers.add_parser("reset", help="Delete indexed/staged runtime records without touching source DATA or configuration")
    _common(reset)

    undo = subparsers.add_parser("undo", help="Restore the immediately preceding approved version")
    _common(undo)

    review = subparsers.add_parser("save-review-decisions", help="Persist Review Flags user decisions for the latest staged run")
    _common(review)
    review.add_argument("--decisions-file", type=Path, required=True)

    source_selection = subparsers.add_parser(
        "set-source-selection",
        help="Include or exclude one source workbook from owner-approved source selection",
    )
    _common(source_selection)
    source_selection.add_argument("--source-file", required=True)
    source_selection.add_argument("--action", choices=("INCLUDE", "EXCLUDE"), required=True)
    source_selection.add_argument("--reason", default="")

    source_selections = subparsers.add_parser(
        "save-source-selections",
        help="Persist all Excel source-selection checkboxes in one operation",
    )
    _common(source_selections)
    source_selections.add_argument("--selections-file", type=Path, required=True)

    context_fields = (
        "message",
        "project-no",
        "document-no",
        "source-file",
        "worksheet",
        "source-row",
        "source-cell",
        "revision",
        "event-identity",
        "flag-code",
        "current-field",
        "current-value",
        "expected-value",
        "user-name",
    )
    for name in ("user-flag", "support-request"):
        request = subparsers.add_parser(name, help="Create an auditable user support record")
        _common(request)
        for field in context_fields:
            request.add_argument(f"--{field}", required=field == "message", default="")

    return parser


def _request_context(args: argparse.Namespace) -> dict[str, str]:
    return {
        "message": args.message,
        "project_no": args.project_no,
        "document_no": args.document_no,
        "source_file": args.source_file,
        "worksheet": args.worksheet,
        "source_row": args.source_row,
        "source_cell": args.source_cell,
        "revision": args.revision,
        "event_identity": args.event_identity,
        "flag_code": args.flag_code,
        "current_field": args.current_field,
        "current_value": args.current_value,
        "expected_value": args.expected_value,
        "user_name": args.user_name,
    }


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        # Column order is presentation-only. Canonical extraction records remain
        # untouched; every Excel export is written in review-first order.
        install_review_first_column_order()

        if args.command == "stage":
            install_resilient_source_access()
            install_layout_compatibility()
            result = stage_runtime_update(
                data_dir=args.data_dir,
                state_dir=args.state_dir,
                config_dir=args.config_dir,
                full_rescan=args.mode == "full",
                parser_version=args.parser_version,
            )
        elif args.command == "export-excel":
            result = export_excel_exchange(args.state_dir, args.exchange_dir)
            export_source_selection(args.state_dir, args.exchange_dir, args.config_dir)
        elif args.command == "approve":
            result = approve_stage(args.state_dir, args.run_id, allow_conflicts=args.allow_conflicts)
        elif args.command == "hold":
            result = hold_stage(args.state_dir, args.run_id, note=args.note)
        elif args.command == "reject":
            result = reject_stage(args.state_dir, args.run_id, note=args.note)
        elif args.command == "reset":
            result = reset_runtime_state(args.state_dir)
        elif args.command == "undo":
            result = undo_last_approval(args.state_dir)
        elif args.command == "save-review-decisions":
            result = apply_review_decisions(args.state_dir, args.decisions_file)
        elif args.command == "set-source-selection":
            result = set_source_selection(
                args.config_dir,
                args.source_file,
                include=args.action == "INCLUDE",
                reason=args.reason,
            )
        elif args.command == "save-source-selections":
            result = set_source_selections_from_file(args.config_dir, args.selections_file)
        elif args.command == "user-flag":
            result = record_user_flag(args.state_dir, **_request_context(args))
        elif args.command == "support-request":
            context = _request_context(args)
            context.pop("flag_code", None)
            result = {"support_file": str(create_support_request(args.state_dir, **context))}
        else:
            raise ValueError(f"Unsupported command: {args.command}")
        print(json.dumps(result, ensure_ascii=False, sort_keys=True, default=str))
        return 0
    except (FileNotFoundError, ValueError, OSError) as exc:
        print(f"NMDC engine error: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        print(f"NMDC engine unexpected error: {exc.__class__.__name__}: {exc}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
