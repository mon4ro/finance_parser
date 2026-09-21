"""
DANGEROUS TOOL - removes real, already-imported transaction rows.

Removes every row belonging to one already-imported source file (matched
by exact SourceFile) from the raw, per-bank/broker raw-export, and
classified sheets of a budgeting or investment workbook. Built after a
real incident: a Nordea export silently changed date format
(YYYY/MM/DD instead of YYYY-MM-DD), which parse_date() mis-parsed
(day/month swapped), corrupting 108 rows including 42 real duplicate
transactions that evaded RawID-based dedup because the wrong date changed
the computed hash. Cleaning that up by hand (cross-referencing RawID,
ImportLog, and Review/Notes manually) is exactly the kind of error-prone
work this tool exists to make safe and repeatable.

This tool ONLY removes rows. It never re-imports - that stays a separate,
deliberate step (run the normal parser again against the same file once
the underlying bug is fixed).

ImportLog/InvestmentImportLog is deliberately left untouched - it is a
historical record of "the parser ran against this file and produced X",
which remains true even after a later cleanup. The removal itself is
recorded as a ChangeLog entry instead (script, source file, rows removed
per sheet, backup path) - ChangeLog is this codebase's existing
general-purpose "what happened and why" audit trail.

Safety design (two layers, plus a backup):
1. Report-first, hard-blocking by default: prints exactly what would be
   removed before anything is touched. Refuses to proceed (even with
   --apply) if any affected classified row has non-blank manual notes
   (Review/Notes on budgeting, Comments on investments) or is part of a
   manual split (UnifiedID has a -S## suffix, or is itself a split
   parent with existing children) - unless explicitly overridden with
   --include-reviewed / --include-split. Those represent real human work
   or structural relationships that blind deletion could corrupt.
2. Typed confirmation: --apply also requires --confirm <exact source
   file name> (or, when run interactively with no --confirm given, a
   prompt asking you to type it) - a bare --apply flag is never enough
   on its own.
3. Backup before any write, via the same proven-safe fresh-rebuild path
   used throughout finance_parser (replace_with_fresh_workbook,
   excel_tables=False - see project_excel_corruption_root_cause memory).

Run via the thin root wrapper:
  python tools/remove_import.py --pipeline budgeting --source-file NAME.csv
      (report only - always safe, never writes)
  python tools/remove_import.py --pipeline budgeting --source-file NAME.csv \
      --apply --confirm NAME.csv
      (real removal, after a backup)
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from finance_parser.budgeting.parsers import cash, investment_dividends, nordea as b_nordea, norwegian, op, spankki
from finance_parser.budgeting.transaction_parser import DEFAULT_OUTPUT_WORKBOOK as DEFAULT_BUDGETING_WORKBOOK
from finance_parser.common import IMPORT_LOG_SHEET, RAW_SHEET, UNIFIED_SHEET
from finance_parser.investments.investment_common import (
    INVESTMENT_IMPORT_LOG_SHEET,
    INVESTMENT_RAW_SHEET,
    INVESTMENT_TRANSACTIONS_SHEET,
)
from finance_parser.investments.investment_parser import DEFAULT_OUTPUT_WORKBOOK as DEFAULT_INVESTMENT_WORKBOOK
from finance_parser.investments.parsers import coinmotion, evli, nordea as i_nordea, nordnet, op_investment, seligson
from finance_parser.utilities.fresh_workbook_writer import (
    append_changelog_row,
    read_workbook_values_only,
    records_to_sheet_values,
    replace_with_fresh_workbook,
    sheet_values_to_records,
)


SCRIPT_NAME = "tools/remove_import.py"

_BUDGETING_PARSERS = [op, norwegian, b_nordea, spankki, cash, investment_dividends]
_INVESTMENT_PARSERS = [nordnet, seligson, evli, op_investment, coinmotion, i_nordea]


@dataclass
class PipelineProfile:
    name: str
    default_workbook: Path
    raw_sheet: str
    classified_sheet: str
    import_log_sheet: str
    raw_id_column: str
    source_column: str  # SourceBank (budgeting) / Broker (investments)
    review_gate_columns: list[str]
    split_check: bool  # only budgeting has the -S## split convention
    raw_export_sheet_by_source: dict[str, str]


def _budgeting_profile() -> PipelineProfile:
    mapping = {mod.SOURCE_BANK: mod.BANK_RAW_SHEET for mod in _BUDGETING_PARSERS}
    return PipelineProfile(
        name="budgeting",
        default_workbook=DEFAULT_BUDGETING_WORKBOOK,
        raw_sheet=RAW_SHEET,
        classified_sheet=UNIFIED_SHEET,
        import_log_sheet=IMPORT_LOG_SHEET,
        raw_id_column="RawID",
        source_column="SourceBank",
        review_gate_columns=["Review/Notes"],
        split_check=True,
        raw_export_sheet_by_source=mapping,
    )


def _investment_profile() -> PipelineProfile:
    mapping = {mod.BROKER: mod.BROKER_RAW_SHEET for mod in _INVESTMENT_PARSERS}
    return PipelineProfile(
        name="investments",
        default_workbook=DEFAULT_INVESTMENT_WORKBOOK,
        raw_sheet=INVESTMENT_RAW_SHEET,
        classified_sheet=INVESTMENT_TRANSACTIONS_SHEET,
        import_log_sheet=INVESTMENT_IMPORT_LOG_SHEET,
        raw_id_column="InvestmentRawID",
        source_column="Broker",
        review_gate_columns=["Comments"],
        split_check=False,
        raw_export_sheet_by_source=mapping,
    )


PROFILES = {
    "budgeting": _budgeting_profile,
    "investments": _investment_profile,
}


_SPLIT_CHILD_SUFFIX_RE = re.compile(r"-S\d{2}$")


def _is_blank(value: object) -> bool:
    if value is None:
        return True
    if pd.isna(value):
        return True
    return str(value).strip() == ""


@dataclass
class Plan:
    profile: PipelineProfile
    source_file: str
    headers: dict[str, list[str]]
    raw_records: list[dict[str, object]]
    raw_matched_idx: list[int]
    classified_records: list[dict[str, object]]
    classified_matched_idx: list[int]
    raw_export_sheet: str | None
    raw_export_records: list[dict[str, object]] | None
    raw_export_matched_idx: list[int]
    blocking_reviewed: list[dict[str, object]]
    blocking_split: list[dict[str, object]]


def build_plan(sheets: dict[str, list[list[object]]], profile: PipelineProfile, source_file: str) -> Plan:
    headers: dict[str, list[str]] = {}

    raw_headers, raw_records = sheet_values_to_records(sheets.get(profile.raw_sheet, []))
    headers[profile.raw_sheet] = raw_headers
    raw_matched_idx = [i for i, r in enumerate(raw_records) if str(r.get("SourceFile", "")) == source_file]
    matched_raw_ids = {str(raw_records[i].get(profile.raw_id_column, "")) for i in raw_matched_idx}

    classified_headers, classified_records = sheet_values_to_records(sheets.get(profile.classified_sheet, []))
    headers[profile.classified_sheet] = classified_headers
    classified_matched_idx = [
        i for i, r in enumerate(classified_records)
        if str(r.get("SourceFile", "")) == source_file
        or str(r.get(profile.raw_id_column, "")) in matched_raw_ids
    ]

    source_values = {str(raw_records[i].get(profile.source_column, "")) for i in raw_matched_idx}
    source_values.discard("")
    raw_export_sheet = None
    raw_export_records: list[dict[str, object]] | None = None
    raw_export_matched_idx: list[int] = []
    if len(source_values) == 1:
        source_value = next(iter(source_values))
        candidate_sheet = profile.raw_export_sheet_by_source.get(source_value)
        if candidate_sheet and candidate_sheet in sheets:
            re_headers, re_records = sheet_values_to_records(sheets[candidate_sheet])
            headers[candidate_sheet] = re_headers
            re_matched_idx = [i for i, r in enumerate(re_records) if str(r.get("SourceFile", "")) == source_file]
            raw_export_sheet = candidate_sheet
            raw_export_records = re_records
            raw_export_matched_idx = re_matched_idx

    blocking_reviewed = []
    blocking_split = []
    for i in classified_matched_idx:
        record = classified_records[i]
        for col in profile.review_gate_columns:
            if col in record and not _is_blank(record.get(col)):
                blocking_reviewed.append(record)
                break

    if profile.split_check:
        all_unified_ids = {str(r.get("UnifiedID", "")) for r in classified_records}
        for i in classified_matched_idx:
            record = classified_records[i]
            uid = str(record.get("UnifiedID", "")).strip()
            if not uid:
                continue
            is_split_child = bool(_SPLIT_CHILD_SUFFIX_RE.search(uid))
            has_children = any(
                other_uid.startswith(f"{uid}-S") and other_uid[len(uid) + 2:].isdigit()
                for other_uid in all_unified_ids
                if other_uid != uid
            )
            if is_split_child or has_children:
                blocking_split.append(record)

    return Plan(
        profile=profile,
        source_file=source_file,
        headers=headers,
        raw_records=raw_records,
        raw_matched_idx=raw_matched_idx,
        classified_records=classified_records,
        classified_matched_idx=classified_matched_idx,
        raw_export_sheet=raw_export_sheet,
        raw_export_records=raw_export_records,
        raw_export_matched_idx=raw_export_matched_idx,
        blocking_reviewed=blocking_reviewed,
        blocking_split=blocking_split,
    )


def print_plan(plan: Plan) -> None:
    p = plan.profile
    print(f"Pipeline: {p.name}")
    print(f"Source file: {plan.source_file}")
    print()
    print(f"{p.raw_sheet}: {len(plan.raw_matched_idx)} row(s) would be removed")
    print(f"{p.classified_sheet}: {len(plan.classified_matched_idx)} row(s) would be removed")
    if plan.raw_export_sheet:
        print(f"{plan.raw_export_sheet}: {len(plan.raw_export_matched_idx)} row(s) would be removed")
    else:
        print("(no matching per-bank/broker raw-export sheet found or SourceBank/Broker is ambiguous - "
              "that sheet will be left untouched)")
    print(f"{p.import_log_sheet}: left untouched (historical record) - see ChangeLog for the removal record instead")
    print()

    preview_n = 20
    for idx in plan.classified_matched_idx[:preview_n]:
        record = plan.classified_records[idx]
        rid_col = p.raw_id_column
        print(f"  - {record.get(rid_col)}  Date/TradeDate={record.get('Date') or record.get('TradeDate')}  "
              f"Amount={record.get('Amount') or record.get('CashAmount')}")
    if len(plan.classified_matched_idx) > preview_n:
        print(f"  ... and {len(plan.classified_matched_idx) - preview_n} more")
    print()

    if plan.blocking_reviewed:
        print(f"BLOCKED: {len(plan.blocking_reviewed)} row(s) have manual notes set "
              f"({'/'.join(p.review_gate_columns)}) - pass --include-reviewed to override.")
    if plan.blocking_split:
        print(f"BLOCKED: {len(plan.blocking_split)} row(s) are part of a manual split "
              f"(a -S## child or a parent with existing children) - pass --include-split to override.")
    if not plan.blocking_reviewed and not plan.blocking_split:
        print("No blocking issues found.")


def apply_plan(sheets: dict[str, list[list[object]]], plan: Plan) -> dict[str, tuple[int, int]]:
    """Returns {sheet_name: (rows_before, rows_after)} for every sheet touched."""
    stats: dict[str, tuple[int, int]] = {}

    def remove(sheet_name: str, headers: list[str], records: list[dict[str, object]], matched_idx: list[int]) -> None:
        before = len(records)
        matched = set(matched_idx)
        kept = [r for i, r in enumerate(records) if i not in matched]
        sheets[sheet_name] = records_to_sheet_values(headers, kept)
        stats[sheet_name] = (before, len(kept))

    remove(plan.profile.raw_sheet, plan.headers[plan.profile.raw_sheet], plan.raw_records, plan.raw_matched_idx)
    remove(plan.profile.classified_sheet, plan.headers[plan.profile.classified_sheet], plan.classified_records, plan.classified_matched_idx)
    if plan.raw_export_sheet:
        remove(plan.raw_export_sheet, plan.headers[plan.raw_export_sheet], plan.raw_export_records, plan.raw_export_matched_idx)

    return stats


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--pipeline", required=True, choices=sorted(PROFILES.keys()))
    parser.add_argument("--source-file", required=True, help="Exact SourceFile value to remove, e.g. NAME.csv")
    parser.add_argument("--workbook", type=Path, default=None, help="Defaults to the pipeline's standard output workbook.")
    parser.add_argument("--apply", action="store_true", help="Actually remove the rows. Without this, report only.")
    parser.add_argument("--confirm", default=None, help="Must exactly match --source-file when used with --apply.")
    parser.add_argument("--include-reviewed", action="store_true", help="Allow removing rows with manual notes set.")
    parser.add_argument("--include-split", action="store_true", help="Allow removing rows that are part of a manual split.")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()
    profile = PROFILES[args.pipeline]()
    workbook = args.workbook or profile.default_workbook

    sheets = read_workbook_values_only(workbook)
    plan = build_plan(sheets, profile, args.source_file)
    print_plan(plan)

    if not plan.raw_matched_idx and not plan.classified_matched_idx:
        print()
        print(f"No rows found for source file {args.source_file!r} - nothing to do.")
        return

    blocked = (plan.blocking_reviewed and not args.include_reviewed) or (plan.blocking_split and not args.include_split)
    if blocked:
        print()
        print("Refusing to proceed while blocking issues remain (see above).")
        return

    if not args.apply:
        print()
        print("Dry-run only (no --apply given). Real workbook was not touched.")
        return

    confirm = args.confirm
    if confirm is None:
        if not sys.stdin.isatty():
            print()
            print("Refusing to --apply without --confirm in a non-interactive context.")
            return
        confirm = input(f"Type the exact source file name to confirm deletion ({args.source_file}): ").strip()

    if confirm != args.source_file:
        print()
        print("Confirmation text did not match the source file name exactly - aborting. Nothing was touched.")
        return

    stats = apply_plan(sheets, plan)

    total_removed = sum(before - after for before, after in stats.values())
    details_parts = [f"{sheet}: {before}->{after} (removed {before - after})" for sheet, (before, after) in stats.items()]
    details = (
        f"Removed source file {args.source_file!r} from pipeline {profile.name!r}. "
        f"{'; '.join(details_parts)}. "
        f"{profile.import_log_sheet} left untouched (historical record)."
    )

    append_changelog_row(
        sheets,
        script=SCRIPT_NAME,
        action=f"Remove import: {args.source_file}",
        sheet=", ".join(stats.keys()),
        rows_updated=total_removed,
        rows_removed=total_removed,
        status="Completed",
        details=details,
        backup_file=None,
    )

    backup_path, _, write_stats = replace_with_fresh_workbook(
        workbook,
        sheets,
        backup_label="before_remove_import",
        excel_tables=False,
    )
    print(f"Backup written: {backup_path}")
    print(f"Removed {total_removed} row(s) across {len(stats)} sheet(s). {write_stats['sheets_written']} sheet(s) written.")
    print("ChangeLog entry appended.")


if __name__ == "__main__":
    main()
