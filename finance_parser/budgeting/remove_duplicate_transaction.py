"""
Removes a mechanically-confirmed duplicate transaction (same real event,
two different RawIDs) from RawTransactions and UnifiedTransactions.

This does NOT try to detect duplicates itself - that stays
find_duplicate_transactions.py's job (a read-only report), with a human
deciding what's actually confirmed. This tool only acts once you hand it a
set of RawIDs you've already confirmed refer to the same real transaction,
and it re-verifies that claim itself before touching anything: every
RawTransactions field except a small set of identity/import-metadata
columns (RawID, UnifiedID, ImportedAt, ExportDate, SourceFile - all of
which are EXPECTED to differ between two imports of the same real
transaction) must match exactly across all given RawIDs, or it refuses.

Real incident this was built for: the same real transaction on a real
household member's account was imported twice, five weeks apart, because
a filename-based SourceAccount inference bug (a stray "TILI " filename
prefix leaking into the inferred account name) changed between the two
imports - the same real transaction hashed to two different RawIDs. A
later fix corrected the *displayed* SourceAccount on the older row via a
data migration, but correctly left its RawID untouched (recomputing an
existing RawID would just move the problem), so the two rows still didn't
dedup against each other.

Safety design (mirrors tools/remove_import.py):
1. Report-first: prints the full comparison (every field, side by side)
   before anything is touched. Refuses outright if any field the
   verification cares about differs - this is not a confidence heuristic,
   it is a hard block.
2. Refuses to remove a row that has non-blank Review/Notes on its
   UnifiedTransactions row (real manual work), unless --include-reviewed.
3. Picks which row to keep by recomputing that bank's *current*
   make_raw_id() against the (verified-identical) raw fields: if exactly
   one candidate's RawID matches what current code would produce right
   now, that one is kept - it's the one any future re-import of the same
   file will naturally match, so keeping the other one would just recreate
   this exact duplicate the next time the source file gets reprocessed
   (old input files are never archived out of input/budgeting/, so this
   is a real, not hypothetical, risk). Falls back to keeping the earliest
   ImportedAt only when the bank's parser can't be resolved or recomputing
   doesn't match either candidate - a real incident found "keep oldest"
   alone silently picks the *wrong* row exactly when the duplicate was
   caused by a parsing-logic fix landing between the two imports (which is
   the common case for this kind of duplicate), since the older row is
   usually the one hashed under the old, now-dead logic.
4. Backup before any write, via the same fresh-rebuild path used
   throughout finance_parser (replace_with_fresh_workbook,
   excel_tables=False).
5. Typed confirmation before a real --apply write: --confirm must exactly
   match the comma-joined, sorted list of RawIDs that would be removed.

Deliberately out of scope: per-bank raw-export sheets (e.g. OPRawExport).
Those sheets don't carry RawID and matching a single row there back to a
specific RawTransactions row would need bank-specific column matching:
generally not needed anyway, since in the incident this was built for the
raw-export sheet's own separate dedup key already correctly avoided
creating a second row there. If a raw-export duplicate is ever found, it
needs separate handling.

Run via the thin root wrapper:
  python tools/remove_duplicate_transaction.py --raw-ids ID1,ID2
      (report only - always safe, never writes)
  python tools/remove_duplicate_transaction.py --raw-ids ID1,ID2 \
      --apply --confirm ID2
      (real removal, after a backup - --confirm must equal the sorted,
      comma-joined RawIDs that would actually be removed)
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from finance_parser.budgeting.parsers import cash, nordea, norwegian, op, spankki
from finance_parser.common import RAW_SHEET, UNIFIED_SHEET
from finance_parser.utilities.fresh_workbook_writer import (
    append_changelog_row,
    read_workbook_values_only,
    records_to_sheet_values,
    replace_with_fresh_workbook,
    sheet_values_to_records,
)


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_BUDGETING_WORKBOOK = PROJECT_ROOT / "output" / "budgeting" / "ParsedTransactions.xlsx"

SCRIPT_NAME = "tools/remove_duplicate_transaction.py"

# Expected to legitimately differ between two imports of the same real
# transaction - excluded from the identity comparison.
IDENTITY_EXCLUDED_COLUMNS = {"RawID", "UnifiedID", "ImportedAt", "ExportDate", "SourceFile"}

# cash.py's make_raw_id() handles both SourceBank values (CASH and VINTED)
# itself, keyed off the row's own SourceBank field.
_MAKE_RAW_ID_BY_SOURCE_BANK = {
    op.SOURCE_BANK: op.make_raw_id,
    nordea.SOURCE_BANK: nordea.make_raw_id,
    spankki.SOURCE_BANK: spankki.make_raw_id,
    norwegian.SOURCE_BANK: norwegian.make_raw_id,
    cash.SOURCE_BANK: cash.make_raw_id,
    cash.VINTED_SOURCE_BANK: cash.make_raw_id,
}


def recompute_current_raw_id(record: dict[str, object]) -> str | None:
    """
    Recompute what THIS bank's current make_raw_id() produces for a raw
    record's own stored fields. Returns None if the SourceBank can't be
    resolved to a known parser, or recomputation itself raises (e.g. a
    genuinely malformed row) - callers should treat None as "can't tell".
    """
    source_bank = str(record.get("SourceBank", ""))
    make_raw_id = _MAKE_RAW_ID_BY_SOURCE_BANK.get(source_bank)
    if make_raw_id is None:
        return None
    try:
        return make_raw_id(pd.Series(record))
    except Exception:
        return None


def _is_blank(value: object) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip() == ""


@dataclass
class Plan:
    raw_ids: list[str]
    raw_headers: list[str]
    raw_records: list[dict[str, object]]
    raw_idx_by_id: dict[str, int]
    unified_headers: list[str]
    unified_records: list[dict[str, object]]
    unified_idx_by_id: list[int]  # indices into unified_records matching any of raw_ids
    keep_raw_id: str | None = None
    remove_raw_ids: list[str] = field(default_factory=list)
    mismatches: list[str] = field(default_factory=list)
    not_found: list[str] = field(default_factory=list)
    blocking_reviewed: list[dict[str, object]] = field(default_factory=list)
    keep_reason: str = ""


def verify_same_transaction(records: list[dict[str, object]]) -> list[str]:
    """
    Returns the list of field names that differ across the given
    RawTransactions records (excluding IDENTITY_EXCLUDED_COLUMNS). Empty
    means every other field matches exactly - the records are verified to
    represent the same real transaction, not just a heuristic guess.
    """
    if len(records) < 2:
        return []
    all_fields: set[str] = set()
    for r in records:
        all_fields.update(r.keys())
    fields = sorted(all_fields - IDENTITY_EXCLUDED_COLUMNS)

    mismatches = []
    base = records[0]
    for f in fields:
        base_value = str(base.get(f, "")).strip()
        for other in records[1:]:
            other_value = str(other.get(f, "")).strip()
            if base_value != other_value:
                mismatches.append(f)
                break
    return mismatches


def build_plan(sheets: dict[str, list[list[object]]], raw_ids: list[str]) -> Plan:
    raw_headers, raw_records = sheet_values_to_records(sheets.get(RAW_SHEET, []))
    raw_idx_by_id = {str(r.get("RawID", "")): i for i, r in enumerate(raw_records)}

    unified_headers, unified_records = sheet_values_to_records(sheets.get(UNIFIED_SHEET, []))
    unified_idx_by_id = [i for i, r in enumerate(unified_records) if str(r.get("RawID", "")) in raw_ids]

    not_found = [rid for rid in raw_ids if rid not in raw_idx_by_id]
    found_raw_ids = [rid for rid in raw_ids if rid in raw_idx_by_id]

    plan = Plan(
        raw_ids=raw_ids,
        raw_headers=raw_headers,
        raw_records=raw_records,
        raw_idx_by_id=raw_idx_by_id,
        unified_headers=unified_headers,
        unified_records=unified_records,
        unified_idx_by_id=unified_idx_by_id,
        not_found=not_found,
    )
    if not_found or len(found_raw_ids) < 2:
        return plan

    candidate_records = [raw_records[raw_idx_by_id[rid]] for rid in found_raw_ids]
    plan.mismatches = verify_same_transaction(candidate_records)
    if plan.mismatches:
        return plan

    def imported_at(rid: str) -> str:
        return str(raw_records[raw_idx_by_id[rid]].get("ImportedAt", ""))

    ordered = sorted(found_raw_ids, key=imported_at)

    # All candidates are verified identical except the excluded identity
    # columns, so recomputing from any one of them is equivalent - use the
    # first available.
    recomputed = recompute_current_raw_id(candidate_records[0])
    matches = [rid for rid in found_raw_ids if rid == recomputed] if recomputed else []

    if len(matches) == 1:
        plan.keep_raw_id = matches[0]
        plan.keep_reason = (
            f"matches what current code recomputes for this row ({recomputed}) - "
            "the one any future re-import of the same source file will produce"
        )
    else:
        plan.keep_raw_id = ordered[0]
        if recomputed is None:
            plan.keep_reason = "earliest ImportedAt (could not recompute a current RawID to compare against - unresolved SourceBank)"
        else:
            plan.keep_reason = (
                f"earliest ImportedAt (current code recomputes {recomputed}, which matches "
                "neither candidate - a future re-import may create a third variant; investigate before relying on this)"
            )
    plan.remove_raw_ids = [rid for rid in ordered if rid != plan.keep_raw_id]

    for i in unified_idx_by_id:
        record = unified_records[i]
        if str(record.get("RawID", "")) in plan.remove_raw_ids and not _is_blank(record.get("Review/Notes")):
            plan.blocking_reviewed.append(record)

    return plan


def print_plan(plan: Plan) -> None:
    if plan.not_found:
        print(f"RawID(s) not found in {RAW_SHEET}: {', '.join(plan.not_found)}")
        print("Nothing to do.")
        return

    print(f"Comparing {len(plan.raw_ids)} RawTransactions row(s): {', '.join(plan.raw_ids)}")
    print()

    compare_cols = [c for c in plan.raw_headers if c not in IDENTITY_EXCLUDED_COLUMNS]
    rows = [plan.raw_records[plan.raw_idx_by_id[rid]] for rid in plan.raw_ids]
    df = pd.DataFrame([{c: r.get(c, "") for c in compare_cols} for r in rows], index=plan.raw_ids)
    print(df.T.to_string())
    print()

    if plan.mismatches:
        print(f"REFUSING: {len(plan.mismatches)} field(s) differ, so these are NOT verified as the same "
              f"real transaction: {', '.join(plan.mismatches)}")
        print("Either these are genuinely different transactions, or --raw-ids named the wrong rows.")
        return

    print(f"Verified: every field matches except {', '.join(sorted(IDENTITY_EXCLUDED_COLUMNS))} (expected to differ).")
    print()
    print(f"Would KEEP:   {plan.keep_raw_id}  ({plan.keep_reason})")
    print(f"Would REMOVE: {', '.join(plan.remove_raw_ids)}")
    print()

    for i in plan.unified_idx_by_id:
        r = plan.unified_records[i]
        print(f"  UnifiedTransactions row: RawID={r.get('RawID')} UnifiedID={r.get('UnifiedID')} "
              f"Supercategory={r.get('Supercategory')!r} Category={r.get('Category')!r} "
              f"Subcategory={r.get('Subcategory')!r} Review/Notes={r.get('Review/Notes')!r}")
    print()

    if plan.blocking_reviewed:
        print(f"BLOCKED: {len(plan.blocking_reviewed)} row(s) to be removed have non-blank Review/Notes "
              "(real manual work) - pass --include-reviewed to override.")
    else:
        print("No blocking issues found.")


def apply_plan(sheets: dict[str, list[list[object]]], plan: Plan) -> dict[str, tuple[int, int]]:
    stats: dict[str, tuple[int, int]] = {}

    raw_before = len(plan.raw_records)
    raw_kept = [r for r in plan.raw_records if str(r.get("RawID", "")) not in plan.remove_raw_ids]
    sheets[RAW_SHEET] = records_to_sheet_values(plan.raw_headers, raw_kept)
    stats[RAW_SHEET] = (raw_before, len(raw_kept))

    unified_before = len(plan.unified_records)
    unified_kept = [r for r in plan.unified_records if str(r.get("RawID", "")) not in plan.remove_raw_ids]
    sheets[UNIFIED_SHEET] = records_to_sheet_values(plan.unified_headers, unified_kept)
    stats[UNIFIED_SHEET] = (unified_before, len(unified_kept))

    return stats


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--raw-ids", required=True, help="Comma-separated RawIDs believed to be the same real transaction.")
    parser.add_argument("--workbook", type=Path, default=DEFAULT_BUDGETING_WORKBOOK)
    parser.add_argument("--apply", action="store_true", help="Actually remove the row(s). Without this, report only.")
    parser.add_argument("--confirm", default=None, help="Must exactly match the sorted, comma-joined RawIDs to be removed.")
    parser.add_argument("--include-reviewed", action="store_true", help="Allow removing a row with non-blank Review/Notes.")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()
    raw_ids = [rid.strip() for rid in args.raw_ids.split(",") if rid.strip()]
    if len(raw_ids) < 2:
        print("Need at least 2 --raw-ids to compare.")
        return

    sheets = read_workbook_values_only(args.workbook)
    plan = build_plan(sheets, raw_ids)
    print_plan(plan)

    if plan.not_found or plan.mismatches or not plan.keep_raw_id:
        return

    blocked = plan.blocking_reviewed and not args.include_reviewed
    if blocked:
        print()
        print("Refusing to proceed while blocking issues remain (see above).")
        return

    if not args.apply:
        print()
        print("Dry-run only (no --apply given). Real workbook was not touched.")
        return

    expected_confirm = ",".join(sorted(plan.remove_raw_ids))
    if args.confirm != expected_confirm:
        print()
        print(f"Refusing to --apply: --confirm must exactly equal {expected_confirm!r} "
              f"(the sorted, comma-joined RawIDs that would be removed).")
        return

    stats = apply_plan(sheets, plan)

    total_removed = sum(before - after for before, after in stats.values())
    details_parts = [f"{sheet}: {before}->{after} (removed {before - after})" for sheet, (before, after) in stats.items()]
    details = (
        f"Removed confirmed duplicate RawID(s) {plan.remove_raw_ids} (kept {plan.keep_raw_id!r}, "
        f"earliest ImportedAt). {'; '.join(details_parts)}."
    )

    append_changelog_row(
        sheets,
        script=SCRIPT_NAME,
        action=f"Remove duplicate transaction: {', '.join(plan.remove_raw_ids)}",
        sheet=", ".join(stats.keys()),
        rows_updated=total_removed,
        rows_removed=total_removed,
        status="Completed",
        details=details,
        backup_file=None,
    )

    backup_path, _, write_stats = replace_with_fresh_workbook(
        args.workbook,
        sheets,
        backup_label="before_remove_duplicate_transaction",
        excel_tables=False,
    )
    print(f"Backup written: {backup_path}")
    print(f"Removed {total_removed} row(s) across {len(stats)} sheet(s). {write_stats['sheets_written']} sheet(s) written.")
    print("ChangeLog entry appended.")


if __name__ == "__main__":
    main()
