from __future__ import annotations

def _print_cli_help_and_exit() -> None:
    """
    Early help handler.

    This runs before heavy imports and workbook setup so `--help` works even if
    project-root imports or dependencies are currently broken.
    """
    import sys

    if not any(arg in {"-h", "--help"} for arg in sys.argv[1:]):
        return

    print("""usage: python utilities/cleanup_budgeting_source_accounts.py [--dry-run] [--help]

Review and optionally clean polluted budgeting SourceAccount labels and duplicate rows in ParsedTransactions.xlsx.

options:
  --dry-run   Show what would be changed without modifying files.
  -h, --help  Show this help message and exit.

examples:
  python utilities/cleanup_budgeting_source_accounts.py --dry-run\n  python utilities/cleanup_budgeting_source_accounts.py
""")
    raise SystemExit(0)


_print_cli_help_and_exit()


from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import argparse
import os
import shutil
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook

from finance_parser.common import (
    IMPORT_LOG_COLUMNS,
    RAW_COLUMNS,
    UNIFIED_COLUMNS,
    IMPORT_LOG_SHEET,
    RAW_SHEET,
    UNIFIED_SHEET,
    add_canonical_transaction_key,
    clean_dataframe_for_excel,
    clean_for_excel,
    normalize_source_metadata,
    append_change_log_entry,
)


DEFAULT_WORKBOOK = PROJECT_ROOT / "output" / "budgeting" / "ParsedTransactions.xlsx"

BANK_RAW_SHEETS = [
    "OPRawExport",
    "NorwegianRawExport",
    "NordeaRawExport",
    "SpankkiRawExport",
]


def timestamp_for_filename() -> str:
    return pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")


def default_review_path(workbook_path: Path) -> Path:
    return workbook_path.with_name(f"{workbook_path.stem}_cleanup_review_{timestamp_for_filename()}.xlsx")


def append_dataframe_to_worksheet(wb, sheet_name: str, df: pd.DataFrame) -> None:
    df = clean_dataframe_for_excel(df)

    # Excel sheet names max 31 chars.
    sheet_name = sheet_name[:31]

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]

    ws = wb.create_sheet(sheet_name)
    ws.append([clean_for_excel(c) for c in df.columns])

    for row in df.itertuples(index=False, name=None):
        ws.append([clean_for_excel(v) for v in row])

    ws.freeze_panes = "A2"

    if len(df) > 0 and len(df.columns) > 0:
        end_col = ws.cell(row=1, column=len(df.columns)).column_letter
        end_row = len(df) + 1
        ws.auto_filter.ref = f"A1:{end_col}{end_row}"


def read_sheet_if_exists(path: Path, sheet_name: str) -> pd.DataFrame | None:
    try:
        return pd.read_excel(path, sheet_name=sheet_name, dtype=object, engine="openpyxl")
    except ValueError:
        return None


def duplicate_rows_by_key(df: pd.DataFrame, key_cols: list[str]) -> pd.DataFrame:
    if len(df) == 0 or not key_cols:
        return pd.DataFrame()

    available_key_cols = [col for col in key_cols if col in df.columns]
    if not available_key_cols:
        return pd.DataFrame()

    keyed = df.copy()
    keyed["_CleanupDuplicateKey"] = keyed[available_key_cols].astype(str).agg(" | ".join, axis=1)

    counts = keyed["_CleanupDuplicateKey"].value_counts(dropna=False)
    duplicate_keys = set(counts[counts > 1].index)

    if not duplicate_keys:
        return pd.DataFrame()

    duplicate_rows = keyed[keyed["_CleanupDuplicateKey"].isin(duplicate_keys)].copy()
    duplicate_rows["_CleanupDuplicateGroupSize"] = duplicate_rows["_CleanupDuplicateKey"].map(counts)

    # Put debug columns first.
    debug_cols = ["_CleanupDuplicateKey", "_CleanupDuplicateGroupSize"]
    other_cols = [col for col in duplicate_rows.columns if col not in debug_cols]
    duplicate_rows = duplicate_rows[debug_cols + other_cols]

    return duplicate_rows.sort_values(["_CleanupDuplicateKey"])


def raw_or_unified_duplicate_rows(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    normalised = normalize_source_metadata(df)
    keyed = add_canonical_transaction_key(normalised)

    duplicates = duplicate_rows_by_key(keyed, ["_CanonicalTransactionKey"])
    deduped = keyed.drop_duplicates(subset=["_CanonicalTransactionKey"], keep="last")
    deduped = deduped.drop(columns=["_CanonicalTransactionKey"])

    return deduped, duplicates


def bank_raw_dedup_key_cols(df: pd.DataFrame) -> list[str]:
    """
    Bank raw sheets preserve closer-to-export syntax.

    For cleanup, duplicates should be evaluated by the original export content,
    not by metadata columns such as SourceFile, ExportDate, ImportedAt, or the
    generated BankRawExportID. This is why SpankkiRawExport may show duplicates
    even when the parser itself was not corrupted: the same bank row may have
    been imported from overlapping exports or repeated test imports.
    """
    metadata_cols = {
        "BankRawExportID",
        "SourceBank",
        "SourceAccount",
        "SourceFile",
        "ExportDate",
        "ImportedAt",
    }
    return [col for col in df.columns if col not in metadata_cols]


def canonicalise_bank_raw_sheet(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    out = df.copy()

    if "SourceBank" in out.columns and "SourceAccount" in out.columns:
        out = normalize_source_metadata(out)

    key_cols = bank_raw_dedup_key_cols(out)
    duplicates = duplicate_rows_by_key(out, key_cols)

    if key_cols:
        deduped = out.drop_duplicates(subset=key_cols, keep="last")
    elif "BankRawExportID" in out.columns:
        deduped = out.drop_duplicates(subset=["BankRawExportID"], keep="last")
    else:
        deduped = out.drop_duplicates(keep="last")

    return deduped, duplicates


def write_review_workbook(
    review_path: Path,
    summary_rows: list[dict[str, object]],
    duplicate_sheets: dict[str, pd.DataFrame],
) -> None:
    review_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    default = wb.active
    wb.remove(default)

    summary = pd.DataFrame(summary_rows)
    append_dataframe_to_worksheet(wb, "Summary", summary)

    for sheet_name, df in duplicate_sheets.items():
        if len(df) == 0:
            continue
        append_dataframe_to_worksheet(wb, sheet_name, df)

    wb.save(review_path)


def cleanup_workbook(
    workbook_path: Path,
    dry_run: bool = False,
    review_path: Path | None = None,
    write_review: bool = True,
) -> None:
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    raw = pd.read_excel(workbook_path, sheet_name=RAW_SHEET, dtype=object, engine="openpyxl")
    unified = pd.read_excel(workbook_path, sheet_name=UNIFIED_SHEET, dtype=object, engine="openpyxl")
    log = pd.read_excel(workbook_path, sheet_name=IMPORT_LOG_SHEET, dtype=object, engine="openpyxl")

    raw_before = len(raw)
    unified_before = len(unified)

    raw_clean, raw_duplicates = raw_or_unified_duplicate_rows(raw)
    unified_clean, unified_duplicates = raw_or_unified_duplicate_rows(unified)

    raw_after = len(raw_clean)
    unified_after = len(unified_clean)

    bank_raw_results: dict[str, pd.DataFrame] = {}
    bank_raw_stats: dict[str, tuple[int, int]] = {}
    duplicate_sheets: dict[str, pd.DataFrame] = {
        "RawTransactionsDuplicates": raw_duplicates,
        "UnifiedDuplicates": unified_duplicates,
    }

    for sheet in BANK_RAW_SHEETS:
        df = read_sheet_if_exists(workbook_path, sheet)
        if df is None:
            continue

        before = len(df)
        cleaned, duplicates = canonicalise_bank_raw_sheet(df)
        after = len(cleaned)

        bank_raw_results[sheet] = cleaned
        bank_raw_stats[sheet] = (before, after)
        duplicate_sheets[f"{sheet}_Duplicates"] = duplicates

    summary_rows = [
        {
            "Sheet": RAW_SHEET,
            "RowsBefore": raw_before,
            "RowsAfter": raw_after,
            "RowsRemovedIfRun": raw_before - raw_after,
            "DuplicateRowsShownInReview": len(raw_duplicates),
            "DeduplicationBasis": "Canonical transaction key after SourceAccount normalisation",
        },
        {
            "Sheet": UNIFIED_SHEET,
            "RowsBefore": unified_before,
            "RowsAfter": unified_after,
            "RowsRemovedIfRun": unified_before - unified_after,
            "DuplicateRowsShownInReview": len(unified_duplicates),
            "DeduplicationBasis": "Canonical transaction key after SourceAccount normalisation",
        },
    ]

    for sheet, (before, after) in bank_raw_stats.items():
        duplicates = duplicate_sheets.get(f"{sheet}_Duplicates", pd.DataFrame())
        summary_rows.append(
            {
                "Sheet": sheet,
                "RowsBefore": before,
                "RowsAfter": after,
                "RowsRemovedIfRun": before - after,
                "DuplicateRowsShownInReview": len(duplicates),
                "DeduplicationBasis": "Original export columns only; metadata ignored",
            }
        )

    print("Budgeting workbook cleanup")
    print(f"Workbook: {workbook_path}")
    print(f"RawTransactions:     {raw_before} -> {raw_after} ({raw_before - raw_after} duplicate rows removed)")
    print(f"UnifiedTransactions: {unified_before} -> {unified_after} ({unified_before - unified_after} duplicate rows removed)")
    for sheet, (before, after) in bank_raw_stats.items():
        print(f"{sheet}: {before} -> {after} ({before - after} duplicate rows removed)")

    if write_review:
        if review_path is None:
            review_path = default_review_path(workbook_path)
        write_review_workbook(review_path, summary_rows, duplicate_sheets)
        print(f"Review workbook written: {review_path}")

    if dry_run:
        print("Dry run only: workbook was not modified.")
        return

    tmp_path = workbook_path.with_name(workbook_path.stem + "_tmp_cleanup" + workbook_path.suffix)
    backup_path = workbook_path.with_name(workbook_path.stem + "_backup_before_cleanup" + workbook_path.suffix)
    shutil.copy2(workbook_path, backup_path)

    wb = load_workbook(workbook_path)

    # Replace canonical sheets while keeping sheet order roughly stable.
    for sheet_name, df in [
        (RAW_SHEET, raw_clean[RAW_COLUMNS] if set(RAW_COLUMNS).issubset(raw_clean.columns) else raw_clean),
        (UNIFIED_SHEET, unified_clean[UNIFIED_COLUMNS] if set(UNIFIED_COLUMNS).issubset(unified_clean.columns) else unified_clean),
        (IMPORT_LOG_SHEET, log[IMPORT_LOG_COLUMNS] if set(IMPORT_LOG_COLUMNS).issubset(log.columns) else log),
    ]:
        if sheet_name in wb.sheetnames:
            old_index = wb.sheetnames.index(sheet_name)
            del wb[sheet_name]
            ws = wb.create_sheet(sheet_name, old_index)
            ws.append(list(df.columns))
            for row in clean_dataframe_for_excel(df).itertuples(index=False, name=None):
                ws.append([clean_for_excel(v) for v in row])
            ws.freeze_panes = "A2"
            if len(df) > 0 and len(df.columns) > 0:
                end_col = ws.cell(row=1, column=len(df.columns)).column_letter
                end_row = len(df) + 1
                ws.auto_filter.ref = f"A1:{end_col}{end_row}"

    for sheet_name, df in bank_raw_results.items():
        append_dataframe_to_worksheet(wb, sheet_name, df)

    wb.save(tmp_path)
    wb.close()

    test_wb = load_workbook(tmp_path, read_only=True)
    test_wb.close()

    os.replace(tmp_path, workbook_path)

    total_rows_removed = (raw_before - raw_after) + (unified_before - unified_after)
    total_rows_removed += sum(before - after for before, after in bank_raw_stats.values())

    append_change_log_entry(
        workbook_path,
        script="utilities/cleanup_budgeting_source_accounts.py",
        action="Clean SourceAccount labels and duplicate rows",
        sheet="RawTransactions, UnifiedTransactions, bank raw sheets",
        rows_before=raw_before + unified_before + sum(before for before, _ in bank_raw_stats.values()),
        rows_after=raw_after + unified_after + sum(after for _, after in bank_raw_stats.values()),
        rows_added=0,
        rows_updated="",
        rows_removed=total_rows_removed,
        review_file=review_path or "",
        backup_file=backup_path,
        status="Completed",
        details=f"Raw removed: {raw_before - raw_after}; unified removed: {unified_before - unified_after}; bank raw removed: {sum(before - after for before, after in bank_raw_stats.values())}",
    )

    print(f"Workbook cleaned. Backup created: {backup_path}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="One-off cleanup for polluted SourceAccount labels and duplicate budgeting rows."
    )
    parser.add_argument(
        "--workbook",
        default=str(DEFAULT_WORKBOOK),
        help="ParsedTransactions workbook. Defaults to output/budgeting/ParsedTransactions.xlsx.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Show changes without modifying workbook.")
    parser.add_argument(
        "--review",
        default="",
        help="Optional review workbook path. By default a timestamped review workbook is written next to ParsedTransactions.xlsx.",
    )
    parser.add_argument(
        "--no-review",
        action="store_true",
        help="Do not write a review workbook.",
    )

    args = parser.parse_args()

    review_path = Path(args.review).expanduser().resolve() if args.review else None

    cleanup_workbook(
        Path(args.workbook).expanduser().resolve(),
        dry_run=args.dry_run,
        review_path=review_path,
        write_review=not args.no_review,
    )


if __name__ == "__main__":
    main()
