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

    print("""usage: python utilities/review_cleanup_unified_duplicates.py [--dry-run] [--help]

Create a review workbook for suspicious UnifiedTransactions duplicates and optionally remove reviewed duplicates.

options:
  --dry-run   Show what would be changed without modifying files.
  -h, --help  Show this help message and exit.

examples:
  python utilities/review_cleanup_unified_duplicates.py --dry-run\n  python utilities/review_cleanup_unified_duplicates.py
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
    UNIFIED_COLUMNS,
    UNIFIED_SHEET,
    clean_dataframe_for_excel,
    clean_for_excel,
    normalize_source_metadata,
    append_change_log_entry,
)


DEFAULT_WORKBOOK = PROJECT_ROOT / "output" / "budgeting" / "ParsedTransactions.xlsx"


TECHNICAL_IDENTITY_COLUMNS = [
    "SourceAccount",
    "SourceBank",
    "TransactionType",
    "Date",
    "Amount",
    "RawReceiver",
    "Description",
    "Message",
    "Currency",
    "Rate",
    "SourceFile",
]

MANUAL_OR_RULE_COLUMNS = [
    "NormalizedReceiver",
    "Include",
    "Owner",
    "Supercategory",
    "Category",
    "Subcategory",
    "Review/Notes",
]


def text(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def timestamp_for_filename() -> str:
    return pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")


def append_dataframe_to_worksheet(wb, sheet_name: str, df: pd.DataFrame) -> None:
    df = clean_dataframe_for_excel(df)
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


def technical_key(row: pd.Series) -> str:
    parts = []
    for col in TECHNICAL_IDENTITY_COLUMNS:
        if col == "SourceFile":
            # SourceFile is useful for review but should not decide whether two
            # unified rows represent the same underlying transaction.
            continue
        parts.append(text(row.get(col, "")).upper())
    return " | ".join(parts)


def row_has_manual_or_rule_values(row: pd.Series) -> bool:
    for col in MANUAL_OR_RULE_COLUMNS:
        if text(row.get(col, "")) != "":
            return True
    return False


def score_row_to_keep(row: pd.Series) -> int:
    """
    Higher score = better row to keep.

    Prefer rows with real manual/rule fields. Penalise the old bug pattern where
    NormalizedReceiver equals RawReceiver and no category/comment fields exist.
    """
    score = 0

    normalized = text(row.get("NormalizedReceiver", ""))
    raw = text(row.get("RawReceiver", ""))

    if normalized:
        score += 10
    if normalized and raw and normalized.upper() == raw.upper():
        score -= 20

    for col in ["Include", "Owner", "Supercategory", "Category", "Subcategory", "Comments"]:
        if text(row.get(col, "")):
            score += 5

    # Prefer rows not created by the bug if everything else ties.
    return score


def find_suspicious_duplicates(unified: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = normalize_source_metadata(unified).copy()
    df["_TechnicalDuplicateKey"] = df.apply(technical_key, axis=1)
    counts = df["_TechnicalDuplicateKey"].value_counts(dropna=False)
    duplicate_keys = set(counts[counts > 1].index)

    if not duplicate_keys:
        return df.drop(columns=["_TechnicalDuplicateKey"]), pd.DataFrame()

    dupes = df[df["_TechnicalDuplicateKey"].isin(duplicate_keys)].copy()
    dupes["_DuplicateGroupSize"] = dupes["_TechnicalDuplicateKey"].map(counts)
    dupes["_KeepScore"] = dupes.apply(score_row_to_keep, axis=1)

    # Mark suggested keep/delete without doing anything automatically in dry-run review.
    dupes["_SuggestedAction"] = ""
    for key, group in dupes.groupby("_TechnicalDuplicateKey", dropna=False):
        best_idx = group["_KeepScore"].idxmax()
        dupes.loc[group.index, "_SuggestedAction"] = "DELETE_CANDIDATE"
        dupes.loc[best_idx, "_SuggestedAction"] = "KEEP_CANDIDATE"

    review_cols = [
        "_SuggestedAction",
        "_KeepScore",
        "_DuplicateGroupSize",
        "_TechnicalDuplicateKey",
    ] + [col for col in dupes.columns if col not in {
        "_SuggestedAction",
        "_KeepScore",
        "_DuplicateGroupSize",
        "_TechnicalDuplicateKey",
    }]

    dupes = dupes[review_cols].sort_values(["_TechnicalDuplicateKey", "_SuggestedAction"], ascending=[True, False])

    # Build cleaned version only by removing suggested delete candidates.
    delete_indices = dupes[dupes["_SuggestedAction"] == "DELETE_CANDIDATE"].index
    cleaned = df.drop(index=delete_indices).drop(columns=["_TechnicalDuplicateKey"])

    return cleaned, dupes


def write_review(review_path: Path, summary: pd.DataFrame, duplicates: pd.DataFrame) -> None:
    wb = Workbook()
    default = wb.active
    wb.remove(default)
    append_dataframe_to_worksheet(wb, "Summary", summary)
    if len(duplicates) > 0:
        append_dataframe_to_worksheet(wb, "UnifiedDuplicateReview", duplicates)
    wb.save(review_path)


def replace_unified_sheet(workbook_path: Path, unified: pd.DataFrame, *, rows_before: int = 0, review_path: Path | None = None) -> None:
    backup_path = workbook_path.with_name(workbook_path.stem + "_backup_before_unified_duplicate_cleanup" + workbook_path.suffix)
    tmp_path = workbook_path.with_name(workbook_path.stem + "_tmp_unified_duplicate_cleanup" + workbook_path.suffix)

    shutil.copy2(workbook_path, backup_path)

    wb = load_workbook(workbook_path)
    if UNIFIED_SHEET not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Sheet not found: {UNIFIED_SHEET}")

    old_index = wb.sheetnames.index(UNIFIED_SHEET)
    del wb[UNIFIED_SHEET]
    ws = wb.create_sheet(UNIFIED_SHEET, old_index)

    out = unified.copy()
    if set(UNIFIED_COLUMNS).issubset(out.columns):
        out = out[UNIFIED_COLUMNS]

    ws.append(list(out.columns))
    for row in clean_dataframe_for_excel(out).itertuples(index=False, name=None):
        ws.append([clean_for_excel(v) for v in row])

    ws.freeze_panes = "A2"
    if len(out) > 0 and len(out.columns) > 0:
        end_col = ws.cell(row=1, column=len(out.columns)).column_letter
        end_row = len(out) + 1
        ws.auto_filter.ref = f"A1:{end_col}{end_row}"

    wb.save(tmp_path)
    wb.close()

    test_wb = load_workbook(tmp_path, read_only=True)
    test_wb.close()

    os.replace(tmp_path, workbook_path)

    append_change_log_entry(
        workbook_path,
        script="utilities/review_cleanup_unified_duplicates.py",
        action="Clean suspicious UnifiedTransactions duplicates",
        sheet=UNIFIED_SHEET,
        rows_before=rows_before,
        rows_after=len(unified),
        rows_added=0,
        rows_updated=0,
        rows_removed=max(rows_before - len(unified), 0) if rows_before else "",
        review_file=review_path or "",
        backup_file=backup_path,
        status="Completed",
        details="Removed duplicate UnifiedTransactions rows after review.",
    )

    print(f"Workbook updated. Backup created: {backup_path}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Review/clean suspicious UnifiedTransactions duplicates caused by receiver-normalisation regression."
    )
    parser.add_argument(
        "--workbook",
        default=str(DEFAULT_WORKBOOK),
        help="ParsedTransactions workbook. Defaults to output/budgeting/ParsedTransactions.xlsx.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Write review only, do not modify workbook.")
    parser.add_argument(
        "--review",
        default="",
        help="Optional review workbook path. Defaults to timestamped file next to the workbook.",
    )

    args = parser.parse_args()
    workbook_path = Path(args.workbook).expanduser().resolve()

    unified = pd.read_excel(workbook_path, sheet_name=UNIFIED_SHEET, dtype=object, engine="openpyxl")
    cleaned, duplicates = find_suspicious_duplicates(unified)

    summary = pd.DataFrame([
        {
            "RowsBefore": len(unified),
            "RowsAfterIfCleaned": len(cleaned),
            "RowsRemovedIfCleaned": len(unified) - len(cleaned),
            "DuplicateReviewRows": len(duplicates),
            "Basis": "Technical duplicate key excluding NormalizedReceiver/category/manual fields and SourceFile",
        }
    ])

    if args.review:
        review_path = Path(args.review).expanduser().resolve()
    else:
        review_path = workbook_path.with_name(f"{workbook_path.stem}_unified_duplicate_review_{timestamp_for_filename()}.xlsx")

    write_review(review_path, summary, duplicates)

    print("Unified duplicate review")
    print(f"Workbook: {workbook_path}")
    print(f"Rows before: {len(unified)}")
    print(f"Rows after if cleaned: {len(cleaned)}")
    print(f"Rows removed if cleaned: {len(unified) - len(cleaned)}")
    print(f"Review workbook: {review_path}")

    if args.dry_run:
        print("Dry run only: workbook was not modified.")
        return

    if len(unified) == len(cleaned):
        print("No rows to remove.")
        return

    replace_unified_sheet(workbook_path, cleaned, rows_before=len(unified), review_path=review_path)


if __name__ == "__main__":
    main()
