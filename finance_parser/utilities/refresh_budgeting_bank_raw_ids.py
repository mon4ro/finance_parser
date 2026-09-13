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

    print("""usage: python utilities/refresh_budgeting_bank_raw_ids.py [--dry-run] [--help]

Refresh budgeting bank raw export IDs after hash logic changes and optionally remove duplicate raw rows.

options:
  --dry-run   Show what would be changed without modifying files.
  -h, --help  Show this help message and exit.

examples:
  python utilities/refresh_budgeting_bank_raw_ids.py --dry-run\n  python utilities/refresh_budgeting_bank_raw_ids.py
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
from openpyxl import load_workbook

from finance_parser.common import clean_dataframe_for_excel, clean_for_excel, normalize_source_metadata, append_change_log_entry
from finance_parser.budgeting.parsers import op, norwegian, nordea, spankki


DEFAULT_WORKBOOK = PROJECT_ROOT / "output" / "budgeting" / "ParsedTransactions.xlsx"

SHEET_TO_PARSER = {
    "OPRawExport": op,
    "NorwegianRawExport": norwegian,
    "NordeaRawExport": nordea,
    "SpankkiRawExport": spankki,
}


def append_dataframe_to_worksheet(wb, sheet_name: str, df: pd.DataFrame) -> None:
    df = clean_dataframe_for_excel(df)

    if sheet_name in wb.sheetnames:
        old_index = wb.sheetnames.index(sheet_name)
        del wb[sheet_name]
        ws = wb.create_sheet(sheet_name, old_index)
    else:
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


def refresh_sheet_ids(sheet_name: str, df: pd.DataFrame, parser_module) -> tuple[pd.DataFrame, int, int]:
    out = df.copy()

    if "BankRawExportID" not in out.columns:
        out.insert(0, "BankRawExportID", "")

    if "SourceBank" in out.columns and "SourceAccount" in out.columns:
        out = normalize_source_metadata(out)

    old_ids = out["BankRawExportID"].astype(str).tolist()

    new_ids = []
    for _, row in out.iterrows():
        new_ids.append(parser_module.make_bank_raw_export_id(row, ""))

    out["BankRawExportID"] = new_ids

    changed = sum(1 for old, new in zip(old_ids, new_ids) if old != new)
    before = len(out)

    # With the new metadata-independent IDs, duplicate overlapping-export rows
    # should collapse here.
    out = out.drop_duplicates(subset=["BankRawExportID"], keep="last")

    removed = before - len(out)
    return out, changed, removed


def refresh_workbook(workbook_path: Path, dry_run: bool = False) -> None:
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    results = {}
    stats = []

    for sheet_name, parser_module in SHEET_TO_PARSER.items():
        df = read_sheet_if_exists(workbook_path, sheet_name)
        if df is None:
            continue

        refreshed, changed, removed = refresh_sheet_ids(sheet_name, df, parser_module)
        results[sheet_name] = refreshed
        stats.append({
            "Sheet": sheet_name,
            "RowsBefore": len(df),
            "RowsAfter": len(refreshed),
            "IDsChanged": changed,
            "DuplicateRowsRemoved": removed,
        })

    print("Bank raw export ID refresh")
    print(f"Workbook: {workbook_path}")
    for row in stats:
        print(
            f"{row['Sheet']}: rows {row['RowsBefore']} -> {row['RowsAfter']}, "
            f"IDs changed {row['IDsChanged']}, duplicates removed {row['DuplicateRowsRemoved']}"
        )

    if dry_run:
        print("Dry run only: workbook was not modified.")
        return

    backup_path = workbook_path.with_name(workbook_path.stem + "_backup_before_raw_id_refresh" + workbook_path.suffix)
    tmp_path = workbook_path.with_name(workbook_path.stem + "_tmp_raw_id_refresh" + workbook_path.suffix)
    shutil.copy2(workbook_path, backup_path)

    wb = load_workbook(workbook_path)

    for sheet_name, df in results.items():
        append_dataframe_to_worksheet(wb, sheet_name, df)

    wb.save(tmp_path)
    wb.close()

    test_wb = load_workbook(tmp_path, read_only=True)
    test_wb.close()

    os.replace(tmp_path, workbook_path)

    total_changed = sum(row["IDsChanged"] for row in stats)
    total_removed = sum(row["DuplicateRowsRemoved"] for row in stats)
    append_change_log_entry(
        workbook_path,
        script="utilities/refresh_budgeting_bank_raw_ids.py",
        action="Refresh bank raw export IDs",
        sheet=", ".join(results.keys()),
        rows_before=sum(row["RowsBefore"] for row in stats),
        rows_after=sum(row["RowsAfter"] for row in stats),
        rows_added=0,
        rows_updated=total_changed,
        rows_removed=total_removed,
        backup_file=backup_path,
        status="Completed",
        details=f"Bank raw IDs changed: {total_changed}; duplicate raw rows removed: {total_removed}",
    )

    print(f"Workbook updated. Backup created: {backup_path}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Refresh bank raw export IDs after hash-definition changes."
    )
    parser.add_argument(
        "--workbook",
        default=str(DEFAULT_WORKBOOK),
        help="ParsedTransactions workbook. Defaults to output/budgeting/ParsedTransactions.xlsx.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Show changes without writing.")
    args = parser.parse_args()

    refresh_workbook(Path(args.workbook).expanduser().resolve(), dry_run=args.dry_run)


if __name__ == "__main__":
    main()
