from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from finance_parser.common import normalise_text
from finance_parser.investments.investment_common import (
    INVESTMENT_IMPORT_LOG_COLUMNS,
    INVESTMENT_IMPORT_LOG_SHEET,
    INVESTMENT_RAW_COLUMNS,
    INVESTMENT_RAW_SHEET,
    INVESTMENT_TRANSACTIONS_COLUMNS,
    INVESTMENT_TRANSACTIONS_SHEET,
    read_dynamic_sheet,
    read_sheet,
    write_investment_output_workbook,
)

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "ParsedInvestments.xlsx"

# Broker-specific raw export dump sheets that may exist alongside the fixed-
# schema sheets - untouched by this migration, just preserved as-is so the
# fresh-rebuild write doesn't drop them.
BROKER_RAW_SHEETS = [
    "CoinmotionRawExport",
    "EvliRawExport",
    "NordnetRawExport",
    "OPInvestmentRawExport",
    "SeligsonRawExport",
]


def migrate_dataframe(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """
    One-time fix for a real bug: EVLI's parser used to set Portfolio to the
    export's plan-cycle label (e.g. "Plan Cycle 2022", "SIS Dividend")
    instead of a real per-account identifier, fragmenting one real holding's
    cumulative quantity across 5 fake "portfolios" - fixed in
    parsers/evli.py, which now hardcodes Portfolio="EVLI" and captures the
    plan-cycle label in its own PlanCycle column instead.

    This moves already-imported rows to the same shape in place: any EVLI
    row whose Portfolio isn't already "EVLI" is by definition an old,
    unmigrated row (the fixed parser never produces anything else) - its
    current Portfolio value becomes PlanCycle, and Portfolio becomes "EVLI".
    Idempotent: a second run finds nothing left to migrate. Deliberately
    does not touch InvestmentRawID/InvestmentTransactionID, so a later
    re-import of the same source files still dedupes correctly against
    these rows.
    """
    out = df.copy()
    is_evli = out["Broker"].map(normalise_text) == "EVLI"
    needs_migration = is_evli & (out["Portfolio"].map(normalise_text) != "EVLI")

    out.loc[needs_migration, "PlanCycle"] = out.loc[needs_migration, "Portfolio"]
    out.loc[needs_migration, "Portfolio"] = "EVLI"

    return out, int(needs_migration.sum())


def migrate_workbook(workbook_path: Path, *, dry_run: bool = False) -> dict[str, object]:
    raw = read_sheet(workbook_path, INVESTMENT_RAW_SHEET, INVESTMENT_RAW_COLUMNS)
    transactions = read_sheet(workbook_path, INVESTMENT_TRANSACTIONS_SHEET, INVESTMENT_TRANSACTIONS_COLUMNS)
    import_log = read_sheet(workbook_path, INVESTMENT_IMPORT_LOG_SHEET, INVESTMENT_IMPORT_LOG_COLUMNS)

    migrated_raw, raw_count = migrate_dataframe(raw)
    migrated_transactions, transactions_count = migrate_dataframe(transactions)

    stats: dict[str, object] = {
        "raw_rows_migrated": raw_count,
        "transaction_rows_migrated": transactions_count,
    }

    if dry_run:
        return stats

    extra_sheets = {}
    for sheet_name in BROKER_RAW_SHEETS:
        sheet_df = read_dynamic_sheet(workbook_path, sheet_name)
        if len(sheet_df) > 0:
            extra_sheets[sheet_name] = sheet_df

    write_investment_output_workbook(
        workbook_path,
        migrated_raw[INVESTMENT_RAW_COLUMNS],
        migrated_transactions[INVESTMENT_TRANSACTIONS_COLUMNS],
        import_log[INVESTMENT_IMPORT_LOG_COLUMNS],
        extra_sheets=extra_sheets,
    )

    return stats


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="One-time migration: EVLI rows had Portfolio set to their plan-cycle label "
                    "instead of a real account identifier (fixed in parsers/evli.py). Moves the "
                    "old Portfolio value into a new PlanCycle column and sets Portfolio to 'EVLI', "
                    "in place. InvestmentRawID/InvestmentTransactionID are never touched, so future "
                    "re-imports of the same source files still dedupe correctly against these rows."
    )
    parser.add_argument("--workbook", default=str(DEFAULT_WORKBOOK))
    parser.add_argument("--dry-run", action="store_true")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()
    workbook_path = Path(args.workbook).expanduser().resolve()

    stats = migrate_workbook(workbook_path, dry_run=args.dry_run)

    print("EVLI portfolio migration dry run complete." if args.dry_run else "EVLI portfolio migration complete.")
    print(f"RawInvestmentTransactions rows migrated: {stats['raw_rows_migrated']}")
    print(f"InvestmentTransactions rows migrated:    {stats['transaction_rows_migrated']}")

    if args.dry_run:
        print()
        print("Dry run only: workbook was not modified.")
    else:
        print()
        print(f"Output workbook: {workbook_path}")


if __name__ == "__main__":
    main()
