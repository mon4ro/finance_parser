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

BROKER_RAW_SHEETS = [
    "CoinmotionRawExport",
    "EvliRawExport",
    "NordnetRawExport",
    "OPInvestmentRawExport",
    "SeligsonRawExport",
]


def _migrate(df: pd.DataFrame, type_column: str, buy_value: str, sell_value: str) -> tuple[pd.DataFrame, int]:
    """
    One-time fix for a real bug: the Seligson parser stored CashAmount
    straight from the export's own "Summa" column (positive = subscription
    from the FUND's perspective) instead of flipping it to this project's
    investor cash-flow convention (negative = money paid out to buy) -
    fixed in parsers/seligson.py. This flips already-imported rows in
    place: a BUY/MERKINTÄ row with a positive CashAmount, or a SELL/
    LUNASTUS row with a negative one, is by definition unmigrated (the
    fixed parser never produces that combination) - its CashAmount is
    negated. Idempotent: a second run finds nothing left to migrate.
    Never touches InvestmentRawID/InvestmentTransactionID.
    """
    out = df.copy()
    is_seligson = out["Broker"].map(normalise_text) == "SELIGSON"
    cash_amount = pd.to_numeric(out["CashAmount"], errors="coerce").fillna(0.0)
    transaction_type = out[type_column].map(normalise_text)

    needs_migration = is_seligson & (
        ((transaction_type == buy_value) & (cash_amount > 0))
        | ((transaction_type == sell_value) & (cash_amount < 0))
    )

    out.loc[needs_migration, "CashAmount"] = -cash_amount.loc[needs_migration]

    return out, int(needs_migration.sum())


def migrate_workbook(workbook_path: Path, *, dry_run: bool = False) -> dict[str, object]:
    raw = read_sheet(workbook_path, INVESTMENT_RAW_SHEET, INVESTMENT_RAW_COLUMNS)
    transactions = read_sheet(workbook_path, INVESTMENT_TRANSACTIONS_SHEET, INVESTMENT_TRANSACTIONS_COLUMNS)
    import_log = read_sheet(workbook_path, INVESTMENT_IMPORT_LOG_SHEET, INVESTMENT_IMPORT_LOG_COLUMNS)

    migrated_raw, raw_count = _migrate(raw, "TransactionTypeRaw", "MERKINTÄ", "LUNASTUS")
    migrated_transactions, transactions_count = _migrate(transactions, "TransactionType", "BUY", "SELL")

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
        description="One-time migration: Seligson rows had CashAmount stored with the export's own "
                    "fund-perspective sign instead of the investor's cash-flow sign (fixed in "
                    "parsers/seligson.py). Flips already-imported rows' CashAmount in place. "
                    "InvestmentRawID/InvestmentTransactionID are never touched, so future re-imports "
                    "of the same source files still dedupe correctly against these rows."
    )
    parser.add_argument("--workbook", default=str(DEFAULT_WORKBOOK))
    parser.add_argument("--dry-run", action="store_true")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()
    workbook_path = Path(args.workbook).expanduser().resolve()

    stats = migrate_workbook(workbook_path, dry_run=args.dry_run)

    print("Seligson cash amount sign migration dry run complete." if args.dry_run else "Seligson cash amount sign migration complete.")
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
