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

_SELL_TYPES_RAW = {"SELL", "SELL OF PURCHASED SHARE"}
_SELL_TYPES_NORMALIZED = {"SELL"}


def _migrate(df: pd.DataFrame, type_column: str, sell_types: set[str]) -> tuple[pd.DataFrame, int]:
    """
    One-time fix for a real bug: every real EVLI Sell/Sell of purchased
    share row was recorded with CashAmount=0 because amount/amountLocal are
    blank in the export and the parser had no fallback - fixed in
    parsers/evli.py, which now estimates proceeds from Quantity x UnitPrice
    for a sell with no recorded amount. This corrects already-imported rows
    in place: an EVLI sell row with CashAmount=0 (and real Quantity/
    UnitPrice to estimate from) is by definition unmigrated - never touches
    InvestmentRawID/InvestmentTransactionID.
    """
    out = df.copy()
    is_evli = out["Broker"].map(normalise_text) == "EVLI"
    transaction_type = out[type_column].map(normalise_text).str.upper()
    cash_amount = pd.to_numeric(out["CashAmount"], errors="coerce").fillna(0.0)
    quantity = pd.to_numeric(out["Quantity"], errors="coerce").fillna(0.0)
    unit_price = pd.to_numeric(out["UnitPrice"], errors="coerce").fillna(0.0)

    needs_migration = (
        is_evli
        & transaction_type.isin(sell_types)
        & (cash_amount == 0)
        & (quantity != 0)
        & (unit_price != 0)
    )

    out.loc[needs_migration, "CashAmount"] = (-quantity * unit_price).round(2)

    return out, int(needs_migration.sum())


def migrate_workbook(workbook_path: Path, *, dry_run: bool = False) -> dict[str, object]:
    raw = read_sheet(workbook_path, INVESTMENT_RAW_SHEET, INVESTMENT_RAW_COLUMNS)
    transactions = read_sheet(workbook_path, INVESTMENT_TRANSACTIONS_SHEET, INVESTMENT_TRANSACTIONS_COLUMNS)
    import_log = read_sheet(workbook_path, INVESTMENT_IMPORT_LOG_SHEET, INVESTMENT_IMPORT_LOG_COLUMNS)

    migrated_raw, raw_count = _migrate(raw, "TransactionTypeRaw", _SELL_TYPES_RAW)
    migrated_transactions, transactions_count = _migrate(transactions, "TransactionType", _SELL_TYPES_NORMALIZED)

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
        description="One-time migration: EVLI Sell/Sell of purchased share rows had CashAmount=0 "
                    "because the export has no amount field for them (fixed in parsers/evli.py, which "
                    "now estimates proceeds from Quantity x UnitPrice). Corrects already-imported rows "
                    "in place. InvestmentRawID/InvestmentTransactionID are never touched."
    )
    parser.add_argument("--workbook", default=str(DEFAULT_WORKBOOK))
    parser.add_argument("--dry-run", action="store_true")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()
    workbook_path = Path(args.workbook).expanduser().resolve()

    stats = migrate_workbook(workbook_path, dry_run=args.dry_run)

    print("EVLI sell proceeds migration dry run complete." if args.dry_run else "EVLI sell proceeds migration complete.")
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
