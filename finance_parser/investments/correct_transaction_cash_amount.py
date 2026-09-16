from __future__ import annotations

import argparse
from pathlib import Path

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


def correct_cash_amount(
    workbook_path: Path,
    investment_transaction_id: str,
    new_cash_amount: float,
    *,
    dry_run: bool = False,
) -> dict[str, object]:
    """
    Manually correct a single already-imported transaction's CashAmount in
    InvestmentTransactions - for real events a broker's export recorded with
    no cost (CashAmount=0/blank) because the real cost basis was never in
    the export at all (e.g. a fund-merger transfer), where the user has
    since worked out or remembered the real value. Only InvestmentTransactions
    is touched (the classified/processed layer) - RawInvestmentTransactions
    is left as the true, unmodified historical record of what the broker
    export actually said, matching this project's raw-vs-classified
    separation. InvestmentTransactionID/InvestmentRawID are never touched.
    """
    transactions = read_sheet(workbook_path, INVESTMENT_TRANSACTIONS_SHEET, INVESTMENT_TRANSACTIONS_COLUMNS)

    match = transactions["InvestmentTransactionID"] == investment_transaction_id
    if not match.any():
        raise ValueError(f"No InvestmentTransactions row found with InvestmentTransactionID={investment_transaction_id!r}")

    old_value = transactions.loc[match, "CashAmount"].iloc[0]
    transactions.loc[match, "CashAmount"] = new_cash_amount

    stats: dict[str, object] = {
        "investment_transaction_id": investment_transaction_id,
        "old_cash_amount": old_value,
        "new_cash_amount": new_cash_amount,
        "rows_matched": int(match.sum()),
    }

    if dry_run:
        return stats

    raw = read_sheet(workbook_path, INVESTMENT_RAW_SHEET, INVESTMENT_RAW_COLUMNS)
    import_log = read_sheet(workbook_path, INVESTMENT_IMPORT_LOG_SHEET, INVESTMENT_IMPORT_LOG_COLUMNS)

    extra_sheets = {}
    for sheet_name in BROKER_RAW_SHEETS:
        sheet_df = read_dynamic_sheet(workbook_path, sheet_name)
        if len(sheet_df) > 0:
            extra_sheets[sheet_name] = sheet_df

    write_investment_output_workbook(
        workbook_path,
        raw[INVESTMENT_RAW_COLUMNS],
        transactions[INVESTMENT_TRANSACTIONS_COLUMNS],
        import_log[INVESTMENT_IMPORT_LOG_COLUMNS],
        extra_sheets=extra_sheets,
    )

    return stats


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Manually correct one already-imported transaction's CashAmount, by "
                    "InvestmentTransactionID - for real events a broker's export recorded with no "
                    "cost basis, where the real value has since been worked out or remembered. "
                    "Only InvestmentTransactions is touched; the raw sheet stays the true original."
    )
    parser.add_argument("--workbook", default=str(DEFAULT_WORKBOOK))
    parser.add_argument("--transaction-id", required=True)
    parser.add_argument("--cash-amount", required=True, type=float)
    parser.add_argument("--dry-run", action="store_true")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()
    workbook_path = Path(args.workbook).expanduser().resolve()

    stats = correct_cash_amount(
        workbook_path,
        args.transaction_id,
        args.cash_amount,
        dry_run=args.dry_run,
    )

    print("Cash amount correction dry run complete." if args.dry_run else "Cash amount correction complete.")
    print(f"InvestmentTransactionID: {stats['investment_transaction_id']}")
    print(f"Old CashAmount:          {stats['old_cash_amount']}")
    print(f"New CashAmount:          {stats['new_cash_amount']}")

    if args.dry_run:
        print()
        print("Dry run only: workbook was not modified.")
    else:
        print()
        print(f"Output workbook: {workbook_path}")


if __name__ == "__main__":
    main()
