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
    "NordeaRawExport",
    "NordnetRawExport",
    "OPInvestmentRawExport",
    "SeligsonRawExport",
]


def correct_trade_date(
    workbook_path: Path,
    investment_transaction_id: str,
    new_trade_date: str,
    *,
    new_settlement_date: str | None = None,
    dry_run: bool = False,
) -> dict[str, object]:
    """
    Manually correct a single already-imported transaction's TradeDate/
    SettlementDate in InvestmentTransactions - for a real event whose
    broker-reported date is legally correct for one purpose (a tax cost-
    basis date) but wrong for another (when the holding actually started
    counting toward net worth). Real case this was built for: OP records an
    inherited security's "OSTO" (purchase) transaction dated at the
    deceased's date of death (the correct Finnish tax cost-basis/holding-
    period start), not when the security actually reached the heir's
    account - which, per the user, can be years later depending on how long
    probate (kuolinpesä) takes to settle. Using the death date for net-worth
    tracking would count wealth the user did not yet have any access to or
    control over.

    Only InvestmentTransactions is touched (the layer PortfolioPositions.py
    and everything downstream actually reads dates from) - the raw sheet is
    deliberately left as the true, unmodified original record, since the
    original tax-basis date remains correct and needed for real future
    capital-gains-tax calculations. InvestmentTransactionID/InvestmentRawID
    are never touched.
    """
    transactions = read_sheet(workbook_path, INVESTMENT_TRANSACTIONS_SHEET, INVESTMENT_TRANSACTIONS_COLUMNS)

    match = transactions["InvestmentTransactionID"] == investment_transaction_id
    if not match.any():
        raise ValueError(f"No InvestmentTransactions row found with InvestmentTransactionID={investment_transaction_id!r}")

    settlement_date = new_settlement_date if new_settlement_date is not None else new_trade_date

    old_trade_date = transactions.loc[match, "TradeDate"].iloc[0]
    old_settlement_date = transactions.loc[match, "SettlementDate"].iloc[0]
    transactions.loc[match, "TradeDate"] = new_trade_date
    transactions.loc[match, "SettlementDate"] = settlement_date

    stats: dict[str, object] = {
        "investment_transaction_id": investment_transaction_id,
        "old_trade_date": old_trade_date,
        "new_trade_date": new_trade_date,
        "old_settlement_date": old_settlement_date,
        "new_settlement_date": settlement_date,
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
        description="Manually correct one already-imported transaction's TradeDate/SettlementDate, by "
                    "InvestmentTransactionID - for a real event whose broker-reported date is correct for "
                    "one purpose (e.g. a tax cost-basis date) but wrong for net-worth tracking. Only "
                    "InvestmentTransactions is touched; the raw sheet stays the true original."
    )
    parser.add_argument("--workbook", default=str(DEFAULT_WORKBOOK))
    parser.add_argument("--transaction-id", required=True)
    parser.add_argument("--trade-date", required=True, help="New date, YYYY-MM-DD.")
    parser.add_argument("--settlement-date", default=None, help="Defaults to --trade-date if omitted.")
    parser.add_argument("--dry-run", action="store_true")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()
    workbook_path = Path(args.workbook).expanduser().resolve()

    stats = correct_trade_date(
        workbook_path,
        args.transaction_id,
        args.trade_date,
        new_settlement_date=args.settlement_date,
        dry_run=args.dry_run,
    )

    print("Trade date correction dry run complete." if args.dry_run else "Trade date correction complete.")
    print(f"InvestmentTransactionID: {stats['investment_transaction_id']}")
    print(f"Old TradeDate:           {stats['old_trade_date']} -> New: {stats['new_trade_date']}")
    print(f"Old SettlementDate:      {stats['old_settlement_date']} -> New: {stats['new_settlement_date']}")

    if args.dry_run:
        print()
        print("Dry run only: workbook was not modified.")
    else:
        print()
        print(f"Output workbook: {workbook_path}")


if __name__ == "__main__":
    main()
