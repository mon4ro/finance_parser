from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from finance_parser.common import normalise_header, normalise_text
from finance_parser.utilities.fresh_workbook_writer import (
    records_to_sheet_values,
    replace_with_fresh_workbook,
    write_fresh_workbook,
)

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_INVESTMENTS_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "ParsedInvestments.xlsx"
DEFAULT_OUTPUT_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "DividendHistory.xlsx"

DIVIDEND_HISTORY_SHEET = "DividendHistory"
DIVIDEND_HISTORY_COLUMNS = [
    "TradeDate", "Year", "Month", "Broker", "Portfolio", "PortfolioOwner",
    "PortfolioType", "NormalizedInstrument", "GrossDividendEUR",
    "TaxWithheldEUR", "NetDividendEUR",
]

# TAX (OP) and ENNAKKOPIDÄTYS (Nordnet) are the two real TransactionType
# values seen for dividend withholding tax - different brokers, same
# meaning. Not every DIVIDEND row has a matching withholding row (e.g. EVLI
# has none at all in the real data) - grouping and summing handles that
# naturally rather than requiring a strict 1:1 pairing.
GROSS_DIVIDEND_TYPES = {"DIVIDEND"}
WITHHOLDING_TAX_TYPES = {"TAX", "ENNAKKOPIDÄTYS"}


def load_dividend_events(investments_workbook: Path) -> pd.DataFrame:
    """
    Return one row per real dividend event: (Broker, Portfolio,
    NormalizedInstrument, TradeDate) with that date's gross dividend and
    withheld tax summed together - grouped, not matched 1:1, since a
    dividend and its withholding aren't guaranteed to appear as exactly one
    row each.
    """
    df = pd.read_excel(investments_workbook, sheet_name="InvestmentTransactions", dtype=object, engine="openpyxl")
    df.columns = [normalise_header(c) for c in df.columns]

    df["Broker"] = df["Broker"].map(normalise_text)
    df["Portfolio"] = df["Portfolio"].map(normalise_text)
    df["PortfolioOwner"] = df["PortfolioOwner"].map(normalise_text) if "PortfolioOwner" in df.columns else ""
    df["PortfolioType"] = df["PortfolioType"].map(normalise_text) if "PortfolioType" in df.columns else ""
    df["NormalizedInstrument"] = df["NormalizedInstrument"].map(normalise_text)
    df["TransactionType"] = df["TransactionType"].map(normalise_text)
    df["TradeDate"] = pd.to_datetime(df["TradeDate"], errors="coerce")
    df["CashAmount"] = pd.to_numeric(df["CashAmount"], errors="coerce").fillna(0.0)

    relevant = df[df["TransactionType"].isin(GROSS_DIVIDEND_TYPES | WITHHOLDING_TAX_TYPES)].copy()
    relevant = relevant.dropna(subset=["TradeDate"])
    if len(relevant) == 0:
        return pd.DataFrame(columns=[
            "Broker", "Portfolio", "PortfolioOwner", "PortfolioType",
            "NormalizedInstrument", "TradeDate", "GrossDividendEUR",
            "TaxWithheldEUR", "NetDividendEUR",
        ])

    relevant["GrossDividendEUR"] = relevant["CashAmount"].where(relevant["TransactionType"].isin(GROSS_DIVIDEND_TYPES), 0.0)
    # Withholding tax is stored as a negative CashAmount (money leaving) -
    # negate for a positive "amount withheld" figure.
    relevant["TaxWithheldEUR"] = (-relevant["CashAmount"]).where(relevant["TransactionType"].isin(WITHHOLDING_TAX_TYPES), 0.0)

    group_cols = ["Broker", "Portfolio", "NormalizedInstrument", "TradeDate"]
    grouped = relevant.groupby(group_cols, as_index=False).agg(
        PortfolioOwner=("PortfolioOwner", "first"),
        PortfolioType=("PortfolioType", "first"),
        GrossDividendEUR=("GrossDividendEUR", "sum"),
        TaxWithheldEUR=("TaxWithheldEUR", "sum"),
    )
    grouped["NetDividendEUR"] = grouped["GrossDividendEUR"] - grouped["TaxWithheldEUR"]

    return grouped.sort_values(group_cols).reset_index(drop=True)


def build_dividend_history(investments_workbook: Path) -> tuple[pd.DataFrame, dict[str, object]]:
    events = load_dividend_events(investments_workbook)

    if len(events) == 0:
        result = pd.DataFrame(columns=DIVIDEND_HISTORY_COLUMNS)
    else:
        result = pd.DataFrame({
            "TradeDate": events["TradeDate"].dt.strftime("%Y-%m-%d"),
            "Year": events["TradeDate"].dt.year,
            "Month": events["TradeDate"].dt.month,
            "Broker": events["Broker"],
            "Portfolio": events["Portfolio"],
            "PortfolioOwner": events["PortfolioOwner"],
            "PortfolioType": events["PortfolioType"],
            "NormalizedInstrument": events["NormalizedInstrument"],
            "GrossDividendEUR": events["GrossDividendEUR"].round(2),
            "TaxWithheldEUR": events["TaxWithheldEUR"].round(2),
            "NetDividendEUR": events["NetDividendEUR"].round(2),
        })

    stats = {
        "events": len(result),
        "total_gross_eur": round(float(result["GrossDividendEUR"].sum()), 2) if len(result) else 0.0,
        "total_tax_eur": round(float(result["TaxWithheldEUR"].sum()), 2) if len(result) else 0.0,
        "total_net_eur": round(float(result["NetDividendEUR"].sum()), 2) if len(result) else 0.0,
    }

    return result[DIVIDEND_HISTORY_COLUMNS], stats


def write_dividend_history_workbook(output_workbook: Path, dividend_history: pd.DataFrame) -> None:
    headers = list(DIVIDEND_HISTORY_COLUMNS)
    records = dividend_history.to_dict("records")
    sheets = {DIVIDEND_HISTORY_SHEET: records_to_sheet_values(headers, records)}

    if output_workbook.exists():
        replace_with_fresh_workbook(
            output_workbook,
            sheets,
            backup_label="before_dividend_history_rebuild",
            basic_formatting=True,
            excel_tables=False,
        )
    else:
        write_fresh_workbook(output_workbook, sheets, basic_formatting=True, excel_tables=False)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Recompute DividendHistory.xlsx from scratch: one row per real dividend event "
                    "(gross, tax withheld, net) from InvestmentTransactions. Always a full rebuild, "
                    "never incremental - same reasoning as build_portfolio_positions.py."
    )
    parser.add_argument("--investments-workbook", default=str(DEFAULT_INVESTMENTS_WORKBOOK))
    parser.add_argument("--output-workbook", default=str(DEFAULT_OUTPUT_WORKBOOK))
    parser.add_argument("--dry-run", action="store_true")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()
    investments_workbook = Path(args.investments_workbook).expanduser().resolve()
    output_workbook = Path(args.output_workbook).expanduser().resolve()

    dividend_history, stats = build_dividend_history(investments_workbook)

    print("Dividend history rebuild complete." if not args.dry_run else "Dry run complete.")
    print(f"Dividend events:       {stats['events']}")
    print(f"Total gross dividends: {stats['total_gross_eur']}")
    print(f"Total tax withheld:    {stats['total_tax_eur']}")
    print(f"Total net dividends:   {stats['total_net_eur']}")

    if args.dry_run:
        print()
        print("Dry run only: workbook was not modified.")
        return

    write_dividend_history_workbook(output_workbook, dividend_history)
    print()
    print(f"Output workbook: {output_workbook}")


if __name__ == "__main__":
    main()
