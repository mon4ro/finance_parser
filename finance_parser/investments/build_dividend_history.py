from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from finance_parser.common import normalise_header, normalise_text
from finance_parser.investments.enrich_dividend_local_currency import (
    enrich_dividend_history,
    load_op_dividend_local_currency_details,
)
from finance_parser.utilities.fresh_workbook_writer import (
    records_to_sheet_values,
    replace_with_fresh_workbook,
    write_fresh_workbook,
)

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_INVESTMENTS_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "ParsedInvestments.xlsx"
DEFAULT_OUTPUT_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "DividendHistory.xlsx"
DEFAULT_BUDGETING_WORKBOOK = PROJECT_ROOT / "output" / "budgeting" / "ParsedTransactions.xlsx"

DIVIDEND_HISTORY_SHEET = "DividendHistory"
DIVIDEND_HISTORY_COLUMNS = [
    "TradeDate", "Year", "Month", "Broker", "Portfolio", "PortfolioOwner",
    "PortfolioType", "NormalizedInstrument", "GrossDividendEUR",
    "TaxWithheldEUR", "NetDividendEUR", "TaxDataAvailable",
    "LocalCurrency", "GrossDividendLocal", "TaxWithheldLocal", "ExchangeRate",
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
            "TaxWithheldEUR", "NetDividendEUR", "TaxDataAvailable",
        ])

    # Real gap found and reported (not silently fixed - see below): a
    # broker whose export never includes a withholding-tax transaction row
    # at all (confirmed against real data: EVLI, an employee share-purchase
    # plan account) makes TaxWithheldEUR sum to exactly 0 for every
    # one of its dividends, which looks identical to "genuinely tax-exempt"
    # even when real tax WAS withheld (just not visible in this data
    # source - possibly settled via payroll instead). Flagging per broker
    # (does this broker's export format ever carry a withholding row for
    # ANY dividend) rather than guessing a withholding amount we have no
    # real source for - this project never fabricates a number it can't
    # derive from real data.
    brokers_with_tax_data = set(
        relevant.loc[relevant["TransactionType"].isin(WITHHOLDING_TAX_TYPES), "Broker"].unique()
    )

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
    grouped["TaxDataAvailable"] = grouped["Broker"].isin(brokers_with_tax_data)

    return grouped.sort_values(group_cols).reset_index(drop=True)


def build_dividend_history(
    investments_workbook: Path,
    budgeting_workbook: Path | None = None,
) -> tuple[pd.DataFrame, dict[str, object]]:
    events = load_dividend_events(investments_workbook)

    if len(events) == 0:
        result = pd.DataFrame(columns=[c for c in DIVIDEND_HISTORY_COLUMNS if c not in ("LocalCurrency", "GrossDividendLocal", "TaxWithheldLocal", "ExchangeRate")])
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
            "TaxDataAvailable": events["TaxDataAvailable"],
        })

    # Best-effort enrichment, not a hard dependency: only Telia has ever
    # needed this (the only real non-EUR-native dividend payer found in the
    # whole portfolio), and a missing/absent budgeting workbook just means
    # the four local-currency columns stay blank - never an error.
    local_currency_matched = 0
    if budgeting_workbook is not None and len(result) > 0:
        local_details = load_op_dividend_local_currency_details(budgeting_workbook)
        result, local_currency_matched = enrich_dividend_history(result, local_details)
    else:
        for col in ["LocalCurrency", "GrossDividendLocal", "TaxWithheldLocal", "ExchangeRate"]:
            result[col] = ""

    brokers_without_tax_data = sorted(
        set(result.loc[~result["TaxDataAvailable"].astype(bool), "Broker"])
    ) if len(result) else []

    stats = {
        "events": len(result),
        "total_gross_eur": round(float(result["GrossDividendEUR"].sum()), 2) if len(result) else 0.0,
        "total_tax_eur": round(float(result["TaxWithheldEUR"].sum()), 2) if len(result) else 0.0,
        "total_net_eur": round(float(result["NetDividendEUR"].sum()), 2) if len(result) else 0.0,
        "local_currency_events_matched": local_currency_matched,
        "brokers_without_tax_data": brokers_without_tax_data,
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
    parser.add_argument(
        "--budgeting-workbook",
        default=str(DEFAULT_BUDGETING_WORKBOOK),
        help="Optional: source of local-currency dividend detail (e.g. Telia's SEK amount/exchange rate) "
             "from OP's own dividend notices. Best-effort - missing or non-matching rows are left blank.",
    )
    parser.add_argument(
        "--no-local-currency",
        action="store_true",
        help="Skip the local-currency enrichment entirely, even if --budgeting-workbook exists.",
    )
    parser.add_argument("--dry-run", action="store_true")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()
    investments_workbook = Path(args.investments_workbook).expanduser().resolve()
    output_workbook = Path(args.output_workbook).expanduser().resolve()
    budgeting_workbook = None if args.no_local_currency else Path(args.budgeting_workbook).expanduser().resolve()

    dividend_history, stats = build_dividend_history(investments_workbook, budgeting_workbook)

    print("Dividend history rebuild complete." if not args.dry_run else "Dry run complete.")
    print(f"Dividend events:       {stats['events']}")
    print(f"Total gross dividends: {stats['total_gross_eur']}")
    print(f"Total tax withheld:    {stats['total_tax_eur']}")
    print(f"Total net dividends:   {stats['total_net_eur']}")
    print(f"Local-currency detail matched: {stats['local_currency_events_matched']}")

    if stats["brokers_without_tax_data"]:
        print()
        print(
            f"WARNING: {', '.join(stats['brokers_without_tax_data'])} export(s) never include a "
            "withholding-tax transaction row for any dividend - GrossDividendEUR == NetDividendEUR "
            "for these is a data-availability gap, NOT a claim that no tax was withheld. Real tax may"
        )
        print("  have been withheld outside this data source (e.g. via payroll) and isn't recoverable from here.")

    if args.dry_run:
        print()
        print("Dry run only: workbook was not modified.")
        return

    write_dividend_history_workbook(output_workbook, dividend_history)
    print()
    print(f"Output workbook: {output_workbook}")


if __name__ == "__main__":
    main()
