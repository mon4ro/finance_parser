from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path

import pandas as pd

from finance_parser.common import normalise_header, normalise_text
from finance_parser.investments.build_dividend_history import load_dividend_events
from finance_parser.utilities.fresh_workbook_writer import (
    read_workbook_values_only,
    records_to_sheet_values,
    replace_with_fresh_workbook,
    write_fresh_workbook,
)

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_POSITIONS_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "PortfolioPositions.xlsx"
DEFAULT_PRICES_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "InstrumentPrices.xlsx"
DEFAULT_FX_RATES_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "FXRates.xlsx"
DEFAULT_INVESTMENTS_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "ParsedInvestments.xlsx"
DEFAULT_OUTPUT_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "MonthlyPositionValue.xlsx"

MONTHLY_VALUE_SHEET = "MonthlyPositionValue"
MONTHLY_VALUE_COLUMNS = [
    "MonthEnd", "Year", "Month", "Broker", "Portfolio", "PortfolioOwner",
    "NormalizedInstrument", "CumulativeQuantity", "PriceLocal", "InstrumentCurrency",
    "PriceDate", "FXRate", "FXDate", "MarketValueEUR",
    "CumulativeNetInvested", "UnrealizedGainEUR", "UnrealizedGainPercent",
    "DividendGrossEUR", "DividendTaxEUR", "DividendNetEUR",
    "CumulativeDividendGrossEUR", "CumulativeDividendNetEUR",
]

DIVIDEND_EVENT_COLUMNS = ["Broker", "Portfolio", "NormalizedInstrument", "TradeDate", "GrossDividendEUR", "TaxWithheldEUR", "NetDividendEUR"]

BASE_CURRENCY = "EUR"
QUANTITY_EPSILON = 1e-6


def load_portfolio_positions(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="PortfolioPositions", dtype=object, engine="openpyxl")
    df.columns = [normalise_header(c) for c in df.columns]

    df["Broker"] = df["Broker"].map(normalise_text)
    df["Portfolio"] = df["Portfolio"].map(normalise_text)
    df["NormalizedInstrument"] = df["NormalizedInstrument"].map(normalise_text)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["CumulativeQuantity"] = pd.to_numeric(df["CumulativeQuantity"], errors="coerce")
    df["CumulativeNetInvested"] = pd.to_numeric(df["CumulativeNetInvested"], errors="coerce")

    if "PortfolioOwner" not in df.columns:
        df["PortfolioOwner"] = ""
    df["PortfolioOwner"] = df["PortfolioOwner"].map(normalise_text)

    return df.dropna(subset=["Date"]).sort_values("Date")


def load_instrument_prices(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="InstrumentPrices", dtype=object, engine="openpyxl")
    df.columns = [normalise_header(c) for c in df.columns]

    df["NormalizedInstrument"] = df["NormalizedInstrument"].map(normalise_text)
    df["Currency"] = df["Currency"].map(normalise_text).str.upper()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Close"] = pd.to_numeric(df["Close"], errors="coerce")

    return df.dropna(subset=["Date"]).sort_values("Date")


def load_fx_rates(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=["Currency", "Date", "Rate"])

    df = pd.read_excel(path, sheet_name="FXRates", dtype=object, engine="openpyxl")
    df.columns = [normalise_header(c) for c in df.columns]

    df["Currency"] = df["Currency"].map(normalise_text).str.upper()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Rate"] = pd.to_numeric(df["Rate"], errors="coerce")

    return df.dropna(subset=["Date"]).sort_values("Date")


def last_completed_month_end(today: date) -> pd.Timestamp:
    first_of_this_month = pd.Timestamp(today).replace(day=1)
    return first_of_this_month - pd.Timedelta(days=1)


def month_end_dates(first_event_date: pd.Timestamp, last_month_end: pd.Timestamp) -> pd.DatetimeIndex:
    if first_event_date > last_month_end:
        return pd.DatetimeIndex([])
    return pd.date_range(start=first_event_date, end=last_month_end, freq="ME")


def _monthly_dividend_table(group_events: pd.DataFrame, months: pd.DatetimeIndex) -> pd.DataFrame:
    """
    For one (Broker, Portfolio, NormalizedInstrument) group's dividend
    events, return one row per candidate month with that month's gross/tax/
    net dividend sum and the running cumulative gross/net total up to and
    including that month - a lifetime running total for the group, not
    reset by a full sell-then-later-rebuy (same convention as
    CumulativeNetInvested).
    """
    result = pd.DataFrame({"MonthEnd": months})

    if len(group_events) == 0:
        result["DividendGrossEUR"] = 0.0
        result["DividendTaxEUR"] = 0.0
        result["DividendNetEUR"] = 0.0
        result["CumulativeDividendGrossEUR"] = 0.0
        result["CumulativeDividendNetEUR"] = 0.0
        return result

    events = group_events.copy()
    events["_MonthPeriod"] = events["TradeDate"].dt.to_period("M")
    monthly = events.groupby("_MonthPeriod")[["GrossDividendEUR", "TaxWithheldEUR", "NetDividendEUR"]].sum()
    monthly = monthly.rename(columns={
        "GrossDividendEUR": "DividendGrossEUR",
        "TaxWithheldEUR": "DividendTaxEUR",
        "NetDividendEUR": "DividendNetEUR",
    })

    result["_MonthPeriod"] = result["MonthEnd"].dt.to_period("M")
    result = result.merge(monthly, how="left", left_on="_MonthPeriod", right_index=True)
    result = result.drop(columns=["_MonthPeriod"])
    for col in ["DividendGrossEUR", "DividendTaxEUR", "DividendNetEUR"]:
        result[col] = result[col].fillna(0.0)

    result["CumulativeDividendGrossEUR"] = result["DividendGrossEUR"].cumsum()
    result["CumulativeDividendNetEUR"] = result["DividendNetEUR"].cumsum()

    return result


def build_monthly_values(
    positions_workbook: Path,
    prices_workbook: Path,
    fx_rates_workbook: Path,
    investments_workbook: Path | None = None,
    *,
    as_of: date | None = None,
) -> tuple[pd.DataFrame, dict[str, object]]:
    positions = load_portfolio_positions(positions_workbook)
    prices = load_instrument_prices(prices_workbook)
    fx_rates = load_fx_rates(fx_rates_workbook)
    if investments_workbook is not None and investments_workbook.exists():
        dividend_events = load_dividend_events(investments_workbook)
    else:
        dividend_events = pd.DataFrame(columns=DIVIDEND_EVENT_COLUMNS)

    last_month_end = last_completed_month_end(as_of or date.today())

    group_cols = ["Broker", "Portfolio", "NormalizedInstrument"]
    rows: list[dict[str, object]] = []
    missing_price: set[tuple[str, str]] = set()
    missing_fx: set[tuple[str, str]] = set()

    for (broker, portfolio, instrument), group in positions.groupby(group_cols, sort=False):
        group = group.sort_values("Date")
        owner = normalise_text(group["PortfolioOwner"].iloc[0]) if "PortfolioOwner" in group.columns else ""
        months = month_end_dates(group["Date"].iloc[0], last_month_end)
        if len(months) == 0:
            continue

        group_dividends = dividend_events[
            (dividend_events["Broker"] == broker)
            & (dividend_events["Portfolio"] == portfolio)
            & (dividend_events["NormalizedInstrument"] == instrument)
        ]
        dividend_by_month = _monthly_dividend_table(group_dividends, months)

        candidates = pd.DataFrame({"MonthEnd": months})
        as_of_position = pd.merge_asof(
            candidates,
            group[["Date", "CumulativeQuantity", "CumulativeNetInvested"]],
            left_on="MonthEnd",
            right_on="Date",
            direction="backward",
        )
        as_of_position = as_of_position[as_of_position["CumulativeQuantity"] > QUANTITY_EPSILON]
        if len(as_of_position) == 0:
            continue

        instrument_prices = prices[prices["NormalizedInstrument"] == instrument][["Date", "Close", "Currency"]]
        with_price = pd.merge_asof(
            as_of_position.sort_values("MonthEnd"),
            instrument_prices.sort_values("Date"),
            left_on="MonthEnd",
            right_on="Date",
            direction="backward",
            suffixes=("", "_price"),
        )
        with_price = with_price.merge(dividend_by_month, on="MonthEnd", how="left")

        for _, r in with_price.iterrows():
            month_end = r["MonthEnd"]

            if pd.isna(r["Close"]):
                missing_price.add((instrument, month_end.strftime("%Y-%m")))
                continue

            currency = r["Currency"] or BASE_CURRENCY

            if currency == BASE_CURRENCY:
                fx_rate = 1.0
                fx_date = month_end
            else:
                currency_rates = fx_rates[fx_rates["Currency"] == currency][["Date", "Rate"]].sort_values("Date")
                fx_match = pd.merge_asof(
                    pd.DataFrame({"MonthEnd": [month_end]}),
                    currency_rates,
                    left_on="MonthEnd",
                    right_on="Date",
                    direction="backward",
                )
                if pd.isna(fx_match["Rate"].iloc[0]):
                    missing_fx.add((currency, month_end.strftime("%Y-%m")))
                    continue
                fx_rate = float(fx_match["Rate"].iloc[0])
                fx_date = fx_match["Date"].iloc[0]

            quantity = float(r["CumulativeQuantity"])
            price_local = float(r["Close"])
            market_value_eur = round(quantity * price_local * fx_rate, 2)
            net_invested = round(float(r["CumulativeNetInvested"]), 2)
            unrealized_gain_eur = round(market_value_eur + net_invested, 2)

            rows.append({
                "MonthEnd": month_end.strftime("%Y-%m-%d"),
                "Year": month_end.year,
                "Month": month_end.month,
                "Broker": broker,
                "Portfolio": portfolio,
                "PortfolioOwner": owner,
                "NormalizedInstrument": instrument,
                "CumulativeQuantity": quantity,
                "PriceLocal": price_local,
                "InstrumentCurrency": currency,
                "PriceDate": r["Date_price"].strftime("%Y-%m-%d") if "Date_price" in r and pd.notna(r["Date_price"]) else r["Date"].strftime("%Y-%m-%d"),
                "FXRate": fx_rate,
                "FXDate": fx_date.strftime("%Y-%m-%d") if pd.notna(fx_date) else "",
                "MarketValueEUR": market_value_eur,
                "CumulativeNetInvested": net_invested,
                # net_invested is a cash-flow figure (negative = money paid
                # out to buy), not a positive cost basis - so gain is value
                # PLUS net_invested, not minus. Confirmed real bug: this used
                # to be `market_value_eur - net_invested`, which for a
                # negative net_invested added the two together instead of
                # subtracting, roughly doubling every gain figure. Both
                # inputs are already rounded above so this always sums
                # exactly to the two stored columns, not off by a cent from
                # rounding market_value_eur/net_invested independently.
                "UnrealizedGainEUR": unrealized_gain_eur,
                # Blank (not 0 or an error) when net_invested is 0 - a
                # fully-sold-then-recovered-exactly position, or a genuinely
                # unknown-cost-basis default - dividing by zero has no
                # sensible percentage to show.
                "UnrealizedGainPercent": (
                    round(unrealized_gain_eur / abs(net_invested) * 100, 2) if net_invested != 0 else ""
                ),
                "DividendGrossEUR": round(float(r["DividendGrossEUR"]), 2),
                "DividendTaxEUR": round(float(r["DividendTaxEUR"]), 2),
                "DividendNetEUR": round(float(r["DividendNetEUR"]), 2),
                "CumulativeDividendGrossEUR": round(float(r["CumulativeDividendGrossEUR"]), 2),
                "CumulativeDividendNetEUR": round(float(r["CumulativeDividendNetEUR"]), 2),
            })

    result = pd.DataFrame(rows, columns=MONTHLY_VALUE_COLUMNS)
    if len(result) > 0:
        result = result.sort_values(["NormalizedInstrument", "Broker", "Portfolio", "MonthEnd"]).reset_index(drop=True)

    stats = {
        "rows_produced": len(result),
        "last_month_end": last_month_end.strftime("%Y-%m-%d"),
        "missing_price_instrument_months": sorted(missing_price),
        "missing_fx_currency_months": sorted(missing_fx),
    }

    return result, stats


def write_monthly_value_workbook(output_workbook: Path, monthly_values: pd.DataFrame) -> None:
    headers = list(MONTHLY_VALUE_COLUMNS)
    records = monthly_values.to_dict("records")
    sheets = {MONTHLY_VALUE_SHEET: records_to_sheet_values(headers, records)}

    if output_workbook.exists():
        replace_with_fresh_workbook(
            output_workbook,
            sheets,
            backup_label="before_monthly_value_rebuild",
            basic_formatting=True,
            excel_tables=False,
        )
    else:
        write_fresh_workbook(output_workbook, sheets, basic_formatting=True, excel_tables=False)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Recompute MonthlyPositionValue.xlsx from scratch: month-end EUR market value per "
                    "(Broker, Portfolio, NormalizedInstrument), from PortfolioPositions x InstrumentPrices x FXRates. "
                    "Always a full rebuild, never incremental - same reasoning as build_portfolio_positions.py."
    )
    parser.add_argument("--positions-workbook", default=str(DEFAULT_POSITIONS_WORKBOOK))
    parser.add_argument("--prices-workbook", default=str(DEFAULT_PRICES_WORKBOOK))
    parser.add_argument("--fx-rates-workbook", default=str(DEFAULT_FX_RATES_WORKBOOK))
    parser.add_argument("--investments-workbook", default=str(DEFAULT_INVESTMENTS_WORKBOOK), help="Source of dividend/tax events.")
    parser.add_argument("--output-workbook", default=str(DEFAULT_OUTPUT_WORKBOOK))
    parser.add_argument("--dry-run", action="store_true")
    return parser


def _summarize_missing(pairs: list[tuple[str, str]]) -> list[tuple[str, int, str, str]]:
    """Group (key, month) pairs into (key, count, earliest_month, latest_month), sorted by count descending."""
    by_key: dict[str, list[str]] = {}
    for key, month in pairs:
        by_key.setdefault(key, []).append(month)

    summary = [(key, len(months), min(months), max(months)) for key, months in by_key.items()]
    return sorted(summary, key=lambda row: row[1], reverse=True)


def main() -> None:
    args = build_arg_parser().parse_args()

    positions_workbook = Path(args.positions_workbook).expanduser().resolve()
    prices_workbook = Path(args.prices_workbook).expanduser().resolve()
    fx_rates_workbook = Path(args.fx_rates_workbook).expanduser().resolve()
    investments_workbook = Path(args.investments_workbook).expanduser().resolve()
    output_workbook = Path(args.output_workbook).expanduser().resolve()

    monthly_values, stats = build_monthly_values(positions_workbook, prices_workbook, fx_rates_workbook, investments_workbook)

    print("Monthly position value rebuild complete." if not args.dry_run else "Dry run complete.")
    print(f"Last completed month-end used: {stats['last_month_end']}")
    print(f"Rows produced:                 {stats['rows_produced']}")

    if stats["missing_price_instrument_months"]:
        print()
        total = len(stats["missing_price_instrument_months"])
        print(f"Instrument-months with no price data available ({total} total), skipped, by instrument:")
        for instrument, count, earliest, latest in _summarize_missing(stats["missing_price_instrument_months"]):
            print(f"  - {instrument}: {count} month(s), {earliest} to {latest}")

    if stats["missing_fx_currency_months"]:
        print()
        total = len(stats["missing_fx_currency_months"])
        print(f"Currency-months with no FX rate available ({total} total), skipped, by currency:")
        for currency, count, earliest, latest in _summarize_missing(stats["missing_fx_currency_months"]):
            print(f"  - {currency}: {count} month(s), {earliest} to {latest}")

    if args.dry_run:
        print()
        print("Dry run only: workbook was not modified.")
        return

    write_monthly_value_workbook(output_workbook, monthly_values)
    print()
    print(f"Output workbook: {output_workbook}")


if __name__ == "__main__":
    main()
