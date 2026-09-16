from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path

import pandas as pd

from finance_parser.common import normalise_header, normalise_text
from finance_parser.investments.build_dividend_history import load_dividend_events
from finance_parser.settings import get_settings
from finance_parser.utilities.fresh_workbook_writer import (
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
    "MonthEnd", "Year", "Month", "Broker", "Portfolio", "PortfolioOwner", "PortfolioType",
    "NormalizedInstrument", "CumulativeQuantity", "PriceLocal", "InstrumentCurrency",
    "PriceDate", "FXRate", "FXDate", "MarketValueEUR",
    "CumulativeNetInvested", "UnrealizedGainEUR", "UnrealizedGainPercent",
    "DividendGrossEUR", "DividendTaxEUR", "DividendNetEUR",
    "CumulativeDividendGrossEUR", "CumulativeDividendNetEUR",
]

DIVIDEND_EVENT_COLUMNS = ["Broker", "Portfolio", "NormalizedInstrument", "TradeDate", "GrossDividendEUR", "TaxWithheldEUR", "NetDividendEUR"]

# Brokers where cash sitting unreinvested in the account (AOT/OST balance)
# is worth tracking as its own net-worth line. Deliberately excludes OP and
# Nordea (their cash lives in ordinary bank accounts already tracked on the
# budgeting side) and Seligson (direct fund purchase only - real data and
# the user both confirm no cash ever sits in a Seligson account).
CASH_TRACKED_BROKERS = {"NORDNET", "EVLI"}
CASH_INSTRUMENT_LABEL = "CASH"

# Real bugs found and fixed - two distinct EVLI TransactionTypeRaw values
# that never actually become real, spendable cash inside EVLI, even though
# their CashAmount is a genuine non-zero number:
#
# - Sell / Sell of purchased share: proceeds are wired straight to the
#   linked bank account at settlement, never held as EVLI cash. Confirmed
#   against real data on the budgeting side: two real EVLI withdrawals
#   ("Withdrawal from EAM account (ESSP)") landed in the bank within days of
#   a sale, each within a few percent of that sale's total proceeds (one
#   matched to the exact cent minus a flat fee).
# - Allocated: an employer-funded plan-cycle bonus that converts directly
#   into free "Matching" shares (which sit right next to it with CashAmount
#   0), not a deposit that ever lands in a liquid balance. Confirmed across
#   FOUR independent real plan cycles (2020/2021/2022/2023): each cycle's
#   Savings/Share-purchase pairs net to exactly 0.00, leaving only that
#   cycle's one-time Allocated amount (175/200/200/200) as a permanent
#   residual - never spent, never withdrawn, never swept into a purchase.
#
# Including either in the EVLI cash cumsum made the reconstructed balance
# grow forever with no corresponding outflow ever recorded (EVLI's own
# "Cash transferred" event type is unrelated - a much smaller, separate
# administrative cleanup of already-accumulated dividend float - see
# build_cash_balance_rows()'s docstring). The user caught the resulting
# figure as impossible against their own real, independently-tracked
# balance both times. Excluding these types entirely (as if they never
# touched EVLI's own float) is correct precisely because the money never
# did.
EVLI_NON_CASH_EVENT_TYPES = {"SELL", "SELL OF PURCHASED SHARE", "ALLOCATED"}

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

    if "PortfolioType" not in df.columns:
        df["PortfolioType"] = ""
    df["PortfolioType"] = df["PortfolioType"].map(normalise_text)

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


def load_cash_balance_source(investments_workbook: Path) -> pd.DataFrame:
    """
    RawInvestmentTransactions rows for CASH_TRACKED_BROKERS only. Reads the
    raw sheet (not InvestmentTransactions) specifically because it's the
    only one carrying Nordnet's real CashBalance ("Saldo") column -
    InvestmentTransactions deliberately doesn't include it (not a normal
    per-transaction field). PortfolioOwner/PortfolioType are NOT read from
    here - the raw sheet never gets apply_portfolio_ownership()'s backfill
    (only InvestmentTransactions does, see investment_parser.py), so they're
    looked up fresh from settings.yaml in build_cash_balance_rows() instead.

    Missing sheet (e.g. an investments workbook written before this raw
    sheet existed, or a test fixture that only sets up InvestmentTransactions)
    is not an error - just no cash-balance rows to add.
    """
    empty = pd.DataFrame(columns=["Broker", "Portfolio", "TradeDate", "CashAmount", "CashBalance", "TransactionTypeRaw"])
    try:
        df = pd.read_excel(investments_workbook, sheet_name="RawInvestmentTransactions", dtype=object, engine="openpyxl")
    except ValueError:
        return empty
    df.columns = [normalise_header(c) for c in df.columns]
    if "Broker" not in df.columns:
        return empty

    df["Broker"] = df["Broker"].map(normalise_text)
    df = df[df["Broker"].isin(CASH_TRACKED_BROKERS)].copy()

    df["Portfolio"] = df["Portfolio"].map(normalise_text)
    df["TradeDate"] = pd.to_datetime(df["TradeDate"], errors="coerce")
    df = df.dropna(subset=["TradeDate"])
    df["CashAmount"] = pd.to_numeric(df["CashAmount"], errors="coerce").fillna(0.0)
    df["CashBalance"] = pd.to_numeric(df["CashBalance"], errors="coerce")
    if "TransactionTypeRaw" not in df.columns:
        df["TransactionTypeRaw"] = ""
    df["TransactionTypeRaw"] = df["TransactionTypeRaw"].map(normalise_text).str.upper()

    return df


def build_cash_balance_rows(cash_source: pd.DataFrame, last_month_end: pd.Timestamp) -> list[dict[str, object]]:
    """
    One row per month-end per (Broker, Portfolio), NormalizedInstrument=
    "CASH" - uninvested cash sitting in a Nordnet or EVLI account (AOT/OST
    balance, not yet reinvested or withdrawn). Two different
    reconstructions, by necessity:

    - Nordnet's export carries its own running CashBalance ("Saldo") on
      every row - use it directly (merge_asof backward to each month-end,
      same pattern as instrument prices/FX rates elsewhere in this file).
    - EVLI's export never populates CashBalance (confirmed against real
      data: 0 of ~130 real rows have it) - reconstruct it as a chronological
      cumsum of every CashAmount the account has ever seen EXCEPT
      EVLI_NON_CASH_EVENT_TYPES (deposits, dividends, purchases, real
      withdrawals - everything but sell proceeds and the Allocated bonus,
      neither of which ever actually becomes spendable EVLI cash - see that
      constant's docstring for the real-data evidence).

    A month with zero cash balance is still a real, meaningful data point
    (fully withdrawn/reinvested) and is kept, not dropped - only a month
    before the account's first-ever event is excluded (mirrors
    month_end_dates()'s existing behavior for instrument positions).
    """
    if len(cash_source) == 0:
        return []

    settings = get_settings()
    rows: list[dict[str, object]] = []

    for (broker, portfolio), group in cash_source.groupby(["Broker", "Portfolio"], sort=False):
        group = group.sort_values("TradeDate")
        months = month_end_dates(group["TradeDate"].iloc[0], last_month_end)
        if len(months) == 0:
            continue

        if broker == "NORDNET":
            balance_source = (
                group.dropna(subset=["CashBalance"])[["TradeDate", "CashBalance"]]
                .rename(columns={"CashBalance": "Balance"})
            )
        else:
            group = group.copy()
            counted_cash = group["CashAmount"].where(~group["TransactionTypeRaw"].isin(EVLI_NON_CASH_EVENT_TYPES), 0.0)
            group["Balance"] = counted_cash.cumsum()
            balance_source = group[["TradeDate", "Balance"]]

        if len(balance_source) == 0:
            continue

        candidates = pd.DataFrame({"MonthEnd": months})
        merged = pd.merge_asof(
            candidates,
            balance_source.sort_values("TradeDate"),
            left_on="MonthEnd",
            right_on="TradeDate",
            direction="backward",
        )
        merged = merged.dropna(subset=["Balance"])

        owner = normalise_text(settings.portfolio_owner(broker, portfolio))
        portfolio_type = normalise_text(settings.portfolio_type(broker, portfolio))

        for _, r in merged.iterrows():
            month_end = r["MonthEnd"]
            rows.append({
                "MonthEnd": month_end.strftime("%Y-%m-%d"),
                "Year": month_end.year,
                "Month": month_end.month,
                "Broker": broker,
                "Portfolio": portfolio,
                "PortfolioOwner": owner,
                "PortfolioType": portfolio_type,
                "NormalizedInstrument": CASH_INSTRUMENT_LABEL,
                "CumulativeQuantity": "",
                "PriceLocal": "",
                "InstrumentCurrency": "EUR",
                "PriceDate": "",
                "FXRate": "",
                "FXDate": "",
                "MarketValueEUR": round(float(r["Balance"]), 2),
                # A cash balance isn't "invested" in an instrument and
                # doesn't have a gain/loss concept of its own - these stay
                # blank rather than 0, same convention as UnrealizedGainPercent's
                # existing blank-when-not-applicable case just below.
                "CumulativeNetInvested": "",
                "UnrealizedGainEUR": "",
                "UnrealizedGainPercent": "",
                "DividendGrossEUR": "",
                "DividendTaxEUR": "",
                "DividendNetEUR": "",
                "CumulativeDividendGrossEUR": "",
                "CumulativeDividendNetEUR": "",
            })

    return rows


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
        portfolio_type = normalise_text(group["PortfolioType"].iloc[0]) if "PortfolioType" in group.columns else ""
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
                "PortfolioType": portfolio_type,
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

    cash_rows: list[dict[str, object]] = []
    if investments_workbook is not None and investments_workbook.exists():
        cash_source = load_cash_balance_source(investments_workbook)
        cash_rows = build_cash_balance_rows(cash_source, last_month_end)
        rows.extend(cash_rows)

    result = pd.DataFrame(rows, columns=MONTHLY_VALUE_COLUMNS)
    if len(result) > 0:
        result = result.sort_values(["NormalizedInstrument", "Broker", "Portfolio", "MonthEnd"]).reset_index(drop=True)

    stats = {
        "rows_produced": len(result),
        "cash_balance_rows": len(cash_rows),
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
    parser.add_argument("--investments-workbook", default=str(DEFAULT_INVESTMENTS_WORKBOOK), help="Source of dividend/tax events and Nordnet/EVLI cash balances.")
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
    print(f"  of which CASH balance rows:  {stats['cash_balance_rows']}")

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
