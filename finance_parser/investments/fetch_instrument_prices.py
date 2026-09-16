from __future__ import annotations

import argparse
import json
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import pandas as pd

from finance_parser.common import normalise_header, normalise_text
from finance_parser.utilities.fresh_workbook_writer import (
    read_workbook_values_only,
    records_to_sheet_values,
    replace_with_fresh_workbook,
    write_fresh_workbook,
)


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_INSTRUMENT_MASTER = PROJECT_ROOT / "rules" / "investments" / "InstrumentMaster.xlsx"
DEFAULT_INVESTMENTS_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "ParsedInvestments.xlsx"
DEFAULT_PRICES_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "InstrumentPrices.xlsx"

PRICES_SHEET = "InstrumentPrices"
PRICES_COLUMNS = [
    "NormalizedInstrument",
    "ISIN",
    "PriceSymbol",
    "Date",
    "Close",
    "Currency",
    "PriceSource",
    "FetchedAt",
]

# Politeness delay between calls to the (unofficial) price API.
REQUEST_DELAY_SECONDS = 0.5

USER_AGENT = "Mozilla/5.0 (compatible; finance_parser price fetcher)"


def imported_at_now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def load_instrument_master_price_config(path: Path) -> pd.DataFrame:
    """
    Read InstrumentMaster.xlsx and return only the columns this script needs,
    for rows that have a PriceSymbol configured.

    This deliberately does not reuse investment_common.load_instrument_master()
    (that function is shaped for transaction-matching priority, not price
    lookups) - here we just need one row per NormalizedInstrument with a price
    symbol.
    """
    if not path.exists():
        return pd.DataFrame(columns=["NormalizedInstrument", "ISIN", "PriceSource", "PriceSymbol", "Currency"])

    df = pd.read_excel(path, sheet_name="InstrumentMaster", dtype=object, engine="openpyxl")
    df.columns = [normalise_header(c) for c in df.columns]

    for col in ["Enabled", "NormalizedInstrument", "ISIN", "PriceSource", "PriceSymbol", "Currency"]:
        if col not in df.columns:
            df[col] = ""

    enabled = df["Enabled"].map(lambda v: normalise_text(v).upper() not in {"NO", "N", "FALSE", "0", "DISABLED"})
    df = df.loc[enabled].copy()

    has_symbol = df["PriceSymbol"].map(normalise_text).ne("")
    df = df.loc[has_symbol].copy()

    df["NormalizedInstrument"] = df["NormalizedInstrument"].map(normalise_text)
    df = df[df["NormalizedInstrument"] != ""].copy()

    # If the same NormalizedInstrument appears more than once (e.g. per-broker
    # rows), keep the first configured PriceSymbol - ambiguity here should be
    # resolved by hand in InstrumentMaster.xlsx, not guessed.
    df = df.drop_duplicates(subset=["NormalizedInstrument"], keep="first")

    return df[["NormalizedInstrument", "ISIN", "PriceSource", "PriceSymbol", "Currency"]]


def owned_normalized_instruments(investments_workbook: Path) -> pd.DataFrame:
    """
    Return one row per NormalizedInstrument that actually appears in
    InvestmentTransactions, plus the earliest TradeDate seen for it (used as
    the default price-history start date when nothing has been fetched yet).
    """
    if not investments_workbook.exists():
        return pd.DataFrame(columns=["NormalizedInstrument", "FirstTradeDate"])

    df = pd.read_excel(investments_workbook, sheet_name="InvestmentTransactions", dtype=object, engine="openpyxl")
    df.columns = [normalise_header(c) for c in df.columns]

    df["NormalizedInstrument"] = df["NormalizedInstrument"].map(normalise_text)
    df = df[df["NormalizedInstrument"] != ""].copy()

    df["_TradeDate"] = pd.to_datetime(df["TradeDate"], errors="coerce")
    grouped = df.groupby("NormalizedInstrument")["_TradeDate"].min().reset_index()
    grouped.columns = ["NormalizedInstrument", "FirstTradeDate"]

    return grouped


def build_fetch_plan(
    instrument_master_path: Path,
    investments_workbook: Path,
) -> tuple[pd.DataFrame, list[str]]:
    """
    Combine "owned" with "has a configured PriceSymbol" to decide what to
    fetch. Returns (plan, missing_price_symbol_instruments).
    """
    price_config = load_instrument_master_price_config(instrument_master_path)
    owned = owned_normalized_instruments(investments_workbook)

    plan = owned.merge(price_config, on="NormalizedInstrument", how="inner")

    owned_set = set(owned["NormalizedInstrument"])
    configured_set = set(price_config["NormalizedInstrument"])
    missing = sorted(owned_set - configured_set)

    return plan, missing


def _unix(dt: date) -> int:
    return int(datetime(dt.year, dt.month, dt.day, tzinfo=timezone.utc).timestamp())


def fetch_yahoo_daily_closes(symbol: str, start: date, end: date) -> tuple[list[tuple[date, float]], str]:
    """
    Fetch daily closing prices for one symbol from Yahoo Finance's chart API.

    Returns (rows, currency). Raises on network/parse failure - callers should
    catch per-instrument so one broken symbol doesn't stop the whole run.
    """
    if start > end:
        return [], ""

    url = (
        "https://query1.finance.yahoo.com/v8/finance/chart/"
        f"{urllib.parse.quote(symbol)}"
        f"?period1={_unix(start)}&period2={_unix(end) + 86400}&interval=1d"
    )

    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=20) as response:
        payload = json.loads(response.read().decode("utf-8"))

    result = payload.get("chart", {}).get("result")
    if not result:
        error = payload.get("chart", {}).get("error")
        raise ValueError(f"No data returned for {symbol!r}: {error}")

    result = result[0]
    currency = normalise_text(result.get("meta", {}).get("currency", ""))
    timestamps = result.get("timestamp") or []
    closes = result.get("indicators", {}).get("quote", [{}])[0].get("close") or []

    rows: list[tuple[date, float]] = []
    for ts, close in zip(timestamps, closes):
        if close is None:
            continue
        day = datetime.fromtimestamp(ts, tz=timezone.utc).date()
        rows.append((day, round(float(close), 6)))

    return rows, currency


def load_existing_prices(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=PRICES_COLUMNS)

    try:
        df = pd.read_excel(path, sheet_name=PRICES_SHEET, dtype=object, engine="openpyxl")
    except ValueError:
        return pd.DataFrame(columns=PRICES_COLUMNS)

    df.columns = [normalise_header(c) for c in df.columns]
    for col in PRICES_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    return df[PRICES_COLUMNS]


def fetch_all(
    instrument_master_path: Path,
    investments_workbook: Path,
    prices_workbook: Path,
    *,
    since_override: date | None = None,
    verbose: bool = True,
) -> tuple[pd.DataFrame, dict[str, object]]:
    plan, missing = build_fetch_plan(instrument_master_path, investments_workbook)
    existing_prices = load_existing_prices(prices_workbook)

    stats: dict[str, object] = {
        "instruments_planned": len(plan),
        "instruments_missing_price_symbol": missing,
        "instruments_fetched": 0,
        "instruments_up_to_date": 0,
        "instruments_failed": [],
        "rows_fetched": 0,
        "currency_mismatches": [],
    }

    new_rows: list[dict[str, object]] = []
    today = datetime.now(timezone.utc).date()

    for _, row in plan.iterrows():
        instrument = row["NormalizedInstrument"]
        symbol = row["PriceSymbol"]
        isin = normalise_text(row.get("ISIN", ""))
        source = normalise_text(row.get("PriceSource", "")) or "YAHOO"
        expected_currency = normalise_text(row.get("Currency", ""))

        existing_for_instrument = existing_prices[existing_prices["NormalizedInstrument"] == instrument]
        if len(existing_for_instrument) > 0:
            last_known = pd.to_datetime(existing_for_instrument["Date"]).max().date()
            start = last_known + timedelta(days=1)
        elif since_override is not None:
            start = since_override
        elif pd.notna(row.get("FirstTradeDate")):
            start = pd.to_datetime(row["FirstTradeDate"]).date()
        else:
            if verbose:
                print(f"  SKIP {instrument}: no existing prices and no known first trade date")
            continue

        if start > today:
            stats["instruments_up_to_date"] += 1
            continue

        if verbose:
            print(f"  Fetching {instrument} ({symbol}) from {start} to {today} ...")

        try:
            price_rows, fetched_currency = fetch_yahoo_daily_closes(symbol, start, today)
        except (urllib.error.URLError, ValueError, TimeoutError, json.JSONDecodeError) as exc:
            stats["instruments_failed"].append(f"{instrument} ({symbol}): {exc}")
            if verbose:
                print(f"    FAILED: {exc}")
            time.sleep(REQUEST_DELAY_SECONDS)
            continue

        if expected_currency and fetched_currency and expected_currency.upper() != fetched_currency.upper():
            stats["currency_mismatches"].append(
                f"{instrument}: InstrumentMaster says {expected_currency}, Yahoo says {fetched_currency}"
            )

        # Prefer InstrumentMaster's human-verified Currency over Yahoo's own
        # chart-meta currency when both are set. Real case that forced this:
        # BGF World Technology's working ticker (0P00000AWU) returns Yahoo
        # meta currency "USD", but its raw values match real EUR NAV history
        # almost exactly (cross-checked against manually-tracked EUR prices
        # at multiple dates, no FX-scaling factor needed) - Yahoo's own
        # currency tag is simply wrong here, not the price data. Falls back
        # to the fetched currency only when InstrumentMaster has none set.
        currency = expected_currency or fetched_currency

        fetched_at = imported_at_now()
        for day, close in price_rows:
            new_rows.append({
                "NormalizedInstrument": instrument,
                "ISIN": isin,
                "PriceSymbol": symbol,
                "Date": day.isoformat(),
                "Close": close,
                "Currency": currency,
                "PriceSource": source,
                "FetchedAt": fetched_at,
            })

        stats["instruments_fetched"] += 1
        stats["rows_fetched"] += len(price_rows)
        time.sleep(REQUEST_DELAY_SECONDS)

    new_prices = pd.DataFrame(new_rows, columns=PRICES_COLUMNS)
    if len(existing_prices) == 0:
        combined = new_prices
    elif len(new_prices) == 0:
        combined = existing_prices
    else:
        combined = pd.concat([existing_prices, new_prices], ignore_index=True)
    combined = combined.drop_duplicates(subset=["NormalizedInstrument", "Date"], keep="last")
    combined = combined.sort_values(["NormalizedInstrument", "Date"]).reset_index(drop=True)

    return combined[PRICES_COLUMNS], stats


def write_prices_workbook(prices_workbook: Path, combined_prices: pd.DataFrame) -> None:
    if prices_workbook.exists():
        sheets = read_workbook_values_only(prices_workbook)
    else:
        sheets = {}

    headers = list(PRICES_COLUMNS)
    records = combined_prices.to_dict("records")
    sheets[PRICES_SHEET] = records_to_sheet_values(headers, records)

    if prices_workbook.exists():
        replace_with_fresh_workbook(
            prices_workbook,
            sheets,
            backup_label="before_price_fetch",
            basic_formatting=True,
            excel_tables=False,
        )
    else:
        write_fresh_workbook(prices_workbook, sheets, basic_formatting=True, excel_tables=False)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Fetch daily closing prices for owned instruments that have a PriceSymbol configured in InstrumentMaster.xlsx."
    )
    parser.add_argument("--instrument-master", default=str(DEFAULT_INSTRUMENT_MASTER))
    parser.add_argument("--investments-workbook", default=str(DEFAULT_INVESTMENTS_WORKBOOK))
    parser.add_argument("--prices-workbook", default=str(DEFAULT_PRICES_WORKBOOK))
    parser.add_argument(
        "--since",
        default=None,
        help="Override the fetch start date (YYYY-MM-DD) for instruments with no price history yet.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Fetch and report, but do not write the workbook.")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()

    instrument_master_path = Path(args.instrument_master).expanduser().resolve()
    investments_workbook = Path(args.investments_workbook).expanduser().resolve()
    prices_workbook = Path(args.prices_workbook).expanduser().resolve()
    since_override = date.fromisoformat(args.since) if args.since else None

    combined_prices, stats = fetch_all(
        instrument_master_path,
        investments_workbook,
        prices_workbook,
        since_override=since_override,
    )

    print()
    print("Price fetch complete." if not args.dry_run else "Price fetch dry run complete.")
    print(f"Instruments planned (owned + PriceSymbol configured): {stats['instruments_planned']}")
    print(f"Instruments fetched this run:                         {stats['instruments_fetched']}")
    print(f"Instruments already up to date:                       {stats['instruments_up_to_date']}")
    print(f"New price rows:                                       {stats['rows_fetched']}")

    if stats["instruments_missing_price_symbol"]:
        print()
        print(f"Owned instruments with NO PriceSymbol configured ({len(stats['instruments_missing_price_symbol'])}):")
        for name in stats["instruments_missing_price_symbol"]:
            print(f"  - {name}")

    if stats["currency_mismatches"]:
        print()
        print("Currency mismatches (InstrumentMaster vs fetched data):")
        for msg in stats["currency_mismatches"]:
            print(f"  - {msg}")

    if stats["instruments_failed"]:
        print()
        print("Fetch failures:")
        for msg in stats["instruments_failed"]:
            print(f"  - {msg}")

    if args.dry_run:
        print()
        print("Dry run only: workbook was not modified.")
        return

    write_prices_workbook(prices_workbook, combined_prices)
    print()
    print(f"Output workbook: {prices_workbook}")


if __name__ == "__main__":
    main()
