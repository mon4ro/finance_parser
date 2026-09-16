from __future__ import annotations

import argparse
import json
import time
import urllib.error
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import pandas as pd

from finance_parser.common import normalise_header, normalise_text
from finance_parser.investments.fetch_instrument_prices import (
    REQUEST_DELAY_SECONDS,
    fetch_yahoo_daily_closes,
    imported_at_now,
)
from finance_parser.utilities.fresh_workbook_writer import (
    read_workbook_values_only,
    records_to_sheet_values,
    replace_with_fresh_workbook,
    write_fresh_workbook,
)


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_INVESTMENTS_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "ParsedInvestments.xlsx"
DEFAULT_FX_RATES_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "FXRates.xlsx"

FX_RATES_SHEET = "FXRates"
FX_RATES_COLUMNS = ["Currency", "Date", "Rate", "Source", "FetchedAt"]

# Every rate is "1 unit of Currency = Rate EUR" - multiply a non-EUR amount by
# Rate to get its EUR value, matching the Master Budget Excel's existing
# "Price x Quantity x Exchange rate to EUR" convention. Confirmed empirically:
# Yahoo's "<CCY>EUR=X" tickers return exactly this direction (meta currency
# comes back as EUR, values are plausible for CCY per 1 EUR being their
# reciprocal), so no inversion is needed anywhere in this file.
BASE_CURRENCY = "EUR"


def fx_symbol(currency: str) -> str:
    return f"{currency}{BASE_CURRENCY}=X"


def owned_non_eur_currencies(investments_workbook: Path) -> pd.DataFrame:
    """
    Distinct non-EUR InstrumentCurrency values actually held (appear in real
    InvestmentTransactions), plus the earliest TradeDate seen for each - used
    as the default FX-history start date when nothing has been fetched yet.
    Mirrors owned_normalized_instruments() in fetch_instrument_prices.py.
    """
    if not investments_workbook.exists():
        return pd.DataFrame(columns=["Currency", "FirstTradeDate"])

    df = pd.read_excel(investments_workbook, sheet_name="InvestmentTransactions", dtype=object, engine="openpyxl")
    df.columns = [normalise_header(c) for c in df.columns]

    df["Currency"] = df["InstrumentCurrency"].map(normalise_text).str.upper()
    df = df[(df["Currency"] != "") & (df["Currency"] != BASE_CURRENCY)].copy()

    df["_TradeDate"] = pd.to_datetime(df["TradeDate"], errors="coerce")
    grouped = df.groupby("Currency")["_TradeDate"].min().reset_index()
    grouped.columns = ["Currency", "FirstTradeDate"]

    return grouped


def load_existing_fx_rates(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=FX_RATES_COLUMNS)

    try:
        df = pd.read_excel(path, sheet_name=FX_RATES_SHEET, dtype=object, engine="openpyxl")
    except ValueError:
        return pd.DataFrame(columns=FX_RATES_COLUMNS)

    df.columns = [normalise_header(c) for c in df.columns]
    for col in FX_RATES_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    return df[FX_RATES_COLUMNS]


def fetch_all(
    investments_workbook: Path,
    fx_rates_workbook: Path,
    *,
    since_override: date | None = None,
    verbose: bool = True,
) -> tuple[pd.DataFrame, dict[str, object]]:
    plan = owned_non_eur_currencies(investments_workbook)
    existing = load_existing_fx_rates(fx_rates_workbook)

    stats: dict[str, object] = {
        "currencies_planned": len(plan),
        "currencies_fetched": 0,
        "currencies_up_to_date": 0,
        "currencies_failed": [],
        "rows_fetched": 0,
    }

    new_rows: list[dict[str, object]] = []
    today = datetime.now(timezone.utc).date()

    for _, row in plan.iterrows():
        currency = row["Currency"]
        symbol = fx_symbol(currency)

        existing_for_currency = existing[existing["Currency"] == currency]
        if len(existing_for_currency) > 0:
            last_known = pd.to_datetime(existing_for_currency["Date"]).max().date()
            start = last_known + timedelta(days=1)
        elif since_override is not None:
            start = since_override
        elif pd.notna(row.get("FirstTradeDate")):
            start = pd.to_datetime(row["FirstTradeDate"]).date()
        else:
            if verbose:
                print(f"  SKIP {currency}: no existing rates and no known first trade date")
            continue

        if start > today:
            stats["currencies_up_to_date"] += 1
            continue

        if verbose:
            print(f"  Fetching {currency} ({symbol}) from {start} to {today} ...")

        try:
            rate_rows, _ = fetch_yahoo_daily_closes(symbol, start, today)
        except (urllib.error.URLError, ValueError, TimeoutError, json.JSONDecodeError) as exc:
            stats["currencies_failed"].append(f"{currency} ({symbol}): {exc}")
            if verbose:
                print(f"    FAILED: {exc}")
            time.sleep(REQUEST_DELAY_SECONDS)
            continue

        fetched_at = imported_at_now()
        for day, rate in rate_rows:
            new_rows.append({
                "Currency": currency,
                "Date": day.isoformat(),
                "Rate": rate,
                "Source": "YAHOO",
                "FetchedAt": fetched_at,
            })

        stats["currencies_fetched"] += 1
        stats["rows_fetched"] += len(rate_rows)
        time.sleep(REQUEST_DELAY_SECONDS)

    new_rates = pd.DataFrame(new_rows, columns=FX_RATES_COLUMNS)
    if len(existing) == 0:
        combined = new_rates
    elif len(new_rates) == 0:
        combined = existing
    else:
        combined = pd.concat([existing, new_rates], ignore_index=True)
    combined = combined.drop_duplicates(subset=["Currency", "Date"], keep="last")
    combined = combined.sort_values(["Currency", "Date"]).reset_index(drop=True)

    return combined[FX_RATES_COLUMNS], stats


def write_fx_rates_workbook(fx_rates_workbook: Path, combined: pd.DataFrame) -> None:
    if fx_rates_workbook.exists():
        sheets = read_workbook_values_only(fx_rates_workbook)
    else:
        sheets = {}

    headers = list(FX_RATES_COLUMNS)
    records = combined.to_dict("records")
    sheets[FX_RATES_SHEET] = records_to_sheet_values(headers, records)

    if fx_rates_workbook.exists():
        replace_with_fresh_workbook(
            fx_rates_workbook,
            sheets,
            backup_label="before_fx_fetch",
            basic_formatting=True,
            excel_tables=False,
        )
    else:
        write_fresh_workbook(fx_rates_workbook, sheets, basic_formatting=True, excel_tables=False)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Fetch daily EUR exchange rates for every non-EUR currency actually held, "
                    "determined from real InvestmentTransactions - not a fixed currency list."
    )
    parser.add_argument("--investments-workbook", default=str(DEFAULT_INVESTMENTS_WORKBOOK))
    parser.add_argument("--fx-rates-workbook", default=str(DEFAULT_FX_RATES_WORKBOOK))
    parser.add_argument(
        "--since",
        default=None,
        help="Override the fetch start date (YYYY-MM-DD) for currencies with no rate history yet.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Fetch and report, but do not write the workbook.")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()

    investments_workbook = Path(args.investments_workbook).expanduser().resolve()
    fx_rates_workbook = Path(args.fx_rates_workbook).expanduser().resolve()
    since_override = date.fromisoformat(args.since) if args.since else None

    combined, stats = fetch_all(investments_workbook, fx_rates_workbook, since_override=since_override)

    print()
    print("FX rate fetch complete." if not args.dry_run else "FX rate fetch dry run complete.")
    print(f"Currencies planned (non-EUR, actually held): {stats['currencies_planned']}")
    print(f"Currencies fetched this run:                 {stats['currencies_fetched']}")
    print(f"Currencies already up to date:                {stats['currencies_up_to_date']}")
    print(f"New rate rows:                                {stats['rows_fetched']}")

    if stats["currencies_failed"]:
        print()
        print("Fetch failures:")
        for msg in stats["currencies_failed"]:
            print(f"  - {msg}")

    if args.dry_run:
        print()
        print("Dry run only: workbook was not modified.")
        return

    write_fx_rates_workbook(fx_rates_workbook, combined)
    print()
    print(f"Output workbook: {fx_rates_workbook}")


if __name__ == "__main__":
    main()
