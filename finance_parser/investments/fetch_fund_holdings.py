from __future__ import annotations

import argparse
import http.cookiejar
import json
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from finance_parser.common import normalise_header, normalise_text
from finance_parser.investments.check_instrument_coverage import active_instruments
from finance_parser.utilities.fresh_workbook_writer import (
    read_workbook_values_only,
    records_to_sheet_values,
    replace_with_fresh_workbook,
    write_fresh_workbook,
)


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_INSTRUMENT_MASTER = PROJECT_ROOT / "rules" / "investments" / "InstrumentMaster.xlsx"
DEFAULT_POSITIONS_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "PortfolioPositions.xlsx"
DEFAULT_HOLDINGS_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "FundHoldingsSnapshot.xlsx"

HOLDINGS_SHEET = "FundHoldingsSnapshot"
HOLDINGS_COLUMNS = [
    "SnapshotMonth",
    "SnapshotDate",
    "PriceSymbol",
    "NormalizedInstrument",
    "HoldingRank",
    "HoldingSymbol",
    "HoldingName",
    "HoldingPercent",
    "SnapshotSource",
]

LIVE_FETCH = "LIVE_FETCH"
BACKFILLED = "BACKFILLED_ASSUMED_CONSTANT"
UNKNOWN_SYMBOL = "UNKNOWN"
UNKNOWN_NO_DATA_NAME = "No holdings data available"
UNKNOWN_RESIDUAL_NAME = "Unmapped remainder (beyond top 10)"

# Yahoo's topHoldings module caps out at 10 entries regardless of fund size -
# confirmed by testing several real funds plus a large, well-diversified US
# ETF as a control (all returned exactly 10). The UNKNOWN residual row below
# is what makes the gap beyond this cap visible instead of silently missing.
MAX_HOLDINGS = 10

# There's no historical/point-in-time endpoint for fund holdings - Yahoo only
# ever exposes the current snapshot. On a fund's first-ever fetch, we assume
# today's holdings were constant for this many months back rather than have
# no data at all for the pre-tracking period. These backfilled rows are
# tagged BACKFILLED_ASSUMED_CONSTANT (never LIVE_FETCH) so downstream
# analysis - or a human - can always tell the difference.
BACKFILL_MONTHS = 60

USER_AGENT = "Mozilla/5.0 (compatible; finance_parser fund holdings fetcher)"
REQUEST_DELAY_SECONDS = 0.5


def imported_at_now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def current_month_str(today: date) -> str:
    return today.strftime("%Y-%m")


def month_str_n_back(today: date, n: int) -> str:
    """n=0 -> today's month, n=1 -> the month before that, etc."""
    year = today.year
    month = today.month - n
    while month <= 0:
        month += 12
        year -= 1
    return f"{year:04d}-{month:02d}"


def load_fund_price_config(path: Path) -> pd.DataFrame:
    """
    Funds only (InstrumentType == FUND) with a PriceSymbol configured.

    Stocks are deliberately excluded - a directly-held stock's exposure to
    itself is already 100% known from PortfolioPositions.xlsx, so there's
    nothing this constituent-holdings data can add for it.
    """
    if not path.exists():
        return pd.DataFrame(columns=["NormalizedInstrument", "PriceSymbol"])

    df = pd.read_excel(path, sheet_name="InstrumentMaster", dtype=object, engine="openpyxl")
    df.columns = [normalise_header(c) for c in df.columns]

    for col in ["Enabled", "InstrumentType", "NormalizedInstrument", "PriceSymbol"]:
        if col not in df.columns:
            df[col] = ""

    enabled = df["Enabled"].map(lambda v: normalise_text(v).upper() not in {"NO", "N", "FALSE", "0", "DISABLED"})
    is_fund = df["InstrumentType"].map(normalise_text).str.upper() == "FUND"
    df = df.loc[enabled & is_fund].copy()

    df["NormalizedInstrument"] = df["NormalizedInstrument"].map(normalise_text)
    df["PriceSymbol"] = df["PriceSymbol"].map(normalise_text)
    df = df[(df["NormalizedInstrument"] != "") & (df["PriceSymbol"] != "")].copy()

    df = df.drop_duplicates(subset=["NormalizedInstrument"], keep="first")
    return df[["NormalizedInstrument", "PriceSymbol"]]


def active_fund_plan(instrument_master_path: Path, positions_workbook: Path) -> pd.DataFrame:
    """Funds that are both configured (PriceSymbol set) and currently actively held."""
    funds = load_fund_price_config(instrument_master_path)
    if not positions_workbook.exists() or len(funds) == 0:
        return funds.iloc[0:0]

    active = active_instruments(positions_workbook)
    active_names = set(active["NormalizedInstrument"])
    return funds[funds["NormalizedInstrument"].isin(active_names)].reset_index(drop=True)


def build_yahoo_opener() -> urllib.request.OpenerDirector:
    return urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))


def get_yahoo_crumb(opener: urllib.request.OpenerDirector) -> str:
    """
    Yahoo's v10 quoteSummary API (needed for fund holdings - unlike the v8
    chart API this project already relies on for prices) now rejects
    anonymous requests. Unofficial, undocumented workaround: an
    unauthenticated cookie fetch against fc.yahoo.com (tolerates a 404 - it
    still sets the cookie we need), then a crumb token from a second
    unauthenticated endpoint using the same cookie jar.

    The returned crumb is only valid together with the session cookies the
    opener now holds - the caller must reuse this exact opener for every
    subsequent fetch_top_holdings() call, not just pass the crumb string
    around on its own (confirmed the hard way: a fresh urlopen() call with no
    cookie jar gets a 401 "Invalid Crumb" even with a crumb that a moment
    earlier was valid).

    This is a meaningfully more fragile dependency than the v8 API - it's
    routing around anti-scraping measures, not a stable public contract -
    so a failure here is handled as "fund holdings unavailable this run",
    never as a hard pipeline error.
    """
    try:
        opener.open(urllib.request.Request("https://fc.yahoo.com", headers={"User-Agent": USER_AGENT}), timeout=15)
    except urllib.error.HTTPError:
        pass  # Expected - this endpoint 404s but still sets the cookie we need.

    response = opener.open(
        urllib.request.Request(
            "https://query1.finance.yahoo.com/v1/test/getcrumb",
            headers={"User-Agent": USER_AGENT},
        ),
        timeout=15,
    )
    return response.read().decode("utf-8").strip()


def fetch_top_holdings(symbol: str, crumb: str, opener: urllib.request.OpenerDirector) -> list[dict[str, object]]:
    """
    Returns a list of {symbol, name, percent} dicts, capped at MAX_HOLDINGS.

    An empty list means "fetched OK, but Yahoo has no holdings data for this
    symbol" (a real, seen case for at least one real fund ticker) - the
    caller can't distinguish that from "the fund genuinely holds nothing" and
    shouldn't try to; both convert to a 100% UNKNOWN row downstream.

    Must use the same opener (and its session cookies) that produced crumb -
    see get_yahoo_crumb()'s docstring.
    """
    url = (
        "https://query1.finance.yahoo.com/v10/finance/quoteSummary/"
        f"{urllib.parse.quote(symbol)}?modules=topHoldings&crumb={urllib.parse.quote(crumb)}"
    )
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with opener.open(request, timeout=20) as response:
        payload = json.loads(response.read().decode("utf-8"))

    result = payload.get("quoteSummary", {}).get("result")
    if not result:
        error = payload.get("quoteSummary", {}).get("error")
        raise ValueError(f"No topHoldings result for {symbol!r}: {error}")

    holdings = result[0].get("topHoldings", {}).get("holdings") or []
    parsed: list[dict[str, object]] = []
    for holding in holdings[:MAX_HOLDINGS]:
        percent_raw = (holding.get("holdingPercent") or {}).get("raw")
        parsed.append({
            "symbol": normalise_text(holding.get("symbol", "")),
            "name": normalise_text(holding.get("holdingName", "")),
            "percent": round(float(percent_raw) * 100, 4) if percent_raw is not None else 0.0,
        })
    return parsed


def build_snapshot_rows(
    price_symbol: str,
    instrument: str,
    snapshot_month: str,
    snapshot_date: date,
    holdings: list[dict[str, object]] | None,
    *,
    source: str,
) -> list[dict[str, object]]:
    """
    One fund/month always sums to exactly 100%: real holdings (if any) plus
    an explicit UNKNOWN row for whatever's left over - either the gap beyond
    the top 10, or the whole 100% if nothing was fetchable at all. A
    look-through/concentration calculation reading this table never needs to
    know *why* a residual exists, just that it does - it's never silently
    missing.
    """
    rows: list[dict[str, object]] = []
    covered = 0.0

    if holdings:
        for rank, holding in enumerate(holdings, start=1):
            rows.append({
                "SnapshotMonth": snapshot_month,
                "SnapshotDate": snapshot_date.isoformat(),
                "PriceSymbol": price_symbol,
                "NormalizedInstrument": instrument,
                "HoldingRank": rank,
                "HoldingSymbol": holding["symbol"] or holding["name"],
                "HoldingName": holding["name"],
                "HoldingPercent": holding["percent"],
                "SnapshotSource": source,
            })
            covered += holding["percent"]

    remainder = round(max(0.0, 100.0 - covered), 4)
    if remainder > 0.0:
        rows.append({
            "SnapshotMonth": snapshot_month,
            "SnapshotDate": snapshot_date.isoformat(),
            "PriceSymbol": price_symbol,
            "NormalizedInstrument": instrument,
            "HoldingRank": MAX_HOLDINGS + 1,
            "HoldingSymbol": UNKNOWN_SYMBOL,
            "HoldingName": UNKNOWN_NO_DATA_NAME if not holdings else UNKNOWN_RESIDUAL_NAME,
            "HoldingPercent": remainder,
            "SnapshotSource": source,
        })

    return rows


def build_backfill_rows(
    price_symbol: str,
    instrument: str,
    live_holdings: list[dict[str, object]] | None,
    today: date,
    *,
    months_back: int,
) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for n in range(1, months_back + 1):
        month_str = month_str_n_back(today, n)
        year, month = (int(part) for part in month_str.split("-"))
        snapshot_date = date(year, month, 1)
        rows.extend(
            build_snapshot_rows(price_symbol, instrument, month_str, snapshot_date, live_holdings, source=BACKFILLED)
        )
    return rows


def load_existing_holdings(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=HOLDINGS_COLUMNS)

    try:
        df = pd.read_excel(path, sheet_name=HOLDINGS_SHEET, dtype=object, engine="openpyxl")
    except ValueError:
        return pd.DataFrame(columns=HOLDINGS_COLUMNS)

    df.columns = [normalise_header(c) for c in df.columns]
    for col in HOLDINGS_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    return df[HOLDINGS_COLUMNS]


def fetch_all(
    instrument_master_path: Path,
    positions_workbook: Path,
    holdings_workbook: Path,
    *,
    force_refetch: bool = False,
    verbose: bool = True,
) -> tuple[pd.DataFrame, dict[str, object]]:
    plan = active_fund_plan(instrument_master_path, positions_workbook)
    existing = load_existing_holdings(holdings_workbook)

    stats: dict[str, object] = {
        "funds_planned": len(plan),
        "funds_fetched": 0,
        "funds_up_to_date": 0,
        "funds_no_data": [],
        "funds_failed": [],
        "funds_backfilled": [],
        "rows_added": 0,
    }

    today = datetime.now().date()
    current_month = current_month_str(today)

    opener = build_yahoo_opener()
    crumb: str | None = None
    crumb_error: str | None = None
    new_rows: list[dict[str, object]] = []

    for _, row in plan.iterrows():
        instrument = row["NormalizedInstrument"]
        symbol = row["PriceSymbol"]

        already_done = (
            (existing["PriceSymbol"] == symbol) & (existing["SnapshotMonth"] == current_month)
        ).any()
        if already_done and not force_refetch:
            stats["funds_up_to_date"] += 1
            continue

        is_first_time = not (existing["PriceSymbol"] == symbol).any()

        if crumb is None and crumb_error is None:
            try:
                crumb = get_yahoo_crumb(opener)
            except (urllib.error.URLError, TimeoutError) as exc:
                crumb_error = str(exc)

        holdings: list[dict[str, object]] | None = None
        if crumb is not None:
            if verbose:
                print(f"  Fetching fund holdings: {instrument} ({symbol}) ...")
            try:
                holdings = fetch_top_holdings(symbol, crumb, opener)
                if not holdings:
                    stats["funds_no_data"].append(instrument)
            except (urllib.error.URLError, ValueError, TimeoutError, json.JSONDecodeError) as exc:
                stats["funds_failed"].append(f"{instrument} ({symbol}): {exc}")
                if verbose:
                    print(f"    FAILED: {exc}")
                holdings = None
            time.sleep(REQUEST_DELAY_SECONDS)
        else:
            stats["funds_failed"].append(f"{instrument} ({symbol}): crumb unavailable ({crumb_error})")

        live_rows = build_snapshot_rows(symbol, instrument, current_month, today, holdings, source=LIVE_FETCH)
        new_rows.extend(live_rows)
        stats["funds_fetched"] += 1

        if is_first_time:
            backfill_rows = build_backfill_rows(symbol, instrument, holdings, today, months_back=BACKFILL_MONTHS)
            new_rows.extend(backfill_rows)
            stats["funds_backfilled"].append(instrument)

    new_df = pd.DataFrame(new_rows, columns=HOLDINGS_COLUMNS)
    if len(existing) == 0:
        combined = new_df
    elif len(new_df) == 0:
        combined = existing
    else:
        combined = pd.concat([existing, new_df], ignore_index=True)

    combined = combined.drop_duplicates(subset=["PriceSymbol", "SnapshotMonth", "HoldingSymbol"], keep="last")
    combined = combined.sort_values(["PriceSymbol", "SnapshotMonth", "HoldingRank"]).reset_index(drop=True)

    stats["rows_added"] = len(new_rows)
    return combined[HOLDINGS_COLUMNS], stats


def write_holdings_workbook(holdings_workbook: Path, combined: pd.DataFrame) -> None:
    if holdings_workbook.exists():
        sheets = read_workbook_values_only(holdings_workbook)
    else:
        sheets = {}

    headers = list(HOLDINGS_COLUMNS)
    records = combined.to_dict("records")
    sheets[HOLDINGS_SHEET] = records_to_sheet_values(headers, records)

    if holdings_workbook.exists():
        replace_with_fresh_workbook(
            holdings_workbook,
            sheets,
            backup_label="before_fund_holdings_fetch",
            basic_formatting=True,
            excel_tables=False,
        )
    else:
        write_fresh_workbook(holdings_workbook, sheets, basic_formatting=True, excel_tables=False)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Fetch top-10 constituent holdings for actively-held funds, once per calendar month, for "
            "portfolio look-through/concentration analysis. Uses an unofficial Yahoo Finance endpoint "
            "(cookie/crumb workaround) - more fragile than the price-fetch API this project otherwise "
            "relies on. A fund with no data, or any fetch failure, is recorded as 100% UNKNOWN rather "
            "than silently skipped. A fund's first-ever fetch also backfills "
            f"{BACKFILL_MONTHS} months of assumed-constant history, tagged distinctly from real snapshots."
        )
    )
    parser.add_argument("--instrument-master", default=str(DEFAULT_INSTRUMENT_MASTER))
    parser.add_argument("--positions-workbook", default=str(DEFAULT_POSITIONS_WORKBOOK))
    parser.add_argument("--holdings-workbook", default=str(DEFAULT_HOLDINGS_WORKBOOK))
    parser.add_argument(
        "--force-refetch",
        action="store_true",
        help="Refetch even for funds that already have a snapshot for the current month.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Fetch and report, but do not write the workbook.")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()

    instrument_master_path = Path(args.instrument_master).expanduser().resolve()
    positions_workbook = Path(args.positions_workbook).expanduser().resolve()
    holdings_workbook = Path(args.holdings_workbook).expanduser().resolve()

    combined, stats = fetch_all(
        instrument_master_path,
        positions_workbook,
        holdings_workbook,
        force_refetch=args.force_refetch,
    )

    print()
    print("Fund holdings fetch complete." if not args.dry_run else "Fund holdings fetch dry run complete.")
    print(f"Active funds planned:         {stats['funds_planned']}")
    print(f"Fetched this run:             {stats['funds_fetched']}")
    print(f"Already up to date (skipped): {stats['funds_up_to_date']}")
    print(f"New rows:                     {stats['rows_added']}")

    if stats["funds_backfilled"]:
        print()
        print(f"First-time funds backfilled with {BACKFILL_MONTHS} months of assumed-constant history:")
        for name in stats["funds_backfilled"]:
            print(f"  - {name}")

    if stats["funds_no_data"]:
        print()
        print("Funds with no holdings data available from Yahoo (recorded as 100% UNKNOWN):")
        for name in stats["funds_no_data"]:
            print(f"  - {name}")

    if stats["funds_failed"]:
        print()
        print("Fetch failures (recorded as 100% UNKNOWN, will retry next run):")
        for msg in stats["funds_failed"]:
            print(f"  - {msg}")

    if args.dry_run:
        print()
        print("Dry run only: workbook was not modified.")
        return

    write_holdings_workbook(holdings_workbook, combined)
    print()
    print(f"Output workbook: {holdings_workbook}")


if __name__ == "__main__":
    main()
