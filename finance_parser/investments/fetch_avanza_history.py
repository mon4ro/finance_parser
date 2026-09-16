from __future__ import annotations

import argparse
import json
import urllib.request
from datetime import date, datetime, timezone
from pathlib import Path

import pandas as pd

from finance_parser.common import normalise_text
from finance_parser.investments.fetch_instrument_prices import (
    DEFAULT_INSTRUMENT_MASTER,
    DEFAULT_PRICES_WORKBOOK,
    PRICES_COLUMNS,
    USER_AGENT,
    imported_at_now,
    load_existing_prices,
    load_instrument_master_price_config,
    write_prices_workbook,
)

# Avanza (a large Swedish broker) exposes a public, no-key/no-login chart API
# for any fund it lists, even ones held elsewhere - verified directly
# (2026-09-16):
#   https://www.avanza.se/_api/fund-guide/chart/325406/2022-01-28/2022-03-07?raw=true
#   -> {"id":"325406","dataSerie":[{"x":<epoch_ms>,"y":<value>},...]}
#
# CRITICAL: Avanza appears to convert a non-SEK fund's NAV to SEK for
# display, even when the fund's own name says otherwise (confirmed: Franklin
# Technology A Acc USD's Avanza value was ~10x Yahoo's same-date USD value -
# right in line with the USD/SEK rate, not a real price difference). Only
# add a fund here if you have independently confirmed Avanza returns its
# real native-currency value - for a SEK-native fund this is a non-issue
# (SEK in, SEK out, nothing to hide), which is why only Spiltan is mapped
# below. Continuity-checked against real data before trusting it: Avanza's
# last pre-gap value (592.51) matched Yahoo's first post-gap value
# (592.51001) almost exactly.
#
# Explicitly NOT included, and do not add without re-verifying currency:
#   - FRANKLIN TECHNOLOGY A ACC (USD-native, Avanza id 397: confirmed SEK-converted)
#   - HANDELSBANKEN USA INDEX (EUR-native A1 EUR class held; the only Avanza
#     fund ID found, 315115, is a DIFFERENT share class, A1 SEK)
AVANZA_FUND_IDS = {
    "SPILTAN AKTIEFOND INVESTMENTBOLAG": "325406",
}

AVANZA_CHART_URL = "https://www.avanza.se/_api/fund-guide/chart/{fund_id}/{start}/{end}?raw=true"
PRICE_SOURCE = "AVANZA"


def fetch_avanza_chart(fund_id: str, start: date, end: date) -> list[tuple[date, float]]:
    url = AVANZA_CHART_URL.format(fund_id=fund_id, start=start.isoformat(), end=end.isoformat())
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=20) as response:
        payload = json.loads(response.read().decode("utf-8"))

    rows: list[tuple[date, float]] = []
    for point in payload.get("dataSerie", []):
        value = point.get("y")
        if value is None:
            # Placeholder points before the fund existed yet, when the
            # requested range starts earlier than the fund's real inception.
            continue
        day = datetime.fromtimestamp(point["x"] / 1000, tz=timezone.utc).date()
        rows.append((day, float(value)))

    return rows


def _instrument_isin_lookup(instrument_master_path: Path) -> dict[str, str]:
    config = load_instrument_master_price_config(instrument_master_path)
    lookup = {}
    for _, row in config.iterrows():
        isin = normalise_text(row.get("ISIN", ""))
        if isin:
            lookup[row["NormalizedInstrument"]] = isin
    return lookup


def backfill_gaps(
    prices_workbook: Path,
    instrument_master_path: Path,
    *,
    fetch_fn=fetch_avanza_chart,
    verbose: bool = True,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """
    Fill gaps only: an existing row for a (instrument, date) already on file
    is never touched or replaced. Only dates genuinely missing get an
    AVANZA-sourced row added. Mirrors fetch_seligson_history.py.
    """
    existing = load_existing_prices(prices_workbook)
    isin_lookup = _instrument_isin_lookup(instrument_master_path)

    stats: dict[str, object] = {"instruments_backfilled": 0, "rows_added": 0}
    new_rows: list[dict[str, object]] = []
    fetched_at = imported_at_now()
    today = datetime.now(timezone.utc).date()

    for instrument, fund_id in AVANZA_FUND_IDS.items():
        existing_for_instrument = existing[existing["NormalizedInstrument"] == instrument]
        existing_dates = set(pd.to_datetime(existing_for_instrument["Date"]).dt.date)

        if existing_dates:
            start = min(existing_dates)
        else:
            start = today

        if verbose:
            print(f"  Fetching Avanza history for {instrument} ({fund_id}) from earliest known date back to fund start ...")

        chart_rows = fetch_fn(fund_id, date(2000, 1, 1), start)

        added = 0
        for day, value in chart_rows:
            if day in existing_dates:
                continue
            new_rows.append({
                "NormalizedInstrument": instrument,
                "ISIN": isin_lookup.get(instrument, ""),
                "PriceSymbol": fund_id,
                "Date": day.isoformat(),
                "Close": value,
                "Currency": "SEK",
                "PriceSource": PRICE_SOURCE,
                "FetchedAt": fetched_at,
            })
            added += 1

        if added:
            stats["instruments_backfilled"] += 1
        stats["rows_added"] += added
        if verbose:
            print(f"    {added} new row(s) added (gap-fill only - existing rows untouched)")

    new_prices = pd.DataFrame(new_rows, columns=PRICES_COLUMNS)
    if len(existing) == 0:
        combined = new_prices
    elif len(new_prices) == 0:
        combined = existing
    else:
        combined = pd.concat([existing, new_prices], ignore_index=True)
    combined = combined.drop_duplicates(subset=["NormalizedInstrument", "Date"], keep="last")
    combined = combined.sort_values(["NormalizedInstrument", "Date"]).reset_index(drop=True)

    return combined[PRICES_COLUMNS], stats


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="One-time backfill: fill InstrumentPrices.xlsx gaps for AVANZA_FUND_IDS-mapped funds from "
                    "Avanza's own free chart API. Never touches an existing (instrument, date) row. Only add a "
                    "fund here after confirming Avanza returns its real native-currency value, not a SEK-converted one."
    )
    parser.add_argument("--prices-workbook", default=str(DEFAULT_PRICES_WORKBOOK))
    parser.add_argument("--instrument-master", default=str(DEFAULT_INSTRUMENT_MASTER))
    parser.add_argument("--dry-run", action="store_true")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()

    prices_workbook = Path(args.prices_workbook).expanduser().resolve()
    instrument_master_path = Path(args.instrument_master).expanduser().resolve()

    combined, stats = backfill_gaps(prices_workbook, instrument_master_path)

    print()
    print("Avanza backfill complete." if not args.dry_run else "Avanza backfill dry run complete.")
    print(f"Instruments backfilled: {stats['instruments_backfilled']}")
    print(f"Rows added:             {stats['rows_added']}")

    if args.dry_run:
        print()
        print("Dry run only: workbook was not modified.")
        return

    write_prices_workbook(prices_workbook, combined)
    print()
    print(f"Output workbook: {prices_workbook}")


if __name__ == "__main__":
    main()
