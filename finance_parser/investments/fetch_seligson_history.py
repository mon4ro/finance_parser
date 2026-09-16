from __future__ import annotations

import argparse
import urllib.request
from datetime import date, datetime
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

# Seligson & Co publishes full daily NAV history per fund as a free public
# CSV, no key/login needed - verified directly (2026-09-16):
#   https://www.seligson.fi/graafit/global-brands_exc.csv
#   18.06.1998;1,6819 ... (oldest to newest, no header)
# Finnish locale: semicolon-delimited, comma decimal, DD.MM.YYYY dates.
# Currency is always EUR for Seligson's own funds. This is a genuine
# exception among the fund houses checked - OP/Nordnet/Handelsbanken/
# Franklin/Spiltan have no equivalent free, no-login source (Investing.com,
# the one lead common to several of them, gates the actual download behind
# login and only shows ~1 month of history in the free view).
#
# No discoverable API/pattern maps NormalizedInstrument -> URL slug
# automatically, so this is a hand-maintained mapping - extend it if another
# Seligson fund gets held.
SELIGSON_CSV_SLUGS = {
    "SELIGSON GLOBAL TOP 25 BRANDS": "global-brands",
}

SELIGSON_CSV_URL = "https://www.seligson.fi/graafit/{slug}_exc.csv"
PRICE_SOURCE = "SELIGSON.FI"


def _parse_seligson_csv(text: str) -> list[tuple[date, float]]:
    rows: list[tuple[date, float]] = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        day_text, value_text = line.split(";")
        day = datetime.strptime(day_text, "%d.%m.%Y").date()
        value = float(value_text.replace(",", "."))
        rows.append((day, value))

    return rows


def fetch_seligson_csv(slug: str) -> list[tuple[date, float]]:
    url = SELIGSON_CSV_URL.format(slug=slug)
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=20) as response:
        text = response.read().decode("iso-8859-1")

    return _parse_seligson_csv(text)


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
    fetch_fn=fetch_seligson_csv,
    verbose: bool = True,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """
    Fill gaps only: an existing Yahoo-sourced row for a (instrument, date)
    already on file is never touched or replaced. Only dates genuinely
    missing get a Seligson-sourced row added.

    fetch_fn is injectable (defaults to the real network fetch) so tests can
    exercise the gap-filling/merge logic against a synthetic CSV without
    hitting the network.
    """
    existing = load_existing_prices(prices_workbook)
    isin_lookup = _instrument_isin_lookup(instrument_master_path)

    stats: dict[str, object] = {"instruments_backfilled": 0, "rows_added": 0}
    new_rows: list[dict[str, object]] = []
    fetched_at = imported_at_now()

    for instrument, slug in SELIGSON_CSV_SLUGS.items():
        if verbose:
            print(f"  Fetching Seligson history for {instrument} ({slug}) ...")

        csv_rows = fetch_fn(slug)

        existing_dates = set(
            pd.to_datetime(existing[existing["NormalizedInstrument"] == instrument]["Date"]).dt.date
        )

        added = 0
        for day, value in csv_rows:
            if day in existing_dates:
                continue
            new_rows.append({
                "NormalizedInstrument": instrument,
                "ISIN": isin_lookup.get(instrument, ""),
                "PriceSymbol": slug,
                "Date": day.isoformat(),
                "Close": value,
                "Currency": "EUR",
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
        description="One-time backfill: fill InstrumentPrices.xlsx gaps for Seligson funds from seligson.fi's "
                    "own free historical NAV CSV. Never touches an existing (instrument, date) row."
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
    print("Seligson backfill complete." if not args.dry_run else "Seligson backfill dry run complete.")
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
