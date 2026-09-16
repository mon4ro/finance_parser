from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

import pandas as pd

from finance_parser.common import normalise_text
from finance_parser.investments.fetch_instrument_prices import (
    DEFAULT_INSTRUMENT_MASTER,
    DEFAULT_PRICES_WORKBOOK,
)
from finance_parser.investments.investment_common import load_instrument_master

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_POSITIONS_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "PortfolioPositions.xlsx"

QUANTITY_EPSILON = 1e-6
DEFAULT_STALE_DAYS = 7

# The three gap classes this checks for, in order of severity:
#
# 1. unclassified - InstrumentMaster has no InstrumentType/Currency for a
#    currently-held instrument. Nothing downstream can price it correctly.
# 2. unpriced - classified, but no PriceSymbol configured (or configured
#    but zero price rows ever fetched). Same practical effect as #1 for net
#    worth purposes, different cause.
# 3. stale - has real price history, but the most recent price is older
#    than --stale-days. Real case this catches: BGF World Technology's
#    dead Yahoo ticker left it "priced" but frozen at an April value for
#    months before being caught by chance, not by any automated check.
#
# Deliberately scoped to ACTIVE positions only (CumulativeQuantity > epsilon
# as of the latest known event per (Broker, Portfolio, NormalizedInstrument)
# in PortfolioPositions.xlsx) - a defunct/fully-sold instrument (e.g.
# OP-Delta A, a real permanent gap accepted earlier this project) must not
# force this check to fail on every single run forever. Verified against
# real data before building this: OP-Delta A and the two Pienyhtiot funds
# all correctly fall out of the active set today.


def active_instruments(positions_workbook: Path) -> pd.DataFrame:
    df = pd.read_excel(positions_workbook, sheet_name="PortfolioPositions", dtype=object)
    df["Date"] = pd.to_datetime(df["Date"])
    df["CumulativeQuantity"] = pd.to_numeric(df["CumulativeQuantity"])

    latest = (
        df.sort_values("Date")
        .groupby(["Broker", "Portfolio", "NormalizedInstrument"], as_index=False)
        .tail(1)
    )
    return latest[latest["CumulativeQuantity"] > QUANTITY_EPSILON]


def _best_master_row_by_name(master: pd.DataFrame) -> dict[str, dict]:
    """
    InstrumentMaster can have multiple rows sharing the same
    NormalizedInstrument - a real stock's row plus related corporate-action
    rows (SUBSCRIPTION_RIGHT, TEMPORARY_SHARE, CORPORATE_ACTION) that share
    its display name but aren't meant to be looked up for pricing. A plain
    set_index().loc[] breaks on this (returns a DataFrame slice instead of a
    row, which normalise_text() then defensively - and silently - treats as
    blank, making a fully-classified instrument look unclassified: a real
    bug caught testing this exact function against real data). For each
    name, prefer the row with a PriceSymbol set; otherwise the first row
    with InstrumentType and Currency both set; otherwise just the first row.
    """
    best: dict[str, dict] = {}
    for _, row in master.iterrows():
        name = row["NormalizedInstrument"]
        row_dict = row.to_dict()
        current = best.get(name)

        if current is None:
            best[name] = row_dict
            continue

        current_has_symbol = bool(normalise_text(current.get("PriceSymbol", "")))
        row_has_symbol = bool(normalise_text(row_dict.get("PriceSymbol", "")))
        if row_has_symbol and not current_has_symbol:
            best[name] = row_dict
            continue

        current_classified = bool(normalise_text(current.get("InstrumentType", "")) and normalise_text(current.get("Currency", "")))
        row_classified = bool(normalise_text(row_dict.get("InstrumentType", "")) and normalise_text(row_dict.get("Currency", "")))
        if row_classified and not current_classified and not current_has_symbol:
            best[name] = row_dict

    return best


def find_coverage_gaps(
    positions_workbook: Path,
    instrument_master_path: Path,
    prices_workbook: Path,
    *,
    stale_days: int = DEFAULT_STALE_DAYS,
    as_of: date | None = None,
) -> dict[str, object]:
    active = active_instruments(positions_workbook)
    master = load_instrument_master(instrument_master_path)
    master_by_name = _best_master_row_by_name(master)

    prices = pd.read_excel(prices_workbook, sheet_name="InstrumentPrices", dtype=object) if prices_workbook.exists() else pd.DataFrame(columns=["NormalizedInstrument", "Date"])
    if len(prices) > 0:
        prices["Date"] = pd.to_datetime(prices["Date"])
    last_price_date_by_name = prices.groupby("NormalizedInstrument")["Date"].max().to_dict() if len(prices) > 0 else {}

    today = as_of or date.today()
    stale_cutoff = pd.Timestamp(today) - pd.Timedelta(days=stale_days)

    unclassified: list[str] = []
    unpriced: list[str] = []
    stale: list[tuple[str, str]] = []

    seen: set[str] = set()
    for instrument in active["NormalizedInstrument"].unique():
        if instrument in seen:
            continue
        seen.add(instrument)

        master_row = master_by_name.get(instrument)
        instrument_type = normalise_text(master_row.get("InstrumentType", "")) if master_row is not None else ""
        currency = normalise_text(master_row.get("Currency", "")) if master_row is not None else ""
        price_symbol = normalise_text(master_row.get("PriceSymbol", "")) if master_row is not None else ""

        if not instrument_type or not currency:
            unclassified.append(instrument)
            continue

        if not price_symbol:
            unpriced.append(instrument)
            continue

        last_price_date = last_price_date_by_name.get(instrument)
        if last_price_date is None or pd.isna(last_price_date):
            unpriced.append(instrument)
            continue

        if last_price_date < stale_cutoff:
            stale.append((instrument, last_price_date.strftime("%Y-%m-%d")))

    return {
        "active_instrument_count": len(seen),
        "unclassified": sorted(unclassified),
        "unpriced": sorted(unpriced),
        "stale": sorted(stale),
    }


def has_gaps(gaps: dict[str, object]) -> bool:
    return bool(gaps["unclassified"] or gaps["unpriced"] or gaps["stale"])


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Check that every currently-active investment position has complete data "
                    "(classification, price symbol, recent price) before it's trusted for net worth. "
                    "Exits non-zero if gaps are found, unless --force. Read-only - never writes anything."
    )
    parser.add_argument("--positions-workbook", default=str(DEFAULT_POSITIONS_WORKBOOK))
    parser.add_argument("--instrument-master", default=str(DEFAULT_INSTRUMENT_MASTER))
    parser.add_argument("--prices-workbook", default=str(DEFAULT_PRICES_WORKBOOK))
    parser.add_argument("--stale-days", type=int, default=DEFAULT_STALE_DAYS, help="Flag a price as stale if older than this many days.")
    parser.add_argument("--force", action="store_true", help="Report gaps but exit 0 anyway - proceed with incomplete data rather than block the pipeline.")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()

    positions_workbook = Path(args.positions_workbook).expanduser().resolve()
    instrument_master_path = Path(args.instrument_master).expanduser().resolve()
    prices_workbook = Path(args.prices_workbook).expanduser().resolve()

    gaps = find_coverage_gaps(
        positions_workbook,
        instrument_master_path,
        prices_workbook,
        stale_days=args.stale_days,
    )

    print(f"Active positions checked: {gaps['active_instrument_count']}")

    if not has_gaps(gaps):
        print("All active positions have complete data (classification, price symbol, recent price).")
        return

    print()
    if gaps["unclassified"]:
        print(f"UNCLASSIFIED ({len(gaps['unclassified'])}) - no InstrumentType/Currency in InstrumentMaster.xlsx:")
        for name in gaps["unclassified"]:
            print(f"  - {name}")

    if gaps["unpriced"]:
        print()
        print(f"UNPRICED ({len(gaps['unpriced'])}) - no PriceSymbol configured, or no price data ever fetched:")
        for name in gaps["unpriced"]:
            print(f"  - {name}")

    if gaps["stale"]:
        print()
        print(f"STALE ({len(gaps['stale'])}) - price data exists but hasn't updated recently:")
        for name, last_date in gaps["stale"]:
            print(f"  - {name}: last price {last_date}")

    print()
    if args.force:
        print(
            "--force set: proceeding anyway. Net worth figures for the instruments above will be "
            "incomplete or missing until resolved."
        )
        return

    print(
        "Resolve these in rules/investments/InstrumentMaster.xlsx (or fetch fresh prices) before "
        "trusting net worth output, or rerun with --force to proceed anyway."
    )
    sys.exit(1)


if __name__ == "__main__":
    main()
