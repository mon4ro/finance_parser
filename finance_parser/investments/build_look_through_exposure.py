from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from finance_parser.common import normalise_header, normalise_text
from finance_parser.investments.build_monthly_position_value import (
    CASH_INSTRUMENT_LABEL,
    MONTHLY_VALUE_SHEET,
)
from finance_parser.investments.check_instrument_coverage import _best_master_row_by_name
from finance_parser.investments.fetch_fund_holdings import (
    HOLDINGS_COLUMNS,
    HOLDINGS_SHEET,
    UNKNOWN_NO_DATA_NAME,
    UNKNOWN_SYMBOL,
)
from finance_parser.investments.investment_common import load_instrument_master
from finance_parser.utilities.fresh_workbook_writer import (
    read_workbook_values_only,
    records_to_sheet_values,
    replace_with_fresh_workbook,
    write_fresh_workbook,
)


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_INSTRUMENT_MASTER = PROJECT_ROOT / "rules" / "investments" / "InstrumentMaster.xlsx"
DEFAULT_MONTHLY_VALUE_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "MonthlyPositionValue.xlsx"
DEFAULT_HOLDINGS_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "FundHoldingsSnapshot.xlsx"
DEFAULT_OUTPUT_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "LookThroughExposure.xlsx"

DETAIL_SHEET = "LookThroughDetail"
SUMMARY_SHEET = "LookThroughByCompany"
FUND_COVERAGE_SHEET = "FundCoverage"

# Per-source-instrument breakdown, before aggregation - e.g. how much of the
# total Nokia exposure came from direct stock vs from OP-Suomi A specifically.
DETAIL_COLUMNS = [
    "SnapshotMonth",
    "SourceInstrument",
    "SourceInstrumentType",
    "SourceMarketValueEUR",
    "HoldingKey",
    "HoldingName",
    "HoldingPercent",
    "ExposureEUR",
]

# One row per real-world company/holding, combining direct + indirect exposure.
SUMMARY_COLUMNS = ["SnapshotMonth", "HoldingKey", "HoldingName", "ExposureEUR", "ExposurePercent"]

# One row per currently-held fund: how much of IT could actually be mapped to
# a named holding vs left as an UNKNOWN residual (top-10 cutoff, a
# fund-of-funds Yahoo can't see into, or no data at all). Answers "which of
# my funds are poorly covered" directly, without having to notice it inside
# the (necessarily much larger, and per-fund-scoped) LookThroughByCompany
# UNKNOWN rows.
FUND_COVERAGE_COLUMNS = ["SnapshotMonth", "NormalizedInstrument", "KnownPercent", "UnknownPercent"]


def _month_key(value: object) -> str:
    ts = pd.to_datetime(value)
    return f"{ts.year:04d}-{ts.month:02d}"


def _month_ordinal(month_str: str) -> int:
    year, month = (int(part) for part in month_str.split("-"))
    return year * 12 + month


def load_position_values(monthly_value_workbook: Path, as_of_month: str | None = None) -> tuple[str, pd.DataFrame]:
    """
    Aggregate MarketValueEUR to one row per NormalizedInstrument for a single
    month - the look-through calculation only cares about total exposure per
    security, not which broker/account holds it. Excludes CASH (a broker's
    unreinvested cash balance isn't exposure to any company) from both the
    returned rows and the total-portfolio-value denominator used later for
    percentages.
    """
    df = pd.read_excel(monthly_value_workbook, sheet_name=MONTHLY_VALUE_SHEET, dtype=object, engine="openpyxl")
    df.columns = [normalise_header(c) for c in df.columns]

    df["_MonthKey"] = df["MonthEnd"].map(_month_key)
    df["NormalizedInstrument"] = df["NormalizedInstrument"].map(normalise_text)
    df["MarketValueEUR"] = pd.to_numeric(df["MarketValueEUR"], errors="coerce").fillna(0.0)

    df = df[df["NormalizedInstrument"] != normalise_text(CASH_INSTRUMENT_LABEL)]

    target_month = as_of_month or df["_MonthKey"].max()
    month_rows = df[df["_MonthKey"] == target_month]

    grouped = month_rows.groupby("NormalizedInstrument", as_index=False)["MarketValueEUR"].sum()
    grouped = grouped[grouped["MarketValueEUR"] > 0].reset_index(drop=True)

    return target_month, grouped


def load_fund_holdings(holdings_workbook: Path) -> pd.DataFrame:
    if not holdings_workbook.exists():
        return pd.DataFrame(columns=HOLDINGS_COLUMNS)

    df = pd.read_excel(holdings_workbook, sheet_name=HOLDINGS_SHEET, dtype=object, engine="openpyxl")
    df.columns = [normalise_header(c) for c in df.columns]
    df["NormalizedInstrument"] = df["NormalizedInstrument"].map(normalise_text)
    df["HoldingPercent"] = pd.to_numeric(df["HoldingPercent"], errors="coerce").fillna(0.0)
    return df


def _nearest_snapshot_rows(fund_snapshots: pd.DataFrame, target_month: str) -> pd.DataFrame:
    """
    Prefer an exact SnapshotMonth match; otherwise fall back to whichever
    month is closest by distance. The 60-month assumed-constant backfill
    (fetch_fund_holdings.py) should make an exact match the common case -
    this only matters for a target month outside that backfilled window.
    """
    exact = fund_snapshots[fund_snapshots["SnapshotMonth"] == target_month]
    if len(exact) > 0:
        return exact

    target_ord = _month_ordinal(target_month)
    with_distance = fund_snapshots.copy()
    with_distance["_Distance"] = with_distance["SnapshotMonth"].map(lambda m: abs(_month_ordinal(m) - target_ord))
    nearest_month = with_distance.sort_values("_Distance").iloc[0]["SnapshotMonth"]
    return fund_snapshots[fund_snapshots["SnapshotMonth"] == nearest_month]


def _direct_holding_key(instrument: str, master_row: dict | None) -> tuple[str, str]:
    """
    A directly-held stock keys on its own PriceSymbol when configured - the
    same Yahoo ticker fund holdings report their constituents under (e.g.
    InstrumentMaster's NOKIA -> PriceSymbol NOKIA.HE matches exactly what a
    fund's topHoldings reports Nokia as) - so direct + fund-indirect exposure
    to the same real company merge into one combined total instead of two
    separate buckets. Falls back to NormalizedInstrument when no PriceSymbol
    is configured.
    """
    price_symbol = normalise_text(master_row.get("PriceSymbol", "")) if master_row else ""
    if price_symbol:
        return price_symbol, instrument
    return instrument, instrument


def build_look_through_exposure(
    instrument_master_path: Path,
    monthly_value_workbook: Path,
    holdings_workbook: Path,
    *,
    as_of_month: str | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, dict[str, object]]:
    target_month, positions = load_position_values(monthly_value_workbook, as_of_month)

    master = load_instrument_master(instrument_master_path)
    master_by_name = _best_master_row_by_name(master)

    holdings = load_fund_holdings(holdings_workbook)
    total_value = float(positions["MarketValueEUR"].sum())

    stats: dict[str, object] = {
        "snapshot_month": target_month,
        "total_portfolio_value_eur": total_value,
        "instruments_processed": len(positions),
        "funds_missing_holdings_data": [],
        "unclassified_instruments_treated_as_direct": [],
    }

    detail_rows: list[dict[str, object]] = []
    fund_coverage_rows: list[dict[str, object]] = []

    for _, row in positions.iterrows():
        instrument = row["NormalizedInstrument"]
        value = float(row["MarketValueEUR"])

        master_row = master_by_name.get(instrument)
        instrument_type = normalise_text(master_row.get("InstrumentType", "")) if master_row else ""

        if not instrument_type:
            stats["unclassified_instruments_treated_as_direct"].append(instrument)

        if instrument_type == "FUND":
            fund_snapshots = holdings[holdings["NormalizedInstrument"] == instrument] if len(holdings) > 0 else holdings
            if len(fund_snapshots) == 0:
                stats["funds_missing_holdings_data"].append(instrument)
                rows_for_month = pd.DataFrame([{
                    "HoldingSymbol": UNKNOWN_SYMBOL,
                    "HoldingName": UNKNOWN_NO_DATA_NAME,
                    "HoldingPercent": 100.0,
                }])
            else:
                rows_for_month = _nearest_snapshot_rows(fund_snapshots, target_month)

            unknown_percent = 0.0
            for _, holding_row in rows_for_month.iterrows():
                percent = float(holding_row["HoldingPercent"])
                raw_key = normalise_text(holding_row.get("HoldingSymbol", "")) or UNKNOWN_SYMBOL
                raw_name = normalise_text(holding_row.get("HoldingName", "")) or raw_key

                if raw_key == UNKNOWN_SYMBOL:
                    # Scope the UNKNOWN bucket to its source fund rather than
                    # sharing one global "UNKNOWN" key across every fund -
                    # otherwise a fund with literally zero holdings data
                    # (e.g. a real one this project has seen: Yahoo returns
                    # nothing for it at all) looks identical to a
                    # well-covered fund's ordinary top-10 cutoff tail, once
                    # merged into a single bucket in the summary. See also
                    # the FundCoverage sheet, which reports this per fund
                    # directly.
                    key = f"UNKNOWN:{instrument}"
                    name = f"Unknown / unmapped ({instrument})"
                    unknown_percent += percent
                else:
                    key, name = raw_key, raw_name

                detail_rows.append({
                    "SnapshotMonth": target_month,
                    "SourceInstrument": instrument,
                    "SourceInstrumentType": "FUND",
                    "SourceMarketValueEUR": value,
                    "HoldingKey": key,
                    "HoldingName": name,
                    "HoldingPercent": percent,
                    "ExposureEUR": round(value * percent / 100.0, 2),
                })

            fund_coverage_rows.append({
                "SnapshotMonth": target_month,
                "NormalizedInstrument": instrument,
                "KnownPercent": round(100.0 - unknown_percent, 4),
                "UnknownPercent": round(unknown_percent, 4),
            })
        else:
            key, name = _direct_holding_key(instrument, master_row)
            detail_rows.append({
                "SnapshotMonth": target_month,
                "SourceInstrument": instrument,
                "SourceInstrumentType": instrument_type or "UNCLASSIFIED",
                "SourceMarketValueEUR": value,
                "HoldingKey": key,
                "HoldingName": name,
                "HoldingPercent": 100.0,
                "ExposureEUR": value,
            })

    detail = pd.DataFrame(detail_rows, columns=DETAIL_COLUMNS)

    summary = (
        detail.groupby(["SnapshotMonth", "HoldingKey"], as_index=False)
        .agg(HoldingName=("HoldingName", "first"), ExposureEUR=("ExposureEUR", "sum"))
    )
    summary["ExposurePercent"] = summary["ExposureEUR"].map(
        lambda v: round(v / total_value * 100.0, 4) if total_value else 0.0
    )
    summary = summary.sort_values("ExposureEUR", ascending=False).reset_index(drop=True)
    summary = summary[SUMMARY_COLUMNS]

    fund_coverage = pd.DataFrame(fund_coverage_rows, columns=FUND_COVERAGE_COLUMNS)
    fund_coverage = fund_coverage.sort_values("UnknownPercent", ascending=False).reset_index(drop=True)

    return detail, summary, fund_coverage, stats


def write_look_through_workbook(
    output_workbook: Path,
    detail: pd.DataFrame,
    summary: pd.DataFrame,
    fund_coverage: pd.DataFrame,
) -> None:
    if output_workbook.exists():
        sheets = read_workbook_values_only(output_workbook)
    else:
        sheets = {}

    sheets[SUMMARY_SHEET] = records_to_sheet_values(SUMMARY_COLUMNS, summary.to_dict("records"))
    sheets[DETAIL_SHEET] = records_to_sheet_values(DETAIL_COLUMNS, detail.to_dict("records"))
    sheets[FUND_COVERAGE_SHEET] = records_to_sheet_values(FUND_COVERAGE_COLUMNS, fund_coverage.to_dict("records"))

    if output_workbook.exists():
        replace_with_fresh_workbook(
            output_workbook,
            sheets,
            backup_label="before_look_through_rebuild",
            basic_formatting=True,
            excel_tables=False,
        )
    else:
        write_fresh_workbook(output_workbook, sheets, basic_formatting=True, excel_tables=False)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Compute total look-through investment exposure per company/security, combining "
            "direct stock holdings with indirect exposure via funds' own constituent holdings "
            "(output/investments/FundHoldingsSnapshot.xlsx). Read-only analysis, on-demand - "
            "not part of the main investment pipeline, since nothing downstream depends on it."
        )
    )
    parser.add_argument("--instrument-master", default=str(DEFAULT_INSTRUMENT_MASTER))
    parser.add_argument("--monthly-value-workbook", default=str(DEFAULT_MONTHLY_VALUE_WORKBOOK))
    parser.add_argument("--holdings-workbook", default=str(DEFAULT_HOLDINGS_WORKBOOK))
    parser.add_argument("--output-workbook", default=str(DEFAULT_OUTPUT_WORKBOOK))
    parser.add_argument(
        "--month",
        default=None,
        help="Override the snapshot month (YYYY-MM). Defaults to the latest month in MonthlyPositionValue.xlsx.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Compute and report, but do not write the workbook.")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()

    instrument_master_path = Path(args.instrument_master).expanduser().resolve()
    monthly_value_workbook = Path(args.monthly_value_workbook).expanduser().resolve()
    holdings_workbook = Path(args.holdings_workbook).expanduser().resolve()
    output_workbook = Path(args.output_workbook).expanduser().resolve()

    detail, summary, fund_coverage, stats = build_look_through_exposure(
        instrument_master_path,
        monthly_value_workbook,
        holdings_workbook,
        as_of_month=args.month,
    )

    print()
    print("Look-through exposure calculation complete." if not args.dry_run else "Look-through exposure dry run complete.")
    print(f"Snapshot month:                 {stats['snapshot_month']}")
    print(f"Total portfolio value (EUR):    {stats['total_portfolio_value_eur']:.2f}")
    print(f"Instruments processed:          {stats['instruments_processed']}")
    print(f"Distinct look-through holdings: {len(summary)}")

    if stats["unclassified_instruments_treated_as_direct"]:
        print()
        print("Unclassified instruments (no InstrumentType in InstrumentMaster) - treated as direct exposure to themselves:")
        for name in stats["unclassified_instruments_treated_as_direct"]:
            print(f"  - {name}")

    if stats["funds_missing_holdings_data"]:
        print()
        print("Funds with no holdings snapshot data available (run fetch_fund_holdings.py first) - counted as 100% UNKNOWN:")
        for name in stats["funds_missing_holdings_data"]:
            print(f"  - {name}")

    print()
    print("Top 10 look-through exposures:")
    for _, row in summary.head(10).iterrows():
        print(f"  {row['HoldingName']:<40} {row['ExposureEUR']:>12.2f} EUR  ({row['ExposurePercent']:.2f}%)")

    if len(fund_coverage) > 0:
        print()
        print("Fund coverage (share of each fund's own value that could NOT be mapped to a named holding):")
        for _, row in fund_coverage.iterrows():
            print(f"  {row['NormalizedInstrument']:<30} {row['UnknownPercent']:>7.2f}% unknown")

    if args.dry_run:
        print()
        print("Dry run only: workbook was not modified.")
        return

    write_look_through_workbook(output_workbook, detail, summary, fund_coverage)
    print()
    print(f"Output workbook: {output_workbook}")


if __name__ == "__main__":
    main()
