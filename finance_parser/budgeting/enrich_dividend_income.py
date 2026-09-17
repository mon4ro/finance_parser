from __future__ import annotations

import argparse
from datetime import timedelta
from pathlib import Path

import pandas as pd

from finance_parser.common import clean_for_excel
from finance_parser.budgeting.parsers.investment_dividends import SYNTHETIC_BROKERS
from finance_parser.utilities.fresh_workbook_writer import (
    append_changelog_row,
    read_workbook_values_only,
    records_to_sheet_values,
    replace_with_fresh_workbook,
    sheet_values_to_records,
)

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_WORKBOOK = PROJECT_ROOT / "output" / "budgeting" / "ParsedTransactions.xlsx"
DEFAULT_DIVIDEND_HISTORY = PROJECT_ROOT / "output" / "investments" / "DividendHistory.xlsx"

UNIFIED_SHEET = "UnifiedTransactions"
DIVIDEND_HISTORY_SHEET = "DividendHistory"
SCRIPT_NAME = "enrich_dividend_income.py"

# Only OP-held instruments' dividends settle same-day into a real bank
# transaction this project can match against - confirmed against real data
# (e.g. a real Kesko B dividend landed in the budgeting side on the exact
# same date, exact net amount, as its investment-side dividend event).
# Nordnet/EVLI dividends sit in the broker's own cash balance instead (no
# matching bank transaction exists at all) and are handled separately by
# investment_dividends.py's synthetic entries, not this cross-reference.
MATCHABLE_BROKERS = {"OP"}

# Every real dividend this project has ever seen comes from one of these
# three brokers (confirmed against real DividendHistory.xlsx: exactly
# OP/NORDNET/EVLI, zero Nordea or Seligson events - Nordea's current
# holding is an accumulating fund with no cash distributions, and Seligson
# is direct-purchase only, no cash ever sits there). Not a hardcoded
# assumption though: find_uncovered_broker_dividends() below checks this
# against real data every run, so a real dividend from any OTHER broker -
# now or in the future - gets flagged loudly instead of silently dropped
# by both this script's MATCHABLE_BROKERS filter and
# investment_dividends.py's SYNTHETIC_BROKERS filter.
COVERED_BROKERS = MATCHABLE_BROKERS | SYNTHETIC_BROKERS

# Small settlement-lag tolerance for the (rare) case a dividend doesn't
# credit same-day - amount match stays exact, only the date gets slack.
DATE_TOLERANCE_DAYS = 3

DIVIDEND_SUPERCATEGORY = "INCOME"
DIVIDEND_CATEGORY = "Passive income"
DIVIDEND_SUBCATEGORY = "Dividend yields"

TARGET_FIELDS = ["Supercategory", "Category", "Subcategory", "Owner"]


def load_dividend_events(dividend_history_path: Path) -> pd.DataFrame:
    df = pd.read_excel(dividend_history_path, sheet_name=DIVIDEND_HISTORY_SHEET, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]

    df = df[df["Broker"].astype(str).isin(MATCHABLE_BROKERS)].copy()
    df["TradeDate"] = pd.to_datetime(df["TradeDate"], errors="coerce")
    df["NetDividendEUR"] = pd.to_numeric(df["NetDividendEUR"], errors="coerce").round(2)
    df = df.dropna(subset=["TradeDate", "NetDividendEUR"])
    df = df[df["NetDividendEUR"] > 0]

    return df


def find_uncovered_broker_dividends(dividend_history_path: Path) -> list[tuple[str, str, str, float]]:
    """
    Real dividend events from a broker that neither this script (OP-held
    instruments, matched against a real bank transaction) nor
    investment_dividends.py (Nordnet/EVLI-held instruments, synthesized
    directly) knows how to handle - see COVERED_BROKERS above. Returns
    (Broker, NormalizedInstrument, TradeDate, NetDividendEUR) tuples for
    reporting; never silently dropped.
    """
    df = pd.read_excel(dividend_history_path, sheet_name=DIVIDEND_HISTORY_SHEET, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]

    df["TradeDate"] = pd.to_datetime(df["TradeDate"], errors="coerce")
    df["NetDividendEUR"] = pd.to_numeric(df["NetDividendEUR"], errors="coerce").round(2)
    df = df.dropna(subset=["TradeDate", "NetDividendEUR"])
    df = df[df["NetDividendEUR"] > 0]

    uncovered = df[~df["Broker"].astype(str).isin(COVERED_BROKERS)]

    return [
        (row["Broker"], row.get("NormalizedInstrument", ""), row["TradeDate"].strftime("%Y-%m-%d"), float(row["NetDividendEUR"]))
        for _, row in uncovered.iterrows()
    ]


def _is_blank(value: object) -> bool:
    return str(value if value is not None else "").strip() == ""


def match_dividend_income(unified: pd.DataFrame, dividend_events: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, object]]:
    """
    Cross-reference real investment-side dividend events (the authoritative
    source, already verified against real broker data) against budgeting-
    side UnifiedTransactions by date + exact net amount, filling
    Supercategory/Category/Subcategory/Owner for the matched row - only into
    currently-blank cells, and only for rows a human hasn't already touched
    (Review/Notes blank), matching this project's existing categoriser
    safety conventions. Never creates rows, never overwrites a manual or
    rule-based value.
    """
    out = unified.copy()
    out["_Date"] = pd.to_datetime(out["Date"], errors="coerce")
    out["_Amount"] = pd.to_numeric(out["Amount"], errors="coerce").round(2)

    stats: dict[str, object] = {
        "dividend_events": len(dividend_events),
        "matched": 0,
        "matched_via_date_tolerance": 0,
        "fields_updated": 0,
        "skipped_review_notes": 0,
        "unmatched": [],
        "ambiguous": [],
    }

    for _, event in dividend_events.iterrows():
        target_amount = float(event["NetDividendEUR"])
        target_date = event["TradeDate"]
        instrument = event.get("NormalizedInstrument", "")

        same_amount = out[(out["_Amount"] == target_amount) & (out["_Amount"] > 0)]
        exact = same_amount[same_amount["_Date"] == target_date]

        if len(exact) == 1:
            candidates = exact
        elif len(exact) == 0:
            window = same_amount[
                (same_amount["_Date"] >= target_date - timedelta(days=DATE_TOLERANCE_DAYS))
                & (same_amount["_Date"] <= target_date + timedelta(days=DATE_TOLERANCE_DAYS))
            ]
            if len(window) == 1:
                candidates = window
                stats["matched_via_date_tolerance"] += 1
            elif len(window) == 0:
                stats["unmatched"].append((instrument, target_date.strftime("%Y-%m-%d"), target_amount))
                continue
            else:
                stats["ambiguous"].append((instrument, target_date.strftime("%Y-%m-%d"), target_amount, len(window)))
                continue
        else:
            stats["ambiguous"].append((instrument, target_date.strftime("%Y-%m-%d"), target_amount, len(exact)))
            continue

        idx = candidates.index[0]

        review_notes = out.at[idx, "Review/Notes"] if "Review/Notes" in out.columns else ""
        if not _is_blank(review_notes):
            stats["skipped_review_notes"] += 1
            continue

        new_values = {
            "Supercategory": DIVIDEND_SUPERCATEGORY,
            "Category": DIVIDEND_CATEGORY,
            "Subcategory": DIVIDEND_SUBCATEGORY,
            "Owner": event.get("PortfolioOwner", ""),
        }
        row_changed = False
        for field, value in new_values.items():
            if field not in out.columns or _is_blank(value):
                continue
            if _is_blank(out.at[idx, field]):
                out.at[idx, field] = value
                stats["fields_updated"] += 1
                row_changed = True

        if row_changed:
            stats["matched"] += 1

    out = out.drop(columns=["_Date", "_Amount"])
    return out, stats


def dataframe_to_records(df: pd.DataFrame) -> list[dict[str, object]]:
    records: list[dict[str, object]] = []
    for _, row in df.iterrows():
        record: dict[str, object] = {}
        for col in df.columns:
            value = row.get(col, "")
            if pd.isna(value):
                value = ""
            record[str(col)] = clean_for_excel(value)
        records.append(record)
    return records


def enrich_dividend_income(
    workbook_path: Path,
    dividend_history_path: Path,
    *,
    dry_run: bool = False,
) -> dict[str, object]:
    sheets = read_workbook_values_only(workbook_path)
    if UNIFIED_SHEET not in sheets:
        raise ValueError(f"Workbook does not contain required sheet: {UNIFIED_SHEET}")

    headers, records = sheet_values_to_records(sheets[UNIFIED_SHEET])
    unified = pd.DataFrame(records, columns=headers)

    dividend_events = load_dividend_events(dividend_history_path)
    updated, stats = match_dividend_income(unified, dividend_events)
    stats["uncovered_broker_dividends"] = find_uncovered_broker_dividends(dividend_history_path)

    if dry_run or stats["fields_updated"] == 0:
        return stats

    new_headers = [str(col) for col in updated.columns]
    new_records = dataframe_to_records(updated)
    sheets[UNIFIED_SHEET] = records_to_sheet_values(new_headers, new_records)

    append_changelog_row(
        sheets,
        script=SCRIPT_NAME,
        action="Cross-reference real dividend income against investment DividendHistory",
        sheet=UNIFIED_SHEET,
        rows_updated=stats["matched"],
        status="Completed",
        details=(
            f"Dividend events checked: {stats['dividend_events']}; matched: {stats['matched']}; "
            f"fields updated: {stats['fields_updated']}; unmatched: {len(stats['unmatched'])}; "
            f"ambiguous: {len(stats['ambiguous'])}; skipped (Review/Notes set): {stats['skipped_review_notes']}"
        ),
        backup_file=None,
    )

    backup_path, _, writer_stats = replace_with_fresh_workbook(
        workbook_path,
        sheets,
        backup_label="before_dividend_income_enrich",
        basic_formatting=True,
        excel_tables=False,
    )
    stats["backup_path"] = str(backup_path)
    stats["sheets_written"] = writer_stats["sheets_written"]

    return stats


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Cross-reference real OP-held-instrument dividend events (investment-side "
                    "DividendHistory.xlsx) against budgeting-side UnifiedTransactions by date + net "
                    "amount, filling Supercategory/Category/Subcategory/Owner for matched rows - only "
                    "into blank cells, never overwriting a manual or rule-based value."
    )
    parser.add_argument("--workbook", default=str(DEFAULT_WORKBOOK))
    parser.add_argument("--dividend-history", default=str(DEFAULT_DIVIDEND_HISTORY))
    parser.add_argument("--dry-run", action="store_true")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()
    workbook_path = Path(args.workbook).expanduser().resolve()
    dividend_history_path = Path(args.dividend_history).expanduser().resolve()

    stats = enrich_dividend_income(workbook_path, dividend_history_path, dry_run=args.dry_run)

    print("Dividend income enrichment dry run complete." if args.dry_run else "Dividend income enrichment complete.")
    print(f"Dividend events checked (OP-held only): {stats['dividend_events']}")
    print(f"Matched:                                {stats['matched']}")
    print(f"  of which via date-tolerance match:    {stats['matched_via_date_tolerance']}")
    print(f"Fields updated:                         {stats['fields_updated']}")
    print(f"Skipped (Review/Notes already set):     {stats['skipped_review_notes']}")

    if stats["unmatched"]:
        print()
        print(f"Unmatched dividend events ({len(stats['unmatched'])}) - no budgeting-side row found within {DATE_TOLERANCE_DAYS} days:")
        for instrument, date, amount in stats["unmatched"]:
            print(f"  - {instrument}: {date}, {amount} EUR")

    if stats["ambiguous"]:
        print()
        print(f"Ambiguous dividend events ({len(stats['ambiguous'])}) - more than one candidate row, skipped:")
        for instrument, date, amount, count in stats["ambiguous"]:
            print(f"  - {instrument}: {date}, {amount} EUR ({count} candidates)")

    if stats["uncovered_broker_dividends"]:
        print()
        print(
            f"WARNING: {len(stats['uncovered_broker_dividends'])} real dividend event(s) from a broker "
            f"neither this script nor investment_dividends.py currently covers (COVERED_BROKERS = "
            f"{sorted(COVERED_BROKERS)}):"
        )
        for broker, instrument, date, amount in stats["uncovered_broker_dividends"]:
            print(f"  - {broker} / {instrument}: {date}, {amount} EUR")
        print("  These are not reflected in budgeting-side income at all - add this broker to")
        print("  MATCHABLE_BROKERS (if it settles same-day into a real bank transaction, like OP) or")
        print("  SYNTHETIC_BROKERS in investment_dividends.py (if it doesn't, like Nordnet/EVLI).")

    if args.dry_run:
        print()
        print("Dry run only: workbook was not modified.")
    elif stats["fields_updated"] == 0:
        print()
        print("No changes to write.")
    else:
        print()
        print(f"Output workbook: {workbook_path}")


if __name__ == "__main__":
    main()
