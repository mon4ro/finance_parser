from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from finance_parser.common import normalise_header, normalise_text
from finance_parser.settings import get_settings
from finance_parser.utilities.fresh_workbook_writer import (
    records_to_sheet_values,
    replace_with_fresh_workbook,
    write_fresh_workbook,
)


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_INVESTMENTS_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "ParsedInvestments.xlsx"
DEFAULT_POSITIONS_WORKBOOK = PROJECT_ROOT / "output" / "investments" / "PortfolioPositions.xlsx"
DEFAULT_INSTRUMENT_MASTER = PROJECT_ROOT / "rules" / "investments" / "InstrumentMaster.xlsx"

OPENING_POSITIONS_SHEET = "OpeningPositions"

POSITIONS_SHEET = "PortfolioPositions"
POSITIONS_COLUMNS = [
    "Broker",
    "Portfolio",
    "PortfolioOwner",
    "PortfolioType",
    "NormalizedInstrument",
    "Date",
    "TransactionType",
    "QuantityDelta",
    "CumulativeQuantity",
    "CashDelta",
    "CumulativeNetInvested",
]

# Quantity effect classification. Real InvestmentTransactions data (2026-09)
# contains ~30 distinct TransactionType values; only the ones below are
# confidently understood. Anything else is deliberately excluded from the
# quantity sum and reported by build_positions() as "unhandled" rather than
# guessed - see docs/patch_notes or ask the user before extending this.
#
# "Delivered in" vs "delivered out" sign for Nordnet's JÄTTÖ/OTTO corporate
# action pairs was confirmed against real reverse-split data (Norwegian Air
# Shuttle, Finnair): JÄTTÖ = shares delivered in (add), OTTO = shares
# delivered out (subtract) - e.g. a 100:1 reverse split shows
# "SPLIT AP JÄTTÖ" with the small post-split count and "SPLIT AP OTTO" with
# the large pre-split count being given up.
QUANTITY_ADD_TYPES = {
    "BUY", "ALLOCATED", "DELIVERY", "MATCHING",
    "VAIHTO AP-JÄTTÖ", "VAIHTO - JÄTTÖ",
    # Lehto 2020 rights issue: only the actual share-issuance step counts -
    # the payment/rights-attachment/rights-removal steps around it are about
    # the *rights*, not real share count, and are left unmapped (neutral) to
    # avoid double-counting the same corporate action twice.
    "MO OTTO EMISSION YHT",
    # Manually-seeded pre-tracking-period holdings (see load_opening_positions()
    # and InstrumentMaster.xlsx's OpeningPositions sheet) - not a real row from
    # any broker export. A plain add is correct here since these are always
    # synthesized as the very first row for their (Broker, Portfolio,
    # NormalizedInstrument) group, so running_total is 0 going in.
    "OPENING BALANCE",
}
QUANTITY_SUBTRACT_TYPES = {
    "SELL", "REDEMPTION",
    "VAIHTO AP-OTTO", "VAIHTO - OTTO",
}

# A stock split is NOT a delta event - it's a ratio replacement. Confirmed
# against two real cases: Norwegian Air Shuttle and Finnair reverse splits.
# "SPLIT AP JÄTTÖ"'s Quantity is the actual resulting share count (verified
# against Finnair: 120 shares x real market price landed exactly in the
# user's remembered real portfolio value range). Treating it as an ADD delta
# only worked for Norwegian Air Shuttle by coincidence (its paired "OTTO"
# magnitude happened to exactly equal the pre-split running total already
# computed here); it broke for Finnair, where extra 2023 rights-issue
# activity meant the OTTO figure referenced a different base than what this
# script was tracking, producing an impossible negative share count.
QUANTITY_RESET_TYPES = {"SPLIT AP JÄTTÖ"}
# The paired "OTTO" side of a split is already fully accounted for by the
# JÄTTÖ reset above - explicitly neutral, not "unhandled".
KNOWN_NEUTRAL_SPLIT_PAIR_TYPES = {"SPLIT AP OTTO"}

# Cash-flow "net invested" classification: real buy/sell cash flows, plus
# the same quantity-add events that can carry an unrecorded (CashAmount=0)
# cost basis - OPENING BALANCE (see load_opening_positions()) and the
# VAIHTO - JÄTTÖ / VAIHTO AP-JÄTTÖ transfer-in events. All default to 0.0
# (unknown cost basis, same as before) unless manually corrected - via
# OpeningPositions' own CashAmount column, or for a real already-imported
# transaction, correct_transaction_cash_amount.py - in which case the
# corrected amount must count the same as an actual purchase would.
# Dividends, interest, fees, tax, and deposits/withdrawals are deliberately
# excluded - they don't represent money invested in the instrument itself.
CASH_FLOW_TYPES = {"BUY", "SELL", "OPENING BALANCE", "VAIHTO - JÄTTÖ", "VAIHTO AP-JÄTTÖ"}

# Confidently understood as having NO effect on quantity or invested cash -
# real cash-flow/informational events, not a gap in coverage. Kept separate
# from "unhandled" reporting so that report only flags genuinely uncertain
# TransactionType values, not deliberately-neutral ones.
KNOWN_NEUTRAL_TYPES = {
    "DIVIDEND", "INTEREST", "FEE", "TAX", "ENNAKKOPIDÄTYS",
    "DEPOSIT", "WITHDRAWAL", "TALLETUS OST.",
} | KNOWN_NEUTRAL_SPLIT_PAIR_TYPES


def imported_at_now() -> str:
    from datetime import datetime
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def load_transactions(investments_workbook: Path) -> pd.DataFrame:
    df = pd.read_excel(investments_workbook, sheet_name="InvestmentTransactions", dtype=object, engine="openpyxl")
    df.columns = [normalise_header(c) for c in df.columns]

    df["NormalizedInstrument"] = df["NormalizedInstrument"].map(normalise_text)
    df = df[df["NormalizedInstrument"] != ""].copy()

    df["_TradeDate"] = pd.to_datetime(df["TradeDate"], errors="coerce")
    df = df[df["_TradeDate"].notna()].copy()

    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0.0)
    df["CashAmount"] = pd.to_numeric(df["CashAmount"], errors="coerce").fillna(0.0)
    df["Broker"] = df["Broker"].map(normalise_text)
    df["Portfolio"] = df["Portfolio"].map(normalise_text)
    df["TransactionType"] = df["TransactionType"].map(normalise_text)

    if "PortfolioOwner" not in df.columns:
        df["PortfolioOwner"] = ""
    df["PortfolioOwner"] = df["PortfolioOwner"].map(normalise_text)

    if "PortfolioType" not in df.columns:
        df["PortfolioType"] = ""
    df["PortfolioType"] = df["PortfolioType"].map(normalise_text)

    return df


_OPENING_POSITIONS_COLUMNS = ["Broker", "Portfolio", "NormalizedInstrument", "_TradeDate", "Quantity", "CashAmount", "TransactionType"]


def load_opening_positions(instrument_master_path: Path) -> pd.DataFrame:
    """
    Manually-seeded pre-tracking-period holdings from InstrumentMaster.xlsx's
    OpeningPositions sheet - real quantities the user remembers holding
    (e.g. small fund positions received as gifts before tracking began) with
    no corresponding row in any parsed broker export. Returns an empty,
    correctly-shaped DataFrame if the sheet doesn't exist yet - this is an
    optional, manually-maintained supplement, not a required input.
    """
    if not instrument_master_path.exists():
        return pd.DataFrame(columns=_OPENING_POSITIONS_COLUMNS)

    try:
        df = pd.read_excel(instrument_master_path, sheet_name=OPENING_POSITIONS_SHEET, dtype=object, engine="openpyxl")
    except ValueError:
        return pd.DataFrame(columns=_OPENING_POSITIONS_COLUMNS)

    if df.empty:
        return pd.DataFrame(columns=_OPENING_POSITIONS_COLUMNS)

    df.columns = [normalise_header(c) for c in df.columns]
    # NormalizedInstrument/Broker/Portfolio must match the uppercase form used
    # everywhere else in the pipeline (see InstrumentMaster's own
    # NormalizedInstrument column) - .upper() here so a manually-typed mixed-case
    # entry in the Excel sheet doesn't silently fail to join.
    df["Broker"] = df["Broker"].map(normalise_text).str.upper()
    df["Portfolio"] = df["Portfolio"].map(normalise_text).str.upper()
    df["NormalizedInstrument"] = df["NormalizedInstrument"].map(normalise_text).str.upper()
    df["_TradeDate"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0.0)
    # Honor a manually-supplied cost basis if the sheet has one (e.g. a
    # gift's tax-assessed value, treated as if that amount had been paid in
    # cash for the position) - defaults to 0.0 (unknown cost basis) only
    # when the sheet has no CashAmount column at all or leaves it blank for
    # a given row, matching every other "manual override, never guessed"
    # pattern in this project.
    if "CashAmount" not in df.columns:
        df["CashAmount"] = 0.0
    df["CashAmount"] = pd.to_numeric(df["CashAmount"], errors="coerce").fillna(0.0)
    df["TransactionType"] = "OPENING BALANCE"

    return df[_OPENING_POSITIONS_COLUMNS]


def load_average_cost_instruments(instrument_master_path: Path) -> set[str]:
    """
    NormalizedInstrument names opted into average-cost-basis tracking via
    InstrumentMaster.xlsx's own CostBasisMethod column ("AVERAGE"; blank or
    anything else keeps the default cash-flow cumsum behavior). Deliberately
    an explicit opt-in per instrument, not a blanket policy: correct for a
    savings-plan-style fund (regular contributions, eventual redemption -
    average cost is the standard method), but wrong for an actively-traded
    stock bought and sold many times, or an incentive-plan holding with
    partial tax-withholding sales - those need FIFO/specific-lot matching to
    track correctly, which is a deliberately separate, deferred feature.
    """
    if not instrument_master_path.exists():
        return set()

    try:
        df = pd.read_excel(instrument_master_path, sheet_name="InstrumentMaster", dtype=object, engine="openpyxl")
    except ValueError:
        return set()

    df.columns = [normalise_header(c) for c in df.columns]
    if "CostBasisMethod" not in df.columns:
        return set()

    is_average = df["CostBasisMethod"].map(normalise_text).str.upper() == "AVERAGE"
    return set(df.loc[is_average, "NormalizedInstrument"].map(normalise_text).str.upper())


def apply_average_cost_basis(active: pd.DataFrame, group_cols: list[str], average_cost_instruments: set[str]) -> pd.Series:
    """
    Walk each (Broker, Portfolio, NormalizedInstrument) group chronologically
    with a running cost-basis total, so that for an opted-in instrument a
    quantity-reducing event (a SELL, or any QUANTITY_SUBTRACT_TYPES) reduces
    the running cost basis proportionally to the fraction of the position
    removed, instead of adding that event's cash delta on top (the default
    cash-flow model). This always yields a correct "remaining cost basis" /
    "unrealized gain on what's still held" for the opted-in instrument, no
    matter how many buy/sell cycles happen - it does not (and isn't meant
    to) compute realized gain/loss for any specific historical sale, which
    needs FIFO/specific-lot matching and is a deliberately separate,
    deferred feature. A non-opted-in instrument's CashDelta passes through
    completely unchanged. Requires `active` to already be sorted by
    group_cols + date, and CumulativeQuantity to already be computed
    (build_positions() guarantees both).
    """
    resolved = active["CashDelta"].copy()
    running_invested = 0.0
    previous_quantity = 0.0
    current_group = None

    for idx, row in active.iterrows():
        group_key = tuple(row[c] for c in group_cols)
        if group_key != current_group:
            current_group = group_key
            running_invested = 0.0
            previous_quantity = 0.0

        quantity_after = row["CumulativeQuantity"]
        is_reduction = row["QuantityDelta"] < 0 and previous_quantity > 1e-9

        if row["NormalizedInstrument"] in average_cost_instruments and is_reduction:
            remaining_fraction = max(quantity_after, 0.0) / previous_quantity
            new_invested = running_invested * remaining_fraction
            resolved.at[idx] = new_invested - running_invested
            running_invested = new_invested
        else:
            resolved.at[idx] = row["CashDelta"]
            running_invested += row["CashDelta"]

        previous_quantity = quantity_after

    return resolved


def classify_quantity_delta(row: pd.Series) -> float:
    """
    Naive per-row delta, ignoring reset semantics - used only as the
    starting point before apply_quantity_resets() rewrites the reset rows.
    """
    tt = row["TransactionType"]
    magnitude = abs(row["Quantity"])

    if tt in QUANTITY_ADD_TYPES:
        return magnitude
    if tt in QUANTITY_SUBTRACT_TYPES:
        return -magnitude
    return 0.0


def apply_quantity_resets(active: pd.DataFrame, group_cols: list[str]) -> pd.Series:
    """
    Walk each (Broker, Portfolio, NormalizedInstrument) group chronologically
    with an explicit running total, so that a QUANTITY_RESET_TYPES row (a
    stock split) can be turned into "whatever delta makes the running total
    land exactly on this row's real resulting share count" - not simply
    added on top of what came before. Requires `active` to already be sorted
    by group_cols + date (build_positions() guarantees this).
    """
    resolved_deltas = active["QuantityDelta"].copy()
    running_total = 0.0
    current_group = None

    for idx, row in active.iterrows():
        group_key = tuple(row[c] for c in group_cols)
        if group_key != current_group:
            current_group = group_key
            running_total = 0.0

        if row["TransactionType"] in QUANTITY_RESET_TYPES:
            target = abs(row["Quantity"])
            resolved_deltas.at[idx] = target - running_total
            running_total = target
        else:
            running_total += row["QuantityDelta"]

    return resolved_deltas


def classify_cash_delta(row: pd.Series) -> float:
    if row["TransactionType"] in CASH_FLOW_TYPES:
        return row["CashAmount"]
    return 0.0


def build_positions(
    investments_workbook: Path,
    instrument_master: Path = DEFAULT_INSTRUMENT_MASTER,
) -> tuple[pd.DataFrame, dict[str, object]]:
    df = load_transactions(investments_workbook)

    opening = load_opening_positions(instrument_master)
    if not opening.empty:
        df = pd.concat([df, opening], ignore_index=True, sort=False)

    # Opening-position rows (seeded straight from InstrumentMaster, not a
    # parsed broker export - see load_opening_positions()) have no
    # PortfolioOwner of their own. Real InvestmentTransactions rows should
    # already carry it (apply_portfolio_ownership() backfills the whole
    # history on every investment_parser.py run), but derive it here too
    # from settings.yaml wherever it's still blank, so every position row
    # ends up attributed regardless of source.
    if "PortfolioOwner" not in df.columns:
        df["PortfolioOwner"] = ""
    df["PortfolioOwner"] = df["PortfolioOwner"].fillna("").map(normalise_text)
    blank_owner = df["PortfolioOwner"] == ""
    if blank_owner.any():
        settings = get_settings()
        df.loc[blank_owner, "PortfolioOwner"] = df.loc[blank_owner].apply(
            lambda r: settings.portfolio_owner(r["Broker"], r["Portfolio"]), axis=1
        )

    if "PortfolioType" not in df.columns:
        df["PortfolioType"] = ""
    df["PortfolioType"] = df["PortfolioType"].fillna("").map(normalise_text)
    blank_type = df["PortfolioType"] == ""
    if blank_type.any():
        settings = get_settings()
        df.loc[blank_type, "PortfolioType"] = df.loc[blank_type].apply(
            lambda r: settings.portfolio_type(r["Broker"], r["Portfolio"]), axis=1
        )

    df["QuantityDelta"] = df.apply(classify_quantity_delta, axis=1)
    df["CashDelta"] = df.apply(classify_cash_delta, axis=1)

    known_types = (
        QUANTITY_ADD_TYPES | QUANTITY_SUBTRACT_TYPES | QUANTITY_RESET_TYPES
        | CASH_FLOW_TYPES | KNOWN_NEUTRAL_TYPES
    )
    unhandled_mask = ~df["TransactionType"].isin(known_types)
    unhandled = df.loc[unhandled_mask]
    unhandled_summary = (
        unhandled.groupby("TransactionType")
        .size()
        .sort_values(ascending=False)
        .to_dict()
    )

    # Only rows that actually move quantity or cash are kept - this is what
    # keeps the sheet sparse (proportional to real activity, not to elapsed
    # calendar time). Reset-type rows (splits) are always kept even though
    # their naive per-row delta is 0 - apply_quantity_resets() below turns
    # them into the real delta.
    is_reset = df["TransactionType"].isin(QUANTITY_RESET_TYPES)
    active = df[(df["QuantityDelta"] != 0) | (df["CashDelta"] != 0) | is_reset].copy()
    active = active.sort_values(["Broker", "Portfolio", "NormalizedInstrument", "_TradeDate"])

    group_cols = ["Broker", "Portfolio", "NormalizedInstrument"]
    active["QuantityDelta"] = apply_quantity_resets(active, group_cols)
    active["CumulativeQuantity"] = active.groupby(group_cols)["QuantityDelta"].cumsum()

    average_cost_instruments = load_average_cost_instruments(instrument_master)
    active["CashDelta"] = apply_average_cost_basis(active, group_cols, average_cost_instruments)
    active["CumulativeNetInvested"] = active.groupby(group_cols)["CashDelta"].cumsum()

    active["Date"] = active["_TradeDate"].dt.strftime("%Y-%m-%d")

    # A real holding can never go negative. If it does, the TransactionType
    # mapping above doesn't correctly model what actually happened for that
    # instrument (e.g. a corporate action referencing a different base than
    # what we're tracking) - flag it rather than silently write an
    # impossible value.
    negative_rows = active[active["CumulativeQuantity"] < -1e-6]
    negative_summary = (
        negative_rows.groupby(group_cols)["CumulativeQuantity"].min().to_dict()
    )

    # Net invested is a cost-basis figure - it should never be positive
    # (money paid in is always <= 0 under this project's cash-flow sign
    # convention). A positive value means the quantity currently held
    # arrived with an unrecorded cost basis (a gift/transfer event with
    # CashDelta=0, e.g. OpeningPositions or a VAIHTO - JÄTTÖ delivery) and a
    # later sell added real proceeds on top of that unknown zero, producing
    # a fictitious "gain" - flag it instead of silently reporting it.
    positive_invested_rows = active[active["CumulativeNetInvested"] > 1e-6]
    positive_invested_summary = (
        positive_invested_rows.groupby(group_cols)["CumulativeNetInvested"].max().to_dict()
    )

    stats = {
        "transactions_scanned": len(df),
        "position_rows": len(active),
        "unhandled_types": unhandled_summary,
        "unhandled_rows": int(unhandled_mask.sum()),
        "negative_quantity_instruments": negative_summary,
        "positive_net_invested_instruments": positive_invested_summary,
    }

    return active[POSITIONS_COLUMNS].reset_index(drop=True), stats


def write_positions_workbook(positions_workbook: Path, positions: pd.DataFrame) -> None:
    headers = list(POSITIONS_COLUMNS)
    records = positions.to_dict("records")
    sheets = {POSITIONS_SHEET: records_to_sheet_values(headers, records)}

    if positions_workbook.exists():
        replace_with_fresh_workbook(
            positions_workbook,
            sheets,
            backup_label="before_position_rebuild",
            basic_formatting=True,
            excel_tables=False,
        )
    else:
        write_fresh_workbook(positions_workbook, sheets, basic_formatting=True, excel_tables=False)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Recompute PortfolioPositions.xlsx from scratch from current InvestmentTransactions history. "
                    "Always a full rebuild, never incremental - cumulative totals depend on the full transaction "
                    "history, so adding older transactions later must recompute everything downstream correctly."
    )
    parser.add_argument("--investments-workbook", default=str(DEFAULT_INVESTMENTS_WORKBOOK))
    parser.add_argument("--positions-workbook", default=str(DEFAULT_POSITIONS_WORKBOOK))
    parser.add_argument("--instrument-master", default=str(DEFAULT_INSTRUMENT_MASTER))
    parser.add_argument("--dry-run", action="store_true")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()
    investments_workbook = Path(args.investments_workbook).expanduser().resolve()
    positions_workbook = Path(args.positions_workbook).expanduser().resolve()
    instrument_master = Path(args.instrument_master).expanduser().resolve()

    positions, stats = build_positions(investments_workbook, instrument_master)

    print("Portfolio positions rebuild complete." if not args.dry_run else "Dry run complete.")
    print(f"Transactions scanned:      {stats['transactions_scanned']}")
    print(f"Position rows produced:    {stats['position_rows']}")

    if stats["unhandled_types"]:
        print()
        print(f"Unhandled TransactionType values ({stats['unhandled_rows']} rows total, excluded from quantity/cash sums):")
        for tt, count in stats["unhandled_types"].items():
            print(f"  - {tt!r}: {count} row(s)")

    if stats["negative_quantity_instruments"]:
        print()
        print("WARNING: computed a negative (impossible) share count for:")
        for key, min_qty in stats["negative_quantity_instruments"].items():
            print(f"  - {key}: minimum computed quantity {min_qty}")
        print("  This TransactionType mapping does not correctly model what actually happened here - review needed.")

    if stats["positive_net_invested_instruments"]:
        print()
        print("WARNING: computed a positive (impossible) net invested amount for:")
        for key, max_invested in stats["positive_net_invested_instruments"].items():
            print(f"  - {key}: maximum computed net invested {max_invested}")
        print("  This quantity likely arrived with an unrecorded cost basis (a gift/transfer event with no")
        print("  recorded cash amount) - review needed, real cost basis may need to be filled in manually.")

    if args.dry_run:
        print()
        print("Dry run only: workbook was not modified.")
        return

    write_positions_workbook(positions_workbook, positions)
    print()
    print(f"Output workbook: {positions_workbook}")


if __name__ == "__main__":
    main()
