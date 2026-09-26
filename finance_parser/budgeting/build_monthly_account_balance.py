from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path

import pandas as pd

from finance_parser.budgeting.account_balance_seed import (
    BALANCE_FROM_RAW_EXPORT_BANKS,
    DEFAULT_BUDGETING_WORKBOOK,
    EXCLUDED_SOURCE_ACCOUNTS,
    load_raw_transactions,
    owner_lookup,
)
from finance_parser.budgeting.parsers.investment_dividends import SOURCE_BANK as INVESTMENT_DIVIDEND_SOURCE_BANK
from finance_parser.settings import AppSettings, get_settings
from finance_parser.utilities.fresh_workbook_writer import (
    records_to_sheet_values,
    replace_with_fresh_workbook,
    write_fresh_workbook,
)


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_OUTPUT_WORKBOOK = PROJECT_ROOT / "output" / "budgeting" / "MonthlyAccountBalance.xlsx"

MONTHLY_BALANCE_SHEET = "MonthlyAccountBalance"
MONTHLY_BALANCE_COLUMNS = ["MonthEnd", "Year", "Month", "SourceAccount", "SourceBank", "Owner", "Balance", "BalanceSource"]

RAW_EXPORT_SOURCE = "RAW_EXPORT"
RECONSTRUCTED_SOURCE = "RECONSTRUCTED"

# A pending card-authorisation hold (real case: Norwegian, a credit card) -
# not a settled transaction, and its own real BookingDate is blank in the
# export (only ValueDate is populated) since it hasn't posted yet. That
# already excludes it from the date-filtered group today, but only as a
# side effect of being dateless - explicitly filtering by type here means
# this stays correct even if a future export ever gives a pending hold a
# real BookingDate.
KATEVARAUS_TRANSACTION_TYPES = {"KATEVARAUS"}


def last_completed_month_end(today: date) -> pd.Timestamp:
    first_of_this_month = pd.Timestamp(today).replace(day=1)
    return first_of_this_month - pd.Timedelta(days=1)


def month_end_dates(first_date: pd.Timestamp, last_month_end: pd.Timestamp) -> pd.DatetimeIndex:
    if first_date > last_month_end:
        return pd.DatetimeIndex([])
    return pd.date_range(start=first_date, end=last_month_end, freq="ME")


def _account_owner(account: str, owners: dict[str, str]) -> str:
    return owners.get(account, "")


def _has_continuous_coverage(seed_period: pd.Period, target_period: pd.Period, active_months: set) -> bool:
    """
    True if every calendar month strictly between seed_period and
    target_period (in whichever direction) has at least one real
    transaction - i.e. reconstruction math can walk that whole chain without
    silently crossing an unimported gap.
    """
    if seed_period == target_period:
        return True
    step = 1 if target_period > seed_period else -1
    cursor = seed_period + step
    while True:
        if cursor not in active_months:
            return False
        if cursor == target_period:
            return True
        cursor += step


def _reliable_reconstructed_months(
    group: pd.DataFrame, months: pd.DatetimeIndex, seeds: list[tuple[str, float]]
) -> pd.DatetimeIndex:
    """
    Filters candidate month-ends down to ones connected to their applicable
    seed by an unbroken run of calendar months with at least one real
    transaction. Reconstruction is a running sum, so any real, unimported
    gap in that chain silently skews every month past it - the further from
    the seed, the worse the drift. Real case that surfaced this: an account
    with a real multi-year stretch of consecutive calendar months with zero
    imported transactions (an import gap, not genuine long-term dormancy)
    produced wildly wrong, deeply negative historical balances before this
    guard existed. A calendar month with genuinely zero real activity looks
    identical to an unimported gap from the data alone, so this is
    deliberately conservative - it may withhold a real, correct month rather
    than risk a wrong one.
    """
    seed_periods = sorted({pd.Timestamp(d).to_period("M") for d, _ in seeds})
    active_months = set(group["_Date"].dt.to_period("M"))

    reliable = []
    for month_end in months:
        target_period = month_end.to_period("M")
        applicable = [p for p in seed_periods if p <= target_period]
        seed_period = applicable[-1] if applicable else seed_periods[0]

        if _has_continuous_coverage(seed_period, target_period, active_months):
            reliable.append(month_end)

    return pd.DatetimeIndex(reliable)


def _raw_export_balance_rows(account: str, bank: str, group: pd.DataFrame, months: pd.DatetimeIndex, owners: dict[str, str]) -> list[dict]:
    """
    Account whose bank export already carries its own running balance
    (Nordea's real "Saldo" field) - use it directly, merge-backward to each
    month-end, same pattern as Nordnet's investment cash-balance tracking.
    """
    balance_source = group.dropna(subset=["_Balance"])[["_Date", "_Balance"]].sort_values("_Date")
    if len(balance_source) == 0:
        return []

    candidates = pd.DataFrame({"MonthEnd": months})
    merged = pd.merge_asof(
        candidates, balance_source, left_on="MonthEnd", right_on="_Date", direction="backward"
    )
    merged = merged.dropna(subset=["_Balance"])

    owner = _account_owner(account, owners)
    rows = []
    for _, r in merged.iterrows():
        month_end = r["MonthEnd"]
        rows.append({
            "MonthEnd": month_end.strftime("%Y-%m-%d"),
            "Year": month_end.year,
            "Month": month_end.month,
            "SourceAccount": account,
            "SourceBank": bank,
            "Owner": owner,
            "Balance": round(float(r["_Balance"]), 2),
            "BalanceSource": RAW_EXPORT_SOURCE,
        })
    return rows


def _reconstructed_balance_rows(
    account: str, bank: str, group: pd.DataFrame, months: pd.DatetimeIndex, seeds: list[tuple[str, float]], owners: dict[str, str]
) -> list[dict]:
    """
    Account with no running balance of its own - reconstructed from the
    nearest seed. For a month at-or-after a seed, project FORWARD (cumsum of
    real Amount after it). For a month before every seed - real transaction
    history predating the earliest seed, so its own months are also
    computable - project BACKWARD from the earliest seed instead (subtract
    real Amount between the month-end and the seed). Both directions are
    equally exact, deterministic sums; backward isn't a guess. A seed's date
    is always end-of-day (see account_balance_seed.py's collector): forward
    projection sums transactions STRICTLY after it, backward projection
    subtracts transactions up to AND INCLUDING it - both exclude the seed's
    own date from being double counted.
    """
    if not seeds:
        return []

    seed_pairs = [(pd.Timestamp(d), b) for d, b in seeds]  # oldest-first, per account_balance_seeds()
    rows = []
    owner = _account_owner(account, owners)

    for month_end in months:
        at_or_before = [p for p in seed_pairs if p[0] <= month_end]
        if at_or_before:
            seed_date, seed_balance = at_or_before[-1]
            window = group[(group["_Date"] > seed_date) & (group["_Date"] <= month_end)]
            balance = seed_balance + window["_Amount"].sum()
        else:
            seed_date, seed_balance = seed_pairs[0]
            window = group[(group["_Date"] > month_end) & (group["_Date"] <= seed_date)]
            balance = seed_balance - window["_Amount"].sum()

        rows.append({
            "MonthEnd": month_end.strftime("%Y-%m-%d"),
            "Year": month_end.year,
            "Month": month_end.month,
            "SourceAccount": account,
            "SourceBank": bank,
            "Owner": owner,
            "Balance": round(float(balance), 2),
            "BalanceSource": RECONSTRUCTED_SOURCE,
        })

    return rows


def build_monthly_balances(
    budgeting_workbook: Path = DEFAULT_BUDGETING_WORKBOOK,
    settings: AppSettings | None = None,
    *,
    as_of: date | None = None,
) -> tuple[pd.DataFrame, dict[str, object]]:
    settings = settings or get_settings()
    df = load_raw_transactions(budgeting_workbook)
    owners = owner_lookup(budgeting_workbook)

    stats: dict[str, object] = {
        "accounts_processed": [],
        "accounts_no_seed_configured": [],
        "accounts_raw_export_no_balance_data": [],
    }

    if len(df) == 0:
        return pd.DataFrame(columns=MONTHLY_BALANCE_COLUMNS), stats

    df["_Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0.0)
    if "Balance" in df.columns:
        df["_Balance"] = pd.to_numeric(df["Balance"], errors="coerce")
    else:
        df["_Balance"] = pd.NA

    last_month_end = last_completed_month_end(as_of or date.today())
    rows: list[dict] = []

    for account, group in df.groupby("SourceAccount"):
        if not account or account in EXCLUDED_SOURCE_ACCOUNTS:
            continue

        group = group[group["_Date"].notna()].sort_values("_Date")
        # Synthetic dividend-income rows (Nordnet/EVLI, injected by
        # investment_dividends.py) never hit this bank account for real -
        # the cash lands in the broker's own cash balance instead (see that
        # module's own docstring). They're correct for budgeting's income
        # view but must not count as a real balance-affecting event here, or
        # they double-book money that's actually still at the broker. Real
        # bug this fixes: found via a real dry-run - these rows were
        # inflating several months' reconstructed balance by exactly their
        # own summed amount.
        group = group[group["SourceBank"] != INVESTMENT_DIVIDEND_SOURCE_BANK]
        # Real bug this fixes: a credit card's pending card-authorisation
        # hold (katevaraus) isn't a real settled transaction - explicit by
        # type, not just relying on it happening to have a blank date (see
        # KATEVARAUS_TRANSACTION_TYPES above).
        if "TransactionTypeRaw" in group.columns:
            group = group[~group["TransactionTypeRaw"].astype(str).str.strip().str.upper().isin(KATEVARAUS_TRANSACTION_TYPES)]
        if len(group) == 0:
            continue

        earliest_transaction_date = group["_Date"].iloc[0]
        latest_transaction_date = group["_Date"].iloc[-1]
        # A month-end is only trustworthy once there's evidence the import
        # continued past it (a real transaction dated after it) - not just
        # because the calendar says the month is over. Real bug this fixes:
        # an account whose latest import landed mid-month (e.g. the 18th)
        # was still getting a full end-of-month balance for that same month,
        # silently assuming zero further transactions that just hadn't been
        # imported yet. Capping here (rather than only via last_month_end)
        # means date_range's own upper bound naturally excludes any
        # not-yet-confirmed trailing month.
        account_last_month_end = min(last_month_end, latest_transaction_date)

        banks = set(group["SourceBank"].unique())
        is_raw_export_account = bool(banks) and banks.issubset(BALANCE_FROM_RAW_EXPORT_BANKS)
        # An account is expected to map to exactly one real bank across its
        # whole history - "/"-joining is a defensive fallback (never
        # observed in real data) rather than silently picking one.
        bank = "/".join(sorted(banks)) if banks else ""

        if is_raw_export_account:
            months = month_end_dates(earliest_transaction_date, account_last_month_end)
            account_rows = _raw_export_balance_rows(account, bank, group, months, owners)
            if account_rows:
                stats["accounts_processed"].append(account)
            else:
                # A bank in BALANCE_FROM_RAW_EXPORT_BANKS is expected to
                # populate Balance - flag rather than silently produce zero
                # rows if it turns out not to for this account.
                stats["accounts_raw_export_no_balance_data"].append(account)
            rows.extend(account_rows)
            continue

        seeds = settings.account_balance_seeds(account)
        if not seeds:
            stats["accounts_no_seed_configured"].append(account)
            continue

        # Start as far back as either the earliest seed or real transaction
        # history reaches, whichever is earlier - a seed alone anchors a
        # known balance from its own date, and real transactions predating
        # it (if imported) extend that further back via backward projection
        # (see _reconstructed_balance_rows).
        earliest_seed_date = pd.Timestamp(seeds[0][0])
        start = min(earliest_transaction_date, earliest_seed_date)
        candidate_months = month_end_dates(start, account_last_month_end)
        if settings.trusts_sparse_activity(account):
            # Explicit per-account opt-in (see AppSettings.trusts_sparse_
            # activity's own docstring) - skip the gap-guard entirely for
            # this account. The underlying math is unaffected: a month-end
            # whose window crosses real transactions still correctly
            # subtracts/adds them; only the "was every month covered"
            # safety check is skipped.
            months = candidate_months
        else:
            months = _reliable_reconstructed_months(group, candidate_months, seeds)
        account_rows = _reconstructed_balance_rows(account, bank, group, months, seeds, owners)
        if account_rows:
            stats["accounts_processed"].append(account)
        rows.extend(account_rows)

    result = pd.DataFrame(rows, columns=MONTHLY_BALANCE_COLUMNS)
    if len(result) > 0:
        result = result.sort_values(["SourceAccount", "MonthEnd"]).reset_index(drop=True)

    return result, stats


def write_monthly_balance_workbook(output_workbook: Path, monthly_balances: pd.DataFrame) -> None:
    sheets = {MONTHLY_BALANCE_SHEET: records_to_sheet_values(MONTHLY_BALANCE_COLUMNS, monthly_balances.to_dict("records"))}

    if output_workbook.exists():
        replace_with_fresh_workbook(
            output_workbook,
            sheets,
            backup_label="before_account_balance_rebuild",
            basic_formatting=True,
            excel_tables=False,
        )
    else:
        write_fresh_workbook(output_workbook, sheets, basic_formatting=True, excel_tables=False)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Build a monthly end-of-month balance per budgeting account - budgeting's "
                    "equivalent of the investment side's MonthlyPositionValue.xlsx. Nordea uses "
                    "its own real running balance; other accounts reconstruct from a manually "
                    "seeded balance (see tools/seed_account_balance.py) plus real transactions since."
    )
    parser.add_argument("--budgeting-workbook", default=str(DEFAULT_BUDGETING_WORKBOOK))
    parser.add_argument("--output-workbook", default=str(DEFAULT_OUTPUT_WORKBOOK))
    parser.add_argument("--dry-run", action="store_true", help="Compute and report, but do not write the workbook.")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()

    budgeting_workbook = Path(args.budgeting_workbook).expanduser().resolve()
    output_workbook = Path(args.output_workbook).expanduser().resolve()

    monthly_balances, stats = build_monthly_balances(budgeting_workbook)

    print()
    print("Monthly account balance build complete." if not args.dry_run else "Monthly account balance dry run complete.")
    print(f"Accounts with a monthly balance: {len(stats['accounts_processed'])}")
    for account in stats["accounts_processed"]:
        print(f"  - {account}")

    if stats["accounts_no_seed_configured"]:
        print()
        print("Accounts with no balance seed configured yet (no monthly rows produced):")
        for account in stats["accounts_no_seed_configured"]:
            print(f"  - {account}")
        print("Run python tools/seed_account_balance.py to add one.")

    if stats["accounts_raw_export_no_balance_data"]:
        print()
        print("Accounts expected to have a real running balance in their own export, but none")
        print("found (no monthly rows produced) - check the raw export actually has it:")
        for account in stats["accounts_raw_export_no_balance_data"]:
            print(f"  - {account}")

    if args.dry_run:
        print()
        print("Dry run only: workbook was not modified.")
        return

    write_monthly_balance_workbook(output_workbook, monthly_balances)
    print()
    print(f"Output workbook: {output_workbook}")


if __name__ == "__main__":
    main()
