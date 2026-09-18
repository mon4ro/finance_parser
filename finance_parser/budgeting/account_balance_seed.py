from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd

from finance_parser.common import normalise_header, normalise_text
from finance_parser.utilities.interactive_prompts import ask, ask_yes_no


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_BUDGETING_WORKBOOK = PROJECT_ROOT / "output" / "budgeting" / "ParsedTransactions.xlsx"

# Not a real bank account with a balance to check against a bank app - the
# manually-maintained CASH ledger tracks small untracked real expenses, not
# an account with its own concept of "current balance".
EXCLUDED_SOURCE_ACCOUNTS = {"CASH"}
# Nordea's own export already carries a real running balance per transaction
# (see build_monthly_account_balance.py) - seeding one here would be at best
# redundant and at worst a confusing second, possibly-inconsistent source of
# truth for the same account.
BALANCE_FROM_RAW_EXPORT_BANKS = {"NORDEA"}


def load_raw_transactions(budgeting_workbook: Path) -> pd.DataFrame:
    """
    Reads RawTransactions directly - deliberately NOT UnifiedTransactions.
    RawTransactions is one row per real imported transaction: no Include
    flag to accidentally lean on, no manually-split child rows (the Excel
    macro only ever touches UnifiedTransactions), no synthetic income rows
    injected by a later pipeline stage, and Balance/BookingDate are already
    native - nothing to fall back or reconstruct. Real bugs found and fixed
    by using UnifiedTransactions instead, before this: a split transaction's
    untouched parent row double-counted alongside its children, and a
    transaction unified before Balance existed on that schema stayed blank
    forever. Reading Raw sidesteps both classes of bug structurally, not by
    patching around them.

    Uses BookingDate (not ValueDate) as the real ledger date - the date a
    transaction actually posted, which is what an end-of-day balance
    snapshot reflects.
    """
    if not budgeting_workbook.exists():
        return pd.DataFrame(columns=["SourceAccount", "SourceBank", "BookingDate", "Amount", "Balance"])

    df = pd.read_excel(budgeting_workbook, sheet_name="RawTransactions", dtype=object, engine="openpyxl")
    df.columns = [normalise_header(c) for c in df.columns]

    df["SourceAccount"] = df["SourceAccount"].map(normalise_text)
    df["SourceBank"] = df["SourceBank"].map(normalise_text)
    df["_Date"] = pd.to_datetime(df["BookingDate"], errors="coerce")

    if "Balance" not in df.columns:
        df["Balance"] = pd.NA

    return df


def owner_lookup(budgeting_workbook: Path) -> dict[str, str]:
    """
    A small, read-only side lookup of SourceAccount -> Owner, sourced from
    UnifiedTransactions - deliberately kept OUT of the balance/date/amount
    path above. Owner is resolved by content-matching OwnershipRules at the
    categoriser layer, which only ever runs against UnifiedTransactions -
    RawTransactions has no Owner field and no way to derive one on its own.
    This is a pure display label (one value per account, never summed or
    used in any date/gap/duplicate logic), so reading it from Unified here
    doesn't reintroduce the risks (Include mixing, split/synthetic rows)
    that motivated moving the actual balance math off Unified. Grouped by
    SourceAccount, never by Owner, so two different accounts that happen to
    share an owner are never merged together.
    """
    if not budgeting_workbook.exists():
        return {}

    try:
        df = pd.read_excel(budgeting_workbook, sheet_name="UnifiedTransactions", dtype=object, engine="openpyxl")
    except (ValueError, KeyError):
        return {}

    df.columns = [normalise_header(c) for c in df.columns]
    if "SourceAccount" not in df.columns or "Owner" not in df.columns:
        return {}

    df["SourceAccount"] = df["SourceAccount"].map(normalise_text)
    df["Owner"] = df["Owner"].map(normalise_text)

    lookup: dict[str, str] = {}
    for account, group in df.groupby("SourceAccount"):
        if not account:
            continue
        non_blank = group["Owner"][group["Owner"] != ""]
        if len(non_blank) > 0:
            lookup[account] = non_blank.iloc[0]

    return lookup


def accounts_needing_seed(budgeting_workbook: Path = DEFAULT_BUDGETING_WORKBOOK) -> list[str]:
    """
    SourceAccounts with real imported transactions whose bank doesn't
    already provide a running balance - i.e. the ones a manual seed would
    actually be useful for. Excludes CASH and any account whose transactions
    are ALL from a BALANCE_FROM_RAW_EXPORT_BANKS bank.
    """
    df = load_raw_transactions(budgeting_workbook)
    if len(df) == 0:
        return []

    accounts: list[str] = []
    for account, group in df.groupby("SourceAccount"):
        if not account or account in EXCLUDED_SOURCE_ACCOUNTS:
            continue
        banks = set(group["SourceBank"].unique())
        if banks and banks.issubset(BALANCE_FROM_RAW_EXPORT_BANKS):
            continue
        accounts.append(account)

    return sorted(accounts)


def latest_imported_date(source_account: str, budgeting_workbook: Path = DEFAULT_BUDGETING_WORKBOOK) -> date | None:
    """
    The most recent transaction date actually imported for this account -
    the safe default (and upper bound) for a balance seed's own date, since
    we can't vouch for a later date's transactions being fully captured yet.
    """
    df = load_raw_transactions(budgeting_workbook)
    if len(df) == 0:
        return None

    sub = df[(df["SourceAccount"] == source_account) & df["_Date"].notna()]
    if len(sub) == 0:
        return None

    return sub["_Date"].max().date()


def collect_account_balance_seed(
    overrides: dict[str, Any],
    account: str,
    *,
    budgeting_workbook: Path = DEFAULT_BUDGETING_WORKBOOK,
    existing_seeds: list[tuple[str, float]] | None = None,
) -> bool:
    """
    Interactively collect one balance seed for a single account, writing it
    into overrides["budgeting"]["account_balance_seeds"][account][date] if
    confirmed. Returns True if a seed was added.

    Shared by seed_account_balance.py (the standalone, run-anytime tool) and
    setup_wizard.py's optional step - same prompts, same safety checks,
    either way.
    """
    print()
    print(f"{account}: current balance")
    if existing_seeds:
        print(f"  Already has {len(existing_seeds)} seed(s):")
        for seed_date, balance in existing_seeds:
            print(f"    - {seed_date}: {balance:.2f}")

    print(
        "  Use the END-OF-DAY balance - after the LAST transaction that posted on the "
        "chosen date, not a mid-day snapshot. If more than one transaction happens that "
        "day, they're all already reflected in this one number."
    )
    print(
        "  Use your BOOKED/LEDGER balance, not \"available balance\" - many bank apps "
        "subtract pending holds (e.g. a card pre-authorisation, \"katevaraus\") from the "
        "figure they show by default. A pending hold isn't a real settled transaction "
        "yet and won't exist in your export, so it would make the seed wrong."
    )

    latest = latest_imported_date(account, budgeting_workbook)
    if latest is None:
        print(f"  No transactions imported yet for {account} - you can still seed a balance,")
        print("  but it won't produce any monthly rows until real transactions exist after it.")
        default_date = None
    else:
        default_date = latest.isoformat()
        print(f"  Latest imported transaction for {account}: {default_date}.")

    date_str = ask("  Balance as of date (YYYY-MM-DD)", default=default_date)
    if not date_str:
        print("  Skipped - no date given.")
        return False

    if latest is not None and date_str > latest.isoformat():
        print()
        print(f"  WARNING: {date_str} is AFTER the latest imported transaction ({latest.isoformat()}).")
        print("  Transactions for a date this recent might not be fully captured yet - if any")
        print("  more show up later dated on or before this date, the reconstructed balance")
        print("  for every month from here on would be wrong.")
        if not ask_yes_no("  Use this date anyway?", default=False):
            print("  Skipped.")
            return False

    balance_str = ask("  Balance (EUR)")
    if not balance_str:
        print("  Skipped - no balance given.")
        return False

    try:
        balance = round(float(balance_str.replace(",", ".")), 2)
    except ValueError:
        print(f"  Skipped - '{balance_str}' isn't a number.")
        return False

    node = overrides.setdefault("budgeting", {}).setdefault("account_balance_seeds", {}).setdefault(account, {})
    node[date_str] = balance
    print(f"  Recorded: {account} = {balance:.2f} EUR as of {date_str}.")
    return True
