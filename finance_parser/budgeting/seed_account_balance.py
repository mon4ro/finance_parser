from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

from finance_parser.budgeting.account_balance_seed import (
    DEFAULT_BUDGETING_WORKBOOK,
    accounts_needing_seed,
    collect_account_balance_seed,
)
from finance_parser.setup_wizard import DEFAULT_SETTINGS_PATH, write_settings
from finance_parser.settings import AppSettings
from finance_parser.utilities.interactive_prompts import ask_multi_select


def run(budgeting_workbook: Path = DEFAULT_BUDGETING_WORKBOOK, settings_path: Path = DEFAULT_SETTINGS_PATH) -> bool:
    print("Account balance seeding")
    print("========================")
    print("Adds/updates a real 'balance as of this date' anchor point for a budgeting")
    print("account, so build_monthly_account_balance.py can reconstruct its monthly")
    print("balance history. Runnable anytime - not just at initial setup.")

    accounts = accounts_needing_seed(budgeting_workbook)
    if not accounts:
        print()
        print("No accounts need seeding right now - either no budgeting transactions have")
        print("been imported yet, or every imported account's bank already provides a real")
        print("balance (e.g. Nordea).")
        return False

    print()
    print("Accounts that can use a seed:")
    chosen = ask_multi_select(accounts)
    if not chosen:
        print("Nothing selected.")
        return False

    settings = AppSettings.load(settings_path=settings_path)
    overrides: dict[str, Any] = {}
    any_added = False
    for account in chosen:
        existing = settings.account_balance_seeds(account)
        added = collect_account_balance_seed(
            overrides, account, budgeting_workbook=budgeting_workbook, existing_seeds=existing
        )
        any_added = any_added or added

    if not any_added:
        print()
        print("Nothing to write.")
        return False

    return write_settings(overrides, settings_path=settings_path)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Interactively add/update a budgeting account's balance-as-of-date "
                    "seed in config/settings.yaml, for accounts whose bank export doesn't "
                    "already carry a running balance."
    )
    parser.add_argument("--budgeting-workbook", default=str(DEFAULT_BUDGETING_WORKBOOK))
    parser.add_argument("--settings-path", default=str(DEFAULT_SETTINGS_PATH))
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()
    run(
        budgeting_workbook=Path(args.budgeting_workbook).expanduser().resolve(),
        settings_path=Path(args.settings_path).expanduser().resolve(),
    )


if __name__ == "__main__":
    main()
