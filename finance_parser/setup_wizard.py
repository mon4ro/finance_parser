from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path
from typing import Any

import pandas as pd
import yaml

from finance_parser.budgeting.account_balance_seed import (
    DEFAULT_BUDGETING_WORKBOOK,
    accounts_needing_seed,
    collect_account_balance_seed,
)
from finance_parser.budgeting.transaction_parser import SUPPORTED_PARSERS as BUDGETING_PARSER_MODULES
from finance_parser.investments.investment_parser import SUPPORTED_INVESTMENT_PARSERS as INVESTMENT_PARSER_MODULES
from finance_parser.settings import AppSettings, _deep_merge
# ask/ask_yes_no/ask_multi_select re-exported for this module's existing
# call sites/tests - the implementations live in interactive_prompts.py so
# finance_parser.budgeting.account_balance_seed can use them too, without an
# awkward import from this top-level orchestration module.
from finance_parser.utilities.interactive_prompts import ask, ask_multi_select, ask_yes_no
from finance_parser.utilities.fresh_workbook_writer import (
    read_workbook_values_only,
    records_to_sheet_values,
    replace_with_fresh_workbook,
    sheet_values_to_records,
    write_fresh_workbook,
)


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SETTINGS_PATH = PROJECT_ROOT / "config" / "settings.yaml"
DEFAULT_BUDGETING_TEMPLATE = PROJECT_ROOT / "rules" / "budgeting" / "TransactionRules.template.xlsx"
DEFAULT_BUDGETING_RULES = PROJECT_ROOT / "rules" / "budgeting" / "TransactionRules.xlsx"
DEFAULT_INVESTMENT_TEMPLATE = PROJECT_ROOT / "rules" / "investments" / "InstrumentMaster.template.xlsx"
DEFAULT_INVESTMENT_RULES = PROJECT_ROOT / "rules" / "investments" / "InstrumentMaster.xlsx"

GITHUB_ISSUES_URL = "https://github.com/mon4ro/finance_parser/issues"

# CASH is a manually maintained ledger, not a bank export - offered as its
# own yes/no toggle rather than living in the broker multi-select list.
CASH_SOURCE_BANK = "CASH"
# Synthetic cross-pipeline parser (reads the investment side's own
# DividendHistory.xlsx) - runs automatically once the investment pipeline
# has produced that file, never something a user manually opts into here.
SYNTHETIC_BUDGETING_SOURCE_BANKS = {"INVESTMENT_DIVIDEND"}

# TransactionRules.template.xlsx deliberately ships with no OwnershipRules
# sheet at all (see its own README sheet) - it's built up as accounts get
# configured, not prescribed upfront. transaction_categoriser.py already
# treats a missing sheet as "zero rules" rather than an error, so creating
# it fresh here (matching the real schema below) is safe on a brand-new
# workbook, not just an existing one.
OWNERSHIP_RULES_SHEET = "OwnershipRules"
OWNERSHIP_RULES_COLUMNS = [
    "RuleID", "Enabled", "Priority", "RuleName", "MatchField", "MatchType",
    "Pattern", "CaseSensitive", "SetOwner", "ClearOwner", "StopIfMatched",
    "OverwriteMode", "Notes",
]


def budgeting_broker_names() -> list[str]:
    """
    Live-derived from the actual parser registry, not a hand-maintained
    list, so this never drifts out of sync as parsers are added/removed.
    """
    names = {module.SOURCE_BANK for module in BUDGETING_PARSER_MODULES}
    names -= {CASH_SOURCE_BANK}
    names -= SYNTHETIC_BUDGETING_SOURCE_BANKS
    return sorted(names)


def investment_broker_names() -> list[str]:
    return sorted({module.BROKER for module in INVESTMENT_PARSER_MODULES})


def choose_scope() -> set[str]:
    print()
    print("Which kind of data will you import?")
    print("  1. Budgeting only (bank transactions)")
    print("  2. Investments only (broker/portfolio exports)")
    print("  3. Both")
    while True:
        choice = input("Choice [3]: ").strip() or "3"
        if choice == "1":
            return {"budgeting"}
        if choice == "2":
            return {"investments"}
        if choice == "3":
            return {"budgeting", "investments"}
        print("Please enter 1, 2, or 3.")


def _set_nested(target: dict[str, Any], path: list[str], value: Any) -> None:
    node = target
    for key in path[:-1]:
        node = node.setdefault(key, {})
    node[path[-1]] = value


def _explain_account_owner(broker: str) -> None:
    print()
    print(f"{broker}: who is the Owner of this account - the household member it belongs to?")
    print("This is just a starting default for the account as a whole. It's separate from")
    print("which household member an individual transaction ends up attributed to for")
    print("budgeting purposes later - that can still be set independently, per transaction,")
    print("regardless of the account's owner.")


def collect_op_budgeting_settings(overrides: dict[str, Any]) -> list[str]:
    _explain_account_owner("OP")
    print("Each OP export file is matched to an account by a prefix in its filename, e.g.")
    print("HOUSEHOLD_tapahtumat20260503-20260603.csv -> HOUSEHOLD.")
    count_raw = ask("How many separate OP export accounts do you have", default="1")
    try:
        count = max(1, int(count_raw))
    except ValueError:
        count = 1

    prefixes: dict[str, str] = {}
    owners: list[str] = []
    for i in range(count):
        owner = ask(f"  Account #{i + 1} Owner (e.g. HOUSEHOLD, PERSONAL, CHILD)")
        if not owner:
            continue
        prefix = ask(f"  Filename prefix for '{owner.upper()}'", default=owner.upper())
        prefixes[owner.upper()] = prefix.upper()
        owners.append(owner.upper())

    if prefixes:
        _set_nested(overrides, ["budgeting", "source_account_inference", "OP", "filename_prefixes"], prefixes)

    return owners


def collect_fixed_account_budgeting_settings(overrides: dict[str, Any], broker: str) -> list[str]:
    _explain_account_owner(broker)
    owner = ask(f"{broker}: Owner", default=broker)
    if not owner:
        return []
    _set_nested(overrides, ["budgeting", "source_account_inference", broker, "fixed_source_account"], owner.upper())
    return [owner.upper()]


def collect_spankki_settings(overrides: dict[str, Any]) -> list[str]:
    print()
    print("S-Pankki (budgeting): often used as a buffer/pass-through account (e.g. a shared")
    print("grocery account topped up from a joint account) whose own balance is reconciled")
    print("separately - if so, its individual transactions shouldn't also count toward budget")
    print("totals, or money moving through it gets double-counted.")
    is_buffer = ask_yes_no(
        "Is this account a buffer/pass-through account reconciled separately elsewhere?",
        default=True,
    )
    _set_nested(overrides, ["budgeting", "source_account_inference", "SPANKKI", "fixed_source_account"], "SPANKKI")
    _set_nested(
        overrides,
        ["budgeting", "source_account_inference", "SPANKKI", "default_include"],
        "NO" if is_buffer else "YES",
    )
    # SPANKKI is the account-type name, not a household member - no Owner to
    # register in OwnershipRules here (unlike OP/Norwegian/Nordea above).
    return []


BUDGETING_SETTINGS_HANDLERS = {
    "OP": collect_op_budgeting_settings,
    "NORWEGIAN": lambda overrides: collect_fixed_account_budgeting_settings(overrides, "NORWEGIAN"),
    "NORDEA": lambda overrides: collect_fixed_account_budgeting_settings(overrides, "NORDEA"),
    "SPANKKI": collect_spankki_settings,
}


# Brokers where the account-wrapper type is a genuine AOT (Arvo-osuustili -
# regular custody/investment account) vs OST (Osakesäästötili - equity
# savings account) choice. EVLI is an employer share-plan account (neither),
# and COINMOTION/SELIGSON have a single known fixed type - see
# FIXED_PORTFOLIO_TYPES below - so none of those three get asked.
AOT_OST_CAPABLE_BROKERS = {"OP", "NORDEA", "NORDNET"}
PORTFOLIO_TYPE_CHOICES = {"1": "Arvo-osuustili", "2": "Osakesäästötili"}

# Account-wrapper type is a known, fixed technical fact for these brokers,
# not a personal choice and not an AOT/OST distinction - auto-filled rather
# than asked.
FIXED_PORTFOLIO_TYPES = {
    "COINMOTION": "Crypto wallet",
    "SELIGSON": "Arvo-osuustili",  # direct fund purchase only, always custody-type
}


def ask_portfolio_type(context: str) -> str:
    print(f"{context} account type:")
    print("  1. AOT - Arvo-osuustili (regular custody/investment account)")
    print("  2. OST - Osakesäästötili (equity savings account)")
    choice = ask("Choice (blank to skip)")
    return PORTFOLIO_TYPE_CHOICES.get(choice.strip(), "")


def _existing_broker_value(existing: dict[str, Any], section: str, broker: str) -> Any:
    return (existing.get("investments", {}) or {}).get(section, {}).get(broker)


def collect_portfolio_owners(overrides: dict[str, Any], broker: str, existing: dict[str, Any] | None = None) -> None:
    existing = existing or {}
    existing_owners = _existing_broker_value(existing, "portfolio_owners", broker)
    existing_types = _existing_broker_value(existing, "portfolio_types", broker)
    existing_is_multi = isinstance(existing_owners, dict) or isinstance(existing_types, dict)

    print()
    if existing_is_multi:
        print(f"{broker}: your current settings already have multiple portfolios configured:")
        if isinstance(existing_owners, dict):
            for portfolio_id, owner in existing_owners.items():
                print(f"  - {portfolio_id}: {owner}")

    multiple = ask_yes_no(
        f"{broker}: do you have more than one portfolio/account under this broker?",
        default=existing_is_multi,
    )

    if not multiple:
        if existing_is_multi:
            print()
            print(f"WARNING: answering with a single owner here will REPLACE your existing")
            print(f"per-portfolio configuration for {broker} above with one flat value.")
            if not ask_yes_no("Are you sure you want to replace it?", default=False):
                print("Keeping your existing per-portfolio configuration unchanged.")
                return

        owner = ask(f"{broker}: who does this portfolio belong to?", default="HOUSEHOLD")
        if owner:
            _set_nested(overrides, ["investments", "portfolio_owners", broker], owner.upper())

        if broker in AOT_OST_CAPABLE_BROKERS:
            portfolio_type = ask_portfolio_type(broker)
            if portfolio_type:
                _set_nested(overrides, ["investments", "portfolio_types", broker], portfolio_type)
        return

    print()
    print(f"One {broker} login can hold several separate portfolios (e.g. one per household")
    print("member) - each needs its own owner. You'll be asked for a portfolio number/id")
    print(f"(check your {broker} export or account overview for this) and who it belongs to,")
    print("one pair at a time. Press Enter on a blank portfolio number/id when you're done")
    print("listing them - you'll then be asked for one last fallback owner, used only if a")
    print("transaction ever shows up from a portfolio you didn't list (e.g. a new account")
    print("opened later), so nothing silently ends up with no owner at all.")

    owners: dict[str, str] = {}
    types: dict[str, str] = {}
    while True:
        portfolio = ask("  Portfolio number/id (blank to finish)")
        if not portfolio:
            break
        owner = ask(f"  Who does portfolio {portfolio} belong to?", default="HOUSEHOLD")
        owners[portfolio] = owner.upper()

        if broker in AOT_OST_CAPABLE_BROKERS:
            portfolio_type = ask_portfolio_type(f"  Portfolio {portfolio}")
            if portfolio_type:
                types[portfolio] = portfolio_type

    default_owner = ask(f"{broker}: fallback owner for any other, unlisted portfolio", default="HOUSEHOLD")
    if default_owner:
        owners["default"] = default_owner.upper()

    if owners:
        _set_nested(overrides, ["investments", "portfolio_owners", broker], owners)
    if types:
        _set_nested(overrides, ["investments", "portfolio_types", broker], types)


def collect_fixed_instrument_name(overrides: dict[str, Any], broker: str, hint: str) -> None:
    print()
    print(hint)
    name = ask(f"{broker}: instrument/fund name for this account")
    if name:
        _set_nested(overrides, ["investments", "broker_defaults", broker, "fixed_instrument_name"], name.upper())


def collect_investment_broker_settings(overrides: dict[str, Any], broker: str, existing: dict[str, Any] | None = None) -> None:
    if broker == "EVLI":
        collect_fixed_instrument_name(
            overrides,
            "EVLI",
            "EVLI's export can list plan/program labels instead of an instrument name (e.g. "
            "'Plan Cycle 2023', 'SIS Dividend') - tell it what the underlying instrument really is.",
        )
    elif broker == "NORDEA":
        collect_fixed_instrument_name(
            overrides,
            "NORDEA",
            "Nordea's own Excel export has no instrument/fund name column at all - one custody "
            "account here holds one fund, so the name has to come from settings.",
        )
    collect_portfolio_owners(overrides, broker, existing)

    if broker in FIXED_PORTFOLIO_TYPES:
        _set_nested(overrides, ["investments", "portfolio_types", broker], FIXED_PORTFOLIO_TYPES[broker])


def collect_unsupported(prompt: str) -> list[str]:
    print()
    raw = ask(prompt)
    if not raw:
        return []
    return [part.strip() for part in raw.split(",") if part.strip()]


def _existing_ownership_patterns(rules_path: Path) -> set[str]:
    if not rules_path.exists():
        return set()
    try:
        df = pd.read_excel(rules_path, sheet_name=OWNERSHIP_RULES_SHEET, dtype=object, engine="openpyxl")
    except ValueError:
        return set()
    if "Pattern" not in df.columns or "MatchField" not in df.columns:
        return set()
    matched = df[df["MatchField"].astype(str).str.strip().str.upper() == "SOURCEACCOUNT"]
    return {str(p).strip().upper() for p in matched["Pattern"].dropna()}


def _next_ownership_rule_number(rules_path: Path) -> int:
    if not rules_path.exists():
        return 1
    try:
        df = pd.read_excel(rules_path, sheet_name=OWNERSHIP_RULES_SHEET, dtype=object, engine="openpyxl")
    except ValueError:
        return 1
    numbers = []
    for raw in df.get("RuleID", []):
        text = str(raw).strip().upper()
        if text.startswith("OR") and text[2:].isdigit():
            numbers.append(int(text[2:]))
    return (max(numbers) + 1) if numbers else 1


def _build_ownership_rule_row(rule_number: int, owner: str) -> dict[str, Any]:
    return {
        "RuleID": f"OR{rule_number:04d}",
        "Enabled": "YES",
        "Priority": 100,
        "RuleName": f"Source account {owner} owns transaction",
        "MatchField": "SourceAccount",
        "MatchType": "EXACT",
        "Pattern": owner,
        "CaseSensitive": "NO",
        "SetOwner": owner,
        "ClearOwner": "NO",
        "StopIfMatched": "YES",
        "OverwriteMode": "BLANK_ONLY",
        "Notes": "Default owner from source account (added by setup_wizard.py).",
    }


def write_ownership_rules(rules_path: Path, owners: list[str]) -> bool:
    """
    Appends one OwnershipRules row per not-yet-covered Owner, so the
    transaction-level Owner field actually gets populated from the first
    pipeline run - not just SourceAccount in config/settings.yaml. Never
    touches any other sheet in the workbook; creates OwnershipRules itself,
    with the real schema, if the workbook doesn't have it yet (a brand-new
    TransactionRules.xlsx copied from the template won't). Same
    show-then-confirm pattern as write_settings() - and the same backup
    safety net via replace_with_fresh_workbook().
    """
    covered = _existing_ownership_patterns(rules_path)
    seen: set[str] = set()
    next_number = _next_ownership_rule_number(rules_path)

    to_add: list[dict[str, Any]] = []
    for owner in owners:
        key = owner.strip().upper()
        if not key or key in covered or key in seen:
            continue
        seen.add(key)
        to_add.append(_build_ownership_rule_row(next_number, key))
        next_number += 1

    if not to_add:
        return False

    print()
    print(f"About to add {len(to_add)} OwnershipRules row(s) to {rules_path.name} - without these, these")
    print("accounts' transactions would have no Owner set on the first pipeline run:")
    for row in to_add:
        print(f"  - {row['RuleID']}: SourceAccount = {row['Pattern']}  ->  Owner = {row['SetOwner']}")

    if not ask_yes_no(f"Add these rows to {rules_path.name}?", default=True):
        print("Not written - nothing changed.")
        return False

    sheets = read_workbook_values_only(rules_path) if rules_path.exists() else {}

    existing_records: list[dict[str, Any]] = []
    if OWNERSHIP_RULES_SHEET in sheets:
        _, existing_records = sheet_values_to_records(sheets[OWNERSHIP_RULES_SHEET])

    sheets[OWNERSHIP_RULES_SHEET] = records_to_sheet_values(OWNERSHIP_RULES_COLUMNS, existing_records + to_add)

    if rules_path.exists():
        replace_with_fresh_workbook(
            rules_path,
            sheets,
            backup_label="before_setup_wizard_ownership_rules",
            basic_formatting=True,
            excel_tables=False,
        )
    else:
        write_fresh_workbook(rules_path, sheets, basic_formatting=True, excel_tables=False)

    print(f"Written: {rules_path}")
    return True


def offer_template_copy(
    scope: set[str],
    *,
    budgeting_template: Path = DEFAULT_BUDGETING_TEMPLATE,
    budgeting_rules: Path = DEFAULT_BUDGETING_RULES,
    investment_template: Path = DEFAULT_INVESTMENT_TEMPLATE,
    investment_rules: Path = DEFAULT_INVESTMENT_RULES,
) -> list[Path]:
    copied: list[Path] = []

    if "budgeting" in scope and not budgeting_rules.exists() and budgeting_template.exists():
        print()
        if ask_yes_no(f"Copy {budgeting_template.name} to {budgeting_rules.name} to get started?", default=True):
            budgeting_rules.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(budgeting_template, budgeting_rules)
            copied.append(budgeting_rules)

    if "investments" in scope and not investment_rules.exists() and investment_template.exists():
        print()
        if ask_yes_no(f"Copy {investment_template.name} to {investment_rules.name} to get started?", default=True):
            investment_rules.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(investment_template, investment_rules)
            copied.append(investment_rules)

    return copied


def load_existing_raw_settings(settings_path: Path) -> dict[str, Any]:
    if not settings_path.exists():
        return {}
    with settings_path.open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle) or {}


def write_settings(overrides: dict[str, Any], settings_path: Path = DEFAULT_SETTINGS_PATH) -> bool:
    """
    Never blindly overwrites an existing config/settings.yaml - deep-merges
    the newly collected answers onto whatever is already there (so re-running
    the wizard only touches what changed), shows the exact result, and
    requires an explicit confirmation before writing anything.

    Returns True if the file was written.
    """
    if not overrides:
        print()
        print("No settings collected - nothing to write.")
        return False

    existing = load_existing_raw_settings(settings_path)
    merged = _deep_merge(existing, overrides)

    print()
    print(f"About to write {settings_path}:")
    print("-" * 60)
    print(yaml.safe_dump(merged, sort_keys=False, allow_unicode=True).rstrip())
    print("-" * 60)

    if not ask_yes_no("Write this to config/settings.yaml?", default=True):
        print("Not written - nothing changed.")
        return False

    settings_path.parent.mkdir(parents=True, exist_ok=True)
    with settings_path.open("w", encoding="utf-8") as handle:
        yaml.safe_dump(merged, handle, sort_keys=False, allow_unicode=True)

    print(f"Written: {settings_path}")
    return True


def print_settings(settings_path: Path = DEFAULT_SETTINGS_PATH) -> None:
    """
    Print the full effective settings (example defaults deep-merged with
    your real config/settings.yaml, exactly as AppSettings.load() sees it
    at runtime) as YAML. Read-only - never prompts, never writes anything.
    """
    settings = AppSettings.load(settings_path=settings_path)
    print(f"Effective settings (example defaults + {settings_path}):")
    print("-" * 60)
    print(yaml.safe_dump(settings.data, sort_keys=False, allow_unicode=True).rstrip())
    print("-" * 60)


def print_seeds(settings_path: Path = DEFAULT_SETTINGS_PATH) -> None:
    """
    Print every configured budgeting.account_balance_seeds entry, one
    account at a time, sorted oldest-first (reuses AppSettings.
    account_balance_seeds() so formatting/sorting stays identical to what
    build_monthly_account_balance.py actually uses). Read-only.
    """
    settings = AppSettings.load(settings_path=settings_path)
    configured = settings.get("budgeting", "account_balance_seeds", default={}) or {}

    if not configured:
        print("No account_balance_seeds configured.")
        return

    print(f"Configured account balance seeds ({settings_path}):")
    print("-" * 60)
    for account in sorted(configured):
        seeds = settings.account_balance_seeds(account)
        print(f"{account}:")
        for date_str, balance in seeds:
            print(f"  {date_str}: {balance:,.2f}")
    print("-" * 60)


def run_wizard(
    *,
    settings_path: Path = DEFAULT_SETTINGS_PATH,
    budgeting_template: Path = DEFAULT_BUDGETING_TEMPLATE,
    budgeting_rules: Path = DEFAULT_BUDGETING_RULES,
    budgeting_workbook: Path = DEFAULT_BUDGETING_WORKBOOK,
    investment_template: Path = DEFAULT_INVESTMENT_TEMPLATE,
    investment_rules: Path = DEFAULT_INVESTMENT_RULES,
    dry_run: bool = False,
) -> None:
    print("finance_parser setup wizard" + (" (DRY RUN)" if dry_run else ""))
    print("============================")
    if dry_run:
        print("Dry run: working against a disposable copy of your real settings/rules.")
        print(f"  settings: {settings_path}")
        print("Nothing in your real config/settings.yaml or rules/ will be touched.")
    else:
        print("This helps you configure config/settings.yaml for your own banks/brokers.")
        print("Nothing is written until you confirm at the end.")

    existing_settings = load_existing_raw_settings(settings_path)

    scope = choose_scope()
    overrides: dict[str, Any] = {}
    unsupported: list[str] = []
    account_owners: list[str] = []

    if "budgeting" in scope:
        print()
        print("Budgeting banks supported today:")
        chosen = ask_multi_select(budgeting_broker_names())
        for broker in chosen:
            handler = BUDGETING_SETTINGS_HANDLERS.get(broker)
            if handler:
                account_owners.extend(handler(overrides) or [])

        print()
        use_cash = ask_yes_no(
            "Do you want to track manually-entered cash-only transactions (e.g. a small loan fee "
            "or insurance payment that never appears in any bank export)?",
            default=False,
        )
        if use_cash:
            print("Good - the CASH source parser is already available, no extra settings needed for it.")

        seedable = accounts_needing_seed(budgeting_workbook)
        if seedable:
            print()
            print(
                "Some of your budgeting accounts don't have a running balance in their own "
                "export - you can seed a starting balance now so monthly account balances can "
                "be tracked (see build_monthly_account_balance.py). Skip this anytime and run "
                "python tools/seed_account_balance.py later instead."
            )
            if ask_yes_no("Seed an account balance now?", default=False):
                existing_app_settings = AppSettings(existing_settings)
                for i, account in enumerate(seedable):
                    existing_seeds = existing_app_settings.account_balance_seeds(account)
                    collect_account_balance_seed(
                        overrides,
                        account,
                        budgeting_workbook=budgeting_workbook,
                        existing_seeds=existing_seeds,
                    )
                    remaining = len(seedable) - i - 1
                    if remaining == 0:
                        break
                    if not ask_yes_no(f"Seed another account ({remaining} more available)?", default=False):
                        break

        unsupported.extend(
            collect_unsupported("Any other banks you use that weren't listed? (comma-separated, blank to skip)")
        )

    if "investments" in scope:
        print()
        print("Investment brokers supported today:")
        chosen = ask_multi_select(investment_broker_names())
        for broker in chosen:
            collect_investment_broker_settings(overrides, broker, existing_settings)

        unsupported.extend(
            collect_unsupported("Any other brokers you use that weren't listed? (comma-separated, blank to skip)")
        )

    if unsupported:
        print()
        print("Not currently supported: " + ", ".join(unsupported))
        print(f"Please open a request here: {GITHUB_ISSUES_URL}")

    copied = offer_template_copy(
        scope,
        budgeting_template=budgeting_template,
        budgeting_rules=budgeting_rules,
        investment_template=investment_template,
        investment_rules=investment_rules,
    )
    write_settings(overrides, settings_path=settings_path)
    write_ownership_rules(budgeting_rules, account_owners)

    print()
    print("Setup wizard dry run complete." if dry_run else "Setup wizard complete.")
    if copied:
        print("Rule workbook(s) created from template:")
        for path in copied:
            print(f"  - {path}")
    if dry_run:
        print()
        print("Dry run only: your real config/settings.yaml and rules/ were not touched.")
        return

    print()
    print("Next steps:")
    if "budgeting" in scope:
        print("  1. Put your bank export files into input/budgeting/")
        print("  2. python run_budgeting_pipeline.py --dry-run")
    if "investments" in scope:
        print("  3. Put your broker export files into input/investments/")
        print("  4. python run_investment_pipeline.py --dry-run")


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Interactively configure config/settings.yaml for your own banks/brokers."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Run against a temporary sandbox copy of your real settings.yaml and rule workbooks "
             "(if they exist) instead of the real files - lets you try the wizard without any risk "
             "of overwriting your real configuration. The sandbox is deleted afterward unless "
             "--keep-temp is also given.",
    )
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        help="With --dry-run, keep the temporary sandbox directory for inspection afterward instead "
             "of deleting it.",
    )
    parser.add_argument(
        "--show-settings",
        action="store_true",
        help="Print the full effective settings (example defaults merged with your real "
             "config/settings.yaml) as YAML and exit. Read-only - does not run the wizard.",
    )
    parser.add_argument(
        "--show-seeds",
        action="store_true",
        help="Print every configured budgeting.account_balance_seeds entry and exit. "
             "Read-only - does not run the wizard.",
    )
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()

    if args.show_settings:
        print_settings()
        return

    if args.show_seeds:
        print_seeds()
        return

    if args.keep_temp and not args.dry_run:
        raise ValueError("--keep-temp can only be used together with --dry-run")

    if not args.dry_run:
        run_wizard()
        return

    temp_dir = Path(tempfile.mkdtemp(prefix="finance_parser_setup_wizard_dry_run_"))
    sandbox_settings = temp_dir / "settings.yaml"
    sandbox_budgeting_rules = temp_dir / "rules" / "budgeting" / DEFAULT_BUDGETING_RULES.name
    sandbox_investment_rules = temp_dir / "rules" / "investments" / DEFAULT_INVESTMENT_RULES.name
    sandbox_budgeting_rules.parent.mkdir(parents=True, exist_ok=True)
    sandbox_investment_rules.parent.mkdir(parents=True, exist_ok=True)

    if DEFAULT_SETTINGS_PATH.exists():
        shutil.copy2(DEFAULT_SETTINGS_PATH, sandbox_settings)
    if DEFAULT_BUDGETING_RULES.exists():
        shutil.copy2(DEFAULT_BUDGETING_RULES, sandbox_budgeting_rules)
    if DEFAULT_INVESTMENT_RULES.exists():
        shutil.copy2(DEFAULT_INVESTMENT_RULES, sandbox_investment_rules)

    try:
        run_wizard(
            settings_path=sandbox_settings,
            budgeting_template=DEFAULT_BUDGETING_TEMPLATE,
            budgeting_rules=sandbox_budgeting_rules,
            investment_template=DEFAULT_INVESTMENT_TEMPLATE,
            investment_rules=sandbox_investment_rules,
            dry_run=True,
        )
    finally:
        if args.keep_temp:
            print()
            print(f"Sandbox kept for inspection: {temp_dir}")
        else:
            shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
