from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

import yaml

from finance_parser.budgeting.transaction_parser import SUPPORTED_PARSERS as BUDGETING_PARSER_MODULES
from finance_parser.investments.investment_parser import SUPPORTED_INVESTMENT_PARSERS as INVESTMENT_PARSER_MODULES
from finance_parser.settings import _deep_merge


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


def ask(prompt: str, *, default: str | None = None) -> str:
    suffix = f" [{default}]" if default else ""
    raw = input(f"{prompt}{suffix}: ").strip()
    return raw if raw else (default or "")


def ask_yes_no(prompt: str, *, default: bool = False) -> bool:
    hint = "Y/n" if default else "y/N"
    raw = input(f"{prompt} [{hint}]: ").strip().lower()
    if not raw:
        return default
    return raw in {"y", "yes"}


def ask_multi_select(options: list[str]) -> list[str]:
    for i, option in enumerate(options, start=1):
        print(f"  {i}. {option}")
    raw = input("Enter numbers separated by commas (blank for none): ").strip()
    if not raw:
        return []

    chosen: list[str] = []
    for part in raw.split(","):
        part = part.strip()
        if not part.isdigit():
            continue
        index = int(part) - 1
        if 0 <= index < len(options):
            chosen.append(options[index])
    return chosen


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


def collect_op_budgeting_settings(overrides: dict[str, Any]) -> None:
    print()
    print("OP (budgeting): each OP export file is matched to an account/owner by a prefix in")
    print("its filename, e.g. HOUSEHOLD_tapahtumat20260503-20260603.csv -> HOUSEHOLD.")
    count_raw = ask("How many separate OP export accounts do you have", default="1")
    try:
        count = max(1, int(count_raw))
    except ValueError:
        count = 1

    prefixes: dict[str, str] = {}
    for i in range(count):
        label = ask(f"  Account #{i + 1} label (e.g. HOUSEHOLD, PERSONAL, CHILD)")
        if not label:
            continue
        prefix = ask(f"  Filename prefix for '{label.upper()}'", default=label)
        prefixes[label.upper()] = prefix.upper()

    if prefixes:
        _set_nested(overrides, ["budgeting", "source_account_inference", "OP", "filename_prefixes"], prefixes)


def collect_fixed_account_budgeting_settings(overrides: dict[str, Any], broker: str) -> None:
    label = ask(f"{broker}: account/owner label for these transactions", default=broker)
    if label:
        _set_nested(overrides, ["budgeting", "source_account_inference", broker, "fixed_source_account"], label.upper())


def collect_spankki_settings(overrides: dict[str, Any]) -> None:
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


BUDGETING_SETTINGS_HANDLERS = {
    "OP": collect_op_budgeting_settings,
    "NORWEGIAN": lambda overrides: collect_fixed_account_budgeting_settings(overrides, "NORWEGIAN"),
    "NORDEA": lambda overrides: collect_fixed_account_budgeting_settings(overrides, "NORDEA"),
    "SPANKKI": collect_spankki_settings,
}


def collect_portfolio_owners(overrides: dict[str, Any], broker: str) -> None:
    print()
    multiple = ask_yes_no(f"{broker}: do you have more than one portfolio/account under this broker?", default=False)

    if not multiple:
        owner = ask(f"{broker}: who does this portfolio belong to?", default="HOUSEHOLD")
        if owner:
            _set_nested(overrides, ["investments", "portfolio_owners", broker], owner.upper())
        return

    owners: dict[str, str] = {}
    while True:
        portfolio = ask("  Portfolio number/id (blank to finish)")
        if not portfolio:
            break
        owner = ask(f"  Who does portfolio {portfolio} belong to?", default="HOUSEHOLD")
        owners[portfolio] = owner.upper()

    default_owner = ask(f"{broker}: default owner for any other, unlisted portfolio", default="HOUSEHOLD")
    if default_owner:
        owners["default"] = default_owner.upper()

    if owners:
        _set_nested(overrides, ["investments", "portfolio_owners", broker], owners)


def collect_fixed_instrument_name(overrides: dict[str, Any], broker: str, hint: str) -> None:
    print()
    print(hint)
    name = ask(f"{broker}: instrument/fund name for this account")
    if name:
        _set_nested(overrides, ["investments", "broker_defaults", broker, "fixed_instrument_name"], name.upper())


def collect_investment_broker_settings(overrides: dict[str, Any], broker: str) -> None:
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
    collect_portfolio_owners(overrides, broker)


def collect_unsupported(prompt: str) -> list[str]:
    print()
    raw = ask(prompt)
    if not raw:
        return []
    return [part.strip() for part in raw.split(",") if part.strip()]


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

    existing: dict[str, Any] = {}
    if settings_path.exists():
        with settings_path.open("r", encoding="utf-8") as handle:
            existing = yaml.safe_load(handle) or {}

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


def main() -> None:
    print("finance_parser setup wizard")
    print("============================")
    print("This helps you configure config/settings.yaml for your own banks/brokers.")
    print("Nothing is written until you confirm at the end.")

    scope = choose_scope()
    overrides: dict[str, Any] = {}
    unsupported: list[str] = []

    if "budgeting" in scope:
        print()
        print("Budgeting banks supported today:")
        chosen = ask_multi_select(budgeting_broker_names())
        for broker in chosen:
            handler = BUDGETING_SETTINGS_HANDLERS.get(broker)
            if handler:
                handler(overrides)

        print()
        use_cash = ask_yes_no(
            "Do you want to track manually-entered cash-only transactions (e.g. a small loan fee "
            "or insurance payment that never appears in any bank export)?",
            default=False,
        )
        if use_cash:
            print("Good - the CASH source parser is already available, no extra settings needed for it.")

        unsupported.extend(
            collect_unsupported("Any other banks you use that weren't listed? (comma-separated, blank to skip)")
        )

    if "investments" in scope:
        print()
        print("Investment brokers supported today:")
        chosen = ask_multi_select(investment_broker_names())
        for broker in chosen:
            collect_investment_broker_settings(overrides, broker)

        unsupported.extend(
            collect_unsupported("Any other brokers you use that weren't listed? (comma-separated, blank to skip)")
        )

    if unsupported:
        print()
        print("Not currently supported: " + ", ".join(unsupported))
        print(f"Please open a request here: {GITHUB_ISSUES_URL}")

    copied = offer_template_copy(scope)
    write_settings(overrides)

    print()
    print("Setup wizard complete.")
    if copied:
        print("Rule workbook(s) created from template:")
        for path in copied:
            print(f"  - {path}")
    print()
    print("Next steps:")
    if "budgeting" in scope:
        print("  1. Put your bank export files into input/budgeting/")
        print("  2. python run_budgeting_pipeline.py --dry-run")
    if "investments" in scope:
        print("  3. Put your broker export files into input/investments/")
        print("  4. python run_investment_pipeline.py --dry-run")


if __name__ == "__main__":
    main()
