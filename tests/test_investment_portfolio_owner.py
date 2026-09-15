from pathlib import Path

import pandas as pd

from finance_parser import settings as settings_module
from finance_parser.settings import AppSettings
from finance_parser.investments.investment_common import (
    INVESTMENT_RAW_COLUMNS,
    INVESTMENT_TRANSACTIONS_COLUMNS,
    apply_portfolio_ownership,
    raw_to_investment_transactions,
)


def _raw_row(**overrides) -> pd.DataFrame:
    row = {col: "" for col in INVESTMENT_RAW_COLUMNS}
    row.update(overrides)
    return pd.DataFrame([row], columns=INVESTMENT_RAW_COLUMNS)


def _with_settings(yaml_text: str, tmp_path: Path):
    settings_path = tmp_path / "settings.yaml"
    settings_path.write_text(yaml_text, encoding="utf-8")
    return AppSettings.load(settings_path=settings_path, example_path=Path("does-not-exist.yaml"))


def test_portfolio_owner_and_type_filled_from_settings_when_blank(tmp_path):
    loaded = _with_settings(
        """
project:
  name: "Test"
  locale: "fi_FI"
  default_currency: "EUR"

paths:
  budgeting_input: "input/budgeting"
  budgeting_output: "output/budgeting/ParsedTransactions.xlsx"
  budgeting_rules: "rules/budgeting/TransactionRules.xlsx"
  investment_input: "input/investments"
  investment_output: "output/investments/ParsedInvestments.xlsx"
  investment_rules: "rules/investments/InstrumentMaster.xlsx"

budgeting:
  default_include: "YES"
  source_bank_aliases: {}
  source_account_inference: {}

investments:
  portfolio_owners:
    EVLI: "PERSON_A"
  portfolio_types:
    EVLI: "Arvo-osuustili"
""",
        tmp_path,
    )

    old_cache = settings_module._SETTINGS_CACHE
    try:
        settings_module._SETTINGS_CACHE = loaded

        raw = _raw_row(Broker="EVLI", Portfolio="Plan Cycle 2023", PortfolioOwner="", PortfolioType="")
        result = raw_to_investment_transactions(raw)

        assert result.iloc[0]["PortfolioOwner"] == "PERSON_A"
        assert result.iloc[0]["PortfolioType"] == "ARVO-OSUUSTILI"
    finally:
        settings_module._SETTINGS_CACHE = old_cache


def test_portfolio_owner_from_parser_is_not_overridden_by_settings(tmp_path):
    """
    If a parser ever does supply PortfolioOwner directly, settings must not
    clobber it - settings only fill a blank, matching how Owner defaults work
    on the budgeting side.
    """
    loaded = _with_settings(
        """
project:
  name: "Test"
  locale: "fi_FI"
  default_currency: "EUR"

paths:
  budgeting_input: "input/budgeting"
  budgeting_output: "output/budgeting/ParsedTransactions.xlsx"
  budgeting_rules: "rules/budgeting/TransactionRules.xlsx"
  investment_input: "input/investments"
  investment_output: "output/investments/ParsedInvestments.xlsx"
  investment_rules: "rules/investments/InstrumentMaster.xlsx"

budgeting:
  default_include: "YES"
  source_bank_aliases: {}
  source_account_inference: {}

investments:
  portfolio_owners:
    EVLI: "PERSON_A"
""",
        tmp_path,
    )

    old_cache = settings_module._SETTINGS_CACHE
    try:
        settings_module._SETTINGS_CACHE = loaded

        raw = _raw_row(Broker="EVLI", Portfolio="Plan Cycle 2023", PortfolioOwner="PERSON_B", PortfolioType="")
        result = raw_to_investment_transactions(raw)

        assert result.iloc[0]["PortfolioOwner"] == "PERSON_B"
    finally:
        settings_module._SETTINGS_CACHE = old_cache


def test_portfolio_owner_blank_when_no_settings_configured(tmp_path):
    loaded = _with_settings(
        """
project:
  name: "Test"
  locale: "fi_FI"
  default_currency: "EUR"

paths:
  budgeting_input: "input/budgeting"
  budgeting_output: "output/budgeting/ParsedTransactions.xlsx"
  budgeting_rules: "rules/budgeting/TransactionRules.xlsx"
  investment_input: "input/investments"
  investment_output: "output/investments/ParsedInvestments.xlsx"
  investment_rules: "rules/investments/InstrumentMaster.xlsx"

budgeting:
  default_include: "YES"
  source_bank_aliases: {}
  source_account_inference: {}

investments: {}
""",
        tmp_path,
    )

    old_cache = settings_module._SETTINGS_CACHE
    try:
        settings_module._SETTINGS_CACHE = loaded

        raw = _raw_row(Broker="UNKNOWN_BROKER", Portfolio="X", PortfolioOwner="", PortfolioType="")
        result = raw_to_investment_transactions(raw)

        assert result.iloc[0]["PortfolioOwner"] == ""
        assert result.iloc[0]["PortfolioType"] == ""
    finally:
        settings_module._SETTINGS_CACHE = old_cache


def test_apply_portfolio_ownership_backfills_existing_rows(tmp_path):
    """
    Real scenario: 1083+ pre-existing InvestmentTransactions rows were
    written before portfolio_owners existed in settings.yaml, so they're
    blank. apply_portfolio_ownership() must backfill them retroactively when
    re-run, the same way apply_instrument_master() backfills ISIN - not just
    apply to brand-new rows.
    """
    loaded = _with_settings(
        """
project:
  name: "Test"
  locale: "fi_FI"
  default_currency: "EUR"

paths:
  budgeting_input: "input/budgeting"
  budgeting_output: "output/budgeting/ParsedTransactions.xlsx"
  budgeting_rules: "rules/budgeting/TransactionRules.xlsx"
  investment_input: "input/investments"
  investment_output: "output/investments/ParsedInvestments.xlsx"
  investment_rules: "rules/investments/InstrumentMaster.xlsx"

budgeting:
  default_include: "YES"
  source_bank_aliases: {}
  source_account_inference: {}

investments:
  portfolio_owners:
    EVLI: "PERSON_A"
  portfolio_types:
    EVLI: "Arvo-osuustili"
""",
        tmp_path,
    )

    old_cache = settings_module._SETTINGS_CACHE
    try:
        settings_module._SETTINGS_CACHE = loaded

        existing_row = {col: "" for col in INVESTMENT_TRANSACTIONS_COLUMNS}
        existing_row.update({"Broker": "EVLI", "Portfolio": "Plan Cycle 2023"})
        transactions = pd.DataFrame([existing_row], columns=INVESTMENT_TRANSACTIONS_COLUMNS)

        result = apply_portfolio_ownership(transactions)

        assert result.iloc[0]["PortfolioOwner"] == "PERSON_A"
        assert result.iloc[0]["PortfolioType"] == "ARVO-OSUUSTILI"
    finally:
        settings_module._SETTINGS_CACHE = old_cache


def test_apply_portfolio_ownership_does_not_overwrite_existing_manual_value(tmp_path):
    loaded = _with_settings(
        """
project:
  name: "Test"
  locale: "fi_FI"
  default_currency: "EUR"

paths:
  budgeting_input: "input/budgeting"
  budgeting_output: "output/budgeting/ParsedTransactions.xlsx"
  budgeting_rules: "rules/budgeting/TransactionRules.xlsx"
  investment_input: "input/investments"
  investment_output: "output/investments/ParsedInvestments.xlsx"
  investment_rules: "rules/investments/InstrumentMaster.xlsx"

budgeting:
  default_include: "YES"
  source_bank_aliases: {}
  source_account_inference: {}

investments:
  portfolio_owners:
    EVLI: "PERSON_A"
""",
        tmp_path,
    )

    old_cache = settings_module._SETTINGS_CACHE
    try:
        settings_module._SETTINGS_CACHE = loaded

        existing_row = {col: "" for col in INVESTMENT_TRANSACTIONS_COLUMNS}
        existing_row.update({"Broker": "EVLI", "Portfolio": "Plan Cycle 2023", "PortfolioOwner": "PERSON_B"})
        transactions = pd.DataFrame([existing_row], columns=INVESTMENT_TRANSACTIONS_COLUMNS)

        result = apply_portfolio_ownership(transactions)

        assert result.iloc[0]["PortfolioOwner"] == "PERSON_B"
    finally:
        settings_module._SETTINGS_CACHE = old_cache
