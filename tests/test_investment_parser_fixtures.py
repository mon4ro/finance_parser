from pathlib import Path

import pytest

from investments.parsers import nordnet, seligson, evli, op_investment
from finance_parser import settings as settings_module
from finance_parser.settings import AppSettings


FIXTURES = Path(__file__).resolve().parent / "fixtures" / "investments"


def test_nordnet_fixture_parses_buy_transaction():
    path = FIXTURES / "nordnet_sample.csv"
    ok, reason = nordnet.can_parse(path)
    assert ok, reason

    df = nordnet.parse_file(path, "2026-06-04 12:00:00")
    row = df.iloc[0]

    assert row["Broker"] == "NORDNET"
    assert row["Portfolio"] == "20429585"
    assert row["TransactionTypeRaw"] == "OSTO"
    assert row["InstrumentName"] == "Nordnet Suomi Indeksi"
    assert row["ISIN"] == "SE0005993102"
    assert row["CashAmount"] == -40.0


def test_seligson_fixture_parses_manual_excel():
    path = FIXTURES / "seligson_sample.xlsx"
    ok, reason = seligson.can_parse(path)
    assert ok, reason

    df = seligson.parse_file(path, "2026-06-04 12:00:00")
    row = df.iloc[0]

    assert row["Broker"] == "SELIGSON"
    assert row["Portfolio"] == "58258"
    assert row["InstrumentName"] == "SELIGSON 58258"
    assert row["TransactionTypeRaw"] == "MERKINTÄ"
    assert row["CashAmount"] == 25.0


def test_evli_fixture_parses_nokia_dividend():
    path = FIXTURES / "evli_sample.xlsx"
    ok, reason = evli.can_parse(path)
    assert ok, reason

    df = evli.parse_file(path, "2026-06-04 12:00:00")
    row = df.iloc[0]

    assert row["Broker"] == "EVLI"
    assert row["Portfolio"] == "SIS Dividend"
    assert row["InstrumentName"] == "NOKIA"
    assert row["TransactionTypeRaw"] == "Dividend"
    assert row["CashAmount"] == 17.88


def test_op_investment_fixture_parses_dividend():
    path = FIXTURES / "op_investment_sample.xlsx"
    ok, reason = op_investment.can_parse(path)
    assert ok, reason

    df = op_investment.parse_file(path, "2026-06-04 12:00:00")
    row = df.iloc[0]

    assert row["Broker"] == "OP"
    assert row["InstrumentName"] == "NOKIA OYJ"
    assert row["TransactionTypeRaw"] == "OSINKO"
    assert row["CashAmount"] == 13.92


def test_investment_policy_can_change_through_settings(tmp_path):
    custom_settings = tmp_path / "settings.yaml"
    custom_settings.write_text(
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
  broker_defaults:
    EVLI:
      fixed_instrument_name: "ACME"
      portfolio_from_export_column: "Instrument"
    SELIGSON:
      instrument_name_template: "CUSTOM FUND {portfolio}"
""",
        encoding="utf-8",
    )

    loaded = AppSettings.load(settings_path=custom_settings, example_path=Path("does-not-exist.yaml"))

    old_cache = settings_module._SETTINGS_CACHE
    try:
        settings_module._SETTINGS_CACHE = loaded

        evli_df = evli.parse_file(FIXTURES / "evli_sample.xlsx", "2026-06-04 12:00:00")
        assert evli_df.iloc[0]["InstrumentName"] == "ACME"

        seligson_df = seligson.parse_file(FIXTURES / "seligson_sample.xlsx", "2026-06-04 12:00:00")
        assert seligson_df.iloc[0]["InstrumentName"] == "CUSTOM FUND 58258"
    finally:
        settings_module._SETTINGS_CACHE = old_cache
