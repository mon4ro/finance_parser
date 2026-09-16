import shutil
from pathlib import Path

import pandas as pd
import pytest

from investments.parsers import nordnet, seligson, evli, op_investment, coinmotion
from finance_parser import settings as settings_module
from finance_parser.investments.investment_common import INVESTMENT_IMPORT_LOG_SHEET
from finance_parser.investments.investment_parser import append_to_output
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
    assert row["Portfolio"] == "EVLI"
    assert row["PlanCycle"] == "SIS Dividend"
    assert row["InstrumentName"] == "NOKIA"
    assert row["TransactionTypeRaw"] == "Dividend"
    assert row["CashAmount"] == 17.88


def test_evli_raw_id_distinguishes_rows_by_plan_cycle_not_portfolio():
    """
    Real bug this replaced: Portfolio is now a constant ("EVLI"), so it can
    no longer be part of the row-ID hash - PlanCycle must be, or two
    otherwise-identical rows from different plan cycles would collide onto
    the same InvestmentRawID and one would silently vanish as a "duplicate".
    """
    base = {
        "TransactionTypeRaw": "Vesting", "Quantity": 10, "UnitPrice": 5.0,
        "CashAmount": 0.0, "Status": "Confirmed", "ValueDate": "2026-01-01",
    }
    row_a = pd.Series({**base, "PlanCycle": "Plan Cycle 2022"})
    row_b = pd.Series({**base, "PlanCycle": "Plan Cycle 2023"})

    assert evli.make_investment_raw_id(row_a) != evli.make_investment_raw_id(row_b)


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


def test_coinmotion_fixture_parses_deposit_and_buy():
    path = FIXTURES / "coinmotion_sample.csv"
    ok, reason = coinmotion.can_parse(path)
    assert ok, reason

    df = coinmotion.parse_file(path, "2026-06-04 12:00:00")

    deposit = df.iloc[0]
    assert deposit["Broker"] == "COINMOTION"
    assert deposit["TransactionTypeRaw"] == "DEPOSIT"
    assert deposit["InstrumentName"] == ""
    assert deposit["CashAmount"] == 1000.0
    assert deposit["TradeDate"] == "2026-03-01"

    buy = df.iloc[1]
    assert buy["TransactionTypeRaw"] == "BUY"
    assert buy["InstrumentName"] == "BTC"
    assert buy["Quantity"] == 0.02
    # Total EUR debited is eurAmount (fee-inclusive), not rate*quantity.
    assert buy["CashAmount"] == -1000.0
    assert buy["BrokerageFee"] == 20.0


def test_coinmotion_fixture_parses_sell_direction_from_currency_pair():
    """
    "market_trade" alone doesn't say which side is being sold - direction
    must come from which currency is EUR (fromCurrency=EUR -> BUY,
    toCurrency=EUR -> SELL).
    """
    path = FIXTURES / "coinmotion_sample.csv"
    df = coinmotion.parse_file(path, "2026-06-04 12:00:00")

    sell = df.iloc[2]
    assert sell["TransactionTypeRaw"] == "SELL"
    assert sell["InstrumentName"] == "BTC"
    assert sell["Quantity"] == 0.01
    assert sell["CashAmount"] == 510.0


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


def test_append_to_output_reports_per_file_new_and_duplicate_counts(tmp_path):
    """
    Real bug: RowsNew/RowsDuplicate were computed once for the whole run and
    broadcast into every file's ImportLog row (same bug already fixed on the
    budgeting side - see transaction_parser.py). A run over two 1-row files
    must show RowsNew=1 (not 2) for each file on first import, and
    RowsDuplicate=1 (not 2) for each file on a rerun where nothing is new.
    """
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    shutil.copy(FIXTURES / "nordnet_sample.csv", input_dir / "nordnet_sample.csv")
    shutil.copy(FIXTURES / "seligson_sample.xlsx", input_dir / "seligson_sample.xlsx")

    output_path = tmp_path / "ParsedInvestments.xlsx"

    append_to_output(input_dir, output_path, None)

    log = pd.read_excel(output_path, sheet_name=INVESTMENT_IMPORT_LOG_SHEET, dtype=object)
    first_run_id = log["ImportRunID"].iloc[0]
    first_run = log[log["ImportRunID"] == first_run_id]
    new_by_file = dict(zip(first_run["SourceFile"], first_run["RowsNew"]))
    assert new_by_file["nordnet_sample.csv"] == 1
    assert new_by_file["seligson_sample.xlsx"] == 1

    append_to_output(input_dir, output_path, None)

    log2 = pd.read_excel(output_path, sheet_name=INVESTMENT_IMPORT_LOG_SHEET, dtype=object)
    second_run_id = log2["ImportRunID"].iloc[-1]
    second_run = log2[log2["ImportRunID"] == second_run_id]
    duplicate_by_file = dict(zip(second_run["SourceFile"], second_run["RowsDuplicate"]))
    assert duplicate_by_file["nordnet_sample.csv"] == 1
    assert duplicate_by_file["seligson_sample.xlsx"] == 1


def test_append_to_output_persists_skipped_status_not_overwritten_to_imported(tmp_path):
    """
    Real bug: append_to_output() used to end with an unconditional
    import_log_new["Status"] = "Imported" applied to every row, silently
    overwriting the "Skipped: unsupported file" / "Skipped: format drift"
    statuses parse_import_files() sets for files it never actually parsed -
    undoing the entire audit trail those failure modes exist to provide.
    A real, unsupported file sitting alongside valid ones must still show as
    skipped in the persisted ImportLog sheet, not silently relabelled
    "Imported".
    """
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    shutil.copy(FIXTURES / "nordnet_sample.csv", input_dir / "nordnet_sample.csv")
    (input_dir / "unrelated_reference.txt").write_text("not a broker export at all")

    output_path = tmp_path / "ParsedInvestments.xlsx"

    append_to_output(input_dir, output_path, None)

    log = pd.read_excel(output_path, sheet_name=INVESTMENT_IMPORT_LOG_SHEET, dtype=object)
    status_by_file = dict(zip(log["SourceFile"], log["Status"]))
    assert status_by_file["nordnet_sample.csv"] == "Imported"
    assert status_by_file["unrelated_reference.txt"] == "Skipped: unsupported file"
