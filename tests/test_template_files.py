from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def test_budgeting_transaction_rules_template_has_expected_headers():
    path = PROJECT_ROOT / "rules" / "budgeting" / "TransactionRules.template.xlsx"
    assert path.exists()

    df = pd.read_excel(path, sheet_name="TransactionRules", nrows=0, engine="openpyxl")
    expected = {
        "RuleID",
        "Enabled",
        "Priority",
        "FieldToSearch",
        "MatchType",
        "Pattern",
        "SetNormalizedReceiver",
        "SetInclude",
        "OverwriteMode",
    }
    assert expected.issubset(set(df.columns))


def test_investment_instrument_master_template_has_expected_headers():
    path = PROJECT_ROOT / "rules" / "investments" / "InstrumentMaster.template.xlsx"
    assert path.exists()

    df = pd.read_excel(path, sheet_name="InstrumentMaster", nrows=0, engine="openpyxl")
    expected = {
        "RuleID",
        "Enabled",
        "Priority",
        "RawInstrumentPattern",
        "MatchType",
        "NormalizedInstrument",
        "InstrumentType",
    }
    assert expected.issubset(set(df.columns))
