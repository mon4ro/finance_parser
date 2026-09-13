from pathlib import Path

from finance_parser.settings import AppSettings, DEFAULT_EXAMPLE_SETTINGS_PATH


def test_example_settings_loads_and_validates():
    settings = AppSettings.load(
        settings_path=Path("does-not-exist-user-settings.yaml"),
        example_path=DEFAULT_EXAMPLE_SETTINGS_PATH,
    )

    assert settings.get("project", "default_currency") == "EUR"
    assert settings.path("budgeting_input").as_posix().endswith("input/budgeting")


def test_settings_canonicalises_source_bank_aliases():
    settings = AppSettings.load(
        settings_path=Path("does-not-exist-user-settings.yaml"),
        example_path=DEFAULT_EXAMPLE_SETTINGS_PATH,
    )

    assert settings.canonical_source_bank("BANK NORWEGIAN") == "NORWEGIAN"
    assert settings.canonical_source_bank("S-PANKKI") == "SPANKKI"


def test_settings_source_account_inference_from_filename():
    settings = AppSettings.load(
        settings_path=Path("does-not-exist-user-settings.yaml"),
        example_path=DEFAULT_EXAMPLE_SETTINGS_PATH,
    )

    assert settings.filename_source_account("OP", "HOUSEHOLD_tapahtumat20260501-20260603.csv") == "HOUSEHOLD"
    assert settings.filename_source_account("OP", "PERSONAL_tapahtumat20260503-20260603.csv") == "PERSONAL"
    assert settings.filename_source_account("OP", "CHILD_tapahtumat20260503-20260603-2.csv") == "CHILD"


def test_settings_fixed_source_accounts_and_defaults():
    settings = AppSettings.load(
        settings_path=Path("does-not-exist-user-settings.yaml"),
        example_path=DEFAULT_EXAMPLE_SETTINGS_PATH,
    )

    assert settings.fixed_source_account("NORDEA") == "PERSONAL"
    assert settings.fixed_source_account("NORWEGIAN") == "NORWEGIAN"
    assert settings.fixed_source_account("SPANKKI") == "SPANKKI"
    assert settings.default_include_for_source_bank("SPANKKI") == "NO"
    assert settings.default_include_for_source_bank("OP") == "YES"


def test_settings_investment_broker_defaults():
    settings = AppSettings.load(
        settings_path=Path("does-not-exist-user-settings.yaml"),
        example_path=DEFAULT_EXAMPLE_SETTINGS_PATH,
    )

    assert settings.broker_default("EVLI", "fixed_instrument_name") == "NOKIA"
    assert settings.broker_default("SELIGSON", "instrument_name_template") == "SELIGSON {portfolio}"
