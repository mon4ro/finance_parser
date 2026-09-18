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


def test_settings_portfolio_owner_plain_string_applies_to_whole_broker():
    settings = AppSettings.load(
        settings_path=Path("does-not-exist-user-settings.yaml"),
        example_path=DEFAULT_EXAMPLE_SETTINGS_PATH,
    )

    assert settings.portfolio_owner("EVLI", "Plan Cycle 2023") == "PERSONAL"
    assert settings.portfolio_owner("EVLI", "SIS Dividend") == "PERSONAL"


def test_settings_portfolio_owner_per_portfolio_mapping_with_default_fallback():
    settings = AppSettings.load(
        settings_path=Path("does-not-exist-user-settings.yaml"),
        example_path=DEFAULT_EXAMPLE_SETTINGS_PATH,
    )

    assert settings.portfolio_owner("NORDNET", "20429585") == "PERSONAL"
    assert settings.portfolio_owner("NORDNET", "99999999") == "HOUSEHOLD"


def test_settings_portfolio_owner_unknown_broker_returns_blank():
    settings = AppSettings.load(
        settings_path=Path("does-not-exist-user-settings.yaml"),
        example_path=DEFAULT_EXAMPLE_SETTINGS_PATH,
    )

    assert settings.portfolio_owner("SOME_UNCONFIGURED_BROKER", "X") == ""


def test_settings_portfolio_type_lookup(tmp_path):
    # Deliberately not testing this against the real example file:
    # settings.example.yaml intentionally has no portfolio_types example
    # value, because config/settings.yaml only overrides keys it explicitly
    # sets - a real-looking example value here would silently leak into any
    # deployment that configures a different broker under portfolio_types
    # without also overriding this one (real incident: an example "EVLI:
    # Arvo-osuustili" value leaked into this project's own real settings.yaml
    # this way before the example was emptied out).
    settings_path = tmp_path / "settings.yaml"
    settings_path.write_text(
        """
investments:
  portfolio_types:
    EVLI: "Arvo-osuustili"
""",
        encoding="utf-8",
    )
    settings = AppSettings.load(settings_path=settings_path, example_path=DEFAULT_EXAMPLE_SETTINGS_PATH)

    assert settings.portfolio_type("EVLI", "Plan Cycle 2023") == "ARVO-OSUUSTILI"


def test_example_settings_has_no_portfolio_type_leak_risk():
    """
    portfolio_types in the example file must stay empty - any value here
    would silently apply to every deployment's real settings.yaml unless
    that deployment happens to override the exact same broker key.
    """
    settings = AppSettings.load(
        settings_path=Path("does-not-exist-user-settings.yaml"),
        example_path=DEFAULT_EXAMPLE_SETTINGS_PATH,
    )

    assert settings.get("investments", "portfolio_types", default={}) == {}


def test_settings_account_balance_seeds_sorted_oldest_first(tmp_path):
    settings_path = tmp_path / "settings.yaml"
    settings_path.write_text(
        """
budgeting:
  account_balance_seeds:
    HOUSEHOLD:
      "2024-06-01": 500.0
      "2023-01-15": 100.0
""",
        encoding="utf-8",
    )
    settings = AppSettings.load(settings_path=settings_path, example_path=DEFAULT_EXAMPLE_SETTINGS_PATH)

    assert settings.account_balance_seeds("HOUSEHOLD") == [
        ("2023-01-15", 100.0),
        ("2024-06-01", 500.0),
    ]


def test_settings_account_balance_seeds_unknown_account_returns_empty(tmp_path):
    settings_path = tmp_path / "settings.yaml"
    settings_path.write_text("budgeting:\n  account_balance_seeds: {}\n", encoding="utf-8")
    settings = AppSettings.load(settings_path=settings_path, example_path=DEFAULT_EXAMPLE_SETTINGS_PATH)

    assert settings.account_balance_seeds("SOME_UNCONFIGURED_ACCOUNT") == []


def test_example_settings_has_no_account_balance_seeds_leak_risk():
    """
    Same real incident class as portfolio_types above: account_balance_seeds
    in the example file must stay absent/empty - a real-looking example
    value here would silently apply to every deployment's real balance
    calculation unless that deployment happens to override the exact same
    account key.
    """
    settings = AppSettings.load(
        settings_path=Path("does-not-exist-user-settings.yaml"),
        example_path=DEFAULT_EXAMPLE_SETTINGS_PATH,
    )

    assert settings.get("budgeting", "account_balance_seeds", default={}) == {}
