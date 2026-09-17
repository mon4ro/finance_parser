import yaml

from finance_parser import setup_wizard as wizard


def _feed(monkeypatch, answers):
    """Replay a canned list of answers for successive input() calls."""
    queue = list(answers)

    def fake_input(prompt=""):
        if not queue:
            raise AssertionError(f"Ran out of scripted answers at prompt: {prompt!r}")
        return queue.pop(0)

    monkeypatch.setattr("builtins.input", fake_input)


def test_budgeting_broker_names_excludes_cash_and_synthetic_parser():
    names = wizard.budgeting_broker_names()

    assert "CASH" not in names
    assert "INVESTMENT_DIVIDEND" not in names
    assert "OP" in names
    assert "SPANKKI" in names


def test_investment_broker_names_are_live_from_registry():
    names = wizard.investment_broker_names()

    assert "NORDNET" in names
    assert "EVLI" in names
    assert "OP" in names


def test_ask_returns_default_on_blank_input(monkeypatch):
    _feed(monkeypatch, [""])
    assert wizard.ask("Label", default="HOUSEHOLD") == "HOUSEHOLD"


def test_ask_yes_no_defaults_and_parses_yes(monkeypatch):
    _feed(monkeypatch, [""])
    assert wizard.ask_yes_no("Buffer account?", default=True) is True

    _feed(monkeypatch, ["y"])
    assert wizard.ask_yes_no("Buffer account?", default=False) is True

    _feed(monkeypatch, ["n"])
    assert wizard.ask_yes_no("Buffer account?", default=True) is False


def test_ask_multi_select_parses_numbers_and_ignores_junk(monkeypatch):
    _feed(monkeypatch, ["1, 3, banana, 99"])
    chosen = wizard.ask_multi_select(["OP", "NORDEA", "SPANKKI"])

    assert chosen == ["OP", "SPANKKI"]


def test_choose_scope_maps_menu_choices(monkeypatch):
    _feed(monkeypatch, ["1"])
    assert wizard.choose_scope() == {"budgeting"}

    _feed(monkeypatch, ["2"])
    assert wizard.choose_scope() == {"investments"}

    _feed(monkeypatch, [""])
    assert wizard.choose_scope() == {"budgeting", "investments"}


def test_collect_op_budgeting_settings_builds_filename_prefixes(monkeypatch):
    # count=2, then label/prefix pairs for each account
    _feed(monkeypatch, ["2", "household", "", "personal", "personb"])
    overrides = {}

    wizard.collect_op_budgeting_settings(overrides)

    assert overrides["budgeting"]["source_account_inference"]["OP"]["filename_prefixes"] == {
        "HOUSEHOLD": "HOUSEHOLD",
        "PERSONAL": "PERSONB",
    }


def test_collect_spankki_settings_maps_buffer_answer_to_default_include(monkeypatch):
    _feed(monkeypatch, [""])  # default Y (buffer account)
    overrides = {}
    wizard.collect_spankki_settings(overrides)

    spankki = overrides["budgeting"]["source_account_inference"]["SPANKKI"]
    assert spankki["fixed_source_account"] == "SPANKKI"
    assert spankki["default_include"] == "NO"

    _feed(monkeypatch, ["n"])  # not a buffer account
    overrides2 = {}
    wizard.collect_spankki_settings(overrides2)
    assert overrides2["budgeting"]["source_account_inference"]["SPANKKI"]["default_include"] == "YES"


def test_collect_portfolio_owners_single_portfolio(monkeypatch):
    _feed(monkeypatch, ["n", "personal"])
    overrides = {}
    wizard.collect_portfolio_owners(overrides, "EVLI")

    assert overrides["investments"]["portfolio_owners"]["EVLI"] == "PERSONAL"


def test_collect_portfolio_owners_multiple_portfolios(monkeypatch):
    _feed(monkeypatch, ["y", "10000001", "personal", "", "household"])
    overrides = {}
    wizard.collect_portfolio_owners(overrides, "NORDNET")

    assert overrides["investments"]["portfolio_owners"]["NORDNET"] == {
        "10000001": "PERSONAL",
        "default": "HOUSEHOLD",
    }


def test_collect_investment_broker_settings_evli_asks_fixed_instrument_name(monkeypatch):
    _feed(monkeypatch, ["nokia", "n", "personal"])
    overrides = {}
    wizard.collect_investment_broker_settings(overrides, "EVLI")

    assert overrides["investments"]["broker_defaults"]["EVLI"]["fixed_instrument_name"] == "NOKIA"
    assert overrides["investments"]["portfolio_owners"]["EVLI"] == "PERSONAL"


def test_offer_template_copy_only_copies_missing_files(tmp_path, monkeypatch):
    budgeting_template = tmp_path / "TransactionRules.template.xlsx"
    budgeting_template.write_bytes(b"template")
    budgeting_rules = tmp_path / "TransactionRules.xlsx"

    investment_template = tmp_path / "InstrumentMaster.template.xlsx"
    investment_template.write_bytes(b"template")
    investment_rules = tmp_path / "InstrumentMaster.xlsx"
    investment_rules.write_bytes(b"already exists")  # should NOT be overwritten

    _feed(monkeypatch, ["y"])  # only budgeting gets prompted, investment_rules already exists

    copied = wizard.offer_template_copy(
        {"budgeting", "investments"},
        budgeting_template=budgeting_template,
        budgeting_rules=budgeting_rules,
        investment_template=investment_template,
        investment_rules=investment_rules,
    )

    assert copied == [budgeting_rules]
    assert budgeting_rules.read_bytes() == b"template"
    assert investment_rules.read_bytes() == b"already exists"


def test_write_settings_merges_onto_existing_file_without_dropping_untouched_keys(tmp_path, monkeypatch):
    settings_path = tmp_path / "settings.yaml"
    settings_path.write_text(yaml.safe_dump({"budgeting": {"default_include": "YES"}}), encoding="utf-8")

    _feed(monkeypatch, ["y"])
    overrides = {"budgeting": {"source_account_inference": {"OP": {"filename_prefixes": {"HOUSEHOLD": "HOUSEHOLD"}}}}}

    written = wizard.write_settings(overrides, settings_path=settings_path)

    assert written is True
    result = yaml.safe_load(settings_path.read_text(encoding="utf-8"))
    assert result["budgeting"]["default_include"] == "YES"  # untouched key survives
    assert result["budgeting"]["source_account_inference"]["OP"]["filename_prefixes"]["HOUSEHOLD"] == "HOUSEHOLD"


def test_write_settings_does_not_write_when_user_declines(tmp_path, monkeypatch):
    settings_path = tmp_path / "settings.yaml"
    _feed(monkeypatch, ["n"])

    written = wizard.write_settings({"budgeting": {"default_include": "YES"}}, settings_path=settings_path)

    assert written is False
    assert not settings_path.exists()


def test_write_settings_with_no_overrides_does_not_prompt(tmp_path):
    settings_path = tmp_path / "settings.yaml"

    written = wizard.write_settings({}, settings_path=settings_path)

    assert written is False
    assert not settings_path.exists()
