from pathlib import Path

import pandas as pd
import pytest
import yaml
from openpyxl import Workbook

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
    # count=2, then owner/prefix pairs for each account
    _feed(monkeypatch, ["2", "household", "", "personal", "personb"])
    overrides = {}

    owners = wizard.collect_op_budgeting_settings(overrides)

    assert overrides["budgeting"]["source_account_inference"]["OP"]["filename_prefixes"] == {
        "HOUSEHOLD": "HOUSEHOLD",
        "PERSONAL": "PERSONB",
    }
    assert owners == ["HOUSEHOLD", "PERSONAL"]


def test_collect_fixed_account_budgeting_settings_returns_owner(monkeypatch):
    _feed(monkeypatch, ["personb"])
    overrides = {}

    owners = wizard.collect_fixed_account_budgeting_settings(overrides, "NORWEGIAN")

    assert overrides["budgeting"]["source_account_inference"]["NORWEGIAN"]["fixed_source_account"] == "PERSONB"
    assert owners == ["PERSONB"]


def test_collect_spankki_settings_maps_buffer_answer_to_default_include(monkeypatch):
    _feed(monkeypatch, [""])  # default Y (buffer account)
    overrides = {}
    owners = wizard.collect_spankki_settings(overrides)

    spankki = overrides["budgeting"]["source_account_inference"]["SPANKKI"]
    assert spankki["fixed_source_account"] == "SPANKKI"
    assert spankki["default_include"] == "NO"
    assert owners == []  # SPANKKI is an account type, not a person - no Owner to register

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
    # NORDNET is AOT/OST-capable, so each portfolio also gets an account-type prompt.
    _feed(monkeypatch, ["y", "10000001", "personal", "1", "", "household"])
    overrides = {}
    wizard.collect_portfolio_owners(overrides, "NORDNET")

    assert overrides["investments"]["portfolio_owners"]["NORDNET"] == {
        "10000001": "PERSONAL",
        "default": "HOUSEHOLD",
    }
    assert overrides["investments"]["portfolio_types"]["NORDNET"] == {
        "10000001": "Arvo-osuustili",
    }


def test_collect_portfolio_owners_asks_account_type_for_aot_ost_capable_broker(monkeypatch):
    _feed(monkeypatch, ["n", "personal", "2"])
    overrides = {}
    wizard.collect_portfolio_owners(overrides, "OP")

    assert overrides["investments"]["portfolio_owners"]["OP"] == "PERSONAL"
    assert overrides["investments"]["portfolio_types"]["OP"] == "Osakesäästötili"


def test_collect_portfolio_owners_does_not_ask_account_type_for_evli(monkeypatch):
    # Only "n" (not multiple) and an owner - no account-type prompt for EVLI,
    # so there must be nothing left over to consume.
    _feed(monkeypatch, ["n", "personal"])
    overrides = {}

    wizard.collect_portfolio_owners(overrides, "EVLI")

    assert overrides["investments"]["portfolio_owners"]["EVLI"] == "PERSONAL"
    assert "portfolio_types" not in overrides["investments"]


def test_collect_portfolio_owners_warns_and_can_abort_when_replacing_existing_multi_portfolio(monkeypatch):
    """
    Real bug caught by the user: answering "no" to "more than one portfolio"
    for a broker that already has several real portfolios configured would
    silently replace that whole per-portfolio breakdown with one flat owner.
    Must warn, and let the user back out without changing anything.
    """
    existing = {
        "investments": {
            "portfolio_owners": {
                "NORDNET": {"11111111": "PERSON_A", "22222222": "PERSON_B", "default": "HOUSEHOLD"},
            },
        },
    }
    # Default for "more than one portfolio?" should be Y (pre-filled from existing data) -
    # answer "n" anyway to trigger the warning, then decline the replacement.
    _feed(monkeypatch, ["n", "n"])
    overrides = {}

    wizard.collect_portfolio_owners(overrides, "NORDNET", existing)

    assert "investments" not in overrides


def test_collect_portfolio_owners_replaces_existing_multi_portfolio_when_confirmed(monkeypatch):
    existing = {
        "investments": {
            "portfolio_owners": {
                "NORDNET": {"11111111": "PERSON_A", "default": "HOUSEHOLD"},
            },
        },
    }
    _feed(monkeypatch, ["n", "y", "personb", "1"])
    overrides = {}

    wizard.collect_portfolio_owners(overrides, "NORDNET", existing)

    assert overrides["investments"]["portfolio_owners"]["NORDNET"] == "PERSONB"


def test_collect_investment_broker_settings_evli_asks_fixed_instrument_name(monkeypatch):
    _feed(monkeypatch, ["nokia", "n", "personal"])
    overrides = {}
    wizard.collect_investment_broker_settings(overrides, "EVLI")

    assert overrides["investments"]["broker_defaults"]["EVLI"]["fixed_instrument_name"] == "NOKIA"
    assert overrides["investments"]["portfolio_owners"]["EVLI"] == "PERSONAL"
    assert "portfolio_types" not in overrides["investments"]


def test_collect_investment_broker_settings_auto_fills_fixed_portfolio_type(monkeypatch):
    _feed(monkeypatch, ["n", "household"])
    overrides = {}

    wizard.collect_investment_broker_settings(overrides, "COINMOTION")

    assert overrides["investments"]["portfolio_types"]["COINMOTION"] == "Crypto wallet"


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


def test_write_ownership_rules_with_no_owners_does_not_prompt(tmp_path):
    rules_path = tmp_path / "TransactionRules.xlsx"

    written = wizard.write_ownership_rules(rules_path, [])

    assert written is False
    assert not rules_path.exists()


def test_write_ownership_rules_creates_sheet_on_workbook_with_no_ownership_rules_sheet(tmp_path, monkeypatch):
    """
    A brand-new TransactionRules.xlsx copied from the template has no
    OwnershipRules sheet at all - the wizard must create it fresh, not
    error out.
    """
    rules_path = tmp_path / "TransactionRules.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "TransactionRules"
    ws.append(["RuleID", "Enabled"])
    ws.append(["BR0001", "YES"])
    wb.save(rules_path)

    _feed(monkeypatch, ["y"])
    written = wizard.write_ownership_rules(rules_path, ["HOUSEHOLD"])

    assert written is True
    df = pd.read_excel(rules_path, sheet_name="OwnershipRules", dtype=object)
    assert list(df["RuleID"]) == ["OR0001"]
    assert df.iloc[0]["Pattern"] == "HOUSEHOLD"
    assert df.iloc[0]["SetOwner"] == "HOUSEHOLD"
    assert df.iloc[0]["MatchField"] == "SourceAccount"

    # The pre-existing TransactionRules sheet must survive untouched.
    other = pd.read_excel(rules_path, sheet_name="TransactionRules", dtype=object)
    assert list(other["RuleID"]) == ["BR0001"]


def test_write_ownership_rules_skips_already_covered_owners_and_continues_numbering(tmp_path, monkeypatch):
    rules_path = tmp_path / "TransactionRules.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "OwnershipRules"
    ws.append(wizard.OWNERSHIP_RULES_COLUMNS)
    ws.append(["OR0001", "YES", 100, "Source account PERSONA owns transaction", "SourceAccount", "EXACT", "PERSONA", "NO", "PERSONA", "NO", "YES", "BLANK_ONLY", "Notes"])
    wb.save(rules_path)

    _feed(monkeypatch, ["y"])
    written = wizard.write_ownership_rules(rules_path, ["PERSONA", "PERSONB"])

    assert written is True
    df = pd.read_excel(rules_path, sheet_name="OwnershipRules", dtype=object)
    assert list(df["RuleID"]) == ["OR0001", "OR0002"]
    assert list(df["Pattern"]) == ["PERSONA", "PERSONB"]


def test_write_ownership_rules_does_not_write_when_user_declines(tmp_path, monkeypatch):
    rules_path = tmp_path / "TransactionRules.xlsx"
    _feed(monkeypatch, ["n"])

    written = wizard.write_ownership_rules(rules_path, ["HOUSEHOLD"])

    assert written is False
    assert not rules_path.exists()


def test_cli_has_dry_run_and_keep_temp():
    parser = wizard.build_arg_parser()
    args = parser.parse_args(["--dry-run", "--keep-temp"])

    assert args.dry_run is True
    assert args.keep_temp is True


def test_keep_temp_requires_dry_run(monkeypatch):
    # main() must reject this combination before touching any real path -
    # the ValueError check runs immediately after argument parsing.
    monkeypatch.setattr("sys.argv", ["setup_wizard.py", "--keep-temp"])

    with pytest.raises(ValueError):
        wizard.main()


def test_run_wizard_dry_run_never_touches_real_default_paths(tmp_path, monkeypatch):
    """
    The whole point of --dry-run: run_wizard() must accept explicit sandbox
    paths and never fall back to the real DEFAULT_* module paths while doing
    so - this is what main()'s --dry-run branch relies on.
    """
    sandbox_settings = tmp_path / "settings.yaml"
    sandbox_budgeting_rules = tmp_path / "TransactionRules.xlsx"
    sandbox_investment_rules = tmp_path / "InstrumentMaster.xlsx"

    real_settings_touched = []

    def guard_open(self, *args, **kwargs):
        if self == wizard.DEFAULT_SETTINGS_PATH:
            real_settings_touched.append(self)
        return original_open(self, *args, **kwargs)

    original_open = Path.open
    monkeypatch.setattr(Path, "open", guard_open)

    # scope=budgeting only, no brokers selected, no cash tracking, no unsupported banks
    _feed(monkeypatch, ["1", "", "n", ""])

    wizard.run_wizard(
        settings_path=sandbox_settings,
        budgeting_template=tmp_path / "missing_template.xlsx",
        budgeting_rules=sandbox_budgeting_rules,
        investment_template=tmp_path / "missing_template2.xlsx",
        investment_rules=sandbox_investment_rules,
        dry_run=True,
    )

    assert real_settings_touched == []
    assert not sandbox_settings.exists()
