from finance_parser.budgeting import transaction_categoriser as categoriser
from finance_parser.budgeting.transaction_categoriser import (
    RECATEGORISE_FIELDS,
    clear_recategorise_fields,
    merge_changes,
)


def test_recategorise_cli_switches_exist():
    parser = categoriser.build_arg_parser()

    args = parser.parse_args(["--recategorise", "--all"])
    assert args.recategorise is True
    assert args.all is True
    assert args.only_automatic is False

    args = parser.parse_args(["--recategorise", "--only-automatic"])
    assert args.recategorise is True
    assert args.all is False
    assert args.only_automatic is True


def test_recategorise_modes_are_mutually_exclusive():
    parser = categoriser.build_arg_parser()

    try:
        parser.parse_args(["--recategorise", "--all", "--only-automatic"])
    except SystemExit as exc:
        assert exc.code != 0
    else:
        raise AssertionError("--all and --only-automatic should be mutually exclusive")


def test_recategorise_confirmation_requires_two_yes_answers(monkeypatch):
    answers = iter(["yes", "no"])
    monkeypatch.setattr("builtins.input", lambda _: next(answers))

    assert categoriser.confirm_recategorise("only_automatic") is False

    answers = iter(["yes", "yes"])
    monkeypatch.setattr("builtins.input", lambda _: next(answers))

    assert categoriser.confirm_recategorise("only_automatic") is True


def test_recategorise_fields_do_not_include_comments_or_include():
    assert "Supercategory" in RECATEGORISE_FIELDS
    assert "Category" in RECATEGORISE_FIELDS
    assert "Subcategory" in RECATEGORISE_FIELDS
    assert "Owner" in RECATEGORISE_FIELDS

    assert "Comments" not in RECATEGORISE_FIELDS
    assert "Comment" not in RECATEGORISE_FIELDS
    assert "Include" not in RECATEGORISE_FIELDS
    assert "NormalizedReceiver" not in RECATEGORISE_FIELDS


def test_clear_recategorise_fields_clears_only_existing_category_fields():
    row = {
        "Supercategory": "OLD",
        "Category": "OLD CAT",
        "Subcategory": "",
        "Owner": "PERSON_1",
        "Comments": "manual note",
        "Include": "YES",
    }
    headers = {key: idx for idx, key in enumerate(row.keys(), start=1)}

    changes = clear_recategorise_fields(row, headers)

    assert row["Supercategory"] == ""
    assert row["Category"] == ""
    assert row["Owner"] == ""
    assert row["Comments"] == "manual note"
    assert row["Include"] == "YES"

    assert changes["Supercategory"] == ("OLD", "")
    assert changes["Category"] == ("OLD CAT", "")
    assert changes["Owner"] == ("PERSON_1", "")


def test_merge_changes_preserves_original_old_value():
    base = {"Category": ("OLD", "")}
    new = {"Category": ("", "Restaurants")}

    merge_changes(base, new)

    assert base["Category"] == ("OLD", "Restaurants")
