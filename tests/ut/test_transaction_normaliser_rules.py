import pandas as pd

from finance_parser.budgeting.transaction_normaliser import Rule, apply_rules, rule_matches


def make_rule(**kwargs):
    defaults = dict(
        rule_id="T1",
        enabled=True,
        priority=100,
        rule_name="test",
        rule_group="RECEIVER",
        field_to_search="RawReceiver",
        match_type="CONTAINS",
        pattern="prisma",
        case_sensitive=False,
        set_values={"NormalizedReceiver": "PRISMA"},
        condition_field="",
        condition_match_type="",
        condition_pattern="",
        stop_if_matched=True,
        overwrite_mode="BLANK_ONLY",
    )
    defaults.update(kwargs)
    return Rule(**defaults)


def test_contains_is_case_insensitive_by_default():
    rule = make_rule(pattern="prisma", case_sensitive=False)
    assert rule_matches(rule, "ABC Prisma Espoo")


def test_all_contains_requires_all_parts():
    rule = make_rule(match_type="ALL_CONTAINS", pattern="Alpha|Beta")
    assert rule_matches(rule, "Alpha and Beta")
    assert not rule_matches(rule, "Alpha only")


def test_dynamic_regex_group_from_message_with_condition():
    df = pd.DataFrame([{
        "RawReceiver": "Trustly group AB",
        "Message": "maksu saaja: Example Merchant, viite 123",
        "NormalizedReceiver": "",
        "Include": "",
        "Owner": "",
        "Supercategory": "",
        "Category": "",
        "Subcategory": "",
        "Comments": "",
    }])

    rule = make_rule(
        field_to_search="Message",
        match_type="REGEX",
        pattern=r"saaja:\s*([^,;\n\r]+)",
        set_values={"NormalizedReceiver": "{REGEX_GROUP_1_UPPER}"},
        condition_field="RawReceiver",
        condition_match_type="CONTAINS",
        condition_pattern="Trustly group AB",
    )

    out, stats = apply_rules(df, [rule])
    assert out.loc[0, "NormalizedReceiver"] == "EXAMPLE MERCHANT"
    assert stats["fields_updated"] == 1
