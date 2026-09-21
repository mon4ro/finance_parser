import pandas as pd

from finance_parser.budgeting.transaction_normaliser import (
    Rule,
    apply_rules,
    count_blank_normalized_receivers,
)


def make_receiver_rule() -> Rule:
    return Rule(
        rule_id="T001",
        enabled=True,
        priority=1,
        rule_name="Test receiver normalisation",
        rule_group="Test",
        field_to_search="RawReceiver",
        match_type="CONTAINS",
        pattern="LIDL",
        case_sensitive=False,
        set_values={"NormalizedReceiver": "LIDL"},
        condition_field="",
        condition_match_type="",
        condition_pattern="",
        stop_if_matched=True,
        overwrite_mode="BLANK_ONLY",
    )


def test_apply_rules_skips_split_rows():
    unified = pd.DataFrame(
        [
            {
                "UnifiedID": "parent-1",
                "TransactionType": "Bank",
                "RawReceiver": "LIDL ESPOO",
                "NormalizedReceiver": "",
                "Include": "YES",
            },
            {
                "UnifiedID": "parent-1-S01",
                "TransactionType": "Split",
                "RawReceiver": "LIDL ESPOO",
                "NormalizedReceiver": "",
                "Include": "YES",
            },
        ]
    )

    updated, stats = apply_rules(unified, [make_receiver_rule()])

    assert updated.loc[0, "NormalizedReceiver"] == "LIDL"
    assert updated.loc[1, "NormalizedReceiver"] == ""
    assert stats["rows_checked"] == 2
    assert stats["rows_skipped_split"] == 1
    assert stats["fields_updated"] == 1


def test_blank_normalized_receiver_counter_excludes_split_rows_by_default():
    unified = pd.DataFrame(
        [
            {"TransactionType": "Bank", "NormalizedReceiver": ""},
            {"TransactionType": "Split", "NormalizedReceiver": ""},
            {"TransactionType": "Split", "NormalizedReceiver": "MANUAL"},
        ]
    )

    assert count_blank_normalized_receivers(unified) == 1
    assert count_blank_normalized_receivers(unified, include_split_rows=True) == 2
