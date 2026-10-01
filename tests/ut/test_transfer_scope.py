from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from finance_parser.common import RAW_COLUMNS, UNIFIED_COLUMNS, raw_to_unified_rows
from finance_parser.budgeting.transaction_normaliser import Rule, apply_rules, load_rules


def test_transfer_scope_is_appended_not_inserted():
    """
    Real constraint: Master Budget's formulas reference UnifiedTransactions
    columns by absolute letter ($C:$C, $S:$S, etc.) - inserting a new column
    anywhere but the very end would shift every column after it and break
    every live formula.
    """
    assert UNIFIED_COLUMNS[-1] == "TransferScope"


def test_raw_to_unified_rows_defaults_transfer_scope_blank():
    row = {col: "" for col in RAW_COLUMNS}
    row.update({
        "RawID": "R1",
        "SourceAccount": "PERSON_A",
        "SourceBank": "OP",
        "ValueDate": "2026-06-01",
        "Amount": -10.0,
        "RawReceiver": "Some Shop",
    })
    df = pd.DataFrame([row], columns=RAW_COLUMNS)

    unified = raw_to_unified_rows(df)

    assert unified.iloc[0]["TransferScope"] == ""


def test_apply_rules_sets_transfer_scope_from_set_values():
    """
    Real motivating case: a transfer that funds a shared/buffer account
    (or, per BR0281, an internal joint-to-personal reallocation) should be
    tagged INTERNAL independently of Include, so "show me every internal
    transfer" is answerable without also pulling in every other reason a
    row might be Include=NO.
    """
    df = pd.DataFrame([{
        "RawReceiver": "SMITH JOHN TAI JANE",
        "Message": "",
        "NormalizedReceiver": "",
        "Include": "YES",
        "Owner": "",
        "Supercategory": "",
        "Category": "",
        "Subcategory": "",
        "Review/Notes": "",
        "TransferScope": "",
    }])
    rule = Rule(
        rule_id="T1",
        enabled=True,
        priority=20,
        rule_name="test internal transfer",
        rule_group="INTERNAL",
        field_to_search="RawReceiver",
        match_type="CONTAINS",
        pattern="SMITH",
        case_sensitive=False,
        set_values={"TransferScope": "INTERNAL", "Include": "NO"},
        condition_field="",
        condition_match_type="",
        condition_pattern="",
        stop_if_matched=True,
        overwrite_mode="FORCE",
    )

    updated, stats = apply_rules(df, [rule])

    assert updated.iloc[0]["TransferScope"] == "INTERNAL"
    assert updated.iloc[0]["Include"] == "NO"


def test_load_rules_reads_set_transfer_scope_column(tmp_path):
    path = tmp_path / "TransactionRules.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "TransactionRules"
    ws.append([
        "RuleID", "Enabled", "Priority", "RuleName", "RuleGroup",
        "FieldToSearch", "MatchType", "Pattern", "CaseSensitive",
        "SetNormalizedReceiver", "SetInclude", "SetTransferScope",
        "StopIfMatched", "OverwriteMode",
    ])
    ws.append([
        "BR0001", "YES", 20, "test", "INTERNAL",
        "RawReceiver", "CONTAINS", "SMITH", "NO",
        None, "NO", "INTERNAL",
        "YES", "FORCE",
    ])
    wb.save(path)

    rules = load_rules(path)

    assert len(rules) == 1
    assert rules[0].set_values["TransferScope"] == "INTERNAL"
    assert rules[0].set_values["Include"] == "NO"
