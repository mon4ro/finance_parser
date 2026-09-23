from pathlib import Path

from openpyxl import Workbook, load_workbook

from finance_parser.budgeting.transaction_normaliser import (
    find_rule_by_id,
    load_rules,
    run_rule_command,
)


RULES_HEADER = [
    "RuleID", "Enabled", "Priority", "RuleName", "RuleGroup",
    "FieldToSearch", "MatchType", "Pattern", "CaseSensitive",
    "SetNormalizedReceiver", "SetInclude", "SetOwner",
    "SetSupercategory", "SetCategory", "SetSubcategory", "SetComments",
    "StopIfMatched", "OverwriteMode", "Notes",
    "SourceBank", "SourceAccount", "ConditionField", "ConditionMatchType", "ConditionPattern",
]


def make_rules_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "TransactionRules"
    ws.append(RULES_HEADER)
    # Higher priority (lower number) rule also targeting NormalizedReceiver
    # for a different, more specific pattern - proves the simulation still
    # respects real priority ordering, not just "the one rule we named".
    ws.append([
        "BR0050", "YES", 50, "Specific override", "TEST",
        "RawReceiver", "EQUALS", "ABC-LATAUS HOK-ELANTO SPECIAL", "NO",
        "SOMETHING_ELSE", None, None,
        None, None, None, None,
        "YES", "BLANK_ONLY", None,
        None, None, None, None, None,
    ])
    ws.append([
        "BR0290", "YES", 100, "S-BONUS receiver normalisation", "RECEIVER",
        "RawReceiver", "EQUALS", "HOK-ELANTO", "NO",
        "S-BONUS", None, None,
        None, None, None, None,
        "YES", "BLANK_ONLY", None,
        None, None, "Description", "REGEX", "^(BONUS|MAKSUTAPAETU)$",
    ])
    ws.append([
        "BR0999", "NO", 100, "Disabled rule", "TEST",
        "RawReceiver", "CONTAINS", "WHATEVER", "NO",
        "SHOULD_NOT_APPEAR", None, None,
        None, None, None, None,
        "YES", "BLANK_ONLY", None,
        None, None, None, None, None,
    ])
    wb.save(path)


def make_transactions_workbook(path: Path, rows: list[list[object]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "UnifiedTransactions"
    ws.append([
        "UnifiedID", "RawID", "SourceBank", "SourceAccount", "TransactionType", "Include", "Amount",
        "RawReceiver", "NormalizedReceiver", "Description", "Message",
        "Supercategory", "Category", "Subcategory", "Owner", "Review/Notes",
    ])
    for row in rows:
        ws.append(row)
    wb.save(path)


def base_row(unified_id, raw_receiver, description="BONUS", normalized_receiver=""):
    return [unified_id, f"R{unified_id}", "SPANKKI", "SPANKKI", "Bank", "YES", -5, raw_receiver, normalized_receiver, description, "", "", "", "", "", ""]


def test_run_rule_fills_a_blank_field(tmp_path):
    workbook = tmp_path / "ParsedTransactions.xlsx"
    rules = tmp_path / "TransactionRules.xlsx"
    make_rules_workbook(rules)
    make_transactions_workbook(workbook, [base_row("U1", "HOK-ELANTO")])

    run_rule_command(workbook, rules, "BR0290", apply_all=True, apply_rows=None, dry_run=False)

    wb = load_workbook(workbook)
    assert wb["UnifiedTransactions"]["I2"].value == "S-BONUS"
    wb.close()


def test_run_rule_corrects_a_stale_non_blank_field(tmp_path):
    """
    The actual new capability: this row is genuinely in BR0290's scope
    (RawReceiver=HOK-ELANTO, Description=MAKSUTAPAETU matches its
    condition) but NormalizedReceiver already holds a stale wrong value -
    a normal rerun would never touch this (BLANK_ONLY protects non-blank
    cells). --run-rule must correct it to S-BONUS.
    """
    workbook = tmp_path / "ParsedTransactions.xlsx"
    rules = tmp_path / "TransactionRules.xlsx"
    make_rules_workbook(rules)
    make_transactions_workbook(workbook, [base_row("U1", "HOK-ELANTO", description="MAKSUTAPAETU", normalized_receiver="WRONG_VALUE")])

    run_rule_command(workbook, rules, "BR0290", apply_all=True, apply_rows=None, dry_run=False)

    wb = load_workbook(workbook)
    assert wb["UnifiedTransactions"]["I2"].value == "S-BONUS"
    wb.close()


def test_run_rule_leaves_out_of_condition_row_untouched_rather_than_clearing(tmp_path):
    """
    This mirrors the real ABC-lataus incident: RawReceiver=HOK-ELANTO
    matches BR0290's own MatchField/Pattern, but Description=KORTTIOSTO
    fails its ConditionField/ConditionPattern (BONUS/MAKSUTAPAETU only) -
    so the row is out of scope. The recomputed value would be blank (no
    rule matches), and --run-rule deliberately never proposes clearing a
    field to blank - so a stale wrong value here is left as-is, not wiped.
    """
    workbook = tmp_path / "ParsedTransactions.xlsx"
    rules = tmp_path / "TransactionRules.xlsx"
    make_rules_workbook(rules)
    make_transactions_workbook(workbook, [base_row("U1", "HOK-ELANTO", description="KORTTIOSTO", normalized_receiver="S-BONUS")])

    run_rule_command(workbook, rules, "BR0290", apply_all=True, apply_rows=None, dry_run=False)

    wb = load_workbook(workbook)
    assert wb["UnifiedTransactions"]["I2"].value == "S-BONUS"
    wb.close()


def test_run_rule_dry_run_does_not_write(tmp_path):
    workbook = tmp_path / "ParsedTransactions.xlsx"
    rules = tmp_path / "TransactionRules.xlsx"
    make_rules_workbook(rules)
    make_transactions_workbook(workbook, [base_row("U1", "HOK-ELANTO")])

    run_rule_command(workbook, rules, "BR0290", apply_all=False, apply_rows=None, dry_run=True)

    wb = load_workbook(workbook)
    assert wb["UnifiedTransactions"]["I2"].value in (None, "")
    wb.close()


def test_run_rule_apply_rows_writes_only_the_named_subset(tmp_path):
    workbook = tmp_path / "ParsedTransactions.xlsx"
    rules = tmp_path / "TransactionRules.xlsx"
    make_rules_workbook(rules)
    make_transactions_workbook(workbook, [
        base_row("U1", "HOK-ELANTO"),
        base_row("U2", "HOK-ELANTO"),
    ])

    run_rule_command(workbook, rules, "BR0290", apply_all=False, apply_rows=["U1"], dry_run=False)

    wb = load_workbook(workbook)
    ws = wb["UnifiedTransactions"]
    assert ws["I2"].value == "S-BONUS"  # U1
    assert ws["I3"].value in (None, "")  # U2 left alone
    wb.close()


def test_find_rule_by_id_errors_clearly_for_a_disabled_rule(tmp_path):
    rules = tmp_path / "TransactionRules.xlsx"
    make_rules_workbook(rules)
    all_rules = load_rules(rules)
    try:
        find_rule_by_id("BR0999", all_rules, path=rules)
        assert False, "expected ValueError"
    except ValueError as exc:
        assert "disabled" in str(exc).lower()


def test_find_rule_by_id_errors_clearly_for_a_categoriser_prefixed_rule_id(tmp_path):
    rules = tmp_path / "TransactionRules.xlsx"
    make_rules_workbook(rules)
    all_rules = load_rules(rules)
    try:
        find_rule_by_id("CR0267", all_rules, path=rules)
        assert False, "expected ValueError"
    except ValueError as exc:
        assert "categoriser" in str(exc).lower()
