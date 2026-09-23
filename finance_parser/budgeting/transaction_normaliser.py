from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from finance_parser.common import clean_for_excel
from finance_parser.utilities.fresh_workbook_writer import (
    append_changelog_row,
    read_workbook_values_only,
    records_to_sheet_values,
    replace_with_fresh_workbook,
    sheet_values_to_records,
)


PROJECT_ROOT = Path(__file__).resolve().parents[2]

BUDGETING_OUTPUT_DIR = PROJECT_ROOT / "output" / "budgeting"
BUDGETING_RULES_DIR = PROJECT_ROOT / "rules" / "budgeting"

DEFAULT_WORKBOOK = BUDGETING_OUTPUT_DIR / "ParsedTransactions.xlsx"
DEFAULT_RULES = BUDGETING_RULES_DIR / "TransactionRules.xlsx"

UNIFIED_SHEET = "UnifiedTransactions"
RULES_SHEET = "TransactionRules"
SCRIPT_NAME = "transaction_normaliser.py"

TARGET_FIELD_BY_SET_COLUMN = {
    "SetNormalizedReceiver": "NormalizedReceiver",
    "SetInclude": "Include",
    "SetOwner": "Owner",
    "SetSupercategory": "Supercategory",
    "SetCategory": "Category",
    "SetSubcategory": "Subcategory",
    "SetComments": "Review/Notes",
    "SetReviewNotes": "Review/Notes",

}

VALID_MATCH_TYPES = {
    "CONTAINS",
    "ALL_CONTAINS",
    "EQUALS",
    "STARTS_WITH",
    "ENDS_WITH",
    "REGEX",
}


def text(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def is_yes(value: object) -> bool:
    return text(value).upper() in {"YES", "Y", "TRUE", "1"}


def is_blank(value: object) -> bool:
    return text(value) == ""


def is_split_row(row: pd.Series) -> bool:
    return text(row.get("TransactionType", "")).upper() == "SPLIT"


def count_blank_normalized_receivers(df: pd.DataFrame, *, include_split_rows: bool = False) -> int:
    """Count blank NormalizedReceiver cells after rules are applied.

    Split rows are manual child rows. They are excluded by default because the
    normaliser intentionally does not touch them.
    """
    if "NormalizedReceiver" not in df.columns:
        return 0

    if include_split_rows or "TransactionType" not in df.columns:
        rows = df
    else:
        rows = df[~df.apply(is_split_row, axis=1)]

    return int(rows["NormalizedReceiver"].apply(is_blank).sum())


@dataclass
class Rule:
    rule_id: str
    enabled: bool
    priority: int
    rule_name: str
    rule_group: str
    field_to_search: str
    match_type: str
    pattern: str
    case_sensitive: bool
    set_values: dict[str, str]
    condition_field: str
    condition_match_type: str
    condition_pattern: str
    stop_if_matched: bool
    overwrite_mode: str


def parse_priority(value: object) -> int:
    try:
        return int(float(text(value)))
    except Exception:
        return 9999


def load_rules(path: Path) -> list[Rule]:
    if not path.exists():
        raise FileNotFoundError(f"Transaction rules workbook not found: {path}")

    df = pd.read_excel(path, sheet_name=RULES_SHEET, dtype=object, engine="openpyxl")
    df.columns = [text(c) for c in df.columns]

    required = [
        "RuleID",
        "Enabled",
        "Priority",
        "RuleName",
        "RuleGroup",
        "FieldToSearch",
        "MatchType",
        "Pattern",
        "CaseSensitive",
        "StopIfMatched",
        "OverwriteMode",
    ]

    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"{RULES_SHEET} is missing required columns: {', '.join(missing)}")

    rules: list[Rule] = []

    for _, row in df.iterrows():
        enabled = is_yes(row.get("Enabled", "YES"))
        rule_id = text(row.get("RuleID", ""))
        pattern = text(row.get("Pattern", ""))
        field_to_search = text(row.get("FieldToSearch", ""))
        match_type = text(row.get("MatchType", "CONTAINS")).upper()
        overwrite_mode = text(row.get("OverwriteMode", "BLANK_ONLY")).upper()

        if not rule_id and not pattern:
            continue

        if not enabled:
            continue

        if not field_to_search:
            continue

        if not pattern:
            continue

        if match_type not in VALID_MATCH_TYPES:
            raise ValueError(f"Rule {rule_id} has unsupported MatchType: {match_type}")

        condition_field = text(row.get("ConditionField", ""))
        condition_match_type = text(row.get("ConditionMatchType", "CONTAINS")).upper()
        condition_pattern = text(row.get("ConditionPattern", ""))

        if condition_field and condition_pattern and condition_match_type not in VALID_MATCH_TYPES:
            raise ValueError(f"Rule {rule_id} has unsupported ConditionMatchType: {condition_match_type}")

        set_values = {}
        for set_column, target_field in TARGET_FIELD_BY_SET_COLUMN.items():
            if set_column in df.columns:
                value = text(row.get(set_column, ""))
                if value:
                    set_values[target_field] = value

        if not set_values:
            continue

        rules.append(
            Rule(
                rule_id=rule_id,
                enabled=enabled,
                priority=parse_priority(row.get("Priority", 9999)),
                rule_name=text(row.get("RuleName", "")),
                rule_group=text(row.get("RuleGroup", "")),
                field_to_search=field_to_search,
                match_type=match_type,
                pattern=pattern,
                case_sensitive=is_yes(row.get("CaseSensitive", "NO")),
                set_values=set_values,
                condition_field=condition_field,
                condition_match_type=condition_match_type,
                condition_pattern=condition_pattern,
                stop_if_matched=is_yes(row.get("StopIfMatched", "YES")),
                overwrite_mode=overwrite_mode or "BLANK_ONLY",
            )
        )

    rules.sort(key=lambda r: (r.priority, r.rule_id))
    return rules


def rule_matches(rule: Rule, value: object) -> bool:
    haystack = text(value)
    needle = rule.pattern

    if not rule.case_sensitive:
        haystack_cmp = haystack.upper()
        needle_cmp = needle.upper()
    else:
        haystack_cmp = haystack
        needle_cmp = needle

    if rule.match_type == "CONTAINS":
        return needle_cmp in haystack_cmp
    if rule.match_type == "ALL_CONTAINS":
        parts = [part.strip() for part in needle_cmp.split("|") if part.strip()]
        return bool(parts) and all(part in haystack_cmp for part in parts)
    if rule.match_type == "EQUALS":
        return haystack_cmp == needle_cmp
    if rule.match_type == "STARTS_WITH":
        return haystack_cmp.startswith(needle_cmp)
    if rule.match_type == "ENDS_WITH":
        return haystack_cmp.endswith(needle_cmp)
    if rule.match_type == "REGEX":
        flags = 0 if rule.case_sensitive else re.IGNORECASE
        return re.search(needle, haystack, flags=flags) is not None

    return False


def condition_matches(rule: Rule, row: pd.Series) -> bool:
    if not rule.condition_field or not rule.condition_pattern:
        return True

    condition_rule = Rule(
        rule_id=rule.rule_id + "_condition",
        enabled=True,
        priority=rule.priority,
        rule_name=rule.rule_name + " condition",
        rule_group=rule.rule_group,
        field_to_search=rule.condition_field,
        match_type=rule.condition_match_type or "CONTAINS",
        pattern=rule.condition_pattern,
        case_sensitive=rule.case_sensitive,
        set_values={},
        condition_field="",
        condition_match_type="",
        condition_pattern="",
        stop_if_matched=True,
        overwrite_mode="BLANK_ONLY",
    )
    return rule_matches(condition_rule, row.get(rule.condition_field, ""))


def extract_regex_group(pattern: str, value: object, case_sensitive: bool, group: int = 1) -> str:
    flags = 0 if case_sensitive else re.IGNORECASE
    match = re.search(pattern, text(value), flags=flags)
    if not match:
        return ""
    try:
        return text(match.group(group)).upper()
    except IndexError:
        return ""


def resolve_set_value(template: str, rule: Rule, row: pd.Series) -> str:
    """
    Resolve dynamic rule outputs.

    Supported templates:
    - {RAW_RECEIVER_UPPER}
    - {MATCHED_FIELD_UPPER}
    - {AFTER_PATTERN_UPPER}
    - {REGEX_GROUP_1_UPPER}
    - {SOURCE_ACCOUNT_UPPER}
    """
    value = text(template)

    matched_field_value = text(row.get(rule.field_to_search, ""))
    raw_receiver_value = text(row.get("RawReceiver", ""))
    source_account_value = text(row.get("SourceAccount", ""))

    if value == "{RAW_RECEIVER_UPPER}":
        return raw_receiver_value.upper()

    if value == "{SOURCE_ACCOUNT_UPPER}":
        return source_account_value.upper()

    if value == "{MATCHED_FIELD_UPPER}":
        return matched_field_value.upper()

    if value == "{AFTER_PATTERN_UPPER}":
        if rule.match_type == "STARTS_WITH":
            source = matched_field_value
            pattern = rule.pattern
            if not rule.case_sensitive:
                if source.upper().startswith(pattern.upper()):
                    return source[len(pattern):].strip().upper()
            elif source.startswith(pattern):
                return source[len(pattern):].strip().upper()
        return ""

    if value == "{REGEX_GROUP_1_UPPER}":
        return extract_regex_group(rule.pattern, matched_field_value, rule.case_sensitive, group=1)

    return value


def can_write_field(current_value: object, overwrite_mode: str) -> bool:
    mode = overwrite_mode.upper()

    if mode in {"FORCE", "ALWAYS", "OVERWRITE"}:
        return True

    # RULE_ONLY is reserved for the later source-tracking model. Until we add
    # field source/applied-value columns, treat it conservatively as BLANK_ONLY.
    if mode in {"BLANK_ONLY", "RULE_ONLY", ""}:
        return is_blank(current_value)

    raise ValueError(f"Unsupported OverwriteMode: {overwrite_mode}")


def apply_rules(unified: pd.DataFrame, rules: list[Rule]) -> tuple[pd.DataFrame, dict[str, int]]:
    out = unified.copy()

    for col in TARGET_FIELD_BY_SET_COLUMN.values():
        if col not in out.columns:
            out[col] = ""

    stats = {
        "rows_checked": len(out),
        "rules_loaded": len(rules),
        "rows_skipped_split": 0,
        "rule_matches": 0,
        "fields_updated": 0,
        "fields_skipped_existing_value": 0,
        "empty_dynamic_outputs": 0,
    }

    for idx, row in out.iterrows():
        # Split rows are manual child transactions. The parent bank row carries
        # the original receiver/source metadata; the child rows are manually
        # categorised by the user and must not be normalised automatically.
        if is_split_row(row):
            stats["rows_skipped_split"] += 1
            continue

        for rule in rules:
            if rule.field_to_search not in out.columns:
                continue

            if not rule_matches(rule, row.get(rule.field_to_search, "")):
                continue

            if not condition_matches(rule, row):
                continue

            stats["rule_matches"] += 1
            any_update_for_rule = False

            for target_field, template_value in rule.set_values.items():
                new_value = resolve_set_value(template_value, rule, row)

                # Empty dynamic outputs should not overwrite or stop useful later rules.
                if new_value == "":
                    stats["empty_dynamic_outputs"] += 1
                    continue

                current_value = out.at[idx, target_field]

                if can_write_field(current_value, rule.overwrite_mode):
                    # Avoid counting no-op FORCE assignments as updates.
                    if text(current_value) != new_value:
                        out.at[idx, target_field] = new_value
                        stats["fields_updated"] += 1
                        any_update_for_rule = True
                else:
                    stats["fields_skipped_existing_value"] += 1

            if rule.stop_if_matched and any_update_for_rule:
                break

    return out, stats


def find_rule_by_id(rule_id: str, rules: list[Rule], *, path: Path) -> Rule:
    """
    rules is already Enabled-only (see load_rules) - if not found there,
    re-check the raw sheet (regardless of Enabled) so a disabled rule gets
    a clear, specific error instead of a bare "not found".
    """
    for rule in rules:
        if rule.rule_id == rule_id:
            return rule

    if rule_id.upper().startswith(("CR", "OR")):
        raise ValueError(
            f"Rule {rule_id!r} looks like a CategoryRules/OwnershipRules rule - "
            "use `transaction_categoriser.py --run-rule` instead, not the normaliser."
        )

    df = pd.read_excel(path, sheet_name=RULES_SHEET, dtype=object, engine="openpyxl")
    df.columns = [text(c) for c in df.columns]
    if "RuleID" in df.columns and (df["RuleID"].astype(str).str.strip() == rule_id).any():
        raise ValueError(f"Rule {rule_id!r} exists in {RULES_SHEET} but is disabled (Enabled != YES) - enable it first.")

    raise ValueError(f"Rule {rule_id!r} not found in {RULES_SHEET}.")


def row_matches_rule(rule: Rule, row: pd.Series) -> bool:
    if rule.field_to_search not in row.index:
        return False
    if not rule_matches(rule, row.get(rule.field_to_search, "")):
        return False
    return condition_matches(rule, row)


def run_single_rule(
    unified: pd.DataFrame,
    all_rules: list[Rule],
    rule_id: str,
    *,
    path: Path,
) -> tuple[pd.DataFrame, list[dict[str, object]]]:
    """
    Narrowly re-apply ONE rule's effect to the rows it matches - including
    correcting a stale non-blank value, which a normal normaliser run
    cannot do (every rule here defaults to BLANK_ONLY). Recomputes by
    blanking the target rule's own Set* fields on just the matching rows,
    then running the *existing*, unmodified apply_rules() (full
    priority-ordered chain across all rules, not just this one) so a
    higher-priority rule targeting the same field still correctly wins.
    Only the rows whose recomputed value actually differs from what's
    there now are proposed as changes.

    Note: unlike the categoriser, this engine has no Review/Notes-style
    manual-override gate (see apply_rules() - it only skips split rows) -
    a manually-corrected NormalizedReceiver/Include on a row matching the
    target rule is not otherwise protected here, consistent with how a
    normal normaliser run already behaves for split rows only.
    """
    target_rule = find_rule_by_id(rule_id, all_rules, path=path)
    target_fields = set(target_rule.set_values.keys())
    if not target_fields:
        raise ValueError(f"Rule {rule_id!r} does not set any field - nothing to recompute.")

    match_mask = unified.apply(lambda row: row_matches_rule(target_rule, row), axis=1)
    matched_idx = unified.index[match_mask]

    simulated = unified.copy()
    for field in target_fields:
        if field in simulated.columns:
            simulated.loc[matched_idx, field] = ""

    updated, _ = apply_rules(simulated, all_rules)

    proposed: list[dict[str, object]] = []
    for idx in matched_idx:
        row_changes: dict[str, tuple[object, object]] = {}
        for field in target_fields:
            if field not in unified.columns:
                continue
            old_value = unified.at[idx, field]
            new_value = updated.at[idx, field]
            if is_blank(new_value):
                continue
            if text(old_value) != text(new_value):
                row_changes[field] = (text(old_value), text(new_value))

        if row_changes:
            proposed.append({
                "index": idx,
                "unified_id": text(unified.at[idx, "UnifiedID"]) if "UnifiedID" in unified.columns else "",
                "raw_receiver": text(unified.at[idx, "RawReceiver"]) if "RawReceiver" in unified.columns else "",
                "changes": row_changes,
            })

    return updated, proposed


def print_run_rule_report(rule_id: str, target_fields: set[str], proposed: list[dict[str, object]], rows_checked: int) -> None:
    print(f"Rule {rule_id!r} targets field(s): {', '.join(sorted(target_fields))}")
    print(f"Rows checked: {rows_checked}")
    if not proposed:
        print("No changes proposed.")
        return
    print(f"Changed rows: {len(proposed)}")
    for item in proposed:
        fields = ", ".join(f"{f}: {old!r} -> {new!r}" for f, (old, new) in item["changes"].items())
        print(f"  UnifiedID={item['unified_id']!r} RawReceiver={item['raw_receiver']!r} {fields}")


def prompt_for_row_selection(proposed: list[dict[str, object]]) -> set[str]:
    choice = input("Apply [a]ll / [s]elect rows / [n]one: ").strip().lower()
    if choice == "a":
        return {item["unified_id"] for item in proposed}
    if choice == "s":
        raw = input("Comma-separated UnifiedIDs to apply: ").strip()
        return {v.strip() for v in raw.split(",") if v.strip()}
    return set()


def run_rule_command(
    workbook_path: Path,
    rules_path: Path,
    rule_id: str,
    *,
    apply_all: bool,
    apply_rows: list[str] | None,
    dry_run: bool,
) -> None:
    rules = load_rules(rules_path)
    target_rule = find_rule_by_id(rule_id, rules, path=rules_path)
    target_fields = set(target_rule.set_values.keys())

    print(f"Read previous workbook values only: {workbook_path}")
    sheets, unified = load_unified_from_workbook_values(workbook_path)
    print(f"Loaded sheets: {len(sheets)}")

    updated, proposed = run_single_rule(unified, rules, rule_id, path=rules_path)
    print_run_rule_report(rule_id, target_fields, proposed, rows_checked=len(unified))

    if not proposed:
        return

    if dry_run:
        print()
        print("Dry run only: workbook was not modified.")
        return

    if apply_rows is not None:
        rows_to_apply = set(apply_rows)
    elif apply_all:
        rows_to_apply = {item["unified_id"] for item in proposed}
    else:
        rows_to_apply = prompt_for_row_selection(proposed)

    if not rows_to_apply:
        print("Nothing selected to apply - aborted.")
        return

    selected = [item for item in proposed if item["unified_id"] in rows_to_apply]
    unmatched = rows_to_apply - {item["unified_id"] for item in proposed}
    if unmatched:
        print(f"WARNING: these selected UnifiedIDs did not match any proposed change and were ignored: {sorted(unmatched)}")

    if not selected:
        print("Nothing selected to apply - aborted.")
        return

    final = unified.copy()
    for item in selected:
        for field, (_, new_value) in item["changes"].items():
            final.at[item["index"], field] = new_value

    headers = [str(col) for col in final.columns]
    records = dataframe_to_records(final)
    sheets[UNIFIED_SHEET] = records_to_sheet_values(headers, records)

    append_changelog_row(
        sheets,
        script=SCRIPT_NAME,
        action=f"--run-rule {rule_id}",
        sheet=UNIFIED_SHEET,
        rows_updated=len(selected),
        status="Completed",
        details=(
            f"Narrow re-application of rule {rule_id} ({RULES_SHEET}), targeting "
            f"field(s) {', '.join(sorted(target_fields))}. {len(proposed)} row(s) "
            f"proposed, {len(selected)} applied."
        ),
        backup_file=None,
    )

    backup_path, _, writer_stats = replace_with_fresh_workbook(
        workbook_path,
        sheets,
        backup_label=f"before_run_rule_{rule_id}",
        basic_formatting=True,
        excel_tables=False,
    )
    print(f"Backed up old workbook: {backup_path}")
    print(f"Sheets written: {writer_stats['sheets_written']}")
    print(f"Applied {len(selected)} row(s).")


def load_unified_from_workbook_values(workbook_path: Path) -> tuple[dict[str, list[list[object]]], pd.DataFrame]:
    """Read workbook values only and return all sheets plus UnifiedTransactions."""
    sheets = read_workbook_values_only(workbook_path)

    if UNIFIED_SHEET not in sheets:
        raise ValueError(f"Workbook does not contain required sheet: {UNIFIED_SHEET}")

    headers, records = sheet_values_to_records(sheets[UNIFIED_SHEET])
    unified = pd.DataFrame(records, columns=headers)
    return sheets, unified


def dataframe_to_records(df: pd.DataFrame) -> list[dict[str, object]]:
    """Convert a DataFrame into clean Excel-safe row dictionaries."""
    records: list[dict[str, object]] = []

    for _, row in df.iterrows():
        record: dict[str, object] = {}
        for col in df.columns:
            value = row.get(col, "")
            if pd.isna(value):
                value = ""
            record[str(col)] = clean_for_excel(value)
        records.append(record)

    return records


def replace_unified_sheet_fresh_rebuild(
    workbook_path: Path,
    sheets: dict[str, list[list[object]]],
    updated: pd.DataFrame,
    *,
    rows_before: int,
    stats: dict[str, int],
    blank_normalized_receivers: int,
) -> Path:
    """Write a fresh workbook package instead of round-tripping the old .xlsx.

    This mirrors the categoriser safe-write strategy. The old workbook is read
    as values only, UnifiedTransactions is replaced in memory, ChangeLog is
    appended in memory, and a brand-new workbook package is written. Excel Table
    objects are intentionally disabled because real workbook testing showed
    Excel repair warnings for generated table parts.
    """
    headers = [str(col) for col in updated.columns]
    records = dataframe_to_records(updated)
    sheets[UNIFIED_SHEET] = records_to_sheet_values(headers, records)

    details = (
        "Fresh-rebuild write mode. Applied TransactionRules to non-split rows; "
        f"rules loaded: {stats['rules_loaded']}; "
        f"rule matches: {stats['rule_matches']}; "
        f"fields updated: {stats['fields_updated']}; "
        f"skipped existing/manual fields: {stats['fields_skipped_existing_value']}; "
        f"split rows skipped: {stats['rows_skipped_split']}; "
        f"blank NormalizedReceiver cells excluding split rows: {blank_normalized_receivers}"
    )

    append_changelog_row(
        sheets,
        script=SCRIPT_NAME,
        action="Apply budgeting transaction rules",
        sheet=UNIFIED_SHEET,
        rows_updated=stats["fields_updated"],
        status="Completed",
        details=details,
        backup_file=None,
    )

    backup_path, _, writer_stats = replace_with_fresh_workbook(
        workbook_path,
        sheets,
        backup_label="before_normalise_fresh_rebuild",
        basic_formatting=True,
        excel_tables=False,
    )

    print("Created fresh workbook from scratch.")
    print("Validated fresh workbook: OK")
    print(f"Backed up old workbook: {backup_path}")
    print(f"Sheets written: {writer_stats['sheets_written']}")
    print(f"Excel tables added: {writer_stats['tables_added']}")
    print("No loaded workbook object was saved.")

    return backup_path

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Apply budgeting TransactionRules.xlsx to UnifiedTransactions."
    )
    parser.add_argument(
        "--workbook",
        default=str(DEFAULT_WORKBOOK),
        help="Parsed transactions workbook. Defaults to output/budgeting/ParsedTransactions.xlsx.",
    )
    parser.add_argument(
        "--rules",
        default=str(DEFAULT_RULES),
        help="Transaction rules workbook. Defaults to rules/budgeting/TransactionRules.xlsx.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be changed, but do not write the workbook.",
    )
    parser.add_argument(
        "--run-rule",
        metavar="RULEID",
        default=None,
        help=(
            "Narrowly re-apply one TransactionRules rule (by RuleID, e.g. BR0290) to the rows it "
            "matches - including correcting a stale non-blank value a normal run would skip. Always "
            "prints a full report first; combine with --dry-run, --apply, or --apply-rows to control "
            "whether/what gets written."
        ),
    )
    parser.add_argument("--apply", action="store_true", help="With --run-rule: apply all proposed changes without prompting.")
    parser.add_argument(
        "--apply-rows",
        metavar="UNIFIEDID1,UNIFIEDID2,...",
        default=None,
        help="With --run-rule: comma-separated UnifiedIDs to apply (skips the prompt, applies only these).",
    )

    args = parser.parse_args()

    workbook_path = Path(args.workbook).expanduser().resolve()
    rules_path = Path(args.rules).expanduser().resolve()

    if not workbook_path.exists():
        raise FileNotFoundError(f"Parsed transactions workbook not found: {workbook_path}")

    if args.run_rule:
        apply_rows = [v.strip() for v in args.apply_rows.split(",") if v.strip()] if args.apply_rows else None
        run_rule_command(
            workbook_path,
            rules_path,
            args.run_rule,
            apply_all=args.apply,
            apply_rows=apply_rows,
            dry_run=args.dry_run,
        )
        return

    if args.apply or args.apply_rows:
        parser.error("--apply and --apply-rows only apply together with --run-rule")

    rules = load_rules(rules_path)
    print(f"Read previous workbook values only: {workbook_path}")
    sheets, unified = load_unified_from_workbook_values(workbook_path)
    print(f"Loaded sheets: {len(sheets)}")
    print("No loaded workbook object will be saved.")

    updated, stats = apply_rules(unified, rules)
    blank_normalized_receivers = count_blank_normalized_receivers(updated)

    print("Budgeting normaliser complete." if not args.dry_run else "Budgeting normaliser dry run complete.")
    print(f"Workbook:                       {workbook_path}")
    print(f"Rules:                          {rules_path}")
    print(f"Rows checked:                   {stats['rows_checked']}")
    print(f"Rules loaded:                   {stats['rules_loaded']}")
    print(f"Split rows skipped:             {stats['rows_skipped_split']}")
    print(f"Rule matches:                   {stats['rule_matches']}")
    print(f"Fields updated:                 {stats['fields_updated']}")
    print(f"Skipped existing/manual fields: {stats['fields_skipped_existing_value']}")
    print(f"Empty dynamic outputs:          {stats['empty_dynamic_outputs']}")
    print(f"Blank NormalizedReceiver cells: {blank_normalized_receivers}  (excluding split rows)")

    if args.dry_run:
        print("Dry run only: workbook was not modified.")
        return

    replace_unified_sheet_fresh_rebuild(
        workbook_path,
        sheets,
        updated,
        rows_before=len(unified),
        stats=stats,
        blank_normalized_receivers=blank_normalized_receivers,
    )
    print("UnifiedTransactions updated.")


if __name__ == "__main__":
    main()
