from pathlib import Path

from finance_parser.common import CHANGE_LOG_COLUMNS
from finance_parser.utilities.fresh_workbook_writer import append_changelog_row


def test_fresh_sheet_header_matches_change_log_columns_exactly():
    """
    Real inconsistency this fixes, found by eye on the live household
    workbook (2026-10-03): append_change_log_entry() (common.py, used by
    scripts still mutating the workbook directly with openpyxl) always
    wrote the full CHANGE_LOG_COLUMNS schema; this function's own default
    header set was a different, smaller 8-column list with no ChangeID,
    Workbook, RowsAdded, or ReviewFile at all - not because callers chose
    to omit them, but because nothing in this function ever set them.
    Since fresh-rebuild is the project's documented default write mode,
    almost every real row ended up missing those four columns.
    """
    sheets: dict[str, list[list[object]]] = {}

    append_changelog_row(
        sheets, script="test.py", action="Do a thing", sheet="SomeSheet",
        rows_updated=1, status="Completed", details="",
    )

    assert sheets["ChangeLog"][0] == CHANGE_LOG_COLUMNS


def test_workbook_rows_added_and_review_file_are_now_populated():
    sheets: dict[str, list[list[object]]] = {}

    append_changelog_row(
        sheets, script="test.py", action="Do a thing", sheet="SomeSheet",
        rows_updated=5, status="Completed", details="",
        workbook=Path("/some/path/ParsedTransactions.xlsx"),
        rows_added=3, review_file="review_2026_10.xlsx",
    )

    headers = sheets["ChangeLog"][0]
    row = dict(zip(headers, sheets["ChangeLog"][1]))
    assert row["Workbook"] == "ParsedTransactions.xlsx"
    assert row["RowsAdded"] == 3
    assert row["ReviewFile"] == "review_2026_10.xlsx"


def test_change_id_is_generated_and_unique_per_row():
    sheets: dict[str, list[list[object]]] = {}

    for _ in range(3):
        append_changelog_row(
            sheets, script="test.py", action="Do a thing", sheet="SomeSheet",
            rows_updated=1, status="Completed", details="",
        )

    headers = sheets["ChangeLog"][0]
    change_id_col = headers.index("ChangeID")
    change_ids = [row[change_id_col] for row in sheets["ChangeLog"][1:]]
    assert all(cid for cid in change_ids)
    assert len(set(change_ids)) == 3


def test_omitted_optional_fields_default_to_blank_not_none():
    sheets: dict[str, list[list[object]]] = {}

    append_changelog_row(
        sheets, script="test.py", action="Do a thing", sheet="SomeSheet",
        rows_updated=1, status="Completed", details="",
    )

    headers = sheets["ChangeLog"][0]
    row = dict(zip(headers, sheets["ChangeLog"][1]))
    assert row["Workbook"] == ""
    assert row["RowsAdded"] == ""
    assert row["ReviewFile"] == ""


def test_existing_sheet_with_the_full_schema_is_carried_forward_correctly():
    """
    Matches the real household workbook's actual state: ChangeLog already
    has the full CHANGE_LOG_COLUMNS header (created by
    append_change_log_entry's first-ever call) by the time a fresh-rebuild
    script's append_changelog_row() call runs against it.
    """
    sheets: dict[str, list[list[object]]] = {
        "ChangeLog": [
            list(CHANGE_LOG_COLUMNS),
            ["CHG-OLD", "2026-01-01 00:00:00", "parser.py", "Import", "ParsedTransactions.xlsx",
             "RawTransactions", 100, 110, 10, "", "", "", "", "Completed", "old entry"],
        ],
    }

    append_changelog_row(
        sheets, script="normaliser.py", action="Apply rules", sheet="UnifiedTransactions",
        rows_updated=42, status="Completed", details="new entry",
        workbook=Path("ParsedTransactions.xlsx"),
    )

    assert len(sheets["ChangeLog"]) == 3
    headers = sheets["ChangeLog"][0]
    new_row = dict(zip(headers, sheets["ChangeLog"][2]))
    assert new_row["Script"] == "normaliser.py"
    assert new_row["Workbook"] == "ParsedTransactions.xlsx"
    assert new_row["RowsUpdated"] == 42
    old_row = dict(zip(headers, sheets["ChangeLog"][1]))
    assert old_row["ChangeID"] == "CHG-OLD"
