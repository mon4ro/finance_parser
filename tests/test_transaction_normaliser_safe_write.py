from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
NORMALISER = PROJECT_ROOT / "finance_parser" / "budgeting" / "transaction_normaliser.py"


def compact_source(source: str) -> str:
    return "".join(source.split())


def test_transaction_normaliser_uses_fresh_rebuild_write_path():
    source = NORMALISER.read_text(encoding="utf-8")
    compact = compact_source(source)

    assert "replace_with_fresh_workbook" in source
    assert "append_changelog_row" in source
    assert "read_workbook_values_only" in source

    # The normaliser must use the same corruption-resistant write strategy as
    # the categoriser: rebuild the workbook package from values, with generated
    # Excel table objects disabled.
    assert "excel_tables=False" in compact


def test_transaction_normaliser_does_not_mutate_existing_workbook_package():
    source = NORMALISER.read_text(encoding="utf-8")

    # These were part of the old openpyxl-mutating write path. That path can
    # leave the workbook package in a state where Excel reports XML repair
    # warnings, so the normaliser should not reintroduce it.
    assert "def replace_sheet" not in source
    assert "ws = wb[sheet_name]" not in source
    assert "wb.save(tmp_path)" not in source
    assert "_tmp_normalise" not in source
    assert "backup_before_normalise" not in source


def test_transaction_normaliser_changelog_is_written_in_memory():
    source = NORMALISER.read_text(encoding="utf-8")

    assert "append_changelog_row" in source
    assert "append_change_log_entry" not in source
