from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]


def test_workbook_utilities_define_help_before_common_import():
    for rel in [
        "finance_parser/utilities/cleanup_budgeting_source_accounts.py",
        "finance_parser/utilities/refresh_budgeting_bank_raw_ids.py",
        "finance_parser/utilities/review_cleanup_unified_duplicates.py",
    ]:
        source = (PROJECT_ROOT / rel).read_text(encoding="utf-8")
        assert "_print_cli_help_and_exit()" in source
        assert source.index("_print_cli_help_and_exit()") < source.index("from finance_parser.common import")
