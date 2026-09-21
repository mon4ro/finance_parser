from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]


def test_utility_scripts_bootstrap_project_root_before_common_import():
    for rel in [
        "finance_parser/utilities/cleanup_budgeting_source_accounts.py",
        "finance_parser/utilities/refresh_budgeting_bank_raw_ids.py",
        "finance_parser/utilities/review_cleanup_unified_duplicates.py",
    ]:
        source = (PROJECT_ROOT / rel).read_text(encoding="utf-8")
        assert "sys.path.insert(0, str(PROJECT_ROOT))" in source
        assert source.index("sys.path.insert(0, str(PROJECT_ROOT))") < source.index("from finance_parser.common import")
