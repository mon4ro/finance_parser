from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]


def test_utilities_doc_mentions_canonical_commands():
    path = PROJECT_ROOT / "docs" / "utilities.md"
    assert path.exists()

    text = path.read_text(encoding="utf-8")
    assert "python utilities/cleanup_budgeting_source_accounts.py --dry-run" in text
    assert "python utilities/refresh_budgeting_bank_raw_ids.py --dry-run" in text
    assert "python utilities/review_cleanup_unified_duplicates.py --dry-run" in text
    assert "python utilities/cleanup_local_artifacts.py --dry-run" in text
    assert "python utilities/<script>.py --help" in text
