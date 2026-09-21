from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def test_project_structure_doc_exists_and_mentions_target_package():
    path = PROJECT_ROOT / "docs" / "project_structure.md"
    assert path.exists()

    text = path.read_text(encoding="utf-8")
    assert "finance_parser/" in text
    assert "budgeting/" in text
    assert "investments/" in text
    assert "tools/seed_account_balance.py" in text
    assert "tools/find_duplicate_transactions.py" in text


def test_cleanup_local_artifacts_doc_exists():
    path = PROJECT_ROOT / "docs" / "cleanup_local_artifacts.md"
    assert path.exists()

    text = path.read_text(encoding="utf-8")
    assert "python utilities/cleanup_local_artifacts.py --dry-run" in text
    assert ".DS_Store" in text
    assert "__pycache__" in text
    assert ".pytest_cache" in text
