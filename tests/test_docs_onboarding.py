from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]


LEGACY_BUDGETING_COMMANDS = [
    "python transaction_parser.py",
    "python transaction_normaliser.py",
    "python transaction_categoriser.py",
]


def test_readme_exists_and_mentions_core_commands():
    readme = PROJECT_ROOT / "README.md"
    assert readme.exists()

    text = readme.read_text(encoding="utf-8")
    assert "python run_budgeting_pipeline.py" in text
    assert "python -m finance_parser.budgeting.transaction_parser" in text
    assert "python -m finance_parser.budgeting.transaction_normaliser" in text
    assert "python -m finance_parser.budgeting.transaction_categoriser" in text
    assert "python investment_parser.py" in text
    assert "python -m pytest" in text

    for legacy_command in LEGACY_BUDGETING_COMMANDS:
        assert legacy_command not in text


def test_setup_doc_mentions_template_copy_commands():
    setup = PROJECT_ROOT / "docs" / "setup.md"
    assert setup.exists()

    text = setup.read_text(encoding="utf-8")
    assert "cp config/settings.example.yaml config/settings.yaml" in text
    assert "cp rules/budgeting/TransactionRules.template.xlsx rules/budgeting/TransactionRules.xlsx" in text
    assert "cp rules/investments/InstrumentMaster.template.xlsx rules/investments/InstrumentMaster.xlsx" in text


def test_workflow_doc_mentions_changelog_and_pipeline():
    workflow = PROJECT_ROOT / "docs" / "workflow.md"
    assert workflow.exists()

    text = workflow.read_text(encoding="utf-8")
    assert "ChangeLog" in text
    assert "python run_budgeting_pipeline.py" in text
    assert "python run_budgeting_pipeline.py --dry-run" in text

    for legacy_command in LEGACY_BUDGETING_COMMANDS:
        assert legacy_command not in text
