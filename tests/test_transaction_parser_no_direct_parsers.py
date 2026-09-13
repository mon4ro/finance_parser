from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
PARSER = PROJECT_ROOT / "finance_parser" / "budgeting" / "transaction_parser.py"


def test_transaction_parser_does_not_reference_missing_parsers_global():
    source = PARSER.read_text(encoding="utf-8")
    assert "for parser_module in PARSERS:" not in source
