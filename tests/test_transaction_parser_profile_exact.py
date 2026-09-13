from pathlib import Path

from finance_parser.budgeting import transaction_parser


PROJECT_ROOT = Path(__file__).resolve().parents[1]
PARSER = PROJECT_ROOT / "finance_parser" / "budgeting" / "transaction_parser.py"


def test_get_parser_modules_returns_supported_parsers():
    modules = transaction_parser.get_parser_modules()
    assert modules == transaction_parser.SUPPORTED_PARSERS
    assert {module.SOURCE_BANK for module in modules} == {"OP", "NORWEGIAN", "NORDEA", "SPANKKI", "CASH"}


def test_transaction_parser_source_has_no_missing_parsers_reference():
    source = PARSER.read_text(encoding="utf-8")
    assert "for parser_module in PARSERS:" not in source


def test_parse_import_files_profile_with_fixture_reaches_parse_loop():
    fixture = Path(__file__).resolve().parent / "fixtures" / "budgeting" / "spankki_sample.csv"

    raw, log, bank_raw = transaction_parser.parse_import_files(
        fixture,
        "2026-06-04 12:00:00",
        profile=True,
    )

    assert len(raw) == 1
    assert len(log) == 1
    assert "SpankkiRawExport" in bank_raw
