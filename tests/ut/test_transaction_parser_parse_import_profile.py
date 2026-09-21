from pathlib import Path

from finance_parser.budgeting import transaction_parser


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


def test_get_parser_modules_returns_budgeting_parsers():
    modules = transaction_parser.get_parser_modules()
    names = {module.SOURCE_BANK for module in modules}
    assert {"OP", "NORWEGIAN", "NORDEA", "SPANKKI"}.issubset(names)
