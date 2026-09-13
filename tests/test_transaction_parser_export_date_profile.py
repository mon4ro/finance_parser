from pathlib import Path

from finance_parser.budgeting import transaction_parser


def test_export_date_from_file_returns_timestamp_for_fixture():
    fixture = Path(__file__).resolve().parent / "fixtures" / "budgeting" / "spankki_sample.csv"
    value = transaction_parser.export_date_from_file(fixture)
    assert isinstance(value, str)
    assert len(value) == 19


def test_parse_import_files_profile_with_fixture_reaches_parse_loop():
    fixture = Path(__file__).resolve().parent / "fixtures" / "budgeting" / "spankki_sample.csv"

    raw, log, bank_raw = transaction_parser.parse_import_files(
        fixture,
        "2026-06-04 12:00:00",
        profile=True,
    )

    assert len(raw) == 1
    assert len(log) == 1
    assert log.iloc[0]["ExportDate"]
    assert "SpankkiRawExport" in bank_raw
