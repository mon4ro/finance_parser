from pathlib import Path

from parsers import op, norwegian, spankki


FIXTURES = Path(__file__).resolve().parent / "fixtures" / "budgeting"


def test_spankki_fixture_parses_expected_fields():
    path = FIXTURES / "spankki_sample.csv"
    ok, reason = spankki.can_parse(path)
    assert ok, reason

    df = spankki.parse_file(path, "2026-06-04 12:00:00")
    row = df.iloc[0]

    assert row["SourceBank"] == "SPANKKI"
    assert row["SourceAccount"] == "SPANKKI"
    assert row["Amount"] == -4.50
    assert row["RawReceiver"] == "S-MARKET"
    assert row["ArchiveID"] == "20260601990000900000"

    raw = spankki.parse_bank_raw_rows(path, "2026-06-04 12:00:00")
    assert raw.iloc[0]["BankRawExportID"].startswith("SPKRAW-")


def test_op_budgeting_fixture_parses_sourceaccount_from_filename():
    path = FIXTURES / "HOUSEHOLD_tapahtumat20260601-20260603.xlsx"
    ok, reason = op.can_parse(path)
    assert ok, reason

    df = op.parse_file(path, "2026-06-04 12:00:00")
    row = df.iloc[0]

    assert row["SourceBank"] == "OP"
    assert row["SourceAccount"] == "HOUSEHOLD"
    assert row["Amount"] == -10.0
    assert row["RawReceiver"] == "PRISMA ESPOO"
    assert row["ArchiveID"] == "OPARCH123"


def test_norwegian_fixture_parses_expected_fields():
    path = FIXTURES / "norwegian_sample.xlsx"
    ok, reason = norwegian.can_parse(path)
    assert ok, reason

    df = norwegian.parse_file(path, "2026-06-04 12:00:00")
    row = df.iloc[0]

    assert row["SourceBank"] == "NORWEGIAN"
    assert row["SourceAccount"] == "NORWEGIAN"
    assert row["Amount"] == -12.34
    assert row["RawReceiver"] == "K-MARKET TEST"

    raw = norwegian.parse_bank_raw_rows(path, "2026-06-04 12:00:00")
    assert raw.iloc[0]["BankRawExportID"].startswith("NWGRAW-")
