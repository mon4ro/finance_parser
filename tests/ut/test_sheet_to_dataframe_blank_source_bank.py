from openpyxl import Workbook

from finance_parser.common import IMPORT_LOG_COLUMNS, RAW_COLUMNS, sheet_to_dataframe


def test_blank_source_bank_row_does_not_crash_and_stays_blank(tmp_path):
    """
    Real bug: sheet_to_dataframe() re-reading an existing ImportLog sheet
    with a blank SourceBank row (e.g. a "Skipped: unsupported file" entry,
    which is never associated with any real bank) called an undefined
    function (infer_source_bank_from_existing_row) and crashed the whole
    pipeline on the very next stage. Fixed by reusing the already-tested
    canonical_source_bank(), which correctly returns "" when nothing can be
    inferred rather than guessing.
    """
    path = tmp_path / "log.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "ImportLog"
    ws.append(IMPORT_LOG_COLUMNS)
    ws.append([
        "RUN-test", "2026-09-20 12:00:00", "", "not_a_bank_export.xlsx",
        0, 0, 0, "Skipped: unsupported file", "",
    ])
    wb.save(path)

    df = sheet_to_dataframe(path, "ImportLog", IMPORT_LOG_COLUMNS)

    assert len(df) == 1
    assert df.iloc[0]["SourceBank"] in ("", None) or str(df.iloc[0]["SourceBank"]).strip() == ""


def test_blank_source_bank_row_infers_from_raw_id_prefix(tmp_path):
    """
    A blank SourceBank on a row that DOES have a recognisable RawID prefix
    (the original documented purpose of this migration path - an older
    output file predating the SourceBank column) should still be inferred.
    """
    path = tmp_path / "raw.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "RawTransactions"
    columns = [c for c in RAW_COLUMNS if c != "SourceBank"]
    ws.append(columns)
    row = {c: "" for c in columns}
    row["RawID"] = "OP-abc123"
    row["SourceAccount"] = "HOUSEHOLD"
    ws.append([row[c] for c in columns])
    wb.save(path)

    df = sheet_to_dataframe(path, "RawTransactions", RAW_COLUMNS)

    assert df.iloc[0]["SourceBank"] == "OP"


def test_tili_prefixed_source_account_self_heals_on_read(tmp_path):
    """
    Real bug: infer_source_account_from_filename() strips a generic "Tili"
    filename prefix for NEW imports (see its own docstring), but that fix
    can't retroactively touch rows already written before it existed - a
    real account ended up split across two different SourceAccount values
    ("CHILD" and "TILI CHILD") purely depending on when each row was
    imported, silently breaking OwnershipRules matching and balance
    reconstruction for the older half. sheet_to_dataframe() must merge
    "TILI <name>" back to "<name>" on every read so both halves reunite.
    """
    path = tmp_path / "raw.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "RawTransactions"
    ws.append(RAW_COLUMNS)
    row = {c: "" for c in RAW_COLUMNS}
    row["RawID"] = "OP-abc123"
    row["SourceAccount"] = "TILI CHILD"
    row["SourceBank"] = "OP"
    ws.append([row[c] for c in RAW_COLUMNS])
    wb.save(path)

    df = sheet_to_dataframe(path, "RawTransactions", RAW_COLUMNS)

    assert df.iloc[0]["SourceAccount"] == "CHILD"


def test_plain_source_account_untouched_by_tili_migration(tmp_path):
    """A normal, already-correct SourceAccount must not be touched just
    because it happens to start with the same letters ("TILING" etc.)."""
    path = tmp_path / "raw.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "RawTransactions"
    ws.append(RAW_COLUMNS)
    row = {c: "" for c in RAW_COLUMNS}
    row["RawID"] = "OP-abc123"
    row["SourceAccount"] = "CHILD"
    row["SourceBank"] = "OP"
    ws.append([row[c] for c in RAW_COLUMNS])
    wb.save(path)

    df = sheet_to_dataframe(path, "RawTransactions", RAW_COLUMNS)

    assert df.iloc[0]["SourceAccount"] == "CHILD"
