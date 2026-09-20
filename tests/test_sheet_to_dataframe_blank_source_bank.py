from openpyxl import Workbook

from finance_parser.common import IMPORT_LOG_COLUMNS, sheet_to_dataframe


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
    from finance_parser.common import RAW_COLUMNS

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
