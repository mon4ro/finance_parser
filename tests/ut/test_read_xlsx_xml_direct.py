from openpyxl import Workbook

from finance_parser.common import read_xlsx_xml_direct


REQUIRED = ["Date", "Amount", "Receiver"]


def _write_workbook(path, sheets: dict):
    wb = Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
    wb.save(path)


def test_single_sheet_file_reads_normally(tmp_path):
    path = tmp_path / "single.xlsx"
    _write_workbook(path, {
        "Sheet1": [
            ["Date", "Amount", "Receiver"],
            ["2026-01-01", -10, "Shop A"],
            ["2026-01-02", -20, "Shop B"],
        ],
    })

    df = read_xlsx_xml_direct(path, required_columns=REQUIRED)

    assert len(df) == 2
    assert list(df.columns) == REQUIRED


def test_multiple_sheets_with_matching_columns_are_merged(tmp_path):
    """
    Real bug found and fixed: a genuine multi-sheet export (e.g. a "current
    year" sheet plus one sheet per historical year, all with the identical
    real column layout) had every sheet but the single best-scoring one
    silently discarded - found via a real dry-run where an entire year of
    real transactions from a legitimate historical sheet vanished.
    """
    path = tmp_path / "multi.xlsx"
    _write_workbook(path, {
        "CurrentYear": [
            ["Date", "Amount", "Receiver"],
            ["2026-01-01", -10, "Shop A"],
        ],
        "HistoricalYear1": [
            ["Date", "Amount", "Receiver"],
            ["2024-01-01", -30, "Shop C"],
            ["2024-06-01", 40, "Shop D"],
        ],
        "HistoricalYear2": [
            ["Date", "Amount", "Receiver"],
            ["2023-01-01", -50, "Shop E"],
        ],
    })

    df = read_xlsx_xml_direct(path, required_columns=REQUIRED)

    assert len(df) == 4
    assert set(df["Receiver"]) == {"Shop A", "Shop C", "Shop D", "Shop E"}


def test_sheet_with_different_column_layout_is_not_merged(tmp_path, capsys):
    """
    A sheet with a genuinely different shape (e.g. a summary/pivot tab, or
    a differently-structured export with extra leading columns and no
    recognisable header) must not be blindly merged in just because it also
    happens to produce some rows - only sheets matching the primary sheet's
    exact column set are merged. It IS still flagged with a loud warning
    though - real case this protects against: a genuine historical-data
    sheet once took exactly this shape (see the merge fix), so a human
    should get the chance to verify a shape-mismatched sheet with real rows
    isn't actually real data being silently left out.
    """
    path = tmp_path / "mixed.xlsx"
    _write_workbook(path, {
        "MainData": [
            ["Date", "Amount", "Receiver"],
            ["2026-01-01", -10, "Shop A"],
            ["2026-01-02", -20, "Shop B"],
        ],
        "SummaryTab": [
            ["Account", "TotalSpend", "Notes"],
            ["HOUSEHOLD", -30, "monthly summary"],
        ],
    })

    df = read_xlsx_xml_direct(path, required_columns=REQUIRED)

    assert len(df) == 2
    assert set(df["Receiver"]) == {"Shop A", "Shop B"}

    captured = capsys.readouterr()
    assert "WARNING" in captured.out
    assert "SummaryTab" in captured.out
    assert "1 row" in captured.out


def test_empty_sheets_are_ignored(tmp_path, capsys):
    path = tmp_path / "with_empty.xlsx"
    _write_workbook(path, {
        "MainData": [
            ["Date", "Amount", "Receiver"],
            ["2026-01-01", -10, "Shop A"],
        ],
        "EmptyPlaceholder": [],
    })

    df = read_xlsx_xml_direct(path, required_columns=REQUIRED)

    assert len(df) == 1
    captured = capsys.readouterr()
    assert "WARNING" not in captured.out
