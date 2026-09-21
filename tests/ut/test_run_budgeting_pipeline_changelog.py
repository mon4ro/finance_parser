import pandas as pd
from openpyxl import Workbook

from finance_parser.budgeting.run_budgeting_pipeline import print_changelog
from finance_parser.common import CHANGE_LOG_COLUMNS


def _write_changelog(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "ChangeLog"
    ws.append(CHANGE_LOG_COLUMNS)
    for row in rows:
        ws.append([row.get(h, "") for h in CHANGE_LOG_COLUMNS])
    wb.save(path)


def _base_row(**overrides):
    row = {h: "" for h in CHANGE_LOG_COLUMNS}
    row.update(overrides)
    return row


def test_print_changelog_shows_blank_timestamp_as_same_run(tmp_path, capsys):
    """
    Real bug found and fixed: a blank ChangedAt cell reads back from Excel
    as a pandas NaN float, and str(nan) produces the literal text "nan" -
    the "or" fallback for a blank string never triggers on that, so the
    naive fix silently printed "[nan]" instead of the intended
    "(same run as above)" placeholder.
    """
    path = tmp_path / "log.xlsx"
    _write_changelog(path, [
        _base_row(ChangedAt="2026-06-01 12:00:00", Script="parser.py", Action="Import", Status="Completed"),
        _base_row(ChangedAt="", Script="normaliser.py", Action="Normalise", Status="Completed"),
    ])

    print_changelog(path)

    out = capsys.readouterr().out
    assert "2026-06-01 12:00:00" in out
    assert "(same run as above)" in out
    assert "nan" not in out.lower().replace("(same run as above)", "")


def test_print_changelog_filters_by_script(tmp_path, capsys):
    path = tmp_path / "log.xlsx"
    _write_changelog(path, [
        _base_row(ChangedAt="2026-06-01 12:00:00", Script="parser.py", Action="Import", Status="Completed"),
        _base_row(ChangedAt="", Script="normaliser.py", Action="Normalise", Status="Completed"),
    ])

    print_changelog(path, script="normaliser.py")

    out = capsys.readouterr().out
    assert "normaliser.py" in out
    assert "parser.py" not in out


def test_print_changelog_last_n_shows_only_the_tail(tmp_path, capsys):
    path = tmp_path / "log.xlsx"
    _write_changelog(path, [
        _base_row(ChangedAt=f"2026-06-0{i} 12:00:00", Script="parser.py", Action=f"Run{i}", Status="Completed")
        for i in range(1, 6)
    ])

    print_changelog(path, last=2, show_all=False)

    out = capsys.readouterr().out
    assert "Run4" in out
    assert "Run5" in out
    assert "Run1" not in out
    assert "(2 of 5 matching entries shown)" in out


def test_print_changelog_all_shows_every_entry(tmp_path, capsys):
    path = tmp_path / "log.xlsx"
    _write_changelog(path, [
        _base_row(ChangedAt=f"2026-06-0{i} 12:00:00", Script="parser.py", Action=f"Run{i}", Status="Completed")
        for i in range(1, 6)
    ])

    print_changelog(path, last=2, show_all=True)

    out = capsys.readouterr().out
    assert "Run1" in out
    assert "Run5" in out
    assert "(5 of 5 matching entries shown)" in out


def test_print_changelog_no_matches_reports_clearly(tmp_path, capsys):
    path = tmp_path / "log.xlsx"
    _write_changelog(path, [
        _base_row(ChangedAt="2026-06-01 12:00:00", Script="parser.py", Action="Import", Status="Completed"),
    ])

    print_changelog(path, script="does_not_exist.py")

    out = capsys.readouterr().out
    assert "No ChangeLog entries match" in out


def test_print_changelog_missing_workbook_reports_clearly(tmp_path, capsys):
    print_changelog(tmp_path / "does_not_exist.xlsx")

    out = capsys.readouterr().out
    assert "does not exist" in out


def test_print_changelog_missing_sheet_reports_clearly(tmp_path, capsys):
    path = tmp_path / "no_changelog.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "SomethingElse"
    ws.append(["A"])
    wb.save(path)

    print_changelog(path)

    out = capsys.readouterr().out
    assert "No ChangeLog sheet found" in out


def test_print_changelog_shows_rows_summary(tmp_path, capsys):
    path = tmp_path / "log.xlsx"
    _write_changelog(path, [
        _base_row(
            ChangedAt="2026-06-01 12:00:00", Script="parser.py", Action="Import", Status="Completed",
            RowsBefore=100, RowsAfter=110, RowsAdded=10,
        ),
    ])

    print_changelog(path)

    out = capsys.readouterr().out
    assert "before=100" in out
    assert "after=110" in out
    assert "added=10" in out
