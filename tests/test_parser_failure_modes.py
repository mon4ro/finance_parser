import shutil
import types
from pathlib import Path

import pandas as pd
import pytest

from finance_parser.budgeting import transaction_parser as budgeting_parser
from finance_parser.common import IMPORT_LOG_SHEET, RAW_COLUMNS
from finance_parser.investments import investment_parser
from finance_parser.investments.investment_common import INVESTMENT_RAW_COLUMNS

BUDGETING_FIXTURES = Path(__file__).resolve().parent / "fixtures" / "budgeting"


def _fake_parser(source_bank_attr, source_bank, *, matches, raises=None, rows=1):
    """
    A minimal fake parser module: can_parse() controlled by `matches`,
    parse_file() either returns a small valid frame or raises `raises`.
    """
    module = types.SimpleNamespace()
    setattr(module, source_bank_attr, source_bank)

    def can_parse(path):
        return matches, f"{source_bank}: {'matched' if matches else 'no match'}"

    def parse_file(path, imported_at):
        if raises is not None:
            raise raises
        columns = RAW_COLUMNS if source_bank_attr == "SOURCE_BANK" else INVESTMENT_RAW_COLUMNS
        id_col = "RawID" if source_bank_attr == "SOURCE_BANK" else "InvestmentRawID"
        return pd.DataFrame(
            [{c: (f"{source_bank}-{path.name}-{i}" if c == id_col else "") for c in columns} for i in range(rows)],
            columns=columns,
        )

    module.can_parse = can_parse
    module.parse_file = parse_file
    return module


# --- Budgeting side ---

def test_budgeting_unsupported_file_is_skipped_not_fatal(tmp_path, monkeypatch):
    good = _fake_parser("SOURCE_BANK", "GOODBANK", matches=True)
    never_matches = _fake_parser("SOURCE_BANK", "NEVERMATCH", matches=False)
    monkeypatch.setattr(budgeting_parser, "get_parser_modules", lambda: [never_matches, good])

    (tmp_path / "real_export.csv").write_text("date,amount\n2026-01-01,10\n")
    (tmp_path / "unrelated.csv").write_text("not a bank export\n")

    combined_raw, import_log, _ = budgeting_parser.parse_import_files(tmp_path, "2026-01-01 00:00:00")

    assert len(combined_raw) == 2  # both files matched "good" since it always matches=True
    statuses = dict(zip(import_log["SourceFile"], import_log["Status"]))
    assert all(s == "Parsed" for s in statuses.values())


def test_budgeting_truly_unsupported_file_is_skipped_and_others_still_import(tmp_path, monkeypatch):
    only_matches_csv_named_real = types.SimpleNamespace(SOURCE_BANK="GOODBANK")

    def can_parse(path):
        return path.name == "real_export.csv", "only matches real_export.csv"

    def parse_file(path, imported_at):
        return pd.DataFrame([{c: ("GOOD-1" if c == "RawID" else "") for c in RAW_COLUMNS}], columns=RAW_COLUMNS)

    only_matches_csv_named_real.can_parse = can_parse
    only_matches_csv_named_real.parse_file = parse_file

    monkeypatch.setattr(budgeting_parser, "get_parser_modules", lambda: [only_matches_csv_named_real])

    (tmp_path / "real_export.csv").write_text("date,amount\n2026-01-01,10\n")
    (tmp_path / "unrelated_reference_doc.csv").write_text("not a bank export at all\n")

    combined_raw, import_log, _ = budgeting_parser.parse_import_files(tmp_path, "2026-01-01 00:00:00")

    # The real file still imports even though the unrelated one has no matching parser.
    assert len(combined_raw) == 1
    statuses = dict(zip(import_log["SourceFile"], import_log["Status"]))
    assert statuses["real_export.csv"] == "Parsed"
    assert statuses["unrelated_reference_doc.csv"] == "Skipped: unsupported file"


def test_budgeting_format_drift_is_skipped_and_others_still_import(tmp_path, monkeypatch):
    good = _fake_parser("SOURCE_BANK", "GOODBANK", matches=False)
    drifted = types.SimpleNamespace(SOURCE_BANK="DRIFTEDBANK")

    def can_parse(path):
        return path.name == "drifted.csv", "matches by name"

    def parse_file(path, imported_at):
        raise ValueError("Missing required column: Amount")

    drifted.can_parse = can_parse
    drifted.parse_file = parse_file

    def good_can_parse(path):
        return path.name == "fine.csv", "matches by name"

    good.can_parse = good_can_parse

    monkeypatch.setattr(budgeting_parser, "get_parser_modules", lambda: [drifted, good])

    (tmp_path / "drifted.csv").write_text("garbage\n")
    (tmp_path / "fine.csv").write_text("date,amount\n2026-01-01,10\n")

    combined_raw, import_log, _ = budgeting_parser.parse_import_files(tmp_path, "2026-01-01 00:00:00")

    assert len(combined_raw) == 1
    statuses = dict(zip(import_log["SourceFile"], import_log["Status"]))
    assert statuses["fine.csv"] == "Parsed"
    assert statuses["drifted.csv"] == "Skipped: format drift"
    assert "Missing required column" in import_log[import_log["SourceFile"] == "drifted.csv"]["Notes"].iloc[0]


# --- Investment side (same pattern, mirrors the budgeting fix) ---

def test_investment_unsupported_and_format_drift_are_both_non_fatal(tmp_path, monkeypatch):
    drifted = types.SimpleNamespace(BROKER="DRIFTEDBROKER")

    def drifted_can_parse(path):
        return path.name == "drifted.csv", "matches by name"

    def drifted_parse_file(path, imported_at):
        raise ValueError("Column layout changed")

    drifted.can_parse = drifted_can_parse
    drifted.parse_file = drifted_parse_file

    good = types.SimpleNamespace(BROKER="GOODBROKER")

    def good_can_parse(path):
        return path.name == "fine.csv", "matches by name"

    def good_parse_file(path, imported_at):
        return pd.DataFrame(
            [{c: ("GOOD-1" if c == "InvestmentRawID" else "") for c in INVESTMENT_RAW_COLUMNS}],
            columns=INVESTMENT_RAW_COLUMNS,
        )

    good.can_parse = good_can_parse
    good.parse_file = good_parse_file

    monkeypatch.setattr(investment_parser, "SUPPORTED_INVESTMENT_PARSERS", [drifted, good])

    (tmp_path / "drifted.csv").write_text("garbage\n")
    (tmp_path / "fine.csv").write_text("garbage2\n")
    (tmp_path / "unrelated.csv").write_text("not a broker export\n")

    combined_raw, import_log, _ = investment_parser.parse_import_files(tmp_path, "2026-01-01 00:00:00")

    assert len(combined_raw) == 1
    statuses = dict(zip(import_log["SourceFile"], import_log["Status"]))
    assert statuses["fine.csv"] == "Parsed"
    assert statuses["drifted.csv"] == "Skipped: format drift"
    assert statuses["unrelated.csv"] == "Skipped: unsupported file"


# --- End-to-end: the persisted ImportLog sheet must keep the skipped status,
# not get overwritten by append_to_output()'s later "Status" = "Imported"
# pass. This is the actual layer the investment-side bug lived in - a
# per-file guard could still leave one blanket unconditional assignment
# downstream that undoes it, which is exactly what happened.

def test_budgeting_append_to_output_persists_skipped_status(tmp_path):
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    shutil.copy(BUDGETING_FIXTURES / "cash_sample.xlsx", input_dir / "cash_sample.xlsx")
    (input_dir / "unrelated_reference.txt").write_text("not a bank export at all")

    output_path = tmp_path / "ParsedTransactions.xlsx"

    budgeting_parser.append_to_output(input_dir, output_path)

    log = pd.read_excel(output_path, sheet_name=IMPORT_LOG_SHEET, dtype=object)
    status_by_file = dict(zip(log["SourceFile"], log["Status"]))
    assert status_by_file["cash_sample.xlsx"] == "Imported"
    assert status_by_file["unrelated_reference.txt"] == "Skipped: unsupported file"
