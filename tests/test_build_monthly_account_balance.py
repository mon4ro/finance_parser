from datetime import date
from pathlib import Path

from openpyxl import Workbook

from finance_parser.settings import AppSettings
from finance_parser.budgeting.build_monthly_account_balance import build_monthly_balances


UNIFIED_COLUMNS = [
    "UnifiedID", "RawID", "SourceAccount", "SourceBank", "TransactionType",
    "Date", "Month", "Year", "Amount", "Balance", "RawReceiver", "NormalizedReceiver",
    "Description", "Message", "Include", "Owner", "Supercategory",
    "Category", "Subcategory", "ReviewStatus", "Review/Notes",
]


def _write_unified(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "UnifiedTransactions"
    ws.append(UNIFIED_COLUMNS)
    for row in rows:
        ws.append([row.get(h, "") for h in UNIFIED_COLUMNS])
    wb.save(path)


_BASE_SETTINGS_YAML = """
project:
  name: "Test"
  locale: "fi_FI"
  default_currency: "EUR"

paths:
  budgeting_input: "input/budgeting"
  budgeting_output: "output/budgeting/ParsedTransactions.xlsx"
  budgeting_rules: "rules/budgeting/TransactionRules.xlsx"
  investment_input: "input/investments"
  investment_output: "output/investments/ParsedInvestments.xlsx"
  investment_rules: "rules/investments/InstrumentMaster.xlsx"

investments: {{}}

{extra}
"""


def _settings_with_seeds(extra_yaml: str, tmp_path: Path) -> AppSettings:
    settings_path = tmp_path / "settings.yaml"
    settings_path.write_text(_BASE_SETTINGS_YAML.format(extra=extra_yaml), encoding="utf-8")
    return AppSettings.load(settings_path=settings_path, example_path=Path("does-not-exist.yaml"))


def test_nordea_uses_raw_balance_directly(tmp_path):
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_unified(path, [
        {"SourceAccount": "PERSONAL", "SourceBank": "NORDEA", "Date": "2026-01-10", "Amount": -50, "Balance": 950, "Owner": "PERSONAL"},
        {"SourceAccount": "PERSONAL", "SourceBank": "NORDEA", "Date": "2026-01-25", "Amount": -20, "Balance": 930, "Owner": "PERSONAL"},
        {"SourceAccount": "PERSONAL", "SourceBank": "NORDEA", "Date": "2026-02-05", "Amount": 100, "Balance": 1030, "Owner": "PERSONAL"},
    ])
    settings = _settings_with_seeds("budgeting:\n  default_include: 'YES'\n", tmp_path)

    result, stats = build_monthly_balances(path, settings, as_of=date(2026, 3, 1))

    rows = result[result["SourceAccount"] == "PERSONAL"]
    assert list(rows["MonthEnd"]) == ["2026-01-31", "2026-02-28"]
    assert list(rows["Balance"]) == [930.0, 1030.0]
    assert list(rows["BalanceSource"]) == ["RAW_EXPORT", "RAW_EXPORT"]
    assert list(rows["Owner"]) == ["PERSONAL", "PERSONAL"]
    assert "PERSONAL" in stats["accounts_processed"]


def test_nordea_account_with_no_balance_data_is_flagged_not_silently_dropped(tmp_path):
    """
    Real bug found and fixed: an account correctly identified as a raw-
    export bank (Nordea) but with zero usable Balance data (real case: the
    column existed but was blank for every row, before UnifiedTransactions
    carried it through at all) must be explicitly flagged, not silently
    disappear from every stats bucket.
    """
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_unified(path, [
        {"SourceAccount": "PERSONAL", "SourceBank": "NORDEA", "Date": "2026-01-10", "Amount": -50, "Owner": "PERSONAL"},
    ])
    settings = _settings_with_seeds("budgeting:\n  default_include: 'YES'\n", tmp_path)

    result, stats = build_monthly_balances(path, settings, as_of=date(2026, 3, 1))

    assert len(result[result["SourceAccount"] == "PERSONAL"]) == 0
    assert stats["accounts_processed"] == []
    assert stats["accounts_no_seed_configured"] == []
    assert stats["accounts_raw_export_no_balance_data"] == ["PERSONAL"]


def test_reconstructed_account_sums_from_seed_forward(tmp_path):
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_unified(path, [
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "Date": "2026-01-15", "Amount": -30, "Owner": "HOUSEHOLD"},
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "Date": "2026-02-10", "Amount": 200, "Owner": "HOUSEHOLD"},
    ])
    settings = _settings_with_seeds(
        """
budgeting:
  account_balance_seeds:
    HOUSEHOLD:
      "2026-01-15": 1000.0
""",
        tmp_path,
    )

    result, stats = build_monthly_balances(path, settings, as_of=date(2026, 3, 1))

    rows = result[result["SourceAccount"] == "HOUSEHOLD"]
    # Seed date 01-15 is end-of-day, so its own -30 is already baked in.
    assert list(rows["MonthEnd"]) == ["2026-01-31", "2026-02-28"]
    assert list(rows["Balance"]) == [1000.0, 1200.0]
    assert list(rows["BalanceSource"]) == ["RECONSTRUCTED", "RECONSTRUCTED"]
    assert "HOUSEHOLD" in stats["accounts_processed"]


def test_account_with_no_seed_is_flagged_not_silently_skipped(tmp_path):
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_unified(path, [
        {"SourceAccount": "SPANKKI", "SourceBank": "SPANKKI", "Date": "2026-01-15", "Amount": -30, "Owner": "SHARED"},
    ])
    settings = _settings_with_seeds("budgeting:\n  default_include: 'YES'\n", tmp_path)

    result, stats = build_monthly_balances(path, settings, as_of=date(2026, 3, 1))

    assert len(result[result["SourceAccount"] == "SPANKKI"]) == 0
    assert stats["accounts_no_seed_configured"] == ["SPANKKI"]
    assert stats["accounts_processed"] == []


def test_cash_source_account_is_excluded(tmp_path):
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_unified(path, [
        {"SourceAccount": "CASH", "SourceBank": "CASH", "Date": "2026-01-15", "Amount": -10, "Owner": "SHARED"},
    ])
    settings = _settings_with_seeds(
        'budgeting:\n  account_balance_seeds:\n    CASH:\n      "2026-01-01": 0\n',
        tmp_path,
    )

    result, stats = build_monthly_balances(path, settings, as_of=date(2026, 3, 1))

    assert len(result) == 0
    assert stats["accounts_processed"] == []
    assert stats["accounts_no_seed_configured"] == []


def test_second_seed_is_used_once_its_date_arrives(tmp_path):
    """
    Periodic recalibration: adding a second seed later must be picked up
    starting from ITS date, not retroactively rewrite months already
    computed from the first seed.
    """
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_unified(path, [
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "Date": "2026-02-01", "Amount": -10, "Owner": "HOUSEHOLD"},
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "Date": "2026-03-05", "Amount": 5, "Owner": "HOUSEHOLD"},
    ])
    settings = _settings_with_seeds(
        """
budgeting:
  account_balance_seeds:
    HOUSEHOLD:
      "2026-01-01": 1000.0
      "2026-02-15": 5000.0
""",
        tmp_path,
    )

    result, stats = build_monthly_balances(path, settings, as_of=date(2026, 4, 1))

    rows = result[result["SourceAccount"] == "HOUSEHOLD"]
    by_month = dict(zip(rows["MonthEnd"], rows["Balance"]))
    # Jan: seed1 (1000) + no transactions after it that month = 1000
    assert by_month["2026-01-31"] == 1000.0
    # Feb: seed2 (5000, effective 02-15) - the 02-01 transaction predates it, ignored
    assert by_month["2026-02-28"] == 5000.0
    # Mar: seed2 + the 03-05 transaction (+5)
    assert by_month["2026-03-31"] == 5005.0
