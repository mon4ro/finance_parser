from datetime import date
from pathlib import Path

from openpyxl import Workbook

from finance_parser.settings import AppSettings
from finance_parser.budgeting.build_monthly_account_balance import build_monthly_balances


RAW_COLUMNS = [
    "RawID", "SourceAccount", "SourceBank", "ExportDate", "ImportedAt",
    "BookingDate", "ValueDate", "Amount", "TransactionTypeRaw", "Description",
    "RawReceiver", "ReceiverAccount", "ReceiverBankBIC", "Reference", "Message",
    "ArchiveID", "Balance", "CurrencyAmount", "Currency", "Rate",
    "MerchantArea", "MerchantCategory", "SourceFile",
]


def _write_raw(path, rows, owners=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "RawTransactions"
    ws.append(RAW_COLUMNS)
    for i, row in enumerate(rows):
        row = {**row}
        row.setdefault("RawID", f"R-{i}")
        ws.append([row.get(h, "") for h in RAW_COLUMNS])

    if owners:
        unified_ws = wb.create_sheet("UnifiedTransactions")
        unified_ws.append(["SourceAccount", "Owner"])
        for account, owner in owners.items():
            unified_ws.append([account, owner])

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
    _write_raw(path, [
        {"SourceAccount": "PERSONAL", "SourceBank": "NORDEA", "BookingDate": "2026-01-10", "Amount": -50, "Balance": 950},
        {"SourceAccount": "PERSONAL", "SourceBank": "NORDEA", "BookingDate": "2026-01-25", "Amount": -20, "Balance": 930},
        {"SourceAccount": "PERSONAL", "SourceBank": "NORDEA", "BookingDate": "2026-02-05", "Amount": 100, "Balance": 1030},
        {"SourceAccount": "PERSONAL", "SourceBank": "NORDEA", "BookingDate": "2026-03-01", "Amount": 1, "Balance": 1031},
    ], owners={"PERSONAL": "PERSONAL"})
    settings = _settings_with_seeds("budgeting:\n  default_include: 'YES'\n", tmp_path)

    result, stats = build_monthly_balances(path, settings, as_of=date(2026, 3, 5))

    rows = result[result["SourceAccount"] == "PERSONAL"]
    # February's own end-of-month row is only trustworthy once there's a
    # real transaction dated AFTER Feb 28 (here, the Mar 1 row) proving the
    # import continued past it - not just because Feb 5 happens to be the
    # latest transaction and the calendar says February is over.
    assert list(rows["MonthEnd"]) == ["2026-01-31", "2026-02-28"]
    assert list(rows["Balance"]) == [930.0, 1030.0]
    assert list(rows["BalanceSource"]) == ["RAW_EXPORT", "RAW_EXPORT"]
    assert list(rows["Owner"]) == ["PERSONAL", "PERSONAL"]
    assert "PERSONAL" in stats["accounts_processed"]


def test_nordea_month_with_no_later_transaction_evidence_is_withheld(tmp_path):
    """
    Real bug found and fixed: the latest imported transaction landing
    mid-month (not exactly on that month's last calendar day) must NOT get
    a full end-of-month balance yet - there's no evidence the import
    covered the rest of that month.
    """
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"SourceAccount": "PERSONAL", "SourceBank": "NORDEA", "BookingDate": "2026-01-10", "Amount": -50, "Balance": 950},
        {"SourceAccount": "PERSONAL", "SourceBank": "NORDEA", "BookingDate": "2026-02-05", "Amount": 100, "Balance": 1030},
    ])
    settings = _settings_with_seeds("budgeting:\n  default_include: 'YES'\n", tmp_path)

    result, stats = build_monthly_balances(path, settings, as_of=date(2026, 3, 1))

    rows = result[result["SourceAccount"] == "PERSONAL"]
    assert list(rows["MonthEnd"]) == ["2026-01-31"]
    assert list(rows["Balance"]) == [950.0]


def test_nordea_account_with_no_balance_data_is_flagged_not_silently_dropped(tmp_path):
    """
    Real bug found and fixed: an account correctly identified as a raw-
    export bank (Nordea) but with zero usable Balance data must be
    explicitly flagged, not silently disappear from every stats bucket.
    """
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"SourceAccount": "PERSONAL", "SourceBank": "NORDEA", "BookingDate": "2026-01-10", "Amount": -50},
    ])
    settings = _settings_with_seeds("budgeting:\n  default_include: 'YES'\n", tmp_path)

    result, stats = build_monthly_balances(path, settings, as_of=date(2026, 3, 1))

    assert len(result[result["SourceAccount"] == "PERSONAL"]) == 0
    assert stats["accounts_processed"] == []
    assert stats["accounts_no_seed_configured"] == []
    assert stats["accounts_raw_export_no_balance_data"] == ["PERSONAL"]


def test_reconstructed_account_sums_from_seed_forward(tmp_path):
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-01-15", "Amount": -30},
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-02-10", "Amount": 200},
        # Proves February is "closed" (import continued past Feb 28) - see
        # test_reconstructed_month_with_no_later_transaction_evidence_is_withheld
        # for the case where this is missing.
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-03-01", "Amount": 1},
    ], owners={"HOUSEHOLD": "HOUSEHOLD"})
    settings = _settings_with_seeds(
        """
budgeting:
  account_balance_seeds:
    HOUSEHOLD:
      "2026-01-15": 1000.0
""",
        tmp_path,
    )

    result, stats = build_monthly_balances(path, settings, as_of=date(2026, 3, 5))

    rows = result[result["SourceAccount"] == "HOUSEHOLD"]
    # Seed date 01-15 is end-of-day, so its own -30 is already baked in.
    assert list(rows["MonthEnd"]) == ["2026-01-31", "2026-02-28"]
    assert list(rows["Balance"]) == [1000.0, 1200.0]
    assert list(rows["BalanceSource"]) == ["RECONSTRUCTED", "RECONSTRUCTED"]
    assert list(rows["Owner"]) == ["HOUSEHOLD", "HOUSEHOLD"]
    assert "HOUSEHOLD" in stats["accounts_processed"]


def test_reconstructed_month_with_no_later_transaction_evidence_is_withheld(tmp_path):
    """
    Real bug found and fixed: the seed forward-projection must not produce a
    row for a month whose end we have no evidence of (no later transaction
    proving the import continued past it) - same rule as the raw-export
    path, see test_nordea_month_with_no_later_transaction_evidence_is_withheld.
    """
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-01-15", "Amount": -30},
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-02-10", "Amount": 200},
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
    assert list(rows["MonthEnd"]) == ["2026-01-31"]
    assert list(rows["Balance"]) == [1000.0]


def test_reconstructed_account_projects_backward_before_the_seed(tmp_path):
    """
    A seed doesn't have to be the account's earliest known point - real
    transaction history before it (e.g. the account was imported well
    before the seed was ever entered) makes those earlier months computable
    too, via the same deterministic subtraction run in reverse.
    """
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2025-11-10", "Amount": -40},
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2025-12-20", "Amount": 60},
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-01-15", "Amount": -10},
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-02-10", "Amount": 5},
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
    by_month = dict(zip(rows["MonthEnd"], rows["Balance"]))
    # Dec: seed (1000) minus everything after Dec 31 through the seed date
    # (-10 on Jan 15) = 1010.
    assert by_month["2025-12-31"] == 1010.0
    # Nov: seed minus everything after Nov 30 through the seed date
    # (60 on Dec 20, -10 on Jan 15) = 1000 - 50 = 950.
    assert by_month["2025-11-30"] == 950.0
    assert by_month["2026-01-31"] == 1000.0
    # Feb's own last transaction (02-10) doesn't prove the month is closed.
    assert "2026-02-28" not in by_month


def test_reconstructed_backward_reconstruction_stops_at_first_gap(tmp_path):
    """
    Real bug found and fixed via a real dry-run: backward reconstruction
    across an unimported gap (a calendar month with zero real transactions)
    silently drifts more wrong the further back it goes - every real
    transaction missing inside the gap never gets added/subtracted back.
    Must stop at the first such month rather than assume the whole span
    back to the earliest transaction is gapless.
    """
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2025-09-05", "Amount": -20},
        # 2025-10 is a real gap: zero imported transactions that month.
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2025-11-10", "Amount": -40},
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2025-12-20", "Amount": 60},
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-01-15", "Amount": -10},
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

    result, stats = build_monthly_balances(path, settings, as_of=date(2026, 2, 1))

    rows = result[result["SourceAccount"] == "HOUSEHOLD"]
    by_month = set(rows["MonthEnd"])
    # Nov and Dec connect to the seed with no gap in between - reconstructed.
    assert "2025-12-31" in by_month
    assert "2025-11-30" in by_month
    # October is the gap itself; September is on the far side of it - unsafe either way.
    assert "2025-10-31" not in by_month
    assert "2025-09-30" not in by_month


def test_reconstructed_forward_reconstruction_also_stops_at_first_gap(tmp_path):
    """
    The same gap guard applies going forward from the seed, not just
    backward - a month past an unimported gap is just as untrustworthy
    either direction.
    """
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-01-15", "Amount": -10},
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-02-05", "Amount": -20},
        # 2026-03 is a real gap: zero imported transactions that month.
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-04-10", "Amount": 40},
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-05-05", "Amount": 5},
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

    result, stats = build_monthly_balances(path, settings, as_of=date(2026, 6, 1))

    rows = result[result["SourceAccount"] == "HOUSEHOLD"]
    by_month = set(rows["MonthEnd"])
    assert "2026-01-31" in by_month
    assert "2026-02-28" in by_month
    assert "2026-03-31" not in by_month  # the gap itself
    assert "2026-04-30" not in by_month  # beyond the gap, unsafe


def test_account_with_no_seed_is_flagged_not_silently_skipped(tmp_path):
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"SourceAccount": "SPANKKI", "SourceBank": "SPANKKI", "BookingDate": "2026-01-15", "Amount": -30},
    ])
    settings = _settings_with_seeds("budgeting:\n  default_include: 'YES'\n", tmp_path)

    result, stats = build_monthly_balances(path, settings, as_of=date(2026, 3, 1))

    assert len(result[result["SourceAccount"] == "SPANKKI"]) == 0
    assert stats["accounts_no_seed_configured"] == ["SPANKKI"]
    assert stats["accounts_processed"] == []


def test_cash_source_account_is_excluded(tmp_path):
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"SourceAccount": "CASH", "SourceBank": "CASH", "BookingDate": "2026-01-15", "Amount": -10},
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
    _write_raw(path, [
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-02-01", "Amount": -10},
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-03-05", "Amount": 5},
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
    # Mar 31 is withheld: the latest real transaction (03-05) doesn't prove
    # the import covered the rest of March.
    assert "2026-03-31" not in by_month


def test_investment_dividend_source_bank_never_counts_toward_balance(tmp_path):
    """
    Synthetic dividend-income rows (Nordnet/EVLI, injected by
    investment_dividends.py) never hit a real bank account - the cash lands
    in the broker's own cash balance instead. Must not count as a real
    balance-affecting event.
    """
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-01-15", "Amount": -30},
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "INVESTMENT_DIVIDEND", "BookingDate": "2026-02-01", "Amount": 500},
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-02-10", "Amount": 200},
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-03-01", "Amount": 1},
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

    result, stats = build_monthly_balances(path, settings, as_of=date(2026, 3, 5))

    rows = result[result["SourceAccount"] == "HOUSEHOLD"]
    by_month = dict(zip(rows["MonthEnd"], rows["Balance"]))
    # Feb: seed (1000) + real -30 (Jan 15, already baked into seed) is
    # already excluded; only the real +200 counts, NOT the +500 dividend row.
    assert by_month["2026-02-28"] == 1200.0


def test_katevaraus_pending_hold_never_counts_toward_balance(tmp_path):
    """
    Real bug found and fixed: a credit card's pending card-authorisation
    hold (katevaraus) isn't a real settled transaction. In real data it
    has a blank BookingDate (only ValueDate is populated), which already
    excludes it from the date-filtered group as a side effect - but this
    test gives it a real BookingDate anyway, proving the exclusion is
    explicit by transaction type and doesn't just accidentally rely on a
    blank date that a future export format could stop guaranteeing.
    """
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-01-15", "Amount": -30},
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-02-01", "TransactionTypeRaw": "Katevaraus", "Amount": -500},
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-02-10", "Amount": 200},
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-03-01", "Amount": 1},
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

    result, stats = build_monthly_balances(path, settings, as_of=date(2026, 3, 5))

    rows = result[result["SourceAccount"] == "HOUSEHOLD"]
    by_month = dict(zip(rows["MonthEnd"], rows["Balance"]))
    # Only the real +200 counts, NOT the -500 pending hold.
    assert by_month["2026-02-28"] == 1200.0
