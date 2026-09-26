from openpyxl import Workbook

from finance_parser.budgeting.account_balance_seed import (
    accounts_needing_seed,
    collect_account_balance_seed,
    latest_imported_date,
    load_raw_transactions,
    owner_lookup,
)


RAW_COLUMNS = [
    "RawID", "SourceAccount", "SourceBank", "ExportDate", "ImportedAt",
    "BookingDate", "ValueDate", "Amount", "TransactionTypeRaw", "Description",
    "RawReceiver", "ReceiverAccount", "ReceiverBankBIC", "Reference", "Message",
    "ArchiveID", "Balance", "CurrencyAmount", "Currency", "Rate",
    "MerchantArea", "MerchantCategory", "SourceFile",
]


def _write_raw(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "RawTransactions"
    ws.append(RAW_COLUMNS)
    for row in rows:
        ws.append([row.get(h, "") for h in RAW_COLUMNS])
    wb.save(path)


def _feed(monkeypatch, answers):
    queue = list(answers)

    def fake_input(prompt=""):
        if not queue:
            raise AssertionError(f"Ran out of scripted answers at prompt: {prompt!r}")
        return queue.pop(0)

    monkeypatch.setattr("builtins.input", fake_input)


def test_accounts_needing_seed_excludes_cash_and_nordea_only_accounts(tmp_path):
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"RawID": "R-1", "SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-01-01", "Amount": -10},
        {"RawID": "R-2", "SourceAccount": "PERSONAL", "SourceBank": "NORDEA", "BookingDate": "2026-01-01", "Amount": -10},
        {"RawID": "R-3", "SourceAccount": "CASH", "SourceBank": "CASH", "BookingDate": "2026-01-01", "Amount": -10},
        {"RawID": "R-4", "SourceAccount": "SPANKKI", "SourceBank": "SPANKKI", "BookingDate": "2026-01-01", "Amount": -10},
    ])

    accounts = accounts_needing_seed(path)

    assert accounts == ["HOUSEHOLD", "SPANKKI"]


def test_accounts_needing_seed_keeps_account_mixing_nordea_with_another_bank(tmp_path):
    """
    An account isn't excluded just because SOME of its transactions are
    Nordea - only if ALL of them are - a partial Nordea presence still
    needs a manual seed since Nordea's own real balance field doesn't cover
    the whole account.
    """
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"RawID": "R-1", "SourceAccount": "MIXED", "SourceBank": "OP", "BookingDate": "2026-01-01", "Amount": -10},
        {"RawID": "R-2", "SourceAccount": "MIXED", "SourceBank": "NORDEA", "BookingDate": "2026-01-02", "Amount": -10},
    ])

    assert accounts_needing_seed(path) == ["MIXED"]


def test_accounts_needing_seed_merges_tili_prefixed_rows_into_the_real_account(tmp_path):
    """
    Real bug: rows imported before the "Tili" filename-prefix fix in
    infer_source_account_from_filename() have "TILI <name>" permanently
    baked into their own SourceAccount value - the code fix only changes
    NEW imports, so a real account ended up listed twice (once as
    "CHILD", once as "TILI CHILD") purely depending on when each row was
    imported. load_raw_transactions() must merge them via
    sheet_to_dataframe()'s migration before accounts_needing_seed() groups
    by SourceAccount, so only one clean entry shows up.
    """
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"RawID": "R-1", "SourceAccount": "CHILD", "SourceBank": "OP", "BookingDate": "2026-01-01", "Amount": -10},
        {"RawID": "R-2", "SourceAccount": "TILI CHILD", "SourceBank": "OP", "BookingDate": "2025-06-01", "Amount": -5},
    ])

    accounts = accounts_needing_seed(path)

    assert accounts == ["CHILD"]


def test_accounts_needing_seed_empty_workbook_returns_empty(tmp_path):
    path = tmp_path / "does_not_exist.xlsx"
    assert accounts_needing_seed(path) == []


def test_latest_imported_date_returns_max_date_for_account(tmp_path):
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"RawID": "R-1", "SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-01-01", "Amount": -10},
        {"RawID": "R-2", "SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-03-15", "Amount": -20},
        {"RawID": "R-3", "SourceAccount": "OTHER", "SourceBank": "OP", "BookingDate": "2026-06-01", "Amount": -30},
    ])

    assert latest_imported_date("HOUSEHOLD", path).isoformat() == "2026-03-15"


def test_latest_imported_date_no_transactions_returns_none(tmp_path):
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [])
    assert latest_imported_date("HOUSEHOLD", path) is None


def test_collect_account_balance_seed_writes_to_overrides(tmp_path, monkeypatch):
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"RawID": "R-1", "SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-03-15", "Amount": -10},
    ])
    _feed(monkeypatch, ["", "1234.56"])  # accept default date, then balance
    overrides = {}

    added = collect_account_balance_seed(overrides, "HOUSEHOLD", budgeting_workbook=path)

    assert added is True
    assert overrides["budgeting"]["account_balance_seeds"]["HOUSEHOLD"]["2026-03-15"] == 1234.56


def test_collect_account_balance_seed_future_date_requires_explicit_confirmation(tmp_path, monkeypatch):
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"RawID": "R-1", "SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-03-15", "Amount": -10},
    ])
    _feed(monkeypatch, ["2026-04-01", "n"])  # later date, decline the warning
    overrides = {}

    added = collect_account_balance_seed(overrides, "HOUSEHOLD", budgeting_workbook=path)

    assert added is False
    assert overrides == {}


def test_collect_account_balance_seed_future_date_accepted_when_confirmed(tmp_path, monkeypatch):
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"RawID": "R-1", "SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-03-15", "Amount": -10},
    ])
    _feed(monkeypatch, ["2026-04-01", "y", "500.00"])
    overrides = {}

    added = collect_account_balance_seed(overrides, "HOUSEHOLD", budgeting_workbook=path)

    assert added is True
    assert overrides["budgeting"]["account_balance_seeds"]["HOUSEHOLD"]["2026-04-01"] == 500.00


def test_collect_account_balance_seed_blank_date_skips(tmp_path, monkeypatch):
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [])
    _feed(monkeypatch, [""])
    overrides = {}

    added = collect_account_balance_seed(overrides, "HOUSEHOLD", budgeting_workbook=path)

    assert added is False


def test_collect_account_balance_seed_non_numeric_balance_skips(tmp_path, monkeypatch):
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"RawID": "R-1", "SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-03-15", "Amount": -10},
    ])
    _feed(monkeypatch, ["", "not a number"])
    overrides = {}

    added = collect_account_balance_seed(overrides, "HOUSEHOLD", budgeting_workbook=path)

    assert added is False
    assert overrides == {}


def test_load_raw_transactions_reads_balance_natively(tmp_path):
    """
    RawTransactions carries Balance directly (Nordea's own "Saldo" field,
    at the raw/parser layer) - no fallback merge needed, unlike the old
    UnifiedTransactions-based approach this replaced.
    """
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"RawID": "R-1", "SourceAccount": "PERSONAL", "SourceBank": "NORDEA", "BookingDate": "2026-03-15", "Amount": -10, "Balance": 456.78},
    ])

    df = load_raw_transactions(path)

    assert float(df.iloc[0]["Balance"]) == 456.78


def test_load_raw_transactions_uses_value_date_not_booking_date(tmp_path):
    """
    Deliberately reads ValueDate, not BookingDate - flipped 2026-09-26 after
    a real NORWEGIAN transaction's ValueDate (2026-06-30) landed a real
    month before its BookingDate (2026-07-02), silently shifting that
    month's real spend into the wrong reconstructed-balance window. A card
    hold reduces real spending power on its value/authorisation date, well
    before the bank formally posts it - BookingDate is provably wrong for
    this purpose. See load_raw_transactions()'s own docstring.
    """
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"RawID": "R-1", "SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-03-18", "ValueDate": "2026-03-15", "Amount": -10},
    ])

    df = load_raw_transactions(path)

    assert df.iloc[0]["_Date"].date().isoformat() == "2026-03-15"


def test_load_raw_transactions_falls_back_to_booking_date_when_lag_too_large(tmp_path):
    """
    Real bug found while building the ValueDate fix above: one merchant's
    refund exports set ValueDate to the *original* order's date, not the
    refund's own - a gap of several weeks, wildly outside any real
    settlement lag (confirmed at most 4 days across every real
    NORWEGIAN/SPANKKI transaction checked). Trusting ValueDate there moved a
    real credit clean out of its correct month. A gap this large means
    ValueDate is answering a different question, not lagging normally.
    """
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"RawID": "R-1", "SourceAccount": "HOUSEHOLD", "SourceBank": "NORWEGIAN", "BookingDate": "2026-02-18", "ValueDate": "2025-12-02", "Amount": 45, "TransactionTypeRaw": "CreditVoucher"},
    ])

    df = load_raw_transactions(path)

    assert df.iloc[0]["_Date"].date().isoformat() == "2026-02-18"


def test_load_raw_transactions_falls_back_to_booking_date_when_value_date_blank(tmp_path):
    """
    A pending card-authorisation hold has no BookingDate yet (see
    KATEVARAUS_TRANSACTION_TYPES in common.py) - the reverse gap, a blank
    ValueDate with a real BookingDate, isn't a real case in any bank export
    seen so far, but falling back keeps this function safe either way
    rather than silently producing a blank/unreconstructable date.
    """
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"RawID": "R-1", "SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-03-15", "ValueDate": "", "Amount": -10},
    ])

    df = load_raw_transactions(path)

    assert df.iloc[0]["_Date"].date().isoformat() == "2026-03-15"


def test_owner_lookup_reads_from_unified_transactions_grouped_by_account(tmp_path):
    """
    Owner is a display-only label resolved by OwnershipRules at the
    categoriser layer, which only ever runs against UnifiedTransactions -
    RawTransactions has no Owner field. This side lookup is deliberately
    isolated from the balance/date/amount math above and grouped by
    SourceAccount, never by Owner, so two different accounts sharing an
    owner are never merged together.
    """
    path = tmp_path / "ParsedTransactions.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "UnifiedTransactions"
    ws.append(["SourceAccount", "Owner"])
    ws.append(["HOUSEHOLD", "HOUSEHOLD_OWNER"])
    ws.append(["OTHER", "OTHER_OWNER"])
    wb.save(path)

    lookup = owner_lookup(path)

    assert lookup == {"HOUSEHOLD": "HOUSEHOLD_OWNER", "OTHER": "OTHER_OWNER"}


def test_owner_lookup_missing_unified_sheet_returns_empty(tmp_path):
    path = tmp_path / "ParsedTransactions.xlsx"
    _write_raw(path, [
        {"RawID": "R-1", "SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "BookingDate": "2026-03-15", "Amount": -10},
    ])

    assert owner_lookup(path) == {}
