from openpyxl import Workbook

from finance_parser.budgeting.parsers import investment_dividends


DIVIDEND_HISTORY_HEADERS = [
    "TradeDate", "Year", "Month", "Broker", "Portfolio", "PortfolioOwner",
    "PortfolioType", "NormalizedInstrument", "GrossDividendEUR",
    "TaxWithheldEUR", "NetDividendEUR",
]


def _write_dividend_history(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = investment_dividends.DIVIDEND_HISTORY_SHEET
    ws.append(DIVIDEND_HISTORY_HEADERS)
    for row in rows:
        ws.append([row.get(h, "") for h in DIVIDEND_HISTORY_HEADERS])
    wb.save(path)


def test_can_parse_real_dividend_history_shape(tmp_path):
    path = tmp_path / "DividendHistory.xlsx"
    _write_dividend_history(path, [])

    ok, reason = investment_dividends.can_parse(path)

    assert ok, reason


def test_only_synthesizes_rows_for_brokers_with_no_bank_match(tmp_path):
    """OP-held instruments settle same-day into a real bank transaction
    (handled by enrich_dividend_income.py's cross-reference) - only
    Nordnet/EVLI dividends (no matching bank transaction exists) get a
    synthetic row here."""
    path = tmp_path / "DividendHistory.xlsx"
    _write_dividend_history(path, [
        {"TradeDate": "2021-06-15", "Broker": "OP", "Portfolio": "OP", "PortfolioOwner": "PERSON_A", "NormalizedInstrument": "SAMPO A", "NetDividendEUR": 100.0},
        {"TradeDate": "2021-06-15", "Broker": "NORDNET", "Portfolio": "1", "PortfolioOwner": "PERSON_A", "NormalizedInstrument": "FORTUM", "NetDividendEUR": 50.0},
        {"TradeDate": "2021-06-15", "Broker": "EVLI", "Portfolio": "EVLI", "PortfolioOwner": "PERSON_A", "NormalizedInstrument": "NOKIA", "NetDividendEUR": 30.0},
    ])

    df = investment_dividends.parse_file(path, "2026-01-01 00:00:00")

    assert len(df) == 2
    assert set(df["Amount"]) == {50.0, 30.0}


def test_source_account_carries_portfolio_owner_not_owner_directly(tmp_path):
    """The parser must never set Owner directly (see CLAUDE.md) - it carries
    PortfolioOwner into SourceAccount so the existing SourceAccount->Owner
    OwnershipRules resolve it the same way they do for every real account."""
    path = tmp_path / "DividendHistory.xlsx"
    _write_dividend_history(path, [
        {"TradeDate": "2021-06-15", "Broker": "NORDNET", "Portfolio": "1", "PortfolioOwner": "PERSON_A", "NormalizedInstrument": "FORTUM", "NetDividendEUR": 50.0},
    ])

    df = investment_dividends.parse_file(path, "2026-01-01 00:00:00")

    row = df.iloc[0]
    assert row["SourceAccount"] == "PERSON_A"
    assert "Owner" not in df.columns


def test_receiver_text_avoids_colliding_with_real_merchant_name_rules(tmp_path):
    """Real bug found and fixed: a bare instrument name (e.g. a company that
    is ALSO a real household utility/telecom bill payee) can exact-match an
    unrelated existing CategoryRules rule and get miscategorized as an
    expense instead of income. The receiver text must never be just the
    bare instrument name."""
    path = tmp_path / "DividendHistory.xlsx"
    _write_dividend_history(path, [
        {"TradeDate": "2021-06-15", "Broker": "NORDNET", "Portfolio": "1", "PortfolioOwner": "PERSON_A", "NormalizedInstrument": "FORTUM", "NetDividendEUR": 50.0},
    ])

    df = investment_dividends.parse_file(path, "2026-01-01 00:00:00")

    assert df.iloc[0]["RawReceiver"] != "FORTUM"
    assert "FORTUM" in df.iloc[0]["RawReceiver"]


def test_zero_or_negative_dividend_amounts_are_excluded(tmp_path):
    path = tmp_path / "DividendHistory.xlsx"
    _write_dividend_history(path, [
        {"TradeDate": "2021-06-15", "Broker": "NORDNET", "Portfolio": "1", "PortfolioOwner": "PERSON_A", "NormalizedInstrument": "FORTUM", "NetDividendEUR": 0.0},
        {"TradeDate": "2021-06-15", "Broker": "NORDNET", "Portfolio": "1", "PortfolioOwner": "PERSON_A", "NormalizedInstrument": "NESTE", "NetDividendEUR": 50.0},
    ])

    df = investment_dividends.parse_file(path, "2026-01-01 00:00:00")

    assert len(df) == 1
    assert df.iloc[0]["Amount"] == 50.0


def test_raw_id_is_stable_across_reparse(tmp_path):
    """Same real dividend event parsed twice must produce the same RawID,
    so re-running the pipeline doesn't duplicate rows."""
    path = tmp_path / "DividendHistory.xlsx"
    _write_dividend_history(path, [
        {"TradeDate": "2021-06-15", "Broker": "NORDNET", "Portfolio": "1", "PortfolioOwner": "PERSON_A", "NormalizedInstrument": "FORTUM", "NetDividendEUR": 50.0},
    ])

    df1 = investment_dividends.parse_file(path, "2026-01-01 00:00:00")
    df2 = investment_dividends.parse_file(path, "2026-01-02 00:00:00")

    assert df1.iloc[0]["RawID"] == df2.iloc[0]["RawID"]
