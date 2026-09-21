from openpyxl import Workbook

from finance_parser.investments.build_dividend_history import build_dividend_history


TRANSACTIONS_HEADERS = [
    "Broker", "Portfolio", "PortfolioOwner", "PortfolioType", "NormalizedInstrument",
    "TransactionType", "TradeDate", "CashAmount",
]


def _write_transactions(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "InvestmentTransactions"
    ws.append(TRANSACTIONS_HEADERS)
    for row in rows:
        ws.append([row.get(h, "") for h in TRANSACTIONS_HEADERS])
    wb.save(path)


def test_dividend_and_tax_on_the_same_date_are_grouped_into_one_event(tmp_path):
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"Broker": "OP", "Portfolio": "OP", "PortfolioOwner": "PERSON_A", "PortfolioType": "AOT", "NormalizedInstrument": "TELIA COMPANY AB", "TransactionType": "DIVIDEND", "TradeDate": "2021-01-15", "CashAmount": 20.0},
        {"Broker": "OP", "Portfolio": "OP", "PortfolioOwner": "PERSON_A", "PortfolioType": "AOT", "NormalizedInstrument": "TELIA COMPANY AB", "TransactionType": "TAX", "TradeDate": "2021-01-15", "CashAmount": -3.0},
    ])

    result, stats = build_dividend_history(path)

    assert len(result) == 1
    row = result.iloc[0]
    assert row["GrossDividendEUR"] == 20.0
    assert row["TaxWithheldEUR"] == 3.0
    assert row["NetDividendEUR"] == 17.0
    assert row["Year"] == 2021
    assert row["Month"] == 1
    assert row["PortfolioOwner"] == "PERSON_A"
    assert row["PortfolioType"] == "AOT"


def test_dividend_with_no_matching_tax_row_still_reports_correctly(tmp_path):
    """
    Real case: EVLI dividends (an employee share-purchase plan account)
    have no withholding row at all in the real data - GrossDividendEUR ends
    up equal to NetDividendEUR, which looks identical to "genuinely
    tax-exempt" even when real tax was withheld outside this data source
    (e.g. via payroll). TaxDataAvailable=False makes that gap visible
    instead of silently implying zero tax was withheld.
    """
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"Broker": "EVLI", "Portfolio": "EVLI", "PortfolioOwner": "PERSON_A", "PortfolioType": "", "NormalizedInstrument": "NOKIA", "TransactionType": "DIVIDEND", "TradeDate": "2021-02-01", "CashAmount": 15.0},
    ])

    result, stats = build_dividend_history(path)

    assert len(result) == 1
    row = result.iloc[0]
    assert row["GrossDividendEUR"] == 15.0
    assert row["TaxWithheldEUR"] == 0.0
    assert row["NetDividendEUR"] == 15.0
    assert row["TaxDataAvailable"] == False
    assert stats["brokers_without_tax_data"] == ["EVLI"]


def test_broker_with_tax_data_elsewhere_is_flagged_available(tmp_path):
    """
    TaxDataAvailable is a broker-level signal (does this broker's export
    format ever carry a withholding row at all), not a per-event one - a
    genuinely tax-exempt individual dividend from a broker that DOES report
    withholding elsewhere must not be mistaken for a data-availability gap.
    """
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"Broker": "OP", "Portfolio": "OP", "PortfolioOwner": "PERSON_A", "PortfolioType": "AOT", "NormalizedInstrument": "SAMPO A", "TransactionType": "DIVIDEND", "TradeDate": "2021-01-15", "CashAmount": 20.0},
        {"Broker": "OP", "Portfolio": "OP", "PortfolioOwner": "PERSON_A", "PortfolioType": "AOT", "NormalizedInstrument": "SAMPO A", "TransactionType": "TAX", "TradeDate": "2021-01-15", "CashAmount": -3.0},
        # Genuinely tax-exempt dividend, same broker, no tax row of its own.
        {"Broker": "OP", "Portfolio": "OP", "PortfolioOwner": "PERSON_A", "PortfolioType": "AOT", "NormalizedInstrument": "KESKO B", "TransactionType": "DIVIDEND", "TradeDate": "2021-02-15", "CashAmount": 10.0},
    ])

    result, stats = build_dividend_history(path)

    assert set(result["TaxDataAvailable"]) == {True}
    assert stats["brokers_without_tax_data"] == []


def test_ennakkopidatys_is_treated_the_same_as_tax(tmp_path):
    """Real case: Nordnet uses ENNAKKOPIDÄTYS, OP uses TAX - same meaning."""
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"Broker": "NORDNET", "Portfolio": "1", "PortfolioOwner": "PERSON_A", "PortfolioType": "AOT", "NormalizedInstrument": "FORTUM", "TransactionType": "DIVIDEND", "TradeDate": "2021-03-01", "CashAmount": 40.0},
        {"Broker": "NORDNET", "Portfolio": "1", "PortfolioOwner": "PERSON_A", "PortfolioType": "AOT", "NormalizedInstrument": "FORTUM", "TransactionType": "ENNAKKOPIDÄTYS", "TradeDate": "2021-03-01", "CashAmount": -6.0},
    ])

    result, stats = build_dividend_history(path)

    row = result.iloc[0]
    assert row["TaxWithheldEUR"] == 6.0
    assert row["NetDividendEUR"] == 34.0


def test_different_dates_produce_separate_events(tmp_path):
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"Broker": "OP", "Portfolio": "OP", "PortfolioOwner": "", "PortfolioType": "", "NormalizedInstrument": "TELIA COMPANY AB", "TransactionType": "DIVIDEND", "TradeDate": "2021-01-15", "CashAmount": 20.0},
        {"Broker": "OP", "Portfolio": "OP", "PortfolioOwner": "", "PortfolioType": "", "NormalizedInstrument": "TELIA COMPANY AB", "TransactionType": "DIVIDEND", "TradeDate": "2021-07-15", "CashAmount": 22.0},
    ])

    result, stats = build_dividend_history(path)

    assert len(result) == 2
    assert stats["events"] == 2
    assert stats["total_gross_eur"] == 42.0


def test_non_dividend_rows_are_excluded(tmp_path):
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"Broker": "OP", "Portfolio": "OP", "PortfolioOwner": "", "PortfolioType": "", "NormalizedInstrument": "TELIA COMPANY AB", "TransactionType": "BUY", "TradeDate": "2021-01-15", "CashAmount": -100},
    ])

    result, stats = build_dividend_history(path)

    assert len(result) == 0
    assert stats["events"] == 0


def test_no_budgeting_workbook_leaves_local_currency_columns_blank(tmp_path):
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"Broker": "OP", "Portfolio": "OP", "PortfolioOwner": "", "PortfolioType": "", "NormalizedInstrument": "TELIA COMPANY AB", "TransactionType": "DIVIDEND", "TradeDate": "2021-01-15", "CashAmount": 20.0},
    ])

    result, stats = build_dividend_history(path, budgeting_workbook=None)

    assert result.iloc[0]["LocalCurrency"] == ""
    assert stats["local_currency_events_matched"] == 0


def test_budgeting_workbook_enriches_matching_telia_row(tmp_path):
    investments_path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(investments_path, [
        {"Broker": "OP", "Portfolio": "OP", "PortfolioOwner": "", "PortfolioType": "", "NormalizedInstrument": "TELIA COMPANY AB", "TransactionType": "DIVIDEND", "TradeDate": "2021-01-15", "CashAmount": 20.0},
    ])

    budgeting_path = tmp_path / "ParsedTransactions.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "UnifiedTransactions"
    ws.append(["Date", "Amount", "Message"])
    ws.append([
        "2021-01-15", 17.0,
        "Viesti: OP Säilytys Oy TELIA COMPANY AB SE0000667925 Osinkotuotto "
        "Osinko 0,50 SEK/KplOmistettu määrä 100Kpl Tuoton määrä 50,00SEK"
        "Lähdevero SE15,0 % 7,50SEKVal.kurssi 11,000000",
    ])
    wb.save(budgeting_path)

    result, stats = build_dividend_history(investments_path, budgeting_workbook=budgeting_path)

    assert stats["local_currency_events_matched"] == 1
    row = result.iloc[0]
    assert row["LocalCurrency"] == "SEK"
    assert row["GrossDividendLocal"] == 50.0
    assert row["ExchangeRate"] == 11.0
