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
        {"Broker": "OP", "Portfolio": "OP", "PortfolioOwner": "PERSON_A", "PortfolioType": "AOT", "NormalizedInstrument": "TELIA COMPANY AB", "TransactionType": "DIVIDEND", "TradeDate": "2017-10-30", "CashAmount": 61.62},
        {"Broker": "OP", "Portfolio": "OP", "PortfolioOwner": "PERSON_A", "PortfolioType": "AOT", "NormalizedInstrument": "TELIA COMPANY AB", "TransactionType": "TAX", "TradeDate": "2017-10-30", "CashAmount": -9.24},
    ])

    result, stats = build_dividend_history(path)

    assert len(result) == 1
    row = result.iloc[0]
    assert row["GrossDividendEUR"] == 61.62
    assert row["TaxWithheldEUR"] == 9.24
    assert row["NetDividendEUR"] == 52.38
    assert row["Year"] == 2017
    assert row["Month"] == 10
    assert row["PortfolioOwner"] == "PERSON_A"
    assert row["PortfolioType"] == "AOT"


def test_dividend_with_no_matching_tax_row_still_reports_correctly(tmp_path):
    """Real case: EVLI dividends have no withholding row at all."""
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"Broker": "EVLI", "Portfolio": "EVLI", "PortfolioOwner": "PERSON_A", "PortfolioType": "", "NormalizedInstrument": "NOKIA", "TransactionType": "DIVIDEND", "TradeDate": "2026-05-07", "CashAmount": 17.88},
    ])

    result, stats = build_dividend_history(path)

    assert len(result) == 1
    row = result.iloc[0]
    assert row["GrossDividendEUR"] == 17.88
    assert row["TaxWithheldEUR"] == 0.0
    assert row["NetDividendEUR"] == 17.88


def test_ennakkopidatys_is_treated_the_same_as_tax(tmp_path):
    """Real case: Nordnet uses ENNAKKOPIDÄTYS, OP uses TAX - same meaning."""
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"Broker": "NORDNET", "Portfolio": "1", "PortfolioOwner": "PERSON_A", "PortfolioType": "AOT", "NormalizedInstrument": "FORTUM", "TransactionType": "DIVIDEND", "TradeDate": "2022-03-29", "CashAmount": 114.0},
        {"Broker": "NORDNET", "Portfolio": "1", "PortfolioOwner": "PERSON_A", "PortfolioType": "AOT", "NormalizedInstrument": "FORTUM", "TransactionType": "ENNAKKOPIDÄTYS", "TradeDate": "2022-03-29", "CashAmount": -29.07},
    ])

    result, stats = build_dividend_history(path)

    row = result.iloc[0]
    assert row["TaxWithheldEUR"] == 29.07
    assert row["NetDividendEUR"] == 84.93


def test_different_dates_produce_separate_events(tmp_path):
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"Broker": "OP", "Portfolio": "OP", "PortfolioOwner": "", "PortfolioType": "", "NormalizedInstrument": "TELIA COMPANY AB", "TransactionType": "DIVIDEND", "TradeDate": "2017-10-30", "CashAmount": 61.62},
        {"Broker": "OP", "Portfolio": "OP", "PortfolioOwner": "", "PortfolioType": "", "NormalizedInstrument": "TELIA COMPANY AB", "TransactionType": "DIVIDEND", "TradeDate": "2018-04-18", "CashAmount": 65.88},
    ])

    result, stats = build_dividend_history(path)

    assert len(result) == 2
    assert stats["events"] == 2
    assert stats["total_gross_eur"] == 127.5


def test_non_dividend_rows_are_excluded(tmp_path):
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"Broker": "OP", "Portfolio": "OP", "PortfolioOwner": "", "PortfolioType": "", "NormalizedInstrument": "TELIA COMPANY AB", "TransactionType": "BUY", "TradeDate": "2017-10-30", "CashAmount": -100},
    ])

    result, stats = build_dividend_history(path)

    assert len(result) == 0
    assert stats["events"] == 0
