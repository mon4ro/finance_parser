from datetime import date, timedelta

import pandas as pd
import pytest
from openpyxl import Workbook

from finance_parser.investments.fetch_fx_rates import (
    FX_RATES_COLUMNS,
    fetch_all,
    fx_symbol,
    owned_non_eur_currencies,
)
from finance_parser.investments.fetch_instrument_prices import fetch_yahoo_daily_closes


TRANSACTIONS_HEADERS = [
    "InvestmentTransactionID", "InvestmentRawID", "Broker", "Portfolio",
    "PortfolioOwner", "PortfolioType", "TransactionType", "TradeDate",
    "SettlementDate", "InstrumentName", "NormalizedInstrument", "ISIN",
    "InstrumentType", "AssetClass", "InstrumentCurrency", "Quantity",
    "UnitPrice", "CashAmount", "CashCurrency", "AcquisitionValue",
    "AcquisitionCurrency", "TotalQuantity", "ExchangeRate", "Description",
    "BrokerageFee", "BrokerageFeeCurrency", "ExportDate", "ImportedAt",
    "SourceFile", "Comments",
]


def _write_transactions(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "InvestmentTransactions"
    ws.append(TRANSACTIONS_HEADERS)
    for row in rows:
        ws.append([row.get(h, "") for h in TRANSACTIONS_HEADERS])
    wb.save(path)


def test_fx_symbol_appends_eur_and_x_suffix():
    assert fx_symbol("NOK") == "NOKEUR=X"
    assert fx_symbol("USD") == "USDEUR=X"


def test_owned_non_eur_currencies_excludes_eur_and_blank(tmp_path):
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"InstrumentCurrency": "EUR", "TradeDate": "2021-01-01"},
        {"InstrumentCurrency": "", "TradeDate": "2021-01-01"},
        {"InstrumentCurrency": "NOK", "TradeDate": "2021-01-01"},
        {"InstrumentCurrency": "SEK", "TradeDate": "2021-06-01"},
    ])

    result = owned_non_eur_currencies(path)

    assert set(result["Currency"]) == {"NOK", "SEK"}


def test_owned_non_eur_currencies_returns_earliest_trade_date(tmp_path):
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"InstrumentCurrency": "USD", "TradeDate": "2021-06-01"},
        {"InstrumentCurrency": "USD", "TradeDate": "2020-01-15"},
    ])

    result = owned_non_eur_currencies(path)
    by_currency = result.set_index("Currency")["FirstTradeDate"]

    assert by_currency["USD"] == pd.Timestamp("2020-01-15")


def test_fetch_all_incremental_skips_already_covered_dates(tmp_path):
    """
    If FXRates already has data up to today, a same-day rerun should not
    attempt to refetch that range.
    """
    transactions_path = tmp_path / "ParsedInvestments.xlsx"
    fx_rates_path = tmp_path / "FXRates.xlsx"

    _write_transactions(transactions_path, [
        {"InstrumentCurrency": "NOK", "TradeDate": "2021-01-01"},
    ])

    wb = Workbook()
    ws = wb.active
    ws.title = "FXRates"
    ws.append(FX_RATES_COLUMNS)
    today = date.today().isoformat()
    ws.append(["NOK", today, 0.1234, "YAHOO", "2026-01-01 00:00:00"])
    wb.save(fx_rates_path)

    combined, stats = fetch_all(transactions_path, fx_rates_path, verbose=False)

    assert stats["currencies_up_to_date"] == 1
    assert stats["currencies_fetched"] == 0
    assert len(combined) == 1
    assert combined.iloc[0]["Date"] == today


@pytest.mark.parametrize("symbol", ["NOKEUR=X"])
def test_fetch_yahoo_daily_closes_fx_real_network_call(symbol):
    """
    Live integration check against the real Yahoo FX chart API - same
    endpoint and function already used for equity prices, just a currency
    pair symbol instead of a stock ticker. Skips gracefully if the network is
    unavailable.
    """
    end = date.today() - timedelta(days=1)
    start = end - timedelta(days=10)

    try:
        rows, currency = fetch_yahoo_daily_closes(symbol, start, end)
    except Exception as exc:
        pytest.skip(f"Live FX rate API unreachable, skipping: {exc}")

    assert currency == "EUR"
    assert len(rows) > 0
    for day, rate in rows:
        # FX trades ~24h/day unlike equities with fixed exchange hours, so a
        # daily bar can land one UTC calendar day outside the requested
        # range on either side (observed in practice: a bar for the "end"
        # session date returned dated end+1, and separately a bar dated
        # start-1) - tolerate that specific one-day slip on both ends, not a
        # wider window.
        assert start - timedelta(days=1) <= day <= end + timedelta(days=1)
        # A NOK/EUR rate should plausibly be a small positive fraction, not
        # e.g. accidentally inverted (~10-11) or zero.
        assert 0 < rate < 1
