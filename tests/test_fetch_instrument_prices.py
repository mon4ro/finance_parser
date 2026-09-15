from datetime import date, timedelta

import pandas as pd
import pytest
from openpyxl import Workbook

from finance_parser.investments.fetch_instrument_prices import (
    PRICES_COLUMNS,
    build_fetch_plan,
    fetch_all,
    fetch_yahoo_daily_closes,
    load_existing_prices,
    load_instrument_master_price_config,
    owned_normalized_instruments,
)


INSTRUMENT_MASTER_HEADERS = [
    "Enabled", "MatchPriority", "Broker", "ISIN", "RawInstrumentName",
    "NormalizedInstrument", "InstrumentType", "AssetClass", "Currency",
    "Ticker", "Exchange", "PriceSource", "PriceSymbol", "Notes",
]

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


def _write_instrument_master(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "InstrumentMaster"
    ws.append(INSTRUMENT_MASTER_HEADERS)
    for row in rows:
        ws.append([row.get(h, "") for h in INSTRUMENT_MASTER_HEADERS])
    wb.save(path)


def _write_transactions(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "InvestmentTransactions"
    ws.append(TRANSACTIONS_HEADERS)
    for row in rows:
        ws.append([row.get(h, "") for h in TRANSACTIONS_HEADERS])
    wb.save(path)


def test_load_instrument_master_price_config_filters_correctly(tmp_path):
    path = tmp_path / "InstrumentMaster.xlsx"
    _write_instrument_master(path, [
        {"Enabled": "YES", "NormalizedInstrument": "SAMPO A", "PriceSource": "YAHOO", "PriceSymbol": "SAMPO.HE", "Currency": "EUR"},
        {"Enabled": "YES", "NormalizedInstrument": "NO SYMBOL YET", "PriceSymbol": ""},
        {"Enabled": "NO", "NormalizedInstrument": "DISABLED ROW", "PriceSymbol": "XYZ"},
    ])

    df = load_instrument_master_price_config(path)

    assert list(df["NormalizedInstrument"]) == ["SAMPO A"]
    assert df.iloc[0]["PriceSymbol"] == "SAMPO.HE"


def test_owned_normalized_instruments_returns_earliest_trade_date(tmp_path):
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"NormalizedInstrument": "SAMPO A", "TradeDate": "2026-05-01"},
        {"NormalizedInstrument": "SAMPO A", "TradeDate": "2026-06-01"},
        {"NormalizedInstrument": "NOKIA", "TradeDate": "2020-01-15"},
    ])

    df = owned_normalized_instruments(path)
    by_name = df.set_index("NormalizedInstrument")["FirstTradeDate"]

    assert by_name["SAMPO A"] == pd.Timestamp("2026-05-01")
    assert by_name["NOKIA"] == pd.Timestamp("2020-01-15")


def test_build_fetch_plan_reports_missing_price_symbol(tmp_path):
    master_path = tmp_path / "InstrumentMaster.xlsx"
    transactions_path = tmp_path / "ParsedInvestments.xlsx"

    _write_instrument_master(master_path, [
        {"Enabled": "YES", "NormalizedInstrument": "SAMPO A", "PriceSource": "YAHOO", "PriceSymbol": "SAMPO.HE", "Currency": "EUR"},
    ])
    _write_transactions(transactions_path, [
        {"NormalizedInstrument": "SAMPO A", "TradeDate": "2026-05-01"},
        {"NormalizedInstrument": "NOKIA", "TradeDate": "2020-01-15"},
    ])

    plan, missing = build_fetch_plan(master_path, transactions_path)

    assert list(plan["NormalizedInstrument"]) == ["SAMPO A"]
    assert missing == ["NOKIA"]


def test_fetch_all_incremental_skips_already_covered_dates(tmp_path):
    """
    If InstrumentPrices already has data up to yesterday, a same-day rerun
    should not attempt to refetch that range - only anything from today on.
    """
    master_path = tmp_path / "InstrumentMaster.xlsx"
    transactions_path = tmp_path / "ParsedInvestments.xlsx"
    prices_path = tmp_path / "InstrumentPrices.xlsx"

    _write_instrument_master(master_path, [
        {"Enabled": "YES", "NormalizedInstrument": "SAMPO A", "PriceSource": "YAHOO", "PriceSymbol": "SAMPO.HE", "Currency": "EUR"},
    ])
    _write_transactions(transactions_path, [
        {"NormalizedInstrument": "SAMPO A", "TradeDate": "2026-05-01"},
    ])

    # Pre-seed InstrumentPrices as already up to date through today.
    wb = Workbook()
    ws = wb.active
    ws.title = "InstrumentPrices"
    ws.append(PRICES_COLUMNS)
    today = date.today().isoformat()
    ws.append(["SAMPO A", "", "SAMPO.HE", today, 9.5, "EUR", "YAHOO", "2026-01-01 00:00:00"])
    wb.save(prices_path)

    combined, stats = fetch_all(master_path, transactions_path, prices_path, verbose=False)

    assert stats["instruments_up_to_date"] == 1
    assert stats["instruments_fetched"] == 0
    # Existing row must survive untouched.
    assert len(combined) == 1
    assert combined.iloc[0]["Date"] == today


@pytest.mark.parametrize("symbol", ["SAMPO.HE"])
def test_fetch_yahoo_daily_closes_real_network_call(symbol):
    """
    Live integration check against the real (unofficial) Yahoo Finance API.
    Skips gracefully if the network is unavailable rather than failing the
    whole suite - this is deliberately testing real external data shape, per
    this project's preference for real data over mocks where practical.
    """
    end = date.today() - timedelta(days=1)
    start = end - timedelta(days=10)

    try:
        rows, currency = fetch_yahoo_daily_closes(symbol, start, end)
    except Exception as exc:  # network unavailable, API down, etc.
        pytest.skip(f"Live price API unreachable, skipping: {exc}")

    assert currency == "EUR"
    assert len(rows) > 0
    for day, close in rows:
        assert start <= day <= end
        assert close > 0
