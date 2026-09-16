from datetime import date

import pytest
from openpyxl import Workbook

from finance_parser.investments.build_monthly_position_value import (
    build_monthly_values,
    last_completed_month_end,
)


POSITIONS_HEADERS = [
    "Broker", "Portfolio", "PortfolioOwner", "NormalizedInstrument", "Date", "TransactionType",
    "QuantityDelta", "CumulativeQuantity", "CashDelta", "CumulativeNetInvested",
]
PRICES_HEADERS = ["NormalizedInstrument", "ISIN", "PriceSymbol", "Date", "Close", "Currency", "PriceSource", "FetchedAt"]
FX_HEADERS = ["Currency", "Date", "Rate", "Source", "FetchedAt"]


def _write_sheet(path, sheet_name, headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(headers)
    for row in rows:
        ws.append([row.get(h, "") for h in headers])
    wb.save(path)


def _positions(path, rows):
    _write_sheet(path, "PortfolioPositions", POSITIONS_HEADERS, rows)


def _prices(path, rows):
    _write_sheet(path, "InstrumentPrices", PRICES_HEADERS, rows)


def _fx(path, rows):
    _write_sheet(path, "FXRates", FX_HEADERS, rows)


def test_last_completed_month_end():
    assert last_completed_month_end(date(2026, 9, 16)).strftime("%Y-%m-%d") == "2026-08-31"
    assert last_completed_month_end(date(2026, 1, 1)).strftime("%Y-%m-%d") == "2025-12-31"


def test_eur_instrument_market_value_and_gain(tmp_path):
    positions_path = tmp_path / "PortfolioPositions.xlsx"
    prices_path = tmp_path / "InstrumentPrices.xlsx"
    fx_path = tmp_path / "FXRates.xlsx"

    _positions(positions_path, [
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "SAMPO A", "Date": "2021-03-15", "CumulativeQuantity": 100, "CumulativeNetInvested": -1000},
    ])
    _prices(prices_path, [
        {"NormalizedInstrument": "SAMPO A", "Date": "2021-03-31", "Close": 12.0, "Currency": "EUR"},
    ])
    _fx(fx_path, [])

    result, stats = build_monthly_values(positions_path, prices_path, fx_path, as_of=date(2021, 4, 5))

    row = result.iloc[0]
    assert row["MonthEnd"] == "2021-03-31"
    assert row["Year"] == 2021
    assert row["Month"] == 3
    assert row["MarketValueEUR"] == 1200.0
    # Real bug caught by the user against real data: this used to assert
    # 2200.0 (market value MINUS a negative net_invested, i.e. added
    # together) instead of the real gain (value minus amount actually paid).
    assert row["UnrealizedGainEUR"] == 200.0
    assert row["FXRate"] == 1.0


def test_portfolio_owner_carried_through_from_positions(tmp_path):
    positions_path = tmp_path / "PortfolioPositions.xlsx"
    prices_path = tmp_path / "InstrumentPrices.xlsx"
    fx_path = tmp_path / "FXRates.xlsx"

    _positions(positions_path, [
        {"Broker": "NORDNET", "Portfolio": "1", "PortfolioOwner": "PERSON_A", "NormalizedInstrument": "SAMPO A", "Date": "2021-03-15", "CumulativeQuantity": 100, "CumulativeNetInvested": -1000},
    ])
    _prices(prices_path, [
        {"NormalizedInstrument": "SAMPO A", "Date": "2021-03-31", "Close": 12.0, "Currency": "EUR"},
    ])
    _fx(fx_path, [])

    result, stats = build_monthly_values(positions_path, prices_path, fx_path, as_of=date(2021, 4, 5))

    assert result.iloc[0]["PortfolioOwner"] == "PERSON_A"


def test_portfolio_owner_blank_when_positions_sheet_predates_the_column(tmp_path):
    """PortfolioPositions.xlsx files written before this column existed must
    not error - just carry an empty PortfolioOwner through."""
    positions_path = tmp_path / "PortfolioPositions.xlsx"
    prices_path = tmp_path / "InstrumentPrices.xlsx"
    fx_path = tmp_path / "FXRates.xlsx"

    _write_sheet(
        positions_path,
        "PortfolioPositions",
        ["Broker", "Portfolio", "NormalizedInstrument", "Date", "TransactionType", "QuantityDelta", "CumulativeQuantity", "CashDelta", "CumulativeNetInvested"],
        [{"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "SAMPO A", "Date": "2021-03-15", "CumulativeQuantity": 100, "CumulativeNetInvested": -1000}],
    )
    _prices(prices_path, [
        {"NormalizedInstrument": "SAMPO A", "Date": "2021-03-31", "Close": 12.0, "Currency": "EUR"},
    ])
    _fx(fx_path, [])

    result, stats = build_monthly_values(positions_path, prices_path, fx_path, as_of=date(2021, 4, 5))

    assert result.iloc[0]["PortfolioOwner"] == ""


def test_non_eur_instrument_converts_using_fx_rate(tmp_path):
    positions_path = tmp_path / "PortfolioPositions.xlsx"
    prices_path = tmp_path / "InstrumentPrices.xlsx"
    fx_path = tmp_path / "FXRates.xlsx"

    _positions(positions_path, [
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "NORWEGIAN AIR SHUTTLE", "Date": "2021-03-15", "CumulativeQuantity": 30, "CumulativeNetInvested": -159.26},
    ])
    _prices(prices_path, [
        {"NormalizedInstrument": "NORWEGIAN AIR SHUTTLE", "Date": "2021-03-31", "Close": 13.0, "Currency": "NOK"},
    ])
    _fx(fx_path, [
        {"Currency": "NOK", "Date": "2021-03-31", "Rate": 0.1},
    ])

    result, stats = build_monthly_values(positions_path, prices_path, fx_path, as_of=date(2021, 4, 5))

    row = result.iloc[0]
    # 30 * 13.0 * 0.1 = 39.0
    assert row["MarketValueEUR"] == 39.0
    assert row["FXRate"] == 0.1


def test_price_and_fx_lookups_backfill_to_most_recent_prior_date(tmp_path):
    """Month-end (e.g. a weekend) may not have an exact price/FX row - use the
    most recent one on or before it, not an exact-date match."""
    positions_path = tmp_path / "PortfolioPositions.xlsx"
    prices_path = tmp_path / "InstrumentPrices.xlsx"
    fx_path = tmp_path / "FXRates.xlsx"

    _positions(positions_path, [
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "NORWEGIAN AIR SHUTTLE", "Date": "2021-03-15", "CumulativeQuantity": 10, "CumulativeNetInvested": -50},
    ])
    _prices(prices_path, [
        {"NormalizedInstrument": "NORWEGIAN AIR SHUTTLE", "Date": "2021-03-29", "Close": 10.0, "Currency": "NOK"},
    ])
    _fx(fx_path, [
        {"Currency": "NOK", "Date": "2021-03-28", "Rate": 0.1},
    ])

    result, stats = build_monthly_values(positions_path, prices_path, fx_path, as_of=date(2021, 4, 5))

    row = result.iloc[0]
    assert row["PriceDate"] == "2021-03-29"
    assert row["FXDate"] == "2021-03-28"
    assert row["MarketValueEUR"] == 10.0


def test_fully_sold_out_instrument_stops_appearing(tmp_path):
    positions_path = tmp_path / "PortfolioPositions.xlsx"
    prices_path = tmp_path / "InstrumentPrices.xlsx"
    fx_path = tmp_path / "FXRates.xlsx"

    _positions(positions_path, [
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "SAMPO A", "Date": "2021-01-15", "CumulativeQuantity": 100, "CumulativeNetInvested": -1000},
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "SAMPO A", "Date": "2021-02-15", "CumulativeQuantity": 0, "CumulativeNetInvested": 200},
    ])
    _prices(prices_path, [
        {"NormalizedInstrument": "SAMPO A", "Date": "2021-01-31", "Close": 10.0, "Currency": "EUR"},
        {"NormalizedInstrument": "SAMPO A", "Date": "2021-02-28", "Close": 11.0, "Currency": "EUR"},
        {"NormalizedInstrument": "SAMPO A", "Date": "2021-03-31", "Close": 12.0, "Currency": "EUR"},
    ])
    _fx(fx_path, [])

    result, stats = build_monthly_values(positions_path, prices_path, fx_path, as_of=date(2021, 4, 5))

    assert list(result["MonthEnd"]) == ["2021-01-31"]


def test_only_fully_completed_months_included_not_current_in_progress(tmp_path):
    positions_path = tmp_path / "PortfolioPositions.xlsx"
    prices_path = tmp_path / "InstrumentPrices.xlsx"
    fx_path = tmp_path / "FXRates.xlsx"

    _positions(positions_path, [
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "SAMPO A", "Date": "2021-03-15", "CumulativeQuantity": 100, "CumulativeNetInvested": -1000},
    ])
    _prices(prices_path, [
        {"NormalizedInstrument": "SAMPO A", "Date": "2021-03-31", "Close": 12.0, "Currency": "EUR"},
        {"NormalizedInstrument": "SAMPO A", "Date": "2021-04-10", "Close": 13.0, "Currency": "EUR"},
    ])
    _fx(fx_path, [])

    result, stats = build_monthly_values(positions_path, prices_path, fx_path, as_of=date(2021, 4, 15))

    assert list(result["MonthEnd"]) == ["2021-03-31"]
    assert stats["last_month_end"] == "2021-03-31"


def test_missing_price_is_reported_and_row_skipped(tmp_path):
    positions_path = tmp_path / "PortfolioPositions.xlsx"
    prices_path = tmp_path / "InstrumentPrices.xlsx"
    fx_path = tmp_path / "FXRates.xlsx"

    _positions(positions_path, [
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "SAMPO A", "Date": "2021-01-15", "CumulativeQuantity": 100, "CumulativeNetInvested": -1000},
    ])
    _prices(prices_path, [])
    _fx(fx_path, [])

    result, stats = build_monthly_values(positions_path, prices_path, fx_path, as_of=date(2021, 2, 5))

    assert len(result) == 0
    assert stats["missing_price_instrument_months"] == [("SAMPO A", "2021-01")]


def test_missing_fx_rate_is_reported_and_row_skipped(tmp_path):
    positions_path = tmp_path / "PortfolioPositions.xlsx"
    prices_path = tmp_path / "InstrumentPrices.xlsx"
    fx_path = tmp_path / "FXRates.xlsx"

    _positions(positions_path, [
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "NORWEGIAN AIR SHUTTLE", "Date": "2021-01-15", "CumulativeQuantity": 10, "CumulativeNetInvested": -50},
    ])
    _prices(prices_path, [
        {"NormalizedInstrument": "NORWEGIAN AIR SHUTTLE", "Date": "2021-01-31", "Close": 10.0, "Currency": "NOK"},
    ])
    _fx(fx_path, [])

    result, stats = build_monthly_values(positions_path, prices_path, fx_path, as_of=date(2021, 2, 5))

    assert len(result) == 0
    assert stats["missing_fx_currency_months"] == [("NOK", "2021-01")]
