from datetime import date

import pytest
from openpyxl import Workbook

from finance_parser.investments.fetch_avanza_history import (
    AVANZA_FUND_IDS,
    backfill_gaps,
    fetch_avanza_chart,
)
from finance_parser.investments.fetch_instrument_prices import PRICES_COLUMNS
from finance_parser.investments.investment_common import INSTRUMENT_MASTER_COLUMNS


def _write_prices(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "InstrumentPrices"
    ws.append(PRICES_COLUMNS)
    for row in rows:
        ws.append([row.get(h, "") for h in PRICES_COLUMNS])
    wb.save(path)


def _write_instrument_master(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "InstrumentMaster"
    ws.append(INSTRUMENT_MASTER_COLUMNS)
    row = {
        "Enabled": "YES",
        "NormalizedInstrument": list(AVANZA_FUND_IDS)[0],
        "ISIN": "SE0011337195",
        "PriceSymbol": "0P0000ULAP.ST",
    }
    ws.append([row.get(h, "") for h in INSTRUMENT_MASTER_COLUMNS])
    wb.save(path)


def test_backfill_gaps_only_adds_dates_not_already_present(tmp_path):
    instrument = list(AVANZA_FUND_IDS)[0]
    prices_path = tmp_path / "InstrumentPrices.xlsx"
    master_path = tmp_path / "InstrumentMaster.xlsx"

    _write_prices(prices_path, [
        {"NormalizedInstrument": instrument, "Date": "2022-03-07", "Close": 592.51, "Currency": "SEK", "PriceSource": "YAHOO"},
    ])
    _write_instrument_master(master_path)

    def fake_fetch(fund_id, start, end):
        return [
            (date(2022, 2, 28), 594.16),
            (date(2022, 3, 7), 999.0),  # already present via Yahoo - must not overwrite it
        ]

    combined, stats = backfill_gaps(prices_path, master_path, fetch_fn=fake_fetch, verbose=False)

    assert stats["rows_added"] == 1
    row = combined[(combined["NormalizedInstrument"] == instrument) & (combined["Date"] == "2022-02-28")].iloc[0]
    assert row["Close"] == 594.16
    assert row["PriceSource"] == "AVANZA"
    assert row["Currency"] == "SEK"
    assert row["ISIN"] == "SE0011337195"

    existing_row = combined[(combined["NormalizedInstrument"] == instrument) & (combined["Date"] == "2022-03-07")].iloc[0]
    assert existing_row["Close"] == 592.51
    assert existing_row["PriceSource"] == "YAHOO"


def test_backfill_gaps_requests_from_fund_start_to_earliest_known_date(tmp_path):
    """The fetch range should reach back to before any existing data, not
    just the most recent gap - Avanza's own history often predates the
    fund's actual holding period, worth capturing."""
    instrument = list(AVANZA_FUND_IDS)[0]
    prices_path = tmp_path / "InstrumentPrices.xlsx"
    master_path = tmp_path / "InstrumentMaster.xlsx"

    _write_prices(prices_path, [
        {"NormalizedInstrument": instrument, "Date": "2022-03-07", "Close": 592.51, "Currency": "SEK", "PriceSource": "YAHOO"},
    ])
    _write_instrument_master(master_path)

    captured_range = {}

    def fake_fetch(fund_id, start, end):
        captured_range["start"] = start
        captured_range["end"] = end
        return []

    backfill_gaps(prices_path, master_path, fetch_fn=fake_fetch, verbose=False)

    assert captured_range["start"] == date(2000, 1, 1)
    assert captured_range["end"] == date(2022, 3, 7)


def test_fetch_avanza_chart_real_network_call():
    """
    Live integration check against the real Avanza chart API, verified
    manually working (no key, no login). Skips gracefully if the network is
    unavailable.
    """
    fund_id = AVANZA_FUND_IDS[list(AVANZA_FUND_IDS)[0]]
    try:
        rows = fetch_avanza_chart(fund_id, date(2022, 1, 1), date(2022, 2, 1))
    except Exception as exc:
        pytest.skip(f"Live Avanza chart API unreachable, skipping: {exc}")

    assert len(rows) > 0
    for day, value in rows:
        assert date(2022, 1, 1) <= day <= date(2022, 2, 1)
        assert value > 0
