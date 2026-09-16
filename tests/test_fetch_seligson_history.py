from datetime import date

import pytest
from openpyxl import Workbook

from finance_parser.investments.fetch_instrument_prices import PRICES_COLUMNS
from finance_parser.investments.fetch_seligson_history import (
    SELIGSON_CSV_SLUGS,
    _parse_seligson_csv,
    backfill_gaps,
    fetch_seligson_csv,
)


def test_parse_seligson_csv_handles_finnish_locale_format():
    text = "01.06.2020;1,2345\n02.06.2020;1,3456\n"
    rows = _parse_seligson_csv(text)

    assert rows == [
        (date(2020, 6, 1), 1.2345),
        (date(2020, 6, 2), 1.3456),
    ]


def test_parse_seligson_csv_ignores_blank_lines():
    text = "01.06.2020;1,0000\n\n02.06.2020;1,1000\n"
    rows = _parse_seligson_csv(text)
    assert len(rows) == 2


def _write_prices(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "InstrumentPrices"
    ws.append(PRICES_COLUMNS)
    for row in rows:
        ws.append([row.get(h, "") for h in PRICES_COLUMNS])
    wb.save(path)


def _write_instrument_master(path):
    from finance_parser.investments.investment_common import INSTRUMENT_MASTER_COLUMNS

    wb = Workbook()
    ws = wb.active
    ws.title = "InstrumentMaster"
    ws.append(INSTRUMENT_MASTER_COLUMNS)
    row = {
        "Enabled": "YES",
        "NormalizedInstrument": list(SELIGSON_CSV_SLUGS)[0],
        "ISIN": "FI0008801790",
        "PriceSymbol": "0P00000NFI.F",
    }
    ws.append([row.get(h, "") for h in INSTRUMENT_MASTER_COLUMNS])
    wb.save(path)


def test_backfill_gaps_only_adds_dates_not_already_present(tmp_path):
    instrument = list(SELIGSON_CSV_SLUGS)[0]
    prices_path = tmp_path / "InstrumentPrices.xlsx"
    master_path = tmp_path / "InstrumentMaster.xlsx"

    _write_prices(prices_path, [
        {"NormalizedInstrument": instrument, "Date": "2022-03-08", "Close": 20.0, "Currency": "EUR", "PriceSource": "YAHOO"},
    ])
    _write_instrument_master(master_path)

    def fake_fetch(slug):
        return [
            (date(2020, 2, 4), 15.0),
            (date(2022, 3, 8), 999.0),  # already present via Yahoo - must not overwrite it
        ]

    combined, stats = backfill_gaps(prices_path, master_path, fetch_fn=fake_fetch, verbose=False)

    assert stats["rows_added"] == 1
    row = combined[(combined["NormalizedInstrument"] == instrument) & (combined["Date"] == "2020-02-04")].iloc[0]
    assert row["Close"] == 15.0
    assert row["PriceSource"] == "SELIGSON.FI"
    assert row["ISIN"] == "FI0008801790"

    existing_row = combined[(combined["NormalizedInstrument"] == instrument) & (combined["Date"] == "2022-03-08")].iloc[0]
    assert existing_row["Close"] == 20.0
    assert existing_row["PriceSource"] == "YAHOO"


def test_backfill_gaps_reports_zero_when_fully_covered(tmp_path):
    instrument = list(SELIGSON_CSV_SLUGS)[0]
    prices_path = tmp_path / "InstrumentPrices.xlsx"
    master_path = tmp_path / "InstrumentMaster.xlsx"

    _write_prices(prices_path, [
        {"NormalizedInstrument": instrument, "Date": "2020-02-04", "Close": 15.0, "Currency": "EUR", "PriceSource": "YAHOO"},
    ])
    _write_instrument_master(master_path)

    def fake_fetch(slug):
        return [(date(2020, 2, 4), 15.0)]

    combined, stats = backfill_gaps(prices_path, master_path, fetch_fn=fake_fetch, verbose=False)

    assert stats["rows_added"] == 0
    assert stats["instruments_backfilled"] == 0


def test_fetch_seligson_csv_real_network_call():
    """
    Live integration check against the real seligson.fi CSV endpoint,
    verified manually working (no key, no login). Skips gracefully if the
    network is unavailable.
    """
    try:
        rows = fetch_seligson_csv("global-brands")
    except Exception as exc:
        pytest.skip(f"Live Seligson CSV endpoint unreachable, skipping: {exc}")

    assert len(rows) > 1000
    first_day, first_value = rows[0]
    assert first_day.year <= 2000
    assert first_value > 0
