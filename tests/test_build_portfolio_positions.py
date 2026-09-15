import pytest
from openpyxl import Workbook

from finance_parser.investments.build_portfolio_positions import build_positions


TRANSACTIONS_HEADERS = [
    "Broker", "Portfolio", "NormalizedInstrument", "TransactionType",
    "TradeDate", "Quantity", "CashAmount",
]


def _write_transactions(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "InvestmentTransactions"
    ws.append(TRANSACTIONS_HEADERS)
    for row in rows:
        ws.append([row.get(h, "") for h in TRANSACTIONS_HEADERS])
    wb.save(path)


OPENING_POSITIONS_HEADERS = ["Broker", "Portfolio", "NormalizedInstrument", "Date", "Quantity", "Notes"]


def _write_instrument_master(path, opening_position_rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "InstrumentMaster"
    ws.append(["NormalizedInstrument"])
    opening = wb.create_sheet("OpeningPositions")
    opening.append(OPENING_POSITIONS_HEADERS)
    for row in opening_position_rows:
        opening.append([row.get(h, "") for h in OPENING_POSITIONS_HEADERS])
    wb.save(path)


def _no_instrument_master(tmp_path):
    """A path that deliberately doesn't exist, so build_positions() gets no
    opening-position seeding unless a test explicitly writes one - keeps
    tests isolated from the real rules/investments/InstrumentMaster.xlsx,
    which build_positions() would otherwise read by default."""
    return tmp_path / "InstrumentMaster.xlsx"


def test_build_positions_simple_buy_sell_tracks_cumulative_totals(tmp_path):
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "SAMPO A", "TransactionType": "BUY", "TradeDate": "2021-01-01", "Quantity": 100, "CashAmount": -1000},
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "SAMPO A", "TransactionType": "SELL", "TradeDate": "2021-06-01", "Quantity": 40, "CashAmount": 450},
    ])

    positions, stats = build_positions(path, _no_instrument_master(tmp_path))

    rows = positions[positions["NormalizedInstrument"] == "SAMPO A"]
    assert list(rows["CumulativeQuantity"]) == [100.0, 60.0]
    assert list(rows["CumulativeNetInvested"]) == [-1000.0, -550.0]
    assert stats["negative_quantity_instruments"] == {}


def test_build_positions_handles_inconsistent_sell_quantity_sign(tmp_path):
    """
    Real data: NORDNET reports SELL as a positive magnitude, EVLI reports it
    already negative. build_positions() must always derive sign from
    TransactionType via abs(Quantity), never trust the raw sign.
    """
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"Broker": "EVLI", "Portfolio": "1", "NormalizedInstrument": "NOKIA", "TransactionType": "BUY", "TradeDate": "2021-01-01", "Quantity": 50, "CashAmount": -500},
        {"Broker": "EVLI", "Portfolio": "1", "NormalizedInstrument": "NOKIA", "TransactionType": "SELL", "TradeDate": "2021-06-01", "Quantity": -20, "CashAmount": 220},
    ])

    positions, _ = build_positions(path, _no_instrument_master(tmp_path))

    rows = positions[positions["NormalizedInstrument"] == "NOKIA"]
    assert list(rows["CumulativeQuantity"]) == [50.0, 30.0]


def test_build_positions_stock_split_resets_to_absolute_quantity(tmp_path):
    """
    A SPLIT AP JÄTTÖ row's Quantity is the real resulting absolute share
    count, not an incremental delta - confirmed against real Finnair/
    Norwegian Air Shuttle data. build_positions() must land the cumulative
    total exactly on that value, not add it on top of the running total.
    """
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "FINNAIR", "TransactionType": "BUY", "TradeDate": "2021-01-01", "Quantity": 1200, "CashAmount": -750},
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "FINNAIR", "TransactionType": "SELL", "TradeDate": "2023-11-03", "Quantity": 400, "CashAmount": 27},
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "FINNAIR", "TransactionType": "SPLIT AP JÄTTÖ", "TradeDate": "2024-03-21", "Quantity": 120, "CashAmount": 0},
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "FINNAIR", "TransactionType": "SPLIT AP OTTO", "TradeDate": "2024-03-21", "Quantity": 800, "CashAmount": 0},
    ])

    positions, stats = build_positions(path, _no_instrument_master(tmp_path))

    rows = positions[positions["NormalizedInstrument"] == "FINNAIR"]
    assert list(rows["CumulativeQuantity"]) == [1200.0, 800.0, 120.0]
    assert stats["negative_quantity_instruments"] == {}


def test_build_positions_flags_negative_quantity_as_warning_not_silent(tmp_path):
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"Broker": "OP", "Portfolio": "OP", "NormalizedInstrument": "OP-DELTA A", "TransactionType": "BUY", "TradeDate": "2020-01-01", "Quantity": 10, "CashAmount": -100},
        {"Broker": "OP", "Portfolio": "OP", "NormalizedInstrument": "OP-DELTA A", "TransactionType": "VAIHTO AP-OTTO", "TradeDate": "2021-01-01", "Quantity": 64, "CashAmount": 0},
    ])

    positions, stats = build_positions(path, _no_instrument_master(tmp_path))

    assert stats["negative_quantity_instruments"] == {("OP", "OP", "OP-DELTA A"): -54.0}


def test_build_positions_excludes_known_neutral_types_from_unhandled_report(tmp_path):
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "SAMPO A", "TransactionType": "DIVIDEND", "TradeDate": "2021-03-01", "Quantity": 0, "CashAmount": 12.5},
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "SAMPO A", "TransactionType": "SOME UNKNOWN TYPE", "TradeDate": "2021-04-01", "Quantity": 0, "CashAmount": 0},
    ])

    _, stats = build_positions(path, _no_instrument_master(tmp_path))

    assert stats["unhandled_types"] == {"SOME UNKNOWN TYPE": 1}


def test_build_positions_excludes_blank_normalized_instrument_rows(tmp_path):
    path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(path, [
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "", "TransactionType": "FEE", "TradeDate": "2021-01-01", "Quantity": 0, "CashAmount": -5},
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "SAMPO A", "TransactionType": "BUY", "TradeDate": "2021-01-01", "Quantity": 10, "CashAmount": -100},
    ])

    positions, stats = build_positions(path, _no_instrument_master(tmp_path))

    assert stats["transactions_scanned"] == 1
    assert len(positions) == 1


def test_build_positions_seeds_opening_balance_before_real_transaction_history(tmp_path):
    """
    A gift/inherited holding with no purchase row in any broker export (real
    case: OP-Eurooppa/OP-Suomi Pienyhtiöt A, received as gifts before
    tracking began) is seeded from InstrumentMaster's OpeningPositions sheet
    as the earliest event for that instrument, so a later SELL doesn't go
    negative.
    """
    investments_path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(investments_path, [
        {"Broker": "OP", "Portfolio": "OP", "NormalizedInstrument": "OP-EUROOPPA PIENYHTIÖT A", "TransactionType": "SELL", "TradeDate": "2024-12-27", "Quantity": 0.7761, "CashAmount": 30.0},
    ])

    instrument_master_path = tmp_path / "InstrumentMaster.xlsx"
    _write_instrument_master(instrument_master_path, [
        # Deliberately mixed-case, matching how a human would actually type it
        # into the sheet - NormalizedInstrument matching must be
        # case-insensitive (real bug: it silently failed to join at first).
        {"Broker": "OP", "Portfolio": "OP", "NormalizedInstrument": "OP-Eurooppa Pienyhtiöt A", "Date": "2013-10-10", "Quantity": 0.78},
    ])

    positions, stats = build_positions(investments_path, instrument_master_path)

    rows = positions[positions["NormalizedInstrument"] == "OP-EUROOPPA PIENYHTIÖT A"]
    assert list(rows["TransactionType"]) == ["OPENING BALANCE", "SELL"]
    assert list(rows["CumulativeQuantity"]) == pytest.approx([0.78, 0.0039])
    assert stats["negative_quantity_instruments"] == {}


def test_build_positions_without_opening_positions_sheet_is_a_no_op(tmp_path):
    """InstrumentMaster.xlsx existing but lacking an OpeningPositions sheet
    (the common case - most InstrumentMaster.xlsx files predate this
    feature) must not error, just contribute nothing."""
    investments_path = tmp_path / "ParsedInvestments.xlsx"
    _write_transactions(investments_path, [
        {"Broker": "NORDNET", "Portfolio": "1", "NormalizedInstrument": "SAMPO A", "TransactionType": "BUY", "TradeDate": "2021-01-01", "Quantity": 100, "CashAmount": -1000},
    ])

    instrument_master_path = tmp_path / "InstrumentMaster.xlsx"
    wb = Workbook()
    wb.active.title = "InstrumentMaster"
    wb.save(instrument_master_path)

    positions, stats = build_positions(investments_path, instrument_master_path)

    assert stats["transactions_scanned"] == 1
    assert len(positions) == 1
