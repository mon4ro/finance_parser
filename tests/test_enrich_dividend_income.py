import pandas as pd
from openpyxl import Workbook

from finance_parser.budgeting.enrich_dividend_income import (
    load_dividend_events,
    match_dividend_income,
)


DIVIDEND_HISTORY_HEADERS = [
    "TradeDate", "Year", "Month", "Broker", "Portfolio", "PortfolioOwner",
    "PortfolioType", "NormalizedInstrument", "GrossDividendEUR",
    "TaxWithheldEUR", "NetDividendEUR",
]


UNIFIED_COLUMNS = [
    "UnifiedID", "RawID", "SourceAccount", "SourceBank", "TransactionType",
    "Date", "Month", "Year", "Amount", "RawReceiver", "NormalizedReceiver",
    "Description", "Message", "Include", "Owner", "Supercategory",
    "Category", "Subcategory", "ReviewStatus", "Review/Notes",
]


def _unified_row(**overrides):
    row = {c: "" for c in UNIFIED_COLUMNS}
    row.update(overrides)
    return row


def _dividend_event(**overrides):
    row = {
        "Broker": "OP", "Portfolio": "OP", "PortfolioOwner": "PERSON_A",
        "NormalizedInstrument": "SAMPO A", "TradeDate": pd.Timestamp("2021-06-15"),
        "NetDividendEUR": 100.0,
    }
    row.update(overrides)
    return row


def test_exact_date_and_amount_match_fills_blank_fields():
    unified = pd.DataFrame([
        _unified_row(UnifiedID="U-1", Date="2021-06-15", Amount=100.0, Owner=""),
    ])
    events = pd.DataFrame([_dividend_event()])

    updated, stats = match_dividend_income(unified, events)

    row = updated.iloc[0]
    assert row["Supercategory"] == "INCOME"
    assert row["Category"] == "Passive income"
    assert row["Subcategory"] == "Dividend yields"
    assert row["Owner"] == "PERSON_A"
    assert stats["matched"] == 1
    assert stats["fields_updated"] == 4


def test_never_overwrites_an_already_filled_cell():
    """No already-filled cell gets touched - whether set by a rule or
    manually, per the user's explicit instruction."""
    unified = pd.DataFrame([
        _unified_row(UnifiedID="U-1", Date="2021-06-15", Amount=100.0, Owner="PERSON_B", Category="Something else"),
    ])
    events = pd.DataFrame([_dividend_event()])

    updated, stats = match_dividend_income(unified, events)

    row = updated.iloc[0]
    assert row["Owner"] == "PERSON_B"
    assert row["Category"] == "Something else"
    # Supercategory/Subcategory were blank and still get filled.
    assert row["Supercategory"] == "INCOME"
    assert row["Subcategory"] == "Dividend yields"


def test_row_with_review_notes_already_set_is_skipped_entirely():
    unified = pd.DataFrame([
        _unified_row(UnifiedID="U-1", Date="2021-06-15", Amount=100.0, **{"Review/Notes": "OK"}),
    ])
    events = pd.DataFrame([_dividend_event()])

    updated, stats = match_dividend_income(unified, events)

    row = updated.iloc[0]
    assert row["Category"] == ""
    assert stats["skipped_review_notes"] == 1
    assert stats["matched"] == 0


def test_date_tolerance_matches_within_window_not_beyond():
    unified = pd.DataFrame([
        _unified_row(UnifiedID="U-1", Date="2021-06-17", Amount=100.0),
    ])
    events = pd.DataFrame([_dividend_event(TradeDate=pd.Timestamp("2021-06-15"))])

    updated, stats = match_dividend_income(unified, events)

    assert stats["matched"] == 1
    assert stats["matched_via_date_tolerance"] == 1
    assert updated.iloc[0]["Category"] == "Passive income"


def test_no_matching_row_is_reported_unmatched_not_guessed():
    unified = pd.DataFrame([
        _unified_row(UnifiedID="U-1", Date="2021-01-01", Amount=50.0),
    ])
    events = pd.DataFrame([_dividend_event()])

    updated, stats = match_dividend_income(unified, events)

    assert stats["matched"] == 0
    assert len(stats["unmatched"]) == 1
    assert updated.iloc[0]["Category"] == ""


def test_multiple_candidates_are_reported_ambiguous_not_guessed():
    unified = pd.DataFrame([
        _unified_row(UnifiedID="U-1", Date="2021-06-15", Amount=100.0),
        _unified_row(UnifiedID="U-2", Date="2021-06-15", Amount=100.0),
    ])
    events = pd.DataFrame([_dividend_event()])

    updated, stats = match_dividend_income(unified, events)

    assert stats["matched"] == 0
    assert len(stats["ambiguous"]) == 1
    assert list(updated["Category"]) == ["", ""]


def test_load_dividend_events_filters_to_matchable_brokers_only(tmp_path):
    """Nordnet/EVLI dividends never settle same-day into a bank account -
    they're handled separately (investment_dividends.py's synthetic rows),
    not by this cross-reference."""
    path = tmp_path / "DividendHistory.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "DividendHistory"
    ws.append(DIVIDEND_HISTORY_HEADERS)
    ws.append(["2021-06-15", 2021, 6, "OP", "OP", "PERSON_A", "AOT", "SAMPO A", 0, 0, 100.0])
    ws.append(["2021-06-15", 2021, 6, "NORDNET", "1", "PERSON_A", "AOT", "FORTUM", 0, 0, 50.0])
    ws.append(["2021-06-15", 2021, 6, "EVLI", "EVLI", "PERSON_A", "", "NOKIA", 0, 0, 30.0])
    wb.save(path)

    events = load_dividend_events(path)

    assert list(events["Broker"].unique()) == ["OP"]
    assert len(events) == 1
