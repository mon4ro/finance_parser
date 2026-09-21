import pandas as pd
from openpyxl import Workbook

from finance_parser.investments.enrich_dividend_local_currency import (
    parse_op_dividend_message,
    load_op_dividend_local_currency_details,
    enrich_dividend_history,
)


# Synthetic message matching the real OP dividend-notice pattern, using
# invented numbers (not copied from any real dividend).
SEK_MESSAGE = (
    "Viesti: OP Säilytys Oy TELIA COMPANY AB SE0000667925 Osinkotuotto "
    "Osinko 0,50 SEK/KplOmistettu määrä 100Kpl Tuoton määrä 50,00SEK"
    "Lähdevero SE15,0 % 7,50SEKVal.kurssi 11,000000"
)

EUR_MESSAGE = (
    "Viesti: OP Säilytys Oy NOKIA OYJ FI0009000681 Osinkotuotto "
    "Osinko 0,04 EUR/KplOmistettu määrä 660Kpl Tuoton määrä 26,40EUR"
    "Ennakonpid. 25,5 % 6,73EUR"
)


def test_parses_sek_dividend_message():
    result = parse_op_dividend_message(SEK_MESSAGE)

    assert result is not None
    assert result["NormalizedInstrument"] == "TELIA COMPANY AB"
    assert result["ISIN"] == "SE0000667925"
    assert result["LocalCurrency"] == "SEK"
    assert result["GrossDividendLocal"] == 50.00
    assert result["TaxWithheldLocal"] == 7.50
    assert result["ExchangeRate"] == 11.0


def test_eur_native_dividend_message_does_not_match():
    """No Val.kurssi in a EUR-native dividend - nothing to extract."""
    assert parse_op_dividend_message(EUR_MESSAGE) is None


def test_unrelated_message_does_not_match():
    assert parse_op_dividend_message("TILISIIRTO") is None
    assert parse_op_dividend_message("") is None
    assert parse_op_dividend_message(None) is None


def test_load_details_from_budgeting_workbook(tmp_path):
    path = tmp_path / "ParsedTransactions.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "UnifiedTransactions"
    ws.append(["Date", "Amount", "Message"])
    ws.append(["2021-01-15", 42.5, SEK_MESSAGE])
    ws.append(["2021-01-16", 19.67, EUR_MESSAGE])
    ws.append(["2021-01-17", -10.0, "TILISIIRTO"])
    wb.save(path)

    details = load_op_dividend_local_currency_details(path)

    assert len(details) == 1
    row = details.iloc[0]
    assert row["NormalizedInstrument"] == "TELIA COMPANY AB"
    assert row["TradeDate"] == pd.Timestamp("2021-01-15")
    assert row["LocalCurrency"] == "SEK"


def test_load_details_missing_workbook_returns_empty(tmp_path):
    details = load_op_dividend_local_currency_details(tmp_path / "does_not_exist.xlsx")
    assert len(details) == 0


def test_enrich_dividend_history_matches_by_instrument_and_date():
    dividend_history = pd.DataFrame([
        {"TradeDate": "2021-01-15", "NormalizedInstrument": "TELIA COMPANY AB", "GrossDividendEUR": 20.0},
        {"TradeDate": "2021-01-16", "NormalizedInstrument": "NOKIA", "GrossDividendEUR": 10.0},
    ])
    local_details = pd.DataFrame([
        {"NormalizedInstrument": "TELIA COMPANY AB", "TradeDate": pd.Timestamp("2021-01-15"), "ISIN": "SE0000667925",
         "LocalCurrency": "SEK", "GrossDividendLocal": 200.0, "TaxWithheldLocal": 30.0, "ExchangeRate": 10.0},
    ])

    enriched, matched = enrich_dividend_history(dividend_history, local_details)

    assert matched == 1
    telia_row = enriched[enriched["NormalizedInstrument"] == "TELIA COMPANY AB"].iloc[0]
    assert telia_row["LocalCurrency"] == "SEK"
    assert telia_row["GrossDividendLocal"] == 200.0
    assert telia_row["ExchangeRate"] == 10.0

    nokia_row = enriched[enriched["NormalizedInstrument"] == "NOKIA"].iloc[0]
    assert nokia_row["LocalCurrency"] == ""


def test_enrich_dividend_history_no_details_leaves_columns_blank():
    dividend_history = pd.DataFrame([
        {"TradeDate": "2021-01-15", "NormalizedInstrument": "TELIA COMPANY AB", "GrossDividendEUR": 20.0},
    ])

    enriched, matched = enrich_dividend_history(dividend_history, pd.DataFrame(columns=["NormalizedInstrument", "TradeDate", "ISIN", "LocalCurrency", "GrossDividendLocal", "TaxWithheldLocal", "ExchangeRate"]))

    assert matched == 0
    assert enriched.iloc[0]["LocalCurrency"] == ""
