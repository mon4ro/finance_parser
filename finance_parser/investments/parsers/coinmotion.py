from __future__ import annotations

from pathlib import Path

import pandas as pd

from finance_parser.common import (
    file_modified_date,
    format_date,
    normalise_amount,
    normalise_header,
    normalise_text,
    read_input_file,
    stable_hash,
)

from finance_parser.investments.investment_common import INVESTMENT_RAW_COLUMNS


BROKER = "COINMOTION"
BROKER_RAW_SHEET = "CoinmotionRawExport"
# Raw investment ID prefix: CMO-

# Coinmotion's "Transaction Statement" CSV export columns:
# fromCurrency,toCurrency,type,eurAmount,cryptoAmount,rate,fee,feeCurrency,time
#
# Example rows (see tests/fixtures/investments/coinmotion_sample.csv):
#   EUR,EUR,deposit,1000.00,,0.00,0.00,EUR,2026-03-01T09:00:00+02:00
#   EUR,BTC,market_trade,1000.00,0.02,49000.00,20.00,EUR,2026-03-01T09:05:00+02:00
#
# eurAmount is the TOTAL EUR debited/credited for the row (fee-inclusive for a
# trade), not just rate*cryptoAmount - confirmed against a real export row:
# rate*cryptoAmount always comes out exactly eurAmount-fee short. So CashAmount
# below uses eurAmount directly (the real total cash movement), with the fee
# tracked separately in BrokerageFee rather than re-derived.

COINMOTION_REQUIRED_COLUMNS = [
    "fromCurrency",
    "toCurrency",
    "type",
    "eurAmount",
    "cryptoAmount",
    "rate",
    "fee",
    "feeCurrency",
    "time",
]

# Single personal wallet - no sub-portfolios in this export.
PORTFOLIO = "COINMOTION"


def _read_coinmotion_file(path: Path) -> pd.DataFrame:
    df = read_input_file(path, required_columns=COINMOTION_REQUIRED_COLUMNS)
    df.columns = [normalise_header(c) for c in df.columns]

    missing = [c for c in COINMOTION_REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"Missing Coinmotion columns: {', '.join(missing)}. "
            f"Available columns: {', '.join(map(str, df.columns))}"
        )

    return df


def can_parse(path: Path) -> tuple[bool, str]:
    try:
        _read_coinmotion_file(path)
        return True, "Matched Coinmotion transaction statement column set"
    except Exception as exc:
        return False, str(exc)


def make_investment_raw_id(row: pd.Series) -> str:
    return "CMO-" + stable_hash([
        BROKER,
        row.get("RawTime", ""),
        row.get("TransactionTypeRaw", ""),
        row.get("InstrumentName", ""),
        row.get("Quantity", ""),
        row.get("CashAmount", ""),
    ], length=16)


def make_broker_raw_export_id(row: pd.Series, source_file: str) -> str:
    return "COINMOTIONRAW-" + stable_hash([
        BROKER,
        source_file,
        row.get("fromCurrency", ""),
        row.get("toCurrency", ""),
        row.get("type", ""),
        row.get("eurAmount", ""),
        row.get("cryptoAmount", ""),
        row.get("time", ""),
    ], length=16)


def _transaction_type_raw(from_currency: str, to_currency: str, raw_type: str) -> str:
    """
    Resolve directional BUY/SELL/DEPOSIT/WITHDRAWAL from the currency pair,
    not just the `type` column, since "market_trade" alone doesn't say which
    side is the crypto. Anything unrecognised is passed through uppercased so
    it surfaces as unhandled downstream rather than being silently guessed.
    """
    key = normalise_text(raw_type).upper()

    if from_currency == to_currency:
        if key == "DEPOSIT":
            return "DEPOSIT"
        if key == "WITHDRAWAL":
            return "WITHDRAWAL"
        return key

    if key == "MARKET_TRADE":
        if from_currency == "EUR":
            return "BUY"
        if to_currency == "EUR":
            return "SELL"

    return key


def parse_file(path: Path, imported_at: str) -> pd.DataFrame:
    df = _read_coinmotion_file(path).copy()

    df = df[df["time"].notna()]
    df = df[df["time"].astype(str).str.strip() != ""]

    export_date = file_modified_date(path)

    rows: list[dict[str, object]] = []

    for _, row in df.iterrows():
        from_currency = normalise_text(row.get("fromCurrency", "")).upper()
        to_currency = normalise_text(row.get("toCurrency", "")).upper()
        raw_type = row.get("type", "")
        transaction_type = _transaction_type_raw(from_currency, to_currency, raw_type)
        raw_time = normalise_text(row.get("time", ""))

        crypto_currency = to_currency if to_currency != "EUR" else from_currency
        is_crypto_trade = transaction_type in {"BUY", "SELL"}

        eur_amount = normalise_amount(row.get("eurAmount", 0) or 0)
        crypto_amount = normalise_amount(row.get("cryptoAmount", 0) or 0)
        rate = normalise_amount(row.get("rate", 0) or 0)
        fee = normalise_amount(row.get("fee", 0) or 0)
        fee_currency = normalise_text(row.get("feeCurrency", "")) or "EUR"

        # Real total EUR movement: a BUY debits the account (negative), a SELL
        # or DEPOSIT credits it (positive), a WITHDRAWAL debits it (negative).
        if transaction_type == "BUY":
            cash_amount = -eur_amount
        elif transaction_type == "WITHDRAWAL":
            cash_amount = -eur_amount
        else:
            cash_amount = eur_amount

        parsed_row = {
            "Broker": BROKER,
            "Portfolio": PORTFOLIO,
            "PortfolioOwner": "",
            "PortfolioType": "",
            "BookingDate": format_date(raw_time),
            "TradeDate": format_date(raw_time),
            "SettlementDate": format_date(raw_time),
            "TransactionTypeRaw": transaction_type,
            "InstrumentName": crypto_currency if is_crypto_trade else "",
            "ISIN": "",
            "Quantity": crypto_amount if is_crypto_trade else 0,
            "UnitPrice": rate if is_crypto_trade else 0,
            "Interest": "",
            "TotalFees": fee,
            "TotalFeesCurrency": fee_currency,
            "CashAmount": cash_amount,
            "CashCurrency": "EUR",
            "AcquisitionValue": eur_amount if is_crypto_trade else "",
            "AcquisitionCurrency": "EUR",
            "Result": "",
            "ResultCurrency": "EUR",
            "TotalQuantity": "",
            "CashBalance": "",
            "ExchangeRate": "",
            "Description": f"Coinmotion {raw_type} {from_currency}->{to_currency}",
            "CancellationDate": "",
            "CalculationID": "",
            "ConfirmationNumber": "",
            "BrokerageFee": fee,
            "BrokerageFeeCurrency": fee_currency,
            "ReferenceExchangeRate": "",
            "OriginalInterest": "",
            "ExportDate": export_date,
            "ImportedAt": imported_at,
            "SourceFile": path.name,
            "RawTime": raw_time,
        }

        parsed_row["InvestmentRawID"] = make_investment_raw_id(pd.Series(parsed_row))
        rows.append({column: parsed_row.get(column, "") for column in INVESTMENT_RAW_COLUMNS})

    return pd.DataFrame(rows, columns=INVESTMENT_RAW_COLUMNS)


def parse_broker_raw_rows(path: Path, imported_at: str) -> pd.DataFrame:
    df = _read_coinmotion_file(path).copy()

    df.insert(0, "BrokerRawExportID", [
        make_broker_raw_export_id(row, path.name) for _, row in df.iterrows()
    ])
    df.insert(1, "Broker", BROKER)
    df.insert(2, "SourceFile", path.name)
    df.insert(3, "ExportDate", file_modified_date(path))
    df.insert(4, "ImportedAt", imported_at)

    return df
