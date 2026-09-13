from __future__ import annotations

from pathlib import Path

import pandas as pd

from finance_parser.settings import get_settings

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


BROKER = "SELIGSON"
BROKER_RAW_SHEET = "SeligsonRawExport"
# Raw investment ID prefix: SEL-

# This parser is for the manually parsed Seligson table syntax, not for the PDF
# itself. Supported visible columns include:
#
# Salkku
# Arvopäivä
# Tapahtumanumero
# Osuuden arvo
# Osuuksien määrä
# Arvo yhteensä (€)
# Palkkio (€)
# Summa (€)

SELIGSON_REQUIRED_COLUMN_GROUPS = {
    "ValueDate": ["Arvopäivä", "Arvopaiva"],
    "TransactionNumber": ["Tapahtumanumero", "Tapahtuma", "TransactionNumber"],
    "CashAmount": ["Summa (€)", "Summa", "Summa EUR"],
}


def _key(value: object) -> str:
    text = normalise_header(value)
    text = text.replace("\u00a0", " ")
    text = " ".join(text.split())
    return text.upper()


def _find_column(df: pd.DataFrame, candidates: list[str], startswith: list[str] | None = None) -> str | None:
    startswith = startswith or []
    candidate_keys = {_key(c) for c in candidates}
    startswith_keys = [_key(c) for c in startswith]

    for col in df.columns:
        col_key = _key(col)
        if col_key in candidate_keys:
            return col

    for col in df.columns:
        col_key = _key(col)
        if any(col_key.startswith(prefix) for prefix in startswith_keys):
            return col

    return None


def _column_map(df: pd.DataFrame) -> dict[str, str | None]:
    return {
        "Portfolio": _find_column(df, ["Salkku", "Portfolio"]),
        "InstrumentName": _find_column(df, ["Rahasto", "Rahaston nimi", "Arvopaperi", "InstrumentName", "Instrument"]),
        "ValueDate": _find_column(df, ["Arvopäivä", "Arvopaiva"]),
        "TransactionNumber": _find_column(df, ["Tapahtumanumero", "Tapahtuma", "TransactionNumber"]),
        "UnitPrice": _find_column(df, ["Osuuden arvo"], startswith=["Osuuden arvo"]),
        "Quantity": _find_column(df, ["Osuuksien määrä", "Osuuksien maara", "Osuuksien"], startswith=["Osuuksien"]),
        "GrossValue": _find_column(df, ["Arvo yhteensä (€)", "Arvo yhteensä", "Arvo yhteensa (€)", "Arvo yhteensa", "Arvo yhte"], startswith=["Arvo yhte"]),
        "Fee": _find_column(df, ["Palkkio (€)", "Palkkio", "Palkkio EUR"], startswith=["Palkkio"]),
        "CashAmount": _find_column(df, ["Summa (€)", "Summa", "Summa EUR"], startswith=["Summa"]),
    }


def _read_seligson_file(path: Path) -> pd.DataFrame:
    # The shared reader is enough for this manually created Excel, but we avoid
    # passing a strict required_columns list because the exact syntax may use
    # Tapahtumanumero instead of Tapahtuma, and headers may contain NBSPs.
    df = read_input_file(path, required_columns=[])
    df.columns = [normalise_header(c).replace("\u00a0", " ").strip() for c in df.columns]

    colmap = _column_map(df)
    missing = [
        logical_name
        for logical_name in ["ValueDate", "TransactionNumber", "CashAmount"]
        if not colmap.get(logical_name)
    ]

    if missing:
        raise ValueError(
            f"Missing Seligson manually parsed columns: {', '.join(missing)}. "
            f"Available columns: {', '.join(map(str, df.columns))}"
        )

    return df


def can_parse(path: Path) -> tuple[bool, str]:
    try:
        df = _read_seligson_file(path)
        colmap = _column_map(df)
        return True, (
            "Matched manually parsed Seligson column set "
            f"(date={colmap.get('ValueDate')}, transaction={colmap.get('TransactionNumber')}, amount={colmap.get('CashAmount')})"
        )
    except Exception as exc:
        return False, str(exc)


def make_investment_raw_id(row: pd.Series) -> str:
    return "SEL-" + stable_hash([
        BROKER,
        row.get("Portfolio", ""),
        row.get("ValueDate", ""),
        row.get("SeligsonTransactionNumber", ""),
        row.get("InstrumentName", ""),
        row.get("Quantity", ""),
        row.get("UnitPrice", ""),
        row.get("CashAmount", ""),
    ], length=16)


def make_broker_raw_export_id(row: pd.Series, source_file: str, colmap: dict[str, str | None]) -> str:
    return "SELIGSONRAW-" + stable_hash([
        BROKER,
        source_file,
        row.get(colmap.get("Portfolio") or "", ""),
        row.get(colmap.get("InstrumentName") or "", ""),
        row.get(colmap.get("ValueDate") or "", ""),
        row.get(colmap.get("TransactionNumber") or "", ""),
        row.get(colmap.get("Quantity") or "", ""),
        row.get(colmap.get("CashAmount") or "", ""),
    ], length=16)




def _instrument_name_from_template(portfolio: str) -> str:
    """
    Return fallback Seligson instrument name from settings.

    Example:
      investments.broker_defaults.SELIGSON.instrument_name_template = "SELIGSON {portfolio}"
    """
    template = get_settings().broker_default(
        BROKER,
        "instrument_name_template",
        default="SELIGSON {portfolio}",
    )

    try:
        return str(template).format(portfolio=portfolio).strip()
    except Exception:
        return f"SELIGSON {portfolio}".strip()


def _infer_transaction_type_raw(cash_amount: float) -> str:
    # In the manually parsed Seligson syntax, positive Summa (€) rows from the
    # provided file represent subscriptions/purchases. Negative rows are treated
    # as redemptions. Adjust later if Seligson rows with other meanings appear.
    if cash_amount < 0:
        return "LUNASTUS"
    return "MERKINTÄ"


def _optional_amount(row: pd.Series, column: str | None, default: object = "") -> object:
    if column is None:
        return default
    value = normalise_text(row.get(column, ""))
    if value == "":
        return default
    return normalise_amount(value)


def parse_file(path: Path, imported_at: str) -> pd.DataFrame:
    df = _read_seligson_file(path).copy()
    colmap = _column_map(df)

    portfolio_col = colmap["Portfolio"]
    instrument_col = colmap["InstrumentName"]
    value_date_col = colmap["ValueDate"]
    transaction_col = colmap["TransactionNumber"]
    unit_price_col = colmap["UnitPrice"]
    quantity_col = colmap["Quantity"]
    gross_value_col = colmap["GrossValue"]
    fee_col = colmap["Fee"]
    cash_amount_col = colmap["CashAmount"]

    df = df[df[value_date_col].notna()]
    df = df[df[value_date_col].astype(str).str.strip() != ""]
    df = df[df[transaction_col].notna()]
    df = df[df[transaction_col].astype(str).str.strip() != ""]

    export_date = file_modified_date(path)

    rows: list[dict[str, object]] = []

    for _, row in df.iterrows():
        cash_amount = normalise_amount(row.get(cash_amount_col, 0))
        portfolio = normalise_text(row.get(portfolio_col, "")) if portfolio_col else "SELIGSON"

        # If the manually parsed file has no explicit fund/instrument name,
        # use the configured fallback template. This keeps personal naming policy
        # out of parser code.
        if instrument_col:
            instrument_name = normalise_text(row.get(instrument_col, ""))
        elif portfolio:
            instrument_name = _instrument_name_from_template(portfolio)
        else:
            instrument_name = _instrument_name_from_template("").strip() or "SELIGSON"

        transaction_number = normalise_text(row.get(transaction_col, ""))

        parsed_row = {
            "Broker": BROKER,
            "Portfolio": portfolio or "SELIGSON",
            "PortfolioOwner": "",
            "PortfolioType": "",
            "BookingDate": format_date(row.get(value_date_col, "")),
            "TradeDate": format_date(row.get(value_date_col, "")),
            "SettlementDate": format_date(row.get(value_date_col, "")),
            "TransactionTypeRaw": _infer_transaction_type_raw(cash_amount),
            "InstrumentName": instrument_name,
            "ISIN": "",
            "Quantity": _optional_amount(row, quantity_col, 0),
            "UnitPrice": _optional_amount(row, unit_price_col, 0),
            "Interest": "",
            "TotalFees": _optional_amount(row, fee_col, 0),
            "TotalFeesCurrency": "EUR",
            "CashAmount": cash_amount,
            "CashCurrency": "EUR",
            "AcquisitionValue": _optional_amount(row, gross_value_col, 0),
            "AcquisitionCurrency": "EUR",
            "Result": "",
            "ResultCurrency": "EUR",
            "TotalQuantity": "",
            "CashBalance": "",
            "ExchangeRate": 1,
            "Description": f"Seligson transaction {transaction_number}",
            "CancellationDate": "",
            "CalculationID": "",
            "ConfirmationNumber": transaction_number,
            "BrokerageFee": _optional_amount(row, fee_col, 0),
            "BrokerageFeeCurrency": "EUR",
            "ReferenceExchangeRate": "",
            "OriginalInterest": "",
            "ExportDate": export_date,
            "ImportedAt": imported_at,
            "SourceFile": path.name,
            "SeligsonTransactionNumber": transaction_number,
        }

        parsed_row["InvestmentRawID"] = make_investment_raw_id(pd.Series(parsed_row))
        rows.append({column: parsed_row.get(column, "") for column in INVESTMENT_RAW_COLUMNS})

    return pd.DataFrame(rows, columns=INVESTMENT_RAW_COLUMNS)


def parse_broker_raw_rows(path: Path, imported_at: str) -> pd.DataFrame:
    df = _read_seligson_file(path).copy()
    colmap = _column_map(df)

    df.insert(0, "BrokerRawExportID", [
        make_broker_raw_export_id(row, path.name, colmap) for _, row in df.iterrows()
    ])
    df.insert(1, "Broker", BROKER)
    df.insert(2, "SourceFile", path.name)
    df.insert(3, "ExportDate", file_modified_date(path))
    df.insert(4, "ImportedAt", imported_at)

    return df
