from __future__ import annotations

import re
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


BROKER = "OP"
BROKER_RAW_SHEET = "OPInvestmentRawExport"
# Raw investment ID prefix: OPINV-

# OP investment export syntax seen in the real uploaded file:
# LAJI | TAPAHTUMA | TAPAHTUMA- / PÄIVÄMÄÄRÄ | MÄÄRÄ | KURSSI |
# KULUT / KOROT / VEROT | YHTEENSÄ | TOIMINNOT
#
# OP may split some headers across two rows. This parser repairs that header
# shape before detecting/parsing columns.

OP_INVESTMENT_REQUIRED_COLUMNS = [
    "LAJI",
    "TAPAHTUMA",
    "YHTEENSÄ",
]

OP_INVESTMENT_OPTIONAL_COLUMNS = [
    "TAPAHTUMA-PÄIVÄMÄÄRÄ",
    "MÄÄRÄ",
    "KURSSI",
    "KULUT/KOROT/VEROT",
    "TOIMINNOT",
]


def _key(value: object) -> str:
    text = normalise_header(value)
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"\s*/\s*", "/", text)
    text = re.sub(r"\s*-\s*", "-", text)
    return text.upper()


def _normalised_columns(df: pd.DataFrame) -> dict[str, str]:
    return {_key(c): c for c in df.columns}


def _get_column(df: pd.DataFrame, *candidates: str, startswith: str | None = None) -> str | None:
    columns = _normalised_columns(df)
    for candidate in candidates:
        key = _key(candidate)
        if key in columns:
            return columns[key]

    if startswith:
        prefix = _key(startswith)
        for key, original in columns.items():
            if key.startswith(prefix):
                return original

    return None


def _looks_like_continuation(value: object) -> bool:
    text = normalise_text(value)
    if not text:
        return False

    # Header continuations in the real file are strings such as PÄIVÄMÄÄRÄ and
    # VEROT, while real data rows have dates, numbers, etc.
    if re.search(r"\d", text):
        return False

    return True


def _repair_split_header_row(df: pd.DataFrame) -> pd.DataFrame:
    """
    OP's Excel export can split headers over two rows.

    Example after pandas reads the first row as headers:

    columns:
      TAPAHTUMA-
      KULUT / KOROT /

    first data row:
      PÄIVÄMÄÄRÄ
      VEROT

    This function combines those into:
      TAPAHTUMA-PÄIVÄMÄÄRÄ
      KULUT/KOROT/VEROT

    and drops the continuation row.
    """
    if df.empty:
        return df

    out = df.copy()
    first_row = out.iloc[0]

    new_columns: list[str] = []
    changed = False

    for col in out.columns:
        header = normalise_header(col).replace("\u00a0", " ").strip()
        continuation = normalise_header(first_row.get(col, "")).replace("\u00a0", " ").strip()

        if continuation and _looks_like_continuation(continuation):
            header_key = _key(header)
            continuation_key = _key(continuation)

            if header_key == "TAPAHTUMA-" and continuation_key == "PÄIVÄMÄÄRÄ":
                new_columns.append("TAPAHTUMA-PÄIVÄMÄÄRÄ")
                changed = True
                continue

            if header_key in {"KULUT/KOROT/", "KULUT/KOROT"} and continuation_key == "VEROT":
                new_columns.append("KULUT/KOROT/VEROT")
                changed = True
                continue

        new_columns.append(header)

    if changed:
        out = out.iloc[1:].copy()
        out.columns = new_columns
        out = out.reset_index(drop=True)

    out.columns = [
        re.sub(r"\s*/\s*", "/", normalise_header(c).replace("\u00a0", " ").strip())
        for c in out.columns
    ]

    return out


def _require_columns(df: pd.DataFrame) -> dict[str, str | None]:
    columns = {
        "InstrumentName": _get_column(df, "LAJI", "Laji"),
        "TransactionTypeRaw": _get_column(df, "TAPAHTUMA", "Tapahtuma"),
        "TradeDate": _get_column(
            df,
            "TAPAHTUMA-PÄIVÄMÄÄRÄ",
            "Tapahtuma-päivämäärä",
            "Tapahtumapäivämäärä",
            startswith="TAPAHTUMA-PÄIV",
        ),
        "Quantity": _get_column(df, "MÄÄRÄ", "Määrä"),
        "UnitPrice": _get_column(df, "KURSSI", "Kurssi"),
        "FeesInterestTaxes": _get_column(
            df,
            "KULUT/KOROT/VEROT",
            "KULUT / KOROT / VEROT",
            "Kulut/korot/verot",
            startswith="KULUT",
        ),
        "Total": _get_column(df, "YHTEENSÄ", "Yhteensä"),
        "Actions": _get_column(df, "TOIMINNOT", "Toiminnot"),
    }

    missing = [
        name
        for name in ["InstrumentName", "TransactionTypeRaw", "TradeDate", "Total"]
        if not columns.get(name)
    ]

    if missing:
        raise ValueError(
            f"Missing OP investment columns: {', '.join(missing)}. "
            f"Available columns: {', '.join(map(str, df.columns))}"
        )

    return columns


def _read_op_investment_file(path: Path) -> pd.DataFrame:
    # Do not pass the strict split-header column list to the shared reader; OP
    # can split TAPAHTUMA-PÄIVÄMÄÄRÄ and KULUT/KOROT/VEROT across two rows.
    df = read_input_file(path, required_columns=[])
    df.columns = [normalise_header(c).replace("\u00a0", " ").strip() for c in df.columns]
    df = _repair_split_header_row(df)
    _require_columns(df)
    return df


def can_parse(path: Path) -> tuple[bool, str]:
    try:
        df = _read_op_investment_file(path)
        columns = _require_columns(df)
        return True, (
            "Matched OP investment export column set "
            f"(date={columns.get('TradeDate')}, total={columns.get('Total')})"
        )
    except Exception as exc:
        return False, str(exc)


def _extract_number(value: object, default: object = "") -> object:
    text = normalise_text(value)
    if not text:
        return default

    # Values can look like:
    # 348,00 kpl
    # 337,98 EUR
    # 50,00 EUR
    match = re.search(r"[-+]?\d+(?:[ .\u00a0]\d{3})*(?:,\d+)?|[-+]?\d+(?:\.\d+)?", text)
    if not match:
        return default

    return normalise_amount(match.group(0))


def _extract_currency(value: object, default: str = "EUR") -> str:
    text = normalise_text(value).upper()
    for currency in ["EUR", "SEK", "NOK", "USD", "DKK", "GBP", "CHF"]:
        if re.search(rf"\b{currency}\b", text):
            return currency
    return default


def _canonical_event(value: object) -> str:
    return normalise_text(value).upper()


def _cash_amount_signed(event_type_raw: object, total_value: float) -> float:
    """
    OP investment export displays totals as positive values. Convert to
    investment cash-flow signs:

    - OSTO / MERKINTÄ: cash outflow
    - OSINKO: cash inflow
    - VERONPIDÄTYS / taxes: cash outflow
    - MYYNTI / LUNASTUS: cash inflow
    """
    event = _canonical_event(event_type_raw)

    if event in {"OSTO", "MERKINTÄ", "MERKINTA", "VERONPIDÄTYS", "ENNAKONPIDÄTYS", "VERO"}:
        return -abs(total_value)

    if event in {"MYYNTI", "LUNASTUS", "OSINKO"}:
        return abs(total_value)

    return total_value


def make_investment_raw_id(row: pd.Series) -> str:
    return "OPINV-" + stable_hash([
        BROKER,
        row.get("TradeDate", ""),
        row.get("TransactionTypeRaw", ""),
        row.get("InstrumentName", ""),
        row.get("Quantity", ""),
        row.get("UnitPrice", ""),
        row.get("CashAmount", ""),
        row.get("TotalFees", ""),
    ], length=16)


def make_broker_raw_export_id(row: pd.Series, source_file: str, columns: dict[str, str | None]) -> str:
    return "OPINVRAW-" + stable_hash([
        BROKER,
        source_file,
        row.get(columns.get("InstrumentName") or "", ""),
        row.get(columns.get("TransactionTypeRaw") or "", ""),
        row.get(columns.get("TradeDate") or "", ""),
        row.get(columns.get("Quantity") or "", ""),
        row.get(columns.get("UnitPrice") or "", ""),
        row.get(columns.get("FeesInterestTaxes") or "", ""),
        row.get(columns.get("Total") or "", ""),
    ], length=16)


def parse_file(path: Path, imported_at: str) -> pd.DataFrame:
    df = _read_op_investment_file(path).copy()
    columns = _require_columns(df)

    instrument_col = columns["InstrumentName"]
    event_col = columns["TransactionTypeRaw"]
    date_col = columns["TradeDate"]
    quantity_col = columns.get("Quantity")
    unit_price_col = columns.get("UnitPrice")
    fees_col = columns.get("FeesInterestTaxes")
    total_col = columns["Total"]
    actions_col = columns.get("Actions")

    df = df[df[date_col].notna()]
    df = df[df[date_col].astype(str).str.strip() != ""]
    df = df[df[event_col].notna()]
    df = df[df[event_col].astype(str).str.strip() != ""]

    export_date = file_modified_date(path)
    rows: list[dict[str, object]] = []

    for _, row in df.iterrows():
        event_type = normalise_text(row.get(event_col, ""))
        instrument_name = normalise_text(row.get(instrument_col, ""))
        total = float(_extract_number(row.get(total_col, ""), 0))
        cash_amount = _cash_amount_signed(event_type, total)

        unit_price = _extract_number(row.get(unit_price_col, ""), 0) if unit_price_col else 0
        quantity = _extract_number(row.get(quantity_col, ""), 0) if quantity_col else 0
        fees_interest_taxes = _extract_number(row.get(fees_col, ""), 0) if fees_col else 0

        currency = _extract_currency(row.get(total_col, ""), "EUR")
        fees_currency = _extract_currency(row.get(fees_col, ""), currency) if fees_col else currency

        parsed_row = {
            "Broker": BROKER,
            "Portfolio": "OP",
            "PortfolioOwner": "",
            "PortfolioType": "",
            "BookingDate": format_date(row.get(date_col, "")),
            "TradeDate": format_date(row.get(date_col, "")),
            "SettlementDate": format_date(row.get(date_col, "")),
            "TransactionTypeRaw": event_type,
            "InstrumentName": instrument_name,
            "ISIN": "",
            "Quantity": quantity,
            "UnitPrice": unit_price,
            "Interest": "",
            "TotalFees": fees_interest_taxes,
            "TotalFeesCurrency": fees_currency,
            "CashAmount": cash_amount,
            "CashCurrency": currency,
            "AcquisitionValue": abs(cash_amount) if _canonical_event(event_type) in {"OSTO", "MERKINTÄ", "MERKINTA"} else "",
            "AcquisitionCurrency": currency,
            "Result": "",
            "ResultCurrency": currency,
            "TotalQuantity": "",
            "CashBalance": "",
            "ExchangeRate": 1,
            "Description": normalise_text(row.get(actions_col, "")) if actions_col else event_type,
            "CancellationDate": "",
            "CalculationID": "",
            "ConfirmationNumber": "",
            "BrokerageFee": fees_interest_taxes if _canonical_event(event_type) in {"OSTO", "MYYNTI", "MERKINTÄ", "MERKINTA", "LUNASTUS"} else "",
            "BrokerageFeeCurrency": fees_currency,
            "ReferenceExchangeRate": "",
            "OriginalInterest": "",
            "ExportDate": export_date,
            "ImportedAt": imported_at,
            "SourceFile": path.name,
        }

        parsed_row["InvestmentRawID"] = make_investment_raw_id(pd.Series(parsed_row))
        rows.append({column: parsed_row.get(column, "") for column in INVESTMENT_RAW_COLUMNS})

    return pd.DataFrame(rows, columns=INVESTMENT_RAW_COLUMNS)


def parse_broker_raw_rows(path: Path, imported_at: str) -> pd.DataFrame:
    df = _read_op_investment_file(path).copy()
    columns = _require_columns(df)

    df.insert(0, "BrokerRawExportID", [
        make_broker_raw_export_id(row, path.name, columns) for _, row in df.iterrows()
    ])
    df.insert(1, "Broker", BROKER)
    df.insert(2, "SourceFile", path.name)
    df.insert(3, "ExportDate", file_modified_date(path))
    df.insert(4, "ImportedAt", imported_at)

    return df
