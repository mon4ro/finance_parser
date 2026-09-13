from __future__ import annotations

import csv
import io
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


BROKER = "NORDNET"
BROKER_RAW_SHEET = "NordnetRawExport"
# Raw investment ID prefix: NDN-

NORDNET_REQUIRED_COLUMNS = [
    "Id",
    "Kirjauspäivä",
    "Kauppapäivä",
    "Maksupäivä",
    "Salkku",
    "Tapahtumatyyppi",
    "Arvopaperi",
    "ISIN",
    "Määrä",
    "Kurssi",
    "Korko",
    "Kokonaiskulut",
    "Summa",
    "Hankinta-arvo",
    "Tulos",
    "Kokonaismäärä",
    "Saldo",
    "Vaihtokurssi",
    "Tapahtumateksti",
    "Mitätöintipäivä",
    "Laskelma",
    "Vahvistusnumero",
    "Välityspalkkio",
    "Viitevaluuttakurssi",
    "Alkuperäinen korko",
]

HEADER_SENTINEL_COLUMNS = [
    "Id",
    "Kirjauspäivä",
    "Kauppapäivä",
    "Tapahtumatyyppi",
    "Arvopaperi",
    "ISIN",
]


def _base_column_name(column: object) -> str:
    """
    Pandas disambiguates duplicate headers as Valuutta, Valuutta.1, etc.
    This returns the original logical header name.
    """
    text = normalise_header(column)
    if "." in text:
        prefix, suffix = text.rsplit(".", 1)
        if suffix.isdigit():
            return prefix
    return text


def _deduplicate_headers(headers: list[object]) -> list[str]:
    """
    Match pandas' duplicate-header style for manual header promotion:
    Valuutta, Valuutta.1, Valuutta.2, ...
    """
    out = []
    seen: dict[str, int] = {}

    for raw_header in headers:
        header = normalise_header(raw_header)
        if not header:
            header = f"BlankColumn{len(out) + 1}"

        count = seen.get(header, 0)
        if count == 0:
            out.append(header)
        else:
            out.append(f"{header}.{count}")

        seen[header] = count + 1

    return out


def _columns_by_base(df: pd.DataFrame, base_name: str) -> list[str]:
    return [c for c in df.columns if _base_column_name(c) == base_name]


def _currency_columns(df: pd.DataFrame) -> list[str]:
    return _columns_by_base(df, "Valuutta")


def _currency(row: pd.Series, currency_cols: list[str], index: int, default: str = "EUR") -> str:
    if index < len(currency_cols):
        value = normalise_text(row.get(currency_cols[index], ""))
        if value:
            return value
    return default


def _optional_amount(row: pd.Series, column: str) -> object:
    value = normalise_text(row.get(column, ""))
    if value == "":
        return ""
    return normalise_amount(value)


def _has_required_column_set(df: pd.DataFrame) -> bool:
    columns = set(_base_column_name(c) for c in df.columns)
    return set(NORDNET_REQUIRED_COLUMNS).issubset(columns)


def _header_score(values: list[object]) -> int:
    headers = set(normalise_header(v) for v in values)
    return len(headers.intersection(set(HEADER_SENTINEL_COLUMNS)))


def _promote_header_row_if_needed(df: pd.DataFrame) -> pd.DataFrame:
    """
    Some Nordnet exports contain one or more metadata / title rows before the
    actual table header. If pandas read the wrong row as headers, scan the first
    rows for the real header and promote it.
    """
    if _has_required_column_set(df):
        return df

    if df.empty:
        return df

    best_idx = None
    best_score = 0

    scan_limit = min(len(df), 30)
    for idx in range(scan_limit):
        row_values = list(df.iloc[idx].values)
        score = _header_score(row_values)
        if score > best_score:
            best_score = score
            best_idx = idx

    # Require several clear Nordnet header fields before promoting.
    if best_idx is None or best_score < 4:
        return df

    headers = _deduplicate_headers(list(df.iloc[best_idx].values))
    promoted = df.iloc[best_idx + 1:].copy()
    promoted.columns = headers

    # Drop fully blank rows after the promoted header.
    promoted = promoted.loc[
        ~promoted.apply(lambda r: all(normalise_text(v) == "" for v in r), axis=1)
    ].copy()

    return promoted


def _read_delimited_with_header_scan(path: Path) -> pd.DataFrame:
    """
    Fallback reader for Nordnet CSV files whose real header is not the first row.
    It scans decoded text for a row containing key Nordnet headers, then builds
    a DataFrame from the following rows.
    """
    raw = path.read_bytes()

    encodings = [
        "utf-8-sig",
        "utf-8",
        "cp1252",
        "iso-8859-15",
        "latin1",
        "utf-16",
        "utf-16-le",
        "utf-16-be",
    ]
    delimiters = [";", "\t", ",", "|"]

    attempts = []

    for encoding in encodings:
        try:
            text = raw.decode(encoding)
        except UnicodeDecodeError as exc:
            attempts.append(f"encoding={encoding} failed decode: {exc}")
            continue

        text = text.replace("\x00", "")
        text = text.replace("\r\n", "\n").replace("\r", "\n")

        for sep in delimiters:
            for quoting in [csv.QUOTE_MINIMAL, csv.QUOTE_NONE]:
                try:
                    reader = csv.reader(
                        io.StringIO(text),
                        delimiter=sep,
                        quotechar='"',
                        quoting=quoting,
                    )
                    rows = [row for row in reader if any(normalise_text(v) for v in row)]
                except Exception as exc:
                    attempts.append(
                        f"encoding={encoding}, sep={repr(sep)}, quoting={quoting} failed: {exc}"
                    )
                    continue

                best_idx = None
                best_score = 0

                for idx, row in enumerate(rows[:80]):
                    score = _header_score(row)
                    if score > best_score:
                        best_score = score
                        best_idx = idx

                attempts.append(
                    f"encoding={encoding}, sep={repr(sep)}, quoting={quoting} -> "
                    f"rows={len(rows)}, best_header_score={best_score}, best_idx={best_idx}"
                )

                if best_idx is None or best_score < 4:
                    continue

                headers = _deduplicate_headers(rows[best_idx])
                data_rows = rows[best_idx + 1:]

                width = len(headers)
                fixed_rows = []
                for row in data_rows:
                    if len(row) < width:
                        row = row + [""] * (width - len(row))
                    elif len(row) > width:
                        row = row[:width]
                    fixed_rows.append(row)

                df = pd.DataFrame(fixed_rows, columns=headers)
                df.columns = [normalise_header(c) for c in df.columns]

                if _has_required_column_set(df):
                    print(
                        f"Read {path.name} as Nordnet delimited text with header scan: "
                        f"encoding={encoding}, delimiter={repr(sep)}, quoting={quoting}, header_row={best_idx + 1}"
                    )
                    return df

    raise ValueError(
        "Could not find Nordnet header row. Recent attempts:\n"
        + "\n".join(f"  - {a}" for a in attempts[-30:])
    )


def read_nordnet_file(path: Path) -> pd.DataFrame:
    """
    Read Nordnet exports robustly.

    First use the project's shared reader. If that reads the file but chooses a
    pre-header row as column names, promote the actual Nordnet header row. If the
    shared reader fails entirely, use a Nordnet-specific header-scan fallback.
    """
    shared_error: Exception | None = None

    try:
        df = read_input_file(path, required_columns=NORDNET_REQUIRED_COLUMNS)
        df.columns = [normalise_header(c) for c in df.columns]
        df = _promote_header_row_if_needed(df)

        if _has_required_column_set(df):
            return df

    except Exception as exc:
        shared_error = exc

    try:
        return _read_delimited_with_header_scan(path)
    except Exception as fallback_error:
        if shared_error is not None:
            raise ValueError(
                f"Shared reader failed or found wrong header: {shared_error}\n"
                f"Nordnet header-scan fallback also failed: {fallback_error}"
            )
        raise fallback_error


def can_parse(path: Path) -> tuple[bool, str]:
    try:
        df = read_nordnet_file(path)
        columns = set(_base_column_name(c) for c in df.columns)
        missing = [c for c in NORDNET_REQUIRED_COLUMNS if c not in columns]
        if missing:
            return False, f"Missing Nordnet columns: {', '.join(missing[:6])}"
        return True, "Matched Nordnet column set"
    except Exception as exc:
        return False, str(exc)


def make_investment_raw_id(row: pd.Series) -> str:
    """
    Stable hash-based ID.

    Nordnet's Id is expected to be stable and unique, but the hash also includes
    broker and portfolio so the public output ID stays format-consistent.
    """
    return "NDN-" + stable_hash([
        BROKER,
        row.get("Portfolio", ""),
        row.get("NordnetID", ""),
        row.get("TradeDate", ""),
        row.get("TransactionTypeRaw", ""),
        row.get("InstrumentName", ""),
        row.get("ISIN", ""),
        row.get("Quantity", ""),
        row.get("CashAmount", ""),
        row.get("ConfirmationNumber", ""),
    ], length=16)


def parse_file(path: Path, imported_at: str) -> pd.DataFrame:
    df = read_nordnet_file(path)

    base_columns = set(_base_column_name(c) for c in df.columns)
    missing = [c for c in NORDNET_REQUIRED_COLUMNS if c not in base_columns]
    if missing:
        available = ", ".join(map(str, df.columns))
        raise ValueError(
            f"{path.name} is missing required Nordnet column(s): {', '.join(missing)}\n"
            f"Available columns detected: {available}"
        )

    df = df.copy()
    df = df[df["Id"].notna()]
    df = df[df["Id"].astype(str).str.strip() != ""]
    df = df[df["Kauppapäivä"].notna()]
    df = df[df["Kauppapäivä"].astype(str).str.strip() != ""]

    currency_cols = _currency_columns(df)
    export_date = file_modified_date(path)

    rows: list[dict[str, object]] = []

    for _, row in df.iterrows():
        # Nordnet export contains repeated "Valuutta" columns in the sample:
        # 0 = Kokonaiskulut currency
        # 1 = Summa currency
        # 2 = Hankinta-arvo currency
        # 3 = Tulos currency
        # 4 = Välityspalkkio currency
        total_fees_currency = _currency(row, currency_cols, 0)
        cash_currency = _currency(row, currency_cols, 1)
        acquisition_currency = _currency(row, currency_cols, 2, cash_currency)
        result_currency = _currency(row, currency_cols, 3, cash_currency)
        brokerage_fee_currency = _currency(row, currency_cols, 4, cash_currency)

        parsed_row = {
            "Broker": BROKER,
            "Portfolio": normalise_text(row.get("Salkku", "")),
            # Owner and account type should later come from a PortfolioMaster
            # or rules file. Do not infer them from the portfolio number yet.
            "PortfolioOwner": "",
            "PortfolioType": "",
            "BookingDate": format_date(row.get("Kirjauspäivä", "")),
            "TradeDate": format_date(row.get("Kauppapäivä", "")),
            "SettlementDate": format_date(row.get("Maksupäivä", "")),
            "TransactionTypeRaw": normalise_text(row.get("Tapahtumatyyppi", "")),
            "InstrumentName": normalise_text(row.get("Arvopaperi", "")),
            "ISIN": normalise_text(row.get("ISIN", "")),
            "Quantity": normalise_amount(row.get("Määrä", 0)),
            "UnitPrice": normalise_amount(row.get("Kurssi", 0)),
            "Interest": normalise_amount(row.get("Korko", 0)),
            "TotalFees": normalise_amount(row.get("Kokonaiskulut", 0)),
            "TotalFeesCurrency": total_fees_currency,
            "CashAmount": normalise_amount(row.get("Summa", 0)),
            "CashCurrency": cash_currency,
            "AcquisitionValue": normalise_amount(row.get("Hankinta-arvo", 0)),
            "AcquisitionCurrency": acquisition_currency,
            "Result": normalise_amount(row.get("Tulos", 0)),
            "ResultCurrency": result_currency,
            "TotalQuantity": normalise_amount(row.get("Kokonaismäärä", 0)),
            "CashBalance": normalise_amount(row.get("Saldo", 0)),
            "ExchangeRate": _optional_amount(row, "Vaihtokurssi"),
            "Description": normalise_text(row.get("Tapahtumateksti", "")),
            "CancellationDate": format_date(row.get("Mitätöintipäivä", "")),
            "CalculationID": normalise_text(row.get("Laskelma", "")),
            "ConfirmationNumber": normalise_text(row.get("Vahvistusnumero", "")),
            "BrokerageFee": normalise_amount(row.get("Välityspalkkio", 0)),
            "BrokerageFeeCurrency": brokerage_fee_currency,
            "ReferenceExchangeRate": _optional_amount(row, "Viitevaluuttakurssi"),
            "OriginalInterest": _optional_amount(row, "Alkuperäinen korko"),
            "ExportDate": export_date,
            "ImportedAt": imported_at,
            "SourceFile": path.name,
            # Helper used only for ID generation, removed before output.
            "NordnetID": normalise_text(row.get("Id", "")),
        }

        parsed_row["InvestmentRawID"] = make_investment_raw_id(pd.Series(parsed_row))

        rows.append({column: parsed_row.get(column, "") for column in INVESTMENT_RAW_COLUMNS})

    return pd.DataFrame(rows, columns=INVESTMENT_RAW_COLUMNS)



def make_broker_raw_export_id(row: pd.Series, source_file: str) -> str:
    return "NORDNETRAW-" + stable_hash([
        BROKER,
        source_file,
        row.get("Id", ""),
        row.get("Salkku", ""),
        row.get("Kauppapäivä", ""),
        row.get("Tapahtumatyyppi", ""),
        row.get("Arvopaperi", ""),
        row.get("ISIN", ""),
        row.get("Määrä", ""),
        row.get("Summa", ""),
        row.get("Vahvistusnumero", ""),
    ], length=16)


def parse_broker_raw_rows(path: Path, imported_at: str) -> pd.DataFrame:
    """
    Preserve a closer-to-original Nordnet export sheet in ParsedInvestments.xlsx.

    This is separate from RawInvestmentTransactions, which is the canonical
    cross-broker raw schema.
    """
    df = read_nordnet_file(path).copy()

    df.insert(0, "BrokerRawExportID", [
        make_broker_raw_export_id(row, path.name) for _, row in df.iterrows()
    ])
    df.insert(1, "Broker", BROKER)
    df.insert(2, "SourceFile", path.name)
    df.insert(3, "ExportDate", file_modified_date(path))
    df.insert(4, "ImportedAt", imported_at)

    return df
