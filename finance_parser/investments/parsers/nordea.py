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


BROKER = "NORDEA"
BROKER_RAW_SHEET = "NordeaRawExport"
TRANSACTIONS_SHEET = "Transactions"
# Raw investment ID prefix: NDA-

# Nordea's own "Profit/Loss" Excel export (downloaded from the web UI, one
# sheet named "Transactions") only carries actual transaction rows within
# some history window - real case: the export used to build this parser had
# only SELL rows (2018 onward), with all older purchase history visible only
# in the web UI itself, not downloadable. That older history is backfilled
# separately via InstrumentMaster.xlsx's OpeningPositions sheet (some as
# real dated purchases transcribed from a screenshot, the rest as a single
# best-guess catch-up entry) - this parser only ever reads real rows that
# are actually present in the Excel export.
#
# The export has no instrument/fund name column at all - a single Nordea
# custody account here holds one fund, so the instrument name comes from
# settings (investments.broker_defaults.NORDEA.fixed_instrument_name),
# matching the same pattern already used for EVLI/Seligson's similarly
# account-per-instrument exports.
NORDEA_REQUIRED_COLUMNS = [
    "Pvm",
    "Tapahtumatyyppi",
    "Määrä",
    "Kurssi",
    "Kauppahinta",
]


def _fixed_instrument_name() -> str:
    return str(
        get_settings().broker_default(BROKER, "fixed_instrument_name", default="NORDEA")
    ).strip() or "NORDEA"


def _read_nordea_file(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=TRANSACTIONS_SHEET, dtype=object, engine="openpyxl")
    df.columns = [normalise_header(c) for c in df.columns]
    return df


def can_parse(path: Path) -> tuple[bool, str]:
    try:
        df = _read_nordea_file(path)
        missing = [c for c in NORDEA_REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            return False, f"Missing Nordea columns: {', '.join(missing)}"
        return True, "Matched Nordea investment export column set"
    except Exception as exc:
        return False, str(exc)


def make_investment_raw_id(row: pd.Series) -> str:
    return "NDA-" + stable_hash([
        BROKER,
        row.get("Portfolio", ""),
        row.get("TradeDate", ""),
        row.get("TransactionTypeRaw", ""),
        row.get("Quantity", ""),
        row.get("UnitPrice", ""),
        row.get("CashAmount", ""),
    ], length=16)


def make_broker_raw_export_id(row: pd.Series, source_file: str) -> str:
    return "NORDEARAW-" + stable_hash([
        BROKER,
        source_file,
        row.get("Pvm", ""),
        row.get("Tapahtumatyyppi", ""),
        row.get("Määrä", ""),
        row.get("Kurssi", ""),
        row.get("Kauppahinta", ""),
    ], length=16)


def parse_file(path: Path, imported_at: str) -> pd.DataFrame:
    df = _read_nordea_file(path)

    missing = [c for c in NORDEA_REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"{path.name} is missing required Nordea column(s): {', '.join(missing)}\n"
            f"Available columns detected: {', '.join(map(str, df.columns))}"
        )

    df = df.copy()
    df = df[df["Pvm"].notna()]
    df = df[df["Pvm"].astype(str).str.strip() != ""]

    export_date = file_modified_date(path)
    fixed_instrument_name = _fixed_instrument_name()

    rows: list[dict[str, object]] = []

    for _, row in df.iterrows():
        transaction_type_raw = normalise_text(row.get("Tapahtumatyyppi", ""))
        quantity = normalise_amount(row.get("Määrä", 0))
        unit_price = normalise_amount(row.get("Kurssi", 0))
        # Kauppahinta is already investor-perspective: positive on a sell
        # (money received), and - by the same convention every other row in
        # this export would use - negative on a buy (money paid), matching
        # this project's CashAmount sign convention everywhere else. No
        # sign flip needed here (unlike Seligson's fund-perspective Summa).
        cash_amount = normalise_amount(row.get("Kauppahinta", 0))
        fees = normalise_amount(row.get("Kulut", 0)) if normalise_text(row.get("Kulut", "")) else 0

        parsed_row = {
            "Broker": BROKER,
            "Portfolio": BROKER,
            "PortfolioOwner": "",
            "PortfolioType": "",
            "BookingDate": format_date(row.get("Pvm", "")),
            "TradeDate": format_date(row.get("Pvm", "")),
            "SettlementDate": format_date(row.get("Pvm", "")),
            "TransactionTypeRaw": transaction_type_raw,
            "InstrumentName": fixed_instrument_name,
            "ISIN": "",
            "Quantity": quantity,
            "UnitPrice": unit_price,
            "Interest": "",
            "TotalFees": fees,
            "TotalFeesCurrency": normalise_text(row.get("Perusvaluutta", "")) or "EUR",
            "CashAmount": cash_amount,
            "CashCurrency": normalise_text(row.get("Selvitysvaluutta", "")) or "EUR",
            "AcquisitionValue": "",
            "AcquisitionCurrency": "EUR",
            "Result": normalise_amount(row.get("Realisoitunut tuotto", 0)) if normalise_text(row.get("Realisoitunut tuotto", "")) else "",
            "ResultCurrency": "EUR",
            "TotalQuantity": "",
            "CashBalance": "",
            "ExchangeRate": 1,
            "Description": f"Nordea {transaction_type_raw}",
            "CancellationDate": "",
            "CalculationID": "",
            "ConfirmationNumber": "",
            "BrokerageFee": fees,
            "BrokerageFeeCurrency": "EUR",
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
    df = _read_nordea_file(path).copy()
    df = df[df["Pvm"].notna()]
    df = df[df["Pvm"].astype(str).str.strip() != ""]

    df.insert(0, "BrokerRawExportID", [
        make_broker_raw_export_id(row, path.name) for _, row in df.iterrows()
    ])
    df.insert(1, "Broker", BROKER)
    df.insert(2, "SourceFile", path.name)
    df.insert(3, "ExportDate", file_modified_date(path))
    df.insert(4, "ImportedAt", imported_at)

    return df
