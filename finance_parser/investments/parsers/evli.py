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


BROKER = "EVLI"
BROKER_RAW_SHEET = "EvliRawExport"
# Raw investment ID prefix: EVL-

# EVLI export syntax:
# Instrument | Event type | Quantity | bonusPercent | amount | amountLocal | Value | Date | Status
#
# The export column named "Instrument" is actually a plan/program label (e.g.
# "Plan Cycle 2022", "SIS Dividend"), not the traded instrument or a real
# per-account identifier - EVLI is one real account holding one real
# instrument (the employee stock plan's own stock), never multiple. Portfolio
# is therefore hardcoded to the broker name, matching how OP's parser already
# does the same thing for its own account-less export. The plan label is kept
# as its own PlanCycle raw-data column instead of being folded into Portfolio
# (real bug this replaced: treating each plan cycle as a separate "portfolio"
# fragmented one real holding's cumulative quantity across 5 fake groups).

EVLI_REQUIRED_COLUMNS = [
    "Instrument",
    "Event type",
    "Date",
    "Status",
]

EVLI_OPTIONAL_COLUMNS = [
    "Quantity",
    "bonusPercent",
    "amount",
    "amountLocal",
    "Value",
]


def _normalised_columns(df: pd.DataFrame) -> dict[str, str]:
    """
    Map uppercase normalised column names to actual DataFrame column names.
    """
    return {normalise_header(c).upper(): c for c in df.columns}


def _get_column(df: pd.DataFrame, name: str) -> str | None:
    return _normalised_columns(df).get(normalise_header(name).upper())


def _require_columns(df: pd.DataFrame) -> dict[str, str | None]:
    columns = {name: _get_column(df, name) for name in EVLI_REQUIRED_COLUMNS + EVLI_OPTIONAL_COLUMNS}

    missing = [name for name in EVLI_REQUIRED_COLUMNS if not columns.get(name)]
    if missing:
        raise ValueError(
            f"Missing EVLI columns: {', '.join(missing)}. "
            f"Available columns: {', '.join(map(str, df.columns))}"
        )

    return columns


def _read_evli_file(path: Path) -> pd.DataFrame:
    df = read_input_file(path, required_columns=EVLI_REQUIRED_COLUMNS)
    df.columns = [normalise_header(c) for c in df.columns]
    _require_columns(df)
    return df


def can_parse(path: Path) -> tuple[bool, str]:
    try:
        df = _read_evli_file(path)
        return True, "Matched EVLI column set"
    except Exception as exc:
        return False, str(exc)


def _optional_amount(row: pd.Series, column: str | None, default: object = "") -> object:
    if column is None:
        return default

    value = normalise_text(row.get(column, ""))
    if value == "":
        return default

    return normalise_amount(value)


def _cash_amount(row: pd.Series, amount_col: str | None, amount_local_col: str | None) -> float:
    """
    Prefer `amountLocal` if present, otherwise `amount`.

    Sell rows in the observed export may have blank amount fields and only a
    Value/unit-price field. In that case we leave cash amount as 0 instead of
    inventing proceeds from Quantity × Value, because the export also contains
    separate cash transfer rows and may omit fees/taxes.
    """
    local = _optional_amount(row, amount_local_col, "")
    if local != "":
        return float(local)

    amount = _optional_amount(row, amount_col, "")
    if amount != "":
        return float(amount)

    return 0.0




def _fixed_instrument_name() -> str:
    """
    Return the configured canonical EVLI instrument name.

    Example:
      investments.broker_defaults.EVLI.fixed_instrument_name = NOKIA
    """
    return str(get_settings().broker_default(BROKER, "fixed_instrument_name", default="EVLI")).strip() or "EVLI"


def _plan_cycle_from_export_instrument(value: object) -> str:
    """
    The EVLI export column named `Instrument` is actually a plan/program
    label (e.g. "Plan Cycle 2022", "SIS Dividend") - captured verbatim as
    PlanCycle, not used as Portfolio (see module docstring above).
    """
    return normalise_text(value)


def make_investment_raw_id(row: pd.Series) -> str:
    # Portfolio is now a constant ("EVLI"), so PlanCycle is what distinguishes
    # otherwise-identical rows from different plan cycles - it must be in the
    # hash, not Portfolio, or two coincidentally-identical rows from
    # different cycles would collide onto the same ID.
    return "EVL-" + stable_hash([
        BROKER,
        row.get("PlanCycle", ""),
        row.get("ValueDate", ""),
        row.get("TransactionTypeRaw", ""),
        row.get("Quantity", ""),
        row.get("UnitPrice", ""),
        row.get("CashAmount", ""),
        row.get("Status", ""),
    ], length=16)


def make_broker_raw_export_id(row: pd.Series, source_file: str, columns: dict[str, str | None]) -> str:
    return "EVLIRAW-" + stable_hash([
        BROKER,
        source_file,
        row.get(columns.get("Instrument") or "", ""),
        row.get(columns.get("Event type") or "", ""),
        row.get(columns.get("Quantity") or "", ""),
        row.get(columns.get("amount") or "", ""),
        row.get(columns.get("amountLocal") or "", ""),
        row.get(columns.get("Value") or "", ""),
        row.get(columns.get("Date") or "", ""),
        row.get(columns.get("Status") or "", ""),
    ], length=16)


def parse_file(path: Path, imported_at: str) -> pd.DataFrame:
    df = _read_evli_file(path).copy()
    columns = _require_columns(df)

    instrument_col = columns["Instrument"]
    event_col = columns["Event type"]
    quantity_col = columns.get("Quantity")
    amount_col = columns.get("amount")
    amount_local_col = columns.get("amountLocal")
    value_col = columns.get("Value")
    date_col = columns["Date"]
    status_col = columns["Status"]

    fixed_instrument_name = _fixed_instrument_name()

    df = df[df[date_col].notna()]
    df = df[df[date_col].astype(str).str.strip() != ""]
    df = df[df[event_col].notna()]
    df = df[df[event_col].astype(str).str.strip() != ""]

    export_date = file_modified_date(path)

    rows: list[dict[str, object]] = []

    for _, row in df.iterrows():
        plan_cycle = _plan_cycle_from_export_instrument(row.get(instrument_col, ""))
        event_type = normalise_text(row.get(event_col, ""))
        status = normalise_text(row.get(status_col, ""))

        quantity = _optional_amount(row, quantity_col, 0)
        unit_price = _optional_amount(row, value_col, 0)
        cash_amount = _cash_amount(row, amount_col, amount_local_col)

        parsed_row = {
            "Broker": BROKER,
            "Portfolio": BROKER,
            "PlanCycle": plan_cycle,
            "PortfolioOwner": "",
            "PortfolioType": "",
            "BookingDate": format_date(row.get(date_col, "")),
            "TradeDate": format_date(row.get(date_col, "")),
            "SettlementDate": format_date(row.get(date_col, "")),
            "TransactionTypeRaw": event_type,
            "InstrumentName": fixed_instrument_name,
            "ISIN": "",
            "Quantity": quantity,
            "UnitPrice": unit_price,
            "Interest": "",
            "TotalFees": "",
            "TotalFeesCurrency": "EUR",
            "CashAmount": cash_amount,
            "CashCurrency": "EUR",
            "AcquisitionValue": "",
            "AcquisitionCurrency": "EUR",
            "Result": "",
            "ResultCurrency": "EUR",
            "TotalQuantity": "",
            "CashBalance": "",
            "ExchangeRate": 1,
            "Description": f"{plan_cycle} | {event_type} | {status}",
            "CancellationDate": "",
            "CalculationID": "",
            "ConfirmationNumber": "",
            "BrokerageFee": "",
            "BrokerageFeeCurrency": "EUR",
            "ReferenceExchangeRate": "",
            "OriginalInterest": "",
            "ExportDate": export_date,
            "ImportedAt": imported_at,
            "SourceFile": path.name,
            "Status": status,
            "ValueDate": format_date(row.get(date_col, "")),
        }

        parsed_row["InvestmentRawID"] = make_investment_raw_id(pd.Series(parsed_row))
        rows.append({column: parsed_row.get(column, "") for column in INVESTMENT_RAW_COLUMNS})

    return pd.DataFrame(rows, columns=INVESTMENT_RAW_COLUMNS)


def parse_broker_raw_rows(path: Path, imported_at: str) -> pd.DataFrame:
    df = _read_evli_file(path).copy()
    columns = _require_columns(df)

    df.insert(0, "BrokerRawExportID", [
        make_broker_raw_export_id(row, path.name, columns) for _, row in df.iterrows()
    ])
    df.insert(1, "Broker", BROKER)
    df.insert(2, "SourceFile", path.name)
    df.insert(3, "ExportDate", file_modified_date(path))
    df.insert(4, "ImportedAt", imported_at)

    return df
