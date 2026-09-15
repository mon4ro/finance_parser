from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from finance_parser.investments.investment_common import (
    INSTRUMENT_MASTER_COLUMNS,
    INVESTMENT_TRANSACTIONS_COLUMNS,
    apply_instrument_master,
    load_instrument_master,
)


def _write_instrument_master(path: Path, rows: list[dict]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "InstrumentMaster"
    ws.append(INSTRUMENT_MASTER_COLUMNS)
    for row in rows:
        ws.append([row.get(h, "") for h in INSTRUMENT_MASTER_COLUMNS])
    wb.save(path)


def _transaction_row(**overrides) -> pd.DataFrame:
    row = {col: "" for col in INVESTMENT_TRANSACTIONS_COLUMNS}
    row.update(overrides)
    return pd.DataFrame([row], columns=INVESTMENT_TRANSACTIONS_COLUMNS)


def test_broker_pinned_master_row_still_matches_a_different_broker(tmp_path):
    """
    Real bug: a NOKIA master row entered with Broker=NORDNET never matched
    NOKIA transactions from EVLI, because the old matching mask required
    the broker to be blank ("") to count as a wildcard - a non-blank,
    different-broker row was excluded entirely. Confirmed against real data
    (481 rows across NOKIA/OP-AASIA/SELIGSON/etc. with no classification)
    that this was never a deliberate broker-specific disambiguation: a scan
    of the whole real InstrumentMaster found zero RawInstrumentName values
    legitimately meaning two different instruments across brokers.
    """
    master_path = tmp_path / "InstrumentMaster.xlsx"
    _write_instrument_master(master_path, [
        {"Enabled": "YES", "Broker": "NORDNET", "RawInstrumentName": "Nokia", "NormalizedInstrument": "NOKIA", "InstrumentType": "STOCK", "AssetClass": "Equity", "Currency": "EUR", "ISIN": "FI0009000681"},
    ])
    master = load_instrument_master(master_path)

    evli_row = _transaction_row(Broker="EVLI", InstrumentName="NOKIA", NormalizedInstrument="NOKIA")

    result = apply_instrument_master(evli_row, master)

    assert result.iloc[0]["InstrumentType"] == "STOCK"
    assert result.iloc[0]["AssetClass"] == "Equity"
    assert result.iloc[0]["InstrumentCurrency"] == "EUR"
    assert result.iloc[0]["ISIN"] == "FI0009000681"


def test_broker_specific_row_still_preferred_when_both_exist(tmp_path):
    master_path = tmp_path / "InstrumentMaster.xlsx"
    _write_instrument_master(master_path, [
        {"Enabled": "YES", "Broker": "NORDNET", "RawInstrumentName": "Nokia", "NormalizedInstrument": "NOKIA", "InstrumentType": "STOCK", "Currency": "EUR"},
        {"Enabled": "YES", "Broker": "EVLI", "RawInstrumentName": "Nokia", "NormalizedInstrument": "NOKIA", "InstrumentType": "STOCK", "Currency": "USD", "Notes": "hypothetical EVLI-specific override"},
    ])
    master = load_instrument_master(master_path)

    evli_row = _transaction_row(Broker="EVLI", InstrumentName="Nokia", NormalizedInstrument="NOKIA")

    result = apply_instrument_master(evli_row, master)

    assert result.iloc[0]["InstrumentCurrency"] == "USD"


def test_share_classes_with_distinct_raw_names_are_not_conflated(tmp_path):
    """
    Real case: Kesko A (Nordnet) and Kesko B (OP) are genuinely different
    instruments. Their raw names differ ("Kesko A" vs "KESKO OYJ B"), so the
    broker-agnostic fallback must not merge them just because they're both
    "Kesko".
    """
    master_path = tmp_path / "InstrumentMaster.xlsx"
    _write_instrument_master(master_path, [
        {"Enabled": "YES", "Broker": "NORDNET", "RawInstrumentName": "Kesko A", "NormalizedInstrument": "KESKO A", "InstrumentType": "STOCK"},
        {"Enabled": "YES", "Broker": "", "RawInstrumentName": "KESKO OYJ B", "NormalizedInstrument": "KESKO OYJ B", "InstrumentType": "STOCK"},
    ])
    master = load_instrument_master(master_path)

    op_row = _transaction_row(Broker="OP", InstrumentName="KESKO OYJ B", NormalizedInstrument="KESKO OYJ B")

    result = apply_instrument_master(op_row, master)

    assert result.iloc[0]["NormalizedInstrument"] == "KESKO OYJ B"


def test_isin_is_backfilled_from_master_when_transaction_row_lacks_one(tmp_path):
    """ISIN was previously never backfilled from InstrumentMaster - only ever
    whatever the raw broker export happened to include natively."""
    master_path = tmp_path / "InstrumentMaster.xlsx"
    _write_instrument_master(master_path, [
        {"Enabled": "YES", "Broker": "EVLI", "RawInstrumentName": "NOKIA", "NormalizedInstrument": "NOKIA", "ISIN": "FI0009000681"},
    ])
    master = load_instrument_master(master_path)

    row = _transaction_row(Broker="EVLI", InstrumentName="NOKIA", NormalizedInstrument="NOKIA", ISIN="")

    result = apply_instrument_master(row, master)

    assert result.iloc[0]["ISIN"] == "FI0009000681"


def test_existing_isin_on_transaction_row_is_not_overwritten(tmp_path):
    master_path = tmp_path / "InstrumentMaster.xlsx"
    _write_instrument_master(master_path, [
        {"Enabled": "YES", "Broker": "NORDNET", "RawInstrumentName": "Nokia", "NormalizedInstrument": "NOKIA", "ISIN": "WRONG-ISIN"},
    ])
    master = load_instrument_master(master_path)

    row = _transaction_row(Broker="NORDNET", InstrumentName="Nokia", NormalizedInstrument="NOKIA", ISIN="FI0009000681")

    result = apply_instrument_master(row, master)

    assert result.iloc[0]["ISIN"] == "FI0009000681"
