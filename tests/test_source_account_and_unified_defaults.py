import pandas as pd
from pathlib import Path

from finance_parser.common import (
    RAW_COLUMNS,
    canonical_source_account,
    infer_source_account_from_filename,
    raw_to_unified_rows,
)


def test_op_filename_source_account_inference():
    assert infer_source_account_from_filename(Path("HOUSEHOLD_tapahtumat20260501-20260603.csv"), "OP") == "HOUSEHOLD"
    assert infer_source_account_from_filename(Path("PERSONAL_tapahtumat20260503-20260603.csv"), "OP") == "PERSONAL"
    assert infer_source_account_from_filename(Path("CHILD_tapahtumat20260503-20260603-2.csv"), "OP") == "CHILD"


def test_canonical_source_account_repairs_polluted_labels():
    assert canonical_source_account("HOUSEHOLD TAPAHTUMAT20260501 20260603", "OP") == "HOUSEHOLD"
    assert canonical_source_account("PERSONAL TAPAHTUMAT20260503 20260603", "OP") == "PERSONAL"
    assert canonical_source_account("CHILD TAPAHTUMAT20260503 20260603 2", "OP") == "CHILD"


def test_raw_to_unified_leaves_normalized_receiver_blank():
    raw = pd.DataFrame([{
        "RawID": "RAW-1",
        "SourceAccount": "HOUSEHOLD",
        "SourceBank": "OP",
        "BookingDate": "2026-06-01",
        "ValueDate": "2026-06-01",
        "Amount": -10.0,
        "TransactionTypeRaw": "KORTTIOSTO",
        "Description": "KORTTIOSTO",
        "RawReceiver": "Prisma",
        "ReceiverAccount": "",
        "ReceiverBankBIC": "",
        "Reference": "",
        "Message": "",
        "ArchiveID": "123",
        "Balance": "",
        "CurrencyAmount": "",
        "Currency": "EUR",
        "Rate": 1,
        "MerchantArea": "",
        "MerchantCategory": "",
        "ExportDate": "2026-06-01",
        "ImportedAt": "2026-06-04 12:00:00",
        "SourceFile": "HOUSEHOLD_tapahtumat20260601.csv",
    }])

    for col in RAW_COLUMNS:
        if col not in raw.columns:
            raw[col] = ""

    unified = raw_to_unified_rows(raw[RAW_COLUMNS])
    row = unified.iloc[0]

    assert row["RawReceiver"] == "Prisma"
    assert row["NormalizedReceiver"] == ""
    assert row["Supercategory"] == ""
    assert row["Category"] == ""
    assert row["Subcategory"] == ""
    assert row["Include"] == "YES"


def test_raw_to_unified_carries_balance_through(tmp_path):
    """
    Real bug found and fixed: Balance (Nordea's real running balance, "Saldo"
    in its raw export) was already part of RAW_COLUMNS and correctly
    populated at the raw layer, but never carried through to
    UnifiedTransactions at all - build_monthly_account_balance.py needs it.
    """
    raw = pd.DataFrame([{
        "RawID": "RAW-1",
        "SourceAccount": "PERSONAL",
        "SourceBank": "NORDEA",
        "BookingDate": "2026-06-01",
        "ValueDate": "2026-06-01",
        "Amount": -10.0,
        "TransactionTypeRaw": "KORTTIOSTO",
        "Description": "KORTTIOSTO",
        "RawReceiver": "Prisma",
        "ReceiverAccount": "",
        "ReceiverBankBIC": "",
        "Reference": "",
        "Message": "",
        "ArchiveID": "123",
        "Balance": 456.78,
        "CurrencyAmount": "",
        "Currency": "EUR",
        "Rate": 1,
        "MerchantArea": "",
        "MerchantCategory": "",
        "ExportDate": "2026-06-01",
        "ImportedAt": "2026-06-04 12:00:00",
        "SourceFile": "PERSONAL_export.csv",
    }])

    for col in RAW_COLUMNS:
        if col not in raw.columns:
            raw[col] = ""

    unified = raw_to_unified_rows(raw[RAW_COLUMNS])

    assert unified.iloc[0]["Balance"] == 456.78


def test_raw_to_unified_spankki_include_default_from_settings():
    raw = pd.DataFrame([{
        "RawID": "SPK-1",
        "SourceAccount": "SPANKKI",
        "SourceBank": "SPANKKI",
        "BookingDate": "2026-06-01",
        "ValueDate": "2026-06-01",
        "Amount": -3.09,
        "TransactionTypeRaw": "KORTTIOSTO",
        "Description": "KORTTIOSTO",
        "RawReceiver": "S-MARKET",
        "ReceiverAccount": "",
        "ReceiverBankBIC": "",
        "Reference": "",
        "Message": "",
        "ArchiveID": "SPKARCH1",
        "Balance": "",
        "CurrencyAmount": "",
        "Currency": "EUR",
        "Rate": 1,
        "MerchantArea": "",
        "MerchantCategory": "",
        "ExportDate": "2026-06-01",
        "ImportedAt": "2026-06-04 12:00:00",
        "SourceFile": "spankki.csv",
    }])

    for col in RAW_COLUMNS:
        if col not in raw.columns:
            raw[col] = ""

    unified = raw_to_unified_rows(raw[RAW_COLUMNS])
    assert unified.iloc[0]["Include"] == "NO"
