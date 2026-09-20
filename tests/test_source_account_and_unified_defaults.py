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


def test_op_filename_with_generic_tili_prefix_strips_it():
    """
    Real bug found and fixed: a newer account with no explicit
    filename_prefixes settings.yaml entry yet fell through to the
    conservative fallback and got named "TILI <name>" instead of just
    <name> - "Tili" is just the generic Finnish word for "account" in the
    real filename convention ("Tili_<name>_tapahtumat..."), not part of the
    actual account name.
    """
    assert infer_source_account_from_filename(Path("Tili_CHILD_tapahtumat20260703-20260722.csv"), "OP") == "CHILD"


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


def test_raw_to_unified_carries_transaction_type_raw_through(tmp_path):
    """
    Real bug found and fixed: TransactionTypeRaw was already part of
    RAW_COLUMNS and correctly populated at the raw layer (e.g. Norwegian's
    KATEVARAUS), but never carried through to UnifiedTransactions at all -
    a content-based CategoryRules/TransactionRules row needed to key off it
    directly (CashEntries.xlsx's optional Type column for Vinted-sourced
    rows) rather than relying on free-text Description matching.
    """
    raw = pd.DataFrame([{
        "RawID": "VINTED-1",
        "SourceAccount": "CASH",
        "SourceBank": "VINTED",
        "BookingDate": "2026-06-01",
        "ValueDate": "2026-06-01",
        "Amount": 20.0,
        "TransactionTypeRaw": "SALE",
        "Description": "Test jacket",
        "RawReceiver": "VINTED",
        "ReceiverAccount": "",
        "ReceiverBankBIC": "",
        "Reference": "",
        "Message": "",
        "ArchiveID": "",
        "Balance": "",
        "CurrencyAmount": "",
        "Currency": "EUR",
        "Rate": 1,
        "MerchantArea": "",
        "MerchantCategory": "",
        "ExportDate": "2026-06-01",
        "ImportedAt": "2026-06-04 12:00:00",
        "SourceFile": "CashEntries.xlsx",
    }])

    for col in RAW_COLUMNS:
        if col not in raw.columns:
            raw[col] = ""

    unified = raw_to_unified_rows(raw[RAW_COLUMNS])

    assert unified.iloc[0]["TransactionTypeRaw"] == "SALE"


def test_raw_to_unified_spankki_include_default_from_settings():
    raw = pd.DataFrame([{
        "RawID": "SPK-1",
        "SourceAccount": "SPANKKI",
        "SourceBank": "SPANKKI",
        "BookingDate": "2026-06-01",
        "ValueDate": "2026-06-01",
        "Amount": -4.50,
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
