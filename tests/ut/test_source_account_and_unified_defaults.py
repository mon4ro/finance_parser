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


def test_raw_to_unified_does_not_carry_balance_or_transaction_type_raw_through():
    """
    Balance and TransactionTypeRaw were briefly carried through to
    UnifiedTransactions (an earlier version of this test asserted exactly
    that), but both turned out to have zero real consumers there:
    build_monthly_account_balance.py's Balance use and its KATEVARAUS
    TransactionTypeRaw filter both read from RawTransactions, not
    UnifiedTransactions (see load_raw_transactions() in
    account_balance_seed.py) - confirmed via a full grep before removing
    them 2026-09-22. Both remain part of RAW_COLUMNS, correctly populated
    there. This test guards against silently re-adding either without a
    real UnifiedTransactions-side consumer that actually needs it.
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

    assert "Balance" not in unified.columns
    assert "TransactionTypeRaw" not in unified.columns


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
