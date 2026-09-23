from __future__ import annotations

from pathlib import Path

import pandas as pd

from finance_parser.common import (
    RAW_COLUMNS,
    file_modified_date,
    format_date,
    normalise_amount,
    normalise_header,
    normalise_text,
    read_input_file,
    stable_hash,
)


SOURCE_BANK = "NORWEGIAN"
BANK_RAW_SHEET = "NorwegianRawExport"
# RawID prefix: NWG-

NORWEGIAN_COLUMNS = [
    "TransactionDate",
    "Text",
    "Type",
    "Currency Amount",
    "Currency Rate",
    "Currency",
    "Amount",
    "Merchant Area",
    "Merchant Category",
    "BookDate",
    "ValueDate",
]

# Optional columns you may add manually to the Norwegian export.
OPTIONAL_METADATA_COLUMNS = [
    "Source",
    "SourceAccount",
    "ExportDate",
]

REQUIRED_COLUMNS = NORWEGIAN_COLUMNS


def can_parse(path: Path) -> tuple[bool, str]:
    try:
        df = read_input_file(path, required_columns=REQUIRED_COLUMNS)
        columns = set(normalise_header(c) for c in df.columns)
        missing = [c for c in REQUIRED_COLUMNS if c not in columns]
        if missing:
            return False, f"Missing Norwegian columns: {', '.join(missing[:6])}"
        return True, "Matched Bank Norwegian column set"
    except Exception as exc:
        return False, str(exc)


def make_raw_id(row: pd.Series) -> str:
    return "NWG-" + stable_hash([
        row.get("SourceAccount", ""),
        format_date(row.get("ValueDate", "")),
        normalise_amount(row.get("Amount", 0)),
        row.get("RawReceiver", ""),
        row.get("TransactionTypeRaw", ""),
        row.get("CurrencyAmount", ""),
        row.get("Currency", ""),
    ])


def metadata_value(df: pd.DataFrame, row_index, preferred: str, fallback: str = "") -> str:
    if preferred in df.columns:
        return normalise_text(df.at[row_index, preferred])
    return fallback


def parse_file(path: Path, imported_at: str) -> pd.DataFrame:
    df = read_input_file(path, required_columns=REQUIRED_COLUMNS)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        available = ", ".join(map(str, df.columns))
        raise ValueError(
            f"{path.name} is missing required Bank Norwegian column(s): {', '.join(missing)}\n"
            f"Available columns detected: {available}"
        )

    df = df.copy()
    df = df[df["Amount"].notna()]
    df = df[df["Amount"].astype(str).str.strip() != ""]
    df = df[df["TransactionDate"].notna()]
    df = df[df["TransactionDate"].astype(str).str.strip() != ""]

    rows = []
    for idx, row in df.iterrows():
        # Bank Norwegian source account is automatic unless you explicitly add
        # SourceAccount/Source to override it.
        source_account = ""
        if "SourceAccount" in df.columns:
            source_account = normalise_text(row.get("SourceAccount", ""))
        if not source_account and "Source" in df.columns:
            source_account = normalise_text(row.get("Source", ""))
        if not source_account:
            source_account = "NORWEGIAN"

        # ExportDate is taken from the file's modified timestamp by default.
        # A manual ExportDate column can still override this if present.
        export_date = file_modified_date(path)
        if "ExportDate" in df.columns:
            manual_export_date = format_date(row.get("ExportDate", ""))
            if manual_export_date:
                export_date = manual_export_date

        # For budgeting, use the purchase/transaction date.
        # We store that in ValueDate because the shared pipeline uses ValueDate
        # as the UnifiedTransactions Date.
        transaction_date = format_date(row.get("TransactionDate", ""))
        book_date = format_date(row.get("BookDate", ""))

        rows.append({
            "SourceAccount": source_account,
            "SourceBank": SOURCE_BANK,
            "ExportDate": export_date,
            "ImportedAt": imported_at,
            "BookingDate": book_date,
            "ValueDate": transaction_date,
            "Amount": normalise_amount(row.get("Amount", 0)),
            "TransactionTypeRaw": normalise_text(row.get("Type", "")),
            # For Norwegian, Type is the best equivalent of OP's Selitys-like
            # transaction descriptor. The merchant/payee text is RawReceiver.
            "Description": normalise_text(row.get("Type", "")),
            "RawReceiver": normalise_text(row.get("Text", "")),
            "ReceiverAccount": "",
            "ReceiverBankBIC": "",
            "Reference": "",
            "Message": "",
            "ArchiveID": "",
            "Balance": "",
            "CurrencyAmount": normalise_amount(row.get("Currency Amount", 0)),
            "Currency": normalise_text(row.get("Currency", "")) or "EUR",
            "Rate": normalise_amount(row.get("Currency Rate", 1)) or 1,
            "MerchantArea": normalise_text(row.get("Merchant Area", "")),
            "MerchantCategory": normalise_text(row.get("Merchant Category", "")),
            "SourceFile": path.name,
            "Owner": "",
        })

    out = pd.DataFrame(rows)
    out["RawID"] = out.apply(make_raw_id, axis=1)

    return out[RAW_COLUMNS]



def source_account_for_raw_export(df: pd.DataFrame) -> pd.Series:
    values = []
    for _, row in df.iterrows():
        source_account = ""
        if "SourceAccount" in df.columns:
            source_account = normalise_text(row.get("SourceAccount", ""))
        if not source_account and "Source" in df.columns:
            source_account = normalise_text(row.get("Source", ""))
        if not source_account:
            source_account = "NORWEGIAN"
        values.append(source_account)
    return pd.Series(values, index=df.index)


def make_bank_raw_export_id(row: pd.Series, source_file: str = "") -> str:
    """
    Stable Bank Norwegian raw-export row identity.

    Deliberately excludes SourceFile / ExportDate / ImportedAt / inferred
    SourceAccount. Norwegian exports do not appear to contain a bank-native
    archive ID, so use a transaction-content composite.
    """
    return "NWGRAW-" + stable_hash([
        SOURCE_BANK,
        row.get("TransactionDate", ""),
        row.get("BookDate", ""),
        row.get("ValueDate", ""),
        row.get("Amount", ""),
        row.get("Text", ""),
        row.get("Type", ""),
        row.get("Currency Amount", ""),
        row.get("Currency", ""),
    ], length=16)


def parse_bank_raw_rows(path: Path, imported_at: str) -> pd.DataFrame:
    """
    Preserve a closer-to-original Bank Norwegian export sheet in ParsedTransactions.xlsx.
    """
    df = read_input_file(path, required_columns=REQUIRED_COLUMNS).copy()
    df.columns = [normalise_header(c) for c in df.columns]

    df = df.loc[
        ~df.apply(lambda r: all(normalise_text(v) == "" for v in r), axis=1)
    ].copy()

    source_accounts = source_account_for_raw_export(df)

    df_for_id = df.copy()
    df_for_id["SourceAccount"] = source_accounts

    df.insert(0, "BankRawExportID", [
        make_bank_raw_export_id(row, path.name) for _, row in df_for_id.iterrows()
    ])
    df.insert(1, "SourceBank", SOURCE_BANK)
    df.insert(2, "SourceAccount", source_accounts)
    df.insert(3, "SourceFile", path.name)
    df.insert(4, "ExportDate", file_modified_date(path))
    df.insert(5, "ImportedAt", imported_at)

    return df
