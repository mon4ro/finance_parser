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


SOURCE_BANK = "CASH"
BANK_RAW_SHEET = "CashRawExport"
# RawID prefix: CASH-

# CashEntries.xlsx is a manually maintained ledger, not a bank export: the
# user hand-types one row per cash transaction (see input/budgeting/CashEntries.xlsx).
# EntryID exists only to guarantee row uniqueness for hashing purposes - unlike
# bank exports, there is no bank-native reference/archive number to lean on, so
# two genuinely different cash purchases with identical date/amount/receiver
# text would otherwise hash to the same RawID and silently collapse into one
# transaction during dedup.
CASH_COLUMNS = [
    "EntryID",
    "Date",
    "Source",
    "Amount",
    "RawReceiver",
]

OPTIONAL_METADATA_COLUMNS = [
    "Currency",
    "Description",
    "ExportDate",
]

REQUIRED_COLUMNS = CASH_COLUMNS


def can_parse(path: Path) -> tuple[bool, str]:
    try:
        df = read_input_file(path, required_columns=REQUIRED_COLUMNS)
        columns = set(normalise_header(c) for c in df.columns)
        missing = [c for c in REQUIRED_COLUMNS if c not in columns]
        if missing:
            return False, f"Missing Cash columns: {', '.join(missing[:6])}"
        return True, "Matched Cash column set"
    except Exception as exc:
        return False, str(exc)


def make_raw_id(row: pd.Series) -> str:
    return "CASH-" + stable_hash([
        row.get("EntryID", ""),
        format_date(row.get("ValueDate", "")),
        normalise_amount(row.get("Amount", 0)),
        row.get("RawReceiver", ""),
        row.get("Description", ""),
    ])


def parse_file(path: Path, imported_at: str) -> pd.DataFrame:
    df = read_input_file(path, required_columns=REQUIRED_COLUMNS)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        available = ", ".join(map(str, df.columns))
        raise ValueError(
            f"{path.name} is missing required Cash column(s): {', '.join(missing)}\n"
            f"Available columns detected: {available}"
        )

    df = df.copy()
    df = df[df["Amount"].notna()]
    df = df[df["Amount"].astype(str).str.strip() != ""]
    df = df[df["Date"].notna()]
    df = df[df["Date"].astype(str).str.strip() != ""]

    out = pd.DataFrame(index=df.index)

    # Cash is always its own bucket, independent of which person's physical
    # cash was used. "Source" (whose cash) is intentionally not carried into
    # RawTransactions/UnifiedTransactions - it is preserved in CashRawExport
    # only, for audit. Do not use it to infer Owner here.
    out["EntryID"] = df["EntryID"].map(normalise_text)
    out["SourceAccount"] = "CASH"
    out["SourceBank"] = SOURCE_BANK

    if "ExportDate" in df.columns:
        parsed_export_dates = df["ExportDate"].map(format_date)
        out["ExportDate"] = parsed_export_dates.where(
            parsed_export_dates.astype(str).str.strip() != "",
            file_modified_date(path),
        )
    else:
        out["ExportDate"] = file_modified_date(path)

    out["ImportedAt"] = imported_at

    # A cash entry has only one date - use it for both BookingDate and
    # ValueDate, same approach as Nordea (which only has Kirjauspäivä). The
    # shared pipeline uses ValueDate as UnifiedTransactions Date.
    out["BookingDate"] = df["Date"].map(format_date)
    out["ValueDate"] = df["Date"].map(format_date)
    out["Amount"] = df["Amount"].map(normalise_amount)
    out["TransactionTypeRaw"] = ""
    out["Description"] = df["Description"].map(normalise_text) if "Description" in df.columns else ""
    out["RawReceiver"] = df["RawReceiver"].map(normalise_text)
    out["ReceiverAccount"] = ""
    out["ReceiverBankBIC"] = ""
    out["Reference"] = ""
    out["Message"] = ""
    out["ArchiveID"] = ""
    out["Balance"] = ""
    out["CurrencyAmount"] = ""
    out["Currency"] = df["Currency"].map(normalise_text) if "Currency" in df.columns else ""
    out["Currency"] = out["Currency"].where(out["Currency"].astype(str).str.strip() != "", "EUR")
    out["Rate"] = 1
    out["MerchantArea"] = ""
    out["MerchantCategory"] = ""
    out["SourceFile"] = path.name

    out["RawID"] = out.apply(make_raw_id, axis=1)

    return out[RAW_COLUMNS]


def make_bank_raw_export_id(row: pd.Series, source_file: str = "") -> str:
    """
    Stable Cash raw-export row identity.

    Includes EntryID, since there is no other bank-native identity to lean on
    for manually entered rows.
    """
    return "CASHRAW-" + stable_hash([
        SOURCE_BANK,
        row.get("EntryID", ""),
        # format_date() normalises the date regardless of how Excel stored
        # the cell (plain text vs a native date type) - without this, fixing
        # a cell's type (e.g. text "2026-05-09" -> a real Excel date) changes
        # the hash and produces a duplicate audit row for the same entry.
        format_date(row.get("Date", "")),
        row.get("Amount", ""),
        row.get("RawReceiver", ""),
        row.get("Description", ""),
    ], length=16)


def parse_bank_raw_rows(path: Path, imported_at: str) -> pd.DataFrame:
    """
    Preserve the manually entered CashEntries.xlsx rows in ParsedTransactions.xlsx,
    including the "Source" (whose cash) column that is deliberately dropped from
    the canonical RawTransactions/UnifiedTransactions schema.
    """
    df = read_input_file(path, required_columns=REQUIRED_COLUMNS).copy()
    df.columns = [normalise_header(c) for c in df.columns]

    df = df.loc[
        ~df.apply(lambda r: all(normalise_text(v) == "" for v in r), axis=1)
    ].copy()

    df.insert(0, "BankRawExportID", [
        make_bank_raw_export_id(row, path.name) for _, row in df.iterrows()
    ])
    df.insert(1, "SourceBank", SOURCE_BANK)
    df.insert(2, "SourceFile", path.name)

    # CashEntries.xlsx already carries its own manually typed ExportDate
    # column (the date the row was scribbled into the ledger). Don't
    # duplicate/overwrite it with a computed bookkeeping stamp - only add
    # one if the source data doesn't already have it.
    if "ExportDate" not in df.columns:
        df.insert(3, "ExportDate", file_modified_date(path))

    df.insert(len(df.columns), "ImportedAt", imported_at)

    return df
