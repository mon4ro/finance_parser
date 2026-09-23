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
# RawID prefix: CASH- (or VINTED- - see source_bank_for_entry_id below)

# CashEntries.xlsx is a manually maintained ledger, not a bank export: the
# user hand-types one row per cash transaction (see input/budgeting/CashEntries.xlsx).
# EntryID exists only to guarantee row uniqueness for hashing purposes - unlike
# bank exports, there is no bank-native reference/archive number to lean on, so
# two genuinely different cash purchases with identical date/amount/receiver
# text would otherwise hash to the same RawID and silently collapse into one
# transaction during dedup.
#
# Vinted marketplace sales/purchases/refunds/withdrawals are hand-entered
# into this same ledger rather than parsed from a real Vinted export - there
# is no reliable machine-readable export to build a standing parser against
# (the one seen was LLM-reconstructed from screenshots, not authoritative).
# They're distinguished purely by an EntryID prefix (VINTED-0001, ...) the
# user/whoever enters the rows commits to using: SourceAccount stays CASH
# (neither a Vinted virtual wallet nor physical cash has a real bank
# statement/balance to reconcile against), but SourceBank becomes VINTED
# instead of CASH, so TransactionType reads "Virtual" (see
# transaction_type_for_source_bank in common.py) and content-based
# CategoryRules/TransactionRules can scope narrowly to just these rows via
# SourceBank, without inventing any new parser-level Include/category logic
# (still just RAW_COLUMNS - normal Include/category rules apply exactly
# like any other source). Owner is the one exception - see CASH_COLUMNS'
# "Owner" field below.
VINTED_SOURCE_BANK = "VINTED"
VINTED_ENTRY_ID_PREFIX = "VINTED-"

CASH_COLUMNS = [
    "EntryID",
    "Date",
    # Renamed from "Source" (2026-09-23) to harmonise with
    # UnifiedTransactions' own "Owner" column - same concept, same name on
    # both sides now. Whoever types a row here already knows whose cash it
    # was; unlike a bank export, this is a real per-row fact, not something
    # that needs inferring. raw_to_unified_rows() in common.py prefers this
    # explicit value over the generic SourceAccount-based default.
    "Owner",
    "Amount",
    "RawReceiver",
]

OPTIONAL_METADATA_COLUMNS = [
    "Currency",
    "Description",
    "ExportDate",
    # Optional, structured alternative to matching content-based rules
    # against free-text Description: SALE/PURCHASE/WITHDRAWAL/DEPOSIT (or
    # any other short code) - carried through to UnifiedTransactions as
    # TransactionTypeRaw, unambiguous and typo-tolerant in a way exact
    # Description-prefix matching isn't. Optional - a plain physical cash
    # entry has no need for it and can leave it blank, same as today.
    "Type",
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


def source_bank_for_entry_id(entry_id: object) -> str:
    if normalise_text(entry_id).upper().startswith(VINTED_ENTRY_ID_PREFIX):
        return VINTED_SOURCE_BANK
    return SOURCE_BANK


def make_raw_id(row: pd.Series) -> str:
    # row is a row of `out` (already-mapped RAW_COLUMNS shape) at the point
    # this is applied, so SourceBank is already resolved per-row - reusing
    # it here (rather than re-deriving from EntryID) keeps the RawID prefix
    # and the SourceBank column always in agreement.
    prefix = "VINTED-" if row.get("SourceBank") == VINTED_SOURCE_BANK else "CASH-"
    return prefix + stable_hash([
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
    # cash was used - SourceAccount is always the fixed literal "CASH",
    # never derived from Owner. Owner itself, unlike SourceAccount, DOES
    # carry through to RawTransactions/UnifiedTransactions (see
    # raw_to_unified_rows() in common.py) - it's a real per-row fact the
    # person entering the row already states directly.
    out["EntryID"] = df["EntryID"].map(normalise_text)
    out["SourceAccount"] = "CASH"
    out["SourceBank"] = df["EntryID"].map(source_bank_for_entry_id)
    out["Owner"] = df["Owner"].map(lambda v: normalise_text(v).upper())

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
    out["TransactionTypeRaw"] = (
        df["Type"].map(lambda v: normalise_text(v).upper()) if "Type" in df.columns else ""
    )
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
    for manually entered rows. The sheet name/ID prefix stay "CashRawExport"/
    "CASHRAW-" regardless of SourceBank (one shared audit sheet for the whole
    file) - only the hash salt varies, so a real CASH-prefixed row's ID is
    unaffected and a VINTED-prefixed row gets a distinct hash.
    """
    return "CASHRAW-" + stable_hash([
        source_bank_for_entry_id(row.get("EntryID", "")),
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
    unchanged, for audit - including the "Owner" column, which (unlike most
    of this raw audit sheet's other columns) also separately flows through
    to the canonical RawTransactions/UnifiedTransactions schema now (see
    parse_file() above).
    """
    df = read_input_file(path, required_columns=REQUIRED_COLUMNS).copy()
    df.columns = [normalise_header(c) for c in df.columns]

    df = df.loc[
        ~df.apply(lambda r: all(normalise_text(v) == "" for v in r), axis=1)
    ].copy()

    df.insert(0, "BankRawExportID", [
        make_bank_raw_export_id(row, path.name) for _, row in df.iterrows()
    ])
    df.insert(1, "SourceBank", df["EntryID"].map(source_bank_for_entry_id))
    df.insert(2, "SourceFile", path.name)

    # CashEntries.xlsx already carries its own manually typed ExportDate
    # column (the date the row was scribbled into the ledger). Don't
    # duplicate/overwrite it with a computed bookkeeping stamp - only add
    # one if the source data doesn't already have it.
    if "ExportDate" not in df.columns:
        df.insert(3, "ExportDate", file_modified_date(path))

    df.insert(len(df.columns), "ImportedAt", imported_at)

    return df
