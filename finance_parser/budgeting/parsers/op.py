from __future__ import annotations

from pathlib import Path

import pandas as pd

from finance_parser.common import (
    RAW_COLUMNS,
    file_modified_date,
    format_date,
    infer_source_account_from_filename,
    normalise_amount,
    normalise_header,
    normalise_reference_text,
    normalise_text,
    read_input_file,
    stable_hash,
)


SOURCE_BANK = "OP"
BANK_RAW_SHEET = "OPRawExport"
# RawID prefix: OP-

OP_COLUMNS = [
    "Kirjauspäivä",
    "Arvopäivä",
    "Määrä EUROA",
    "Laji",
    "Selitys",
    "Saaja/Maksaja",
    "Saajan tilinumero",
    "Saajan pankin BIC",
    "Viite",
    "Viesti",
    "Arkistointitunnus",
]

OPTIONAL_METADATA_COLUMNS = [
    "Source",
    "SourceAccount",
    "ExportDate",
]

REQUIRED_COLUMNS = OP_COLUMNS


def can_parse(path: Path) -> tuple[bool, str]:
    try:
        df = read_input_file(path, required_columns=REQUIRED_COLUMNS)
        columns = set(normalise_header(c) for c in df.columns)
        missing = [c for c in REQUIRED_COLUMNS if c not in columns]
        if missing:
            return False, f"Missing OP columns: {', '.join(missing[:6])}"
        return True, "Matched OP column set"
    except Exception as exc:
        return False, str(exc)


def make_raw_id(row: pd.Series) -> str:
    return "OP-" + stable_hash([
        row.get("SourceAccount", ""),
        format_date(row.get("ValueDate", "")),
        normalise_amount(row.get("Amount", 0)),
        row.get("RawReceiver", ""),
        row.get("Description", ""),
        row.get("Reference", ""),
        row.get("Message", ""),
        row.get("ArchiveID", ""),
    ])


def parse_file(path: Path, imported_at: str) -> pd.DataFrame:
    df = read_input_file(path, required_columns=REQUIRED_COLUMNS)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        available = ", ".join(map(str, df.columns))
        raise ValueError(
            f"{path.name} is missing required OP column(s): {', '.join(missing)}\n"
            f"Available columns detected: {available}"
        )

    df = df.copy()
    df = df[df["Määrä EUROA"].notna()]
    df = df[df["Määrä EUROA"].astype(str).str.strip() != ""]
    df = df[df["Arvopäivä"].notna()]
    df = df[df["Arvopäivä"].astype(str).str.strip() != ""]

    out = pd.DataFrame(index=df.index)

    # Keep OP exports untouched: SourceAccount can come from an optional
    # SourceAccount/Source column, otherwise it is inferred from the filename.
    if "SourceAccount" in df.columns:
        out["SourceAccount"] = df["SourceAccount"].map(normalise_text)
    elif "Source" in df.columns:
        out["SourceAccount"] = df["Source"].map(normalise_text)
    else:
        out["SourceAccount"] = infer_source_account_from_filename(path, SOURCE_BANK)

    out["SourceBank"] = SOURCE_BANK

    # ExportDate now aligns with Norwegian: default is file modified date.
    # A manual ExportDate column can still override it if present.
    if "ExportDate" in df.columns:
        parsed_export_dates = df["ExportDate"].map(format_date)
        out["ExportDate"] = parsed_export_dates.where(
            parsed_export_dates.astype(str).str.strip() != "",
            file_modified_date(path),
        )
    else:
        out["ExportDate"] = file_modified_date(path)

    out["ImportedAt"] = imported_at
    out["BookingDate"] = df["Kirjauspäivä"].map(format_date)
    out["ValueDate"] = df["Arvopäivä"].map(format_date)
    out["Amount"] = df["Määrä EUROA"].map(normalise_amount)
    out["TransactionTypeRaw"] = df["Laji"].map(normalise_text)
    out["Description"] = df["Selitys"].map(normalise_text)
    out["RawReceiver"] = df["Saaja/Maksaja"].map(normalise_text)
    out["ReceiverAccount"] = df["Saajan tilinumero"].map(normalise_text)
    out["ReceiverBankBIC"] = df["Saajan pankin BIC"].map(normalise_text)
    out["Reference"] = df["Viite"].map(normalise_reference_text)
    out["Message"] = df["Viesti"].map(normalise_text)
    out["ArchiveID"] = df["Arkistointitunnus"].map(normalise_text)
    out["Balance"] = ""
    out["CurrencyAmount"] = ""
    out["Currency"] = "EUR"
    out["Rate"] = 1
    out["MerchantArea"] = ""
    out["MerchantCategory"] = ""
    out["SourceFile"] = path.name

    out["RawID"] = out.apply(make_raw_id, axis=1)

    return out[RAW_COLUMNS]



def source_account_for_raw_export(df: pd.DataFrame, path: Path) -> pd.Series:
    if "SourceAccount" in df.columns:
        return df["SourceAccount"].map(normalise_text)
    if "Source" in df.columns:
        return df["Source"].map(normalise_text)
    return pd.Series([infer_source_account_from_filename(path, SOURCE_BANK)] * len(df), index=df.index)


def make_bank_raw_export_id(row: pd.Series, source_file: str = "") -> str:
    """
    Stable OP raw-export row identity.

    Deliberately excludes SourceFile / ExportDate / ImportedAt / inferred
    SourceAccount so overlapping OP exports produce the same ID for the same
    bank-native transaction row.
    """
    return "OPRAW-" + stable_hash([
        SOURCE_BANK,
        row.get("Arkistointitunnus", ""),
        row.get("Arvopäivä", ""),
        row.get("Kirjauspäivä", ""),
        row.get("Määrä EUROA", ""),
        row.get("Saaja/Maksaja", ""),
        row.get("Viite", ""),
        row.get("Viesti", ""),
        row.get("Laji", ""),
    ], length=16)


def parse_bank_raw_rows(path: Path, imported_at: str) -> pd.DataFrame:
    """
    Preserve a closer-to-original OP export sheet in ParsedTransactions.xlsx.
    """
    df = read_input_file(path, required_columns=REQUIRED_COLUMNS).copy()
    df.columns = [normalise_header(c) for c in df.columns]

    df = df.loc[
        ~df.apply(lambda r: all(normalise_text(v) == "" for v in r), axis=1)
    ].copy()

    source_accounts = source_account_for_raw_export(df, path)

    df.insert(0, "BankRawExportID", [
        make_bank_raw_export_id(row, path.name) for _, row in df.assign(SourceAccount=source_accounts).iterrows()
    ])
    df.insert(1, "SourceBank", SOURCE_BANK)
    df.insert(2, "SourceAccount", source_accounts)
    df.insert(3, "SourceFile", path.name)
    df.insert(4, "ExportDate", file_modified_date(path))
    df.insert(5, "ImportedAt", imported_at)

    return df
