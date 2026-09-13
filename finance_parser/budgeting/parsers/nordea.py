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
    normalise_text,
    read_input_file,
    stable_hash,
)


SOURCE_BANK = "NORDEA"
BANK_RAW_SHEET = "NordeaRawExport"
# RawID prefix: NOD-

NORDEA_COLUMNS = [
    "Kirjauspäivä",
    "Määrä",
    "Maksaja",
    "Maksunsaaja",
    "Nimi",
    "Otsikko",
    "Viesti",
    "Viitenumero",
    "Saldo",
    "Valuutta",
]

OPTIONAL_METADATA_COLUMNS = [
    "Source",
    "SourceAccount",
    "ExportDate",
]

REQUIRED_COLUMNS = NORDEA_COLUMNS


def can_parse(path: Path) -> tuple[bool, str]:
    try:
        df = read_input_file(path, required_columns=REQUIRED_COLUMNS)
        columns = set(normalise_header(c) for c in df.columns)
        missing = [c for c in REQUIRED_COLUMNS if c not in columns]
        if missing:
            return False, f"Missing Nordea columns: {', '.join(missing[:6])}"
        return True, "Matched Nordea column set"
    except Exception as exc:
        return False, str(exc)


def choose_raw_receiver(row: pd.Series) -> str:
    """
    Nordea gives several counterparty-ish columns.

    Preferred:
    - Nimi, if present, because it is usually the readable counterparty name
    - Maksunsaaja, for outgoing payments
    - Maksaja, for incoming payments
    """
    for col in ["Nimi", "Maksunsaaja", "Maksaja"]:
        value = normalise_text(row.get(col, ""))
        if value:
            return value
    return ""


def make_raw_id(row: pd.Series) -> str:
    # Stable identity: avoid fields such as Description/Message that may change
    # when parser mappings are refined.
    return "NOD-" + stable_hash([
        row.get("SourceAccount", ""),
        format_date(row.get("ValueDate", "")),
        normalise_amount(row.get("Amount", 0)),
        row.get("RawReceiver", ""),
        row.get("Reference", ""),
        row.get("Currency", ""),
    ])


def parse_file(path: Path, imported_at: str) -> pd.DataFrame:
    df = read_input_file(path, required_columns=REQUIRED_COLUMNS)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        available = ", ".join(map(str, df.columns))
        raise ValueError(
            f"{path.name} is missing required Nordea column(s): {', '.join(missing)}\n"
            f"Available columns detected: {available}"
        )

    df = df.copy()
    df = df[df["Määrä"].notna()]
    df = df[df["Määrä"].astype(str).str.strip() != ""]
    df = df[df["Kirjauspäivä"].notna()]
    df = df[df["Kirjauspäivä"].astype(str).str.strip() != ""]

    out = pd.DataFrame(index=df.index)

    # Nordea exports in this workbook always represent one fixed account,
    # configured per-workbook in settings (budgeting.source_account_inference.
    # NORDEA.fixed_source_account) rather than hardcoded here.
    out["SourceAccount"] = infer_source_account_from_filename(path, SOURCE_BANK)
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

    # Nordea only gives Kirjauspäivä, so use it for both BookingDate and
    # ValueDate. The shared pipeline uses ValueDate as UnifiedTransactions Date.
    out["BookingDate"] = df["Kirjauspäivä"].map(format_date)
    out["ValueDate"] = df["Kirjauspäivä"].map(format_date)
    out["Amount"] = df["Määrä"].map(normalise_amount)

    # Keep Otsikko as the raw transaction type, but use Viesti as Description
    # because it is the more useful budget-facing explanatory text for Nordea.
    out["TransactionTypeRaw"] = df["Otsikko"].map(normalise_text)
    out["Description"] = df["Viesti"].map(normalise_text)

    out["RawReceiver"] = df.apply(choose_raw_receiver, axis=1)
    out["ReceiverAccount"] = ""
    out["ReceiverBankBIC"] = ""
    out["Reference"] = df["Viitenumero"].map(normalise_text)
    out["Message"] = ""
    out["ArchiveID"] = ""
    out["Balance"] = df["Saldo"].map(normalise_amount)
    out["CurrencyAmount"] = ""
    out["Currency"] = df["Valuutta"].map(normalise_text)
    out["Currency"] = out["Currency"].where(out["Currency"].astype(str).str.strip() != "", "EUR")
    out["Rate"] = 1
    out["MerchantArea"] = ""
    out["MerchantCategory"] = ""
    out["SourceFile"] = path.name

    out["RawID"] = out.apply(make_raw_id, axis=1)

    return out[RAW_COLUMNS]



def make_bank_raw_export_id(row: pd.Series, source_file: str = "") -> str:
    """
    Stable Nordea raw-export row identity.

    Deliberately excludes SourceFile / ExportDate / ImportedAt / script-produced
    SourceAccount. Uses transaction-content fields available in the export.
    """
    return "NODRAW-" + stable_hash([
        SOURCE_BANK,
        row.get("Kirjauspäivä", ""),
        row.get("Määrä", ""),
        row.get("Maksaja", ""),
        row.get("Maksunsaaja", ""),
        row.get("Nimi", ""),
        row.get("Otsikko", ""),
        row.get("Viesti", ""),
        row.get("Viitenumero", ""),
        row.get("Valuutta", ""),
    ], length=16)


def parse_bank_raw_rows(path: Path, imported_at: str) -> pd.DataFrame:
    """
    Preserve a closer-to-original Nordea export sheet in ParsedTransactions.xlsx.
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
    df.insert(2, "SourceAccount", infer_source_account_from_filename(path, SOURCE_BANK))
    df.insert(3, "SourceFile", path.name)
    df.insert(4, "ExportDate", file_modified_date(path))
    df.insert(5, "ImportedAt", imported_at)

    return df
