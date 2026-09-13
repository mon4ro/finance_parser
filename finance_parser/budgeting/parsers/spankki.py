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


SOURCE_BANK = "SPANKKI"
BANK_RAW_SHEET = "SpankkiRawExport"
# RawID prefix: SPK-

SPANKKI_COLUMNS = [
    "Kirjauspäivä",
    "Maksupäivä",
    "Summa",
    "Tapahtumalaji",
    "Maksaja",
    "Saajan nimi",
    "Saajan tilinumero",
    "Saajan BIC-tunnus",
    "Viitenumero",
    "Viesti",
    "Arkistointitunnus",
]

OPTIONAL_METADATA_COLUMNS = [
    "Source",
    "SourceAccount",
    "ExportDate",
]

REQUIRED_COLUMNS = SPANKKI_COLUMNS


def can_parse(path: Path) -> tuple[bool, str]:
    try:
        df = read_input_file(path, required_columns=REQUIRED_COLUMNS)
        columns = set(normalise_header(c) for c in df.columns)
        missing = [c for c in REQUIRED_COLUMNS if c not in columns]
        if missing:
            return False, f"Missing S-Pankki columns: {', '.join(missing[:6])}"
        return True, "Matched S-Pankki column set"
    except Exception as exc:
        return False, str(exc)


def choose_raw_receiver(row: pd.Series) -> str:
    """
    For outgoing purchases/payments, Saajan nimi is usually the merchant/recipient.
    For incoming payments, Maksaja is usually the useful counterparty.
    """
    amount = normalise_amount(row.get("Summa", 0))

    if amount < 0:
        for col in ["Saajan nimi", "Maksaja"]:
            value = normalise_text(row.get(col, ""))
            if value:
                return value
    else:
        for col in ["Maksaja", "Saajan nimi"]:
            value = normalise_text(row.get(col, ""))
            if value:
                return value

    return ""


def make_raw_id(row: pd.Series) -> str:
    return "SPK-" + stable_hash([
        row.get("SourceAccount", ""),
        format_date(row.get("ValueDate", "")),
        normalise_amount(row.get("Amount", 0)),
        row.get("RawReceiver", ""),
        row.get("Reference", ""),
        row.get("ArchiveID", ""),
        row.get("Currency", ""),
    ])


def parse_file(path: Path, imported_at: str) -> pd.DataFrame:
    df = read_input_file(path, required_columns=REQUIRED_COLUMNS)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        available = ", ".join(map(str, df.columns))
        raise ValueError(
            f"{path.name} is missing required S-Pankki column(s): {', '.join(missing)}\n"
            f"Available columns detected: {available}"
        )

    df = df.copy()
    df = df[df["Summa"].notna()]
    df = df[df["Summa"].astype(str).str.strip() != ""]
    df = df[df["Maksupäivä"].notna()]
    df = df[df["Maksupäivä"].astype(str).str.strip() != ""]

    out = pd.DataFrame(index=df.index)

    # S-Pankki is a dedicated groceries buffer account in this workbook.
    out["SourceAccount"] = "SPANKKI"
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
    out["BookingDate"] = df["Kirjauspäivä"].map(format_date)
    out["ValueDate"] = df["Maksupäivä"].map(format_date)
    out["Amount"] = df["Summa"].map(normalise_amount)

    out["TransactionTypeRaw"] = df["Tapahtumalaji"].map(normalise_text)
    out["Description"] = df["Tapahtumalaji"].map(normalise_text)
    out["RawReceiver"] = df.apply(choose_raw_receiver, axis=1)

    out["ReceiverAccount"] = df["Saajan tilinumero"].map(normalise_text)
    out["ReceiverBankBIC"] = df["Saajan BIC-tunnus"].map(normalise_text)
    out["Reference"] = df["Viitenumero"].map(normalise_text)
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


def make_bank_raw_export_id(row: pd.Series, source_file: str = "") -> str:
    """
    Stable S-Pankki raw-export row identity.

    Deliberately excludes SourceFile / ExportDate / ImportedAt. Anchored on
    Arkistointitunnus plus date/amount/counterparty sanity fields.
    """
    return "SPKRAW-" + stable_hash([
        SOURCE_BANK,
        row.get("Arkistointitunnus", ""),
        row.get("Maksupäivä", ""),
        row.get("Kirjauspäivä", ""),
        row.get("Summa", ""),
        row.get("Saajan nimi", ""),
        row.get("Maksaja", ""),
        row.get("Tapahtumalaji", ""),
    ], length=16)


def parse_bank_raw_rows(path: Path, imported_at: str) -> pd.DataFrame:
    """
    Preserve a closer-to-original S-Pankki export sheet in ParsedTransactions.xlsx.
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
    df.insert(2, "SourceAccount", "SPANKKI")
    df.insert(3, "SourceFile", path.name)
    df.insert(4, "ExportDate", file_modified_date(path))
    df.insert(5, "ImportedAt", imported_at)

    return df
