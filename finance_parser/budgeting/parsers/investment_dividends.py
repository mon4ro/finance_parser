from __future__ import annotations

from pathlib import Path

import pandas as pd

from finance_parser.common import (
    RAW_COLUMNS,
    file_modified_date,
    normalise_amount,
    normalise_header,
    normalise_text,
    stable_hash,
)


SOURCE_BANK = "INVESTMENT_DIVIDEND"
BANK_RAW_SHEET = "InvestmentDividendRawExport"
DIVIDEND_HISTORY_SHEET = "DividendHistory"
# RawID prefix: INVDIV-

# Real dividend events from Nordnet/EVLI-held instruments have no matching
# budgeting-side bank transaction at all: the cash lands in the broker's own
# cash balance (already tracked separately via the investment pipeline's
# cash-balance feature), not a bank account, unlike OP-held instruments
# which settle same-day into a real transaction (see enrich_dividend_income.py,
# the OP-side cross-reference for those). This parser bridges that gap by
# generating one real budgeting-side row per real dividend event straight
# from the investment pipeline's own DividendHistory.xlsx - not a guess, a
# real recorded dividend, just sourced from a different pipeline instead of
# a bank export file.
SYNTHETIC_BROKERS = {"NORDNET", "EVLI"}

REQUIRED_COLUMNS = [
    "TradeDate",
    "Broker",
    "Portfolio",
    "PortfolioOwner",
    "NormalizedInstrument",
    "NetDividendEUR",
]


def can_parse(path: Path) -> tuple[bool, str]:
    try:
        df = pd.read_excel(path, sheet_name=DIVIDEND_HISTORY_SHEET, dtype=object, engine="openpyxl")
        df.columns = [normalise_header(c) for c in df.columns]
        missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            return False, f"Missing DividendHistory columns: {', '.join(missing)}"
        return True, "Matched investment DividendHistory column set"
    except Exception as exc:
        return False, str(exc)


def make_raw_id(row: pd.Series) -> str:
    return "INVDIV-" + stable_hash([
        SOURCE_BANK,
        row.get("Broker", ""),
        row.get("Portfolio", ""),
        row.get("NormalizedInstrument", ""),
        row.get("ValueDate", ""),
        row.get("Amount", ""),
    ])


def parse_file(path: Path, imported_at: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=DIVIDEND_HISTORY_SHEET, dtype=object, engine="openpyxl")
    df.columns = [normalise_header(c) for c in df.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"{path.name} is missing required DividendHistory column(s): {', '.join(missing)}\n"
            f"Available columns detected: {', '.join(map(str, df.columns))}"
        )

    df = df.copy()
    df["Broker"] = df["Broker"].map(normalise_text)
    df = df[df["Broker"].isin(SYNTHETIC_BROKERS)]
    df = df[df["TradeDate"].notna()]
    df = df[df["TradeDate"].astype(str).str.strip() != ""]
    df["NetDividendEUR"] = pd.to_numeric(df["NetDividendEUR"], errors="coerce")
    df = df[df["NetDividendEUR"].notna()]
    df = df[df["NetDividendEUR"] > 0]

    export_date = file_modified_date(path)

    rows: list[dict[str, object]] = []

    for _, row in df.iterrows():
        instrument = normalise_text(row.get("NormalizedInstrument", ""))
        broker = normalise_text(row.get("Broker", ""))
        portfolio = normalise_text(row.get("Portfolio", ""))
        owner = normalise_text(row.get("PortfolioOwner", ""))
        trade_date = normalise_text(row.get("TradeDate", ""))
        amount = normalise_amount(row.get("NetDividendEUR", 0))

        parsed_row = {
            # SourceAccount carries the real investment-side PortfolioOwner
            # so this project's normal SourceAccount->Owner rules resolve it
            # the same way they already do for every real bank account -
            # this parser does not set Owner directly (see CLAUDE.md).
            "SourceAccount": owner,
            "SourceBank": SOURCE_BANK,
            "ExportDate": export_date,
            "ImportedAt": imported_at,
            "BookingDate": trade_date,
            "ValueDate": trade_date,
            "Amount": amount,
            "TransactionTypeRaw": "DIVIDEND",
            "Description": f"{instrument} dividend ({broker} {portfolio})",
            # Real bug found and fixed: the bare instrument name alone (e.g.
            # "FORTUM", "ELISA") is also a real household utility/telecom
            # bill payee with its own existing CategoryRules EXACT-match
            # rule (electricity, phone subscription) - a dividend row for
            # the SAME company name collided with that rule and got
            # miscategorized as a real utility expense instead of income.
            # Suffixing " DIVIDEND" keeps it readable while never exact-
            # matching an existing merchant-name rule.
            "RawReceiver": f"{instrument} DIVIDEND",
            "ReceiverAccount": "",
            "ReceiverBankBIC": "",
            "Reference": "",
            "Message": f"Real dividend event from investment pipeline DividendHistory.xlsx - {broker}/{portfolio}, net of withheld tax.",
            "ArchiveID": "",
            "Balance": "",
            "CurrencyAmount": "",
            "Currency": "EUR",
            "Rate": 1,
            "MerchantArea": "",
            "MerchantCategory": "",
            "SourceFile": path.name,
        }

        parsed_row["RawID"] = make_raw_id(pd.Series(parsed_row))
        rows.append({column: parsed_row.get(column, "") for column in RAW_COLUMNS})

    return pd.DataFrame(rows, columns=RAW_COLUMNS)


def parse_bank_raw_rows(path: Path, imported_at: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=DIVIDEND_HISTORY_SHEET, dtype=object, engine="openpyxl")
    df.columns = [normalise_header(c) for c in df.columns]
    df["Broker"] = df["Broker"].map(normalise_text)
    df = df[df["Broker"].isin(SYNTHETIC_BROKERS)].copy()

    df.insert(0, "BankRawExportID", [
        make_raw_id(pd.Series({
            "Broker": row.get("Broker", ""),
            "Portfolio": row.get("Portfolio", ""),
            "NormalizedInstrument": row.get("NormalizedInstrument", ""),
            "ValueDate": row.get("TradeDate", ""),
            "Amount": row.get("NetDividendEUR", ""),
        }))
        for _, row in df.iterrows()
    ])
    df.insert(1, "SourceBank", SOURCE_BANK)
    df.insert(2, "SourceFile", path.name)
    df.insert(3, "ExportDate", file_modified_date(path))
    df.insert(4, "ImportedAt", imported_at)

    return df
