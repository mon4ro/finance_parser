from __future__ import annotations

import re

import pandas as pd

from finance_parser.common import normalise_header, normalise_text

# OP's own dividend-notice text (as it appears in the budgeting side's
# UnifiedTransactions "Message" column) for a dividend that required
# currency conversion - shape (invented numbers, not a real dividend):
#
#   Viesti: OP Säilytys Oy TELIA COMPANY AB SE0000667925 Osinkotuotto
#   Osinko 0,50 SEK/KplOmistettu määrä 100Kpl Tuoton määrä 50,00SEK
#   Lähdevero SE15,0 % 7,50SEKVal.kurssi 11,000000
#
# A EUR-native dividend (e.g. Nokia, Kesko) uses a different withholding
# label ("Ennakonpid.") and never includes "Val.kurssi" at all - this
# pattern deliberately only matches the currency-conversion case, since
# that's the only one with anything worth extracting. Real data check:
# of every real "Osinkotuotto" message, only Telia's SEK-denominated ones
# match this; the EUR-native ones correctly don't.
_OP_DIVIDEND_MESSAGE_PATTERN = re.compile(
    r"OP Säilytys Oy (?P<company>.+?) (?P<isin>[A-Z]{2}\d{10}) "
    r"Osinkotuotto Osinko [\d,]+ (?P<currency>[A-Z]{3})/Kpl"
    r"Omistettu määrä [\d,]+Kpl "
    r"Tuoton määrä (?P<gross>[\d,]+)(?P=currency)"
    r"Lähdevero .+? % (?P<tax>[\d,]+)(?P=currency)"
    r"Val\.kurssi (?P<fxrate>[\d,]+)"
)


def _finnish_decimal(text: str) -> float:
    return float(text.replace(",", "."))


def parse_op_dividend_message(message: object) -> dict[str, object] | None:
    """
    Extract the local-currency dividend detail from one OP custody dividend
    notice, if the message matches that pattern. Returns None for anything
    else (a EUR-native dividend, a different bank's message, an unrelated
    transaction) - this is a best-effort enrichment, not a required parse.
    """
    text = normalise_text(message)
    if not text:
        return None

    match = _OP_DIVIDEND_MESSAGE_PATTERN.search(text)
    if not match:
        return None

    return {
        "NormalizedInstrument": normalise_text(match.group("company")),
        "ISIN": match.group("isin"),
        "LocalCurrency": match.group("currency"),
        "GrossDividendLocal": round(_finnish_decimal(match.group("gross")), 2),
        "TaxWithheldLocal": round(_finnish_decimal(match.group("tax")), 2),
        "ExchangeRate": round(_finnish_decimal(match.group("fxrate")), 6),
    }


def load_op_dividend_local_currency_details(budgeting_workbook) -> pd.DataFrame:
    """
    Scan the budgeting side's UnifiedTransactions for OP dividend notices
    that required currency conversion, and return one row per match:
    (NormalizedInstrument, TradeDate) plus the local-currency detail. This
    is the only place this project reads budgeting-side output from the
    investment pipeline - deliberately kept as an isolated, optional
    enrichment (see build_dividend_history.py's --budgeting-workbook flag)
    rather than a hard dependency between the two otherwise-separate
    pipelines.
    """
    columns = ["NormalizedInstrument", "TradeDate", "ISIN", "LocalCurrency", "GrossDividendLocal", "TaxWithheldLocal", "ExchangeRate"]

    if not budgeting_workbook.exists():
        return pd.DataFrame(columns=columns)

    df = pd.read_excel(budgeting_workbook, sheet_name="UnifiedTransactions", dtype=object, engine="openpyxl")
    df.columns = [normalise_header(c) for c in df.columns]

    rows: list[dict[str, object]] = []
    for _, row in df.iterrows():
        parsed = parse_op_dividend_message(row.get("Message", ""))
        if parsed is None:
            continue
        parsed["TradeDate"] = pd.to_datetime(row.get("Date", ""), errors="coerce")
        rows.append(parsed)

    if not rows:
        return pd.DataFrame(columns=columns)

    result = pd.DataFrame(rows)
    result = result.dropna(subset=["TradeDate"])
    return result[columns]


def enrich_dividend_history(dividend_history: pd.DataFrame, local_currency_details: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """
    Left-join local-currency dividend detail onto dividend_history by
    (NormalizedInstrument, TradeDate) - matched rows gain
    GrossDividendLocal/LocalCurrency/TaxWithheldLocal/ExchangeRate; rows
    with no match (the vast majority - only Telia currently has any) get
    blank values, never an error. Returns (enriched, rows_matched).
    """
    out = dividend_history.copy()
    for col in ["LocalCurrency", "GrossDividendLocal", "TaxWithheldLocal", "ExchangeRate"]:
        out[col] = ""

    if len(local_currency_details) == 0:
        return out, 0

    details = local_currency_details.copy()
    details["_TradeDateStr"] = pd.to_datetime(details["TradeDate"]).dt.strftime("%Y-%m-%d")
    details["_Key"] = details["NormalizedInstrument"].str.upper() + "|" + details["_TradeDateStr"]
    details = details.drop_duplicates(subset=["_Key"], keep="first").set_index("_Key")

    out_key = out["NormalizedInstrument"].astype(str).str.upper() + "|" + out["TradeDate"].astype(str)
    matched = 0
    for idx, key in out_key.items():
        if key in details.index:
            detail = details.loc[key]
            out.at[idx, "LocalCurrency"] = detail["LocalCurrency"]
            out.at[idx, "GrossDividendLocal"] = detail["GrossDividendLocal"]
            out.at[idx, "TaxWithheldLocal"] = detail["TaxWithheldLocal"]
            out.at[idx, "ExchangeRate"] = detail["ExchangeRate"]
            matched += 1

    return out, matched
