from __future__ import annotations

import argparse
import json
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from finance_parser.common import normalise_header, normalise_text
from finance_parser.investments.fetch_instrument_prices import (
    DEFAULT_INSTRUMENT_MASTER,
    DEFAULT_INVESTMENTS_WORKBOOK,
    build_fetch_plan,
    USER_AGENT,
)


REQUEST_DELAY_SECONDS = 0.5


@dataclass
class Suggestion:
    normalized_instrument: str
    real_currency: str
    query: str
    symbol: str | None
    symbol_name: str | None
    symbol_exchange: str | None
    symbol_currency: str | None
    confidence: str
    reason: str
    all_candidates: list[str]


def real_instrument_currency(investments_workbook: Path, normalized_instrument: str) -> str:
    """
    Most common InstrumentCurrency actually recorded for this instrument in
    real transaction history - used to disambiguate share classes.
    """
    df = pd.read_excel(investments_workbook, sheet_name="InvestmentTransactions", dtype=object, engine="openpyxl")
    df.columns = [normalise_header(c) for c in df.columns]
    df["NormalizedInstrument"] = df["NormalizedInstrument"].map(normalise_text)

    rows = df[df["NormalizedInstrument"] == normalized_instrument]
    currencies = [normalise_text(c) for c in rows["InstrumentCurrency"] if normalise_text(c)]
    if not currencies:
        currencies = [normalise_text(c) for c in rows["CashCurrency"] if normalise_text(c)]

    if not currencies:
        return ""

    return Counter(currencies).most_common(1)[0][0]


def raw_instrument_name(instrument_master_path: Path, normalized_instrument: str) -> str:
    """
    Prefer the human-written RawInstrumentName from InstrumentMaster as the
    search query (it's usually cleaner than the all-caps NormalizedInstrument,
    e.g. "Fortum" vs "FORTUM OYJ").
    """
    if not instrument_master_path.exists():
        return normalized_instrument

    df = pd.read_excel(instrument_master_path, sheet_name="InstrumentMaster", dtype=object, engine="openpyxl")
    df.columns = [normalise_header(c) for c in df.columns]
    df["NormalizedInstrument"] = df["NormalizedInstrument"].map(normalise_text)

    matches = df[df["NormalizedInstrument"] == normalized_instrument]
    if len(matches) > 0:
        raw = normalise_text(matches.iloc[0].get("RawInstrumentName", ""))
        if raw:
            return raw

    return normalized_instrument


def yahoo_search(query: str) -> list[dict]:
    url = "https://query1.finance.yahoo.com/v1/finance/search?" + urllib.parse.urlencode({
        "q": query, "quotesCount": 8, "newsCount": 0,
    })
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=20) as response:
        data = json.loads(response.read().decode("utf-8"))
    return data.get("quotes", [])


def yahoo_quote_currency(symbol: str) -> tuple[str, str, str]:
    """Returns (currency, longName, exchange) for a symbol, or ("", "", "") on failure."""
    url = (
        "https://query1.finance.yahoo.com/v8/finance/chart/"
        f"{urllib.parse.quote(symbol)}?range=5d&interval=1d"
    )
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            data = json.loads(response.read().decode("utf-8"))
        meta = data["chart"]["result"][0]["meta"]
        return (
            normalise_text(meta.get("currency", "")),
            normalise_text(meta.get("longName") or meta.get("shortName", "")),
            normalise_text(meta.get("fullExchangeName", "")),
        )
    except Exception:
        return "", "", ""


def suggest_for_instrument(
    normalized_instrument: str,
    instrument_master_path: Path,
    investments_workbook: Path,
) -> Suggestion:
    real_currency = real_instrument_currency(investments_workbook, normalized_instrument)
    query = raw_instrument_name(instrument_master_path, normalized_instrument)

    try:
        candidates = yahoo_search(query)
    except (urllib.error.URLError, json.JSONDecodeError, KeyError) as exc:
        return Suggestion(
            normalized_instrument, real_currency, query, None, None, None, None,
            "FAILED", f"Search request failed: {exc}", [],
        )

    time.sleep(REQUEST_DELAY_SECONDS)

    equity_and_fund_candidates = [
        c for c in candidates
        if c.get("quoteType") in {"EQUITY", "MUTUALFUND", "ETF"}
    ]

    if not equity_and_fund_candidates:
        return Suggestion(
            normalized_instrument, real_currency, query, None, None, None, None,
            "NOT_FOUND", "No stock/fund/ETF candidates returned by search.", [],
        )

    all_candidate_labels = [
        f"{c.get('symbol')} ({c.get('longname') or c.get('shortname')}, {c.get('exchDisp')})"
        for c in equity_and_fund_candidates
    ]

    # If we know the real currency, prefer a candidate whose live quote
    # currency matches it - this is what catches share-class mismatches
    # (e.g. Handelsbanken A1 EUR vs A9 EUR vs A1 NOK).
    if real_currency:
        for candidate in equity_and_fund_candidates:
            symbol = candidate.get("symbol")
            currency, long_name, exchange = yahoo_quote_currency(symbol)
            time.sleep(REQUEST_DELAY_SECONDS)
            if currency.upper() == real_currency.upper():
                return Suggestion(
                    normalized_instrument, real_currency, query, symbol, long_name, exchange, currency,
                    "HIGH", f"Currency matches real holding ({real_currency}).", all_candidate_labels,
                )

        # No candidate matched the real currency - top search hit only, flagged low confidence.
        top = equity_and_fund_candidates[0]
        symbol = top.get("symbol")
        currency, long_name, exchange = yahoo_quote_currency(symbol)
        return Suggestion(
            normalized_instrument, real_currency, query, symbol, long_name, exchange, currency,
            "LOW", f"No candidate matched real currency {real_currency!r} (this one is {currency!r}) - likely wrong share class or not listed.",
            all_candidate_labels,
        )

    # No known real currency to disambiguate against - just report the top hit.
    top = equity_and_fund_candidates[0]
    symbol = top.get("symbol")
    currency, long_name, exchange = yahoo_quote_currency(symbol)
    return Suggestion(
        normalized_instrument, real_currency, query, symbol, long_name, exchange, currency,
        "LOW", "No real currency on file to cross-check against - verify manually.", all_candidate_labels,
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Suggest PriceSymbol candidates for owned instruments missing one in InstrumentMaster.xlsx. "
                    "Report only - never writes to InstrumentMaster.xlsx."
    )
    parser.add_argument("--instrument-master", default=str(DEFAULT_INSTRUMENT_MASTER))
    parser.add_argument("--investments-workbook", default=str(DEFAULT_INVESTMENTS_WORKBOOK))
    args = parser.parse_args()

    instrument_master_path = Path(args.instrument_master).expanduser().resolve()
    investments_workbook = Path(args.investments_workbook).expanduser().resolve()

    _, missing = build_fetch_plan(instrument_master_path, investments_workbook)

    high, low, not_found, failed = [], [], [], []

    for name in missing:
        print(f"Searching: {name} ...")
        suggestion = suggest_for_instrument(name, instrument_master_path, investments_workbook)
        if suggestion.confidence == "HIGH":
            high.append(suggestion)
        elif suggestion.confidence == "LOW":
            low.append(suggestion)
        elif suggestion.confidence == "NOT_FOUND":
            not_found.append(suggestion)
        else:
            failed.append(suggestion)

    def print_section(title: str, items: list[Suggestion]) -> None:
        print()
        print(f"=== {title} ({len(items)}) ===")
        for s in items:
            print(f"  {s.normalized_instrument!r} (real currency: {s.real_currency or '?'})")
            print(f"    -> {s.symbol}  |  {s.symbol_name}  |  {s.symbol_exchange}  |  {s.symbol_currency}")
            print(f"    reason: {s.reason}")
            if s.confidence == "LOW":
                print(f"    all candidates: {s.all_candidates}")

    print_section("HIGH confidence (currency confirmed)", high)
    print_section("LOW confidence - needs manual check", low)
    print_section("NOT FOUND", not_found)
    print_section("FAILED (network/parse error)", failed)


if __name__ == "__main__":
    main()
