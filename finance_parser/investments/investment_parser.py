from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from finance_parser.common import imported_at_now, make_import_run_id, normalise_header

from finance_parser.investments.investment_common import (
    INVESTMENT_IMPORT_LOG_COLUMNS,
    INVESTMENT_IMPORT_LOG_SHEET,
    INVESTMENT_RAW_COLUMNS,
    INVESTMENT_RAW_SHEET,
    INVESTMENT_TRANSACTIONS_COLUMNS,
    INVESTMENT_TRANSACTIONS_SHEET,
    apply_instrument_master,
    apply_portfolio_ownership,
    load_instrument_master,
    raw_to_investment_transactions,
    read_dynamic_sheet,
    read_sheet,
    write_investment_output_workbook,
)

from finance_parser.investments.parsers import nordnet, seligson, evli, op_investment, coinmotion, nordea


PROJECT_ROOT = Path(__file__).resolve().parents[2]
INVESTMENT_INPUT_DIR = PROJECT_ROOT / "input" / "investments"
INVESTMENT_OUTPUT_DIR = PROJECT_ROOT / "output" / "investments"
INVESTMENT_RULES_DIR = PROJECT_ROOT / "rules" / "investments"
DEFAULT_INPUT_PATH = INVESTMENT_INPUT_DIR
DEFAULT_OUTPUT_WORKBOOK = INVESTMENT_OUTPUT_DIR / "ParsedInvestments.xlsx"
DEFAULT_INSTRUMENT_MASTER = INVESTMENT_RULES_DIR / "InstrumentMaster.xlsx"

SUPPORTED_INVESTMENT_PARSERS = [
    nordnet,
    seligson,
    evli,
    op_investment,
    coinmotion,
    nordea,
]


def collect_files(input_path: Path) -> list[Path]:
    if input_path.is_file():
        return [input_path]

    files = sorted([
        p for p in input_path.iterdir()
        if p.is_file()
        and p.suffix.lower() in {".xlsx", ".xlsm", ".xls", ".csv", ".txt", ".html", ".htm"}
        and not p.name.startswith("~$")
    ])

    if not files:
        raise FileNotFoundError(f"No supported investment files found in {input_path}")

    return files


def detect_parser(path: Path):
    matches = []
    reasons = []

    for parser_module in SUPPORTED_INVESTMENT_PARSERS:
        ok, reason = parser_module.can_parse(path)
        if ok:
            matches.append(parser_module)
        reasons.append(f"{parser_module.BROKER}: {reason}")

    if len(matches) == 1:
        return matches[0], "; ".join(reasons)

    if len(matches) == 0:
        raise ValueError(
            f"Could not identify investment export type for {path.name}.\n"
            + "\n".join(f"  - {r}" for r in reasons)
        )

    raise ValueError(
        f"Ambiguous investment parser match for {path.name}: "
        + ", ".join(m.BROKER for m in matches)
    )


def parse_import_files(
    input_path: Path,
    imported_at: str,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, pd.DataFrame]]:
    files = collect_files(input_path)
    import_run_id = make_import_run_id(imported_at)

    print("Investment parser modular v1")
    print("Files selected for investment import:")
    for file in files:
        print(f"  - {file.name}")

    parsed_frames = []
    broker_raw_frames: dict[str, list[pd.DataFrame]] = {}
    log_rows = []

    # Two distinct failure modes, mirroring the same fix on the budgeting
    # side (transaction_parser.py) - deliberately handled differently:
    #
    # 1. Unsupported file - no parser recognises it at all (or an ambiguous
    #    match). Almost always just an unrelated file in the input folder
    #    (real case: a manual reference workbook sitting next to real broker
    #    exports). Warn, skip that one file, keep processing everything else.
    #
    # 2. Format drift - a parser matched the file but then failed to
    #    actually parse it, meaning the broker's export format probably
    #    changed underneath us. Louder signal, but still only skips that one
    #    file - an unrelated broker's file having a problem shouldn't block
    #    everything else from importing.
    #
    # Previously both were lumped into one `errors` list that aborted the
    # ENTIRE run with nothing written, even for files that parsed fine -
    # confirmed harmful in practice (a stray non-broker file killed an
    # entire investment import batch).
    for file in files:
        try:
            parser_module, detection_reason = detect_parser(file)
        except ValueError as exc:
            print(f"WARNING: skipping {file.name} - {exc}")
            log_rows.append({
                "ImportRunID": import_run_id,
                "ImportedAt": imported_at,
                "Broker": "",
                "SourceFile": file.name,
                "RowsRead": 0,
                "RowsNew": "",
                "RowsDuplicate": "",
                "Status": "Skipped: unsupported file",
                "Error": str(exc),
            })
            continue

        try:
            parsed = parser_module.parse_file(file, imported_at)
            parsed_frames.append(parsed)

            if hasattr(parser_module, "parse_broker_raw_rows"):
                sheet_name = getattr(parser_module, "BROKER_RAW_SHEET", f"{parser_module.BROKER}RawExport")
                broker_raw = parser_module.parse_broker_raw_rows(file, imported_at)
                if len(broker_raw) > 0:
                    broker_raw_frames.setdefault(sheet_name, []).append(broker_raw)

            log_rows.append({
                "ImportRunID": import_run_id,
                "ImportedAt": imported_at,
                "Broker": parser_module.BROKER,
                "SourceFile": file.name,
                "RowsRead": len(parsed),
                "RowsNew": "",
                "RowsDuplicate": "",
                "Status": "Parsed",
                "Error": "",
            })

            print(f"Parsed {file.name}: {parser_module.BROKER}, {len(parsed)} rows")

        except Exception as exc:
            print(
                f"ERROR: {file.name} looks like a {parser_module.BROKER} export, but parsing it "
                f"failed - the format may have changed. Skipping this file.\n  {exc}"
            )
            log_rows.append({
                "ImportRunID": import_run_id,
                "ImportedAt": imported_at,
                "Broker": parser_module.BROKER,
                "SourceFile": file.name,
                "RowsRead": 0,
                "RowsNew": 0,
                "RowsDuplicate": 0,
                "Status": "Skipped: format drift",
                "Error": str(exc),
            })

    if parsed_frames:
        combined_raw = pd.concat(parsed_frames, ignore_index=True)
        combined_raw = combined_raw.drop_duplicates(subset=["InvestmentRawID"], keep="first")
    else:
        combined_raw = pd.DataFrame(columns=INVESTMENT_RAW_COLUMNS)

    combined_broker_raw = {
        sheet_name: pd.concat(frames, ignore_index=True)
        for sheet_name, frames in broker_raw_frames.items()
    }

    return (
        combined_raw[INVESTMENT_RAW_COLUMNS],
        pd.DataFrame(log_rows, columns=INVESTMENT_IMPORT_LOG_COLUMNS),
        combined_broker_raw,
    )


def _combine_dynamic_broker_raw_sheets(
    output_path: Path,
    imported_sheets: dict[str, pd.DataFrame],
) -> dict[str, pd.DataFrame]:
    combined: dict[str, pd.DataFrame] = {}

    for sheet_name, imported_df in imported_sheets.items():
        imported_df = imported_df.copy()
        imported_df.columns = [normalise_header(c) for c in imported_df.columns]

        existing_df = read_dynamic_sheet(output_path, sheet_name)

        if len(existing_df) == 0:
            merged = imported_df
        else:
            # Preserve columns from both old and new versions if the raw export
            # shape evolves.
            all_columns = list(existing_df.columns)
            for col in imported_df.columns:
                if col not in all_columns:
                    all_columns.append(col)

            for col in all_columns:
                if col not in existing_df.columns:
                    existing_df[col] = ""
                if col not in imported_df.columns:
                    imported_df[col] = ""

            merged = pd.concat([existing_df[all_columns], imported_df[all_columns]], ignore_index=True)

        if "BrokerRawExportID" in merged.columns:
            merged = merged.drop_duplicates(subset=["BrokerRawExportID"], keep="first")

        combined[sheet_name] = merged

    return combined


def append_to_output(input_path: Path, output_path: Path, instrument_master_path: Path | None) -> tuple[int, int, int]:
    if output_path.suffix.lower() == ".xlsm":
        raise ValueError(
            "Safety stop: do not write investment parser output to .xlsm. "
            "Use a separate .xlsx output, e.g. output/investments/ParsedInvestments.xlsx."
        )

    imported_at = imported_at_now()
    imported_raw, import_log_new, imported_broker_raw_sheets = parse_import_files(input_path, imported_at)

    existing_raw = read_sheet(output_path, INVESTMENT_RAW_SHEET, INVESTMENT_RAW_COLUMNS)
    existing_transactions = read_sheet(
        output_path,
        INVESTMENT_TRANSACTIONS_SHEET,
        INVESTMENT_TRANSACTIONS_COLUMNS,
    )
    existing_log = read_sheet(output_path, INVESTMENT_IMPORT_LOG_SHEET, INVESTMENT_IMPORT_LOG_COLUMNS)

    existing_raw_ids = set(existing_raw["InvestmentRawID"].dropna().astype(str))
    existing_transaction_raw_ids = set(existing_transactions["InvestmentRawID"].dropna().astype(str))

    new_raw = imported_raw[
        ~imported_raw["InvestmentRawID"].astype(str).isin(existing_raw_ids)
    ].copy()

    raw_needing_transactions = imported_raw[
        ~imported_raw["InvestmentRawID"].astype(str).isin(existing_transaction_raw_ids)
    ].copy()

    new_transactions = raw_to_investment_transactions(raw_needing_transactions)

    combined_raw = pd.concat([existing_raw, new_raw], ignore_index=True)
    combined_raw = combined_raw.drop_duplicates(subset=["InvestmentRawID"], keep="first")

    combined_transactions = pd.concat([existing_transactions, new_transactions], ignore_index=True)
    combined_transactions = combined_transactions.drop_duplicates(
        subset=["InvestmentTransactionID"],
        keep="first",
    )

    instrument_master = load_instrument_master(instrument_master_path)
    combined_transactions = apply_instrument_master(combined_transactions, instrument_master)
    combined_transactions = apply_portfolio_ownership(combined_transactions)

    rows_read = len(imported_raw)
    rows_new = len(new_raw)

    # Fill file-level new/duplicate counts into the log.
    #
    # Earlier versions wrote the whole-run RowsNew/RowsDuplicate totals into
    # every file's import-log row (same bug already fixed on the budgeting
    # side - see transaction_parser.py). That made every file in a run
    # appear to have contributed the same number of new/duplicate rows as
    # the run as a whole, which is impossible whenever a run touches more
    # than one file (e.g. a file with 77 RowsRead showing 679 duplicates).
    is_new_mask = ~imported_raw["InvestmentRawID"].astype(str).isin(existing_raw_ids)
    file_counts = (
        imported_raw.assign(_IsNew=is_new_mask)
        .groupby("SourceFile", dropna=False)["_IsNew"]
        .agg(["count", "sum"])
        .reset_index()
    )
    file_count_lookup = {
        str(row["SourceFile"]): {
            "RowsNew": int(row["sum"]),
            "RowsDuplicate": int(row["count"] - row["sum"]),
        }
        for _, row in file_counts.iterrows()
    }

    for idx, log_row in import_log_new.iterrows():
        source_file = str(log_row.get("SourceFile", ""))
        counts = file_count_lookup.get(source_file)
        if counts is None:
            # Real bug caught here: this used to be an unconditional
            # `import_log_new["Status"] = "Imported"` applied to every row,
            # which silently overwrote the "Skipped: unsupported file" /
            # "Skipped: format drift" statuses parse_import_files() sets for
            # files it never actually parsed - undoing the entire audit
            # trail those failure modes exist to provide. A file with no
            # entry in file_count_lookup was never successfully parsed
            # (contributed zero rows to imported_raw), so its Status must be
            # left exactly as parse_import_files() set it, not overwritten.
            continue
        import_log_new.at[idx, "RowsNew"] = counts["RowsNew"]
        import_log_new.at[idx, "RowsDuplicate"] = counts["RowsDuplicate"]
        import_log_new.at[idx, "Status"] = "Imported"

    combined_log = pd.concat([existing_log, import_log_new], ignore_index=True)

    combined_broker_raw_sheets = _combine_dynamic_broker_raw_sheets(
        output_path,
        imported_broker_raw_sheets,
    )

    write_investment_output_workbook(
        output_path,
        combined_raw[INVESTMENT_RAW_COLUMNS],
        combined_transactions[INVESTMENT_TRANSACTIONS_COLUMNS],
        combined_log[INVESTMENT_IMPORT_LOG_COLUMNS],
        extra_sheets=combined_broker_raw_sheets,
    )

    return rows_read, rows_new, len(new_transactions)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Modular investment transaction parser. Currently supports Nordnet, manually parsed Seligson, EVLI, OP, Coinmotion, and Nordea investment exports."
    )
    parser.add_argument(
        "--input",
        default=str(DEFAULT_INPUT_PATH),
        help="Investment export file or folder. Defaults to input/investments/.",
    )
    parser.add_argument(
        "--output",
        default=str(DEFAULT_OUTPUT_WORKBOOK),
        help="Separate investment output .xlsx file. Defaults to output/investments/ParsedInvestments.xlsx.",
    )
    parser.add_argument(
        "--instrument-master",
        default=str(DEFAULT_INSTRUMENT_MASTER),
        help="Instrument master workbook. Defaults to rules/investments/InstrumentMaster.xlsx.",
    )
    parser.add_argument(
        "--no-instrument-master",
        action="store_true",
        help="Skip InstrumentMaster enrichment even if the workbook exists.",
    )

    args = parser.parse_args()

    INVESTMENT_INPUT_DIR.mkdir(parents=True, exist_ok=True)
    INVESTMENT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    INVESTMENT_RULES_DIR.mkdir(parents=True, exist_ok=True)

    input_path = Path(args.input).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()

    instrument_master_path = None
    if not args.no_instrument_master:
        instrument_master_path = Path(args.instrument_master).expanduser().resolve()

    rows_read, rows_new, rows_transactions = append_to_output(
        input_path,
        output_path,
        instrument_master_path,
    )

    print()
    print("Investment import complete.")
    print(f"Rows found in input files:          {rows_read}")
    print(f"New raw investment rows appended:   {rows_new}")
    print(f"New investment transaction rows:    {rows_transactions}")
    print(f"Output workbook:                    {output_path}")
    if instrument_master_path is not None:
        if instrument_master_path.exists():
            print(f"Instrument master:                  {instrument_master_path}")
        else:
            print(f"Instrument master:                  not found ({instrument_master_path})")


if __name__ == "__main__":
    main()
