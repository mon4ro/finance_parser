from __future__ import annotations

import argparse
from pathlib import Path
from time import perf_counter
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook

from finance_parser.common import (
    IMPORT_LOG_COLUMNS,
    RAW_COLUMNS,
    UNIFIED_COLUMNS,
    IMPORT_LOG_SHEET,
    RAW_SHEET,
    UNIFIED_SHEET,
    imported_at_now,
    add_canonical_transaction_key,
    make_import_run_id,
    raw_to_unified_rows,
    normalize_source_metadata,
    merge_blank_metadata_before_dedup,
    sheet_to_dataframe,
    write_clean_output_workbook,
    clean_for_excel,
    clean_dataframe_for_excel,
    normalise_header,
    append_change_log_entry,
)

from finance_parser.budgeting.parsers import op, norwegian, nordea, spankki, cash


PROJECT_ROOT = Path(__file__).resolve().parents[2]
BUDGETING_INPUT_DIR = PROJECT_ROOT / "input" / "budgeting"
BUDGETING_OUTPUT_DIR = PROJECT_ROOT / "output" / "budgeting"
BUDGETING_RULES_DIR = PROJECT_ROOT / "rules" / "budgeting"
DEFAULT_INPUT_PATH = BUDGETING_INPUT_DIR
DEFAULT_OUTPUT_WORKBOOK = BUDGETING_OUTPUT_DIR / "ParsedTransactions.xlsx"


SUPPORTED_PARSERS = [
    op,
    norwegian,
    nordea,
    spankki,
    cash,
]





def get_parser_modules() -> list:
    """
    Return the configured budgeting parser modules.

    Use this helper instead of referencing a hardcoded parser-list variable
    inside parsing logic. This prevents parser-list rename bugs such as
    referencing PARSERS when the actual list is SUPPORTED_PARSERS.
    """
    return SUPPORTED_PARSERS

def profile_log(enabled: bool, label: str, start_time: float) -> None:
    if enabled:
        print(f"[profile] {label}: {perf_counter() - start_time:.3f}s")


def verbose_log(enabled: bool, message: str) -> None:
    if enabled:
        print(message)


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
        raise FileNotFoundError(f"No supported files found in {input_path}")

    return files




def export_date_from_file(path: Path) -> str:
    """
    Return export timestamp metadata for an input file.

    For bank exports, the file modified timestamp is the safest common fallback.
    Parser modules may also store this in their parsed rows, but ImportLog needs
    the same value here.
    """
    return datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")

def detect_parser(path: Path):
    matches = []
    reasons = []

    for parser_module in SUPPORTED_PARSERS:
        ok, reason = parser_module.can_parse(path)
        if ok:
            matches.append(parser_module)
        reasons.append(f"{parser_module.SOURCE_BANK}: {reason}")

    if len(matches) == 1:
        return matches[0], "; ".join(reasons)

    if len(matches) == 0:
        raise ValueError(
            f"Could not identify bank/export type for {path.name}.\n"
            + "\n".join(f"  - {r}" for r in reasons)
        )

    raise ValueError(
        f"Ambiguous parser match for {path.name}: "
        + ", ".join(m.SOURCE_BANK for m in matches)
    )


def parse_import_files(
    input_path: Path,
    imported_at: str,
    profile: bool = True,
    verbose: bool = False,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, pd.DataFrame]]:
    t_collect = perf_counter()
    files = collect_files(input_path)
    profile_log(profile, "collect input files", t_collect)

    import_run_id = make_import_run_id(imported_at)

    verbose_log(verbose, "Transaction parser modular v1")
    verbose_log(verbose, f"Input:  {input_path}")
    verbose_log(verbose, f"Files:  {len(files)}")

    parsed_frames: list[pd.DataFrame] = []
    import_log_frames: list[pd.DataFrame] = []
    bank_raw_sheets: dict[str, list[pd.DataFrame]] = {}

    # Two distinct failure modes, deliberately handled differently:
    #
    # 1. Unsupported file - no parser recognises it at all. Almost always
    #    just an unrelated file that ended up in the input folder (real case:
    #    a manual reference workbook sitting next to real broker exports).
    #    Warn, skip that one file, keep processing everything else.
    #
    # 2. Format drift - a parser matched the file (by name/shape heuristics)
    #    but then failed to actually parse it, meaning the broker's export
    #    format probably changed underneath us. This is a much louder signal
    #    something needs human attention, but still only skips that one file
    #    - an unrelated broker's file having a real problem shouldn't block
    #    everything else from importing.
    #
    # Previously both cases were lumped into one `errors` list that aborted
    # the ENTIRE run with nothing written, even for files that parsed fine -
    # confirmed harmful in practice (a stray non-broker file killed an entire
    # investment import batch).
    unsupported_files: list[str] = []
    format_drift_files: list[str] = []

    t_parse = perf_counter()

    for file in files:
        matched_parser = None
        match_reasons = []

        for parser_module in get_parser_modules():
            ok, reason = parser_module.can_parse(file)
            if ok:
                matched_parser = parser_module
                break
            match_reasons.append(f"{parser_module.SOURCE_BANK}: {reason}")

        if matched_parser is None:
            reason_text = "\n".join(f"  - {reason}" for reason in match_reasons)
            print(f"WARNING: skipping {file.name} - no matching parser.\n{reason_text}")
            unsupported_files.append(file.name)
            import_log_frames.append(
                pd.DataFrame([{
                    "ImportRunID": import_run_id,
                    "SourceFile": file.name,
                    "SourceBank": "",
                    "ImportedAt": imported_at,
                    "ExportDate": export_date_from_file(file),
                    "RowsRead": 0,
                    "RowsNew": 0,
                    "RowsDuplicate": 0,
                    "Status": "Skipped: unsupported file",
                    "Notes": reason_text,
                }])
            )
            continue

        print(f"Parsing input file: {file.name} ({matched_parser.SOURCE_BANK})")

        try:
            parsed = matched_parser.parse_file(file, imported_at)
            parsed_frames.append(parsed)

            if hasattr(matched_parser, "parse_bank_raw_rows"):
                bank_raw = matched_parser.parse_bank_raw_rows(file, imported_at)
                sheet_name = getattr(matched_parser, "BANK_RAW_SHEET", f"{matched_parser.SOURCE_BANK}RawExport")
                bank_raw_sheets.setdefault(sheet_name, []).append(bank_raw)

            import_log_frames.append(
                pd.DataFrame([{
                    "ImportRunID": import_run_id,
                    "SourceFile": file.name,
                    "SourceBank": matched_parser.SOURCE_BANK,
                    "ImportedAt": imported_at,
                    "ExportDate": export_date_from_file(file),
                    "RowsRead": len(parsed),
                    "RowsNew": 0,
                    "RowsDuplicate": 0,
                    "Status": "Parsed",
                    "Notes": "",
                }])
            )
        except Exception as exc:
            error_text = f"{type(exc).__name__}: {exc}"
            print(
                f"ERROR: {file.name} looks like a {matched_parser.SOURCE_BANK} export, but parsing it "
                f"failed - the format may have changed. Skipping this file.\n  {error_text}"
            )
            format_drift_files.append(f"{file.name} ({matched_parser.SOURCE_BANK}): {error_text}")
            import_log_frames.append(
                pd.DataFrame([{
                    "ImportRunID": import_run_id,
                    "SourceFile": file.name,
                    "SourceBank": matched_parser.SOURCE_BANK,
                    "ImportedAt": imported_at,
                    "ExportDate": export_date_from_file(file),
                    "RowsRead": 0,
                    "RowsNew": 0,
                    "RowsDuplicate": 0,
                    "Status": "Skipped: format drift",
                    "Notes": error_text,
                }])
            )

    profile_log(profile, "parse input files", t_parse)

    non_empty_frames = [
        frame
        for frame in parsed_frames
        if frame is not None and len(frame) > 0
    ]

    if not non_empty_frames:
        problems = []
        if unsupported_files:
            problems.append("Unsupported files: " + ", ".join(unsupported_files))
        if format_drift_files:
            problems.append("Format drift: " + "; ".join(format_drift_files))
        raise ValueError(
            "No transaction rows were parsed from the input files.\n" + "\n".join(problems)
        )

    combined_raw = pd.concat(non_empty_frames, ignore_index=True)
    combined_raw = combined_raw.drop_duplicates(subset=["RawID"], keep="first")

    import_log = pd.concat(import_log_frames, ignore_index=True) if import_log_frames else pd.DataFrame(columns=IMPORT_LOG_COLUMNS)

    combined_bank_raw_sheets = {
        sheet_name: pd.concat(frames, ignore_index=True).drop_duplicates(subset=["BankRawExportID"], keep="first")
        for sheet_name, frames in bank_raw_sheets.items()
        if frames
    }

    return combined_raw, import_log, combined_bank_raw_sheets

def known_bank_raw_sheet_names() -> list[str]:
    names = []
    for parser_module in SUPPORTED_PARSERS:
        sheet_name = getattr(parser_module, "BANK_RAW_SHEET", "")
        if sheet_name:
            names.append(sheet_name)
    return names


def read_dynamic_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    """
    Read optional bank-specific raw-export sheets with their existing columns.

    These sheets intentionally retain a closer-to-export shape and therefore do
    not use the fixed RAW_COLUMNS schema.
    """
    if not path.exists():
        return pd.DataFrame()

    try:
        df = pd.read_excel(path, sheet_name=sheet_name, dtype=object, engine="openpyxl")
    except ValueError:
        return pd.DataFrame()

    df.columns = [normalise_header(c) for c in df.columns]
    return df


def combine_dynamic_bank_raw_sheets(
    output_path: Path,
    imported_sheets: dict[str, pd.DataFrame],
) -> dict[str, pd.DataFrame]:
    """
    Preserve existing bank-specific raw sheets and merge newly imported rows.

    This is needed because write_clean_output_workbook rewrites the standard
    sheets from scratch. Without this step, bank raw sheets from previous runs
    would disappear whenever a different bank is imported.
    """
    combined: dict[str, pd.DataFrame] = {}

    sheet_names = set(known_bank_raw_sheet_names()) | set(imported_sheets.keys())

    for sheet_name in sorted(sheet_names):
        existing_df = read_dynamic_sheet(output_path, sheet_name)
        imported_df = imported_sheets.get(sheet_name, pd.DataFrame()).copy()

        if len(existing_df) == 0 and len(imported_df) == 0:
            continue

        if len(existing_df) == 0:
            merged = imported_df
        elif len(imported_df) == 0:
            merged = existing_df
        else:
            existing_df = existing_df.copy()
            imported_df.columns = [normalise_header(c) for c in imported_df.columns]

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

        if "BankRawExportID" in merged.columns:
            merged = merged.drop_duplicates(subset=["BankRawExportID"], keep="first")

        combined[sheet_name] = merged

    return combined


def append_dataframe_to_worksheet(wb, sheet_name: str, df: pd.DataFrame) -> None:
    df = clean_dataframe_for_excel(df)

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]

    ws = wb.create_sheet(sheet_name)
    ws.append([clean_for_excel(c) for c in df.columns])

    for row in df.itertuples(index=False, name=None):
        ws.append([clean_for_excel(v) for v in row])

    ws.freeze_panes = "A2"

    if len(df) > 0 and len(df.columns) > 0:
        end_col = ws.cell(row=1, column=len(df.columns)).column_letter
        end_row = len(df) + 1
        ws.auto_filter.ref = f"A1:{end_col}{end_row}"


def write_extra_sheets_to_existing_workbook(
    workbook_path: Path,
    extra_sheets: dict[str, pd.DataFrame],
) -> None:
    if not extra_sheets:
        return

    wb = load_workbook(workbook_path)

    for sheet_name, df in extra_sheets.items():
        append_dataframe_to_worksheet(wb, sheet_name, df)

    wb.save(workbook_path)
    wb.close()


def append_to_output(
    input_path: Path,
    output_path: Path,
    profile: bool = True,
    dry_run: bool = False,
    verbose: bool = False,
) -> tuple[int, int, int]:
    if output_path.suffix.lower() == ".xlsm":
        raise ValueError(
            "Safety stop: do not write parser output to .xlsm. "
            "Use a separate .xlsx output, e.g. ParsedTransactions.xlsx."
        )

    imported_at = imported_at_now()
    imported_raw, import_log_new, imported_bank_raw_sheets = parse_import_files(
        input_path,
        imported_at,
        profile=profile,
        verbose=verbose,
    )

    t_read = perf_counter()
    existing_raw = sheet_to_dataframe(output_path, RAW_SHEET, RAW_COLUMNS)
    existing_unified = sheet_to_dataframe(output_path, UNIFIED_SHEET, UNIFIED_COLUMNS)
    existing_log = sheet_to_dataframe(output_path, IMPORT_LOG_SHEET, IMPORT_LOG_COLUMNS)
    profile_log(profile, "read existing workbook sheets", t_read)

    t_dedupe = perf_counter()

    existing_raw_ids = set(existing_raw["RawID"].dropna().astype(str))
    existing_unified_raw_ids = set(existing_unified["RawID"].dropna().astype(str))

    existing_raw_keyed = add_canonical_transaction_key(existing_raw)
    imported_raw_keyed = add_canonical_transaction_key(imported_raw)
    existing_keys = set(existing_raw_keyed["_CanonicalTransactionKey"].dropna().astype(str))

    # Reporting count: genuinely new means neither exact RawID nor canonical key
    # existed before this run.
    new_mask = (
        ~imported_raw_keyed["RawID"].astype(str).isin(existing_raw_ids)
        & ~imported_raw_keyed["_CanonicalTransactionKey"].astype(str).isin(existing_keys)
    )
    new_raw = imported_raw_keyed.loc[new_mask].drop(columns=["_CanonicalTransactionKey"]).copy()

    # Important: combine existing rows with ALL imported rows, not only new rows.
    # Otherwise an older preserved duplicate row with blank ExportDate/ImportedAt
    # never gets a chance to copy metadata from the freshly parsed duplicate.
    imported_raw_for_merge = imported_raw_keyed.drop(columns=["_CanonicalTransactionKey"]).copy()

    combined_raw = pd.concat([existing_raw, imported_raw_for_merge], ignore_index=True)
    combined_raw = add_canonical_transaction_key(combined_raw)
    combined_raw = merge_blank_metadata_before_dedup(combined_raw)
    combined_raw = combined_raw.drop_duplicates(subset=["_CanonicalTransactionKey"], keep="first")
    combined_raw = combined_raw.drop(columns=["_CanonicalTransactionKey"])

    # Create candidate Unified rows for all imported rows whose exact Unified RawID
    # does not already exist. Duplicate cleanup keeps existing/manual rows while
    # merge_blank_metadata_before_dedup copies missing technical metadata from the
    # fresh candidate rows.
    raw_needing_unified = imported_raw_for_merge[
        ~imported_raw_for_merge["RawID"].astype(str).isin(existing_unified_raw_ids)
    ].copy()

    new_unified_candidates = raw_to_unified_rows(raw_needing_unified)

    # Reporting count: only genuinely new canonical rows are counted as new unified rows.
    new_unified = raw_to_unified_rows(new_raw)

    combined_unified = pd.concat([existing_unified, new_unified_candidates], ignore_index=True)
    combined_unified = add_canonical_transaction_key(combined_unified)
    combined_unified = merge_blank_metadata_before_dedup(combined_unified)
    combined_unified = combined_unified.drop_duplicates(subset=["_CanonicalTransactionKey"], keep="first")
    combined_unified = combined_unified.drop(columns=["_CanonicalTransactionKey"])

    # Do not de-duplicate solely by UnifiedID.
    #
    # Manual split rows are commonly created by copying an existing unified row,
    # editing Amount and manual budget fields. Such rows may intentionally share
    # the original UnifiedID unless the user later gives them split-specific IDs.
    # Canonical-key de-duplication above is enough for imported duplicate cleanup;
    # UnifiedID-only de-duplication can accidentally delete manual split rows.

    # Fill file-level new/duplicate counts into the log.
    #
    # Earlier versions wrote the whole-run RowsNew/RowsDuplicate totals into
    # every file's import-log row. That made old files appear to have the same
    # number of new rows as the one file that actually introduced new rows.
    rows_read = len(imported_raw)
    rows_new = len(new_raw)
    rows_duplicate = rows_read - rows_new

    file_counts = (
        imported_raw_keyed.assign(_IsNewCanonicalRow=new_mask)
        .groupby("SourceFile", dropna=False)["_IsNewCanonicalRow"]
        .agg(["count", "sum"])
        .reset_index()
    )
    file_count_lookup = {
        str(row["SourceFile"]): {
            "RowsRead": int(row["count"]),
            "RowsNew": int(row["sum"]),
            "RowsDuplicate": int(row["count"] - row["sum"]),
        }
        for _, row in file_counts.iterrows()
    }

    for idx, log_row in import_log_new.iterrows():
        source_file = str(log_row.get("SourceFile", ""))
        counts = file_count_lookup.get(source_file)

        if counts is None:
            continue

        import_log_new.at[idx, "RowsRead"] = counts["RowsRead"]
        import_log_new.at[idx, "RowsNew"] = counts["RowsNew"]
        import_log_new.at[idx, "RowsDuplicate"] = counts["RowsDuplicate"]
        import_log_new.at[idx, "Status"] = "Imported"

    combined_log = pd.concat([existing_log, import_log_new], ignore_index=True)

    # Final metadata normalisation is important after duplicate cleanup because
    # the row kept may be an older row with blank SourceBank/SourceAccount.
    combined_raw = normalize_source_metadata(combined_raw)
    combined_unified = normalize_source_metadata(combined_unified)
    profile_log(profile, "deduplicate and prepare output data", t_dedupe)

    combined_bank_raw_sheets = combine_dynamic_bank_raw_sheets(
        output_path,
        imported_bank_raw_sheets,
    )

    if dry_run:
        return rows_read, rows_new, len(new_unified)

    t_write = perf_counter()
    write_clean_output_workbook(
        output_path,
        combined_raw[RAW_COLUMNS],
        combined_unified[UNIFIED_COLUMNS],
        combined_log[IMPORT_LOG_COLUMNS],
    )
    profile_log(profile, "write canonical workbook sheets", t_write)

    t_write_extra = perf_counter()
    write_extra_sheets_to_existing_workbook(
        output_path,
        combined_bank_raw_sheets,
    )

    profile_log(profile, "write bank raw sheets", t_write_extra)

    t_changelog = perf_counter()
    append_change_log_entry(
        output_path,
        script="transaction_parser.py",
        action="Import budgeting transactions",
        sheet="RawTransactions, UnifiedTransactions, ImportLog, bank raw sheets",
        rows_before=len(existing_raw),
        rows_after=len(combined_raw),
        rows_added=rows_new,
        rows_updated="",
        rows_removed="",
        status="Completed",
        details=f"Rows read: {rows_read}; raw rows added: {rows_new}; new unified rows: {len(new_unified)}",
    )

    profile_log(profile, "append ChangeLog entry", t_changelog)

    return rows_read, rows_new, len(new_unified)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Modular transaction parser. Currently supports OP, Bank Norwegian, Nordea, and S-Pankki."
    )
    parser.add_argument(
        "--input",
        default=str(DEFAULT_INPUT_PATH),
        help="Import file or folder. Defaults to input/budgeting/.",
    )
    parser.add_argument(
        "--profile",
        dest="profile",
        action="store_true",
        default=True,
        help="Print simple timing information for parser phases. Enabled by default.",
    )
    parser.add_argument(
        "--no-profile",
        dest="profile",
        action="store_false",
        help="Disable timing information.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Parse input files and report what would change, but do not write the workbook.",
    )
    parser.add_argument(
        "--verbose",
        "--debug",
        dest="verbose",
        action="store_true",
        help="Print extra diagnostic information. Normal output only lists parsed input files and summary counts.",
    )
    parser.add_argument(
        "--output",
        default=str(DEFAULT_OUTPUT_WORKBOOK),
        help="Separate budgeting output .xlsx file. Defaults to output/budgeting/ParsedTransactions.xlsx.",
    )
    return parser


def main() -> None:
    parser = build_arg_parser()
    args = parser.parse_args()

    BUDGETING_INPUT_DIR.mkdir(parents=True, exist_ok=True)
    BUDGETING_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    BUDGETING_RULES_DIR.mkdir(parents=True, exist_ok=True)

    input_path = Path(args.input).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()

    rows_read, rows_new, rows_unified = append_to_output(
        input_path,
        output_path,
        profile=args.profile,
        dry_run=args.dry_run,
        verbose=args.verbose,
    )

    print()
    print("Import dry run complete." if args.dry_run else "Import complete.")
    print(f"Rows found in input files:       {rows_read}")
    print(f"New raw rows appended:           {rows_new}")
    print(f"New UnifiedTransactions rows:    {rows_unified}")
    print(f"Output workbook:                 {output_path}")
    if args.dry_run:
        print("Dry run only: workbook was not modified.")


if __name__ == "__main__":
    main()
