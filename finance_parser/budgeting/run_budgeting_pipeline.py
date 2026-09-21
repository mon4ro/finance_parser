from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

from finance_parser.common import CHANGE_LOG_COLUMNS, CHANGE_LOG_SHEET


PROJECT_ROOT = Path(__file__).resolve().parents[2]

BUDGETING_INPUT_DIR = PROJECT_ROOT / "input" / "budgeting"
BUDGETING_OUTPUT_DIR = PROJECT_ROOT / "output" / "budgeting"
BUDGETING_RULES_DIR = PROJECT_ROOT / "rules" / "budgeting"

DEFAULT_INPUT_PATH = BUDGETING_INPUT_DIR
DEFAULT_WORKBOOK = BUDGETING_OUTPUT_DIR / "ParsedTransactions.xlsx"
DEFAULT_RULES = BUDGETING_RULES_DIR / "TransactionRules.xlsx"
DEFAULT_BALANCE_WORKBOOK = BUDGETING_OUTPUT_DIR / "MonthlyAccountBalance.xlsx"
# Read-only cross-pipeline input, not owned by this pipeline - see
# investment_dividends.py (Nordnet/EVLI synthetic dividend income, no
# matching bank transaction exists) and enrich_dividend_income.py (OP-held
# instruments, cross-referenced against an existing real bank transaction).
DEFAULT_DIVIDEND_HISTORY = PROJECT_ROOT / "output" / "investments" / "DividendHistory.xlsx"


@dataclass(frozen=True)
class PipelineCommand:
    name: str
    argv: list[str]


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Run the full budgeting import pipeline: parser, normaliser, "
            "and categoriser."
        )
    )
    parser.add_argument(
        "--input",
        default=str(DEFAULT_INPUT_PATH),
        help="Import file or folder. Defaults to input/budgeting/.",
    )
    parser.add_argument(
        "--workbook",
        default=str(DEFAULT_WORKBOOK),
        help="Budgeting output workbook. Defaults to output/budgeting/ParsedTransactions.xlsx.",
    )
    parser.add_argument(
        "--rules",
        default=str(DEFAULT_RULES),
        help="Transaction rules workbook. Defaults to rules/budgeting/TransactionRules.xlsx.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help=(
            "Run the full pipeline against a temporary copied workbook. "
            "The real workbook is not modified."
        ),
    )
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        help=(
            "With --dry-run, keep the temporary workbook/folder for inspection "
            "instead of deleting it after a successful run."
        ),
    )
    parser.add_argument(
        "--no-profile",
        action="store_true",
        help="Disable transaction parser timing output.",
    )
    parser.add_argument(
        "--verbose",
        "--debug",
        dest="verbose",
        action="store_true",
        help="Pass verbose/debug output to pipeline steps that support it.",
    )
    parser.add_argument(
        "--dividend-history",
        default=str(DEFAULT_DIVIDEND_HISTORY),
        help="Read-only: investment pipeline's DividendHistory.xlsx, source for real dividend "
             "income (Nordnet/EVLI synthetic rows, OP-held cross-reference enrichment).",
    )
    parser.add_argument(
        "--skip-dividends",
        action="store_true",
        help="Skip both dividend-income stages (useful if the investment pipeline hasn't been "
             "run yet, or DividendHistory.xlsx doesn't exist).",
    )
    parser.add_argument(
        "--balance-workbook",
        default=str(DEFAULT_BALANCE_WORKBOOK),
        help="Monthly account balance output. Defaults to output/budgeting/MonthlyAccountBalance.xlsx.",
    )
    parser.add_argument(
        "--skip-account-balance",
        action="store_true",
        help="Skip the monthly account balance stage.",
    )
    parser.add_argument(
        "--show-changelog",
        action="store_true",
        help="Print the workbook's ChangeLog (what every pipeline stage did: rows read/added/"
             "updated, status, details) and exit. Read-only - does not run the pipeline.",
    )
    parser.add_argument(
        "--last",
        type=int,
        default=20,
        help="With --show-changelog, how many of the most recent entries to show. Ignored if "
             "--all is also given. Default: 20.",
    )
    parser.add_argument(
        "--all",
        dest="show_all_changelog",
        action="store_true",
        help="With --show-changelog, show every entry instead of just the last N.",
    )
    parser.add_argument(
        "--script",
        default=None,
        help="With --show-changelog, only show entries from this Script value (exact match).",
    )
    return parser


def _clean_cell(value: object) -> str:
    """
    A blank Excel cell reads back via pandas as a NaN float, not an empty
    string - str(nan) renders the literal text "nan", and "nan or ''" is
    still truthy (only 0.0 is falsy for floats), so a naive `str(x or "")`
    fallback never catches it. Real bug this fixed: both the ChangedAt
    timestamp and the Details column rendered the literal word "nan"
    instead of blank for the several ChangeLog rows that don't stamp them
    (only the transaction_parser.py import stage does).
    """
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def print_changelog(
    workbook: Path,
    *,
    last: int = 20,
    show_all: bool = False,
    script: str | None = None,
) -> None:
    """
    Read-only summary of a budgeting workbook's ChangeLog sheet - one row
    per pipeline-stage run (transaction_parser.py, transaction_normaliser.py,
    transaction_categoriser.py, enrich_dividend_income.py, and any other
    script that calls append_changelog_row()/append_change_log_entry()).

    Shown in the sheet's own append order (oldest first) rather than
    reversed, deliberately: several stages append a row with a blank
    ChangedAt/ChangeID (only the transaction_parser.py import stage stamps
    those) immediately after the real-timestamped row for the same run, so
    natural order keeps "which run did this blank-timestamp row belong to"
    visually obvious - reversing the list would separate them.
    """
    if not workbook.exists():
        print(f"{workbook} does not exist - nothing to show.")
        return

    try:
        df = pd.read_excel(workbook, sheet_name=CHANGE_LOG_SHEET, dtype=object)
    except (ValueError, KeyError):
        print(f"No {CHANGE_LOG_SHEET} sheet found in {workbook}.")
        return

    df.columns = [str(c).strip() for c in df.columns]
    for col in CHANGE_LOG_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    if script:
        df = df[df["Script"].astype(str) == script]

    if df.empty:
        print("No ChangeLog entries match.")
        return

    matching_count = len(df)
    if not show_all:
        df = df.tail(last)

    print(f"ChangeLog for {workbook} ({len(df)} of {matching_count} matching entries shown):")
    print("-" * 100)
    for _, row in df.iterrows():
        timestamp = _clean_cell(row["ChangedAt"]) or "(same run as above)"
        rows_summary_parts = []
        for label, col in (("before", "RowsBefore"), ("after", "RowsAfter"), ("added", "RowsAdded"), ("updated", "RowsUpdated"), ("removed", "RowsRemoved")):
            value = _clean_cell(row.get(col))
            if value:
                rows_summary_parts.append(f"{label}={value}")
        rows_summary = ", ".join(rows_summary_parts)

        print(f"[{timestamp}] {_clean_cell(row['Script'])} - {_clean_cell(row['Action'])}")
        if rows_summary:
            print(f"    rows: {rows_summary}")
        print(f"    status: {_clean_cell(row['Status'])}")
        details = _clean_cell(row.get("Details"))
        if details:
            print(f"    details: {details}")
        print()


def timestamp_for_path() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def prepare_dry_run_workbook(real_workbook: Path, *, timestamp: str | None = None) -> tuple[Path, Path]:
    """
    Prepare a disposable workbook path for a true full-pipeline dry-run.

    The real workbook is copied when it exists. If it does not exist yet, the
    parser step can create the temporary workbook from scratch.
    """
    stamp = timestamp or timestamp_for_path()
    temp_dir = real_workbook.parent / "_pipeline_dry_run" / stamp
    temp_dir.mkdir(parents=True, exist_ok=True)

    temp_workbook = temp_dir / real_workbook.name

    if real_workbook.exists():
        shutil.copy2(real_workbook, temp_workbook)

    return temp_workbook, temp_dir


def build_pipeline_commands(
    *,
    python_executable: str,
    input_path: Path,
    workbook_path: Path,
    rules_path: Path,
    no_profile: bool = False,
    verbose: bool = False,
    dividend_history_path: Path | None = None,
    skip_dividends: bool = False,
    balance_workbook_path: Path | None = None,
    skip_account_balance: bool = False,
) -> list[PipelineCommand]:
    parser_cmd = [
        python_executable,
        "-m",
        "finance_parser.budgeting.transaction_parser",
        "--input",
        str(input_path),
        "--output",
        str(workbook_path),
    ]

    if no_profile:
        parser_cmd.append("--no-profile")

    if verbose:
        parser_cmd.append("--verbose")

    # Reuses transaction_parser.py against a single cross-pipeline file
    # (not input/budgeting/) - investment_dividends.py only claims this one
    # specific file shape, so it's safe to run as its own targeted step.
    dividend_parser_cmd = [
        python_executable,
        "-m",
        "finance_parser.budgeting.transaction_parser",
        "--input",
        str(dividend_history_path),
        "--output",
        str(workbook_path),
    ]
    if no_profile:
        dividend_parser_cmd.append("--no-profile")

    normaliser_cmd = [
        python_executable,
        "-m",
        "finance_parser.budgeting.transaction_normaliser",
        "--workbook",
        str(workbook_path),
        "--rules",
        str(rules_path),
    ]

    # Runs BEFORE the categoriser so its high-confidence, cross-referenced
    # values win by default (categoriser rules default to BLANK_ONLY, so
    # they'll correctly skip cells this step already filled).
    dividend_enrich_cmd = [
        python_executable,
        "-m",
        "finance_parser.budgeting.enrich_dividend_income",
        "--workbook",
        str(workbook_path),
        "--dividend-history",
        str(dividend_history_path),
    ]

    categoriser_cmd = [
        python_executable,
        "-m",
        "finance_parser.budgeting.transaction_categoriser",
        "--workbook",
        str(workbook_path),
        "--rules",
        str(rules_path),
    ]

    balance_cmd = [
        python_executable,
        "-m",
        "finance_parser.budgeting.build_monthly_account_balance",
        "--budgeting-workbook",
        str(workbook_path),
        "--output-workbook",
        str(balance_workbook_path),
    ]

    run_dividend_stages = not skip_dividends and dividend_history_path is not None

    commands = [PipelineCommand("transaction parser", parser_cmd)]
    if run_dividend_stages:
        commands.append(PipelineCommand("investment dividend parser", dividend_parser_cmd))
    commands.append(PipelineCommand("transaction normaliser", normaliser_cmd))
    if run_dividend_stages:
        commands.append(PipelineCommand("dividend income enrichment", dividend_enrich_cmd))
    commands.append(PipelineCommand("transaction categoriser", categoriser_cmd))
    if not skip_account_balance and balance_workbook_path is not None:
        commands.append(PipelineCommand("monthly account balance", balance_cmd))
    return commands


def run_command(command: PipelineCommand) -> None:
    print()
    print(f"==> Running {command.name}")
    subprocess.run(command.argv, check=True)


def run_pipeline(args: argparse.Namespace) -> int:
    input_path = Path(args.input).expanduser().resolve()
    real_workbook = Path(args.workbook).expanduser().resolve()
    rules_path = Path(args.rules).expanduser().resolve()
    real_balance_workbook = Path(args.balance_workbook).expanduser().resolve()

    BUDGETING_INPUT_DIR.mkdir(parents=True, exist_ok=True)
    BUDGETING_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    BUDGETING_RULES_DIR.mkdir(parents=True, exist_ok=True)

    if args.keep_temp and not args.dry_run:
        raise ValueError("--keep-temp can only be used together with --dry-run")

    workbook_for_run = real_workbook
    balance_workbook_for_run = real_balance_workbook
    temp_dir: Path | None = None

    if args.dry_run:
        workbook_for_run, temp_dir = prepare_dry_run_workbook(real_workbook)
        balance_workbook_for_run = temp_dir / real_balance_workbook.name
        if real_balance_workbook.exists():
            shutil.copy2(real_balance_workbook, balance_workbook_for_run)
        print("Budgeting pipeline dry-run.")
        print("The real workbook will not be modified.")
        print(f"Real workbook:      {real_workbook}")
        print(f"Temporary workbook: {workbook_for_run}")
    else:
        print("Budgeting pipeline.")
        print(f"Workbook: {real_workbook}")

    dividend_history_path = Path(args.dividend_history).expanduser().resolve()
    skip_dividends = args.skip_dividends
    if not skip_dividends and not dividend_history_path.exists():
        print()
        print(f"Note: dividend history not found at {dividend_history_path} - skipping dividend-income stages.")
        print("Run the investment pipeline first (run_investment_pipeline.py) to enable them.")
        skip_dividends = True

    commands = build_pipeline_commands(
        python_executable=sys.executable,
        input_path=input_path,
        workbook_path=workbook_for_run,
        rules_path=rules_path,
        no_profile=args.no_profile,
        verbose=args.verbose,
        dividend_history_path=dividend_history_path,
        skip_dividends=skip_dividends,
        balance_workbook_path=balance_workbook_for_run,
        skip_account_balance=args.skip_account_balance,
    )

    try:
        for command in commands:
            run_command(command)
    except Exception:
        if args.dry_run and temp_dir is not None:
            print()
            print("Dry-run failed. Temporary files were kept for inspection:")
            print(temp_dir)
        raise

    print()
    if args.dry_run:
        print("Budgeting pipeline dry-run complete.")
        if args.keep_temp:
            print("Temporary files kept for inspection:")
            print(temp_dir)
        elif temp_dir is not None:
            shutil.rmtree(temp_dir)
            print("Temporary files deleted.")
        print("Real workbook was not modified.")
    else:
        print("Budgeting pipeline complete.")
        print(f"Updated workbook: {real_workbook}")

    return 0


def main() -> None:
    parser = build_arg_parser()
    args = parser.parse_args()

    if args.show_changelog:
        print_changelog(
            Path(args.workbook),
            last=args.last,
            show_all=args.show_all_changelog,
            script=args.script,
        )
        return

    raise SystemExit(run_pipeline(args))


if __name__ == "__main__":
    main()
