from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]

BUDGETING_INPUT_DIR = PROJECT_ROOT / "input" / "budgeting"
BUDGETING_OUTPUT_DIR = PROJECT_ROOT / "output" / "budgeting"
BUDGETING_RULES_DIR = PROJECT_ROOT / "rules" / "budgeting"

DEFAULT_INPUT_PATH = BUDGETING_INPUT_DIR
DEFAULT_WORKBOOK = BUDGETING_OUTPUT_DIR / "ParsedTransactions.xlsx"
DEFAULT_RULES = BUDGETING_RULES_DIR / "TransactionRules.xlsx"


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
    return parser


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

    normaliser_cmd = [
        python_executable,
        "-m",
        "finance_parser.budgeting.transaction_normaliser",
        "--workbook",
        str(workbook_path),
        "--rules",
        str(rules_path),
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

    return [
        PipelineCommand("transaction parser", parser_cmd),
        PipelineCommand("transaction normaliser", normaliser_cmd),
        PipelineCommand("transaction categoriser", categoriser_cmd),
    ]


def run_command(command: PipelineCommand) -> None:
    print()
    print(f"==> Running {command.name}")
    subprocess.run(command.argv, check=True)


def run_pipeline(args: argparse.Namespace) -> int:
    input_path = Path(args.input).expanduser().resolve()
    real_workbook = Path(args.workbook).expanduser().resolve()
    rules_path = Path(args.rules).expanduser().resolve()

    BUDGETING_INPUT_DIR.mkdir(parents=True, exist_ok=True)
    BUDGETING_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    BUDGETING_RULES_DIR.mkdir(parents=True, exist_ok=True)

    if args.keep_temp and not args.dry_run:
        raise ValueError("--keep-temp can only be used together with --dry-run")

    workbook_for_run = real_workbook
    temp_dir: Path | None = None

    if args.dry_run:
        workbook_for_run, temp_dir = prepare_dry_run_workbook(real_workbook)
        print("Budgeting pipeline dry-run.")
        print("The real workbook will not be modified.")
        print(f"Real workbook:      {real_workbook}")
        print(f"Temporary workbook: {workbook_for_run}")
    else:
        print("Budgeting pipeline.")
        print(f"Workbook: {real_workbook}")

    commands = build_pipeline_commands(
        python_executable=sys.executable,
        input_path=input_path,
        workbook_path=workbook_for_run,
        rules_path=rules_path,
        no_profile=args.no_profile,
        verbose=args.verbose,
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
    raise SystemExit(run_pipeline(args))


if __name__ == "__main__":
    main()
