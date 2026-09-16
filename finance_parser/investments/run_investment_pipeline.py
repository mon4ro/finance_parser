from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from time import perf_counter


PROJECT_ROOT = Path(__file__).resolve().parents[2]

INVESTMENT_INPUT_DIR = PROJECT_ROOT / "input" / "investments"
INVESTMENT_OUTPUT_DIR = PROJECT_ROOT / "output" / "investments"
INVESTMENT_RULES_DIR = PROJECT_ROOT / "rules" / "investments"

DEFAULT_INPUT_PATH = INVESTMENT_INPUT_DIR
DEFAULT_INSTRUMENT_MASTER = INVESTMENT_RULES_DIR / "InstrumentMaster.xlsx"
DEFAULT_INVESTMENTS_WORKBOOK = INVESTMENT_OUTPUT_DIR / "ParsedInvestments.xlsx"
DEFAULT_PRICES_WORKBOOK = INVESTMENT_OUTPUT_DIR / "InstrumentPrices.xlsx"
DEFAULT_FX_RATES_WORKBOOK = INVESTMENT_OUTPUT_DIR / "FXRates.xlsx"
DEFAULT_POSITIONS_WORKBOOK = INVESTMENT_OUTPUT_DIR / "PortfolioPositions.xlsx"
DEFAULT_MONTHLY_VALUE_WORKBOOK = INVESTMENT_OUTPUT_DIR / "MonthlyPositionValue.xlsx"
DEFAULT_DIVIDEND_HISTORY_WORKBOOK = INVESTMENT_OUTPUT_DIR / "DividendHistory.xlsx"
# Read-only cross-pipeline input, not part of this pipeline's own
# dry-run-copied workbooks - see build_dividend_history.py's
# --budgeting-workbook (best-effort local-currency dividend enrichment).
DEFAULT_BUDGETING_WORKBOOK = PROJECT_ROOT / "output" / "budgeting" / "ParsedTransactions.xlsx"

DEFAULT_STALE_DAYS = 7

# Six pipeline-owned output files, each written by exactly one stage. These
# get dry-run temp-copied. input/investments/ and InstrumentMaster.xlsx are
# read-only inputs to every stage (never written by the pipeline itself) and
# are never copied - same distinction the budgeting orchestrator already
# makes between --workbook (copied) and --rules (not copied).
WORKBOOK_KEYS = ["investments", "prices", "fx_rates", "positions", "monthly_value", "dividend_history"]


@dataclass(frozen=True)
class PipelineCommand:
    name: str
    argv: list[str]
    # If True, a non-zero exit is the coverage gate saying "resolve this
    # first" (or --force overriding it) rather than a real crash - still
    # propagated as a hard stop either way, but worth a clearer message.
    is_gate: bool = False


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Run the full investment pipeline: parse raw exports, fetch prices, fetch FX rates, "
            "build positions, check coverage, build the monthly value rollup."
        )
    )
    parser.add_argument("--input", default=str(DEFAULT_INPUT_PATH), help="Import file or folder. Defaults to input/investments/.")
    parser.add_argument("--instrument-master", default=str(DEFAULT_INSTRUMENT_MASTER))
    parser.add_argument("--investments-workbook", default=str(DEFAULT_INVESTMENTS_WORKBOOK))
    parser.add_argument("--prices-workbook", default=str(DEFAULT_PRICES_WORKBOOK))
    parser.add_argument("--fx-rates-workbook", default=str(DEFAULT_FX_RATES_WORKBOOK))
    parser.add_argument("--positions-workbook", default=str(DEFAULT_POSITIONS_WORKBOOK))
    parser.add_argument("--monthly-value-workbook", default=str(DEFAULT_MONTHLY_VALUE_WORKBOOK))
    parser.add_argument("--dividend-history-workbook", default=str(DEFAULT_DIVIDEND_HISTORY_WORKBOOK))
    parser.add_argument(
        "--budgeting-workbook",
        default=str(DEFAULT_BUDGETING_WORKBOOK),
        help="Read-only: source of local-currency dividend detail (e.g. Telia's SEK amount). "
             "Best-effort - missing file or non-matching rows just leave those columns blank.",
    )
    parser.add_argument(
        "--no-local-currency",
        action="store_true",
        help="Skip the local-currency dividend enrichment entirely.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Run the full pipeline against temporary copied workbooks. The real workbooks are not modified.",
    )
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        help="With --dry-run, keep the temporary workbooks/folder for inspection instead of deleting them after a successful run.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Pass through to the coverage check: report gaps but continue anyway instead of halting the pipeline.",
    )
    parser.add_argument(
        "--stale-days",
        type=int,
        default=DEFAULT_STALE_DAYS,
        help="Pass through to the coverage check: flag a price as stale if older than this many days.",
    )
    parser.add_argument(
        "--skip-pricing",
        action="store_true",
        help="Skip the instrument price fetch stage (network call). For offline dev/debug iteration on "
             "positions/coverage/rollup logic - downstream stages use whatever price data already exists.",
    )
    parser.add_argument(
        "--skip-fx",
        action="store_true",
        help="Skip the FX rate fetch stage (network call). Same use case as --skip-pricing.",
    )
    return parser


def timestamp_for_path() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def prepare_dry_run_workbooks(real_workbooks: dict[str, Path], *, timestamp: str | None = None) -> tuple[dict[str, Path], Path]:
    """
    Prepare disposable copies of every pipeline-owned output file for a true
    full-pipeline dry-run. A workbook that doesn't exist yet (e.g. a brand
    new deployment) is simply not copied - the relevant stage creates it
    fresh in the temp dir, same as a real first run would.
    """
    stamp = timestamp or timestamp_for_path()
    temp_dir = INVESTMENT_OUTPUT_DIR / "_pipeline_dry_run" / stamp
    temp_dir.mkdir(parents=True, exist_ok=True)

    temp_workbooks: dict[str, Path] = {}
    for key, real_path in real_workbooks.items():
        temp_path = temp_dir / real_path.name
        if real_path.exists():
            shutil.copy2(real_path, temp_path)
        temp_workbooks[key] = temp_path

    return temp_workbooks, temp_dir


def build_pipeline_commands(
    *,
    python_executable: str,
    input_path: Path,
    instrument_master_path: Path,
    workbooks: dict[str, Path],
    force: bool,
    stale_days: int,
    skip_pricing: bool = False,
    skip_fx: bool = False,
    budgeting_workbook: Path | None = None,
    no_local_currency: bool = False,
) -> list[PipelineCommand]:
    parse_cmd = [
        python_executable, "-m", "finance_parser.investments.investment_parser",
        "--input", str(input_path),
        "--output", str(workbooks["investments"]),
        "--instrument-master", str(instrument_master_path),
    ]

    prices_cmd = [
        python_executable, "-m", "finance_parser.investments.fetch_instrument_prices",
        "--instrument-master", str(instrument_master_path),
        "--investments-workbook", str(workbooks["investments"]),
        "--prices-workbook", str(workbooks["prices"]),
    ]

    fx_cmd = [
        python_executable, "-m", "finance_parser.investments.fetch_fx_rates",
        "--investments-workbook", str(workbooks["investments"]),
        "--fx-rates-workbook", str(workbooks["fx_rates"]),
    ]

    positions_cmd = [
        python_executable, "-m", "finance_parser.investments.build_portfolio_positions",
        "--investments-workbook", str(workbooks["investments"]),
        "--instrument-master", str(instrument_master_path),
        "--positions-workbook", str(workbooks["positions"]),
    ]

    coverage_cmd = [
        python_executable, "-m", "finance_parser.investments.check_instrument_coverage",
        "--positions-workbook", str(workbooks["positions"]),
        "--instrument-master", str(instrument_master_path),
        "--prices-workbook", str(workbooks["prices"]),
        "--stale-days", str(stale_days),
    ]
    if force:
        coverage_cmd.append("--force")

    dividend_history_cmd = [
        python_executable, "-m", "finance_parser.investments.build_dividend_history",
        "--investments-workbook", str(workbooks["investments"]),
        "--output-workbook", str(workbooks["dividend_history"]),
    ]
    if no_local_currency:
        dividend_history_cmd.append("--no-local-currency")
    elif budgeting_workbook is not None:
        dividend_history_cmd.extend(["--budgeting-workbook", str(budgeting_workbook)])

    rollup_cmd = [
        python_executable, "-m", "finance_parser.investments.build_monthly_position_value",
        "--positions-workbook", str(workbooks["positions"]),
        "--prices-workbook", str(workbooks["prices"]),
        "--fx-rates-workbook", str(workbooks["fx_rates"]),
        "--investments-workbook", str(workbooks["investments"]),
        "--output-workbook", str(workbooks["monthly_value"]),
    ]

    commands = [PipelineCommand("investment parser", parse_cmd)]
    if not skip_pricing:
        commands.append(PipelineCommand("instrument price fetch", prices_cmd))
    if not skip_fx:
        commands.append(PipelineCommand("FX rate fetch", fx_cmd))
    commands.append(PipelineCommand("portfolio positions", positions_cmd))
    commands.append(PipelineCommand("dividend history", dividend_history_cmd))
    commands.append(PipelineCommand("instrument coverage check", coverage_cmd, is_gate=True))
    commands.append(PipelineCommand("monthly position value rollup", rollup_cmd))
    return commands


def run_command(command: PipelineCommand) -> float:
    print()
    print(f"==> Running {command.name}")
    start = perf_counter()
    try:
        subprocess.run(command.argv, check=True)
    except subprocess.CalledProcessError:
        if command.is_gate:
            print()
            print(
                "Coverage check failed: some active positions are missing classification, a price "
                "symbol, or a recent price. Resolve them in InstrumentMaster.xlsx, or rerun with "
                "--force to proceed with incomplete net worth data anyway."
            )
        raise
    elapsed = perf_counter() - start
    print(f"    ({command.name} done in {elapsed:.1f}s)")
    return elapsed


def run_pipeline(args: argparse.Namespace) -> int:
    input_path = Path(args.input).expanduser().resolve()
    instrument_master_path = Path(args.instrument_master).expanduser().resolve()
    real_workbooks = {
        "investments": Path(args.investments_workbook).expanduser().resolve(),
        "prices": Path(args.prices_workbook).expanduser().resolve(),
        "fx_rates": Path(args.fx_rates_workbook).expanduser().resolve(),
        "positions": Path(args.positions_workbook).expanduser().resolve(),
        "monthly_value": Path(args.monthly_value_workbook).expanduser().resolve(),
        "dividend_history": Path(args.dividend_history_workbook).expanduser().resolve(),
    }

    INVESTMENT_INPUT_DIR.mkdir(parents=True, exist_ok=True)
    INVESTMENT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    INVESTMENT_RULES_DIR.mkdir(parents=True, exist_ok=True)

    if args.keep_temp and not args.dry_run:
        raise ValueError("--keep-temp can only be used together with --dry-run")

    workbooks_for_run = real_workbooks
    temp_dir: Path | None = None

    if args.dry_run:
        workbooks_for_run, temp_dir = prepare_dry_run_workbooks(real_workbooks)
        print("Investment pipeline dry-run.")
        print("The real workbooks will not be modified.")
        for key, real_path in real_workbooks.items():
            print(f"  {key}: {real_path} -> {workbooks_for_run[key]}")
    else:
        print("Investment pipeline.")
        for key, real_path in real_workbooks.items():
            print(f"  {key}: {real_path}")

    if args.skip_pricing or args.skip_fx:
        print()
        if args.skip_pricing:
            print("Skipping instrument price fetch (--skip-pricing): using existing price data as-is.")
        if args.skip_fx:
            print("Skipping FX rate fetch (--skip-fx): using existing FX rate data as-is.")

    commands = build_pipeline_commands(
        python_executable=sys.executable,
        input_path=input_path,
        instrument_master_path=instrument_master_path,
        workbooks=workbooks_for_run,
        force=args.force,
        stale_days=args.stale_days,
        skip_pricing=args.skip_pricing,
        skip_fx=args.skip_fx,
        budgeting_workbook=Path(args.budgeting_workbook).expanduser().resolve(),
        no_local_currency=args.no_local_currency,
    )

    stage_timings: list[tuple[str, float]] = []
    try:
        for command in commands:
            elapsed = run_command(command)
            stage_timings.append((command.name, elapsed))
    except Exception:
        if args.dry_run and temp_dir is not None:
            print()
            print("Dry-run failed. Temporary files were kept for inspection:")
            print(temp_dir)
        raise

    print()
    print("Stage timing:")
    for name, elapsed in stage_timings:
        print(f"  {name:<32} {elapsed:6.1f}s")
    print(f"  {'total':<32} {sum(elapsed for _, elapsed in stage_timings):6.1f}s")

    print()
    if args.dry_run:
        print("Investment pipeline dry-run complete.")
        if args.keep_temp:
            print("Temporary files kept for inspection:")
            print(temp_dir)
        elif temp_dir is not None:
            shutil.rmtree(temp_dir)
            print("Temporary files deleted.")
        print("Real workbooks were not modified.")
    else:
        print("Investment pipeline complete.")
        for key, real_path in real_workbooks.items():
            print(f"  {key}: {real_path}")

    return 0


def main() -> None:
    parser = build_arg_parser()
    args = parser.parse_args()
    raise SystemExit(run_pipeline(args))


if __name__ == "__main__":
    main()
