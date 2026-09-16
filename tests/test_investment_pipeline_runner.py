from pathlib import Path

import pytest

from finance_parser.investments import run_investment_pipeline as pipeline


def test_run_command_returns_elapsed_seconds_and_prints_timing(capsys, monkeypatch):
    monkeypatch.setattr(pipeline.subprocess, "run", lambda argv, check: None)

    elapsed = pipeline.run_command(pipeline.PipelineCommand("dummy stage", ["true"]))

    assert isinstance(elapsed, float)
    assert elapsed >= 0
    assert "dummy stage done in" in capsys.readouterr().out


def test_run_command_propagates_failure_without_swallowing_it(monkeypatch):
    import subprocess as subprocess_module

    def raise_error(argv, check):
        raise subprocess_module.CalledProcessError(1, argv)

    monkeypatch.setattr(pipeline.subprocess, "run", raise_error)

    with pytest.raises(subprocess_module.CalledProcessError):
        pipeline.run_command(pipeline.PipelineCommand("dummy stage", ["false"]))


def test_pipeline_cli_has_dry_run_keep_temp_and_force():
    parser = pipeline.build_arg_parser()
    args = parser.parse_args(["--dry-run", "--keep-temp", "--force"])

    assert args.dry_run is True
    assert args.keep_temp is True
    assert args.force is True


def test_build_pipeline_commands_runs_stages_in_dependency_order():
    workbooks = {
        "investments": Path("output/investments/ParsedInvestments.xlsx"),
        "prices": Path("output/investments/InstrumentPrices.xlsx"),
        "fx_rates": Path("output/investments/FXRates.xlsx"),
        "positions": Path("output/investments/PortfolioPositions.xlsx"),
        "monthly_value": Path("output/investments/MonthlyPositionValue.xlsx"),
    }
    commands = pipeline.build_pipeline_commands(
        python_executable="python",
        input_path=Path("input/investments"),
        instrument_master_path=Path("rules/investments/InstrumentMaster.xlsx"),
        workbooks=workbooks,
        force=False,
        stale_days=7,
    )

    assert [c.name for c in commands] == [
        "investment parser",
        "instrument price fetch",
        "FX rate fetch",
        "portfolio positions",
        "instrument coverage check",
        "monthly position value rollup",
    ]

    assert commands[0].argv[:3] == ["python", "-m", "finance_parser.investments.investment_parser"]
    assert commands[1].argv[:3] == ["python", "-m", "finance_parser.investments.fetch_instrument_prices"]
    assert commands[2].argv[:3] == ["python", "-m", "finance_parser.investments.fetch_fx_rates"]
    assert commands[3].argv[:3] == ["python", "-m", "finance_parser.investments.build_portfolio_positions"]
    assert commands[4].argv[:3] == ["python", "-m", "finance_parser.investments.check_instrument_coverage"]
    assert commands[5].argv[:3] == ["python", "-m", "finance_parser.investments.build_monthly_position_value"]


def test_coverage_check_is_marked_as_the_gate_stage():
    workbooks = {key: Path(f"{key}.xlsx") for key in pipeline.WORKBOOK_KEYS}
    commands = pipeline.build_pipeline_commands(
        python_executable="python",
        input_path=Path("input/investments"),
        instrument_master_path=Path("rules/investments/InstrumentMaster.xlsx"),
        workbooks=workbooks,
        force=False,
        stale_days=7,
    )

    gate_commands = [c for c in commands if c.is_gate]
    assert [c.name for c in gate_commands] == ["instrument coverage check"]


def test_force_flag_passes_through_to_coverage_check_only():
    workbooks = {key: Path(f"{key}.xlsx") for key in pipeline.WORKBOOK_KEYS}
    commands = pipeline.build_pipeline_commands(
        python_executable="python",
        input_path=Path("input/investments"),
        instrument_master_path=Path("rules/investments/InstrumentMaster.xlsx"),
        workbooks=workbooks,
        force=True,
        stale_days=7,
    )

    coverage_cmd = next(c for c in commands if c.name == "instrument coverage check")
    assert "--force" in coverage_cmd.argv

    other_cmds = [c for c in commands if c.name != "instrument coverage check"]
    assert all("--force" not in c.argv for c in other_cmds)


def test_stale_days_passes_through_to_coverage_check():
    workbooks = {key: Path(f"{key}.xlsx") for key in pipeline.WORKBOOK_KEYS}
    commands = pipeline.build_pipeline_commands(
        python_executable="python",
        input_path=Path("input/investments"),
        instrument_master_path=Path("rules/investments/InstrumentMaster.xlsx"),
        workbooks=workbooks,
        force=False,
        stale_days=14,
    )

    coverage_cmd = next(c for c in commands if c.name == "instrument coverage check")
    assert "--stale-days" in coverage_cmd.argv
    assert "14" in coverage_cmd.argv


def test_build_pipeline_commands_does_not_pass_step_dry_run():
    workbooks = {key: Path(f"{key}.xlsx") for key in pipeline.WORKBOOK_KEYS}
    commands = pipeline.build_pipeline_commands(
        python_executable="python",
        input_path=Path("input/investments"),
        instrument_master_path=Path("rules/investments/InstrumentMaster.xlsx"),
        workbooks=workbooks,
        force=False,
        stale_days=7,
    )

    # Pipeline --dry-run is implemented by running the real steps against
    # temp-copied workbooks, not by passing isolated --dry-run to each stage.
    assert all("--dry-run" not in c.argv for c in commands)


def test_keep_temp_requires_dry_run():
    parser = pipeline.build_arg_parser()
    args = parser.parse_args(["--keep-temp"])

    import pytest

    with pytest.raises(ValueError):
        pipeline.run_pipeline(args)


def test_prepare_dry_run_workbooks_skips_copy_for_nonexistent_files(tmp_path, monkeypatch):
    monkeypatch.setattr(pipeline, "INVESTMENT_OUTPUT_DIR", tmp_path)

    real = {
        "investments": tmp_path / "does_not_exist.xlsx",
        "prices": tmp_path / "also_missing.xlsx",
    }
    temp_workbooks, temp_dir = pipeline.prepare_dry_run_workbooks(real, timestamp="20260101_000000")

    assert temp_dir.exists()
    assert not temp_workbooks["investments"].exists()
    assert not temp_workbooks["prices"].exists()
    assert temp_workbooks["investments"].parent == temp_dir
