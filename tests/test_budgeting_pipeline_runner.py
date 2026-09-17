from pathlib import Path

from finance_parser.budgeting import run_budgeting_pipeline as pipeline


def test_pipeline_cli_has_dry_run_and_keep_temp():
    parser = pipeline.build_arg_parser()
    args = parser.parse_args(["--dry-run", "--keep-temp"])

    assert args.dry_run is True
    assert args.keep_temp is True


def test_pipeline_cli_supports_verbose_debug_alias():
    parser = pipeline.build_arg_parser()

    verbose_args = parser.parse_args(["--verbose"])
    debug_args = parser.parse_args(["--debug"])

    assert verbose_args.verbose is True
    assert debug_args.verbose is True


def test_pipeline_cli_supports_no_profile():
    parser = pipeline.build_arg_parser()
    args = parser.parse_args(["--no-profile"])

    assert args.no_profile is True


def test_build_pipeline_commands_runs_steps_in_expected_order():
    commands = pipeline.build_pipeline_commands(
        python_executable="python",
        input_path=Path("input/budgeting"),
        workbook_path=Path("output/budgeting/ParsedTransactions.xlsx"),
        rules_path=Path("rules/budgeting/TransactionRules.xlsx"),
    )

    assert [command.name for command in commands] == [
        "transaction parser",
        "transaction normaliser",
        "transaction categoriser",
    ]

    assert commands[0].argv[:3] == ["python", "-m", "finance_parser.budgeting.transaction_parser"]
    assert commands[1].argv[:3] == ["python", "-m", "finance_parser.budgeting.transaction_normaliser"]
    assert commands[2].argv[:3] == ["python", "-m", "finance_parser.budgeting.transaction_categoriser"]


def test_build_pipeline_commands_uses_output_workbook_for_all_steps():
    workbook = Path("sandbox/ParsedTransactions.xlsx")
    commands = pipeline.build_pipeline_commands(
        python_executable="python",
        input_path=Path("input/budgeting"),
        workbook_path=workbook,
        rules_path=Path("rules/budgeting/TransactionRules.xlsx"),
    )

    assert "--output" in commands[0].argv
    assert str(workbook) in commands[0].argv

    assert "--workbook" in commands[1].argv
    assert str(workbook) in commands[1].argv

    assert "--workbook" in commands[2].argv
    assert str(workbook) in commands[2].argv


def test_build_pipeline_commands_does_not_pass_step_dry_run():
    commands = pipeline.build_pipeline_commands(
        python_executable="python",
        input_path=Path("input/budgeting"),
        workbook_path=Path("sandbox/ParsedTransactions.xlsx"),
        rules_path=Path("rules/budgeting/TransactionRules.xlsx"),
    )

    # Pipeline --dry-run is implemented by running the real non-dry steps
    # against a copied temporary workbook, not by passing isolated --dry-run
    # to each step.
    assert all("--dry-run" not in command.argv for command in commands)


def test_dividend_stages_included_when_history_path_given():
    commands = pipeline.build_pipeline_commands(
        python_executable="python",
        input_path=Path("input/budgeting"),
        workbook_path=Path("output/budgeting/ParsedTransactions.xlsx"),
        rules_path=Path("rules/budgeting/TransactionRules.xlsx"),
        dividend_history_path=Path("output/investments/DividendHistory.xlsx"),
    )

    assert [command.name for command in commands] == [
        "transaction parser",
        "investment dividend parser",
        "transaction normaliser",
        "dividend income enrichment",
        "transaction categoriser",
    ]
    # Dividend enrichment runs BEFORE the categoriser so its cross-referenced
    # values win by default (categoriser rules default to BLANK_ONLY).
    names = [c.name for c in commands]
    assert names.index("dividend income enrichment") < names.index("transaction categoriser")


def test_dividend_stages_skipped_when_history_path_omitted():
    commands = pipeline.build_pipeline_commands(
        python_executable="python",
        input_path=Path("input/budgeting"),
        workbook_path=Path("output/budgeting/ParsedTransactions.xlsx"),
        rules_path=Path("rules/budgeting/TransactionRules.xlsx"),
    )

    assert [command.name for command in commands] == [
        "transaction parser",
        "transaction normaliser",
        "transaction categoriser",
    ]


def test_dividend_stages_skipped_when_explicitly_requested():
    commands = pipeline.build_pipeline_commands(
        python_executable="python",
        input_path=Path("input/budgeting"),
        workbook_path=Path("output/budgeting/ParsedTransactions.xlsx"),
        rules_path=Path("rules/budgeting/TransactionRules.xlsx"),
        dividend_history_path=Path("output/investments/DividendHistory.xlsx"),
        skip_dividends=True,
    )

    assert [command.name for command in commands] == [
        "transaction parser",
        "transaction normaliser",
        "transaction categoriser",
    ]


def test_build_pipeline_commands_can_disable_parser_profile_and_enable_verbose():
    commands = pipeline.build_pipeline_commands(
        python_executable="python",
        input_path=Path("input/budgeting"),
        workbook_path=Path("sandbox/ParsedTransactions.xlsx"),
        rules_path=Path("rules/budgeting/TransactionRules.xlsx"),
        no_profile=True,
        verbose=True,
    )

    assert "--no-profile" in commands[0].argv
    assert "--verbose" in commands[0].argv
    assert "--no-profile" not in commands[1].argv
    assert "--verbose" not in commands[1].argv
    assert "--no-profile" not in commands[2].argv
    assert "--verbose" not in commands[2].argv
