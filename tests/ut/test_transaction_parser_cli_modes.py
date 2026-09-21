import inspect

from finance_parser.budgeting import transaction_parser as parser_module


def test_transaction_parser_cli_defaults_profile_on_and_dry_run_off():
    parser = parser_module.build_arg_parser()
    args = parser.parse_args([])

    assert args.profile is True
    assert args.dry_run is False
    assert args.verbose is False


def test_transaction_parser_cli_can_disable_profile():
    parser = parser_module.build_arg_parser()
    args = parser.parse_args(["--no-profile"])

    assert args.profile is False


def test_transaction_parser_cli_supports_dry_run():
    parser = parser_module.build_arg_parser()
    args = parser.parse_args(["--dry-run"])

    assert args.dry_run is True


def test_transaction_parser_cli_supports_verbose_and_debug_aliases():
    parser = parser_module.build_arg_parser()

    verbose_args = parser.parse_args(["--verbose"])
    debug_args = parser.parse_args(["--debug"])

    assert verbose_args.verbose is True
    assert debug_args.verbose is True


def test_append_to_output_accepts_dry_run_and_verbose_flags():
    signature = inspect.signature(parser_module.append_to_output)

    assert "dry_run" in signature.parameters
    assert "verbose" in signature.parameters
    assert signature.parameters["dry_run"].default is False
    assert signature.parameters["verbose"].default is False


def test_dry_run_branch_precedes_workbook_writes():
    source = inspect.getsource(parser_module.append_to_output)

    dry_run_index = source.index("if dry_run:")
    canonical_write_index = source.index("write_clean_output_workbook(")
    bank_raw_write_index = source.index("write_extra_sheets_to_existing_workbook(")
    changelog_index = source.index("append_change_log_entry(")

    assert dry_run_index < canonical_write_index
    assert dry_run_index < bank_raw_write_index
    assert dry_run_index < changelog_index
