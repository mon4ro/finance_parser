import inspect

import pytest

from finance_parser.budgeting import transaction_parser


def test_append_to_output_has_profile_argument():
    signature = inspect.signature(transaction_parser.append_to_output)
    assert "profile" in signature.parameters


def test_append_to_output_has_dry_run_and_verbose_arguments():
    signature = inspect.signature(transaction_parser.append_to_output)
    assert "dry_run" in signature.parameters
    assert "verbose" in signature.parameters


def test_parse_import_files_has_profile_argument():
    signature = inspect.signature(transaction_parser.parse_import_files)
    assert "profile" in signature.parameters


def test_parse_import_files_profile_no_nameerror_on_empty_folder(tmp_path):
    with pytest.raises((FileNotFoundError, ValueError, RuntimeError)) as excinfo:
        transaction_parser.parse_import_files(tmp_path, "2026-06-04 12:00:00", profile=True)

    assert excinfo.type is not NameError
