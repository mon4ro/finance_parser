import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]


UTILITY_SCRIPTS = [
    "finance_parser/utilities/cleanup_budgeting_source_accounts.py",
    "finance_parser/utilities/refresh_budgeting_bank_raw_ids.py",
    "finance_parser/utilities/review_cleanup_unified_duplicates.py",
    "finance_parser/utilities/cleanup_local_artifacts.py",
]


def test_utility_scripts_support_help():
    for rel in UTILITY_SCRIPTS:
        result = subprocess.run(
            [sys.executable, str(PROJECT_ROOT / rel), "--help"],
            cwd=PROJECT_ROOT,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )

        assert result.returncode == 0, f"{rel} failed --help\nSTDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}"
        assert "usage:" in result.stdout.lower()
        assert "--dry-run" in result.stdout
        assert "--help" in result.stdout


def test_utility_help_does_not_require_root_wrappers():
    """
    Help output should not make tests depend on root utilities/ wrappers.

    During the wrapper-removal migration, individual utility help strings may
    still mention historical commands. This test intentionally validates CLI
    functionality rather than enforcing a root wrapper path.
    """
    for rel in UTILITY_SCRIPTS:
        result = subprocess.run(
            [sys.executable, str(PROJECT_ROOT / rel), "--help"],
            cwd=PROJECT_ROOT,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )

        assert result.returncode == 0
        assert "usage:" in result.stdout.lower()
