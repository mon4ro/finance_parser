import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]


TOOLS_SCRIPTS = [
    "tools/seed_account_balance.py",
    "tools/find_duplicate_transactions.py",
    "tools/remove_import.py",
]


def test_tools_scripts_run_as_real_subprocess_with_help():
    """
    Real bug: tools/seed_account_balance.py and tools/find_duplicate_transactions.py
    were moved from the repo root into tools/ (one level deeper) without adding
    the sys.path shim that every other nested wrapper script needs (see
    finance_parser/utilities/cleanup_local_artifacts.py) - "python -m pytest"
    still passed throughout because pytest.ini's pythonpath=. masks the issue
    entirely, but a real `python tools/seed_account_balance.py` subprocess
    invocation (as the README documents) failed with ModuleNotFoundError.
    Running these as an actual subprocess - not just importing them in-process -
    is the only way this class of bug gets caught.
    """
    for rel in TOOLS_SCRIPTS:
        result = subprocess.run(
            [sys.executable, str(PROJECT_ROOT / rel), "--help"],
            cwd=PROJECT_ROOT,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )

        assert result.returncode == 0, f"{rel} failed --help\nSTDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}"
        assert "usage:" in result.stdout.lower()
