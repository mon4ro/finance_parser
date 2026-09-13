from __future__ import annotations

import argparse
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]

TARGET_FILES = [
    "finance_parser/common.py",
    "finance_parser/budgeting/transaction_parser.py",
    "finance_parser/budgeting/transaction_normaliser.py",
    "finance_parser/budgeting/parsers/op.py",
    "finance_parser/budgeting/parsers/norwegian.py",
    "finance_parser/budgeting/parsers/nordea.py",
    "finance_parser/budgeting/parsers/spankki.py",
    "finance_parser/investments/investment_parser.py",
    "finance_parser/investments/investment_common.py",
    "finance_parser/investments/parsers/nordnet.py",
    "finance_parser/investments/parsers/seligson.py",
    "finance_parser/investments/parsers/evli.py",
    "finance_parser/investments/parsers/op_investment.py",
    "finance_parser/utilities/cleanup_budgeting_source_accounts.py",
    "finance_parser/utilities/refresh_budgeting_bank_raw_ids.py",
    "finance_parser/utilities/review_cleanup_unified_duplicates.py",
]

REPLACEMENTS = {
    "finance_parser/common.py": [
        ("from settings import get_settings", "from finance_parser.settings import get_settings"),
    ],
    "finance_parser/budgeting/transaction_parser.py": [
        ("from common import (", "from finance_parser.common import ("),
        ("from parsers import op, norwegian, nordea, spankki", "from finance_parser.budgeting.parsers import op, norwegian, nordea, spankki"),
    ],
    "finance_parser/budgeting/transaction_normaliser.py": [
        ("from common import append_change_log_entry", "from finance_parser.common import append_change_log_entry"),
    ],
    "finance_parser/budgeting/parsers/op.py": [("from common import (", "from finance_parser.common import (")],
    "finance_parser/budgeting/parsers/norwegian.py": [("from common import (", "from finance_parser.common import (")],
    "finance_parser/budgeting/parsers/nordea.py": [("from common import (", "from finance_parser.common import (")],
    "finance_parser/budgeting/parsers/spankki.py": [("from common import (", "from finance_parser.common import (")],
    "finance_parser/investments/investment_parser.py": [
        ("from common import imported_at_now, make_import_run_id, normalise_header", "from finance_parser.common import imported_at_now, make_import_run_id, normalise_header"),
        ("from investments.investment_common import (", "from finance_parser.investments.investment_common import ("),
        ("from investments.parsers import nordnet, seligson, evli, op_investment", "from finance_parser.investments.parsers import nordnet, seligson, evli, op_investment"),
    ],
    "finance_parser/investments/investment_common.py": [("from common import (", "from finance_parser.common import (")],
    "finance_parser/investments/parsers/nordnet.py": [
        ("from common import (", "from finance_parser.common import ("),
        ("from investments.investment_common import INVESTMENT_RAW_COLUMNS", "from finance_parser.investments.investment_common import INVESTMENT_RAW_COLUMNS"),
    ],
    "finance_parser/investments/parsers/seligson.py": [
        ("from settings import get_settings", "from finance_parser.settings import get_settings"),
        ("from common import (", "from finance_parser.common import ("),
        ("from investments.investment_common import INVESTMENT_RAW_COLUMNS", "from finance_parser.investments.investment_common import INVESTMENT_RAW_COLUMNS"),
    ],
    "finance_parser/investments/parsers/evli.py": [
        ("from settings import get_settings", "from finance_parser.settings import get_settings"),
        ("from common import (", "from finance_parser.common import ("),
        ("from investments.investment_common import INVESTMENT_RAW_COLUMNS", "from finance_parser.investments.investment_common import INVESTMENT_RAW_COLUMNS"),
    ],
    "finance_parser/investments/parsers/op_investment.py": [
        ("from common import (", "from finance_parser.common import ("),
        ("from investments.investment_common import INVESTMENT_RAW_COLUMNS", "from finance_parser.investments.investment_common import INVESTMENT_RAW_COLUMNS"),
    ],
    "finance_parser/utilities/cleanup_budgeting_source_accounts.py": [("from common import (", "from finance_parser.common import (")],
    "finance_parser/utilities/review_cleanup_unified_duplicates.py": [("from common import (", "from finance_parser.common import (")],
    "finance_parser/utilities/refresh_budgeting_bank_raw_ids.py": [
        ("from common import clean_dataframe_for_excel, clean_for_excel, normalize_source_metadata, append_change_log_entry", "from finance_parser.common import clean_dataframe_for_excel, clean_for_excel, normalize_source_metadata, append_change_log_entry"),
        ("from parsers import op, norwegian, nordea, spankki", "from finance_parser.budgeting.parsers import op, norwegian, nordea, spankki"),
    ],
}

FORBIDDEN_PREFIXES = [
    "from common import",
    "import common",
    "from settings import",
    "import settings",
    "from parsers import",
    "import parsers",
    "from investments.investment_common import",
    "from investments.parsers import",
]

def apply_replacements(path: Path, dry_run: bool) -> tuple[bool, list[str]]:
    rel = path.relative_to(PROJECT_ROOT).as_posix()
    text = path.read_text(encoding="utf-8")
    original = text
    notes = []
    for old, new in REPLACEMENTS.get(rel, []):
        if old in text:
            text = text.replace(old, new)
            notes.append(f"{old} -> {new}")
    if text != original and not dry_run:
        path.write_text(text, encoding="utf-8")
    return text != original, notes

def scan_forbidden_imports() -> list[str]:
    findings = []
    for path in (PROJECT_ROOT / "finance_parser").rglob("*.py"):
        if "__pycache__" in path.parts:
            continue
        for lineno, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            stripped = line.strip()
            if any(stripped.startswith(prefix) for prefix in FORBIDDEN_PREFIXES):
                findings.append(f"{path.relative_to(PROJECT_ROOT)}:{lineno}: {stripped}")
    return findings

def main() -> None:
    parser = argparse.ArgumentParser(description="Migrate finance_parser internal imports to package imports.")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()

    if args.check:
        findings = scan_forbidden_imports()
        if findings:
            print("Old shim imports still found inside finance_parser/:")
            for finding in findings:
                print(f"  {finding}")
            raise SystemExit(1)
        print("No old shim imports found inside finance_parser/.")
        return

    any_changed = False
    for rel in TARGET_FILES:
        path = PROJECT_ROOT / rel
        if not path.exists():
            print(f"SKIP missing: {rel}")
            continue
        changed, notes = apply_replacements(path, args.dry_run)
        any_changed = any_changed or changed
        action = "Would update" if args.dry_run else "Updated"
        if changed:
            print(f"{action}: {rel}")
            for note in notes:
                print(f"  {note}")
        else:
            print(f"No change: {rel}")

    findings = scan_forbidden_imports()
    if findings:
        print("\nRemaining old shim imports inside finance_parser/:")
        for finding in findings:
            print(f"  {finding}")
    elif not args.dry_run:
        print("\nNo old shim imports found inside finance_parser/.")

    if not any_changed:
        print("\nNo files changed.")

if __name__ == "__main__":
    main()
