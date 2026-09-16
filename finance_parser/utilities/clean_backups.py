from __future__ import annotations

def _print_cli_help_and_exit() -> None:
    """
    Early help handler.

    This runs before heavy imports and workbook setup so `--help` works even if
    project-root imports or dependencies are currently broken.
    """
    import sys

    if not any(arg in {"-h", "--help"} for arg in sys.argv[1:]):
        return

    print("""usage: python -m finance_parser.utilities.clean_backups [--root PATH] [--keep N] [--dry-run] [--help]

Sweep output/ and rules/ (or --root) for timestamped Excel backup files
(<stem>_backup_<label>_<timestamp><suffix>, as created by
replace_with_fresh_workbook()) and delete all but the --keep most recent per
base file. replace_with_fresh_workbook() already does this automatically on
every write - this command is for a one-off sweep of existing accumulation,
or for retroactively applying a different --keep value.

Does not touch the separate, non-accumulating "<stem>_backup_before_last_run"
pattern (always overwritten in place, never needs pruning).

options:
  --root PATH  Directory to scan recursively. Defaults to the whole project.
  --keep N     Backups to keep per base file. Defaults to 5.
  --dry-run    Show what would be removed without deleting anything.
  -h, --help   Show this help message and exit.

examples:
  python -m finance_parser.utilities.clean_backups --dry-run
  python -m finance_parser.utilities.clean_backups --keep 3
""")
    raise SystemExit(0)


_print_cli_help_and_exit()


from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import argparse
from collections import defaultdict

from finance_parser.utilities.fresh_workbook_writer import (
    DEFAULT_BACKUP_RETENTION,
    is_timestamped_backup,
    recover_base_path_from_backup,
)


def find_all_timestamped_backups(root: Path) -> dict[Path, list[Path]]:
    """
    Scan root recursively for timestamped backup files and group them by the
    base workbook path they belong to, newest first within each group.
    """
    groups: dict[Path, list[Path]] = defaultdict(list)

    for path in root.rglob("*"):
        if ".git" in path.parts:
            continue
        if not path.is_file() or not is_timestamped_backup(path):
            continue

        base_path = recover_base_path_from_backup(path)
        if base_path is not None:
            groups[base_path].append(path)

    for backups in groups.values():
        backups.sort(key=lambda p: p.stat().st_mtime, reverse=True)

    return dict(groups)


def sweep(root: Path, *, keep: int, dry_run: bool) -> list[Path]:
    removed: list[Path] = []

    for base_path, backups in find_all_timestamped_backups(root).items():
        for path in backups[keep:]:
            removed.append(path)
            if not dry_run:
                path.unlink()

    return removed


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Sweep timestamped Excel backup files down to --keep most recent per base file."
    )
    parser.add_argument("--root", default=str(PROJECT_ROOT), help="Directory to scan recursively. Defaults to the whole project.")
    parser.add_argument("--keep", type=int, default=DEFAULT_BACKUP_RETENTION, help="Backups to keep per base file. Defaults to 5.")
    parser.add_argument("--dry-run", action="store_true", help="Show what would be removed without deleting anything.")
    args = parser.parse_args()

    root = Path(args.root).expanduser().resolve()
    removed = sweep(root, keep=args.keep, dry_run=args.dry_run)

    action = "Would remove" if args.dry_run else "Removed"
    if not removed:
        print(f"No backups beyond the {args.keep} most recent per file found under {root}.")
        return

    print(f"{action} {len(removed)} backup file(s) (keeping {args.keep} most recent per base file):")
    for path in sorted(removed):
        try:
            print(f"  {path.relative_to(root)}")
        except ValueError:
            print(f"  {path}")


if __name__ == "__main__":
    main()
