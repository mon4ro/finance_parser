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

    print("""usage: python utilities/cleanup_local_artifacts.py [--dry-run] [--help]

Remove local generated artifacts such as .DS_Store, __pycache__, *.pyc, and .pytest_cache.

options:
  --dry-run   Show what would be changed without modifying files.
  -h, --help  Show this help message and exit.

examples:
  python utilities/cleanup_local_artifacts.py --dry-run\n  python utilities/cleanup_local_artifacts.py
""")
    raise SystemExit(0)


_print_cli_help_and_exit()


from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import argparse
import shutil
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def iter_cleanup_targets(project_root: Path):
    # macOS Finder metadata.
    yield from project_root.rglob(".DS_Store")

    # Python bytecode caches.
    yield from project_root.rglob("__pycache__")
    yield from project_root.rglob("*.pyc")

    # Pytest local cache.
    pytest_cache = project_root / ".pytest_cache"
    if pytest_cache.exists():
        yield pytest_cache


def should_skip(path: Path) -> bool:
    # Never delete Git internals.
    return ".git" in path.parts


def cleanup(dry_run: bool = False) -> list[Path]:
    removed: list[Path] = []
    seen: set[Path] = set()

    for path in iter_cleanup_targets(PROJECT_ROOT):
        path = path.resolve()

        if path in seen or should_skip(path):
            continue

        seen.add(path)

        if not path.exists():
            continue

        removed.append(path)

        if dry_run:
            continue

        if path.is_dir():
            shutil.rmtree(path)
        else:
            path.unlink()

    return removed


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Remove local/generated artifacts such as .DS_Store, __pycache__, *.pyc, and .pytest_cache."
    )
    parser.add_argument("--dry-run", action="store_true", help="Show what would be removed without deleting anything.")
    args = parser.parse_args()

    targets = cleanup(dry_run=args.dry_run)

    action = "Would remove" if args.dry_run else "Removed"
    if not targets:
        print("No local artifacts found.")
        return

    print(f"{action} {len(targets)} local artifact(s):")
    for path in targets:
        print(f"  {path.relative_to(PROJECT_ROOT)}")


if __name__ == "__main__":
    main()
