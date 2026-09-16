import os
import time
from pathlib import Path

from finance_parser.utilities.clean_backups import find_all_timestamped_backups, sweep


def _touch(path: Path, *, age_seconds: int) -> Path:
    path.write_text("stub")
    now = time.time()
    os.utime(path, (now - age_seconds, now - age_seconds))
    return path


def test_find_all_timestamped_backups_groups_by_base_file(tmp_path):
    a_new = _touch(tmp_path / "A_backup_x_20260101_120000.xlsx", age_seconds=0)
    a_old = _touch(tmp_path / "A_backup_y_20260101_110000.xlsx", age_seconds=100)
    b_only = _touch(tmp_path / "B_backup_x_20260101_120000.xlsx", age_seconds=0)
    # Non-timestamped pattern - not a pruning candidate, must be ignored.
    _touch(tmp_path / "A_backup_before_last_run.xlsx", age_seconds=0)
    # Unrelated file.
    _touch(tmp_path / "notes.txt", age_seconds=0)

    groups = find_all_timestamped_backups(tmp_path)

    assert groups[tmp_path / "A.xlsx"] == [a_new, a_old]
    assert groups[tmp_path / "B.xlsx"] == [b_only]
    assert len(groups) == 2


def test_sweep_removes_beyond_keep_per_base_file(tmp_path):
    files = [
        _touch(tmp_path / f"A_backup_run{i}_2026010{i}_120000.xlsx", age_seconds=(10 - i))
        for i in range(1, 6)
    ]

    removed = sweep(tmp_path, keep=3, dry_run=False)

    assert set(removed) == set(files[:2])
    for path in files[:2]:
        assert not path.exists()
    for path in files[2:]:
        assert path.exists()


def test_sweep_dry_run_reports_without_deleting(tmp_path):
    files = [
        _touch(tmp_path / f"A_backup_run{i}_2026010{i}_120000.xlsx", age_seconds=(10 - i))
        for i in range(1, 4)
    ]

    removed = sweep(tmp_path, keep=1, dry_run=True)

    assert len(removed) == 2
    for path in files:
        assert path.exists()


def test_sweep_ignores_non_timestamped_backup_pattern(tmp_path):
    _touch(tmp_path / "A_backup_before_last_run.xlsx", age_seconds=0)

    removed = sweep(tmp_path, keep=0, dry_run=False)

    assert removed == []
    assert (tmp_path / "A_backup_before_last_run.xlsx").exists()
