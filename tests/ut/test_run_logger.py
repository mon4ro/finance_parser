import subprocess
import sys
from pathlib import Path

from finance_parser.utilities.run_logger import (
    DEFAULT_LOG_RETENTION,
    TeeLog,
    _LineDeduper,
    find_logs,
    log_file_path,
    prune_old_logs,
    start_run_log,
)


def test_log_file_path_uses_label_and_timestamp():
    from datetime import datetime

    path = log_file_path(Path("/tmp/logs"), "mypipeline", now=datetime(2026, 3, 5, 14, 30, 0))
    assert path == Path("/tmp/logs/mypipeline_20260305_143000.log")


def test_find_logs_only_matches_own_label_newest_first(tmp_path):
    for name in [
        "budgeting_pipeline_20260101_000000.log",
        "budgeting_pipeline_20260301_000000.log",
        "budgeting_pipeline_20260201_000000.log",
        "other_label_20260401_000000.log",
        "not_a_log.txt",
    ]:
        (tmp_path / name).write_text("x")

    found = find_logs(tmp_path, "budgeting_pipeline")
    assert [p.name for p in found] == [
        "budgeting_pipeline_20260301_000000.log",
        "budgeting_pipeline_20260201_000000.log",
        "budgeting_pipeline_20260101_000000.log",
    ]


def test_find_logs_empty_dir_returns_empty_list(tmp_path):
    assert find_logs(tmp_path / "does_not_exist", "budgeting_pipeline") == []


def test_prune_old_logs_keeps_only_the_newest_n(tmp_path):
    for stamp in ["20260101_000000", "20260102_000000", "20260103_000000", "20260104_000000"]:
        (tmp_path / f"budgeting_pipeline_{stamp}.log").write_text("x")

    removed = prune_old_logs(tmp_path, "budgeting_pipeline", keep=2)

    assert len(removed) == 2
    remaining = {p.name for p in tmp_path.glob("*.log")}
    assert remaining == {
        "budgeting_pipeline_20260104_000000.log",
        "budgeting_pipeline_20260103_000000.log",
    }


def test_prune_old_logs_default_retention_matches_documented_value():
    assert DEFAULT_LOG_RETENTION == 5


def test_line_deduper_passes_through_distinct_lines_unchanged():
    d = _LineDeduper()
    out = d.feed(b"first line\nsecond line\n")
    assert out == "first line\nsecond line\n"


def test_line_deduper_collapses_consecutive_repeats_with_a_count():
    """
    Real motivating case: a parser probing several candidate readers per
    file prints the identical "Successfully read X using: Y" line multiple
    times in a row for one file - zero extra information the 2nd/3rd/Nth
    time. Only consecutive repeats collapse; the same line reappearing
    later (not back-to-back) is real signal and must not be swallowed.
    """
    d = _LineDeduper()
    out = d.feed(b"Successfully read X\nSuccessfully read X\nSuccessfully read X\nDone\n")
    assert out == "Successfully read X\n  ... (previous line repeated 2 more times)\nDone\n"


def test_line_deduper_non_consecutive_repeat_is_not_collapsed():
    d = _LineDeduper()
    out = d.feed(b"A\nB\nA\n")
    assert out == "A\nB\nA\n"


def test_line_deduper_flush_emits_trailing_partial_line_and_pending_repeat_summary():
    d = _LineDeduper()
    d.feed(b"same\nsame\n")
    tail = d.flush()
    assert tail == "  ... (previous line repeated 1 more time)\n"

    d2 = _LineDeduper()
    d2.feed(b"complete\nno newline yet")
    tail2 = d2.flush()
    assert tail2 == "no newline yet"


def test_tee_log_captures_both_own_prints_and_a_real_subprocess(tmp_path, capfd):
    """
    The whole reason this uses os.dup2 on the real file descriptor instead
    of reassigning sys.stdout: a subprocess.run() call with no stdout=
    override inherits the parent's fd directly, bypassing sys.stdout
    entirely. A plain "redirect sys.stdout to a file" approach would miss
    every pipeline stage's own output (each stage is its own subprocess -
    see run_command() in run_budgeting_pipeline.py). This proves both paths
    land in the log.

    capfd.disabled() is required here: TeeLog does its own real fd-level
    dup2, which pytest's own default fd-level capture otherwise fights over
    (confirmed - this exact scenario passes as a standalone script outside
    pytest; only interferes when pytest's own capture is also holding fd 1).
    """
    log_dir = tmp_path / "logs"
    with capfd.disabled():
        with TeeLog(log_dir / "test_run.log", label="test", mode="TEST", argv=["prog", "--flag"]) as log:
            print("from this process")
            subprocess.run([sys.executable, "-c", "print('from a real subprocess')"], check=True)

    content = log.log_path.read_text()
    assert "from this process" in content
    assert "from a real subprocess" in content
    assert "TEST" in content
    assert "COMPLETED" in content


def test_tee_log_records_failure_status_when_the_block_raises(tmp_path):
    log_dir = tmp_path / "logs"
    log_path = log_dir / "failing_run.log"
    try:
        with TeeLog(log_path, label="test", mode="TEST", argv=["prog"]):
            raise RuntimeError("boom")
    except RuntimeError:
        pass

    content = log_path.read_text()
    assert "FAILED" in content


def test_start_run_log_builds_timestamped_path_under_the_given_dir(tmp_path):
    with start_run_log(tmp_path, "mylabel", mode="TEST", argv=["prog"]) as log:
        print("hello")

    assert log.log_path.parent == tmp_path
    assert log.log_path.name.startswith("mylabel_")
    assert "hello" in log.log_path.read_text()
