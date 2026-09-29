from __future__ import annotations

import os
import re
import sys
import threading
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path


DEFAULT_LOG_RETENTION = 5

_LOG_FILENAME_RE = re.compile(r"^(?P<label>.+)_(?P<stamp>\d{8}_\d{6})\.log$")


def log_file_path(log_dir: Path, label: str, *, now: datetime | None = None) -> Path:
    stamp = (now or datetime.now()).strftime("%Y%m%d_%H%M%S")
    return log_dir / f"{label}_{stamp}.log"


def find_logs(log_dir: Path, label: str) -> list[Path]:
    """All log files for this label, newest first (by the timestamp in the
    filename, not mtime - a log file's own name is authoritative for when
    the run it records started, and sorts correctly with no filesystem
    dependency)."""
    if not log_dir.exists():
        return []
    matches = []
    for path in log_dir.glob(f"{label}_*.log"):
        m = _LOG_FILENAME_RE.match(path.name)
        if m and m.group("label") == label:
            matches.append((m.group("stamp"), path))
    matches.sort(key=lambda pair: pair[0], reverse=True)
    return [path for _, path in matches]


def prune_old_logs(log_dir: Path, label: str, *, keep: int = DEFAULT_LOG_RETENTION) -> list[Path]:
    """Delete all but the `keep` most recent log files for this label. Returns
    the paths removed."""
    logs = find_logs(log_dir, label)
    to_remove = logs[keep:]
    for path in to_remove:
        path.unlink(missing_ok=True)
    return to_remove


@dataclass
class _LineDeduper:
    """
    Collapses immediately-repeated identical lines before they reach the log
    file - real pipeline output includes runs of literally-identical lines
    (a parser probing several candidate readers per file, each printing the
    same "Successfully read X using: Y" line) that carry zero extra
    information the second/third/Nth time. The live terminal/pipe stream
    passed through by the caller is never touched by this - only what gets
    written to the log file is deduplicated, so nothing about the real
    command's own output behaviour changes.
    """

    last_line: str | None = None
    repeat_count: int = 0
    _buffer: bytes = field(default=b"", repr=False)

    def feed(self, chunk: bytes) -> str:
        self._buffer += chunk
        out_lines: list[str] = []
        while b"\n" in self._buffer:
            raw_line, self._buffer = self._buffer.split(b"\n", 1)
            out_lines.append(raw_line.decode("utf-8", errors="replace"))
        return "".join(self._emit(line) for line in out_lines)

    def flush(self) -> str:
        text = self._buffer.decode("utf-8", errors="replace") if self._buffer else ""
        self._buffer = b""
        tail = self._emit_repeat_summary()
        return tail + text

    def _emit(self, line: str) -> str:
        if line == self.last_line:
            self.repeat_count += 1
            return ""
        summary = self._emit_repeat_summary()
        self.last_line = line
        self.repeat_count = 0
        return summary + line + "\n"

    def _emit_repeat_summary(self) -> str:
        if self.repeat_count <= 0:
            return ""
        n = self.repeat_count
        self.repeat_count = 0
        return f"  ... (previous line repeated {n} more time{'s' if n != 1 else ''})\n"


def _run_header(label: str, *, mode: str, argv: list[str]) -> str:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    bar = "=" * 70
    lines = [
        bar,
        f"{label}  -  {mode}",
        f"Started:  {now}",
        f"Command:  {' '.join(argv)}",
        bar,
        "",
    ]
    return "\n".join(lines)


def _run_footer(*, ok: bool, started: datetime) -> str:
    elapsed = (datetime.now() - started).total_seconds()
    bar = "=" * 70
    status = "COMPLETED" if ok else "FAILED"
    lines = [
        "",
        bar,
        f"{status}  -  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  ({elapsed:.1f}s elapsed)",
        bar,
    ]
    return "\n".join(lines)


class TeeLog:
    """
    Mirrors everything written to real stdout/stderr - including output from
    child processes spawned via subprocess.run(), which write straight to
    the inherited OS file descriptor and never pass through Python's
    sys.stdout object - into a log file, without changing what appears live
    on the terminal at all.

    Works at the file-descriptor level (os.dup2), not by reassigning
    sys.stdout: a subprocess with no stdout=/stderr= override inherits
    whatever fd 1/2 currently point to, so redirecting the underlying fd
    (rather than the Python-level object) is the only way to capture that
    output too, not just this process's own print() calls.
    """

    def __init__(self, log_path: Path, *, label: str, mode: str, argv: list[str]):
        self.log_path = log_path
        self._label = label
        self._mode = mode
        self._argv = argv
        self._started = datetime.now()
        self._log_file = None
        self._saved_stdout_fd: int | None = None
        self._saved_stderr_fd: int | None = None
        self._pump_thread: threading.Thread | None = None

    def __enter__(self) -> "TeeLog":
        self.log_path.parent.mkdir(parents=True, exist_ok=True)
        self._log_file = open(self.log_path, "w", encoding="utf-8")
        self._log_file.write(_run_header(self._label, mode=self._mode, argv=self._argv))
        self._log_file.flush()

        sys.stdout.flush()
        sys.stderr.flush()
        # Real bug found and fixed: Python's stdout is line-buffered when
        # connected to a real terminal but switches to full block-buffering
        # once the underlying fd points somewhere else (as dup2 below makes
        # it do) - this process's own print() calls then sat in memory and
        # only landed in the pipe in one big batch near exit, while each
        # subprocess stage wrote its own output immediately. Net effect: the
        # "==> Running X" section headers appeared wildly out of chronological
        # order, both in the log and live on the terminal. Forcing line
        # buffering here (independent of what fd sys.stdout happens to wrap)
        # keeps this process's own output interleaved correctly with every
        # subprocess's.
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(line_buffering=True)
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(line_buffering=True)
        stdout_fd = sys.stdout.fileno()
        stderr_fd = sys.stderr.fileno()
        self._saved_stdout_fd = os.dup(stdout_fd)
        self._saved_stderr_fd = os.dup(stderr_fd)

        read_fd, write_fd = os.pipe()
        os.dup2(write_fd, stdout_fd)
        os.dup2(write_fd, stderr_fd)
        os.close(write_fd)

        real_terminal = os.fdopen(self._saved_stdout_fd, "wb", buffering=0, closefd=False)
        log_file = self._log_file

        def pump() -> None:
            deduper = _LineDeduper()
            with os.fdopen(read_fd, "rb", buffering=0) as pipe_reader:
                while True:
                    chunk = pipe_reader.read(4096)
                    if not chunk:
                        break
                    real_terminal.write(chunk)
                    text = deduper.feed(chunk)
                    if text:
                        log_file.write(text)
                        log_file.flush()
            tail = deduper.flush()
            if tail:
                log_file.write(tail)
                log_file.flush()

        self._pump_thread = threading.Thread(target=pump, daemon=True)
        self._pump_thread.start()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        sys.stdout.flush()
        sys.stderr.flush()
        stdout_fd = sys.stdout.fileno()
        stderr_fd = sys.stderr.fileno()
        # Restoring the real fds closes the pipe's write end (nothing else
        # holds it open), which is what lets the pump thread's read() finally
        # return b"" and exit its loop. Must join the thread BEFORE closing
        # the saved fds below - the pump thread's real_terminal handle wraps
        # those same fd numbers (closefd=False), so closing them first races
        # the thread's own in-flight write() and can raise "Bad file
        # descriptor" if it loses.
        os.dup2(self._saved_stdout_fd, stdout_fd)
        os.dup2(self._saved_stderr_fd, stderr_fd)
        if self._pump_thread is not None:
            self._pump_thread.join(timeout=5)
        os.close(self._saved_stdout_fd)
        os.close(self._saved_stderr_fd)
        if self._log_file is not None:
            self._log_file.write(_run_footer(ok=exc_type is None, started=self._started))
            self._log_file.close()


def start_run_log(log_dir: Path, label: str, *, mode: str, argv: list[str]) -> TeeLog:
    """Convenience constructor - see TeeLog. Caller is expected to use this
    as a context manager: `with start_run_log(...) as log: ...`."""
    return TeeLog(log_file_path(log_dir, label), label=label, mode=mode, argv=argv)
