#!/usr/bin/env python3
"""Startet watcher.py OS-abhängig im Hintergrund und schreibt eine PID-Datei."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
VENV_PYTHON = BASE / ".venv" / ("Scripts/python.exe" if sys.platform == "win32" else "bin/python")
PYTHON = str(VENV_PYTHON) if VENV_PYTHON.exists() else sys.executable
WATCHER = BASE / "watcher.py"
LOG_FILE = BASE / "watcher.log"
PID_FILE = BASE / "watcher.pid"


def running_pid() -> int | None:
    if not PID_FILE.exists():
        return None
    try:
        pid = int(PID_FILE.read_text().strip())
    except ValueError:
        return None
    if sys.platform == "win32":
        import ctypes

        handle = ctypes.windll.kernel32.OpenProcess(0x1000, 0, pid)
        if not handle:
            return None
        ctypes.windll.kernel32.CloseHandle(handle)
        return pid
    import os

    try:
        os.kill(pid, 0)
        return pid
    except OSError:
        return None


def start() -> subprocess.Popen:
    log = open(LOG_FILE, "a")
    common = dict(stdout=log, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL, cwd=str(BASE))
    if sys.platform == "win32":
        DETACHED_PROCESS = 0x00000008
        CREATE_NEW_PROCESS_GROUP = 0x00000200
        return subprocess.Popen(
            [PYTHON, str(WATCHER)],
            creationflags=DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP,
            **common,
        )
    return subprocess.Popen([PYTHON, str(WATCHER)], start_new_session=True, **common)


def main() -> None:
    pid = running_pid()
    if pid:
        print(f"watcher.py läuft bereits (PID {pid}).")
        return

    proc = start()
    PID_FILE.write_text(str(proc.pid))
    stop_hint = f"taskkill /PID {proc.pid} /F" if sys.platform == "win32" else f"kill {proc.pid}"
    print(f"watcher.py gestartet im Hintergrund (PID {proc.pid}).")
    print(f"Log: {LOG_FILE}")
    print(f"Beenden mit: {stop_hint}")


if __name__ == "__main__":
    main()
