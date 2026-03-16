"""
CLI Runner Tool
---------------
Executes ExcelMCP3 CLI commands via subprocess, captures output,
detects failures, and returns structured results.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
import time
from dataclasses import dataclass
from pathlib import Path

from crewai.tools import tool


# ---------------------------------------------------------------------------
# Paths — configurable via environment variables
# ---------------------------------------------------------------------------

EXCEL_MCP_PYTHON = os.getenv(
    "EXCEL_MCP_PYTHON",
    r"c:\_Internal_ViewPoint_Projects\Jon-ExcelMCP3\venv\Scripts\python.exe",
)
EXCEL_MCP_MAIN = os.getenv(
    "EXCEL_MCP_MAIN",
    r"c:\_Internal_ViewPoint_Projects\Jon-ExcelMCP3\main.py",
)
EXCEL_MCP_CWD = os.getenv(
    "EXCEL_MCP_CWD",
    r"c:\_Internal_ViewPoint_Projects\Jon-ExcelMCP3",
)


# ---------------------------------------------------------------------------
# Result model
# ---------------------------------------------------------------------------

@dataclass
class CLIResult:
    command: str
    returncode: int
    stdout: str
    stderr: str
    duration: float
    success: bool

    def to_dict(self) -> dict:
        return {
            "command": self.command,
            "returncode": self.returncode,
            "success": self.success,
            "duration_seconds": round(self.duration, 2),
            "stdout_tail": self.stdout[-1000:] if len(self.stdout) > 1000 else self.stdout,
            "stderr_tail": self.stderr[-500:] if len(self.stderr) > 500 else self.stderr,
        }


# ---------------------------------------------------------------------------
# Error detection patterns
# ---------------------------------------------------------------------------

FAILURE_PATTERNS = [
    "traceback (most recent call last)",
    "error:",
    "exception:",
    "failed:",
    "cannot find",
    "file not found",
    "no such file",
    "permission denied",
    "modulenotfounderror",
    "valueerror:",
    "keyerror:",
    "typeerror:",
    "attributeerror:",
    "wf1 failed",
    "wf2 failed",
    "wf3 failed",
    "wf4 failed",
    "wf5 failed",
]

SUCCESS_PATTERNS = [
    "complete",
    "completed",
    "success",
    "finished",
    "saved",
    "generated",
    "wf1 complete",
    "wf3 complete",
    "wf5 request complete",
    "wf5 receive complete",
    "wf5 receive passed",
    "wf5 validate passed",
]


def detect_failure(stdout: str, stderr: str, returncode: int) -> bool:
    """
    Heuristic failure detection.

    Priority order:
    1. Non-zero returncode → always a failure.
    2. Clear success indicators in stdout → never a failure (overrides error patterns).
    3. Failure patterns in stdout only (stderr has HTTP/info logs — skip it).
    """
    if returncode != 0:
        return True
    stdout_lower = stdout.lower()
    # Success overrides all error-pattern matching
    for pat in SUCCESS_PATTERNS:
        if pat in stdout_lower:
            return False
    # Only scan stdout for errors, not stderr (stderr has noisy HTTP request logs)
    for pat in FAILURE_PATTERNS:
        if pat in stdout_lower:
            return True
    return False


# ---------------------------------------------------------------------------
# Progress ticker — prints elapsed time every 30s so the user sees activity
# ---------------------------------------------------------------------------

def _progress_ticker(workflow: str, start: float, stop_event: threading.Event) -> None:
    """Background thread: prints 'still running' every 30s until stop_event is set."""
    interval = 30
    next_tick = start + interval
    while not stop_event.wait(timeout=max(0, next_tick - time.time())):
        elapsed = int(time.time() - start)
        mins, secs = divmod(elapsed, 60)
        print(
            f"  [AI Crew] {workflow} still running... {mins}m {secs:02d}s elapsed",
            flush=True,
        )
        next_tick += interval


# ---------------------------------------------------------------------------
# Core runner — silent (used by agent tools)
# ---------------------------------------------------------------------------

def run_cli(args: list[str], timeout: int = 600) -> CLIResult:
    """
    Execute with a progress ticker, capture all output, return CLIResult.
    Used by the AI crew's run_workflow_command tool.
    """
    cmd = [EXCEL_MCP_PYTHON, EXCEL_MCP_MAIN] + args
    cmd_str = " ".join(str(c) for c in cmd)
    workflow = args[0].upper() if args else "workflow"

    print(f"\n  [AI Crew] Starting {workflow} — output captured, please wait...", flush=True)

    start = time.time()
    stop_event = threading.Event()
    ticker = threading.Thread(
        target=_progress_ticker, args=(workflow, start, stop_event), daemon=True
    )
    ticker.start()

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=EXCEL_MCP_CWD,
            timeout=timeout,
        )
        stop_event.set()
        ticker.join()
        duration = time.time() - start
        mins, secs = divmod(int(duration), 60)
        success = not detect_failure(result.stdout, result.stderr, result.returncode)
        status = "OK" if success else "FAILED"
        print(f"  [AI Crew] {workflow} {status} — {mins}m {secs:02d}s", flush=True)
        return CLIResult(
            command=cmd_str,
            returncode=result.returncode,
            stdout=result.stdout,
            stderr=result.stderr,
            duration=duration,
            success=success,
        )
    except subprocess.TimeoutExpired:
        stop_event.set()
        duration = time.time() - start
        print(f"  [AI Crew] {workflow} TIMEOUT after {timeout}s", flush=True)
        return CLIResult(
            command=cmd_str,
            returncode=-1,
            stdout="",
            stderr=f"Timeout after {timeout}s",
            duration=duration,
            success=False,
        )
    except FileNotFoundError as exc:
        stop_event.set()
        duration = time.time() - start
        print(f"  [AI Crew] {workflow} ERROR — executable not found", flush=True)
        return CLIResult(
            command=cmd_str,
            returncode=-1,
            stdout="",
            stderr=f"Executable not found: {exc}",
            duration=duration,
            success=False,
        )


# ---------------------------------------------------------------------------
# Streaming runner — live output for interactive use
# ---------------------------------------------------------------------------

def run_cli_streaming(args: list[str], timeout: int = 900) -> CLIResult:
    """
    Execute with live stdout printed to the terminal line-by-line.
    Used by the interactive --no-crew path so users see progress in real time.
    Returns a CLIResult with the full collected stdout once done.
    """
    cmd = [EXCEL_MCP_PYTHON, EXCEL_MCP_MAIN] + args
    cmd_str = " ".join(str(c) for c in cmd)

    stdout_lines: list[str] = []
    start = time.time()

    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,   # suppress noisy HTTP log lines
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=EXCEL_MCP_CWD,
        )

        # Stream stdout line by line — user sees output as it happens
        for line in proc.stdout:
            stripped = line.rstrip("\n")
            print(stripped, flush=True)
            stdout_lines.append(stripped)

        proc.wait(timeout=timeout)
        returncode = proc.returncode

    except subprocess.TimeoutExpired:
        proc.kill()
        return CLIResult(
            command=cmd_str, returncode=-1,
            stdout="\n".join(stdout_lines),
            stderr=f"Timeout after {timeout}s",
            duration=time.time() - start,
            success=False,
        )
    except FileNotFoundError as exc:
        return CLIResult(
            command=cmd_str, returncode=-1,
            stdout="", stderr=f"Executable not found: {exc}",
            duration=time.time() - start, success=False,
        )

    duration = time.time() - start
    full_stdout = "\n".join(stdout_lines)
    success = not detect_failure(full_stdout, "", returncode)
    return CLIResult(
        command=cmd_str,
        returncode=returncode,
        stdout=full_stdout,
        stderr="",
        duration=duration,
        success=success,
    )


# ---------------------------------------------------------------------------
# Agent @tool — thin wrapper used by ExecutorAgent
# ---------------------------------------------------------------------------

@tool("run_workflow_command")
def run_workflow_command(workflow_args_json: str) -> str:
    """
    Execute an ExcelMCP3 CLI workflow command.

    workflow_args_json must be a JSON string with:
    {
        "args": ["analyze", "--workbook", "...", "--client", "..."],
        "timeout": 600
    }

    Returns a JSON string with: command, success, returncode, stdout_tail, stderr_tail, duration_seconds.
    """
    try:
        payload = json.loads(workflow_args_json)
        args = payload.get("args", [])
        timeout = int(payload.get("timeout", 600))
    except Exception as exc:
        return json.dumps({"success": False, "error": f"Bad input JSON: {exc}"})

    result = run_cli(args, timeout=timeout)
    return json.dumps(result.to_dict(), indent=2)
