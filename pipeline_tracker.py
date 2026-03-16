"""
Pipeline Tracker
----------------
SQLite-backed tracker that records every workflow execution in a pipeline run.
Produces the final execution report.
"""

from __future__ import annotations

import json
import sqlite3
import time
from contextlib import contextmanager
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterator

from workflow_logic import WF, WORKFLOW_NAMES


DB_PATH = Path(__file__).parent / "pipeline_runs.db"


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------

@dataclass
class WorkflowRun:
    run_id: str
    workflow: str           # WF1 .. WF5
    command: str            # full CLI command as string
    status: str = "pending" # pending | running | success | failed
    start_time: float = field(default_factory=time.time)
    end_time: float = 0.0
    duration_seconds: float = 0.0
    outputs: str = "[]"     # JSON list of output paths
    error: str = ""
    stdout_tail: str = ""   # last 500 chars of stdout


@dataclass
class PipelineRun:
    run_id: str
    input_file: str
    client: str
    created_at: str = field(default_factory=lambda: datetime.now().isoformat())
    status: str = "running"  # running | success | partial | failed
    pipeline_desc: str = ""  # e.g. "WF1 → WF3"
    mode: str = "interactive"


# ---------------------------------------------------------------------------
# Tracker
# ---------------------------------------------------------------------------

class PipelineTracker:
    """
    Tracks a complete pipeline run from start to finish.

    Usage::

        tracker = PipelineTracker(run_id="abc123", input_file="x.xlsx", client="Delnor")
        with tracker.track_workflow("WF1", "python main.py analyze ...") as run:
            # ... execute ...
            run.outputs = json.dumps(["processed/Delnor/x/..."])
        report = tracker.generate_report()
    """

    def __init__(self, run_id: str, input_file: str, client: str,
                 mode: str = "interactive"):
        self.run_id = run_id
        self.input_file = input_file
        self.client = client
        self.mode = mode
        self.workflow_runs: list[WorkflowRun] = []
        self._start_time = time.time()
        self._ensure_db()
        self._save_pipeline_run(PipelineRun(
            run_id=run_id, input_file=input_file, client=client, mode=mode
        ))

    # ------------------------------------------------------------------
    # Context manager for tracking one workflow execution
    # ------------------------------------------------------------------

    @contextmanager
    def track_workflow(
        self, workflow: str, command: str
    ) -> Iterator[WorkflowRun]:
        """
        Context manager that records start/end time and status of a workflow run.

        Example::

            with tracker.track_workflow("WF1", cmd_str) as run:
                result = subprocess.run(...)
                if result.returncode == 0:
                    run.status = "success"
                    run.outputs = json.dumps(output_paths)
                else:
                    run.status = "failed"
                    run.error = result.stderr[-500:]
        """
        wrun = WorkflowRun(
            run_id=self.run_id,
            workflow=workflow,
            command=command,
            status="running",
            start_time=time.time(),
        )
        self.workflow_runs.append(wrun)
        self._save_workflow_run(wrun)
        try:
            yield wrun
        finally:
            wrun.end_time = time.time()
            wrun.duration_seconds = wrun.end_time - wrun.start_time
            if wrun.status == "running":
                wrun.status = "success"
            self._update_workflow_run(wrun)

    # ------------------------------------------------------------------
    # Report generation
    # ------------------------------------------------------------------

    def generate_report(self) -> str:
        """Return a rich-formatted text report of the pipeline execution."""
        total_elapsed = time.time() - self._start_time
        minutes, seconds = divmod(int(total_elapsed), 60)

        success_count = sum(1 for r in self.workflow_runs if r.status == "success")
        fail_count = sum(1 for r in self.workflow_runs if r.status == "failed")
        overall = "Successful" if fail_count == 0 else f"Partial ({fail_count} failed)"

        pipeline_str = " → ".join(r.workflow for r in self.workflow_runs)

        # Collect all outputs
        all_outputs: list[str] = []
        for r in self.workflow_runs:
            try:
                all_outputs.extend(json.loads(r.outputs))
            except Exception:
                pass

        lines = [
            "=" * 60,
            "  ExcelMCP Pipeline Execution Report",
            "=" * 60,
            f"  Run ID:        {self.run_id}",
            f"  Input File:    {self.input_file}",
            f"  Client:        {self.client}",
            f"  Mode:          {self.mode}",
            "",
            f"  Workflows:     {pipeline_str}",
            f"  Status:        {overall}",
            f"  Execution Time:{minutes}m {seconds}s",
            "",
            "  Workflow Details:",
        ]

        for i, r in enumerate(self.workflow_runs, 1):
            dur = f"{r.duration_seconds:.1f}s"
            status_icon = "OK" if r.status == "success" else "FAIL"
            lines.append(f"    {i}. {r.workflow} ({WORKFLOW_NAMES.get(WF(r.workflow), r.workflow)}) — {status_icon} [{dur}]")
            if r.error:
                lines.append(f"       Error: {r.error[:120]}")

        if all_outputs:
            lines += ["", "  Outputs Generated:"]
            for out in all_outputs:
                lines.append(f"    • {out}")

        lines += ["", "=" * 60]
        return "\n".join(lines)

    # ------------------------------------------------------------------
    # SQLite persistence
    # ------------------------------------------------------------------

    def _ensure_db(self) -> None:
        with self._conn() as conn:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS pipeline_runs (
                    run_id TEXT PRIMARY KEY,
                    input_file TEXT,
                    client TEXT,
                    created_at TEXT,
                    status TEXT,
                    pipeline_desc TEXT,
                    mode TEXT
                );
                CREATE TABLE IF NOT EXISTS workflow_runs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    run_id TEXT,
                    workflow TEXT,
                    command TEXT,
                    status TEXT,
                    start_time REAL,
                    end_time REAL,
                    duration_seconds REAL,
                    outputs TEXT,
                    error TEXT,
                    stdout_tail TEXT
                );
            """)

    def _save_pipeline_run(self, p: PipelineRun) -> None:
        with self._conn() as conn:
            conn.execute(
                "INSERT OR REPLACE INTO pipeline_runs VALUES (?,?,?,?,?,?,?)",
                (p.run_id, p.input_file, p.client, p.created_at,
                 p.status, p.pipeline_desc, p.mode)
            )

    def _save_workflow_run(self, r: WorkflowRun) -> None:
        with self._conn() as conn:
            conn.execute(
                "INSERT INTO workflow_runs "
                "(run_id,workflow,command,status,start_time,end_time,"
                "duration_seconds,outputs,error,stdout_tail) "
                "VALUES (?,?,?,?,?,?,?,?,?,?)",
                (r.run_id, r.workflow, r.command, r.status, r.start_time,
                 r.end_time, r.duration_seconds, r.outputs, r.error,
                 r.stdout_tail)
            )

    def _update_workflow_run(self, r: WorkflowRun) -> None:
        with self._conn() as conn:
            conn.execute(
                "UPDATE workflow_runs SET status=?,end_time=?,duration_seconds=?,"
                "outputs=?,error=?,stdout_tail=? "
                "WHERE run_id=? AND workflow=? AND start_time=?",
                (r.status, r.end_time, r.duration_seconds, r.outputs,
                 r.error, r.stdout_tail, r.run_id, r.workflow, r.start_time)
            )

    @contextmanager
    def _conn(self) -> Iterator[sqlite3.Connection]:
        conn = sqlite3.connect(DB_PATH)
        try:
            yield conn
            conn.commit()
        finally:
            conn.close()
