"""
Workflow Tasks
--------------
Factory functions that create CrewAI-style Tasks for each stage of an
ExcelMCP pipeline.

Each task is parameterised by the runtime context (selected workflow,
arguments, mode) rather than being static, so the same task definitions
handle every pipeline combination.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field

from crewai import Task, Agent


# ---------------------------------------------------------------------------
# Context — flows through the entire pipeline
# ---------------------------------------------------------------------------

@dataclass
class PipelineContext:
    """
    Runtime context shared across all tasks in a pipeline run.

    Populated incrementally as each task completes.
    """
    input_file: str
    client: str
    workbook_stem: str
    initial_workflow: str                  # "WF1", "WF2", ...
    initial_args: dict[str, str]           # CLI arguments for the first workflow
    mode: str = "interactive"              # "interactive" or "autonomous"
    client_url: str = ""

    # Filled in as pipeline progresses
    analysis_dir: str = ""                 # processed/{client}/{stem}/
    blueprint_db_path: str = ""
    selected_pipeline: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Task factories
# ---------------------------------------------------------------------------

def make_discovery_task(agent: "Agent", input_dir: str) -> Task:
    """Task 1 — Scan input directory and list available Excel files."""
    return Task(
        description=(
            f"Scan the directory '{input_dir}' for Excel workbooks.\n\n"
            "Use the scan_excel_files tool with that directory path.\n"
            "Report the count of files found and list each file with its name, "
            "path, and size. If no files are found, report that clearly."
        ),
        agent=agent,
        expected_output=(
            "A numbered list of Excel files found in the directory, "
            "with name, full path, and size in KB."
        ),
    )


def make_execution_task(
    agent: "Agent",
    ctx: PipelineContext,
    workflow: str,
    args: dict[str, str],
    prior_tasks: list[Task] | None = None,
) -> Task:
    """Task — Execute a specific ExcelMCP3 workflow via CLI."""
    from workflow_logic import WORKFLOW_COMMANDS, WF, WORKFLOW_NAMES

    wf_enum = WF(workflow)
    subcommand = WORKFLOW_COMMANDS[wf_enum]
    wf_name = WORKFLOW_NAMES[wf_enum]

    # Build the args array for the tool
    args_list = [subcommand]
    for flag, value in args.items():
        if value == "":
            args_list.append(flag)
        else:
            args_list.extend([flag, value])

    tool_payload = json.dumps({
        "args": args_list,
        "timeout": 900,  # 15 min max per workflow
    })

    description = (
        f"Execute {workflow} — {wf_name}.\n\n"
        f"Use the run_workflow_command tool with this exact JSON payload:\n"
        f"{tool_payload}\n\n"
        "After execution:\n"
        "1. Report whether it succeeded or failed (check the 'success' field).\n"
        "2. If failed, quote the error from 'stderr_tail'.\n"
        "3. If succeeded, report the execution duration and any key lines from stdout_tail "
        "   (look for 'complete', 'saved', 'generated', output paths).\n"
        "4. Extract any output file paths mentioned in stdout_tail and list them."
    )

    return Task(
        description=description,
        agent=agent,
        expected_output=(
            f"Execution status (success/failed), duration, output file paths, "
            f"and a brief summary of what {workflow} produced."
        ),
        context=prior_tasks or [],
    )


def make_output_analysis_task(
    agent: "Agent",
    ctx: PipelineContext,
    workflow_completed: str,
    prior_task: Task,
) -> Task:
    """Task — Inspect the outputs of a completed workflow."""
    analysis_dir = ctx.analysis_dir or (
        f"processed/{ctx.client}/{ctx.workbook_stem}"
        if ctx.client and ctx.workbook_stem else ""
    )

    if workflow_completed in ("WF1", "WF3"):
        tool_name = "parse_wf1_outputs" if workflow_completed == "WF1" else "parse_wf3_outputs"
        base_args = f'"{analysis_dir}", "{ctx.client}"'
        if workflow_completed == "WF1":
            base_args += f', "{ctx.input_file}"'

        description = (
            f"Analyse the output of {workflow_completed} at: {analysis_dir}\n\n"
            f"Call {tool_name} with:\n"
            f"  analysis_dir = '{analysis_dir}'\n"
            f"  client = '{ctx.client}'\n"
            + (f"  workbook_path = '{ctx.input_file}'\n" if workflow_completed == "WF1" else "")
            + "\n"
            "The tool reads the *-Analysis.json file in the documentation/ subdirectory.\n"
            "That file has a 'summary' dict with 'Power Queries Found' and 'VBA Modules Found' counts.\n\n"
            "Report ALL of these signals:\n"
            "- sql_request_found: true/false (sql_query_request/ directory has files)\n"
            "- power_query_found: true/false (Power Queries Found > 0 in summary)\n"
            "- vba_found: true/false (VBA Modules Found > 0 in summary)\n"
            "- blueprint_db_found: true/false and its full path\n\n"
            "Give a clear summary line: 'WF1 detected: Power Query (yes/no), VBA (yes/no), SQL request (yes/no), Blueprint DB (yes/no)'."
        )
    else:
        description = (
            f"Summarise the outputs of {workflow_completed}.\n\n"
            f"Based on the execution result above, list what files were generated "
            f"and whether the workflow was fully successful."
        )

    return Task(
        description=description,
        agent=agent,
        expected_output=(
            "Structured summary of detected signals: SQL request found (yes/no), "
            "Power Query found (yes/no), VBA found (yes/no), blueprint DB path."
        ),
        context=[prior_task],
    )


def make_planning_task(
    agent: "Agent",
    ctx: PipelineContext,
    analysis_task: Task,
    mode: str = "interactive",
) -> Task:
    """Task — Recommend next workflows based on output analysis."""
    description = (
        "Based on the output analysis above, recommend the next workflow(s) to run.\n\n"
        "Apply the decision matrix:\n"
        "  - SQL request found  → recommend WF5 (sql-gen --phase request --auto-vgpt2)\n"
        "  - Power Query found  → recommend WF3 (improve --analysis-dir <path>)\n"
        "  - VBA found          → recommend WF3 (improve --analysis-dir <path>)\n"
        "  - Blueprint DB found → recommend WF2 (create --analysis-db <path> --output <path>)\n\n"
        f"Client: {ctx.client}\n"
        f"Workbook stem: {ctx.workbook_stem}\n"
        f"Analysis dir:  processed/{ctx.client}/{ctx.workbook_stem}\n\n"
        "For each recommendation provide:\n"
        "  1. The workflow name (WF1/WF2/WF3/WF5)\n"
        "  2. The exact CLI command\n"
        "  3. A brief reason why it is recommended\n\n"
    )
    if mode == "autonomous":
        description += (
            "Since we are running in AUTONOMOUS mode, list the recommended workflows "
            "in priority order. The system will execute them automatically."
        )
    else:
        description += (
            "Present this as a numbered menu for the user to choose from."
        )

    return Task(
        description=description,
        agent=agent,
        expected_output=(
            "Ordered list of recommended next workflows with exact CLI commands and reasons."
        ),
        context=[analysis_task],
    )


def make_report_task(
    agent: "Agent",
    ctx: PipelineContext,
    prior_tasks: list[Task],
    tracker_report: str = "",
) -> Task:
    """Task — Generate the final pipeline execution report."""
    description = (
        "Generate the final pipeline execution report.\n\n"
        f"Pipeline: {' → '.join(ctx.selected_pipeline) if ctx.selected_pipeline else ctx.initial_workflow}\n"
        f"Input File: {ctx.input_file}\n"
        f"Client: {ctx.client}\n\n"
        "Synthesise all prior task outputs into a complete report using this format:\n"
        "  - Run ID\n"
        "  - Input file and client\n"
        "  - Workflows executed (with status and duration)\n"
        "  - All generated output files\n"
        "  - Overall status\n"
        "  - Recommended next actions (if any)\n\n"
    )
    if tracker_report:
        description += f"Pipeline Tracker Summary:\n{tracker_report}\n\n"

    description += "End with a clear 'PIPELINE COMPLETE' or 'PIPELINE FAILED' line."

    return Task(
        description=description,
        agent=agent,
        expected_output=(
            "Complete formatted pipeline execution report with all workflow statuses, "
            "output files, and recommended next steps."
        ),
        context=prior_tasks,
    )


# ---------------------------------------------------------------------------
# High-level builder
# ---------------------------------------------------------------------------

@dataclass
class TaskChain:
    """A complete set of tasks for a pipeline run."""
    tasks: list[Task]
    context: PipelineContext


def build_task_chain(
    ctx: PipelineContext,
    agents: dict,
) -> TaskChain:
    """
    Build the complete task chain for a pipeline run.

    agents dict keys:
      "file"     → FileDiscoveryAgent
      "executor" → CLIExecutionAgent
      "output"   → OutputAnalysisAgent
      "planner"  → WorkflowPlannerAgent
      "report"   → ReportAgent
    """
    tasks: list[Task] = []

    # Task 1: Execute initial workflow
    exec_task = make_execution_task(
        agent=agents["executor"],
        ctx=ctx,
        workflow=ctx.initial_workflow,
        args=ctx.initial_args,
    )
    tasks.append(exec_task)

    # Task 2: Analyse outputs (for WF1 and WF3)
    if ctx.initial_workflow in ("WF1", "WF3"):
        analysis_task = make_output_analysis_task(
            agent=agents["output"],
            ctx=ctx,
            workflow_completed=ctx.initial_workflow,
            prior_task=exec_task,
        )
        tasks.append(analysis_task)

        # Task 3: Plan next steps
        plan_task = make_planning_task(
            agent=agents["planner"],
            ctx=ctx,
            analysis_task=analysis_task,
            mode=ctx.mode,
        )
        tasks.append(plan_task)

        # Task 4: Report
        report_task = make_report_task(
            agent=agents["report"],
            ctx=ctx,
            prior_tasks=[exec_task, analysis_task, plan_task],
        )
        tasks.append(report_task)
    else:
        # For WF2/WF4/WF5 — just run and report
        report_task = make_report_task(
            agent=agents["report"],
            ctx=ctx,
            prior_tasks=[exec_task],
        )
        tasks.append(report_task)

    return TaskChain(tasks=tasks, context=ctx)
