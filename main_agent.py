"""
ExcelMCP Agent — Main Entry Point
==================================

AI agent that autonomously orchestrates the ExcelMCP3 workflow pipeline.

Usage
-----
Interactive mode (default):
    python main_agent.py

Autonomous mode (auto-selects WF1→WF3/WF5 based on output analysis):
    python main_agent.py --autonomous

Pre-selected workflow (skip interactive prompts):
    python main_agent.py --workbook "Delnor Security Workbook V11.1.xlsx" --workflow 1

Batch autonomous:
    python main_agent.py --autonomous --workflow 4 --input-dir inputs/
"""

from __future__ import annotations

import argparse
import os
import sys
import time
import uuid
from pathlib import Path

from dotenv import load_dotenv
from rich.console import Console
from rich.panel import Panel
from rich.prompt import Confirm, Prompt
from rich.table import Table

# Load .env from this project directory
load_dotenv(Path(__file__).parent / ".env")

# Add project root to sys.path so all imports resolve
sys.path.insert(0, str(Path(__file__).parent))

from agents import (
    create_cli_execution_agent,
    create_file_discovery_agent,
    create_output_analysis_agent,
    create_report_agent,
    create_workflow_planner_agent,
)
from crewai import Crew, Process
from pipeline_tracker import PipelineTracker
from tasks.workflow_tasks import PipelineContext, TaskChain, build_task_chain
from tools.file_scanner import scan_excel_files_raw
from workflow_logic import (
    ARGUMENT_PROMPTS,
    WF,
    WORKFLOW_COMMANDS,
    WORKFLOW_NAMES,
    WORKFLOW_REQUIRED_ARGS,
    WorkflowLogic,
)

console = Console()

EXCEL_MCP_INPUT_DIR = os.getenv(
    "EXCEL_MCP_INPUT_DIR",
    r"c:\_Internal_ViewPoint_Projects\Jon-ExcelMCP3\inputs",
)


# ---------------------------------------------------------------------------
# Interactive helpers
# ---------------------------------------------------------------------------

def select_file_interactive(input_dir: str) -> Path | None:
    """Prompt user to select from available Excel files."""
    files = scan_excel_files_raw(input_dir)
    if not files:
        console.print(f"[red]No Excel files found in: {input_dir}[/red]")
        return None

    table = Table(title=f"Excel Files in {Path(input_dir).name}/", show_lines=True)
    table.add_column("#", style="cyan bold", width=4)
    table.add_column("File", style="white")
    table.add_column("Size", style="dim", justify="right")

    for i, f in enumerate(files, 1):
        size_kb = f"{f.stat().st_size / 1024:.1f} KB"
        table.add_row(str(i), f.name, size_kb)

    console.print(table)
    console.print(f"\n[dim]Found {len(files)} Excel file(s)[/dim]")

    raw = Prompt.ask(
        "\nSelect file number",
        default="1",
        console=console,
    )
    try:
        idx = int(raw.strip()) - 1
        if 0 <= idx < len(files):
            return files[idx]
        console.print("[red]Invalid selection.[/red]")
        return None
    except ValueError:
        console.print("[red]Please enter a number.[/red]")
        return None


def select_workflow_interactive() -> WF | None:
    """Prompt user to select a workflow."""
    table = Table(title="Available Workflows", show_lines=True)
    table.add_column("#", style="cyan bold", width=4)
    table.add_column("Workflow", style="white")
    table.add_column("Description", style="dim")

    descriptions = {
        WF.WF1: "Analyze a workbook — extract VBA, Power Query, formulas, SQL",
        WF.WF2: "Create a workbook from YAML config or WF1 blueprint database",
        WF.WF3: "Improve a workbook — refactor VBA, M code, and formulas with AI",
        WF.WF4: "Batch process multiple workbooks in parallel",
        WF.WF5: "Generate optimized SQL queries for Power Query connections",
    }

    for i, wf in enumerate([WF.WF1, WF.WF2, WF.WF3, WF.WF4, WF.WF5], 1):
        table.add_row(str(i), f"{wf.value} — {WORKFLOW_NAMES[wf]}", descriptions[wf])

    console.print(table)

    raw = Prompt.ask("\nSelect workflow number", default="1", console=console)
    wf_map = {1: WF.WF1, 2: WF.WF2, 3: WF.WF3, 4: WF.WF4, 5: WF.WF5}
    try:
        return wf_map.get(int(raw.strip()))
    except ValueError:
        return None


def collect_args_interactive(
    wf: WF,
    selected_file: Path | None = None,
    client: str = "",
) -> dict[str, str]:
    """
    Collect required and optional CLI arguments for the selected workflow.
    Pre-fills defaults where possible.
    """
    args: dict[str, str] = {}

    console.print(f"\n[bold]Collecting arguments for {wf.value} — {WORKFLOW_NAMES[wf]}[/bold]")

    if wf == WF.WF1:
        # --workbook
        default_wb = str(selected_file) if selected_file else ""
        args["--workbook"] = Prompt.ask(
            "[cyan]--workbook[/cyan] (path to workbook)",
            default=default_wb,
            console=console,
        )
        # --client (derive from file stem if not provided)
        derived_client = client or (
            selected_file.stem.split()[0] if selected_file else ""
        )
        args["--client"] = Prompt.ask(
            "[cyan]--client[/cyan] (client name)",
            default=derived_client,
            console=console,
        )
        # --client-url (optional)
        url = Prompt.ask(
            "[cyan]--client-url[/cyan] (optional, press Enter to skip)",
            default="",
            console=console,
        )
        if url.strip():
            args["--client-url"] = url.strip()

    elif wf == WF.WF2:
        mode = Prompt.ask(
            "WF2 mode: [cyan]1[/cyan]=blueprint (from WF1 .db), [cyan]2[/cyan]=config (YAML)",
            choices=["1", "2"],
            default="1",
            console=console,
        )
        if mode == "1":
            args["--analysis-db"] = Prompt.ask(
                "[cyan]--analysis-db[/cyan] (path to WF1 .db file)",
                console=console,
            )
            default_output = "generated_workbooks/rebuilt.xlsx"
            args["--output"] = Prompt.ask(
                "[cyan]--output[/cyan] (output .xlsx path)",
                default=default_output,
                console=console,
            )
        else:
            args["--config"] = Prompt.ask(
                "[cyan]--config[/cyan] (path to YAML config)",
                console=console,
            )
        if Confirm.ask("Enable --live-validation (requires DB credentials)?", default=False):
            args["--live-validation"] = ""

    elif wf == WF.WF3:
        # Derive analysis-dir from file/client if available
        default_dir = ""
        if selected_file and client:
            default_dir = f"processed/{client}/{selected_file.stem}"
        args["--analysis-dir"] = Prompt.ask(
            "[cyan]--analysis-dir[/cyan] (path to WF1 analysis output dir)",
            default=default_dir,
            console=console,
        )
        if Confirm.ask("Enable --auto-vgpt2 (auto-invoke VGPT2 for SQL)?", default=False):
            args["--auto-vgpt2"] = ""

    elif wf == WF.WF4:
        default_input = str(Path(EXCEL_MCP_INPUT_DIR))
        args["--input-dir"] = Prompt.ask(
            "[cyan]--input-dir[/cyan] (directory to scan for workbooks)",
            default=default_input,
            console=console,
        )
        mode = Prompt.ask(
            "[cyan]--mode[/cyan] (analyze or improve)",
            choices=["analyze", "improve"],
            default="analyze",
            console=console,
        )
        args["--mode"] = mode
        workers = Prompt.ask(
            "[cyan]--workers[/cyan] (parallel workers)",
            default="4",
            console=console,
        )
        args["--workers"] = workers
        if client:
            args["--client"] = client
        if Confirm.ask("Enable --resume (skip already-processed workbooks)?", default=True):
            args["--resume"] = ""

    elif wf == WF.WF5:
        derived_client = client or Prompt.ask("[cyan]--client[/cyan]", console=console)
        args["--client"] = derived_client
        default_stem = selected_file.stem if selected_file else ""
        args["--workbook"] = Prompt.ask(
            "[cyan]--workbook[/cyan] (workbook stem / folder name under processed/{client}/)",
            default=default_stem,
            console=console,
        )
        phase = Prompt.ask(
            "[cyan]--phase[/cyan] (request / receive / validate / status)",
            choices=["request", "receive", "validate", "status"],
            default="request",
            console=console,
        )
        args["--phase"] = phase
        if Confirm.ask("Enable --auto-vgpt2?", default=True):
            args["--auto-vgpt2"] = ""

    return args


# ---------------------------------------------------------------------------
# Autonomous pipeline runner
# ---------------------------------------------------------------------------

def run_autonomous_pipeline(
    ctx: PipelineContext,
    tracker: PipelineTracker,
    agents: dict,
) -> str:
    """
    Autonomous mode: runs WF1, inspects outputs, then automatically runs
    recommended follow-on workflows (WF3 and/or WF5).

    Returns the final report string.
    """
    from tools.cli_runner import run_cli
    from tools.output_parser import parse_wf1_outputs_internal
    from workflow_logic import WorkflowLogic, WORKFLOW_COMMANDS

    console.rule("[bold yellow]AUTONOMOUS MODE — WF1[/bold yellow]")

    # Step 1: Run WF1
    wf1_args = [WORKFLOW_COMMANDS[WF.WF1]]
    for flag, value in ctx.initial_args.items():
        if value == "":
            wf1_args.append(flag)
        else:
            wf1_args.extend([flag, value])

    cmd_str = " ".join(wf1_args)
    with tracker.track_workflow("WF1", cmd_str) as run:
        console.print("[cyan]Running WF1 — Expert Analysis...[/cyan]")
        result = run_cli(wf1_args, timeout=900)
        run.stdout_tail = result.stdout[-500:]
        run.status = "success" if result.success else "failed"
        run.error = result.stderr[-300:] if not result.success else ""
        import json as _json
        import re as _re
        # Extract output paths from stdout
        output_paths = _re.findall(r"processed/[^\s]+", result.stdout)
        run.outputs = _json.dumps(list(set(output_paths)))
        if result.success:
            console.print(f"[green]WF1 completed in {result.duration:.1f}s[/green]")
        else:
            console.print(f"[red]WF1 FAILED after {result.duration:.1f}s[/red]")
            console.print(f"[red]{result.stderr[-300:]}[/red]")
            return tracker.generate_report()

    # Step 2: Derive analysis dir and parse outputs
    analysis_dir = WorkflowLogic.derive_wf1_analysis_dir(
        ctx.initial_args.get("--workbook", ""), ctx.client
    )
    ctx.analysis_dir = analysis_dir
    ctx.blueprint_db_path = WorkflowLogic.derive_blueprint_db_path(
        ctx.initial_args.get("--workbook", ""), ctx.client
    )

    signals = parse_wf1_outputs_internal(analysis_dir, ctx.client,
                                          ctx.initial_args.get("--workbook", ""))
    console.print(
        f"\n[bold]WF1 Output Signals:[/bold]\n"
        f"  SQL Request:   {'[green]YES[/green]' if signals.sql_request_found else '[dim]no[/dim]'}\n"
        f"  Power Query:   {'[green]YES[/green]' if signals.power_query_found else '[dim]no[/dim]'}\n"
        f"  VBA Modules:   {'[green]YES[/green]' if signals.vba_found else '[dim]no[/dim]'}\n"
        f"  Blueprint DB:  {'[green]YES[/green]' if signals.blueprint_db_found else '[dim]no[/dim]'}"
    )

    # Step 3: Get recommendations
    recs = WorkflowLogic.recommendations_after_wf1(signals)
    if not recs:
        console.print("\n[yellow]No follow-on workflows recommended.[/yellow]")
        return tracker.generate_report()

    console.print("\n[bold]Recommended follow-on workflows:[/bold]")
    for i, rec in enumerate(recs, 1):
        console.print(f"  {i}. [cyan]{rec.workflow.value}[/cyan] — {rec.reason}")

    # Step 4: Execute recommendations autonomously
    for rec in recs:
        ctx.selected_pipeline.append(rec.workflow.value)
        wf_args = [WORKFLOW_COMMANDS[rec.workflow]]
        for flag, value in rec.suggested_args.items():
            if value == "":
                wf_args.append(flag)
            else:
                wf_args.extend([flag, value])

        console.rule(f"[bold yellow]AUTONOMOUS — {rec.workflow.value}[/bold yellow]")
        console.print(f"[cyan]Running {rec.workflow.value}...[/cyan]")
        cmd_str2 = " ".join(wf_args)

        with tracker.track_workflow(rec.workflow.value, cmd_str2) as run:
            result = run_cli(wf_args, timeout=900)
            run.stdout_tail = result.stdout[-500:]
            run.status = "success" if result.success else "failed"
            run.error = result.stderr[-300:] if not result.success else ""
            import json as _json2
            import re as _re2
            paths = _re2.findall(r"processed/[^\s]+", result.stdout)
            run.outputs = _json2.dumps(list(set(paths)))
            status_txt = "green" if result.success else "red"
            console.print(
                f"[{status_txt}]{rec.workflow.value} {'completed' if result.success else 'FAILED'} "
                f"in {result.duration:.1f}s[/{status_txt}]"
            )

    return tracker.generate_report()


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="ExcelMCP Agent — AI-Powered Workflow Orchestrator"
    )
    parser.add_argument(
        "--autonomous", action="store_true",
        help="Autonomous mode: auto-run WF1 then follow recommended pipelines"
    )
    parser.add_argument(
        "--workbook", default=None,
        help="Pre-select workbook path (skips file selection prompt)"
    )
    parser.add_argument(
        "--client", default=None,
        help="Pre-select client name"
    )
    parser.add_argument(
        "--workflow", type=int, choices=[1, 2, 3, 4, 5], default=None,
        help="Pre-select workflow (1=WF1, 2=WF2, 3=WF3, 4=WF4, 5=WF5)"
    )
    parser.add_argument(
        "--input-dir", default=EXCEL_MCP_INPUT_DIR,
        help="Input directory to scan for Excel files"
    )
    parser.add_argument(
        "--no-crew", action="store_true",
        help="Skip AI crew steps (execute + report only) for faster runs"
    )
    parser.add_argument(
        "--verbose", action="store_true", default=True,
        help="Enable verbose agent output"
    )
    args = parser.parse_args()

    # ---------------------------------------------------------------------------
    # Banner
    # ---------------------------------------------------------------------------
    console.print(Panel.fit(
        "[bold blue]ExcelMCP Agent[/bold blue]\n"
        "[dim]AI-Powered Workflow Orchestrator for ExcelMCP3[/dim]",
        border_style="blue",
    ))

    # Check that at least one LLM API key is present when using AI crew mode
    if not args.no_crew and not args.autonomous:
        has_key = any([
            os.getenv("ANTHROPIC_API_KEY"),
            os.getenv("XAI_API_KEY"),
            os.getenv("GROK_API_KEY"),
            os.getenv("OPENAI_API_KEY"),
        ])
        if not has_key:
            console.print(
                "[yellow]Warning: No LLM API key found (ANTHROPIC_API_KEY / GROK_API_KEY / "
                "OPENAI_API_KEY). AI crew mode will fail. Use --no-crew for CLI-only mode.[/yellow]"
            )

    # ---------------------------------------------------------------------------
    # Step 1: File selection
    # ---------------------------------------------------------------------------
    if args.workbook:
        selected_file = Path(args.workbook)
        console.print(f"[green]Using workbook:[/green] {selected_file.name}")
    else:
        console.rule("[bold]Step 1 — Discover Excel Files[/bold]")
        selected_file = select_file_interactive(args.input_dir)
        if not selected_file:
            sys.exit(1)
        console.print(f"\n[green]Selected:[/green] {selected_file.name}")

    # ---------------------------------------------------------------------------
    # Step 2: Workflow selection
    # ---------------------------------------------------------------------------
    wf_map = {1: WF.WF1, 2: WF.WF2, 3: WF.WF3, 4: WF.WF4, 5: WF.WF5}
    if args.autonomous:
        # Autonomous always starts with WF1
        selected_wf = WF.WF1
        console.print("\n[bold cyan]Autonomous Mode — starting with WF1 (Analyze)[/bold cyan]")
    elif args.workflow:
        selected_wf = wf_map[args.workflow]
        console.print(f"\n[green]Workflow:[/green] {selected_wf.value} — {WORKFLOW_NAMES[selected_wf]}")
    else:
        console.rule("[bold]Step 2 — Select Workflow[/bold]")
        selected_wf = select_workflow_interactive()
        if not selected_wf:
            console.print("[red]Invalid workflow selection.[/red]")
            sys.exit(1)
        console.print(f"\n[green]Selected:[/green] {selected_wf.value} — {WORKFLOW_NAMES[selected_wf]}")

    # ---------------------------------------------------------------------------
    # Step 3: Collect arguments
    # ---------------------------------------------------------------------------
    client = args.client or ""
    if selected_wf == WF.WF1 and not client:
        # Auto-derive client from filename stem (first word)
        client = selected_file.stem.split()[0] if selected_file else ""

    cli_args: dict[str, str]
    if args.autonomous:
        # For autonomous WF1, just need --workbook and --client
        console.rule("[bold]Step 3 — Configure WF1 Arguments[/bold]")
        if not client:
            client = Prompt.ask("[cyan]--client[/cyan] (client name)", console=console)
        cli_args = {
            "--workbook": str(selected_file),
            "--client": client,
        }
        url = Prompt.ask(
            "[cyan]--client-url[/cyan] (optional, press Enter to skip)",
            default="",
            console=console,
        )
        if url.strip():
            cli_args["--client-url"] = url.strip()
    else:
        console.rule("[bold]Step 3 — Collect Arguments[/bold]")
        cli_args = collect_args_interactive(selected_wf, selected_file, client)
        # Always sync client from what user actually typed (may differ from auto-derived default)
        if "--client" in cli_args:
            client = cli_args["--client"]

    # ---------------------------------------------------------------------------
    # Build pipeline context
    # ---------------------------------------------------------------------------
    run_id = str(uuid.uuid4())[:8]
    workbook_stem = selected_file.stem if selected_file else ""

    ctx = PipelineContext(
        input_file=str(selected_file),
        client=client,
        workbook_stem=workbook_stem,
        initial_workflow=selected_wf.value,
        initial_args=cli_args,
        mode="autonomous" if args.autonomous else "interactive",
        client_url=cli_args.get("--client-url", ""),
        analysis_dir=WorkflowLogic.derive_wf1_analysis_dir(str(selected_file), client),
    )
    ctx.selected_pipeline = [selected_wf.value]

    tracker = PipelineTracker(
        run_id=run_id,
        input_file=str(selected_file),
        client=client,
        mode=ctx.mode,
    )

    # ---------------------------------------------------------------------------
    # Step 4: Execute
    # ---------------------------------------------------------------------------
    console.rule("[bold]Step 4 — Execute[/bold]")

    # ---- Autonomous mode ----
    if args.autonomous:
        report = run_autonomous_pipeline(ctx, tracker, agents={})
        console.print(Panel(report, title="[bold green]Pipeline Report[/bold green]", expand=False))
        return

    # ---- CLI-only mode (--no-crew) ----
    if args.no_crew:
        import json as _j
        import re as _r
        from tools.cli_runner import run_cli_streaming
        from tools.output_parser import parse_wf1_outputs_internal

        # Loop so user can chain follow-on workflows after WF1
        current_wf = selected_wf
        current_args = cli_args

        while True:
            wf_args_list = [WORKFLOW_COMMANDS[current_wf]]
            for flag, value in current_args.items():
                if value == "":
                    wf_args_list.append(flag)
                else:
                    wf_args_list.extend([flag, value])

            # Print what we're about to run
            cmd_str = " ".join(f'"{a}"' if " " in a else a for a in wf_args_list)
            console.rule(f"[bold cyan]{current_wf.value} — {WORKFLOW_NAMES[current_wf]}[/bold cyan]")
            console.print(f"[dim]python main.py {cmd_str}[/dim]\n")

            # Execute with live streaming output
            with tracker.track_workflow(current_wf.value, cmd_str) as run:
                result = run_cli_streaming(wf_args_list, timeout=900)
                run.stdout_tail = result.stdout[-500:]
                run.status = "success" if result.success else "failed"
                run.error = result.stderr[-300:] if not result.success else ""
                run.outputs = _j.dumps(list(set(
                    _r.findall(r"processed/[^\s\)\"]+", result.stdout)
                )))

            console.print()
            if result.success:
                console.print(
                    f"[green]{current_wf.value} completed in {result.duration:.0f}s[/green]"
                )
            else:
                console.print(
                    f"[red]{current_wf.value} FAILED after {result.duration:.0f}s[/red]"
                )

            # ---- After WF1: inspect outputs and offer follow-on workflows ----
            if current_wf == WF.WF1 and result.success:
                analysis_dir = WorkflowLogic.derive_wf1_analysis_dir(
                    current_args.get("--workbook", str(selected_file)), client
                )
                signals = parse_wf1_outputs_internal(analysis_dir, client)
                recs = WorkflowLogic.recommendations_after_wf1(signals)

                if recs:
                    console.print()
                    console.rule("[bold]Recommended Next Workflows[/bold]")
                    for i, rec in enumerate(recs, 1):
                        console.print(
                            f"  [cyan]{i}.[/cyan] [bold]{rec.workflow.value}[/bold] "
                            f"— {WORKFLOW_NAMES[rec.workflow]}\n"
                            f"     [dim]{rec.reason}[/dim]"
                        )
                    console.print(f"  [cyan]{len(recs)+1}.[/cyan] Stop — print report and exit")
                    console.print()

                    choice = Prompt.ask(
                        "Select next workflow (or stop)",
                        default=str(len(recs) + 1),
                        console=console,
                    )
                    try:
                        idx = int(choice.strip()) - 1
                        if 0 <= idx < len(recs):
                            rec = recs[idx]
                            current_wf = rec.workflow
                            current_args = rec.suggested_args
                            ctx.selected_pipeline.append(current_wf.value)
                            console.print()
                            continue   # run the chosen follow-on workflow
                    except ValueError:
                        pass

            # User chose stop, or no recommendations, or non-WF1 finished — exit loop
            break

        report = tracker.generate_report()
        console.print()
        console.print(Panel(report, title="[bold green]Pipeline Report[/bold green]", expand=False))
        return

    # ---- AI Crew mode ----
    all_agents = {
        "file": create_file_discovery_agent(verbose=args.verbose),
        "planner": create_workflow_planner_agent(verbose=args.verbose),
        "executor": create_cli_execution_agent(verbose=args.verbose),
        "output": create_output_analysis_agent(verbose=args.verbose),
        "report": create_report_agent(verbose=args.verbose),
    }

    chain = build_task_chain(ctx, all_agents)
    crew = Crew(
        agents=list(all_agents.values()),
        tasks=chain.tasks,
        process=Process.sequential,
        verbose=args.verbose,
    )

    console.print(
        f"[dim]Run ID: {run_id} | Mode: {ctx.mode} | "
        f"Tasks: {len(chain.tasks)}[/dim]\n"
    )

    # Track WF1 (+ analysis/planning) as a single tracker entry
    wf1_cmd = (
        f"analyze --workbook \"{ctx.input_file}\" --client {ctx.client}"
    )
    with tracker.track_workflow(ctx.initial_workflow, wf1_cmd) as wf1_run:
        crew_output = crew.kickoff(inputs={
            "input_file": ctx.input_file,
            "client": ctx.client,
            "workbook_stem": ctx.workbook_stem,
            "analysis_dir": ctx.analysis_dir,
        })
        wf1_run.status = "success"

    # ---------------------------------------------------------------------------
    # Step 5: After WF1 crew — show signals, ask user, then run confirmed workflows
    # ---------------------------------------------------------------------------
    import json as _j
    import re as _r
    from tools.cli_runner import run_cli
    from tools.output_parser import parse_wf1_outputs_internal

    wf1_signals = None
    wf1_recs = []
    followon_results: list[dict] = []  # track results for detailed report

    if ctx.initial_workflow == "WF1":
        wf1_signals = parse_wf1_outputs_internal(
            ctx.analysis_dir, ctx.client, ctx.input_file
        )
        wf1_recs = WorkflowLogic.recommendations_after_wf1(wf1_signals)

        # ---- Show detected signals ----
        console.print()
        console.rule("[bold]WF1 Analysis — Detected Signals[/bold]")
        sig_table = Table(show_header=False, box=None, padding=(0, 2))
        sig_table.add_column("Signal", style="bold")
        sig_table.add_column("Value")
        sig_table.add_row(
            "Power Queries",
            "[green]FOUND[/green]" if wf1_signals.power_query_found else "[dim]not found[/dim]"
        )
        sig_table.add_row(
            "VBA Modules",
            "[green]FOUND[/green]" if wf1_signals.vba_found else "[dim]not found[/dim]"
        )
        sig_table.add_row(
            "SQL Request",
            "[green]FOUND[/green]" if wf1_signals.sql_request_found else "[dim]not found[/dim]"
        )
        sig_table.add_row(
            "Blueprint DB",
            f"[green]FOUND[/green]  [dim]{wf1_signals.blueprint_db_path}[/dim]"
            if wf1_signals.blueprint_db_found else "[dim]not found[/dim]"
        )
        console.print(sig_table)

        if not wf1_recs:
            console.print("\n[yellow]No follow-on workflows recommended for this workbook.[/yellow]")
        else:
            # ---- Show recommendations with reasons ----
            console.print()
            console.rule("[bold]Recommended Next Workflows[/bold]")
            rec_table = Table(show_header=True, box=None, padding=(0, 2))
            rec_table.add_column("#", style="cyan bold", width=3)
            rec_table.add_column("Workflow", style="bold", width=12)
            rec_table.add_column("Name", width=22)
            rec_table.add_column("Why recommended", style="dim")
            for i, rec in enumerate(wf1_recs, 1):
                rec_table.add_row(
                    str(i),
                    rec.workflow.value,
                    WORKFLOW_NAMES[rec.workflow],
                    rec.reason,
                )
            console.print(rec_table)
            console.print()

            # ---- Ask user which to run ----
            console.print(
                "  Type [cyan]all[/cyan] to run all, "
                "[cyan]1 2 3[/cyan] for specific ones, "
                "or [cyan]none[/cyan] to skip."
            )
            choice_raw = Prompt.ask(
                "  Which workflows should run?",
                default="all",
                console=console,
            ).strip().lower()

            selected_recs: list = []
            if choice_raw == "all":
                selected_recs = list(wf1_recs)
            elif choice_raw not in ("none", "n", ""):
                for token in choice_raw.replace(",", " ").split():
                    try:
                        idx = int(token) - 1
                        if 0 <= idx < len(wf1_recs):
                            selected_recs.append(wf1_recs[idx])
                    except ValueError:
                        pass

            # ---- Execute confirmed workflows ----
            for rec in selected_recs:
                wf_args_list = [WORKFLOW_COMMANDS[rec.workflow]]
                for flag, value in rec.suggested_args.items():
                    if value == "":
                        wf_args_list.append(flag)
                    else:
                        wf_args_list.extend([flag, value])

                cmd_str = " ".join(
                    f'"{a}"' if " " in a else a for a in wf_args_list
                )
                console.print()
                console.rule(
                    f"[bold cyan]{rec.workflow.value} — {WORKFLOW_NAMES[rec.workflow]}[/bold cyan]"
                )
                console.print(f"[dim]Command: python main.py {cmd_str}[/dim]\n")
                ctx.selected_pipeline.append(rec.workflow.value)

                with tracker.track_workflow(rec.workflow.value, cmd_str) as run:
                    result = run_cli(wf_args_list, timeout=900)
                    run.stdout_tail = result.stdout[-500:]
                    run.status = "success" if result.success else "failed"
                    run.error = result.stderr[-300:] if not result.success else ""
                    run.outputs = _j.dumps(list(set(
                        _r.findall(r"processed/[^\s\)\"]+", result.stdout)
                    )))

                status_color = "green" if result.success else "red"
                status_word = "COMPLETED" if result.success else "FAILED"
                console.print(
                    f"\n[{status_color}]{rec.workflow.value} {status_word} "
                    f"in {result.duration:.0f}s[/{status_color}]"
                )
                followon_results.append({
                    "workflow": rec.workflow.value,
                    "name": WORKFLOW_NAMES[rec.workflow],
                    "success": result.success,
                    "duration": result.duration,
                    "error": result.stderr[-300:] if not result.success else "",
                    "outputs": list(set(_r.findall(r"processed/[^\s\)\"]+", result.stdout))),
                })

                if not result.success:
                    console.print(f"[red]Error detail: {result.stderr[-400:]}[/red]")
                    console.print("[yellow]Skipping remaining workflows due to failure.[/yellow]")
                    break

    # ---------------------------------------------------------------------------
    # Step 6: Detailed Final Report
    # ---------------------------------------------------------------------------
    console.print()
    console.rule("[bold green]Pipeline Complete — Final Report[/bold green]")

    import time as _time
    total_dur = _time.time() - (_time.time() - tracker._start_time if hasattr(tracker, '_start_time') else 0)

    report_lines = [
        "=" * 68,
        "  ExcelMCP Agent — Pipeline Execution Report",
        "=" * 68,
        f"  Run ID        : {run_id}",
        f"  Input File    : {Path(ctx.input_file).name}",
        f"  Client        : {ctx.client}",
        f"  Mode          : AI Crew + Follow-on Execution",
        f"  Pipeline      : {' → '.join(ctx.selected_pipeline) if ctx.selected_pipeline else ctx.initial_workflow}",
        "",
    ]

    # WF1 signals section
    if wf1_signals:
        report_lines += [
            "  ── WF1 Analysis Signals ──────────────────────────────",
            f"  Power Queries : {'YES' if wf1_signals.power_query_found else 'no'}",
            f"  VBA Modules   : {'YES' if wf1_signals.vba_found else 'no'}",
            f"  SQL Request   : {'YES' if wf1_signals.sql_request_found else 'no'}",
            f"  Blueprint DB  : {'YES → ' + wf1_signals.blueprint_db_path if wf1_signals.blueprint_db_found else 'no'}",
            "",
        ]

    # Each workflow run
    report_lines.append("  ── Workflow Execution Details ────────────────────────")
    for i, wrun in enumerate(tracker.workflow_runs, 1):
        dur_str = f"{wrun.duration_seconds:.1f}s"
        status_str = "OK" if wrun.status == "success" else "FAILED"
        report_lines.append(f"  {i}. {wrun.workflow} ({WORKFLOW_NAMES.get(WF(wrun.workflow), wrun.workflow)}) — {status_str} [{dur_str}]")
        report_lines.append(f"     Command: python main.py {wrun.command}")
        try:
            outs = _j.loads(wrun.outputs)
            if outs:
                for o in outs:
                    report_lines.append(f"     Output : {o}")
        except Exception:
            pass
        if wrun.error:
            report_lines.append(f"     ERROR  : {wrun.error[:200]}")
        report_lines.append("")

    # Overall status
    fail_count = sum(1 for r in tracker.workflow_runs if r.status == "failed")
    overall = "ALL WORKFLOWS SUCCESSFUL" if fail_count == 0 else f"{fail_count} WORKFLOW(S) FAILED"
    overall_color = "green" if fail_count == 0 else "red"
    report_lines += [
        "  ── Overall Status ────────────────────────────────────",
        f"  Result        : {overall}",
        f"  Total Runs    : {len(tracker.workflow_runs)}",
        "=" * 68,
    ]

    report_text = "\n".join(report_lines)
    console.print(Panel(
        report_text,
        title="[bold green]Pipeline Execution Report[/bold green]",
        expand=False,
        border_style=overall_color,
    ))


if __name__ == "__main__":
    main()
