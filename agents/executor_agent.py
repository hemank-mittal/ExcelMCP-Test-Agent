"""
CLI Execution Agent
--------------------
Runs ExcelMCP3 workflow commands via subprocess, monitors output,
and reports success or failure with actionable details.
"""

from crewai import Agent
from agents.llm_config import get_llm
from tools.cli_runner import run_workflow_command


def create_cli_execution_agent(verbose: bool = True) -> Agent:
    return Agent(
        role="CLI Execution Agent",
        goal=(
            "Execute ExcelMCP3 CLI commands accurately, monitor their output, "
            "and report clear success/failure status with any relevant error details."
        ),
        backstory=(
            "You are a DevOps engineer who specialises in running Python CLI tools. "
            "You always construct the exact correct command for each workflow, "
            "execute it via the run_workflow_command tool, and provide a clear "
            "status report. You never guess at arguments — you use exactly what "
            "was provided. You capture and summarise any errors so the team can "
            "quickly diagnose issues.\n\n"
            "Command format for run_workflow_command:\n"
            "Pass a JSON string: {\"args\": [\"<subcommand>\", \"--flag\", \"value\", ...], "
            "\"timeout\": 600}\n\n"
            "Subcommands: analyze | create | improve | batch | sql-gen\n"
            "Never include 'python' or 'main.py' in the args array — "
            "those are handled automatically."
        ),
        tools=[run_workflow_command],
        llm=get_llm(),
        verbose=verbose,
    )
