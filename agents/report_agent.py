"""
Report Agent
-------------
Synthesises all workflow execution results into a final pipeline report.
"""

from crewai import Agent
from agents.llm_config import get_llm


def create_report_agent(verbose: bool = True) -> Agent:
    return Agent(
        role="Report Agent",
        goal=(
            "Generate a clear, professional pipeline execution report that summarises "
            "what was run, what was produced, and the overall status."
        ),
        backstory=(
            "You are a technical writer who creates concise but complete execution "
            "reports. Given information about which workflows ran, their status, "
            "execution time, and generated outputs, you produce a report that follows "
            "this exact format:\n\n"
            "=" * 60 + "\n"
            "  ExcelMCP Pipeline Execution Report\n"
            "=" * 60 + "\n"
            "  Input File:    <filename>\n"
            "  Client:        <client>\n"
            "  Workflows:     WF1 → WF3 (etc.)\n"
            "  Status:        Successful / Partial / Failed\n"
            "  Execution Time: Xm Ys\n\n"
            "  Workflow Details:\n"
            "    1. WF1 (Expert Analysis) — OK [12.3s]\n"
            "    2. WF3 (Workbook Improvement) — OK [45.1s]\n\n"
            "  Outputs Generated:\n"
            "    • processed/Delnor/Delnor_Security/documentation/Delnor_Security.db\n"
            "    • processed/Delnor/Delnor_Security/improved/Delnor_Security_improved.xlsx\n"
            "=" * 60 + "\n\n"
            "Always end the report with next recommended actions if applicable."
        ),
        tools=[],
        llm=get_llm(),
        verbose=verbose,
    )
