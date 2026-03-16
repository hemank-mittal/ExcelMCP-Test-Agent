"""
Output Analysis Agent
----------------------
Inspects ExcelMCP3 workflow output directories to detect signals
(SQL queries, Power Query connections, VBA, blueprint databases)
that drive the next-workflow decision.
"""

from crewai import Agent
from agents.llm_config import get_llm
from tools.output_parser import parse_wf1_outputs, parse_wf3_outputs


def create_output_analysis_agent(verbose: bool = True) -> Agent:
    return Agent(
        role="Output Analysis Agent",
        goal=(
            "Inspect workflow output directories to detect what was produced: "
            "SQL query requests, Power Query connections, VBA modules, and blueprint "
            "databases. Report findings clearly so the Planner can decide next steps."
        ),
        backstory=(
            "You are a data detective who specialises in reading ExcelMCP3 output "
            "structures. After a workflow runs, you scan the output directory and "
            "report exactly what was generated. You use the parse tools to extract "
            "structured signals from the output.\n\n"
            "Key output locations:\n"
            "  WF1: processed/{client}/{workbook_stem}/\n"
            "       - documentation/*.json  (VBA, Power Query, formulas)\n"
            "       - documentation/*.db    (blueprint database)\n"
            "       - sql_query_request/    (SQL export package, if present)\n"
            "  WF3: processed/{client}/{workbook_stem}/improved/\n"
            "       - <stem>_improved.xlsx  (refactored workbook)\n"
            "       - refactoring_report.md\n\n"
            "Always call parse_wf1_outputs after WF1 and parse_wf3_outputs after WF3."
        ),
        tools=[parse_wf1_outputs, parse_wf3_outputs],
        llm=get_llm(),
        verbose=verbose,
    )
