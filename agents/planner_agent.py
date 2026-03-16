"""
Workflow Planner Agent
-----------------------
Understands the ExcelMCP workflow system and decides which workflow(s) to run
based on user intent and output analysis signals.

This agent has full knowledge of the decision matrix:
  WF1 → WF5 (SQL detected)
  WF1 → WF3 (Power Query / VBA detected)
  WF1 → WF2 (blueprint DB created)
  WF3 → WF5 (SQL in Power Query)
"""

from crewai import Agent
from agents.llm_config import get_llm


# Embed the decision logic directly in the agent's backstory so Claude
# can reason about it without needing a separate tool call.
_DECISION_KNOWLEDGE = """
ExcelMCP3 Workflow Decision Matrix:

WF1 (analyze) → Next workflows based on output:
  - SQL query request directory found → WF5 (sql-gen) is strongly recommended
  - Power Query connections found     → WF3 (improve) is recommended
  - VBA modules found                 → WF3 (improve) is recommended
  - Blueprint .db database created    → WF2 (create) is possible (rebuild the workbook)

WF3 (improve) → Next workflows:
  - SQL embedded in Power Query found → WF5 (sql-gen) with --auto-vgpt2

Standard pipelines:
  1. Full improvement:   WF1 → WF3
  2. SQL generation:     WF1 → WF5
  3. Full rebuild:       WF1 → WF2
  4. Complete pipeline:  WF1 → WF3 → WF5

WF5 phases (must run in order):
  1. request  — generates sql_query_request/ package for VGPT2
  2. receive  — validates VGPT2 SQL output from sql_query_response/
  3. validate — live DB validation (requires VPN)

CLI commands:
  WF1: python main.py analyze --workbook <path> --client <name> [--client-url <url>]
  WF2: python main.py create --analysis-db <path.db> --output <path.xlsx>
       OR: python main.py create --config <path.yaml>
  WF3: python main.py improve --analysis-dir <path> [--auto-vgpt2]
  WF4: python main.py batch --input-dir <path> [--mode analyze|improve] [--workers N]
  WF5: python main.py sql-gen --client <name> --workbook <stem> [--phase request|receive|validate] [--auto-vgpt2]

WF1 output location: processed/{client}/{workbook_stem}/
WF3 output location: processed/{client}/{workbook_stem}/improved/
WF5 request output:  processed/{client}/{workbook_stem}/sql_query_request/
"""


def create_workflow_planner_agent(verbose: bool = True) -> Agent:
    return Agent(
        role="Workflow Planner Agent",
        goal=(
            "Analyze workflow output signals and recommend the optimal next workflow(s) "
            "to run. Build intelligent pipelines that maximize value from the ExcelMCP system."
        ),
        backstory=(
            "You are an expert in the ExcelMCP3 workflow system with deep knowledge "
            "of how workflows chain together. You always follow the decision matrix "
            "precisely. Given analysis signals (SQL found, Power Query found, VBA found, "
            "blueprint DB created), you recommend the correct next steps with exact CLI "
            "arguments.\n\n"
            + _DECISION_KNOWLEDGE
        ),
        tools=[],
        llm=get_llm(),
        verbose=verbose,
    )
