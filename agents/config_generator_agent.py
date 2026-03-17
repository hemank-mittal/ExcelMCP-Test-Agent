"""
Config Generator Agent
----------------------
CrewAI agent used for the single-shot YAML generation step in chat mode.
The agent knows the full WF2 YAML schema and uses the scan_templates tool
to reference available templates.
"""

from __future__ import annotations

from crewai import Agent

from agents.llm_config import get_llm
from tools.template_scanner import scan_templates


def create_config_generator_agent(verbose: bool = True) -> Agent:
    return Agent(
        role="ExcelMCP Workbook Configuration Expert",
        goal=(
            "Generate a valid, complete WF2 YAML configuration file from a user's "
            "natural-language description of the workbook they want to create."
        ),
        backstory=(
            "You are an expert in ExcelMCP3's workbook generation system. "
            "You know the complete WF2 YAML schema and can produce valid configs "
            "that pass all validation checks.\n\n"

            "WF2 YAML SCHEMA (required fields):\n"
            "  output:\n"
            "    path: 'generated_workbooks/<name>/<name>.xlsx'  # REQUIRED\n"
            "    format: xlsx\n"
            "    template_mode: true   # Always true — offline, no live DB\n"
            "    overwrite: true\n\n"
            "  connection:             # Required block (values are placeholders in template_mode)\n"
            "    server: '${SQL_SERVER}'\n"
            "    database: '${SQL_DATABASE}'\n"
            "    auth_type: sql\n"
            "    username: '${SQL_USERNAME}'\n"
            "    password: '${SQL_PASSWORD}'\n\n"
            "  sheets:                 # At least 1 sheet REQUIRED\n"
            "    - name: SheetName     # REQUIRED\n"
            "      type: data          # data | cover | summary | dashboard\n"
            "      sql: |              # REQUIRED if type=data\n"
            "        SELECT ... FROM ...\n"
            "      purpose: '...'\n\n"
            "  metadata:               # Optional but recommended\n"
            "    name: WorkbookName\n"
            "    title: Full Title\n"
            "    description: '...'\n\n"
            "  formatting:             # Optional\n"
            "    semantic_colors: true\n"
            "    auto_filter: true\n"
            "    freeze_panes: B2\n\n"

            "SHEET TYPES:\n"
            "  - data: requires sql field, generates Power Query connection\n"
            "  - cover: title/intro page, no sql\n"
            "  - summary: dashboard summary, no sql\n"
            "  - dashboard: KPI view, no sql\n\n"

            "SQL NOTES:\n"
            "  - Do not use ORDER BY at top level (use sort_by instead)\n"
            "  - Use @ParamName for parameters, list them in params[]\n"
            "  - Reference Viewpoint tables: JCJM, JCCD, JCCO, GLBL, GLAC, etc.\n\n"

            "Use scan_templates to see available template structures as reference. "
            "Always set template_mode: true. "
            "Output ONLY valid YAML — no markdown fences, no explanation text."
        ),
        tools=[scan_templates],
        llm=get_llm(),
        verbose=verbose,
    )
