"""
Workflow Decision Logic Engine
-------------------------------
Encodes all rules for chaining ExcelMCP workflows based on analysis output.

Decision matrix
---------------
WF1 output               →  Recommended next workflows
-----------------------  -  --------------------------------
SQL query request found  →  WF5 (SQL Generation)
Power Query detected     →  WF3 (Improvement)
VBA code detected        →  WF3 (Improvement)
Blueprint DB created     →  WF2 (Workbook Creation)

WF3 output               →  Recommended next workflows
SQL in Power Query       →  WF5 (SQL Generation)
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import NamedTuple


# ---------------------------------------------------------------------------
# Workflow catalogue
# ---------------------------------------------------------------------------

class WF(str, Enum):
    WF1 = "WF1"
    WF2 = "WF2"
    WF3 = "WF3"
    WF4 = "WF4"
    WF5 = "WF5"


WORKFLOW_NAMES = {
    WF.WF1: "Expert Analysis",
    WF.WF2: "Workbook Creation",
    WF.WF3: "Workbook Improvement",
    WF.WF4: "Batch Processing",
    WF.WF5: "SQL Generation",
}

WORKFLOW_COMMANDS = {
    WF.WF1: "analyze",
    WF.WF2: "create",
    WF.WF3: "improve",
    WF.WF4: "batch",
    WF.WF5: "sql-gen",
}

# Required CLI arguments per workflow
WORKFLOW_REQUIRED_ARGS: dict[WF, list[str]] = {
    WF.WF1: ["--workbook", "--client"],
    WF.WF2: [],   # mutually exclusive: --config OR (--analysis-db + --output)
    WF.WF3: ["--analysis-dir"],
    WF.WF4: ["--input-dir"],
    WF.WF5: ["--client", "--workbook"],
}

# Human-readable prompts for each required argument
ARGUMENT_PROMPTS: dict[str, str] = {
    "--workbook":     "Path to the Excel workbook (e.g. inputs/Delnor_Security.xlsx)",
    "--client":       "Client name (e.g. Delnor)",
    "--client-url":   "Client website URL (optional, press Enter to skip)",
    "--analysis-dir": "Path to WF1 analysis output directory (e.g. processed/Delnor/Delnor_Security_Workbook_V11.1)",
    "--analysis-db":  "Path to WF1 SQLite blueprint database (.db file)",
    "--output":       "Output path for the generated workbook (e.g. generated_workbooks/Delnor_rebuilt.xlsx)",
    "--config":       "Path to YAML workbook configuration file",
    "--input-dir":    "Directory containing workbooks to process (searched recursively)",
    "--phase":        "WF5 phase: request / receive / validate / status (default: request)",
}


# ---------------------------------------------------------------------------
# Detection signals from WF output analysis
# ---------------------------------------------------------------------------

@dataclass
class WF1Signals:
    """Signals detected after WF1 (analyze) runs."""
    sql_request_found: bool = False      # sql_query_request/ directory present
    power_query_found: bool = False      # Power Query connections in documentation JSON
    vba_found: bool = False              # VBA modules in documentation JSON
    blueprint_db_found: bool = False     # .db blueprint database created
    analysis_dir: str = ""              # processed/{client}/{stem}/
    blueprint_db_path: str = ""         # documentation/{stem}.db
    client: str = ""
    workbook_stem: str = ""


@dataclass
class WF3Signals:
    """Signals detected after WF3 (improve) runs."""
    improved_workbook_found: bool = False
    sql_in_power_query: bool = False
    analysis_dir: str = ""
    client: str = ""
    workbook_stem: str = ""


class Recommendation(NamedTuple):
    workflow: WF
    reason: str
    suggested_args: dict[str, str]


# ---------------------------------------------------------------------------
# Decision engine
# ---------------------------------------------------------------------------

class WorkflowLogic:
    """
    Stateless engine that converts workflow output signals
    into an ordered list of next-workflow recommendations.
    """

    EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb", ".xls"}

    # --- Public API --------------------------------------------------------

    @staticmethod
    def recommendations_after_wf1(signals: WF1Signals) -> list[Recommendation]:
        recs: list[Recommendation] = []

        if signals.sql_request_found:
            recs.append(Recommendation(
                workflow=WF.WF5,
                reason="WF1 detected raw SQL / data export — WF5 can generate optimized Power Query SQL.",
                suggested_args={
                    "--client": signals.client,
                    "--workbook": signals.workbook_stem,
                    "--phase": "request",
                    "--auto-vgpt2": "",
                },
            ))

        if signals.power_query_found or signals.vba_found:
            what = []
            if signals.power_query_found:
                what.append("Power Query connections")
            if signals.vba_found:
                what.append("VBA modules")
            recs.append(Recommendation(
                workflow=WF.WF3,
                reason=f"WF1 found {' and '.join(what)} — WF3 can refactor them with AI.",
                suggested_args={
                    "--analysis-dir": signals.analysis_dir,
                },
            ))

        if signals.blueprint_db_found:
            recs.append(Recommendation(
                workflow=WF.WF2,
                reason="WF1 created a blueprint database — WF2 can rebuild the workbook from it.",
                suggested_args={
                    "--analysis-db": signals.blueprint_db_path,
                    "--output": f"generated_workbooks/{signals.workbook_stem}_rebuilt.xlsx",
                },
            ))

        return recs

    @staticmethod
    def recommendations_after_wf3(signals: WF3Signals) -> list[Recommendation]:
        recs: list[Recommendation] = []
        if signals.sql_in_power_query:
            recs.append(Recommendation(
                workflow=WF.WF5,
                reason="WF3 found SQL embedded in Power Query — WF5 can optimize those queries.",
                suggested_args={
                    "--client": signals.client,
                    "--workbook": signals.workbook_stem,
                    "--auto-vgpt2": "",
                },
            ))
        return recs

    @staticmethod
    def derive_wf1_analysis_dir(workbook_path: str, client: str) -> str:
        """
        Compute the canonical WF1 output directory for a given workbook + client.
        Absolute path: {EXCEL_MCP_CWD}/processed/{client}/{stem}
        """
        stem = Path(workbook_path).stem
        base = os.getenv(
            "EXCEL_MCP_CWD",
            r"c:\_Internal_ViewPoint_Projects\Jon-ExcelMCP3",
        )
        return str(Path(base) / "processed" / client / stem)

    @staticmethod
    def derive_blueprint_db_path(workbook_path: str, client: str) -> str:
        """Path to the WF1 SQLite blueprint database."""
        stem = Path(workbook_path).stem
        base = os.getenv(
            "EXCEL_MCP_CWD",
            r"c:\_Internal_ViewPoint_Projects\Jon-ExcelMCP3",
        )
        return str(Path(base) / "processed" / client / stem / "documentation" / f"{stem}.db")

    @staticmethod
    def describe_pipeline(workflows: list[WF]) -> str:
        """Return a human-readable pipeline description, e.g. 'WF1 → WF3 → WF5'."""
        return " → ".join(w.value for w in workflows)

    @staticmethod
    def build_cli_command(
        wf: WF,
        args: dict[str, str],
        python_exe: str,
        main_py: str,
    ) -> list[str]:
        """
        Build a subprocess command list for a given workflow + argument dict.

        Boolean flags (empty-string values) are added without a value.
        """
        cmd = [python_exe, main_py, WORKFLOW_COMMANDS[wf]]
        for flag, value in args.items():
            if value == "":
                # Boolean flag
                cmd.append(flag)
            else:
                cmd.extend([flag, value])
        return cmd
