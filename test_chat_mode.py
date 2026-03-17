"""
test_chat_mode.py
-----------------
Non-interactive end-to-end test of the Chat Mode feature.

Simulates a user saying:
  "I need a WIP (Work-in-Progress) report with a project summary sheet
   and a cost detail sheet, filtered to open jobs."

Steps:
  1. Template scan
  2. LLM config detection
  3. Full multi-turn conversation (scripted)
  4. YAML generation + validation
  5. Save the generated config (dry-run — does NOT call WF2)

Run from the excel-mcp-agent directory:
  python test_chat_mode.py
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

from dotenv import load_dotenv
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

load_dotenv(Path(__file__).parent / ".env")
sys.path.insert(0, str(Path(__file__).parent))

console = Console()

# ─── Step 1: Template scan ────────────────────────────────────────────────────

console.rule("[bold cyan]Step 1 — Template Scan[/bold cyan]")
from tools.template_scanner import scan_templates_raw, format_templates_for_prompt

templates = scan_templates_raw()
console.print(f"Found [bold]{len(templates)}[/bold] template(s):\n")
t_table = Table(show_header=True, box=None, padding=(0, 2))
t_table.add_column("Name", style="cyan")
t_table.add_column("Description", style="dim")
t_table.add_column("Sheets", style="dim")
for t in templates:
    t_table.add_row(
        t["name"],
        t["description"] or "—",
        ", ".join(t["sheet_names"]) or "—",
    )
console.print(t_table)
templates_text = format_templates_for_prompt(templates)

# ─── Step 2: LLM config ───────────────────────────────────────────────────────

console.rule("[bold cyan]Step 2 — LLM Config[/bold cyan]")
from agents.llm_config import get_litellm_config

llm_cfg = get_litellm_config()
console.print(f"Provider model : [green]{llm_cfg['model']}[/green]")
console.print(f"Base URL       : [dim]{llm_cfg.get('base_url', 'default')}[/dim]")
console.print(f"API key        : [dim]{str(llm_cfg.get('api_key',''))[:12]}...[/dim]")

# ─── Step 3: Scripted conversation ────────────────────────────────────────────

console.rule("[bold cyan]Step 3 — Scripted Conversation[/bold cyan]")
import litellm

system_prompt = (
    "You are an ExcelMCP workbook configuration expert. "
    "Your job is to understand the user's workbook requirements through conversation, "
    "then generate a valid WF2 YAML configuration.\n\n"
    "YAML SCHEMA (required):\n"
    "  metadata:\n"
    "    name: WorkbookName\n"
    "    title: Full Title\n"
    "    description: '...'\n"
    "  output:\n"
    "    path: 'generated_workbooks/<Name>/<Name>.xlsx'\n"
    "    format: xlsx\n"
    "    template_mode: true\n"
    "    overwrite: true\n"
    "  connection:\n"
    "    server: 'placeholder'\n"
    "    database: 'placeholder'\n"
    "    auth_type: sql\n"
    "    username: 'placeholder'\n"
    "    password: 'placeholder'\n"
    "  sheets:\n"
    "    - name: SheetName\n"
    "      type: data         # data | cover | summary | dashboard\n"
    "      purpose: '...'\n"
    "      sql: |\n"
    "        SELECT ... FROM ...\n"
    "  formatting:\n"
    "    semantic_colors: true\n"
    "    auto_filter: true\n"
    "    freeze_panes: B2\n\n"
    "RULES:\n"
    "- Always set template_mode: true (offline mode, no live DB needed)\n"
    "- connection values must be the literal string 'placeholder' — do NOT use ${VAR} syntax\n"
    "- Each data sheet MUST have a sql field\n"
    "- Sheet names MUST NOT contain special characters: no &, /, \\, ?, *, [, ]\n"
    "  Use underscores or full words instead (e.g. 'Profit_Loss' not 'P&L', 'Cash_Flow' not 'C/F')\n"
    "- Do NOT use ORDER BY at top level of SQL — use sort_by instead\n"
    "- Viewpoint tables: JCJM (job cost master), JCCD (cost details), GLBL (GL), APVM (vendors)\n\n"
    f"AVAILABLE TEMPLATES (for reference):\n{templates_text}\n\n"
    "BEHAVIOR:\n"
    "- Ask clarifying questions: report name, sheets needed, data/filters, client\n"
    "- Keep responses concise\n"
    "- When the user says 'done' or 'generate', output ONLY the raw YAML — "
    "no markdown fences, no explanation, just the YAML starting with 'metadata:' or 'output:'"
)

# Scripted turns that simulate a real user session
scripted_turns = [
    "I need a WIP (Work-in-Progress) report for construction projects. "
    "It should show open jobs with budgeted vs actual costs, and a cost detail breakdown.",

    "Client is Delnor Construction. Report name is 'Delnor WIP Report'. "
    "Filter to jobs where status is open. Two sheets: Summary by job, then Cost Detail.",

    "Generate the YAML configuration now. "
    "Output ONLY the raw YAML — no markdown fences, no explanation text.",
]

messages: list[dict] = [{"role": "system", "content": system_prompt}]

for turn_num, user_text in enumerate(scripted_turns, 1):
    console.print(f"\n[bold cyan]Turn {turn_num}[/bold cyan]")
    console.print(f"[cyan]You:[/cyan] {user_text}\n")

    messages.append({"role": "user", "content": user_text})

    try:
        resp = litellm.completion(messages=messages, **llm_cfg)
        assistant_text = resp.choices[0].message.content or ""
        messages.append({"role": "assistant", "content": assistant_text})
        console.print(f"[magenta]Agent:[/magenta] {assistant_text}\n")
    except Exception as exc:
        console.print(f"[red]LLM error on turn {turn_num}: {exc}[/red]")
        sys.exit(1)

# ─── Step 4: Extract + validate YAML ─────────────────────────────────────────

console.rule("[bold cyan]Step 4 — YAML Extraction & Validation[/bold cyan]")

# The last assistant message should be the YAML
raw_yaml = messages[-1]["content"]

# Strip any markdown fences the LLM might have added anyway
raw_yaml = re.sub(r"^```[a-z]*\n?", "", raw_yaml, flags=re.MULTILINE)
raw_yaml = re.sub(r"\n?```$", "", raw_yaml, flags=re.MULTILINE)
raw_yaml = raw_yaml.strip()

try:
    import yaml as _yaml
    config_dict = _yaml.safe_load(raw_yaml)
    if not isinstance(config_dict, dict):
        raise ValueError("Parsed content is not a YAML mapping")
    console.print("[green]YAML is valid[/green]")
except Exception as exc:
    console.print(f"[red]YAML validation failed: {exc}[/red]")
    console.print(Panel(raw_yaml, title="Raw Output", border_style="red"))
    sys.exit(1)

console.print()
console.print(Panel(
    raw_yaml,
    title="[bold cyan]Generated WF2 Config[/bold cyan]",
    border_style="cyan",
    expand=False,
))

# ─── Step 5: Save (dry-run — no WF2 execution) ────────────────────────────────

console.rule("[bold cyan]Step 5 — Save Config (dry-run)[/bold cyan]")

import os
EXCEL_MCP_CWD = os.getenv("EXCEL_MCP_CWD", r"c:\_Internal_ViewPoint_Projects\Jon-ExcelMCP3")

wb_name = (
    config_dict.get("metadata", {}).get("name")
    or config_dict.get("output", {}).get("filename", "").replace(".xlsx", "")
    or "generated_workbook"
)
safe_name = re.sub(r"[^\w\-]", "_", wb_name).strip("_") or "generated_workbook"
config_dir = Path(EXCEL_MCP_CWD) / "generated_workbooks" / safe_name / "config"
config_dir.mkdir(parents=True, exist_ok=True)
yaml_path = config_dir / f"{safe_name}.yaml"
yaml_path.write_text(raw_yaml, encoding="utf-8")

console.print(f"[green]Config saved:[/green] {yaml_path}")
console.print()
console.print(Panel(
    "\n".join([
        f"  Workbook Name : {wb_name}",
        f"  Config Path   : {yaml_path}",
        f"  Sheets        : {', '.join(s.get('name','?') for s in config_dict.get('sheets', []))}",
        f"  Template Mode : {config_dict.get('output', {}).get('template_mode', '?')}",
        "",
        "  [dim]To run WF2 with this config:[/dim]",
        f"  [cyan]cd {EXCEL_MCP_CWD}[/cyan]",
        f"  [cyan]python main.py create --config \"{yaml_path}\"[/cyan]",
    ]),
    title="[bold green]Test PASSED — Chat Mode Working[/bold green]",
    border_style="green",
    expand=False,
))
