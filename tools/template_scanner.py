"""
Template Scanner Tool
---------------------
Scans Jon-ExcelMCP3's config/templates/ directory for available WF2 YAML templates
and returns their metadata. Used by the chat mode to inform the LLM of what
pre-built templates exist as starting points.
"""

from __future__ import annotations

import json
import os
from pathlib import Path

from crewai.tools import tool

try:
    import yaml as _yaml
    _YAML_AVAILABLE = True
except ImportError:
    _YAML_AVAILABLE = False


TEMPLATES_DIR = str(
    Path(os.getenv("EXCEL_MCP_CWD", r"c:\_Internal_ViewPoint_Projects\Jon-ExcelMCP3"))
    / "config" / "templates"
)


def scan_templates_raw(templates_dir: str = TEMPLATES_DIR) -> list[dict]:
    """
    Read all *.yaml files from the templates directory and return a list of
    template summaries: {name, path, description, tags, sheet_names}.
    """
    root = Path(templates_dir)
    if not root.exists():
        return []

    results = []
    for yaml_file in sorted(root.glob("*.yaml")):
        entry: dict = {
            "name": yaml_file.stem,
            "path": str(yaml_file),
            "description": "",
            "tags": [],
            "sheet_names": [],
        }

        if _YAML_AVAILABLE:
            try:
                raw = yaml_file.read_text(encoding="utf-8", errors="replace")
                doc = _yaml.safe_load(raw)
                if isinstance(doc, dict):
                    meta = doc.get("metadata", {}) or {}
                    entry["description"] = (
                        meta.get("description", "")
                        or meta.get("name", "")
                        or doc.get("title", "")
                        or doc.get("name", "")
                        or ""
                    )
                    entry["tags"] = meta.get("tags", []) or []
                    sheets = doc.get("sheets", []) or []
                    entry["sheet_names"] = [
                        s.get("name", "") for s in sheets if isinstance(s, dict)
                    ]
            except Exception:
                pass
        else:
            # Fallback: read first line of file as description
            try:
                first_line = yaml_file.read_text(encoding="utf-8").split("\n")[0]
                entry["description"] = first_line.lstrip("# ").strip()
            except Exception:
                pass

        results.append(entry)

    return results


def format_templates_for_prompt(templates: list[dict]) -> str:
    """
    Format template list as a concise string for inclusion in an LLM prompt.
    """
    if not templates:
        return "No templates available."
    lines = []
    for t in templates:
        desc = t["description"] or t["name"]
        sheets = ", ".join(t["sheet_names"]) if t["sheet_names"] else "unknown"
        lines.append(f"- {t['name']}: {desc}  (sheets: {sheets})")
    return "\n".join(lines)


@tool("scan_templates")
def scan_templates(templates_dir: str = "") -> str:
    """
    Scan the ExcelMCP3 templates directory and return available WF2 YAML templates.

    Returns JSON list of templates with name, description, tags, and sheet names.
    Pass templates_dir="" to use the default location from EXCEL_MCP_CWD env var.
    """
    directory = templates_dir.strip() or TEMPLATES_DIR
    templates = scan_templates_raw(directory)
    return json.dumps(templates, indent=2)
