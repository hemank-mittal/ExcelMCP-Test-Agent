"""
File Scanner Tool
-----------------
Scans a directory for Excel workbooks and returns a structured listing.
"""

from __future__ import annotations

import json
from pathlib import Path

from crewai.tools import tool

EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb", ".xls"}


@tool("scan_excel_files")
def scan_excel_files(directory: str) -> str:
    """
    Scan a directory for Excel files (.xlsx, .xlsm, .xlsb, .xls).
    Returns a JSON string with: directory, count, and a list of files
    (each with name, path, size_kb, extension).
    Skips temporary lock files starting with ~$.
    """
    root = Path(directory)
    if not root.exists():
        return json.dumps({"error": f"Directory not found: {directory}", "files": []})

    results = []
    for entry in sorted(root.iterdir()):
        if entry.is_file() and entry.suffix.lower() in EXCEL_EXTENSIONS:
            # Skip temp lock files
            if entry.name.startswith("~$"):
                continue
            results.append({
                "name": entry.name,
                "path": str(entry),
                "size_kb": round(entry.stat().st_size / 1024, 1),
                "extension": entry.suffix.lower(),
            })

    return json.dumps({
        "directory": str(root),
        "count": len(results),
        "files": results,
    }, indent=2)


def scan_excel_files_raw(directory: str) -> list[Path]:
    """Return raw Path objects for Excel files (used by interactive prompts)."""
    root = Path(directory)
    if not root.exists():
        return []
    return [
        e for e in sorted(root.iterdir())
        if e.is_file()
        and e.suffix.lower() in EXCEL_EXTENSIONS
        and not e.name.startswith("~$")
    ]
