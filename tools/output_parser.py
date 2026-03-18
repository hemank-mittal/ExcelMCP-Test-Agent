"""
Output Parser Tool
------------------
Inspects ExcelMCP3 workflow output directories to detect signals that drive
next-workflow recommendations.

Key detections:
  WF1 outputs  →  SQL request dir, Power Query, VBA, blueprint DB
  WF3 outputs  →  improved workbook, SQL in Power Query
"""

from __future__ import annotations

import json
from pathlib import Path

from crewai.tools import tool
from workflow_logic import WF1Signals, WF3Signals


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _read_json_safe(path: Path) -> dict | list | None:
    try:
        return json.loads(path.read_text(encoding="utf-8", errors="replace"))
    except Exception:
        return None


def _detect_power_query(doc: dict | list | None) -> bool:
    if not doc or not isinstance(doc, dict):
        return False
    # documentation JSON has "power_queries" or "connections" key
    pq = doc.get("power_queries") or doc.get("connections") or doc.get("power_query")
    if isinstance(pq, list) and len(pq) > 0:
        return True
    if isinstance(pq, dict) and pq:
        return True
    return False


def _detect_vba(doc: dict | list | None) -> bool:
    if not doc or not isinstance(doc, dict):
        return False
    vba = doc.get("vba_modules") or doc.get("vba") or doc.get("modules")
    if isinstance(vba, list) and len(vba) > 0:
        return True
    if isinstance(vba, dict) and vba:
        return True
    return False


def _detect_sql_in_power_query(doc: dict | list | None) -> bool:
    """Check if any Power Query connection contains an embedded SQL SELECT statement."""
    if not doc or not isinstance(doc, dict):
        return False
    pq = doc.get("power_queries") or doc.get("connections") or doc.get("power_query") or []
    queries = pq if isinstance(pq, list) else list(pq.values()) if isinstance(pq, dict) else []
    for q in queries:
        m_code = ""
        if isinstance(q, dict):
            m_code = q.get("m_code", "") or q.get("formula", "") or q.get("query", "") or ""
        elif isinstance(q, str):
            m_code = q
        if "select " in m_code.lower() or "sql.database" in m_code.lower():
            return True
    return False


# ---------------------------------------------------------------------------
# WF1 output analysis
# ---------------------------------------------------------------------------

def parse_wf1_outputs_internal(
    analysis_dir: str, client: str = "", workbook_path: str = ""
) -> WF1Signals:
    """
    Inspect a WF1 output directory and return WF1Signals.

    analysis_dir should be the path to processed/{client}/{stem}/
    """
    root = Path(analysis_dir)
    signals = WF1Signals(
        analysis_dir=str(root),
        client=client,
        workbook_stem=root.name,
    )

    if not root.exists():
        return signals

    # 1. Check for sql_query_request/ directory
    sql_req = root / "sql_query_request"
    if sql_req.exists() and any(sql_req.iterdir()):
        signals.sql_request_found = True

    # 2. Find documentation JSON — must be the *-Analysis.json file
    doc_dir = root / "documentation"
    doc_json: dict | None = None
    if doc_dir.exists():
        # Target the Analysis JSON specifically (contains power_queries + vba_modules)
        analysis_files = list(doc_dir.glob("*-Analysis.json"))
        if not analysis_files:
            # Fallback: any JSON that isn't a known non-analysis file
            _skip = {"complexity_assessment.json"}
            analysis_files = [
                f for f in doc_dir.glob("*.json")
                if f.name not in _skip
                and "validation" not in f.name.lower()
                and "assumptions" not in f.name.lower()
                and "profile" not in f.name.lower()
                and "summary" not in f.name.lower()
            ]

        if analysis_files:
            doc_json = _read_json_safe(analysis_files[0])
            if isinstance(doc_json, dict):
                # Primary: use summary counts (fast and reliable)
                summary = doc_json.get("summary", {})
                if isinstance(summary, dict):
                    pq_count = summary.get("Power Queries Found", 0)
                    vba_count = summary.get("VBA Modules Found", 0)
                    if isinstance(pq_count, (int, float)) and pq_count > 0:
                        signals.power_query_found = True
                    if isinstance(vba_count, (int, float)) and vba_count > 0:
                        signals.vba_found = True

                # Fallback: check arrays directly
                if not signals.power_query_found:
                    signals.power_query_found = _detect_power_query(doc_json)
                if not signals.vba_found:
                    signals.vba_found = _detect_vba(doc_json)

        # Check for blueprint DB — only mark as found if the DB actually has data
        db_files = list(doc_dir.glob("*.db"))
        if db_files:
            import sqlite3 as _sqlite3
            _db_has_data = False
            try:
                _con = _sqlite3.connect(str(db_files[0]))
                _cur = _con.cursor()
                _cur.execute("SELECT COUNT(*) FROM workbooks")
                _db_has_data = (_cur.fetchone()[0] or 0) > 0
                _con.close()
            except Exception:
                pass
            if _db_has_data:
                signals.blueprint_db_found = True
                signals.blueprint_db_path = str(db_files[0])

    return signals


def parse_wf3_outputs_internal(
    analysis_dir: str, client: str = ""
) -> WF3Signals:
    """Inspect a WF3 output directory and return WF3Signals."""
    root = Path(analysis_dir)
    signals = WF3Signals(
        analysis_dir=str(root),
        client=client,
        workbook_stem=root.name,
    )

    if not root.exists():
        return signals

    improved_dir = root / "improved"
    if improved_dir.exists():
        improved_files = list(improved_dir.glob("*_improved.*"))
        signals.improved_workbook_found = len(improved_files) > 0

    # Check if SQL was embedded in Power Query (from original doc JSON)
    doc_dir = root / "documentation"
    if doc_dir.exists():
        # Target *-Analysis.json (contains power_queries with M code)
        analysis_files = list(doc_dir.glob("*-Analysis.json"))
        if not analysis_files:
            _skip = {"complexity_assessment.json"}
            analysis_files = [
                f for f in doc_dir.glob("*.json")
                if f.name not in _skip
                and "validation" not in f.name.lower()
                and "assumptions" not in f.name.lower()
                and "profile" not in f.name.lower()
                and "summary" not in f.name.lower()
            ]
        if analysis_files:
            doc_json = _read_json_safe(analysis_files[0])
            signals.sql_in_power_query = _detect_sql_in_power_query(doc_json)

    return signals


# ---------------------------------------------------------------------------
# Agent @tool wrappers
# ---------------------------------------------------------------------------

@tool("parse_wf1_outputs")
def parse_wf1_outputs(analysis_dir: str, client: str, workbook_path: str) -> str:
    """
    Parse WF1 output directory to detect SQL requests, Power Query, VBA, and blueprint DB.

    Returns JSON with fields: sql_request_found, power_query_found, vba_found,
    blueprint_db_found, blueprint_db_path, analysis_dir, client, workbook_stem.
    """
    signals = parse_wf1_outputs_internal(analysis_dir, client, workbook_path)
    return json.dumps({
        "sql_request_found": signals.sql_request_found,
        "power_query_found": signals.power_query_found,
        "vba_found": signals.vba_found,
        "blueprint_db_found": signals.blueprint_db_found,
        "blueprint_db_path": signals.blueprint_db_path,
        "analysis_dir": signals.analysis_dir,
        "client": signals.client,
        "workbook_stem": signals.workbook_stem,
    }, indent=2)


@tool("parse_wf3_outputs")
def parse_wf3_outputs(analysis_dir: str, client: str) -> str:
    """
    Parse WF3 output directory to detect improved workbook and SQL-in-Power-Query signals.

    Returns JSON with fields: improved_workbook_found, sql_in_power_query,
    analysis_dir, client, workbook_stem.
    """
    signals = parse_wf3_outputs_internal(analysis_dir, client)
    return json.dumps({
        "improved_workbook_found": signals.improved_workbook_found,
        "sql_in_power_query": signals.sql_in_power_query,
        "analysis_dir": signals.analysis_dir,
        "client": signals.client,
        "workbook_stem": signals.workbook_stem,
    }, indent=2)
