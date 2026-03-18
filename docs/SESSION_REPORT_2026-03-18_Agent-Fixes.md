# Agent Session Report — 2026-03-18
## excel-mcp-agent: 3 Improvements Shipped

---

## Quick Summary

| # | Improvement | File |
|---|-------------|------|
| 1 | LLM switched to Claude Opus + model indicator in chat header | `agents/llm_config.py`, `main_agent.py` |
| 2 | DB credentials auto-injected from `.env` into generated YAML | `main_agent.py` |
| 3 | Power Query connections now visible in Excel after WF2 build | `main_agent.py` |

---

## 1 — LLM Provider + Model Indicator

**Problem:** Default LLM was Grok. No visible indicator of which model was active in chat mode.

**Fix (two changes, one theme):**

**`agents/llm_config.py`** — Anthropic now checked first in auto-detect, default model bumped to Opus:
```python
_DEFAULTS = { "anthropic": "claude-opus-4-6", "grok": "grok-4-1-fast-non-reasoning", ... }

if not provider:
    if anthropic_key:  provider = "anthropic"   # ← was last, now first
    elif xai_key:      provider = "grok"
    elif openai_key:   provider = "openai"
```
Override via env: `AGENT_PROVIDER=anthropic|grok|openai`, `AGENT_MODEL=<model-id>`

**`main_agent.py`** — Model + provider shown at top of every chat session:
```
  Reasoning with  claude-opus-4-6  (Anthropic)
```
Color-coded: Anthropic = orange · Grok = white · OpenAI = green

---

## 2 — DB Credential Auto-Injection from `.env`

**Problem:** LLM always generates YAML with `placeholder` credentials (`template_mode: true`). Users had to manually edit the YAML to add real DB credentials before WF2 could refresh live data.

**Fix:** After YAML is generated and parsed, agent reads `.env` for SQL credentials and offers to swap them in — with a confirm prompt and masked password display.

```
Real DB credentials found in .env:
  server    SQLPROD01
  database  ViewpointDB
  username  svc_excel
  password  ************

Replace 'placeholder' with these credentials? [Y/n]:
```

**Env vars read:**

| `.env` Variable | → YAML field |
|-----------------|--------------|
| `SQL_SERVER` | `connection.server` |
| `SQL_DATABASE` | `connection.database` |
| `SQL_USERNAME` | `connection.username` |
| `SQL_PASSWORD` | `connection.password` *(masked in display)* |

If no vars are set → `No DB credentials found — using placeholders.` and WF2 runs in offline/template mode.

**Also fixed:** System prompt now explicitly tells the LLM:
> `connection.*` values must be the literal string `'placeholder'` — do NOT use `${VAR}` syntax (env vars are not resolved in template mode)

This prevents the LLM from generating broken `${SQL_SERVER}` references in the YAML.

---

## 3 — Power Query Connections Visible in Excel

**Problem:** Generated workbooks showed **0 queries** in Excel's Queries & Connections panel, even though M code was correctly embedded in the DataMashup binary.

**Root cause:** WF2 `query_only` mode writes M code to `xl/customXml/` (DataMashup binary) but never creates `xl/connections.xml`. Excel requires this XML registration to surface queries in the panel.

**What didn't work first:** Setting `template_query_load_mode: worksheet_table` — this triggers COM `ListObjects.Add(xlSrcExternal)`, which immediately validates the connection string. Placeholder credentials cause the validation to fail, marking the workbook as **"Repaired"** on next open.

**Fix — ZIP-level injection (`_inject_query_connections()` in `main_agent.py`):**

After every successful WF2 run, the function directly patches the XLSX ZIP:

| Step | What happens |
|------|-------------|
| 1 | Open XLSX as ZIP in memory |
| 2 | Build `xl/connections.xml` — one `<connection>` per query (standard params + data sheet queries) |
| 3 | Patch `[Content_Types].xml` — add `Override` for `/xl/connections.xml` |
| 4 | Patch `xl/_rels/workbook.xml.rels` — add `Relationship` of type `connections` |
| 5 | Write to temp file → `os.replace()` onto original (atomic) |

No COM automation = no credential validation = no "Repaired" state.

**Result:** `contract_revenue_by_pm.xlsx` → 5 connections visible. `delnor_wip_fy2025.xlsx` → all queries visible, clean open.

---

## Related Pipeline Fixes (Jon-ExcelMCP3)

These agent fixes pair with three pipeline bugs fixed in the same session:

| Bug | File | Fix |
|-----|------|-----|
| `json_to_sqlite.py` inserted NULLs (wrong key case) | WF1 | Defensive `or`-chain key lookup |
| WF2 built empty workbook when DB had no data | WF2 | JSON fallback in `_load_blueprint()` |
| Agent recommended WF2 on empty DB | agent `output_parser.py` | `COUNT(*) > 0` check before setting flag |

→ Full details: `Jon-ExcelMCP3/docs/SESSION_REPORT_2026-03-18_Pipeline-Fixes.md`
