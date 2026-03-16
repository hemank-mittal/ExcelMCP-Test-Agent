# ExcelMCP Agent Launcher
# Usage: .\run.ps1 [options]
#   .\run.ps1                          # interactive, AI crew
#   .\run.ps1 --no-crew                # interactive, no API key needed
#   .\run.ps1 --no-crew --autonomous   # fully automatic pipeline
param([Parameter(ValueFromRemainingArguments=$true)][string[]]$AgentArgs)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
& "$scriptDir\venv\Scripts\Activate.ps1"
python "$scriptDir\main_agent.py" @AgentArgs
