@echo off
:: ExcelMCP Agent Launcher
:: Double-click this file OR run: run.bat [options]
::
:: Examples:
::   run.bat                          <- interactive mode (AI crew, needs API key)
::   run.bat --no-crew                <- interactive mode, no API key needed
::   run.bat --no-crew --autonomous   <- auto-pipeline, no prompts
::
cd /d "%~dp0"
call venv\Scripts\activate.bat
python main_agent.py %*
