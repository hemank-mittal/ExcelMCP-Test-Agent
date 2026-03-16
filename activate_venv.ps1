# Activate the excel-mcp-agent virtual environment
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
& "$scriptDir\venv\Scripts\Activate.ps1"
Write-Host "excel-mcp-agent venv activated." -ForegroundColor Green
