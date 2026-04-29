$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonExe = Join-Path $projectRoot ".venv\Scripts\python.exe"
$serverPath = Join-Path $projectRoot "server.py"

if (-not (Test-Path $pythonExe)) {
    throw "Missing .venv Python at $pythonExe. Run .\setup_venv.ps1 first."
}

Set-Location $projectRoot
& $pythonExe $serverPath
