<#
.SYNOPSIS
    Installs the AR Backlog Streamlit app as a Windows Service using NSSM,
    so it starts automatically on boot and doesn't depend on an RDP session
    staying open.

.PREREQUISITES
    - Python installed and the app's dependencies already set up
      (python -m venv .venv; .venv\Scripts\pip install -r requirements.txt)
    - NSSM (https://nssm.cc/download) — download nssm.exe and place it in
      this deploy\ folder, or anywhere on PATH.

.USAGE
    Run as Administrator from the repo root:
        .\deploy\install_windows_service.ps1
#>

$ErrorActionPreference = "Stop"

$ServiceName = "ARBacklogStreamlit"
$RepoRoot    = (Resolve-Path "$PSScriptRoot\..").Path
$PythonExe   = Join-Path $RepoRoot ".venv\Scripts\python.exe"
$Nssm        = Join-Path $PSScriptRoot "nssm.exe"

if (-not (Test-Path $PythonExe)) {
    throw "Python venv not found at $PythonExe. Create it first: python -m venv .venv; .venv\Scripts\pip install -r requirements.txt"
}
if (-not (Test-Path $Nssm)) {
    throw "nssm.exe not found at $Nssm. Download it from https://nssm.cc/download and place it in deploy\."
}

if (Get-Service $ServiceName -ErrorAction SilentlyContinue) {
    Write-Host "Service '$ServiceName' already exists, removing it first..."
    & $Nssm stop $ServiceName
    & $Nssm remove $ServiceName confirm
}

& $Nssm install $ServiceName $PythonExe "-m streamlit run app.py"
& $Nssm set $ServiceName AppDirectory $RepoRoot
& $Nssm set $ServiceName AppStdout (Join-Path $RepoRoot "deploy\streamlit-service.log")
& $Nssm set $ServiceName AppStderr (Join-Path $RepoRoot "deploy\streamlit-service-error.log")
& $Nssm set $ServiceName Start SERVICE_AUTO_START
& $Nssm set $ServiceName AppRestartDelay 5000

& $Nssm start $ServiceName

Write-Host "Service '$ServiceName' installed and started. Streamlit is listening on 127.0.0.1:8501."
