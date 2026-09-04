<#
.SYNOPSIS
    Installs Caddy as a Windows Service using NSSM. Caddy reverse-proxies
    mindbots.mindware.net -> 127.0.0.1:8501 (the Streamlit app) and
    automatically obtains/renews its own SSL certificate via Let's Encrypt.

.PREREQUISITES
    - Ports 80 and 443 open to the internet on this server (Let's Encrypt
      needs to reach port 80 to verify domain ownership).
    - DNS for mindbots.mindware.net already pointing at this server's
      public address (ask Imad — see deploy\Caddyfile).
    - Caddy (https://caddyserver.com/download) — download caddy.exe and
      place it in this deploy\ folder, or anywhere on PATH.
    - NSSM (https://nssm.cc/download) — download nssm.exe and place it in
      this deploy\ folder, or anywhere on PATH.

.USAGE
    Run as Administrator from the repo root:
        .\deploy\install_caddy_service.ps1
#>

$ErrorActionPreference = "Stop"

$ServiceName = "Caddy"
$RepoRoot    = (Resolve-Path "$PSScriptRoot\..").Path
$CaddyExe    = Join-Path $PSScriptRoot "caddy.exe"
$Nssm        = Join-Path $PSScriptRoot "nssm.exe"
$Caddyfile   = Join-Path $PSScriptRoot "Caddyfile"

if (-not (Test-Path $CaddyExe)) {
    throw "caddy.exe not found at $CaddyExe. Download it from https://caddyserver.com/download and place it in deploy\."
}
if (-not (Test-Path $Nssm)) {
    throw "nssm.exe not found at $Nssm. Download it from https://nssm.cc/download and place it in deploy\."
}

if (Get-Service $ServiceName -ErrorAction SilentlyContinue) {
    Write-Host "Service '$ServiceName' already exists, removing it first..."
    & $Nssm stop $ServiceName
    & $Nssm remove $ServiceName confirm
}

& $Nssm install $ServiceName $CaddyExe "run --config `"$Caddyfile`" --adapter caddyfile"
& $Nssm set $ServiceName AppDirectory $PSScriptRoot
& $Nssm set $ServiceName AppStdout (Join-Path $PSScriptRoot "caddy-service.log")
& $Nssm set $ServiceName AppStderr (Join-Path $PSScriptRoot "caddy-service-error.log")
& $Nssm set $ServiceName Start SERVICE_AUTO_START
& $Nssm set $ServiceName AppRestartDelay 5000

& $Nssm start $ServiceName

Write-Host "Service '$ServiceName' installed and started. Once DNS resolves, https://mindbots.mindware.net should be live within a minute or two (Caddy needs that time to obtain its certificate)."
