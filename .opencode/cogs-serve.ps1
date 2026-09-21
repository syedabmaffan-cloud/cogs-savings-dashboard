<#
  cogs-serve.ps1
  --------------
  Publishes the COGS Savings Dashboard as a real web page on your network.

    .\.opencode\cogs-serve.ps1                      # serve on port 8080, no password
    .\.opencode\cogs-serve.ps1 -User affan -Pass secret   # require a login
    .\.opencode\cogs-serve.ps1 -Stop                # stop the server

  It copies the latest COGS_Savings_Dashboard.html to .\site\index.html and starts a tiny
  static web server (node). Colleagues on the same network open  http://<your-ip>:<port>/
  Nothing is exposed to the public internet.
#>
[CmdletBinding()]
param(
  [int]$Port = 8080,
  [string]$User = '',
  [string]$Pass = '',
  [switch]$Stop,
  [switch]$Open
)
$ErrorActionPreference = 'Stop'
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$PidFile = Join-Path $ProjectRoot '.opencode\logs\cogs-server.pid'

if ($Stop) {
  if (Test-Path -LiteralPath $PidFile) {
    $p = Get-Content -LiteralPath $PidFile -Raw
    try { Stop-Process -Id ([int]$p) -Force -ErrorAction Stop; Remove-Item -LiteralPath $PidFile -Force; Write-Host "[cogs] server (pid $p) stopped." }
    catch { Write-Host "[cogs] could not stop pid $p : $($_.Exception.Message)" }
  } else { Write-Host '[cogs] no pid file - server not running (or started elsewhere).' }
  return
}

# --- publish latest dashboard as the site index ---
$src = Join-Path $ProjectRoot 'COGS_Savings_Dashboard.html'
$site = Join-Path $ProjectRoot 'docs'
New-Item -ItemType Directory -Force -Path $site | Out-Null
if (-not (Test-Path -LiteralPath $src)) { throw "Dashboard not built yet. Run .opencode\cogs-build.ps1 first." }
Copy-Item -LiteralPath $src -Destination (Join-Path $site 'index.html') -Force
Write-Host ("[cogs] published {0} -> docs\index.html" -f (Split-Path -Leaf $src))

# --- node path ---
$node = (Get-Command node -ErrorAction SilentlyContinue).Source
if (-not $node) { throw 'node not found on PATH.' }

# --- start server (detached) ---
$env:COGS_PORT = "$Port"; $env:COGS_USER = $User; $env:COGS_PASS = $Pass; $env:COGS_ROOT = 'docs'
$srv = Join-Path $PSScriptRoot 'cogs-server.js'
$proc = Start-Process -FilePath $node -ArgumentList ('"{0}"' -f $srv) -WorkingDirectory $ProjectRoot -WindowStyle Hidden -PassThru
$proc.Id | Set-Content -LiteralPath $PidFile -Encoding ascii
Start-Sleep -Milliseconds 800
if ($proc.HasExited) { throw "Server exited immediately (port $Port in use?). Try -Port 8090." }

# --- report URLs ---
$urls = @()
try {
  $urls = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
           Where-Object { $_.IPAddress -ne '127.0.0.1' -and $_.AddressState -eq 'Preferred' -and $_.IPAddress -notlike '169.254.*' }).IPAddress
} catch {}
if (-not $urls) { $urls = @($env:COMPUTERNAME) }

Write-Host ''
Write-Host ('[cogs] COGS Savings Dashboard is live (pid {0})' -f $proc.Id)
Write-Host ('   Local:     http://localhost:{0}/' -f $Port)
foreach ($ip in $urls) { Write-Host ('   Network:   http://{0}:{1}/' -f $ip, $Port) }
if ($User) { Write-Host ('   Login:     {0} / (password supplied)' -f $User) } else { Write-Host '   Login:     none (anyone on the network can open it)' }
Write-Host ''
Write-Host '   Share a Network URL with colleagues on the same network / VPN.'
Write-Host ('   Stop with: powershell -File .opencode\cogs-serve.ps1 -Stop')
Write-Host '   If Windows Firewall prompts, allow Private networks.'
if ($Open) { Start-Process ('http://localhost:{0}/' -f $Port) }
