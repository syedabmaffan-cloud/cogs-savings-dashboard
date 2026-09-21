<#
  cogs-build.ps1
  --------------
  Builds the publishable dashboard from .opencode\logs\cogs-data.json:

     site\data.json                  <- the refreshable dataset (page loads this at runtime)
     site\index.html                 <- dashboard; loads ./data.json, falls back to embedded snapshot
     COGS_Savings_Dashboard.html     <- same page as a single offline-capable file

  Usage:
     powershell -File .opencode\cogs-build.ps1
     powershell -File .opencode\cogs-build.ps1 -DataFile <json> -SiteDir <dir>
#>
[CmdletBinding()]
param(
  [string]$DataFile,
  [string]$SiteDir
)
$ErrorActionPreference = 'Stop'

$ProjectRoot = Split-Path -Parent $PSScriptRoot
if (-not $DataFile) { $DataFile = Join-Path $ProjectRoot '.opencode\logs\cogs-data.json' }
if (-not $SiteDir)  { $SiteDir  = Join-Path $ProjectRoot 'docs' }
$Template = Join-Path $ProjectRoot '.opencode\cogs-dashboard-template.html'
$Standalone = Join-Path $ProjectRoot 'COGS_Savings_Dashboard.html'

foreach ($f in @($DataFile, $Template)) { if (-not (Test-Path -LiteralPath $f)) { throw "Missing file: $f" } }
New-Item -ItemType Directory -Force -Path $SiteDir | Out-Null

$json = [IO.File]::ReadAllText($DataFile, [Text.Encoding]::UTF8)

# 1) runtime dataset the page fetches (pretty enough for diffing in git)
try {
  $obj = $json | ConvertFrom-Json
  $pretty = $obj | ConvertTo-Json -Depth 12 -Compress
} catch { $pretty = $json }

# 2) embedded fallback must be safe inside a <script> block
$embed = $json.Replace('</', '<\/')
$tpl = [IO.File]::ReadAllText($Template, [Text.Encoding]::UTF8)
if ($tpl -notmatch '/\*__COGS_DATA__\*/') { throw 'Template placeholder /*__COGS_DATA__*/ not found.' }
$html = $tpl.Replace('/*__COGS_DATA__*/ null', $embed)

$utf8 = New-Object Text.UTF8Encoding($false)
[IO.File]::WriteAllText((Join-Path $SiteDir 'data.json'), $pretty, $utf8)
[IO.File]::WriteAllText((Join-Path $SiteDir 'index.html'), $html, $utf8)
[IO.File]::WriteAllText($Standalone, $html, $utf8)

Write-Host ("[cogs] data.json  : {0:N0} KB" -f ((Get-Item (Join-Path $SiteDir 'data.json')).Length / 1KB))
Write-Host ("[cogs] site built : {0}" -f (Join-Path $SiteDir 'index.html'))
Write-Host ("[cogs] standalone : {0}" -f $Standalone)
