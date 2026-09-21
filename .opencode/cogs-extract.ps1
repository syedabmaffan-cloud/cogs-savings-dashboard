<#
  cogs-extract.ps1
  ----------------
  Extracts the COGS Savings Dashboard dataset from the iBOS MCP server (read-only)
  plus the budget-rate reference from the ABP Google Sheet, and writes:

     .opencode\logs\cogs-data.json      (actuals + budget, ready for the dashboard)

  Actuals  : direct RM/PM "Issue For Shop Floor" lines from wms.tblInventoryTransaction*,
             aggregated to SBU x Item x calendar-month (qty + value), for FY2025-26 (LY)
             and FY2026-27 (current) - i.e. everything from 2025-07-01 onward.
  Budget   : only the "Consolidated Budget Rate all SBU" tab of
             https://docs.google.com/spreadsheets/d/17WltMox1pa-GxnbxNGfBVIFVeymy8n-4q7B7M03u9e4
             (monthly BDT budget rate per SBU/item, Jul-2026..Jun-2027).

  The MCP server caps every response at 200 rows, so extraction pages with OFFSET/FETCH.
  Nothing is ever written back to any database: only SELECT statements are issued.
#>
[CmdletBinding()]
param(
  [string]$SheetId = '17WltMox1pa-GxnbxNGfBVIFVeymy8n-4q7B7M03u9e4',
  [string]$BudgetTab = 'Consolidated Budget Rate all SBU',
  [int]$PageSize = 200,
  [switch]$NoBudget
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ---------------------------------------------------------------- paths / config
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$LogDir      = Join-Path $ProjectRoot '.opencode\logs'
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$OutJson     = Join-Path $LogDir 'cogs-data.json'

$cfgPath = 'C:\Users\kille\.config\opencode\opencode.jsonc'
$cfg = $null
if (Test-Path -LiteralPath $cfgPath) {
  $cfg = (Get-Content -LiteralPath $cfgPath -Raw) -replace '(?m)^\s*//.*$', '' | ConvertFrom-Json
}
# secrets resolve from environment first (used by the scheduled GitHub Actions refresh),
# then from the local opencode config.
$McpUrl = if ($env:COGS_MCP_URL) { $env:COGS_MCP_URL } elseif ($cfg) { $cfg.mcp.AssetMcpServer.url } else { $null }
$ApiKey = if ($env:COGS_MCP_KEY) { $env:COGS_MCP_KEY } elseif ($cfg) { $cfg.mcp.AssetMcpServer.headers.'X-API-Key' } else { $null }
if (-not $McpUrl -or -not $ApiKey) { throw 'MCP url / key not found (set COGS_MCP_URL and COGS_MCP_KEY, or provide opencode.jsonc).' }
$McpHdr  = @{ 'X-API-Key' = $ApiKey; 'Accept' = 'application/json, text/event-stream'; 'Content-Type' = 'application/json' }

# ---------------------------------------------------------------- SBU scope
# Requested SBU code -> iBOSDDD business unit id.
# NOTE: two BUs share the code 'AIL'. BU 224 = Akij Ispat Ltd. (has issue-to-shop-floor data);
#       BU 135 = Akij Insaf Ltd. (none). AIL therefore maps to 224.
$Sbus = @(
  [pscustomobject]@{ bu = 4;   code = 'ACCL';  name = 'Akij Cement Company Ltd.' }
  [pscustomobject]@{ bu = 224; code = 'AIL';   name = 'Akij Ispat Ltd.' }
  [pscustomobject]@{ bu = 175; code = 'ARMCL'; name = 'Akij Ready Mix Concrete Ltd.' }
  [pscustomobject]@{ bu = 220; code = 'ABSL';  name = 'Akij Building Solutions Ltd.' }
  [pscustomobject]@{ bu = 8;   code = 'APFIL'; name = 'Akij Poly Fibre Industries Ltd.' }
  [pscustomobject]@{ bu = 188; code = 'HRML';  name = 'Hashem Rice Mills Ltd.' }
  [pscustomobject]@{ bu = 144; code = 'AEL';   name = 'Akij Essentials Ltd.' }
  [pscustomobject]@{ bu = 189; code = 'FAL';   name = 'Fariq Agro Ltd.' }
  [pscustomobject]@{ bu = 232; code = 'AAFL';  name = 'Akij Agro Feed Ltd.' }
  [pscustomobject]@{ bu = 237; code = 'ALEL';  name = 'Akij Light Engineering Ltd.' }
)
$BuList = ($Sbus | ForEach-Object { $_.bu }) -join ','

# ---------------------------------------------------------------- helpers
function Invoke-McpQuery([string]$Sql) {
  $body = @{ jsonrpc = '2.0'; id = 1; method = 'tools/call';
             params = @{ name = 'ExecuteReadOnlyQueryAsync'; arguments = @{ sqlQuery = $Sql; limit = 200 } } } |
          ConvertTo-Json -Depth 8
  for ($try = 1; $try -le 3; $try++) {
    try {
      $resp = Invoke-RestMethod -Method Post -Uri $McpUrl -Headers $McpHdr -Body $body -TimeoutSec 240
      if ($resp.result.isError) { throw ("MCP error: " + $resp.result.content[0].text) }
      return $resp.result.content[0].text
    } catch {
      if ($try -eq 3) { throw }
      Start-Sleep -Seconds (2 * $try)
    }
  }
}

function ConvertTo-MdRows([string]$Text) {
  $rows = New-Object System.Collections.Generic.List[object]
  $seenHeader = $false
  foreach ($ln in ($Text -split "`r?`n")) {
    $t = $ln.Trim()
    if (-not $t.StartsWith('|')) { continue }
    if (($t -replace '[|\-:\s]', '') -eq '') { continue }   # separator row
    if (-not $seenHeader) { $seenHeader = $true; continue } # header row
    $cells = @($t.Trim('|') -split '\|' | ForEach-Object { $_.Trim() })
    $rows.Add($cells)
  }
  return $rows
}

function ToNum([object]$v) {
  if ($null -eq $v) { return $null }
  $s = "$v".Trim()
  if ($s -eq '' -or $s -eq '-') { return $null }
  $d = 0.0
  if ([double]::TryParse(($s -replace ',', ''), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
  return $null
}

function Get-GoogleToken {
  if ($env:COGS_GCLIENT_ID -and $env:COGS_GCLIENT_SECRET -and $env:COGS_GREFRESH) {
    $cid = $env:COGS_GCLIENT_ID; $csec = $env:COGS_GCLIENT_SECRET; $rtok = $env:COGS_GREFRESH
  } else {
    $tokPath = 'C:\Users\kille\.config\google-workspace-mcp\tokens.json'
    $envGws  = $cfg.mcp.GoogleWorkspaceMcpServer.environment
    $tok     = Get-Content -LiteralPath $tokPath -Raw | ConvertFrom-Json
    if (-not $tok.refresh_token) { throw 'No refresh_token in google-workspace-mcp tokens.json.' }
    $cid = $envGws.GOOGLE_CLIENT_ID; $csec = $envGws.GOOGLE_CLIENT_SECRET; $rtok = $tok.refresh_token
  }
  $body = @{ client_id = $cid; client_secret = $csec; refresh_token = $rtok; grant_type = 'refresh_token' }
  return (Invoke-RestMethod -Method Post -Uri 'https://oauth2.googleapis.com/token' -Body $body `
            -ContentType 'application/x-www-form-urlencoded').access_token
}

# month-label -> 'yyyy-MM' for the budget tab (Jul 2026 .. Jun 2027)
$MonMap = @{ 'Jan'='01'; 'Feb'='02'; 'Mar'='03'; 'Apr'='04'; 'May'='05'; 'Jun'='06';
             'Jul'='07'; 'Aug'='08'; 'Sep'='09'; 'Oct'='10'; 'Nov'='11'; 'Dec'='12' }
$SectionSbu = @{
  'AKIJ CEMENT'                      = 'ACCL'
  'AKIJ ISPAT'                       = 'AIL'
  'AKIJ Readymix'                    = 'ARMCL'
  'AKIJ Building Solution Ashphalt'  = 'ABSL'
  'AKIJ Polyfibre'                   = 'APFIL'
  'AKIJ Essential'                   = 'AEL'
}

# ================================================================ 1. ACTUALS
Write-Host '[cogs] extracting direct RM/PM issue-to-shop-floor actuals from MCP ...'

$innerBase = @"
SELECT t.BU, t.ItemId, t.Code, t.Name, t.Kind, t.Cat, MAX(t.Uom) AS Uom,
       STRING_AGG(CONCAT(t.Pd,'~',CAST(t.Qty AS varchar(40)),'~',CAST(t.Val AS varchar(40))),';')
         WITHIN GROUP (ORDER BY t.Pd) AS Months
FROM (
  SELECT h.intBusinessUnitId AS BU, r.intItemId AS ItemId, i.strItemCode AS Code,
         REPLACE(REPLACE(REPLACE(i.strItemName,'|','/'),CHAR(10),' '),'~','-') AS Name,
         CASE WHEN i.strItemTypeName='Packaging Materials' OR i.strItemName LIKE '%Cement Bag%'
              THEN 'PM' ELSE 'RM' END AS Kind,
         i.strItemCategoryName AS Cat, MAX(r.strUoMName) AS Uom,
         CONVERT(char(7), h.dteTransactionDate, 120) AS Pd,
         CAST(SUM(ABS(CAST(r.numTransactionQuantity AS decimal(28,4)))) AS decimal(28,2)) AS Qty,
         CAST(SUM(ABS(CAST(r.monTransactionValue  AS decimal(28,2)))) AS decimal(28,2)) AS Val
  FROM wms.tblInventoryTransactionHeader h WITH (NOLOCK)
  JOIN wms.tblInventoryTransactionRow r WITH (NOLOCK) ON r.intInventoryTransactionId = h.intInventoryTransactionId
  JOIN itm.tblItem i WITH (NOLOCK) ON i.intItemId = r.intItemId
  WHERE h.intBusinessUnitId IN ($BuList)
    AND h.TransactionGroupName = 'Issue Inventory'
    AND h.strTransactionTypeName IN ('Issue For ShopFloor','Issue For Shop Floor')
    AND h.isActive = 1 AND r.isActive = 1
    AND h.dteTransactionDate >= '2025-07-01'
    AND ( i.strItemCategoryName LIKE 'Direct Raw Material%'
          OR i.strItemTypeName = 'Packaging Materials'
          OR i.strItemName LIKE '%Cement Bag%' )
  GROUP BY h.intBusinessUnitId, r.intItemId, i.strItemCode, i.strItemName, i.strItemTypeName,
           i.strItemCategoryName, CONVERT(char(7), h.dteTransactionDate, 120)
) t
GROUP BY t.BU, t.ItemId, t.Code, t.Name, t.Kind, t.Cat
"@

$items    = New-Object System.Collections.Generic.List[object]
$offset   = 0
$page     = 0
$maxMonth = ''
while ($true) {
  $sql = "SELECT BU, ItemId, Code, Name, Kind, Cat, Uom, Months FROM ( $innerBase ) x ORDER BY BU, ItemId OFFSET $offset ROWS FETCH NEXT $PageSize ROWS ONLY"
  $text = Invoke-McpQuery $sql
  $rows = ConvertTo-MdRows $text
  if ($rows.Count -eq 0) { break }
  foreach ($c in $rows) {
    if ($c.Count -lt 8) { continue }
    $months = @{}
    foreach ($seg in ("$($c[7])" -split ';')) {
      if (-not $seg) { continue }
      $p = $seg -split '~'
      if ($p.Count -lt 3) { continue }
      $mm = $p[0]
      $q  = ToNum $p[1]
      $v  = ToNum $p[2]
      if ($q -eq $null -and $v -eq $null) { continue }
      if ($months.ContainsKey($mm)) { $months[$mm] = @{ q = $months[$mm].q + [double]$q; v = $months[$mm].v + [double]$v } }
      else { $months[$mm] = @{ q = [double]$q; v = [double]$v } }
      if ($mm -gt $maxMonth) { $maxMonth = $mm }
    }
    $items.Add([pscustomobject]@{
      bu = [int]$c[0]; itemId = [int]$c[1]; code = "$($c[2])"; name = "$($c[3])"
      kind = "$($c[4])"; cat = "$($c[5])"; uom = "$($c[6])"; months = $months
    })
  }
  $page++
  Write-Host ("[cogs]   page {0}: {1} rows (total {2})" -f $page, $rows.Count, $items.Count)
  if ($rows.Count -lt $PageSize) { break }
  $offset += $PageSize
}
Write-Host ("[cogs] actuals: {0} items, latest month {1}" -f $items.Count, $maxMonth)

# ================================================================ 2. BUDGET
$budget = New-Object System.Collections.Generic.List[object]
if (-not $NoBudget) {
  Write-Host '[cogs] reading budget rates from the ABP Google Sheet ...'
  $token = Get-GoogleToken
  $gsHdr = @{ Authorization = "Bearer $token" }
  $rng   = [uri]::EscapeDataString("${BudgetTab}!A1:O1006")
  $url   = 'https://sheets.googleapis.com/v4/spreadsheets/' + $SheetId + '/values/' + $rng + '?majorDimension=ROWS'
  $vals  = (Invoke-RestMethod -Uri $url -Headers $gsHdr -Method Get).values

  $monthCols = @{}   # column index -> yyyy-MM
  $hdr = $vals[0]
  for ($i = 0; $i -lt $hdr.Count; $i++) {
    $h = "$($hdr[$i])"
    if ($h -match '^([A-Za-z]{3})\s+(\d{4})$') {
      $mon = $MonMap[$Matches[1].Substring(0,1).ToUpper() + $Matches[1].Substring(1,2).ToLower()]
      if (-not $mon) { $mon = $MonMap[$Matches[1]] }
      if ($mon) { $monthCols[$i] = "$($Matches[2])-$mon" }
    }
  }
  $section = ''
  for ($r = 1; $r -lt $vals.Count; $r++) {
    $row = $vals[$r]
    $a = if ($row.Count -gt 0) { "$($row[0])".Trim() } else { '' }
    if ($a -ne '') { $section = $a }
    $code = if ($row.Count -gt 1) { "$($row[1])".Trim() } else { '' }
    $desc = if ($row.Count -gt 2) { "$($row[2])".Trim() } else { '' }
    if ($desc -eq '' -and $code -eq '') { continue }
    $sbu = if ($SectionSbu.ContainsKey($section)) { $SectionSbu[$section] } else { '' }
    $rates = @{}
    $hasRate = $false
    foreach ($ci in $monthCols.Keys) {
      if ($ci -ge $row.Count) { continue }
      $val = ToNum $row[$ci]
      if ($val -ne $null) { $rates[$monthCols[$ci]] = $val; $hasRate = $true }
    }
    if (-not $hasRate) { continue }
    $budget.Add([pscustomobject]@{ sbu = $sbu; section = $section; code = $code; desc = $desc; rates = $rates })
  }
  Write-Host ("[cogs] budget rows: {0}" -f $budget.Count)
}

# ================================================================ 3. WRITE
$data = [pscustomobject]@{
  generatedAt = (Get-Date).ToString('s')
  asOf        = $maxMonth
  source      = 'iBOSDDD (MCP) + ABP Google Sheet'
  sbus        = $Sbus
  items       = $items
  budget      = $budget
}
$json = $data | ConvertTo-Json -Depth 12 -Compress
[IO.File]::WriteAllText($OutJson, $json, (New-Object Text.UTF8Encoding($false)))
Write-Host ("[cogs] wrote {0} ({1:N1} KB)" -f $OutJson, ((Get-Item $OutJson).Length / 1KB))
