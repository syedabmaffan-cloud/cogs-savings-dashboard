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

  Endpoints cap each response (enterprise-api-gateway 500 rows, legacy 200), so extraction
  pages with OFFSET/FETCH. Nothing is ever written back: only SELECT statements are issued.
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

# Resolve the MCP endpoint. Priority:
#   1. environment (used by the scheduled GitHub Actions refresh): COGS_MCP_URL + COGS_MCP_KEY
#   2. opencode config: enterprise-api-gateway (current) -> JSON rows, 500/page
#   3. opencode config: AssetMcpServer (legacy) -> markdown rows, 200/page
$ep = $null
if ($env:COGS_MCP_URL -and $env:COGS_MCP_TOKEN) {
  # enterprise-api-gateway style: bearer token, JSON rows, 500/page
  $ep = [pscustomobject]@{ url=$env:COGS_MCP_URL
    headers=@{ 'Authorization'=("Bearer " + $env:COGS_MCP_TOKEN); 'Accept'='application/json, text/event-stream'; 'Content-Type'='application/json' }
    tool='execute_readonly_query'; arg='sql'; format='json'; page=500 }
} elseif ($env:COGS_MCP_URL -and $env:COGS_MCP_KEY) {
  # legacy style: X-API-Key, markdown rows, 200/page
  $ep = [pscustomobject]@{ url=$env:COGS_MCP_URL
    headers=@{ 'X-API-Key'=$env:COGS_MCP_KEY; 'Accept'='application/json, text/event-stream'; 'Content-Type'='application/json' }
    tool='ExecuteReadOnlyQueryAsync'; arg='sqlQuery'; format='md'; page=200 }
} elseif ($cfg -and ($cfg.mcp.PSObject.Properties.Name -contains 'enterprise-api-gateway')) {
  $g = $cfg.mcp.'enterprise-api-gateway'
  $ep = [pscustomobject]@{ url=$g.url
    headers=@{ 'Authorization'=$g.headers.Authorization; 'Accept'='application/json, text/event-stream'; 'Content-Type'='application/json' }
    tool='execute_readonly_query'; arg='sql'; format='json'; page=500 }
} elseif ($cfg -and ($cfg.mcp.PSObject.Properties.Name -contains 'AssetMcpServer')) {
  $g = $cfg.mcp.AssetMcpServer
  $ep = [pscustomobject]@{ url=$g.url
    headers=@{ 'X-API-Key'=$g.headers.'X-API-Key'; 'Accept'='application/json, text/event-stream'; 'Content-Type'='application/json' }
    tool='ExecuteReadOnlyQueryAsync'; arg='sqlQuery'; format='md'; page=200 }
}
if (-not $ep) { throw 'No MCP endpoint resolved (set COGS_MCP_URL + COGS_MCP_TOKEN, or provide opencode.jsonc with enterprise-api-gateway).' }
$McpUrl = $ep.url; $McpHdr = $ep.headers; $McpTool = $ep.tool; $McpArg = $ep.arg; $McpFormat = $ep.format
if (-not $PSBoundParameters.ContainsKey('PageSize')) { $PageSize = $ep.page }
Write-Host ("[cogs] MCP {0}  tool={1}  format={2}  page={3}" -f $McpUrl, $McpTool, $McpFormat, $PageSize)

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
  $qargs = @{ $McpArg = $Sql; limit = [Math]::Min($PageSize, $ep.page) }
  $body = @{ jsonrpc = '2.0'; id = 1; method = 'tools/call';
             params = @{ name = $McpTool; arguments = $qargs } } | ConvertTo-Json -Depth 8
  for ($try = 1; $try -le 3; $try++) {
    try {
      $resp = Invoke-RestMethod -Method Post -Uri $McpUrl -Headers $McpHdr -Body $body -TimeoutSec 240
      if ($resp.result.isError) { throw ("MCP error: " + $resp.result.content[0].text) }
      $text = $resp.result.content[0].text
      if ($McpFormat -eq 'json') {
        $obj = $text | ConvertFrom-Json
        $out = New-Object System.Collections.Generic.List[object]
        $props = $null
        foreach ($r in $obj.rows) {
          if (-not $props) { $props = $r.PSObject.Properties.Name }
          $cells = @(); foreach ($cn in $props) { $cells += "$($r.$cn)" }
          $out.Add($cells)
        }
        return ,$out
      }
      return ,(ConvertTo-MdRows $text)
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

$baseSql = @"
SELECT h.intBusinessUnitId AS BU, r.intItemId AS ItemId, i.strItemCode AS Code,
       i.strItemCategoryName AS Cat,
       CASE WHEN i.strItemTypeName='Packaging Materials' OR i.strItemName LIKE '%Cement Bag%'
            THEN 'PM' ELSE 'RM' END AS Kind,
       MAX(r.strUoMName) AS Uom,
       CONVERT(char(7), h.dteTransactionDate, 120) AS Pd,
       CAST(SUM(ABS(CAST(r.numTransactionQuantity AS decimal(28,4)))) AS decimal(28,2)) AS Qty,
       CAST(SUM(ABS(CAST(r.monTransactionValue  AS decimal(28,2)))) AS decimal(28,2)) AS Val
FROM wms.tblInventoryTransactionHeader h WITH (NOLOCK)
JOIN wms.tblInventoryTransactionRow r WITH (NOLOCK) ON r.intInventoryTransactionId = h.intInventoryTransactionId
JOIN itm.tblItem i WITH (NOLOCK) ON i.intItemId = r.intItemId
WHERE h.intBusinessUnitId IN ($BuList)
  AND h.TransactionGroupName = 'Issue Inventory'
  AND ( h.strTransactionTypeName = 'Issue For ShopFloor' OR h.strTransactionTypeName = 'Issue For Shop Floor' )
  AND h.isActive = 1 AND r.isActive = 1
  AND h.dteTransactionDate >= '2025-07-01'
  AND ( i.strItemCategoryName LIKE 'Direct Raw Material%'
        OR i.strItemTypeName = 'Packaging Materials'
        OR i.strItemName LIKE '%Cement Bag%' )
GROUP BY h.intBusinessUnitId, r.intItemId, i.strItemCode, i.strItemCategoryName, i.strItemTypeName, i.strItemName,
         CONVERT(char(7), h.dteTransactionDate, 120)
"@

$items    = New-Object System.Collections.Generic.List[object]
$byKey    = @{}   # 'bu|itemId' -> item object
$offset   = 0
$page     = 0
$maxMonth = ''
while ($true) {
  $sql = "$baseSql ORDER BY h.intBusinessUnitId, r.intItemId, CONVERT(char(7), h.dteTransactionDate, 120) OFFSET $offset ROWS FETCH NEXT $PageSize ROWS ONLY"
  $rows = Invoke-McpQuery $sql
  if ($rows.Count -eq 0) { break }
  foreach ($c in $rows) {
    if ($c.Count -lt 9) { continue }
    $bu = [int]$c[0]; $iid = [int]$c[1]; $key = "$bu|$iid"
    if (-not $byKey.ContainsKey($key)) {
      $it = [pscustomobject]@{ bu = $bu; itemId = $iid; code = "$($c[2])"; name = ''
                               kind = "$($c[4])"; cat = "$($c[3])"; uom = "$($c[5])"; months = @{} }
      $byKey[$key] = $it; $items.Add($it)
    }
    $mm = "$($c[6])"
    $q = ToNum $c[7]; $v = ToNum $c[8]
    if ($null -eq $q) { $q = 0 }; if ($null -eq $v) { $v = 0 }
    $byKey[$key].months[$mm] = @{ q = [double]$q; v = [double]$v }
    if ($mm -gt $maxMonth) { $maxMonth = $mm }
  }
  $page++
  Write-Host ("[cogs]   page {0}: {1} rows (items so far {2})" -f $page, $rows.Count, $items.Count)
  if ($rows.Count -lt $PageSize) { break }
  $offset += $PageSize
}

# Resolve item names separately. The gateway's SQL rewriter returns 502 for the (Name + GROUP BY +
# ORDER BY/OFFSET) combination, so names are fetched by item id in chunks instead.
$ids = @($items | ForEach-Object { $_.itemId } | Sort-Object -Unique)
$nameMap = @{}
for ($i = 0; $i -lt $ids.Count; $i += 300) {
  $last  = [Math]::Min($i + 299, $ids.Count - 1)
  $chunk = ($ids[$i..$last]) -join ','
  $nrows = Invoke-McpQuery "SELECT intItemId, strItemName FROM itm.tblItem WITH (NOLOCK) WHERE intItemId IN ($chunk)"
  foreach ($n in $nrows) { if ($n.Count -ge 2) { $nameMap[[int]$n[0]] = "$($n[1])" } }
}
foreach ($it in $items) { if ($nameMap.ContainsKey($it.itemId)) { $it.name = $nameMap[$it.itemId] } }
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
