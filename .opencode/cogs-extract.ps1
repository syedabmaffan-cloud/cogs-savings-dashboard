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
  [switch]$NoBudget,
  [switch]$MrpLiveOnly
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
  for ($try = 1; $try -le 5; $try++) {
    try {
      $resp = Invoke-RestMethod -Method Post -Uri $McpUrl -Headers $McpHdr -Body $body -TimeoutSec 240
      if ($null -eq $resp.result) { throw 'MCP endpoint returned a non-JSON response (service unavailable / 502).' }
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

function Split-Chunk([object[]]$Arr, [int]$Size) {
  for ($i = 0; $i -lt $Arr.Count; $i += $Size) {
    $last = [Math]::Min($i + $Size - 1, $Arr.Count - 1)
    ,($Arr[$i..$last])
  }
}

<#
  MRP (Live): all inputs come from the iBOS MCP gateway.
    - requirement for C / C+1 / C+2  : latest active MRP run per BU (pro.tblMRPFromProductionSchedule*), numQty1..3
    - short till C / C+1 / C+2      : same formula as the existing Material Planning sheet, i.e.
                                      current stock - cumulative requirement (negative = shortage)
    - average monthly consumption   : actual Issue For Shop Floor quantities (wms.tblInventoryTransaction*),
                                      grouped once per BU x item x month, last 24 months, zeros after first issue
    - current stock today           : inventory ledger balance, i.e. SUM(wms.tblInventoryTransactionRow.numTransactionQuantity)
                                      for every active movement (receive, issue, transfer, adjustment) of the item in the BU.
                                      wms.tblItemPlantWarehouse.numCurrentStock is a stale cache for many BUs, so it is
                                      only used as a fallback (via the MRP snapshot) when an item has no ledger rows.
    - stock coverage (days)         : current stock / (average monthly consumption / 30); N/A when consumption is zero
#>
function Get-MrpLiveData {
  $empty = [pscustomobject]@{
    generatedAt = (Get-Date).ToString('s'); source = 'unavailable'; months = @(); sbus = @(); rows = @(); error = ''
  }
  try {
    $today    = Get-Date
    $todayS   = $today.ToString('yyyy-MM-dd')
    $winStart = ([datetime]::new($today.Year, $today.Month, 1)).AddMonths(-23).ToString('yyyy-MM-dd')
    $winEnd   = ([datetime]::new($today.Year, $today.Month, 1)).AddMonths(1).ToString('yyyy-MM-dd')

    # --- 1. current MRP run per business unit (latest dteFromDate still in force) ---
    Write-Host '[cogs] MRP live: reading current MRP runs from iBOS ...'
    $hdrRows = Invoke-McpQuery @"
SELECT intMRPFromProductionHeaderId AS HdrId, intMRPProductionScheduleCode AS Code,
       intBusinessUnitId AS BU, dteFromDate AS FromDate, dteToDate AS ToDate
FROM pro.tblMRPFromProductionScheduleHeader WITH (NOLOCK)
WHERE isActive = 1 AND dteFromDate <= '$todayS' AND dteToDate >= '$todayS'
ORDER BY intBusinessUnitId, dteFromDate DESC, intMRPFromProductionHeaderId DESC
"@
    $hdr = @{}; $hdrOrder = New-Object System.Collections.Generic.List[int]
    foreach ($c in $hdrRows) {
      if ($c.Count -lt 5) { continue }
      $bu = [int]$c[2]
      if ($hdr.ContainsKey($bu)) { continue }
      $code = "$($c[1])"
      # Use the SBU code from the MRP run code (MRP-AEL-... -> AEL) so tabs filter by the same SBUs as the rest of the dashboard.
      $sbu = if ($code -match '^MRP-([A-Z]+)-') { $Matches[1] } else { "BU$bu" }
      $hdr[$bu] = [pscustomobject]@{ id = [int]$c[0]; code = $code; bu = $bu; sbu = $sbu; from = [datetime]::Parse("$($c[3])") }
      $hdrOrder.Add($bu)
    }
    if (-not $hdrOrder.Count) { throw 'No active MRP run found for the current period.' }

    $baseMonth = $hdr[$hdrOrder[0]].from
    $months = @(0..2 | ForEach-Object { $baseMonth.AddMonths($_).ToString('MMM-yy') })
    $sbus = @($hdrOrder | ForEach-Object { $hdr[$_].sbu } | Select-Object -Unique)

    # --- 2. requirement rows (one page set per MRP run) ---
    $mrp = @{}
    foreach ($bu in $hdrOrder) {
      $hid = $hdr[$bu].id
      $offset = 0
      while ($true) {
        $rows = Invoke-McpQuery @"
SELECT intMrpfromProductionScheduleRowId AS RowId, intItemId AS ItemId, strItemCode AS Code,
       strItemName AS Name, intUoMId AS UomId, numQty1 AS Q1, numQty2 AS Q2, numQty3 AS Q3,
       numStockQty AS MrpStock
FROM pro.tblMRPFromProductionScheduleRow WITH (NOLOCK)
WHERE intMrpfromProductionScheduleHeaderId = $hid
ORDER BY intMrpfromProductionScheduleRowId OFFSET $offset ROWS FETCH NEXT $PageSize ROWS ONLY
"@
        if ($rows.Count -eq 0) { break }
        foreach ($c in $rows) {
          if ($c.Count -lt 9) { continue }
          $itemId = [int]$c[1]
          $rq = @(0..2 | ForEach-Object { $v = ToNum $c[5 + $_]; if ($null -eq $v) { 0.0 } else { [double]$v } })
          $mrpStock = ToNum $c[8]
          $mrp["$bu|$itemId"] = [pscustomobject]@{ bu = $bu; sbu = $hdr[$bu].sbu; mrpCode = $hdr[$bu].code
            itemId = $itemId; code = "$($c[2])"; name = "$($c[3])"; uomId = "$($c[4])"; req = $rq; mrpStock = $mrpStock }
        }
        if ($rows.Count -lt $PageSize) { break }
        $offset += $PageSize
      }
    }
    Write-Host ("[cogs] MRP live: {0} requirement rows" -f $mrp.Count)
    if (-not $mrp.Count) { throw 'MRP runs contained no requirement rows.' }

    $ids = @($mrp.Keys | ForEach-Object { [int](($_ -split '\|')[1]) } | Sort-Object -Unique)
    $chunks = @(Split-Chunk $ids 300)

    # --- 3. item master (description + base UOM fallback) ---
    $nameMap = @{}; $itemUom = @{}
    foreach ($chunk in $chunks) {
      $list = $chunk -join ','
      $nr = Invoke-McpQuery "SELECT intItemId, strItemName, strBaseUomName FROM itm.tblItem WITH (NOLOCK) WHERE intItemId IN ($list)"
      foreach ($n in $nr) {
        if ($n.Count -lt 2) { continue }
        $nameMap[[int]$n[0]] = "$($n[1])"
        if ($n.Count -ge 3 -and $null -ne $n[2] -and "$($n[2])".Trim() -ne '') { $itemUom[[int]$n[0]] = "$($n[2])".Trim() }
      }
    }

    # MRP requirement rows carry intUoMId. Resolve it to the UOM text actually used by
    # the MRP run/stock transactions so every displayed quantity is in one consistent UOM.
    $uomIds = @($mrp.Values | ForEach-Object { "$($_.uomId)" } | Where-Object { $_ -ne '' } | Sort-Object -Unique)
    $mrpUom = @{}
    foreach ($chunk in $chunks) {
      $list = $chunk -join ','
      $uomList = $uomIds -join ','
      $ur = Invoke-McpQuery "SELECT intItemId, intConvertedUOM, strConvertedUomName FROM itm.tblItemUOMConversion WITH (NOLOCK) WHERE isActive = 1 AND intItemId IN ($list) AND intConvertedUOM IN ($uomList)"
      foreach ($u in $ur) {
        if ($u.Count -lt 3) { continue }
        $un = (("$($u[2])" -replace '[\r\n]+', ' ').Trim())
        if ($un -ne '') { $mrpUom["$([int]$u[0])|$([int]$u[1])"] = $un }
      }
    }

    # --- 4. current stock today: inventory ledger balance (all warehouses; negatives preserved) ---
    # wms.tblItemPlantWarehouse.numCurrentStock is a cached balance that stops updating for many
    # BUs (last action Sep/Oct 2026), so stock is rebuilt from every active ledger movement:
    # receipts +, issues -, transfers +/- and inventory adjustments.
    $stock = @{}
    foreach ($chunk in $chunks) {
      $list = $chunk -join ','
      $base = @"
SELECT h.intBusinessUnitId AS BU, r.intItemId AS ItemId, SUM(r.numTransactionQuantity) AS Stock,
       MAX(h.dteTransactionDate) AS Last
FROM wms.tblInventoryTransactionHeader h WITH (NOLOCK)
JOIN wms.tblInventoryTransactionRow r WITH (NOLOCK) ON r.intInventoryTransactionId = h.intInventoryTransactionId
WHERE h.intBusinessUnitId IN ($BuList) AND h.isActive = 1 AND r.isActive = 1
  AND r.intItemId IN ($list)
GROUP BY h.intBusinessUnitId, r.intItemId
"@
      $offset = 0
      while ($true) {
        $sql = "$base ORDER BY h.intBusinessUnitId, r.intItemId OFFSET $offset ROWS FETCH NEXT $PageSize ROWS ONLY"
        $sr = Invoke-McpQuery $sql
        if ($sr.Count -eq 0) { break }
        foreach ($s in $sr) {
          if ($s.Count -lt 3) { continue }
          $v = ToNum $s[2]; if ($null -eq $v) { $v = 0.0 }
          $stock["$([int]$s[0])|$([int]$s[1])"] = @{ v = [double]$v; last = "$($s[3])" }
        }
        if ($sr.Count -lt $PageSize) { break }
        $offset += $PageSize
      }
    }

    # --- 5. actual consumption: grouped once per BU x item x month (no double counting) ---
    $cons = @{}; $consUom = @{}
    foreach ($chunk in $chunks) {
      $list = $chunk -join ','
      $base = @"
SELECT h.intBusinessUnitId AS BU, r.intItemId AS ItemId, CONVERT(char(7), h.dteTransactionDate, 120) AS Pd,
       CAST(SUM(ABS(CAST(r.numTransactionQuantity AS decimal(28,4)))) AS decimal(28,2)) AS Qty,
       MAX(r.strUoMName) AS Uom
FROM wms.tblInventoryTransactionHeader h WITH (NOLOCK)
JOIN wms.tblInventoryTransactionRow r WITH (NOLOCK) ON r.intInventoryTransactionId = h.intInventoryTransactionId
WHERE h.intBusinessUnitId IN ($BuList) AND h.isActive = 1 AND r.isActive = 1
  AND h.TransactionGroupName = 'Issue Inventory' AND h.strTransactionTypeName LIKE 'Issue%Shop%'
  AND h.dteTransactionDate >= '$winStart' AND h.dteTransactionDate < '$winEnd'
  AND r.intItemId IN ($list)
GROUP BY h.intBusinessUnitId, r.intItemId, CONVERT(char(7), h.dteTransactionDate, 120)
"@
      $offset = 0
      while ($true) {
        $sql = "$base ORDER BY h.intBusinessUnitId, r.intItemId, CONVERT(char(7), h.dteTransactionDate, 120) OFFSET $offset ROWS FETCH NEXT $PageSize ROWS ONLY"
        $cr = Invoke-McpQuery $sql
        if ($cr.Count -eq 0) { break }
        foreach ($x in $cr) {
          if ($x.Count -lt 4) { continue }
          $k = "$([int]$x[0])|$([int]$x[1])"
          $q = ToNum $x[3]; if ($null -eq $q) { $q = 0.0 }
          $pd = "$($x[2])".Trim()
          if ($pd -notmatch '^\d{4}-\d{2}$') {
            Write-Host ("[cogs] MRP live bad period key='{0}' cells={1}" -f $pd, ($x -join '|'))
            continue
          }
          if (-not $cons.ContainsKey($k)) { $cons[$k] = @{} }
          $cons[$k][$pd] = [double]$q
          if ($x.Count -ge 5 -and "$($x[4])" -ne '') { $consUom[$k] = "$($x[4])" }
        }
        if ($cr.Count -lt $PageSize) { break }
        $offset += $PageSize
      }
    }

    $maxPd = ''
    foreach ($k in $cons.Keys) { foreach ($p in $cons[$k].Keys) { $pp = "$p".Trim(); if ($pp -gt $maxPd) { $maxPd = $pp } } }
    if (-not $maxPd) { $maxPd = ([datetime]::new($today.Year, $today.Month, 1)).ToString('yyyy-MM') }
    $anchorDt = [datetime]::ParseExact($maxPd.Trim(), 'yyyy-MM', [Globalization.CultureInfo]::InvariantCulture)
    $windowStart = $anchorDt.AddMonths(-23)

    # --- 6. assemble rows; short uses the established sheet formula: stock - cumulative requirement ---
    $out = New-Object System.Collections.Generic.List[object]
    $uomMismatch = 0
    foreach ($k in $mrp.Keys) {
      $m = $mrp[$k]
      $st = if ($stock.ContainsKey($k)) { $stock[$k] } else { $null }
      if ($st) { $stockV = [double]$st.v; $stockAsOf = "$($st.last)" }
      elseif ($null -ne $m.mrpStock) { $stockV = [double]$m.mrpStock; $stockAsOf = 'MRP snapshot' }
      else { $stockV = 0.0; $stockAsOf = '' }
      $cum = 0.0; $short = @()
      foreach ($q in $m.req) { $cum += [double]$q; $short += [Math]::Round(($stockV - $cum), 4) }

      $avg = $null; $daily = $null; $cover = $null; $nMonths = 0; $consFrom = ''
      if ($cons.ContainsKey($k) -and $cons[$k].Count -gt 0) {
        $first = "$((@($cons[$k].Keys | Sort-Object))[0])".Trim()
        $consFrom = $first
        try { $fd = [datetime]::ParseExact($first, 'yyyy-MM', [Globalization.CultureInfo]::InvariantCulture) }
        catch { Write-Host ("[cogs] MRP live skipped bad first='{0}' key='{1}'" -f $first, $k); continue }
        if ($fd -lt $windowStart) { $fd = $windowStart }
        $sum = 0.0
        for ($d = $fd; $d -le $anchorDt; $d = $d.AddMonths(1)) {
          $pm = '{0:yyyy-MM}' -f $d
          $q = 0.0
          if ($cons[$k].ContainsKey($pm)) { $q = [double]$cons[$k][$pm] }
          $sum += $q; $nMonths++
        }
        if ($nMonths -gt 0) {
          $avg = [Math]::Round(($sum / $nMonths), 4)
          $daily = [Math]::Round(($avg / 30.0), 6)
          if ($daily -ne 0) { $cover = [Math]::Round(($stockV / $daily), 2) }
        }
      }

      $uom = ''
      if ($mrpUom.ContainsKey("$($m.itemId)|$($m.uomId)")) { $uom = $mrpUom["$($m.itemId)|$($m.uomId)"] }
      elseif ($consUom.ContainsKey($k)) { $uom = $consUom[$k] }
      elseif ($itemUom.ContainsKey($m.itemId)) { $uom = $itemUom[$m.itemId] }
      if ($uom -and $consUom.ContainsKey($k) -and $consUom[$k] -ne $uom) {
        $uomMismatch++
        Write-Host ("[cogs] MRP live UOM mismatch {0} {1}: display='{2}' tx='{3}'" -f $m.sbu, $m.code, $uom, $consUom[$k])
      }
      $name = if ($nameMap.ContainsKey($m.itemId)) { $nameMap[$m.itemId] } else { $m.name }

      $out.Add([pscustomobject]@{
        sbu = $m.sbu; code = $m.code; name = $name; uom = $uom; mrpCode = $m.mrpCode
        req = @($m.req); short = @($short)
        avgCons = $avg; dailyCons = $daily; stock = $stockV; coverage = $cover
        consMonths = $nMonths; consFrom = $consFrom; consTo = $maxPd
        stockAsOf = $stockAsOf
      })
    }

    $sorted = @($out | Sort-Object sbu, code)
    Write-Host ("[cogs] MRP live: {0} rows, consumption through {1}, UOM warnings {2}" -f $sorted.Count, $maxPd, $uomMismatch)
    return [pscustomobject]@{
      generatedAt = (Get-Date).ToString('s')
      source = 'iBOSDDD MCP - MRP runs ' + ((@($hdrOrder | ForEach-Object { $hdr[$_].code })) -join ', ') +
               ' | issue-to-shop-floor actuals | inventory ledger stock'
      months = $months; sbus = $sbus; rows = $sorted; uomWarnings = $uomMismatch; error = ''
    }
  } catch {
    Write-Host ("[cogs] MRP live skipped: " + $_.Exception.Message)
    $empty.error = $_.Exception.Message
    return $empty
  }
}

if ($MrpLiveOnly) {
  $mrpLive = Get-MrpLiveData
  $basePath = Join-Path $ProjectRoot 'docs\data.json'
  if (-not (Test-Path -LiteralPath $basePath)) { throw "MrpLiveOnly needs an existing dataset at $basePath" }
  $data = Get-Content -LiteralPath $basePath -Raw -Encoding UTF8 | ConvertFrom-Json
  $data | Add-Member -Force -NotePropertyName mrpLive -NotePropertyValue $mrpLive
  $json = $data | ConvertTo-Json -Depth 12 -Compress
  [IO.File]::WriteAllText($OutJson, $json, (New-Object Text.UTF8Encoding($false)))
  Write-Host ("[cogs] wrote {0} ({1:N1} KB) - MRP Live only, other sections preserved" -f $OutJson, ((Get-Item $OutJson).Length / 1KB))
  return
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

function Invoke-SheetsGet([string]$Url, [hashtable]$Headers) {
  for ($try = 1; $try -le 10; $try++) {
    try { return Invoke-RestMethod -Uri $Url -Headers $Headers -TimeoutSec 90 }
    catch { if ($try -eq 10) { throw }; Start-Sleep -Seconds 2 }
  }
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
  AND h.strTransactionTypeName LIKE 'Issue%Shop%'
  AND h.isActive = 1 AND r.isActive = 1
  AND h.dteTransactionDate >= '2025-07-01'
  AND ( i.strItemCategoryName LIKE 'Direct Raw Material%'
        OR i.strItemTypeName = 'Packaging Materials'
        OR i.strItemName LIKE '%Cement Bag%' )
GROUP BY h.intBusinessUnitId, r.intItemId, i.strItemCode, i.strItemName, i.strItemTypeName,
         i.strItemCategoryName, CONVERT(char(7), h.dteTransactionDate, 120)
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
                               kind = "$($c[3])"; cat = "$($c[4])"; uom = "$($c[5])"; months = @{}; stockRate = 0.0 }
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
$rateMap = @{}
for ($i = 0; $i -lt $ids.Count; $i += 300) {
  $last  = [Math]::Min($i + 299, $ids.Count - 1)
  $chunk = ($ids[$i..$last]) -join ','
  $nrows = Invoke-McpQuery "SELECT intItemId, strItemName, numAverageRate FROM itm.tblItem WITH (NOLOCK) WHERE intItemId IN ($chunk)"
  foreach ($n in $nrows) {
    if ($n.Count -ge 2) { $nameMap[[int]$n[0]] = "$($n[1])" }
    if ($n.Count -ge 3) { $rv = ToNum $n[2]; if ($null -ne $rv) { $rateMap[[int]$n[0]] = [double]$rv } }
  }
}
foreach ($it in $items) {
  if ($nameMap.ContainsKey($it.itemId)) { $it.name = $nameMap[$it.itemId] }
  if ($rateMap.ContainsKey($it.itemId)) { $it.stockRate = [double]$rateMap[$it.itemId] }
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

# ================================================================ 2b. BUDGETED COGS + ACTUAL SALES
Write-Host '[cogs] reading budgeted COGS (GL 4810001) and actual sales from MCP ...'
$budgetCogs = New-Object System.Collections.Generic.List[object]
$salesList  = New-Object System.Collections.Generic.List[object]

$budRows = Invoke-McpQuery @"
SELECT intBusinessUnitId AS BU, intYearId AS Yr, intMonthId AS Mn, SUM(numAmount) AS Amt
FROM bgt.tblBudgetIncomeExpenseRow WITH (NOLOCK)
WHERE intBusinessUnitId IN ($BuList) AND intGeneralLedgerId = 103 AND isActive = 1 AND intYearId IN (2025,2026,2027)
GROUP BY intBusinessUnitId, intYearId, intMonthId
ORDER BY intBusinessUnitId, intYearId, intMonthId
"@
foreach ($r in $budRows) {
  if ($r.Count -lt 4) { continue }
  $v = ToNum $r[3]; if ($null -eq $v) { continue }
  $mm = "{0}-{1:00}" -f [int]$r[1], [int]$r[2]
  $budgetCogs.Add([pscustomobject]@{ bu = [int]$r[0]; m = $mm; v = [math]::Abs([double]$v) })
}

$salesRows = Invoke-McpQuery @"
SELECT intBusinessUnitId AS BU, YEAR(dteSalesOrderDate) AS Yr, MONTH(dteSalesOrderDate) AS Mn, SUM(numNetOrderValue) AS Amt
FROM oms.tblSalesOrderHeader WITH (NOLOCK)
WHERE intBusinessUnitId IN ($BuList) AND isActive = 1 AND isRejected = 0 AND dteSalesOrderDate >= '2025-07-01' AND dteSalesOrderDate < '2027-07-01'
GROUP BY intBusinessUnitId, YEAR(dteSalesOrderDate), MONTH(dteSalesOrderDate)
ORDER BY intBusinessUnitId, YEAR(dteSalesOrderDate), MONTH(dteSalesOrderDate)
"@
foreach ($r in $salesRows) {
  if ($r.Count -lt 4) { continue }
  $v = ToNum $r[3]; if ($null -eq $v) { continue }
  $mm = "{0}-{1:00}" -f [int]$r[1], [int]$r[2]
  $salesList.Add([pscustomobject]@{ bu = [int]$r[0]; m = $mm; v = [double]$v })
}
Write-Host ("[cogs] budgetCogs rows: {0} | sales rows: {1}" -f $budgetCogs.Count, $salesList.Count)

# ================================================================ 2d. MATERIAL PLANNING (MRP workbook)
$materialPlan = [pscustomobject]@{
  source = "MRP Run_Nov'26-Jan'27 (Google Sheet)"
  months = @('Nov-26','Dec-26','Jan-27')
  sbus   = @()
  rows   = @()
}
try {
  $mpSheetId = '1QwUvRKQpz46D962gXSLKFqEBgdVgdnV9SZyYBSJ7hoI'
  Write-Host '[cogs] reading MRP / material planning workbook ...'
  $mpToken = Get-GoogleToken
  $mpHdr = @{ Authorization = "Bearer $mpToken" }
  $metaUrl = 'https://sheets.googleapis.com/v4/spreadsheets/' + $mpSheetId + '?fields=sheets.properties(sheetId,title)'
  $metaMp = Invoke-SheetsGet $metaUrl $mpHdr
  $mpTabs = @($metaMp.sheets | Where-Object { $_.properties.title -like '*_MRP Nov*' })
  $mpRows = New-Object System.Collections.Generic.List[object]
  $mpSbus = New-Object System.Collections.Generic.List[string]
  foreach ($t in $mpTabs) {
    $title = $t.properties.title
    $sbu = ($title -split '_')[0]
    if (-not $mpSbus.Contains($sbu)) { $mpSbus.Add($sbu) }
    $esc = $title.Replace("'","''")
    $rng = [uri]::EscapeDataString("'${esc}'!A1:AC2000")
    $u = 'https://sheets.googleapis.com/v4/spreadsheets/' + $mpSheetId + '/values/' + $rng + '?majorDimension=ROWS'
    $vals = (Invoke-SheetsGet $u $mpHdr).values
    if (-not $vals -or $vals.Count -lt 2) { continue }
    $hdrN = @($vals[0] | ForEach-Object { ("$($_)".Trim().ToLower() -replace '[^a-z0-9]','') })
    $find = {
      param($opts)
      foreach ($o in $opts) {
        $k = ($o.ToLower() -replace '[^a-z0-9]','')
        for ($i = 0; $i -lt $hdrN.Count; $i++) { if ($hdrN[$i] -eq $k) { return $i } }
      }
      return -1
    }
    $iCode = & $find @('Item Code'); $iName = & $find @('Item Name'); $iUom = & $find @('UOM')
    $iReq = @((& $find @('Nov Req Qty')), (& $find @('Dec Req Qty')), (& $find @('Jan Req Qty')))
    $iSh  = @((& $find @('Nov Short (Excl. PR&PO)')), (& $find @('Dec Short (Excl. PR&PO)')), (& $find @('Jan Short (Excl. PR&PO)')))
    $iCov = & $find @('Coverage (Days)'); $iCovDate = & $find @('Coverage Date'); $i3m = & $find @('Total QTY (3M)')
    $iWh = & $find @('Warehouse Stock'); $iIt = & $find @('In Transit'); $iItA = & $find @('Actual in transit'); $iOpenPR = & $find @('Open PR')
    $iDead = & $find @('Dead Stock'); $iAvail = & $find @('Available Stock'); $iReqd = & $find @('Requirment','Requirement'); $iPR = & $find @('PR Code')
    $iRol = & $find @('Reorder Level'); $iRoq = & $find @('Reorder Quantity'); $iSs = & $find @('Safety Stock')
    $iSsQ = & $find @('Safety stock in qty (Inventory input)'); $iSsD = & $find @('Safety Stock in Days (Procurement input)')
    for ($r = 1; $r -lt $vals.Count; $r++) {
      $row = $vals[$r]
      $cell = { param($idx) if ($idx -ge 0 -and $idx -lt $row.Count) { return "$($row[$idx])" } else { return '' } }
      $code = "$(& $cell $iCode)".Trim()
      if (-not $code) { continue }
      $rawName = "$(& $cell $iName)".Trim()
      $parts = @($rawName -split "`r?`n" | ForEach-Object { "$_".Trim() } | Where-Object { $_ -ne '' })
      $src = ''
      if ($parts.Count -gt 1 -and $parts[-1] -match '^(?i)(local|import)$') { $src = $parts[-1]; $parts = @($parts[0..($parts.Count-2)]) }
      $covRaw = "$(& $cell $iCov)"
      $mpRows.Add([pscustomobject]@{
        sbu = $sbu; code = $code; name = ($parts -join ' '); source = $src; uom = "$(& $cell $iUom)".Trim()
        req   = @((ToNum (& $cell $iReq[0])), (ToNum (& $cell $iReq[1])), (ToNum (& $cell $iReq[2])))
        short = @((ToNum (& $cell $iSh[0])), (ToNum (& $cell $iSh[1])), (ToNum (& $cell $iSh[2])))
        covDays = (ToNum ($covRaw -replace '[^0-9\.\-]','')); covDate = "$(& $cell $iCovDate)".Trim(); total3m = (ToNum (& $cell $i3m))
        whStock = (ToNum (& $cell $iWh)); inTransit = (ToNum (& $cell $iIt)); actualInTransit = (ToNum (& $cell $iItA)); openPR = (ToNum (& $cell $iOpenPR))
        deadStock = (ToNum (& $cell $iDead)); available = (ToNum (& $cell $iAvail)); requirement = (ToNum (& $cell $iReqd)); prCode = "$(& $cell $iPR)".Trim()
        reorderLevel = (ToNum (& $cell $iRol)); reorderQty = (ToNum (& $cell $iRoq)); safetyStock = (ToNum (& $cell $iSs))
        ssQty = (ToNum (& $cell $iSsQ)); ssDays = (ToNum (& $cell $iSsD))
      })
    }
  }
  $materialPlan = [pscustomobject]@{
    source = "MRP Run_Nov'26-Jan'27 (Google Sheet)"
    months = @('Nov-26','Dec-26','Jan-27')
    sbus   = $mpSbus.ToArray()
    rows   = $mpRows.ToArray()
  }
  Write-Host ("[cogs] material planning: {0} rows across {1} SBUs" -f $mpRows.Count, $mpSbus.Count)
} catch {
  Write-Host ("[cogs] material planning skipped: " + $_.Exception.Message)
}

# ================================================================ 2e. MRP (LIVE)
$mrpLive = Get-MrpLiveData

# ================================================================ 3. WRITE
$data = [pscustomobject]@{
  generatedAt = (Get-Date).ToString('s')
  asOf        = $maxMonth
  source      = 'iBOSDDD (MCP) + ABP Google Sheet'
  sbus        = $Sbus
  items       = $items
  budget      = $budget
  budgetCogs  = $budgetCogs
  sales       = $salesList
  materialPlanning = $materialPlan
  mrpLive     = $mrpLive
}
$json = $data | ConvertTo-Json -Depth 12 -Compress
[IO.File]::WriteAllText($OutJson, $json, (New-Object Text.UTF8Encoding($false)))
Write-Host ("[cogs] wrote {0} ({1:N1} KB)" -f $OutJson, ((Get-Item $OutJson).Length / 1KB))
