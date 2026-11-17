<#
  add-material-planning.ps1
  -------------------------
  Fetches the "MRP Run_Nov'26-Jan'27" Google Sheet (one tab per SBU) and merges a
  `materialPlanning` block into an existing dashboard data.json (preserving actuals,
  budget, budgetCogs and sales). Use this when the ERP/MCP gateway is unavailable but
  you still want to refresh the Material Planning tab.

  Usage:
    powershell -File .opencode\add-material-planning.ps1 -DataFile <path to data.json>
#>
param(
  [Parameter(Mandatory=$true)][string]$DataFile,
  [string]$SheetId = '1QwUvRKQpz46D962gXSLKFqEBgdVgdnV9SZyYBSJ7hoI'
)
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function ToNum([object]$v) {
  if ($null -eq $v) { return $null }
  $s = "$v".Trim(); if ($s -eq '' -or $s -eq '-') { return $null }
  $d = 0.0
  if ([double]::TryParse(($s -replace ',', ''), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
  return $null
}
function Invoke-SheetsGet([string]$Url, [hashtable]$Headers) {
  for ($try = 1; $try -le 12; $try++) {
    try { return Invoke-RestMethod -Uri $Url -Headers $Headers -TimeoutSec 90 }
    catch { if ($try -eq 12) { throw }; Start-Sleep -Seconds 2 }
  }
}

$cfg = (Get-Content -LiteralPath 'C:\Users\kille\.config\opencode\opencode.jsonc' -Raw) -replace '(?m)^\s*//.*$','' | ConvertFrom-Json
if ($env:COGS_GCLIENT_ID -and $env:COGS_GCLIENT_SECRET -and $env:COGS_GREFRESH) {
  $cid=$env:COGS_GCLIENT_ID; $csec=$env:COGS_GCLIENT_SECRET; $rtok=$env:COGS_GREFRESH
} else {
  $envGws = $cfg.mcp.GoogleWorkspaceMcpServer.environment
  $tok = Get-Content -LiteralPath 'C:\Users\kille\.config\google-workspace-mcp\tokens.json' -Raw | ConvertFrom-Json
  $cid=$envGws.GOOGLE_CLIENT_ID; $csec=$envGws.GOOGLE_CLIENT_SECRET; $rtok=$tok.refresh_token
}
$body=@{client_id=$cid;client_secret=$csec;refresh_token=$rtok;grant_type='refresh_token'}
$at=(Invoke-RestMethod -Method Post -Uri 'https://oauth2.googleapis.com/token' -Body $body -ContentType 'application/x-www-form-urlencoded').access_token
$h=@{ Authorization = "Bearer $at" }

$meta = Invoke-SheetsGet ('https://sheets.googleapis.com/v4/spreadsheets/' + $SheetId + '?fields=sheets.properties(sheetId,title)') $h
$tabs = @($meta.sheets | Where-Object { $_.properties.title -like '*_MRP Nov*' })
$rowsOut = New-Object System.Collections.Generic.List[object]
$sbus = New-Object System.Collections.Generic.List[string]
foreach ($t in $tabs) {
  $title = $t.properties.title
  $sbu = ($title -split '_')[0]
  if (-not $sbus.Contains($sbu)) { $sbus.Add($sbu) }
  $esc = $title.Replace("'","''")
  $rng = [uri]::EscapeDataString("'${esc}'!A1:AC2000")
  $u = 'https://sheets.googleapis.com/v4/spreadsheets/' + $SheetId + '/values/' + $rng + '?majorDimension=ROWS'
  $vals = (Invoke-SheetsGet $u $h).values
  if (-not $vals -or $vals.Count -lt 2) { continue }
  $hdrN = @($vals[0] | ForEach-Object { ("$($_)".Trim().ToLower() -replace '[^a-z0-9]','') })
  $find = {
    param($opts)
    foreach ($o in $opts) { $k=($o.ToLower() -replace '[^a-z0-9]',''); for($i=0;$i -lt $hdrN.Count;$i++){ if($hdrN[$i] -eq $k){ return $i } } }
    return -1
  }
  $iCode=& $find @('Item Code'); $iName=& $find @('Item Name'); $iUom=& $find @('UOM')
  $iReq=@((& $find @('Nov Req Qty')),(& $find @('Dec Req Qty')),(& $find @('Jan Req Qty')))
  $iSh=@((& $find @('Nov Short (Excl. PR&PO)')),(& $find @('Dec Short (Excl. PR&PO)')),(& $find @('Jan Short (Excl. PR&PO)')))
  $iCov=& $find @('Coverage (Days)'); $iCovDate=& $find @('Coverage Date'); $i3m=& $find @('Total QTY (3M)')
  $iWh=& $find @('Warehouse Stock'); $iIt=& $find @('In Transit'); $iItA=& $find @('Actual in transit'); $iOpenPR=& $find @('Open PR')
  $iDead=& $find @('Dead Stock'); $iAvail=& $find @('Available Stock'); $iReqd=& $find @('Requirment','Requirement'); $iPR=& $find @('PR Code')
  $iRol=& $find @('Reorder Level'); $iRoq=& $find @('Reorder Quantity'); $iSs=& $find @('Safety Stock')
  $iSsQ=& $find @('Safety stock in qty (Inventory input)'); $iSsD=& $find @('Safety Stock in Days (Procurement input)')
  for ($r=1; $r -lt $vals.Count; $r++) {
    $row=$vals[$r]
    $cell={ param($idx) if($idx -ge 0 -and $idx -lt $row.Count){ return "$($row[$idx])" } else { return '' } }
    $code="$(& $cell $iCode)".Trim(); if(-not $code){ continue }
    $rawName="$(& $cell $iName)".Trim()
    $parts=@($rawName -split "`r?`n" | ForEach-Object { "$_".Trim() } | Where-Object { $_ -ne '' })
    $src=''; if($parts.Count -gt 1 -and $parts[-1] -match '^(?i)(local|import)$'){ $src=$parts[-1]; $parts=@($parts[0..($parts.Count-2)]) }
    $covRaw="$(& $cell $iCov)"
    $rowsOut.Add([pscustomobject]@{
      sbu=$sbu; code=$code; name=($parts -join ' '); source=$src; uom="$(& $cell $iUom)".Trim()
      req=@((ToNum (& $cell $iReq[0])),(ToNum (& $cell $iReq[1])),(ToNum (& $cell $iReq[2])))
      short=@((ToNum (& $cell $iSh[0])),(ToNum (& $cell $iSh[1])),(ToNum (& $cell $iSh[2])))
      covDays=(ToNum ($covRaw -replace '[^0-9\.\-]','')); covDate="$(& $cell $iCovDate)".Trim(); total3m=(ToNum (& $cell $i3m))
      whStock=(ToNum (& $cell $iWh)); inTransit=(ToNum (& $cell $iIt)); actualInTransit=(ToNum (& $cell $iItA)); openPR=(ToNum (& $cell $iOpenPR))
      deadStock=(ToNum (& $cell $iDead)); available=(ToNum (& $cell $iAvail)); requirement=(ToNum (& $cell $iReqd)); prCode="$(& $cell $iPR)".Trim()
      reorderLevel=(ToNum (& $cell $iRol)); reorderQty=(ToNum (& $cell $iRoq)); safetyStock=(ToNum (& $cell $iSs))
      ssQty=(ToNum (& $cell $iSsQ)); ssDays=(ToNum (& $cell $iSsD))
    })
  }
  Write-Host ("[mp] {0}: {1} rows" -f $sbu, $rowsOut.Count)
}
Write-Host ("[mp] total {0} rows across {1} SBUs" -f $rowsOut.Count, $sbus.Count)

$mp = [pscustomobject]@{
  source = "MRP Run_Nov'26-Jan'27 (Google Sheet)"
  months = @('Nov-26','Dec-26','Jan-27')
  sbus   = $sbus.ToArray()
  rows   = $rowsOut.ToArray()
}
$outFile = Join-Path (Split-Path -Parent $DataFile) 'material-planning.json'
$json = $mp | ConvertTo-Json -Depth 12 -Compress
[IO.File]::WriteAllText($outFile, $json, (New-Object Text.UTF8Encoding($false)))
Write-Host ("[mp] wrote {0} ({1} rows)" -f $outFile, $rowsOut.Count)
