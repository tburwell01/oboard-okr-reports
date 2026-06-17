#requires -Version 5.1
<#
.SYNOPSIS
  Monthly Corporate OKR export from Oboard → HTML slide deck (FY27: Apr 2026 – Mar 2027).

.DESCRIPTION
  Reads corporate-okr-monthly\config.json, pulls FY27 + quarterly intervals for the
  configured workspace, writes JSON + fills slide-template.html.

.NOTES
  Auth: $env:OBOARD_API_TOKEN or oboard-api-token.txt next to this script.

.PARAMETER PreserveBaseline
  When set, the run does NOT overwrite output\grades-baseline.json. Use for a Test Run so the
  month-over-month comparison point stays frozen until a real (committing) run is done.
#>
param([switch]$PreserveBaseline)
$ErrorActionPreference = 'Stop'
$Root = $PSScriptRoot
$configPath = Join-Path $Root 'config.json'
if (-not (Test-Path $configPath)) { throw "Missing config: $configPath" }
$config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json

$base = $config.apiBaseUrl
if (-not $base) { $base = 'https://backend.okr-api.com' }

$token = [Environment]::GetEnvironmentVariable('OBOARD_API_TOKEN', 'User')
if (-not $token) { $token = $env:OBOARD_API_TOKEN }
if (-not $token) {
  $tf = Join-Path $Root 'oboard-api-token.txt'
  if (Test-Path $tf) { $token = (Get-Content $tf -Raw -Encoding UTF8).Trim() }
}
if (-not $token) { throw 'Set OBOARD_API_TOKEN or create oboard-api-token.txt (see HOWTO.txt).' }
$h = @{ 'API-Token' = $token }

function Strip-Html([string]$s) { if ([string]::IsNullOrEmpty($s)) { return '' }; return ($s -replace '<[^>]+>', '').Trim() }
function Get-ElementCount([int]$wid, [int]$intid, [int]$typeid) {
  if ($intid -le 0) { return 0 }
  $off = 0; $lim = 200; $t = 0
  do {
    $u = "$base/api/v3/elements?searchType=1&workspaceIds=$wid&intervalIds=$intid&typeIds=$typeid&limit=$lim&offset=$off"
    $c = Invoke-RestMethod -Uri $u -Headers $h
    $n = if ($null -eq $c) { 0 } elseif ($c -is [array]) { $c.Count } else { 0 }
    $t += $n; $off += $lim
    if ($n -lt $lim) { break }
    Start-Sleep -Milliseconds 180
  } while ($true)
  return $t
}
function Get-AllObjectives([int]$wid, [int]$iid) {
  if ($iid -le 0) { return @() }
  $all = @(); $off = 0; $lim = 100
  do {
    $u = "$base/api/v3/elements?searchType=1&workspaceIds=$wid&intervalIds=$iid&typeIds=1&limit=$lim&offset=$off&order=1"
    $chunk = Invoke-RestMethod -Uri $u -Headers $h
    $n = if ($null -eq $chunk) { 0 } elseif ($chunk -is [array]) { $chunk.Count } else { 0 }
    if ($n -eq 0) { break }
    $all += $chunk
    $off += $lim
    if ($n -lt $lim) { break }
    Start-Sleep -Milliseconds 180
  } while ($true)
  return $all
}
function Avg-Grade($objs) {
  $nums = @()
  foreach ($o in $objs) {
    if ($null -ne $o.grade) { try { $nums += [double]$o.grade } catch {} }
  }
  if ($nums.Count -eq 0) { return $null }
  return [math]::Round((($nums | Measure-Object -Average).Average), 1)
}
$confMap = @{ 1 = 'On track'; 2 = 'Behind'; 3 = 'At risk'; 4 = 'Not started'; 5 = 'Closed'; 6 = 'Abandoned'; 9 = 'Backlog' }
function Obj-ToRow($e) {
  $owner = if ($e.users -and $e.users[0].displayName) { $e.users[0].displayName } else { '—' }
  $grp = if ($e.groups -and $e.groups[0].name) { $e.groups[0].name } else { '' }
  $st = if ($confMap.ContainsKey([int]$e.confidenceLevelId)) { $confMap[[int]$e.confidenceLevelId] } else { ('Level ' + $e.confidenceLevelId) }
  [pscustomobject]@{
    displayId            = $e.displayId
    name                 = (Strip-Html $e.name)
    grade                = if ($null -ne $e.grade) { [math]::Round([double]$e.grade, 1) } else { $null }
    childCount           = $e.childCount
    owner                = $owner
    groupName            = $grp
    statusLabel          = $st
  }
}
function Build-IntervalBlock($wid, $intid) {
  if ($intid -le 0) {
    return @{ objectives = 0; keyResults = 0; avgGradeAchievedPct = $null; objectiveRows = @() }
  }
  $o = Get-ElementCount $wid $intid 1; Start-Sleep -Milliseconds 180
  $k = Get-ElementCount $wid $intid 4; Start-Sleep -Milliseconds 180
  $raw = Get-AllObjectives $wid $intid
  $avg = Avg-Grade $raw
  $rows = @(); foreach ($e in $raw) { $rows += Obj-ToRow $e }
  return @{ objectives = $o; keyResults = $k; avgGradeAchievedPct = $avg; objectiveRows = @($rows) }
}

$wid = [int]$config.corporateWorkspaceId
$ws = Invoke-RestMethod "$base/api/v2/workspaces" -Headers $h
$wsFlat = @($ws)
$me = $wsFlat | Where-Object { [int]$_.id -eq $wid } | Select-Object -First 1
if (-not $me) {
  throw "Workspace id $wid not found in /api/v2/workspaces. Update corporateWorkspaceId in config.json."
}
$wsName = $me.name

$exclude = @($config.excludeWorkspaceIds | ForEach-Object { [int]$_ })
if ($exclude -contains $wid) { throw "Corporate workspace $wid is listed in excludeWorkspaceIds; fix config.json." }

$annualName = $config.intervalNamesAnnual
if (-not $annualName) { $annualName = 'FY27' }
$qNames = @($config.intervalNamesQuarters)
if ($qNames.Count -eq 0) { $qNames = @('FY27Q1', 'FY27Q2', 'FY27Q3', 'FY27Q4') }

$intv = Invoke-RestMethod "$base/api/v1/intervals?workspaceId=$wid" -Headers $h
Start-Sleep -Milliseconds 200

$intervalMap = @{}
$keys = @($annualName) + $qNames
foreach ($nm in $keys) {
  $inv = @($intv) | Where-Object { $_.name -eq $nm } | Select-Object -First 1
  $iid = if ($inv) { [int]$inv.id } else { 0 }
  $intervalMap[$nm] = Build-IntervalBlock $wid $iid
}

$payload = [ordered]@{
  meta              = @{
    generatedAt       = (Get-Date).ToUniversalTime().ToString('o')
    generatedLocal    = (Get-Date).ToString('yyyy-MM-dd HH:mm')
    fiscalYearLabel   = $config.fiscalYearLabel
    periodLabel       = "$(($config.periodStart)) to $(($config.periodEnd))"
    periodStart       = $config.periodStart
    periodEnd         = $config.periodEnd
    reportTitle       = $config.reportTitle
    subtitle          = $config.subtitle
    apiBaseUrl        = $base
  }
  workspace         = @{ id = $wid; name = $wsName }
  intervals         = $intervalMap
  intervalOrder     = @($annualName) + $qNames
  previousGrades    = @()
}

# Load previous grades baseline (from last run) for delta calculation
$baselinePath = Join-Path $Root 'output\grades-baseline.json'
if (Test-Path $baselinePath) {
  try {
    $payload.previousGrades = @(Get-Content $baselinePath -Raw -Encoding UTF8 | ConvertFrom-Json)
  } catch { $payload.previousGrades = @() }
}

$outDir = Join-Path $Root 'output'
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

$json = ($payload | ConvertTo-Json -Depth 12 -Compress)
$jsonPath = Join-Path $outDir 'corporate-fy27-latest.json'
[System.IO.File]::WriteAllText($jsonPath, $json, [Text.UTF8Encoding]::new($false))

# Save current grades as the new baseline for next month (skipped on a Test Run)
if (-not $PreserveBaseline) {
  $annualBlock = $intervalMap[$annualName]
  $currentGrades = @()
  if ($annualBlock -and $annualBlock.objectiveRows) {
    foreach ($row in $annualBlock.objectiveRows) {
      $currentGrades += @{ displayId = $row.displayId; grade = $row.grade }
    }
  }
  $gradesJson = ($currentGrades | ConvertTo-Json -Depth 4 -Compress)
  [System.IO.File]::WriteAllText($baselinePath, $gradesJson, [Text.UTF8Encoding]::new($false))
}

$tplPath = Join-Path $Root 'slide-template.html'
if (-not (Test-Path $tplPath)) { throw "Missing slide-template.html in $Root" }
$tpl = [System.IO.File]::ReadAllText($tplPath, [Text.UTF8Encoding]::new($false))
$html = $tpl.Replace('__CORPORATE_OKR_JSON__', $json)

$stamp = Get-Date -Format 'yyyy-MM-dd'
$htmlPath = Join-Path $outDir "Corporate-OKRs-FY27-monthly-$stamp.html"
[System.IO.File]::WriteAllText($htmlPath, $html, [Text.UTF8Encoding]::new($false))

Write-Host "OK"
Write-Host "  JSON: $jsonPath"
Write-Host "  HTML: $htmlPath"
if ($PreserveBaseline) { Write-Host "  Baseline PRESERVED (test run, not reset)" }
else { Write-Host "  Baseline saved: $baselinePath" }
