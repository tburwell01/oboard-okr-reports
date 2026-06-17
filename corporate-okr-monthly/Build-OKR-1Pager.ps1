#requires -Version 5.1
<#
.SYNOPSIS
  Builds the FY27 Corporate OKR one-pager from corporate-fy27-latest.json, styled on the v5 template.
.DESCRIPTION
  Reads the freshly pulled JSON (run Invoke-CorporateOKRMonthly.ps1 first), regenerates the five
  progress rings from live grades, and updates each card's KR count + owner. Layout/branding come
  from the v5 template, which is copied and edited in place.
.PARAMETER FirstRun
  When set, the change figure under each ring equals the current percent (no prior baseline).
  Otherwise the change is the real month-over-month delta = currentGrade - previousGrade.
#>
param(
  [switch]$FirstRun,
  [string]$OutputDir    = "C:\Users\Troy.Burwell\OneDrive - Sophos Ltd\Documents\AI\OKRs",
  [string]$TemplatePath = (Join-Path $PSScriptRoot 'templates\onepager-template.pptx'),
  [string]$JsonPath
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$Root     = $PSScriptRoot
if (-not $JsonPath) { $JsonPath = Join-Path $Root 'output\corporate-fy27-latest.json' }
if (-not (Test-Path $JsonPath)) { throw "Missing $JsonPath - run Invoke-CorporateOKRMonthly.ps1 first." }
$data    = Get-Content $JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
$fy      = $data.intervals.FY27

$srcV5  = $TemplatePath
if (-not (Test-Path $srcV5)) { throw "Missing template: $srcV5" }
if (-not (Test-Path $OutputDir)) { throw "Output folder not found: $OutputDir`nIf this is the SharePoint OKR library, sync it to OneDrive first, then set the path in config." }
$stamp  = Get-Date -Format 'yyyy-MM-dd'
$dest   = Join-Path $OutputDir "Corporate-OKR-Board-FY27-$stamp.pptx"
$ringDir = Join-Path $env:TEMP 'okr-rings'
if (Test-Path $ringDir) { Remove-Item $ringDir -Recurse -Force }
New-Item -ItemType Directory -Path $ringDir | Out-Null

# ---------- ring generator (transparent PNG: dark track + colored progress arc + centered text) ----------
function New-Ring {
  param([string]$path,[double]$pct,[int]$displayNum,[string]$deltaTxt,[System.Drawing.Color]$col,[System.Drawing.Color]$deltaCol)
  $S = 320
  $bmp = New-Object System.Drawing.Bitmap($S,$S,[System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.Clear([System.Drawing.Color]::Transparent)
  $pad = 34; $pen = 30
  $rect = New-Object System.Drawing.RectangleF($pad,$pad,($S-2*$pad),($S-2*$pad))
  $track = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255,32,64,107)), $pen
  $g.DrawArc($track,$rect,0,360)
  if ($pct -gt 0) {
    $sweep = [float]([Math]::Max(4,([Math]::Min(100,$pct)/100.0)*360.0))
    $pp = New-Object System.Drawing.Pen $col, $pen
    $pp.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pp.EndCap   = [System.Drawing.Drawing2D.LineCap]::Round
    $g.DrawArc($pp,$rect,-90,$sweep)
    $pp.Dispose()
  }
  $sf = New-Object System.Drawing.StringFormat
  $sf.Alignment = [System.Drawing.StringAlignment]::Center
  $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
  $fNum = New-Object System.Drawing.Font('Segoe UI Semibold',68,[System.Drawing.FontStyle]::Bold)
  $fPct = New-Object System.Drawing.Font('Segoe UI',22)
  $fDel = New-Object System.Drawing.Font('Segoe UI Semibold',26,[System.Drawing.FontStyle]::Bold)
  $white = [System.Drawing.Brushes]::White
  $g.DrawString("$displayNum",$fNum,$white,(New-Object System.Drawing.RectangleF(0,84,$S,96)),$sf)
  $g.DrawString("%",$fPct,$white,(New-Object System.Drawing.RectangleF(0,184,$S,28)),$sf)
  $db = New-Object System.Drawing.SolidBrush $deltaCol
  $g.DrawString($deltaTxt,$fDel,$db,(New-Object System.Drawing.RectangleF(0,220,$S,44)),$sf)
  $g.Dispose(); $bmp.Save($path,[System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
}

# colors
$teal    = [System.Drawing.Color]::FromArgb(255,78,205,196)
$magenta = [System.Drawing.Color]::FromArgb(255,181,41,247)
$blue    = [System.Drawing.Color]::FromArgb(255,59,142,247)
$green   = [System.Drawing.Color]::FromArgb(255,46,204,113)
$red     = [System.Drawing.Color]::FromArgb(255,255,107,107)
$mutedC  = [System.Drawing.Color]::FromArgb(255,150,170,195)
function DC($d){ if($d -gt 0){$green}elseif($d -lt 0){$red}else{$mutedC} }
function DT($d){ if($d -gt 0){"+$d"}elseif($d -lt 0){"$d"}else{"0"} }

# ---------- stable layout binding: which objective sits in which v5 slot ----------
# ringId = shape Id of the v5 ring image; krLabel = wording in that card's footer.
$layout = [ordered]@{
  'SOP-523' = @{ ringId=8;  color=$teal;    krLabel='KRs' }          # People  - top
  'SOP-614' = @{ ringId=13; color=$teal;    krLabel='KRs' }          # People  - bottom
  'SOP-506' = @{ ringId=22; color=$magenta; krLabel='KRs' }          # Process - top
  'SOP-528' = @{ ringId=27; color=$magenta; krLabel='KRs' }          # Process - bottom
  'SOP-513' = @{ ringId=36; color=$blue;    krLabel='Key Results' }  # Technology
}

# ---------- index live data + previous baseline by displayId ----------
$cur = @{}; foreach ($o in $fy.objectiveRows) { $cur[$o.displayId] = $o }
$prev = @{}
if ($data.previousGrades -and $data.previousGrades[0].value) {
  foreach ($p in $data.previousGrades[0].value) { $prev[$p.displayId] = [double]$p.grade }
}

# ---------- build per-objective render + meta from data ----------
$dot = [char]0x00B7
$rings = @()
$metaUpdates = @{}
foreach ($id in $layout.Keys) {
  $L = $layout[$id]
  $o = $cur[$id]
  $grade = if ($o -and $null -ne $o.grade) { [double]$o.grade } else { 0 }
  $num   = [int][Math]::Round($grade)
  if ($FirstRun) {
    $d = $num
  } else {
    $pg = if ($prev.ContainsKey($id)) { $prev[$id] } else { 0 }
    $d  = [int][Math]::Round($grade - $pg)
  }
  $rings += @{ id=$L.ringId; pct=$grade; num=$num; d=$d; c=$L.color }

  $kr    = if ($o) { [int]$o.childCount } else { 0 }
  $owner = if ($o -and $o.owner) { ($o.owner -replace '\.', ' ').Trim() } else { '' }
  $metaUpdates[$id] = "$id  $dot  $kr $($L.krLabel)  $dot  $owner"
}

# generate ring PNGs
foreach ($r in $rings) {
  $p = Join-Path $ringDir ("ring-$($r.id).png")
  New-Ring -path $p -pct $r.pct -displayNum $r.num -deltaTxt (DT $r.d) -col $r.c -deltaCol (DC $r.d)
  $r.png = $p
}

# ---------- open LOCAL copy of v5 (avoid OneDrive lock/modal), swap rings + meta ----------
$work = Join-Path $env:TEMP 'okr-1pager-work.pptx'
Copy-Item $srcV5 $work -Force
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.DisplayAlerts = 1   # ppAlertsNone
$pres = $ppt.Presentations.Open($work, $false, $false, $false)
$sl = $pres.Slides.Item(1)

# replace ring images at their original coordinates
foreach ($r in $rings) {
  $old = $sl.Shapes | Where-Object { $_.Id -eq $r.id } | Select-Object -First 1
  if ($old) {
    $x=$old.Left; $y=$old.Top; $w=$old.Width; $h=$old.Height
    $old.Delete()
    $null = $sl.Shapes.AddPicture($r.png, $false, $true, $x, $y, $w, $h)
  }
}

# update each card's footer meta (match the shape whose text starts with the displayId)
foreach ($sh in @($sl.Shapes)) {
  if ($sh.HasTextFrame -and $sh.TextFrame.HasText) {
    $t = $sh.TextFrame.TextRange.Text
    foreach ($id in $metaUpdates.Keys) {
      if ($t -like "$id*") {
        $sh.TextFrame.TextRange.Text = $metaUpdates[$id]
        Write-Host "updated meta: $id -> $($metaUpdates[$id])"
      }
    }
  }
}

$pres.SaveAs($work, 24)   # 24 = ppSaveAsOpenXMLPresentation
$pres.Close(); $ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Copy-Item $work $dest -Force
$mode = if ($FirstRun) { 'FIRST RUN (change = percent)' } else { 'MoM deltas vs previous baseline' }
Write-Host "SAVED: $dest"
Write-Host "Mode: $mode  |  Data generated: $($data.meta.generatedLocal)"
