#requires -Version 5.1
<#
  Corporate OKR Board (FY27) - local web app.
  Serves a small UI at http://localhost:<port>/ so anyone on the team can:
    - choose First Run (change = percent) or Next Run (real month-over-month deltas)
    - pull fresh data from Oboard and generate the dated one-pager
    - have it saved to the OneDrive-synced SharePoint OKR folder
  No servers, no installs - just Windows + Office (for PowerPoint generation).
#>
$ErrorActionPreference = 'Stop'
$here     = $PSScriptRoot
$repo     = Split-Path $here -Parent                       # ...\corporate-okr-monthly
$pull     = Join-Path $repo 'Invoke-CorporateOKRMonthly.ps1'
$build    = Join-Path $repo 'Build-OKR-1Pager.ps1'
$cfgPath  = Join-Path $here 'config.json'
if (-not (Test-Path $cfgPath)) { throw "Missing config.json next to this script." }
$cfg      = Get-Content $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
$port     = if ($cfg.port) { [int]$cfg.port } else { 8787 }

# The synced SharePoint OKR library is per-user but the relative shape is constant.
# If outputFolder is blank/AUTO/not found, fall back to this auto-detected path.
function Resolve-OutputFolder {
  $configured = $cfg.outputFolder
  if ($configured -and $configured -notmatch 'REPLACE_|^AUTO$' -and (Test-Path $configured)) { return $configured }
  $auto = Join-Path $env:USERPROFILE 'Sophos Ltd\Office of Strategic Programs (OSP) - OKR'
  if (Test-Path $auto) { return $auto }
  return $configured  # return whatever was configured so the error message is informative
}

function Get-State {
  $tokenOk  = $cfg.oboardApiToken -and $cfg.oboardApiToken -notmatch 'PASTE_'
  $folder   = Resolve-OutputFolder
  $folderOk = $folder -and $folder -notmatch 'REPLACE_|^AUTO$' -and (Test-Path $folder)
  [pscustomobject]@{
    tokenOk    = [bool]$tokenOk
    folder     = $folder
    folderOk   = [bool]$folderOk
    ready      = ([bool]$tokenOk -and [bool]$folderOk)
    lastFile   = (Get-Date -Format 'yyyy-MM-dd')
  }
}

function Invoke-Generate([string]$mode) {
  $log = New-Object System.Text.StringBuilder
  try {
    $st = Get-State
    $outFolder = Resolve-OutputFolder
    if (-not $st.tokenOk)  { throw "Oboard token not set in config.json." }
    if (-not $st.folderOk) { throw "Output folder not found. Sync the SharePoint OKR library to OneDrive, or set 'outputFolder' in config.json. Tried: $outFolder" }

    # close any stray PowerPoint that could lock files
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Milliseconds 800

    $env:OBOARD_API_TOKEN = $cfg.oboardApiToken
    # Test Run pulls fresh data but keeps the baseline frozen (no reset).
    [void]$log.AppendLine("== Pulling fresh data from Oboard ==")
    $pullArgs = @{}
    if ($mode -eq 'test') { $pullArgs['PreserveBaseline'] = $true }
    $pullOut = & $pull @pullArgs *>&1 | Out-String
    [void]$log.AppendLine($pullOut.Trim())

    [void]$log.AppendLine("== Building one-pager ($mode) ==")
    $args = @{ OutputDir = $outFolder }
    if ($mode -eq 'first') { $args['FirstRun'] = $true }
    $buildOut = & $build @args *>&1 | Out-String
    [void]$log.AppendLine($buildOut.Trim())

    $stamp = Get-Date -Format 'yyyy-MM-dd'
    $file  = Join-Path $outFolder "Corporate-OKR-Board-FY27-$stamp.pptx"
    if (-not (Test-Path $file)) { throw "Build finished but output file not found: $file" }
    return [pscustomobject]@{ ok=$true; file=$file; fileName=(Split-Path $file -Leaf); mode=$mode; log=$log.ToString() }
  }
  catch {
    [void]$log.AppendLine("ERROR: " + $_.Exception.Message)
    return [pscustomobject]@{ ok=$false; error=$_.Exception.Message; log=$log.ToString() }
  }
  finally { $env:OBOARD_API_TOKEN = $null }
}

$html = @'
<!DOCTYPE html><html><head><meta charset="utf-8"><title>Corporate OKR Board - FY27</title>
<style>
 :root{--bg:#071C3E;--panel:#122D56;--teal:#4ECDC4;--mag:#B529F7;--blue:#3B8EF7;--muted:#96AAC3;--ok:#2ECC71;--err:#FF6B6B}
 *{box-sizing:border-box;font-family:'Segoe UI',Calibri,system-ui,sans-serif}
 body{margin:0;background:var(--bg);color:#fff;min-height:100vh;display:flex;justify-content:center}
 .wrap{width:680px;max-width:94vw;padding:36px 0}
 h1{font-size:30px;font-weight:800;margin:0 0 4px}
 .sub{color:var(--muted);margin:0 0 24px}
 .card{background:var(--panel);border-radius:14px;padding:22px 24px;margin-bottom:18px}
 .opt{display:flex;gap:12px;align-items:flex-start;padding:14px;border:1px solid #20406B;border-radius:10px;margin-bottom:10px;cursor:pointer}
 .opt:hover{border-color:var(--teal)}
 .opt input{margin-top:4px}
 .opt b{font-size:15px}.opt span{color:var(--muted);font-size:13px;display:block;margin-top:2px}
 button{background:var(--teal);color:#06243f;border:0;border-radius:10px;padding:13px 22px;font-size:16px;font-weight:700;cursor:pointer;width:100%}
 button:disabled{opacity:.5;cursor:not-allowed}
 .row{display:flex;justify-content:space-between;font-size:13px;color:var(--muted);margin:4px 0}
 .dot{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:6px}
 .green{background:var(--ok)}.red{background:var(--err)}
 pre{background:#06203c;border-radius:8px;padding:14px;font-size:12px;color:#c8d6ea;white-space:pre-wrap;max-height:260px;overflow:auto;display:none}
 .result{padding:14px;border-radius:8px;margin-top:14px;display:none;font-size:14px}
 .result.ok{background:rgba(46,204,113,.15);border:1px solid var(--ok)}
 .result.bad{background:rgba(255,107,107,.15);border:1px solid var(--err)}
 .spin{display:none;margin-left:8px}
 a{color:var(--teal)}
</style></head><body><div class="wrap">
 <h1>Corporate OKR Board &mdash; FY27</h1>
 <p class="sub">Pull fresh data from Oboard and generate the dated one-pager. Saved to the SharePoint OKR folder.</p>

 <div class="card" id="status">
   <div class="row"><span>Oboard token</span><span id="s-token">checking&hellip;</span></div>
   <div class="row"><span>Output folder</span><span id="s-folder">checking&hellip;</span></div>
 </div>

 <div class="card">
   <label class="opt"><input type="radio" name="mode" value="test" checked>
     <span><b>Test Run</b><span>Preview with fresh data and real month-over-month deltas, but <b>does not reset the baseline</b>. Run as many times as you like; the comparison point stays put.</span></span></label>
   <label class="opt"><input type="radio" name="mode" value="next">
     <span><b>Next Run</b><span>Commit the monthly run. Same as a Test Run but <b>resets the baseline</b> to today's grades, so next month compares against these.</span></span></label>
   <label class="opt"><input type="radio" name="mode" value="first">
     <span><b>First Run</b><span>Use only the very first time. No prior baseline, so the change figure equals the current percent (and sets the baseline).</span></span></label>
 </div>

 <button id="go">Generate OKR Board</button>
 <span class="spin" id="spin">Working&hellip; (pulling data + building, ~20s)</span>

 <div class="result" id="result"></div>
 <pre id="log"></pre>
</div>
<script>
async function refresh(){
  const r = await fetch('/state'); const s = await r.json();
  document.getElementById('s-token').innerHTML = s.tokenOk
    ? '<span class="dot green"></span>configured' : '<span class="dot red"></span>not set in config.json';
  document.getElementById('s-folder').innerHTML = s.folderOk
    ? '<span class="dot green"></span>'+s.folder : '<span class="dot red"></span>not found &mdash; sync SharePoint OKR folder';
  document.getElementById('go').disabled = !s.ready;
}
refresh();
document.getElementById('go').onclick = async ()=>{
  const mode = document.querySelector('input[name=mode]:checked').value;
  const btn=document.getElementById('go'), spin=document.getElementById('spin');
  const res=document.getElementById('result'), log=document.getElementById('log');
  btn.disabled=true; spin.style.display='inline'; res.style.display='none'; log.style.display='none';
  try{
    const r = await fetch('/generate',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({mode})});
    const d = await r.json();
    res.style.display='block';
    if(d.ok){
      var label = mode==='first' ? 'First Run' : (mode==='test' ? 'Test Run' : 'Next Run');
      var baseNote = mode==='test' ? 'Baseline NOT reset — run again or commit with Next Run when ready.' : 'Baseline reset to today’s grades.';
      res.className='result ok';
      res.innerHTML='&#10003; Created <b>'+d.fileName+'</b><br>Saved to the OKR folder (syncing to SharePoint).<br>Mode: '+label+' &middot; '+baseNote;
    }
    else { res.className='result bad'; res.innerHTML='&#10007; '+d.error; }
    if(d.log){ log.style.display='block'; log.textContent=d.log; }
  }catch(e){ res.style.display='block'; res.className='result bad'; res.textContent='Request failed: '+e; }
  finally{ btn.disabled=false; spin.style.display='none'; refresh(); }
};
</script></body></html>
'@

function Send-Response($ctx, $body, $type='application/json', [int]$code=200) {
  $buf = [Text.Encoding]::UTF8.GetBytes($body)
  $ctx.Response.StatusCode = $code
  $ctx.Response.ContentType = $type
  $ctx.Response.ContentLength64 = $buf.Length
  $ctx.Response.OutputStream.Write($buf, 0, $buf.Length)
  $ctx.Response.OutputStream.Close()
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()
Write-Host "Corporate OKR web app running at http://localhost:$port/  (press Ctrl+C to stop)" -ForegroundColor Cyan
Start-Process "http://localhost:$port/"

try {
  while ($listener.IsListening) {
    $ctx = $listener.GetContext()
    $path = $ctx.Request.Url.AbsolutePath
    $method = $ctx.Request.HttpMethod
    try {
      if ($path -eq '/' -and $method -eq 'GET') {
        Send-Response $ctx $html 'text/html; charset=utf-8'
      }
      elseif ($path -eq '/state' -and $method -eq 'GET') {
        Send-Response $ctx ((Get-State) | ConvertTo-Json)
      }
      elseif ($path -eq '/generate' -and $method -eq 'POST') {
        $reader = New-Object IO.StreamReader($ctx.Request.InputStream, $ctx.Request.ContentEncoding)
        $payload = $reader.ReadToEnd() | ConvertFrom-Json
        $mode = switch ($payload.mode) { 'first' { 'first' } 'test' { 'test' } default { 'next' } }
        $result = Invoke-Generate $mode
        Send-Response $ctx ($result | ConvertTo-Json -Depth 6)
      }
      else { Send-Response $ctx '{"error":"not found"}' 'application/json' 404 }
    }
    catch {
      Send-Response $ctx (@{ ok=$false; error=$_.Exception.Message } | ConvertTo-Json) 'application/json' 500
    }
  }
}
finally { $listener.Stop(); $listener.Close() }
