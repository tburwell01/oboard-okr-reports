<!--
FOR THE TEAMMATE (human): You don't need to run any commands yourself. Open Claude (Claude Code or
the desktop app with file/terminal access) on your work laptop, paste this whole document in, and say:
"Please follow these instructions to set up the OKR web app on my machine." Claude will do the rest
and ask you only for the things it can't do itself (e.g., signing into GitHub, clicking Sync in
SharePoint, pasting the Oboard token you were sent securely).
-->

# Set up the Corporate OKR Board web app — instructions for Claude

**You are Claude, running on a teammate's Windows laptop (Claude Code / desktop).** Your job is to
set up and launch a small local web app that generates the FY27 Corporate OKR one-pager from the
Oboard API and saves it to the user's synced SharePoint OKR folder. Work through this end to end,
adapting to what you find. The user is likely **non-technical** — do the technical steps yourself,
explain in plain language, and don't ask them to run raw commands unless necessary.

---

## What you're setting up

A localhost web app (`http://localhost:8787/`) with three buttons:
- **Test Run** — pull fresh Oboard data, build the deck with real month-over-month deltas, **do NOT reset the baseline** (safe to run repeatedly to preview).
- **Next Run** — same, but **resets the baseline** to today's grades (the monthly commit).
- **First Run** — only the very first time; the change figure equals the current percent.

It builds the slide by editing a branded PowerPoint template via PowerPoint automation (COM), then
writes `Corporate-OKR-Board-FY27-YYYY-MM-DD.pptx` to the synced SharePoint OKR folder, which OneDrive
syncs back to SharePoint.

Source repo: **https://github.com/tburwell01/oboard-okr-reports** (private).
All scripts live under `corporate-okr-monthly/` and `corporate-okr-monthly/webapp/`.

---

## Prerequisites — verify these first, report any that fail

1. **Windows + Microsoft PowerPoint (desktop) installed.** The generator automates PowerPoint via COM;
   it cannot work without desktop PowerPoint. Check:
   ```powershell
   try { $p = New-Object -ComObject PowerPoint.Application; "PowerPoint OK v$($p.Version)"; $p.Quit() } catch { "NO POWERPOINT: $($_.Exception.Message)" }
   ```
2. **The SharePoint OKR library is synced to this PC.** In SharePoint, the user must open
   *Office of Strategic Programs OSP → Documents → General → OKR* and click **Sync** (or
   *Add shortcut to OneDrive*). After syncing it appears at:
   ```
   %USERPROFILE%\Sophos Ltd\Office of Strategic Programs (OSP) - OKR
   ```
   Check: `Test-Path "$env:USERPROFILE\Sophos Ltd\Office of Strategic Programs (OSP) - OKR"`
   If False, walk the user through the Sync step before continuing.
3. **The shared Oboard API token.** The user should have received it via a secure channel
   (password manager / protected message), NOT plain email. You'll put it in `config.json`.

---

## Step 1 — Get the app files

**Preferred: clone the repo.**
```powershell
git --version   # if this errors, Git isn't installed (see fallback below)
cd $env:USERPROFILE
git clone https://github.com/tburwell01/oboard-okr-reports.git
```
The user needs **collaborator access** to the private repo (the owner, Troy Burwell, grants this on
GitHub → repo Settings → Collaborators). If the clone prompts for auth, have them sign in to GitHub.

**Verify the clone is COMPLETE** (a common failure mode is a partial/empty folder):
```powershell
$wa = "$env:USERPROFILE\oboard-okr-reports\corporate-okr-monthly\webapp"
Test-Path "$wa\Start-OKRWebApp.ps1"   # must be True
```
> Note: `config.json` is intentionally NOT in the repo (it's git-ignored because it holds the token).
> If you ever see a `webapp` folder that contains ONLY `config.json` and none of the `.ps1`/`.bat`
> files, the repo was never properly cloned — re-clone it.

**Fallback if Git is unavailable or blocked:** ask the owner to send the repo as a ZIP
(GitHub → **Code ▸ Download ZIP**). Extract it, then **clear the internet "mark of the web"** so
Endpoint Protection / execution policy doesn't block the scripts:
```powershell
Get-ChildItem "$env:USERPROFILE\oboard-okr-reports" -Recurse -File | Unblock-File
```

---

## Step 2 — Create config.json

```powershell
$wa = "$env:USERPROFILE\oboard-okr-reports\corporate-okr-monthly\webapp"
if (-not (Test-Path "$wa\config.json")) { Copy-Item "$wa\config.example.json" "$wa\config.json" }
```
Then open `$wa\config.json` and set:
- `"oboardApiToken"` → the shared token (paste it in).
- `"outputFolder"` → leave as **`"AUTO"`** (it auto-detects the synced SharePoint OKR folder).

Do NOT commit `config.json` or paste the token into any chat/log. It stays only in this file.

---

## Step 3 — Launch the app (Endpoint-Protection–friendly)

**Do NOT use a `.bat` launcher or `-ExecutionPolicy Bypass` baked into a file** — on Sophos/Intercept X
endpoints that pattern (a batch file spawning PowerShell, or `WScript.Shell` shortcut creation) is
commonly blocked as malicious behavior. Launch the script interactively instead:

```powershell
cd "$env:USERPROFILE\oboard-okr-reports\corporate-okr-monthly\webapp"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force   # process-only, not system-wide
.\Start-OKRWebApp.ps1
```
This prints `running at http://localhost:8787/` and opens the browser. Keep this PowerShell window
open while in use; closing it stops the app.

**If Endpoint Protection still blocks it**, capture the exact Sophos detection name/message and tell
the user. Likely causes and fixes:
- Block on the batch/shortcut/Bypass pattern → the interactive launch above avoids it.
- Block on *a script automating PowerPoint* or *opening a local listener* → this is behavior-based and
  needs the user's **Sophos admin to add an exclusion/exemption** for this folder or these scripts in
  Sophos Central. There is no client-side workaround; route it to IT.

---

## Step 4 — Generate

In the browser at `http://localhost:8787/`:
1. The status card should show green dots for token and output folder. (Red = fix config/sync first.)
2. Pick a mode — **Test Run** to preview, **Next Run** to commit the monthly baseline, **First Run**
   only the first time.
3. Click **Generate OKR Board**. After ~20s the file lands in the SharePoint OKR folder and syncs.

Verify:
```powershell
Get-ChildItem "$env:USERPROFILE\Sophos Ltd\Office of Strategic Programs (OSP) - OKR" -Filter "Corporate-OKR-Board-FY27-*.pptx" | Sort-Object LastWriteTime -Desc | Select-Object -First 1 Name,LastWriteTime
```

---

## Guardrails — do NOT
- Do **not** expose this app to the internet or bind it to anything other than `localhost`. It has no
  authentication and holds an API token.
- Do **not** commit `config.json` or echo the Oboard token into chat, logs, or files other than `config.json`.
- Do **not** change the objective titles/branding unless asked — those come from
  `corporate-okr-monthly/templates/onepager-template.pptx`.

## If you need deeper detail
- `corporate-okr-monthly/webapp/README.md` — usage and architecture.
- `corporate-okr-monthly/README.md` + `reference/Oboard-Public-API-Collection.json` — Oboard API
  (base URL `https://backend.okr-api.com`, workspace 15855 = Sophos, interval `FY27`).
- The generator is `corporate-okr-monthly/Build-OKR-1Pager.ps1`; the data pull (with the
  `-PreserveBaseline` switch used by Test Run) is `Invoke-CorporateOKRMonthly.ps1`.
