# Corporate OKR Board (FY27) — Web App

A small local web app so anyone on the team can generate the FY27 Corporate OKR one-pager
(pulling fresh data from Oboard) and have it saved, dated, to the SharePoint OKR folder.

It runs entirely on your own PC — no servers. It uses PowerPoint (which you already have)
to build the slide, and writes the file to the OneDrive-synced SharePoint folder so it lands
in SharePoint automatically.

---

## One-time setup (per person)

1. **Get the files.** Clone/copy the `oboard-okr-reports` repo to your PC (or a shared location).

2. **Sync the SharePoint OKR folder to your PC.**
   - Open the OKR library in SharePoint:
     *Office of Strategic Programs OSP → Documents → General → OKR*
   - Click **Sync** (or **Add shortcut to OneDrive**).
   - After it syncs, it appears in File Explorer, typically at something like:
     `C:\Users\<you>\OneDrive - Sophos Ltd\Office of Strategic Programs OSP - General\OKR`

3. **Edit `webapp\config.json`:**
   - `oboardApiToken` — the shared Oboard API token.
   - `outputFolder` — the **local synced path** from step 2 (use double backslashes, e.g.
     `C:\\Users\\you\\OneDrive - Sophos Ltd\\Office of Strategic Programs OSP - General\\OKR`).

   > Note: the token is stored in this config file. Keep the repo private and rotate the token
   > periodically (Oboard → Settings → Integrations → API Tokens).

---

## Running it

- Double-click **`Run-OKR-WebApp.bat`** (or run `Start-OKRWebApp.ps1` in PowerShell).
- Your browser opens to `http://localhost:8787/`.
- The status card shows green dots when the token and folder are configured correctly.
- Choose a mode and click **Generate OKR Board**:
  - **Test Run** *(default)* — pulls fresh data and shows real month-over-month deltas, but
    **does not reset the baseline**. Run it as many times as you like to preview; the comparison
    point stays frozen until you commit.
  - **Next Run** — same output, but **resets the baseline** to today's grades (commit). Use this
    once you're happy with the Test Run, so next month compares against these numbers.
  - **First Run** — use only the very first time; the change figure equals the current percent
    (no prior baseline yet). Also sets the baseline.

  Typical monthly flow: **Test Run → review the deck → Next Run** to commit.
- After ~20 seconds you'll get the filename. The file is saved as
  `Corporate-OKR-Board-FY27-YYYY-MM-DD.pptx` in the OKR folder and syncs to SharePoint.

Stop the app by closing the window or pressing **Ctrl+C**.

---

## How it works (for maintainers)

```
Run-OKR-WebApp.bat
  └─ Start-OKRWebApp.ps1        # localhost web UI (PowerShell HttpListener)
       ├─ Invoke-CorporateOKRMonthly.ps1   # pulls FY27 data from Oboard -> output\corporate-fy27-latest.json
       └─ Build-OKR-1Pager.ps1  -OutputDir <synced folder> [-FirstRun]
            ├─ reads grades / KR counts / owners from the JSON
            ├─ regenerates the 5 progress rings (System.Drawing)
            └─ edits a copy of templates\onepager-template.pptx via PowerPoint COM
```

- **Template:** `templates\onepager-template.pptx` (the v5 layout/branding). Update this file if
  objective *titles/descriptions*, colors, or the logo need to change.
- **Objective ↔ slot mapping:** the `$layout` block in `Build-OKR-1Pager.ps1` maps each SOP-ID to
  its ring slot, column color, and footer wording. Update it if the five corporate objectives change.
- **Deltas:** `Build-OKR-1Pager.ps1` computes month-over-month deltas from `previousGrades` in the
  JSON; `-FirstRun` overrides them to equal the current percent.

## Requirements
- Windows + Microsoft PowerPoint (desktop) installed.
- The SharePoint OKR library synced locally via OneDrive.
- A valid Oboard API token in `config.json`.
