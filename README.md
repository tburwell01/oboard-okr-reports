# Oboard OKR Reports

Sophos OKR reporting dashboards powered by the [Oboard](https://oboard.io/) Public API.

**Repo is private.** Do not commit API tokens or credentials.

---

## Reports in this Repo

### 1. Corporate OKR Board Dashboard (corporate-okr-monthly/)

A single-slide (1280x720) HTML dashboard showing the **5 Corporate Objectives** (Sophos workspace, group = "Sophos") with progress rings and month-over-month deltas.

**Files:**
| File | Purpose |
|------|---------|
| `Invoke-CorporateOKRMonthly.ps1` | PowerShell script that pulls data from Oboard API and generates the HTML |
| `config.json` | Configuration (workspace ID, interval names, excluded workspaces) |
| `slide-template.html` | HTML/JS template with `__CORPORATE_OKR_JSON__` placeholder |
| `Run-CorporateOKR-Monthly.bat` | Double-click runner for Windows |
| `HOWTO.txt` | Quick-start instructions |
| `output/` | Generated outputs (HTML slides, JSON data, grades baseline) |

#### How to Regenerate

1. **Set your Oboard API token** (one of these methods):
   - Set environment variable: `9189bf758f33098f7cd9850213c95b30585e0a0a68002e16ded6f27c52e01e92 = 'your-token-here'`
   - Or create `corporate-okr-monthly/oboard-api-token.txt` containing just the token
2. **Run the script:**
   `powershell
   cd corporate-okr-monthly
   .\Invoke-CorporateOKRMonthly.ps1
   `
   Or double-click `Run-CorporateOKR-Monthly.bat`
3. **Output** is written to `output/Corporate-OKRs-FY27-monthly-YYYY-MM-DD.html`
4. Open the HTML file in a browser — it renders as a 1280x720 slide ready for screenshots or PDF export.

#### Configuration (config.json)
- `corporateWorkspaceId`: Workspace ID for the Corporate/Sophos workspace (currently `15855`)
- `intervalNamesAnnual`: Annual interval name (`FY27`)
- `intervalNamesQuarters`: Array of quarterly interval names
- `excludeWorkspaceIds`: Workspaces to skip (currently `[19388]` = Oboard Tech Team)

---

### 2. FY27 Readiness Report (y27-readiness-report/)

A full-page HTML report showing **all workspaces' OKR creation status** for FY27 (annual + Q1–Q4), including:
- Summary cards (workspace count, coverage %)
- Readiness progress bar
- Quarterly breakdown table (Obj/KR per interval)
- Per-workspace objective listings
- Gaps & observations section

#### How to Regenerate

This report is generated via **Cursor Agent** (or any environment with PowerShell and internet access):

1. **Set your Oboard API token:**
   `powershell
   9189bf758f33098f7cd9850213c95b30585e0a0a68002e16ded6f27c52e01e92 = 'your-token-here'
   `

2. **Run these API calls** (the agent does this automatically):
   `powershell
   https://api.github.com/repos/tburwell01/oboard-okr-reports/contents = 'https://backend.okr-api.com'
    = @{ 'API-Token' = 9189bf758f33098f7cd9850213c95b30585e0a0a68002e16ded6f27c52e01e92 }

   # Get workspaces (exclude id 19388)
    = Invoke-RestMethod -Uri "https://api.github.com/repos/tburwell01/oboard-okr-reports/contents/api/v2/workspaces" -Headers 
    =  | Where-Object { .id -ne 19388 }

   # For each workspace, get intervals and find FY27/FY27Q1-Q4
   # Then count objectives (typeId=1) and key results (typeId=4):
   # GET /api/v3/elements?searchType=1&workspaceIds={id}&intervalIds={id}&typeIds=1&limit=200
   # GET /api/v3/elements?searchType=1&workspaceIds={id}&intervalIds={id}&typeIds=4&limit=200
   `

3. **Or simply ask Cursor Agent:**
   > "Regenerate the FY27 Readiness Report pulling updated data from Oboard"

   The agent will call the API, collect counts, and rewrite the HTML.

---

## API Reference

- **Base URL:** `https://backend.okr-api.com`
- **Auth header:** `API-Token: <your-token>`
- **Key endpoints:**
  - `GET /api/v2/workspaces` — list all workspaces
  - `GET /api/v1/intervals?workspaceId={id}` — intervals for a workspace
  - `GET /api/v3/elements?searchType=1&workspaceIds={id}&intervalIds={id}&typeIds={1|4}&limit=200&offset=0` — objectives (type 1) or key results (type 4)
- **Postman collection:** See `reference/Oboard-Public-API-Collection.json`

---

## Security

- **Never commit your API token** to this repo.
- Use environment variables or a local `.txt` file (listed in `.gitignore`).
- Rotate tokens periodically via Oboard Settings > Integrations > API Tokens.

---

## Generated

Last refreshed: **May 28, 2026**