# Runbook — Targeted Connected Apps & Chat Audit
### `Get-TargetedUserReport.ps1`

---

## Table of Contents

1. [Purpose & Scope](#1-purpose--scope)
2. [How It Differs from a Full-Domain Scan](#2-how-it-differs-from-a-full-domain-scan)
3. [Prerequisites](#3-prerequisites)
4. [Files in This Package](#4-files-in-this-package)
5. [Step 1 — Prepare Your Target User List](#5-step-1--prepare-your-target-user-list)
6. [Step 2 — Install & Configure GAM (First-Time Only)](#6-step-2--install--configure-gam-first-time-only)
7. [Step 3 — Verify GAM Authorisation & Required Scopes](#7-step-3--verify-gam-authorisation--required-scopes)
8. [Step 4 — Run the Script](#8-step-4--run-the-script)
9. [GAM Auto-Detection Logic](#9-gam-auto-detection-logic)
10. [Script Execution Flow (Section by Section)](#10-script-execution-flow-section-by-section)
11. [Output Files Reference](#11-output-files-reference)
12. [HTML Report Walkthrough](#12-html-report-walkthrough)
13. [Interpreting the Results](#13-interpreting-the-results)
14. [Troubleshooting](#14-troubleshooting)
15. [Scheduling & Automation](#15-scheduling--automation)
16. [Security & Compliance Notes](#16-security--compliance-notes)

---

## 1. Purpose & Scope

This script performs a **targeted** Google Workspace security and compliance audit for a
**specific list of users** (e.g. 10 out of 100,000). It answers:

| Question | Data Source |
|----------|-------------|
| Which third-party apps has this user authorised? | OAuth token records |
| What permissions did they grant each app? | OAuth scopes |
| Which Chat Spaces is this user a member of? | Google Chat API |
| Have any bots/apps interacted with this user in Chat? | Chat audit log |

**What it does NOT do** (by design):
- It does not enumerate all domain users (`gam print users` is never called).
- It does not scan users outside the supplied list.
- It does not modify any data — read-only queries only.

---

## 2. How It Differs from a Full-Domain Scan

| Aspect | `Get-ConnectedAppsReport.ps1` | `Get-TargetedUserReport.ps1` |
|--------|-------------------------------|------------------------------|
| Users scanned | All domain users | Only your supplied list |
| `gam print users` called? | Yes (slow on 100k tenants) | **No** |
| OAuth tokens | All users | Target users only |
| Runtime (10 users) | 20–60 min | **30–120 seconds** |
| Chat Spaces | Admin-wide dump | Per-user, target users only |
| Chat Bots | Not included | Filtered from audit log |
| Use case | Periodic full audit | Targeted investigation / spot-check |

---

## 3. Prerequisites

### 3.1 System Requirements

| Requirement | Minimum | Notes |
|-------------|---------|-------|
| Operating System | Windows 10 / Server 2016 | Script uses `.exe` paths; Linux/macOS supported with `gam` binary |
| PowerShell | **5.1** or newer | Check: `$PSVersionTable.PSVersion` |
| GAM binary | GAM 6, GAM7, or GAMADV-XTD3 | GAM7/GAMADV required for Chat features |
| Internet access | Required | GAM calls Google APIs |
| Disk space | ~50 MB per run | For CSVs and HTML report |

### 3.2 Google Workspace Requirements

| Requirement | Detail |
|-------------|--------|
| Admin account | Super Admin **or** delegated admin with Reports + Security roles |
| Domain | Google Workspace Business Starter / Standard / Plus / Enterprise |
| GAM OAuth | Must be authorised (see Step 2) |

---

## 4. Files in This Package

```
connected-Apps/
├── Get-TargetedUserReport.ps1   ← Main script (this runbook covers this file)
├── Get-ConnectedAppsReport.ps1  ← Full-domain version (all users)
├── Get-ChatBots (1).ps1         ← Standalone Chat bot audit (tenant-wide)
├── .gampath                     ← Auto-created after first successful GAM detection
└── Runbook.md                   ← This document
```

**`.gampath`** is created automatically the first time you confirm a GAM path interactively.
It stores the full path to `gam.exe` so every subsequent run skips detection entirely.
You can delete this file at any time to force re-detection.

---

## 5. Step 1 — Prepare Your Target User List

The script accepts users in three ways. Choose whichever fits your workflow.

### Option A — Inline on the command line (best for ≤ 10 users)

```powershell
.\Get-TargetedUserReport.ps1 `
    -Users "alice@corp.com","bob@corp.com","carol@corp.com"
```

### Option B — Plain text file (one email per line)

Create `targets.txt`:
```
alice@corp.com
bob@corp.com
carol@corp.com
```

```powershell
.\Get-TargetedUserReport.ps1 -UsersFile ".\targets.txt"
```

### Option C — CSV file with a `primaryEmail` column

Create `targets.csv`:
```csv
primaryEmail,Department,Notes
alice@corp.com,Engineering,Leaver review
bob@corp.com,Finance,Security flag raised
carol@corp.com,HR,Offboarding
```

> **Tip:** The script also accepts a column named `email` or, as a last resort, the first column of the CSV.
> Extra columns (Department, Notes, etc.) are silently ignored.

```powershell
.\Get-TargetedUserReport.ps1 -UsersFile ".\targets.csv"
```

### Combining both

You can mix `-Users` and `-UsersFile` together. Duplicates are removed automatically.

```powershell
.\Get-TargetedUserReport.ps1 `
    -Users "dave@corp.com" `
    -UsersFile ".\targets.csv"
```

---

## 6. Step 2 — Install & Configure GAM (First-Time Only)

> Skip this section if GAM is already installed and authorised on this machine.

### 6.1 Install GAM7 / GAMADV-XTD3 (recommended)

1. Download the latest release from:
   **https://github.com/taers232c/GAMADV-XTD3/releases/latest**
2. Extract to a permanent folder, e.g. `C:\GAM7\`
3. Open PowerShell in that folder and run:
   ```powershell
   .\gam.exe config drive_dir C:\GAMWork create
   ```

### 6.2 Authorise GAM against your Google Workspace tenant

Run the following and follow the browser prompts with your **Super Admin** account:

```powershell
C:\GAM7\gam.exe oauth create
```

Authorise **all scopes** presented. At minimum these must be enabled:

| Scope | Required for |
|-------|-------------|
| Admin SDK - Directory (read) | Listing users |
| Admin SDK - Reports | OAuth token audit log, Chat report |
| Google OAuth Token (read/delete) | `gam print tokens` |
| Google Chat - Admin | Chat Spaces & members (`-IncludeChatSpaces`) |
| Cloud Identity - Policies | Marketplace apps (used by full report) |

### 6.3 Verify authorisation

```powershell
C:\GAM7\gam.exe oauth info
C:\GAM7\gam.exe user alice@corp.com print tokens
```

Both commands should return data without `403` or `401` errors.

---

## 7. Step 3 — Verify GAM Authorisation & Required Scopes

Run these quick checks before executing the main script.

```powershell
# 1. Check GAM version
gam version

# 2. Verify OAuth is configured
gam oauth info

# 3. Test token retrieval for one user
gam user alice@corp.com print tokens

# 4. Test Chat access (only needed if using -IncludeChatSpaces / -IncludeChatBots)
gam user alice@corp.com print chatspaces
gam report chat daysago 1
```

**Expected output for test 3** — a CSV table listing app name, client ID, and scopes.
If you see `403 insufficientPermissions`, run `gam oauth update` and re-authorise.

---

## 8. Step 4 — Run the Script

Open **PowerShell 5.1+** in the `connected-Apps` folder (or any folder if you specify `-GamPath`).

### Execution Policy (one-time unlock)

If you see a security warning on first run:
```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
# or to unblock just this file:
Unblock-File -Path ".\Get-TargetedUserReport.ps1"
```

### Scenario A — Connected Apps only (fastest, no Chat)

```powershell
.\Get-TargetedUserReport.ps1 -Users "alice@corp.com","bob@corp.com"
```

**Runtime:** ~5–30 seconds for 10 users.
**What it collects:** OAuth tokens + app aggregation per user.

---

### Scenario B — Connected Apps + Chat Spaces

```powershell
.\Get-TargetedUserReport.ps1 `
    -UsersFile ".\targets.csv" `
    -IncludeChatSpaces
```

**Runtime:** ~30–90 seconds for 10 users.
**What it adds:** Each target user's Chat Space memberships + members of those spaces.
**Requires:** GAM7 / GAMADV-XTD3 + Google Chat API admin scope.

---

### Scenario C — Full targeted run (Connected Apps + Chat Spaces + Chat Bots)

```powershell
.\Get-TargetedUserReport.ps1 `
    -UsersFile ".\targets.csv" `
    -IncludeChatSpaces `
    -IncludeChatBots `
    -ChatDaysAgo 60
```

**Runtime:** ~60–180 seconds for 10 users (Chat audit log pull adds most of the time).
**What it adds:** Bot/app events in Chat where a target user was involved, last 60 days.
**Requires:** GAM7 / GAMADV-XTD3 + Reports API scope.

---

### Scenario D — Custom GAM path and output folder

```powershell
.\Get-TargetedUserReport.ps1 `
    -Users "alice@corp.com" `
    -GamPath "C:\GAM7\gam.exe" `
    -OutputDir "C:\Reports\Alice_$(Get-Date -f 'yyyyMMdd')" `
    -IncludeChatSpaces `
    -IncludeChatBots
```

---

### Scenario E — Larger batch from a CSV, extended Chat history

```powershell
.\Get-TargetedUserReport.ps1 `
    -UsersFile ".\leavers_q2.csv" `
    -IncludeChatSpaces `
    -IncludeChatBots `
    -ChatDaysAgo 90 `
    -DwdTimeoutSeconds 180 `
    -OutputDir ".\Reports\Leavers_Q2"
```

---

### All Parameters Reference

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Users` | `string[]` | `@()` | One or more email addresses |
| `-UsersFile` | `string` | `""` | Path to CSV or TXT file with emails |
| `-GamPath` | `string` | *(auto-detect)* | Full path to `gam.exe` |
| `-OutputDir` | `string` | `.\TargetedReport_<timestamp>` | Folder for all output files |
| `-IncludeChatSpaces` | switch | off | Collect Chat Space memberships |
| `-IncludeChatBots` | switch | off | Pull Chat audit log and filter for bots |
| `-ChatDaysAgo` | `int` | `30` | Days of Chat audit history to retrieve |
| `-DwdTimeoutSeconds` | `int` | `90` | Per-GAM-job timeout in seconds |

---

## 9. GAM Auto-Detection Logic

The script searches for `gam.exe` in the following priority order.
It **validates** each candidate by running `gam version` — so a wrong binary (e.g. a game's `gam.exe`) is never accepted.

| Priority | Source | Example path |
|----------|--------|--------------|
| 1 | `-GamPath` parameter | Whatever you pass |
| 2 | `.gampath` config file | Saved from a previous run |
| 3 | Current working directory | `.\gam.exe` |
| 4 | Script's own directory | `$PSScriptRoot\gam.exe` |
| 5 | PowerShell's `$PWD` | Explicit working directory |
| 6 | System `PATH` (bare `gam`) | Any directory in `$env:PATH` |
| 7 | Well-known install paths | `C:\GAM7\`, `%LOCALAPPDATA%\GAM7\`, etc. |
| 8 | Every PATH directory | `<path-entry>\gam.exe` |
| 9 | Windows Registry | `HKLM\SOFTWARE\GAM\InstallPath` |

### If auto-detection fails

The script shows an **interactive menu** instead of exiting:

```
  What would you like to do?
    [1] Enter the full path to gam / gam.exe manually
    [2] Search for gam.exe on this machine (scans C:\, D:\, E:\)
    [3] Open the GAM7 / GAMADV-XTD3 download page in your browser
    [Q] Quit
```

- **Option 1:** Paste the full path. The script runs `gam version` to confirm it works, then asks if you want to save the path to `.gampath` for future runs.
- **Option 2:** Triggers a recursive filesystem scan across `C:\`, `D:\`, `E:\`. Up to 10 hits are shown as a numbered list. Select one and it is validated before use.
- **Option 3:** Opens `https://github.com/taers232c/GAMADV-XTD3/releases/latest` in your default browser.

### Saving the GAM path permanently

After any successful interactive selection, you will be asked:
```
  Save this path for future runs? (y/n)
```

Answering `y` writes the confirmed path to `.gampath` (next to the script file).
On every subsequent run the script reads `.gampath` as priority #2 — detection completes in milliseconds.

To **reset** detection (e.g. after moving GAM to a new folder):
```powershell
Remove-Item ".\connected-Apps\.gampath"
```

---

## 10. Script Execution Flow (Section by Section)

```
START
  │
  ├─ 0. Locate GAM binary  ────────────────────────── auto-detect → interactive if needed
  │
  ├─ 0. Load Target Users  ────────────────────────── -Users / -UsersFile → deduplicate
  │       └─ writes _target_users.csv (temp, used by GAM multiprocess)
  │
  ├─ 1. OAuth Tokens  ─────────────────────────────── gam multiprocess → tokens_target_users.csv
  │       ├─ auto-detect column names (GAM version differences)
  │       ├─ build per-user summary → tokens_per_user.csv
  │       └─ aggregate by app → apps_aggregated.csv
  │
  ├─ 2. Chat Spaces  ──────────────────────────────── [only if -IncludeChatSpaces]
  │       ├─ per-user: gam user <email> print chatspaces
  │       ├─ merge all rows → chat_spaces_target.csv
  │       └─ fetch members of each unique space → chat_members_target.csv
  │
  ├─ 3. Chat Bot Activity  ────────────────────────── [only if -IncludeChatBots]
  │       ├─ gam report chat daysago <N>  (tenant-wide pull, runs in background job)
  │       ├─ filter rows: actor OR target must be a target user
  │       ├─ filter rows: event must be bot/app-related
  │       ├─ save filtered events → chat_bot_events_target.csv
  │       └─ aggregate per bot → chat_bot_summary_target.csv
  │
  ├─ 4. Build HTML Report  ────────────────────────── all data → TargetedReport.html
  │
  └─ DONE  ────────────────────────────────────────── print summary, offer to open HTML
```

### Section 1 — OAuth Tokens detail

The GAM command used internally is:
```
gam redirect csv tokens_target_users.csv multiprocess csv _target_users.csv
    gam user "~primaryEmail" print tokens
```

This runs one `gam user <email> print tokens` per user **in parallel** (GAM multiprocess),
so 10 users complete in roughly the same time as 1.

**Column auto-detection:** Different GAM versions use different column names:
- User column: tries `user` then `userEmail`
- App column: tries `displayText` then `appName`

### Section 2 — Chat Spaces detail

For each target user, the script runs:
```
gam user <email> print chatspaces
```
This returns only spaces the user is **currently a member of** — no admin-wide dump.

After collecting all spaces, it fetches members of each unique space:
```
gam print chatmembers <space-resource-name> asadmin
```

### Section 3 — Chat Bot Activity detail

The Chat audit log (`gam report chat`) is **tenant-wide** — there is no per-user filter at the GAM level.
The script pulls the entire log for the specified period, then filters in PowerShell for:
1. Rows where `actor.email` OR `target user` is in the target list, AND
2. Rows where the event name contains `app` or `bot`, OR `resourceDetails.1.type` is `APPLICATION`

The raw chat report file is deleted after processing to save disk space.

---

## 11. Output Files Reference

All files are written to the output folder (`.\TargetedReport_<timestamp>\` by default).

### Always produced

| File | Description | Key columns |
|------|-------------|-------------|
| `_target_users.csv` | Temp: list of emails fed to GAM multiprocess | `primaryEmail` |
| `tokens_target_users.csv` | Raw OAuth token records | `user/userEmail`, `displayText/appName`, `clientId`, `scopes`, `anonymous`, `nativeApp` |
| `tokens_per_user.csv` | One row per target user | `User`, `AppCount`, `Apps` |
| `apps_aggregated.csv` | One row per distinct app | `AppName`, `ClientId`, `UserCount`, `Users`, `Scopes` |
| `TargetedReport.html` | Interactive HTML report | *(see Section 12)* |

### Produced when `-IncludeChatSpaces` is used

| File | Description | Key columns |
|------|-------------|-------------|
| `chat_spaces_target.csv` | Chat Space memberships for target users | `name` (resource), `displayName`, `spaceType`, `TargetUser` |
| `chat_members_target.csv` | All members of the spaces found | `member.name`, `member.type`, `role`, `SpaceResourceName` |

### Produced when `-IncludeChatBots` is used

| File | Description | Key columns |
|------|-------------|-------------|
| `chat_bot_events_target.csv` | Raw bot events filtered to target users | `EventTime`, `EventName`, `TargetUserInvolved`, `ActorEmail`, `AppName`, `AppId`, `SpaceID`, `SpaceName` |
| `chat_bot_summary_target.csv` | One row per unique bot | `AppName`, `AppId`, `TotalEvents`, `EventTypes`, `TargetUsersActive`, `UniqueSpaces`, `SpaceNames`, `FirstSeen`, `LastSeen` |

### Column glossary

| Column | Meaning |
|--------|---------|
| `clientId` | OAuth client ID — unique identifier for the app in Google's system |
| `scopes` | Space-separated list of OAuth scopes the user granted |
| `anonymous` | `true` if the app is not verified by Google |
| `nativeApp` | `true` if installed as a native/desktop app |
| `spaceType` | `SPACE` (named room), `GROUP_CHAT`, or `DIRECT_MESSAGE` |
| `AppId` | GCP resource ID of the Chat bot (format: `gcp/<number>`) |
| `EventName` | E.g. `app_added`, `app_removed`, `message_posted` |
| `TargetUserInvolved` | Which target user triggered or was targeted by this event |

---

## 12. HTML Report Walkthrough

The `TargetedReport.html` file is self-contained — open it in any browser, no server needed.

### Summary Cards (top row)

Six stat cards show at a glance:

```
┌──────────────┬──────────────┬──────────────┬──────────────┬──────────────┬──────────────┐
│ Target Users │  OAuth Apps  │Token Records │Space Records │ Unique Bots  │  Bot Events  │
└──────────────┴──────────────┴──────────────┴──────────────┴──────────────┴──────────────┘
```

### Section: Target Users Scanned

A bulleted list (3 columns) of every email address that was scanned.

### Section: Connected Apps (OAuth Tokens)

Three tabs:

| Tab | Contents |
|-----|----------|
| **By App** | One row per distinct app — app name, client ID, how many target users authorised it, which users, and scopes |
| **By User** | One row per target user — how many apps they have authorised, pipe-separated list of app names |
| **Raw Tokens** | Every raw token record (up to 500 rows; full data in CSV) |

### Section: Chat Spaces

Table of Chat Space memberships. Each row shows:
- The Space name and type (SPACE / GROUP_CHAT / DM)
- The GAM resource name (`spaces/XXXXXXXX`)
- Which target user (`TargetUser` column) this row came from

### Section: Chat Bot Activity

Two tabs:

| Tab | Contents |
|-----|----------|
| **Bot Summary** | One row per unique bot — total events, event types, which target users interacted, spaces involved, first/last seen |
| **All Bot Events** | Every individual bot event filtered for target users |

### Section: Output Files

A table listing every CSV produced, with the record count from this run.

---

## 13. Interpreting the Results

### Red flags — Connected Apps

| Finding | Risk | Recommended action |
|---------|------|--------------------|
| App with `anonymous: true` | High — unverified app | Review the OAuth scopes granted; revoke if not needed |
| App with `https://www.googleapis.com/auth/gmail.modify` or `.readonly` scope | High | App can read/modify all email; confirm it is intentional |
| App with `https://www.googleapis.com/auth/drive` scope | High | Full Drive access; confirm legitimate business use |
| App with `https://www.googleapis.com/auth/admin.directory` scope | Critical | Admin Directory access; investigate immediately |
| Same app authorised by multiple target users | Medium | Could indicate a phishing OAuth app |
| App with no recognisable name (`displayText` is a URL or ID) | Medium | May be a developer or test app |

### Revoking a token for a specific user and app

```powershell
# Revoke by client ID
gam user alice@corp.com delete token clientid <clientId>

# Revoke ALL tokens for a user (nuclear option)
gam user alice@corp.com delete tokens
```

### Red flags — Chat Bots

| Finding | Risk | Recommended action |
|---------|------|--------------------|
| Bot added to a sensitive space (e.g. Executive, Finance) | Medium | Confirm the bot is approved |
| Bot with no recognisable `AppName` (shows GCP resource ID only) | Medium | Look up the GCP project in the Admin Console |
| Bot with very high `TotalEvents` | Low/Medium | May indicate automation abuse |
| `app_removed` event — bot was removed | Informational | Note the date and who removed it |

---

## 14. Troubleshooting

### GAM not found

**Symptom:** Script shows the interactive GAM menu immediately.

**Fix:**
1. Choose `[1]` and paste the full path, e.g. `C:\GAM7\gam.exe`
2. Answer `y` to save the path — future runs will be instant.

OR pass `-GamPath` on every run:
```powershell
.\Get-TargetedUserReport.ps1 -Users "alice@corp.com" -GamPath "C:\GAM7\gam.exe"
```

---

### `403 insufficientPermissions` on token query

**Symptom:** `tokens_target_users.csv` is empty or contains only error lines.

**Fix:**
```powershell
gam oauth update
```
Re-authorise with a Super Admin account. Ensure **Admin SDK - Directory API** and
**Google OAuth Token** scopes are ticked.

---

### `403` on Chat Spaces

**Symptom:** `chat_spaces_target.csv` is empty.

**Fix:**
```powershell
gam oauth update
```
Enable the **Google Chat API - Admin** scope. This requires GAM7 or GAMADV-XTD3.

---

### Chat bot report times out

**Symptom:** `[WARN] Chat report timed out.`

**Fix:** Increase the timeout. For large tenants the Chat log pull can take 3–5 minutes:
```powershell
.\Get-TargetedUserReport.ps1 -UsersFile ".\targets.csv" -IncludeChatBots -DwdTimeoutSeconds 300
```

---

### No bot events found despite `-IncludeChatBots`

**Possible causes:**
1. No bots were active in the date range — try `-ChatDaysAgo 90`
2. The target users never interacted with any bot in Chat
3. The Chat audit log has a 48-hour data delay — results may not appear for recent events
4. GAM column names differ in your version — check `chat_report_raw.csv` (temporarily comment out the `Remove-Item` line in Section 3 to keep the raw file)

---

### HTML report is blank or partially rendered

**Cause:** Some corporate browsers block local HTML files.
**Fix:** Open the file in a different browser, or copy the output folder to a web server path.

---

### `Set-StrictMode` error during run

**Symptom:** Red error mentioning a variable or property not found.

**Fix:** This usually means GAM returned an unexpected column name. Open `tokens_target_users.csv`
and check what the column headers actually are. The script auto-detects `user`/`userEmail` and
`displayText`/`appName` — if your GAM version uses something else, pass it to the developers.

---

## 15. Scheduling & Automation

### Run on a schedule (Windows Task Scheduler)

Create a scheduled task that runs the script silently and saves output to a dated folder:

```powershell
# Save as: Run-TargetedAudit.ps1
$date    = Get-Date -f "yyyyMMdd"
$outDir  = "C:\AuditReports\Targeted_$date"
$users   = "alice@corp.com","bob@corp.com"

& "C:\GAM-Scripts\connected-Apps\Get-TargetedUserReport.ps1" `
    -Users $users `
    -GamPath "C:\GAM7\gam.exe" `
    -OutputDir $outDir `
    -IncludeChatSpaces `
    -IncludeChatBots `
    -ChatDaysAgo 7
```

In Task Scheduler:
- **Program:** `powershell.exe`
- **Arguments:** `-NonInteractive -ExecutionPolicy Bypass -File "C:\GAM-Scripts\Run-TargetedAudit.ps1"`
- **Run as:** A service account with GAM credentials configured

> **Note:** Remove or replace the `Read-Host` prompt at the end of the script when running unattended
> (comment out the last two lines).

### Email the report automatically

Append this to your wrapper script after the main script finishes:

```powershell
$report = "C:\AuditReports\Targeted_$date\TargetedReport.html"
Send-MailMessage `
    -To "security-team@corp.com" `
    -From "audit-noreply@corp.com" `
    -Subject "Targeted User Audit - $date" `
    -Body "See attached report." `
    -Attachments $report `
    -SmtpServer "smtp.corp.com"
```

---

## 16. Security & Compliance Notes

### Data sensitivity

The output files contain **PII and sensitive security data**:
- User email addresses
- Names of third-party apps with access to corporate data
- OAuth scopes (may reveal what data apps can read/modify)
- Chat Space names and membership

**Store output folders in a restricted location.** Delete them when no longer needed.

### GAM credentials

GAM stores its OAuth tokens in the GAM config directory (typically `%USERPROFILE%\.gam\` or `C:\GAMWork\`).
These tokens grant broad admin-level access to your Google Workspace tenant.

- **Restrict access** to the GAM config directory to the service account or admin that runs audits.
- **Rotate credentials** periodically via `gam oauth create`.
- **Do not share** the GAM config directory or its files.

### Principle of least privilege

If you do not need Chat data, omit `-IncludeChatSpaces` and `-IncludeChatBots`.
This avoids the need to authorise the Google Chat Admin scope.

### Audit trail

Every run creates a timestamped output folder. Keep these folders (or archive the HTML report)
as evidence of the audit having been performed.

---

*Last updated: June 2026 — covers `Get-TargetedUserReport.ps1` v1.0*

