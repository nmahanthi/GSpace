# Targeted Google Tasks Scan — Runbook

**Version:** 1.0 | **Platform:** Windows PowerShell 5.1 or PowerShell 7+  
**Tool:** GAM7 (Google Workspace Admin Manager)

---

## Table of Contents

1. [Overview](#1-overview)
2. [Prerequisites](#2-prerequisites)
3. [File Inventory](#3-file-inventory)
4. [One-Time Setup](#4-one-time-setup)
   - 4.1 [Verify GAM is working](#41-verify-gam-is-working)
   - 4.2 [Edit the config file](#42-edit-the-config-file)
   - 4.3 [Add users to the user list](#43-add-users-to-the-user-list)
5. [Running the Scan](#5-running-the-scan)
   - 5.1 [Standard run](#51-standard-run)
   - 5.2 [Inline user list (no file)](#52-inline-user-list-no-file)
   - 5.3 [Include completed or hidden tasks](#53-include-completed-or-hidden-tasks)
   - 5.4 [Enable Chat space scan](#54-enable-chat-space-scan)
   - 5.5 [Safe interrupted run (checkpoint)](#55-safe-interrupted-run-checkpoint)
6. [Understanding the Output](#6-understanding-the-output)
   - 6.1 [Detail CSV columns](#61-detail-csv-columns)
   - 6.2 [Summary CSV columns](#62-summary-csv-columns)
   - 6.3 [SurfaceType values](#63-surfacetype-values)
   - 6.4 [Origin values](#64-origin-values)
7. [How the Scan Works](#7-how-the-scan-works)
8. [Configuration Reference](#8-configuration-reference)
9. [Troubleshooting](#9-troubleshooting)
10. [FAQ](#10-faq)

---

## 1. Overview

This toolkit scans a **selected list of Google Workspace users** (e.g. 10–50 users) for all Google Tasks activity, without touching the rest of the tenant (100,000+ users).

For each user it identifies:

| What | How |
|---|---|
| Tasks the user created themselves | GAM `print tasks` |
| Tasks assigned to the user from a **Google Doc** | REST API (`showAssigned=true`) |
| Tasks assigned to the user from a **Chat Space** | REST API + Chat message correlation |
| Docs **action-item comments** assigned by the user | Drive comment scan |
| Orphaned / unassigned Chat space tasks | Chat message thread analysis |

Two output CSV files are produced per run — a full detail file and a sharer-centric summary.

---

## 2. Prerequisites

| Requirement | Details |
|---|---|
| **GAM7** installed and authorised | https://github.com/GAM-team/GAM/wiki/How-to-Install-Advanced-GAM |
| **Google Workspace super-admin** account used for GAM | Must have domain-wide delegation |
| **Windows PowerShell 5.1** or **PowerShell 7+** | Both supported. PS7 (`pwsh`) additionally enables the REST-based assigned-task fetch. |
| **GAM OAuth scopes** | `tasks.readonly`, `drive.readonly`, `chat.spaces.readonly`, `chat.messages.readonly` |
| **Script execution policy** | `RemoteSigned` or `Bypass` — see §4.1 |

> **Note on PowerShell 7 (pwsh):** If `pwsh` is installed alongside Windows PowerShell, the script automatically uses it for the parallel REST fetch of Docs/Chat-assigned tasks. If only Windows PowerShell 5 is available, the REST pass is skipped and only GAM-visible tasks are returned. No configuration is needed — detection is automatic.

---

## 3. File Inventory

All five files must remain in the **same folder**.

| File | Role | Edit? |
|---|---|---|
| `Invoke-TargetedTaskScan.ps1` | **Entry point — run this** | No |
| `TaskScan.config.psd1` | **Configuration — edit once** | ✅ Yes |
| `TargetUsers.txt` | **User list — edit per run** | ✅ Yes |
| `Get-GoogleTasksWithCreator.ps1` | Engine — called automatically | No |
| `_Get-AssignedTasks.ps1` | REST helper — called automatically | No |

---

## 4. One-Time Setup

### 4.1 Verify GAM is working

Open **PowerShell as Administrator** and run:

```powershell
gam version
```

Expected output: GAM version line, e.g. `GAM 6.x.x`.  
If this fails, locate `gam.exe` and note its full path (e.g. `C:\GAM7\gam.exe`) — you will need it in §4.2.

Also check the execution policy:

```powershell
Get-ExecutionPolicy
```

If it returns `Restricted`, fix it:

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

### 4.2 Edit the config file

Open `TaskScan.config.psd1` in Notepad or any text editor.

```powershell
notepad ".\TaskScan.config.psd1"
```

The file is self-documenting. Key settings:

```powershell
@{
    # Set this only if GAM auto-detection fails (see §9 Troubleshooting).
    # Leave blank to let the script find GAM automatically.
    GamPath = ''                          # e.g. 'C:\GAM7\gam.exe'

    # Where to write output CSVs.
    # Leave blank to save in the same folder as the script.
    OutputDir = ''                        # e.g. 'C:\Scans\Output'

    # Leave blank — defaults to TargetUsers.txt in the script folder.
    UsersFile = ''

    # Scan options — change to $true to enable
    IncludeCompleted   = $false
    IncludeHidden      = $false
    IncludeDeleted     = $false
    ScanSpaces         = $false           # expensive — see §5.4
    SkipDocCommentScan = $false
}
```

Save the file. **This is done once.** You do not need to edit the config again unless your environment changes.

---

### 4.3 Add users to the user list

Open `TargetUsers.txt` and replace the placeholder lines with real email addresses — one per line:

```
# Lines starting with # are comments — ignored by the script
alice@yourdomain.com
bob@yourdomain.com
carol@yourdomain.com
dave@yourdomain.com
```

Save the file. **Update this file for each new batch of users.**

---

## 5. Running the Scan

Open **PowerShell** (does not need to be Administrator for the scan itself), navigate to the script folder, and run:

### 5.1 Standard run

```powershell
cd "C:\Path\To\Task Details"
.\Invoke-TargetedTaskScan.ps1
```

The script will print a pre-flight summary before making any API calls:

```
Config loaded : C:\...\TaskScan.config.psd1
GAM detected  : C:\GAM7\gam.exe

=== Targeted Google Tasks Scan ===
  Users to scan : 9
  Doc scan      : ON (default for targeted)
  Space scan    : SKIPPED (use -ScanSpaces to enable)

  Users:
    - alice@yourdomain.com
    - bob@yourdomain.com
    ...
```

When complete:

```
=== Done ===
  Detail CSV  : C:\...\Tasks_Targeted_20260614_143022.csv
  Summary CSV : C:\...\Tasks_Targeted_20260614_143022_Summary.csv
```

---

### 5.2 Inline user list (no file)

```powershell
.\Invoke-TargetedTaskScan.ps1 -Users alice@corp.com,bob@corp.com,carol@corp.com
```

---

### 5.3 Include completed or hidden tasks

```powershell
# Include completed tasks only
.\Invoke-TargetedTaskScan.ps1 -IncludeCompleted

# Include completed + hidden
.\Invoke-TargetedTaskScan.ps1 -IncludeCompleted -IncludeHidden

# Include everything (completed + hidden + deleted)
.\Invoke-TargetedTaskScan.ps1 -IncludeCompleted -IncludeHidden -IncludeDeleted
```

> These flags can also be permanently enabled in `TaskScan.config.psd1`.

---

### 5.4 Enable Chat space scan

The tenant-wide Chat space scan is **off by default** because it enumerates every Chat space visible to the first user in your list — this can be thousands of spaces and adds significant time.

Enable it only when you need to find orphaned / unassigned Chat tasks:

```powershell
.\Invoke-TargetedTaskScan.ps1 -ScanSpaces
```

---

### 5.5 Safe interrupted run (checkpoint)

For larger user lists, use a checkpoint file. If the script is interrupted, already-processed users are preserved:

```powershell
.\Invoke-TargetedTaskScan.ps1 -CheckpointCsv .\checkpoint.csv
```

The checkpoint CSV is updated after each user completes.

---

## 6. Understanding the Output

Two CSV files are written per run, both timestamped so repeated runs never overwrite each other.

### 6.1 Detail CSV — `Tasks_Targeted_YYYYMMDD_HHMMSS.csv`

One row per task. All tasks for all scanned users.

| Column | Description |
|---|---|
| `AssigneeEmail` | The user who owns / is assigned this task |
| `TaskListId` | Internal ID of the Google Tasks list |
| `TaskListTitle` | Display name of the task list (e.g. `My Tasks`) |
| `TaskId` | Unique task identifier |
| `Title` | Task title / description |
| `Status` | `needsAction`, `completed`, `unassigned` |
| `Due` | Due date (ISO 8601), if set |
| `Completed` | Completion timestamp, if completed |
| `Updated` | Last-modified timestamp |
| `WebViewLink` | Direct URL to the task in Google Tasks |
| `SurfaceType` | Where the task originated — see §6.3 |
| `SourceRef` | Drive file ID (DOCUMENT) or Chat space name (SPACE) |
| `SourceLink` | Deep link to the source document or chat |
| `CreatorEmail` | Email of who created / assigned the task |
| `CreatorName` | Display name of the creator (Docs tasks only) |
| `Shared` | `True` if creator ≠ assignee |
| `Origin` | Which scan path found this row — see §6.4 |
| `Notes` | Human-readable context (doc name, space name, caveats) |

---

### 6.2 Summary CSV — `Tasks_Targeted_YYYYMMDD_HHMMSS_Summary.csv`

One row per **creator** who has shared at least one task. Useful for identifying power users who assign tasks to others.

| Column | Description |
|---|---|
| `CreatorEmail` | The person who created/assigned tasks |
| `SharedCount` | Number of tasks they assigned to others |
| `AssigneeEmails` | Semicolon-separated list of people they assigned to |
| `Titles` | Semicolon-separated list of task titles |
| `SurfaceTypes` | Which surfaces were used (DOCUMENT; SPACE) |

> The Summary CSV is only written when at least one shared task exists.

---

### 6.3 SurfaceType values

| Value | Meaning |
|---|---|
| `SELF` | User created the task themselves (no assignmentInfo) |
| `DOCUMENT` | Task was assigned from a Google Doc, Sheet, or Slide comment |
| `SPACE` | Task was assigned from a Google Chat space |

---

### 6.4 Origin values

| Value | Meaning |
|---|---|
| `GAM` | Found via GAM `print tasks` — standard user-created tasks |
| `REST` | Found via direct REST API with `showAssigned=true` — Docs/Chat-assigned tasks not visible to GAM CLI |
| `DOCSCAN` | Found via Drive comment scan — action-item comment in a Doc owned by the scanned user |
| `SPACESCAN` | Found via Chat space message scan — orphaned or unassigned Chat task |

---

## 7. How the Scan Works

The engine (`Get-GoogleTasksWithCreator.ps1`) runs four scan passes per user:

```
┌─────────────────────────────────────────────────────────────┐
│  For each user in TargetUsers.txt                           │
│                                                             │
│  Pass 1 — GAM bulk fetch (tasks + tasklists)                │
│    gam file <users> print tasks formatjson                  │
│    → Captures self-created and standard assigned tasks      │
│    → Origin: GAM                                            │
│                                                             │
│  Pass 2 — REST assigned-task fetch (PS7 parallel)           │
│    Tasks API: GET /lists/{id}/tasks?showAssigned=true        │
│    → Captures Docs/Chat tasks invisible to GAM CLI          │
│    → Origin: REST                                           │
│                                                             │
│  Pass 3 — Drive comment scan (per owned Doc)                │
│    gam user print filelist + print filecomments             │
│    → Captures Docs action items not yet mirrored to Tasks   │
│    → Origin: DOCSCAN                                        │
│                                                             │
│  Pass 4 — Chat space scan (optional, -ScanSpaces)           │
│    gam user print chatspaces + chatmessages                 │
│    → Captures orphaned/unassigned Chat space tasks          │
│    → Origin: SPACESCAN                                      │
└─────────────────────────────────────────────────────────────┘
```

**Deduplication:** A `HashSet` of task IDs prevents the same task appearing twice across passes (e.g. a task found by both GAM and REST).

**Bulk prefetch:** For efficiency, Passes 1 and 2 are executed as single GAM/pwsh processes for all users combined — not once per user. This means scanning 10 users takes roughly the same time as scanning 1.

**Creator derivation:** The Google Tasks API does not expose a creator field.  
- `SELF` tasks: creator = assignee (same user).  
- `DOCUMENT` tasks: creator = owner of the linked Drive file (via `show fileinfo`).  
- `SPACE` tasks: creator = Chat message sender correlated by timestamp (±10 seconds).

---

## 8. Configuration Reference

### `TaskScan.config.psd1` — all settings

| Key | Type | Default | Description |
|---|---|---|---|
| `GamPath` | String | `''` (auto-detect) | Full path to `gam.exe` |
| `UsersFile` | String | `''` (→ `TargetUsers.txt`) | Path to user list file |
| `OutputDir` | String | `''` (→ script folder) | Folder for output CSVs |
| `IncludeCompleted` | Boolean | `$false` | Include completed tasks |
| `IncludeHidden` | Boolean | `$false` | Include hidden tasks |
| `IncludeDeleted` | Boolean | `$false` | Include deleted tasks |
| `ScanSpaces` | Boolean | `$false` | Enable Chat space scan |
| `SkipDocCommentScan` | Boolean | `$false` | Disable Drive comment scan |

### GAM auto-detection order

When `GamPath` is blank, the script probes these locations in order:

1. System `PATH` (`Get-Command gam`)
2. `C:\GAM7\gam.exe`
3. `C:\GAM\gam.exe`
4. `%LOCALAPPDATA%\GAM7\gam.exe`
5. `%LOCALAPPDATA%\GAM\gam.exe`
6. `%USERPROFILE%\GAM7\gam.exe`
7. `%USERPROFILE%\GAM\gam.exe`
8. `%ProgramData%\GAM7\gam.exe`
9. `%ProgramData%\GAM\gam.exe`
10. `%ProgramFiles%\GAM7\gam.exe`
11. `%ProgramFiles%\GAM\gam.exe`
12. `%ProgramFiles(x86)%\GAM7\gam.exe`
13. `~\.gam\gampath` hint file (written by `gam config`)

### Parameter precedence (highest wins)

```
Command-line parameter  ›  Config file value  ›  Auto-detect / built-in default
```

---

## 9. Troubleshooting

### GAM not found

**Symptom:**
```
Cannot find the GAM executable. Tried PATH, common install folders...
```
**Fix:** Find where `gam.exe` is installed, then either:
- Set `GamPath = 'C:\GAM7\gam.exe'` in `TaskScan.config.psd1`, **or**
- Pass it on the command line: `.\Invoke-TargetedTaskScan.ps1 -GamPath 'C:\GAM7\gam.exe'`

---

### Script is not digitally signed / execution policy error

**Symptom:**
```
File cannot be loaded because running scripts is disabled on this system.
```
**Fix:**
```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

### GAM exits with code 60 (permission denied for a user)

**Symptom:** Yellow `WARNING` lines like `GAM failed (60): user alice@...`  
**Cause:** The GAM service account does not have delegation rights for that user, or the user's account is suspended.  
**Fix:** Code 60 is silently suppressed — those users produce zero rows. Check that the GAM service account has the correct OAuth scopes and domain-wide delegation in the Google Admin console.

---

### "pwsh not available" warning — REST pass skipped

**Symptom:**
```
WARNING: pwsh (PowerShell 7+) or _Get-AssignedTasks.ps1 not available;
Docs/Chat-assigned tasks will not be fetched.
```
**Impact:** Tasks that were assigned from Google Docs or Chat and not yet visible to the GAM CLI will be missing from the output. Self-created tasks are unaffected.  
**Fix:** Install PowerShell 7: https://aka.ms/powershell  
Once installed, re-run the script — detection is automatic, no configuration needed.

---

### Empty output / zero rows

**Possible causes:**
- The users in `TargetUsers.txt` have no tasks at all (correct result).
- The email addresses in `TargetUsers.txt` do not match the tenant domain exactly (check for typos).
- GAM is not authorised for the `tasks.readonly` scope. Run:
  ```powershell
  gam user alice@corp.com print tasklists
  ```
  If this fails, re-authorise GAM scopes.

---

### Output CSV is not produced

**Symptom:** Script finishes but the CSV file is missing.  
**Cause:** An unhandled error occurred before the export step.  
**Fix:** Re-run with verbose output:
```powershell
.\Invoke-TargetedTaskScan.ps1 -Verbose
```
Look for the first red `ERROR` line.

---

### Scan is slow

The Doc-comment scan (Pass 3) is the heaviest path — it calls `gam print filecomments` once per Drive file owned by each user. For users with many Docs files this can take several minutes per user.

**To skip it:**
```powershell
.\Invoke-TargetedTaskScan.ps1 -SkipDocCommentScan
```
Or set `SkipDocCommentScan = $true` in the config file.

---

## 10. FAQ

**Q: Can I scan all users in the tenant?**  
A: Yes, but use `Get-GoogleTasksWithCreator.ps1 -Users all` directly — not this targeted wrapper. For 100,000 users the Doc-comment scan is automatically disabled.

**Q: Will running the script affect users' tasks?**  
A: No. All API calls are read-only (`tasks.readonly`, `drive.readonly`). Nothing is modified.

**Q: The `CreatorEmail` column is blank for some Chat tasks — why?**  
A: The Google Tasks API does not expose the creator for Chat-assigned tasks. The script attempts to correlate the creator by matching Chat message timestamps (±10 seconds). If no matching Chat message is found, `CreatorEmail` is left blank and the `Notes` column explains why.

**Q: Some tasks appear in both GAM and REST origins — is that a bug?**  
A: No — deduplication by task ID ensures each task appears exactly once. The `Origin` column shows which pass found it first.

**Q: Can I run multiple batches and combine the CSVs?**  
A: Yes. Each run produces a timestamped file. Import all files into Excel, append the rows, and remove duplicates on the `TaskId` column.

**Q: The Summary CSV was not produced.**  
A: The Summary CSV is only written when at least one `Shared = True` row exists. If all scanned users only have self-created tasks, no summary is produced — this is expected.

**Q: How do I add more users mid-way through without re-scanning the ones already done?**  
A: Use `-CheckpointCsv` on the first run. Then add the new users to `TargetUsers.txt` and run again — results accumulate in the checkpoint file. Alternatively, run the two batches separately and merge the CSVs.
