# GAM Work — Complete Documentation

This document explains, in detail, every piece of [GAM](https://github.com/GAM-team/GAM)
(Google Workspace Admin CLI) work built for the GSites Assessment Toolkit:
what each script does, the exact GAM commands used and why, the problems
each one solves, and how they fit together. For general pipeline usage see
`CUSTOMER_SETUP.md`; for the non-GAM steps (crawling/scoring) see
`TOOLKIT_DOCUMENTATION.md`.

---

## 1. What GAM Is Used For

GAM is a command-line tool that talks to the Google Workspace Admin SDK and
Drive API on behalf of an admin/user account. In this toolkit it is the
**only** component that talks to Google APIs for inventory purposes — it
answers three questions per Google Site:

1. **Where does it live and who owns it?** (inventory)
2. **Who can access it, and how?** (permissions)
3. *(Shared Drive variant only)* **What are the drive-level sharing rules
   and who administers the drive?** (drive metadata)

No OAuth app registration, service account, or Sites API scope is needed
for any GAM step — GAM handles auth via its own pre-configured OAuth/DWD
credentials tied to a Workspace admin.

---

## 2. GAM Path Resolution (shared by every `.cmd` script)

Every GAM script (`01_run_gam_exports.cmd`, `01d_run_gam_exports_shareddrive.cmd`,
`01e_list_shareddrive_metadata.cmd`) resolves the `gam.exe` binary the same way,
in priority order:

1. **`GAM_PATH` environment variable** — if already set (by the shell, CI,
   or the calling PowerShell orchestrator), used immediately.
2. **`gam.cfg` file** next to the script — a one-line `GAM_PATH=<path>` file,
   kept out of source control (machine-local).
3. **System `PATH` search** — looks for `gam.exe` or `gam` on `PATH`.

If none resolve, the script fails fast with an actionable error instead of
silently breaking. This exists because the original script had a
**hardcoded personal path** (`C:\Users\v-nmahanthi\...\gam.exe`) baked into
source control — leaking a username and breaking on every other machine.
See `GAM_PATH_FIX.md` for the full before/after and rationale.

Every script also verifies the resolved path actually exists on disk
(`:verify_gam`) before running any real GAM command.

---

## 3. `01_run_gam_exports.cmd` — My Drive GSites Export (main pipeline, Step 1)

**Purpose:** inventory every Google Site owned in **My Drive** across the
domain (or a restricted set of users), plus its sharing/permission data.

**Scan target** (`GAM_USER_TARGET`), chosen by the calling orchestrator via
env vars:
- `GAM_TARGET_FILE` set → `csv "<file>" gam user "~Email"` — scans only the
  specific user emails listed in that CSV (built from `-TargetUsersCsv` or
  owners discovered in `-SelectedSitesCsv`).
- Otherwise → `all users` — GAM first fetches every user in the Workspace
  domain, then scans each one's My Drive in turn.

**Sites query filter** (`SITES_QUERY`), always:
```
mimeType='application/vnd.google-apps.site' and trashed=false
```
optionally AND'd with a `(name='...' or name='...')` clause
(`GAM_SITES_FILTER`, built by the orchestrator's `Build-GamNameFilter` from
`-SelectedSitesCsv`) to shrink scope for large tenants — GAM only looks for
those exact Drive file names instead of scanning everything.

**Three exports, in order:**

1. **`GSites_Inventory_Min.csv`** — sanity check. Two forms depending on
   whether a name filter is active:
   ```
   gam <target> print filelist query "<query>" fields id,name,mimetype
   ```
   or (no filter) a broader gsite-flagged listing:
   ```
   gam <target> print filelist fields id,name,mimetype filepath showmimetype gsite
   ```
2. **`GSites_Inventory_Detailed.csv`** — full metadata per site: id, name,
   `webviewlink` (the Drive **edit URL** — used later for crawling),
   created/modified time, owners, size, sharing flags, and `capabilities.*`
   flags (canShare/canEdit/canDownload/canCopy/canRemoveChildren/canDelete).
3. **`GSites_Permissions.csv`** — one row per (site, grantee) via
   `oneitemperrow`, containing `basicpermissions` (type/role/emailAddress
   per grantee) plus sharing-policy flags
   (`copyrequireswriterpermission`, `viewerscancopycontent`,
   `writerscanshare`, `inheritedpermissionsdisabled`).

**Tuning:** `GAM_NUM_THREADS` (env var, set from the orchestrator's
`-GamThreads`, default 10) controls `num_threads` for GAM's `multiprocess`
mode — how many worker processes scan users in parallel. Higher = faster
but more memory and higher risk of `BrokenPipeError`/`BufferError` on
Windows with large user counts (50-100+); see §6.3.

**Post-processing:** after all three exports succeed, a PowerShell
one-liner strips numeric array-index infixes from CSV headers (e.g.
`owners.0.emailAddress` → `owners.emailAddress`) so downstream scripts can
reference fields by a stable name regardless of how many array entries GAM
happened to emit for a given row.

---

## 4. `01d_run_gam_exports_shareddrive.cmd` — Shared Drive GSites Export

**Purpose:** a **separate, opt-in** utility to inventory Google Sites hosted
on a specific, known list of Shared Drives. The main pipeline (`01_`)
deliberately does **not** touch Shared Drives (see §6.2) — this script is
how a customer can still assess Shared-Drive-hosted sites, on demand,
without reintroducing the problems that caused `01_` to exclude them.

**Usage:**
```
01d_run_gam_exports_shareddrive.cmd <SharedDriveIDs.csv> [DefaultScanningUserEmail]
```

**Input CSV** — the key design decision here — has two required columns:

```csv
driveId,account
0AbCDeFGhIJKLmnUK9PVA,owner1@rocheua.com
0XyzTeamDriveIdHere123,owner2@rocheua.com
```

- `driveId` — the Shared Drive's ID.
- `account` — the Workspace user account GAM should impersonate to scan
  **that specific drive**. This can be the drive's owner, an organizer, a
  member, or a Super Admin.

**Why per-row accounts instead of one admin scanning everything:** Shared
Drives in a real tenant are often governed by different business units/
owners, and a single Super Admin may not have (or be granted) access to
every drive. By pairing each `driveId` with its own `account`, GAM can
impersonate the right person per drive — no single super-user needs
visibility into all of them. A `DefaultScanningUserEmail` (2nd CLI arg, or
`GAM_ADMIN_USER` env var) can be supplied as a fallback for any row that
leaves `account` blank.

**Validation step (PowerShell inline):** before running any GAM command,
the input CSV is validated to confirm both `driveId` and `account` columns
exist, and that every row ends up with a non-blank effective account
(filling blanks from the default if given). Missing accounts with no
default fail the whole run immediately with the exact `driveId`s at fault,
rather than partially succeeding. The validated result is written to a
temp file `_SharedDriveIDs_effective.csv`, which is what GAM's
`multiprocess csv ...` substitution actually iterates over.

**Two exports:**

1. **`GSites_SharedDrive_Inventory.csv`**:
   ```
   gam multiprocess csv "<effective.csv>" gam user "~account" print filelist
       select teamdriveid "~driveId" query "<SITES_QUERY>"
       fields id,name,mimetype,description,webviewlink,createdtime,modifiedtime,
              owners,lastmodifyinguser,shared,driveid,size,hasaugmentedpermissions,
              capabilities.*,spaces,thumbnaillink
       showdrivename
   ```
   `select teamdriveid "~driveId"` scopes the file listing to exactly that
   one Shared Drive per row (as opposed to `corpora alldrives`, which
   expands **every** drive a user can see — the behavior that caused
   per-member duplication in the old design, §6.1). `showdrivename` adds
   the human-readable drive name.

2. **`GSites_SharedDrive_Permissions.csv`** — same per-drive scoping, plus:
   - `permissiondetails` and `showshareddrivepermissions` — resolves
     whether each grantee's access on a Shared Drive item is **direct**
     (granted on the file itself) or **inherited** (granted at the drive/
     folder level) — a distinction that doesn't exist for My Drive files
     and is only meaningful/available for Shared Drive content.

**Threading default:** `GAM_NUM_THREADS` defaults to **5** here (vs. 10 in
`01_`), since each worker now expands a full Shared Drive's file tree
(potentially large) rather than one user's My Drive filtered to sites only.

**Post-processing:** same CSV header normalization as `01_`, scoped to
`GSites_SharedDrive_*.csv`.

---

## 5. `01e_list_shareddrive_metadata.cmd` — Shared Drive-Level Metadata

**Purpose:** companion to `01d` — where `01d` reports on the **Sites** living
on a drive, `01e` reports on the **drive itself**: its sharing restrictions
and who administers it. Same input CSV/validation/account-resolution logic
as `01d` (see §4) since it targets the same list of drives.

**Two exports:**

1. **`GSites_SharedDrive_Settings.csv`** — one row per drive, sourced from:
   ```
   gam config num_threads 1 redirect stdout "<file>.jsonl"
       multiprocess csv "<effective.csv>" gam user "~account"
       info shareddrive "~driveId" fields id,name,createdtime,restrictions
       formatjson
   ```
   - Forced to `num_threads 1` (unlike every other export in the toolkit):
     the output here is one JSON object per drive written to a shared
     stdout stream. Concurrent workers writing JSON objects to the same
     stream risks interleaving fragments from different objects together,
     corrupting the JSON. Single-threading avoids that entirely.
   - GAM's `formatjson` pretty-prints each object across multiple lines and
     concatenates them back-to-back with no separator or enclosing array.
     A PowerShell post-processing step inserts a `},{` separator at every
     `}{` boundary, wraps the whole thing in `[...]`, and parses it with
     `ConvertFrom-Json`. If parsing fails, it warns and continues instead
     of crashing the whole export.
   - Extracted fields per drive: `id`, `name`, `createdTime`, and the
     `restrictions` sub-object flattened out —
     `adminManagedRestrictions`, `domainUsersOnly`, `driveMembersOnly`,
     `copyRequiresWriterPermission`,
     `sharingFoldersRequiresOrganizerPermission`.

2. **`GSites_SharedDrive_Organizers.csv`** — one row per (drive, organizer):
   ```
   gam multiprocess csv "<effective.csv>" gam user "~account"
       print shareddriveorganizers shareddrives "~driveId"
       includetypes user,group oneorganizer false
   ```
   Lists drive-level Manager-role principals (users and groups) per drive.
   Note: full restriction visibility from Step 1 above requires the scanning
   `account` to itself be a Manager/Organizer of that drive.

**Post-processing:** header normalization applied to
`GSites_SharedDrive_Organizers.csv` only (the Settings CSV is built directly
by the PowerShell JSON conversion, not raw GAM CSV output, so it has no
array-index artifacts to strip).

---

## 6. Key Design Decisions & Problems Solved

### 6.1 Per-user duplication from `corpora alldrives`

**History:** Shared Drive support was originally bolted onto the *main*
pipeline (`01_run_gam_exports.cmd`) by adding `corpora alldrives` to its
`print filelist` calls (commits `c3b4e03`, `bf76826`). This told GAM's file
listing to search **every Shared Drive a user has access to**, in addition
to their My Drive. Combined with `all users` scanning, the effect was: for
every domain user, GAM re-listed every Shared-Drive-hosted site visible to
them. A site shared with 50 members was emitted **50 times** — once per
member — massively inflating inventory/permission row counts and causing
the crawler and scorer to reprocess the same site dozens of times.

**First fix (`feb21c7`):** `Deduplicate-GamExports`, added to
`Run-FullAssessment.ps1`, collapsed rows post-export by `id` (inventory) or
`id`+`permission.id` (permissions), preferring the row logged under the
site's actual owner. This fixed correctness but not the underlying cost —
GAM still had to expand every Shared Drive per user during the scan itself.

### 6.2 Shared Drive exclusion from the main pipeline

**Decision (`5aceaaa`):** per customer requirement, Shared Drive data was
removed entirely from `01_run_gam_exports.cmd` — not just deduplicated
after the fact. `corpora alldrives` was removed from all three
`print filelist` commands, `showshareddrivepermissions` was removed from
the permissions export, and the now-meaningless `driveid`/`drivename`
fields were dropped. Effects:
- Shared-Drive-hosted sites **no longer appear at all** in the main
  pipeline's output.
- Eliminates the memory blowup previously observed (~97% usage) from
  workers expanding full Shared Drive trees per user.
- Removes the per-user duplication at the source, rather than papering
  over it after export.
- `Deduplicate-GamExports` remains in `Run-FullAssessment.ps1` as a safety
  net for any residual duplication (e.g. an externally supplied
  `-InventoryCsv`), but should rarely find anything to collapse now.

### 6.3 Configurable GAM thread count

**Problem (`832d59a`):** GAM's own default `num_threads` is 5; the toolkit
originally hardcoded a higher value for speed. On Windows, each thread is a
full spawned Python process, and scanning 50-100+ users at a high thread
count was observed to crash GAM mid-export with `BrokenPipeError` /
`BufferError` inside its multiprocessing pool — losing all progress.

**Fix:** thread count is now a first-class parameter end-to-end:
`-GamThreads` (orchestrator) → `GAM_NUM_THREADS` env var → `num_threads`
in every GAM command. Default is 10 for the main pipeline (`01_`) and 5 for
the Shared Drive scripts (`01d`/`01e`, since each worker there expands a
whole drive tree). Lowering it (e.g. to 5) trades speed for stability on
large tenants.

### 6.4 Per-drive scanning accounts (no single all-access admin required)

**Problem (`ba77201`):** an earlier iteration of `01d`/`01e` took one
`AdminUser` for the whole run, requiring a single account with access to
*every* targeted Shared Drive — often unrealistic in real tenants where
different business units own different drives independently.

**Fix:** switched the input format to `driveId,account` pairs (§4), letting
each drive be scanned by whichever user actually has access to it. A
`DefaultScanningUserEmail` fallback covers rows that don't need a specific
override.

---

## 7. Command Reference — Where Each GAM Verb Is Used

| GAM verb / flag | Script(s) | Purpose |
|---|---|---|
| `print filelist` | `01_`, `01d` | List Drive files (Sites) matching a query |
| `query "mimeType='application/vnd.google-apps.site' and trashed=false"` | `01_`, `01d` | Restrict listing to non-trashed Google Sites only |
| `select teamdriveid "~driveId"` | `01d` | Scope listing to one specific Shared Drive (no cross-drive expansion) |
| `showshareddrivepermissions` / `permissiondetails` | `01d` | Resolve direct vs. inherited permission on Shared Drive items |
| `showdrivename` | `01d` | Include the human-readable Shared Drive name |
| `oneitemperrow` | `01_`, `01d` | One CSV row per (item, permission) instead of nested columns |
| `info shareddrive "~driveId" ... formatjson` | `01e` | Fetch a single Shared Drive's metadata/restrictions as JSON |
| `print shareddriveorganizers` | `01e` | List Manager-role principals (users/groups) per Shared Drive |
| `multiprocess csv "<file>" gam user "~<col>"` | `01_`, `01d`, `01e` | Fan out a GAM command across every row of a CSV, substituting a per-row value |
| `config auto_batch_min 1 num_threads N` | `01_`, `01d`, `01e` | Control export parallelism (see §6.3) |
| ~~`corpora alldrives`~~ | *(removed, §6.2)* | Previously expanded every Shared Drive a user could see — root cause of §6.1 |

---

## 8. Output Files Summary

| File | Produced by | Key columns |
|---|---|---|
| `GSites_Inventory_Min.csv` | `01_` | id, name, mimetype |
| `GSites_Inventory_Detailed.csv` | `01_` | id, name, webviewlink, owners, size, capabilities.* |
| `GSites_Permissions.csv` | `01_` | id, basicpermissions (type/role/emailAddress per grantee) |
| `GSites_SharedDrive_Inventory.csv` | `01d` | id, name, webviewlink, owners, driveid, hasaugmentedpermissions |
| `GSites_SharedDrive_Permissions.csv` | `01d` | id, basicpermissions, permissiondetails (direct/inherited) |
| `GSites_SharedDrive_Settings.csv` | `01e` | id, name, domainUsersOnly, driveMembersOnly, adminManagedRestrictions |
| `GSites_SharedDrive_Organizers.csv` | `01e` | driveId, organizer email/group |
| `_SharedDriveIDs_effective.csv` | `01d`/`01e` (temp) | driveId + resolved (default-filled) account per row |
