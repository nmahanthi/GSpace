# Fix: Dynamic GAM Path Resolution in `01_run_gam_exports.cmd`

## Problem

The script `01_run_gam_exports.cmd` previously contained a hardcoded absolute path to the GAM executable:

```bat
set GAM_PATH=C:\Users\v-nmahanthi\OneDrive - Microsoft\Documents\gam7\gam.exe
```

### Why This Was a Problem

| Issue | Detail |
|---|---|
| **Identity exposure** | The path contained a personal Windows username (`v-nmahanthi`), leaking internal user identity information into source control. |
| **Not portable** | The script failed on any machine where GAM was not installed at that exact path. |
| **Maintenance burden** | Every user or CI/CD environment required manual edits to the script to point to their own GAM installation. |

---

## Fix

The hardcoded line was replaced with a three-strategy dynamic resolution block that tries each option in priority order and exits with a clear error if none succeeds.

### Resolution Order

```
GAM_PATH env var  →  gam.cfg config file  →  system %PATH% search
```

#### Strategy 1 — `GAM_PATH` Environment Variable (highest priority)

If `GAM_PATH` is already defined in the shell or system environment, it is used immediately without any file or PATH lookup.

```bat
if defined GAM_PATH goto :verify_gam
```

**How to set it:**

- Temporarily (current shell session):
  ```bat
  set GAM_PATH=C:\tools\gam\gam.exe
  ```
- Permanently (system-wide, via PowerShell as Administrator):
  ```powershell
  [System.Environment]::SetEnvironmentVariable("GAM_PATH", "C:\tools\gam\gam.exe", "Machine")
  ```
- In a CI/CD pipeline: add `GAM_PATH` as a pipeline environment variable or secret.

---

#### Strategy 2 — `gam.cfg` Config File

If no environment variable is set, the script looks for a file named `gam.cfg` in the same directory as the script. It reads the first line matching `GAM_PATH=<value>`.

```bat
set GAM_CFG=%SCRIPT_DIR%gam.cfg
if exist "%GAM_CFG%" (
    for /f "usebackq tokens=1,* delims==" %%A in ("%GAM_CFG%") do (
        if /i "%%A"=="GAM_PATH" set "GAM_PATH=%%B"
    )
)
```

**Example `gam.cfg`:**
```
GAM_PATH=C:\tools\gam\gam.exe
```

> **Recommendation:** Add `gam.cfg` to `.gitignore` so machine-local paths are never committed to source control.

---

#### Strategy 3 — System `%PATH%` Search (lowest priority)

If neither of the above resolves GAM, the script searches the system `%PATH%` for `gam.exe` or `gam`.

```bat
for %%X in (gam.exe gam) do (
    set "_found=%%~$PATH:X"
    if defined _found (
        set "GAM_PATH=%%~$PATH:X"
        goto :verify_gam
    )
)
```

**How to configure:** Add the folder containing `gam.exe` to your Windows `PATH` environment variable.

---

#### Verification Step

Regardless of which strategy resolved `GAM_PATH`, the script verifies the file actually exists before proceeding:

```bat
:verify_gam
if not exist "%GAM_PATH%" (
    echo ERROR: GAM not found at "%GAM_PATH%"
    echo   Check your GAM_PATH environment variable or gam.cfg configuration.
    exit /b 1
)
echo [GAM] Using GAM at: %GAM_PATH%
```

---

#### Error if GAM Cannot Be Found

If all three strategies fail, the script exits with a clear, actionable message rather than silently proceeding:

```
ERROR: GAM executable not found.
  Set the GAM_PATH environment variable, add gam.exe to your PATH,
  or create a gam.cfg file in the script directory with:
    GAM_PATH=<full path to gam.exe>
```

---

## Quick Reference

| Scenario | Recommended approach |
|---|---|
| Local developer machine | Add GAM folder to `%PATH%`, or create `gam.cfg` |
| Shared / team repository | Use `GAM_PATH` env var per machine; never commit a hardcoded path |
| CI/CD pipeline | Set `GAM_PATH` as a pipeline secret or environment variable |
| Multiple GAM versions | Use `gam.cfg` per project to pin each project to its own GAM binary |
