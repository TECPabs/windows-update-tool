# Windows Update Checker

A portable PowerShell WinForms GUI tool for verifying Windows Update status on local or remote servers. No external modules or installers required.

## Features

- **Local & Remote scanning** — scan the local machine or any remote server by hostname/IP
- **Windows Update API (WUA COM)** — authoritative source, same data as the Windows Update UI
- **Fallback to WMI** — if WUA COM is unavailable, falls back to `Win32_QuickFixEngineering`
- **Color-coded results** — green (Installed), red (Missing), plus a separate Reboot Pending indicator
- **Install missing updates** — check missing updates in the grid and click **Install Selected** (local machine only)
- **Filters** — filter by Status and Severity (exports respect the active filters)
- **Detail popup** — double-click any row for full update details + Microsoft KB link
- **Export** — save results as CSV or self-contained HTML report

## Requirements

- Windows PowerShell 5.1 or later
- The tool **self-elevates at startup** (UAC prompt) — installing updates requires Administrator. If elevation is declined, it shows a message and exits.
- Remote scanning requires **WinRM** enabled on the target server

## Usage

```powershell
# Right-click and Run with PowerShell (elevation is requested automatically)
.\WinUpdateChecker.ps1
```

Or from a PowerShell prompt:

```powershell
powershell -ExecutionPolicy Bypass -File .\WinUpdateChecker.ps1
```

## Enabling WinRM on a Remote Server

If you get a "WinRM unreachable" error, run this on the target server as Administrator:

```powershell
winrm quickconfig
```

## Notes

- Exported CSV and HTML files are excluded from Git via `.gitignore`
- The WUA COM `Search()` call can take 30-120 seconds depending on the machine and network
- **Remote scans show installed hotfixes only** — Windows does not allow the WUA API to be called remotely over WinRM, so remote results come from WMI and cannot include missing updates
- **Installing** is available after a local scan only (remote install is planned for a future version). Only missing updates found via WUA can be checked; changing filters clears the selection
- After an install completes, a results dialog is shown and a fresh scan runs automatically. Updates that need a reboot may still show as Missing until the machine restarts
- Updates that require user input during installation are skipped automatically
