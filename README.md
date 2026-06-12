# Windows Update Checker

A portable PowerShell WinForms GUI tool for verifying Windows Update status on local or remote servers. No external modules or installers required.

## Features

- **Local & Remote scanning** — scan the local machine or any remote server by hostname/IP
- **Windows Update API (WUA COM)** — authoritative source, same data as the Windows Update UI
- **Fallback to WMI** — if WUA COM is unavailable, falls back to `Win32_QuickFixEngineering`
- **Color-coded results** — green (Installed), red (Missing), orange (Pending Reboot)
- **Filters** — filter by Status and Severity
- **Detail popup** — double-click any row for full update details + Microsoft KB link
- **Export** — save results as CSV or self-contained HTML report

## Requirements

- Windows PowerShell 5.1 or later
- Run as **Administrator** for best results (WUA COM requires elevation)
- Remote scanning requires **WinRM** enabled on the target server

## Usage

```powershell
# Right-click and Run with PowerShell (as Administrator)
.\WinUpdateChecker.ps1
```

Or from an elevated PowerShell prompt:

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
