# Windows Update Checker

A portable PowerShell WinForms GUI tool for verifying Windows Update status on local or remote servers. No external modules or installers required.

## Features

- **Local & remote scanning** — scan the local machine or any remote server by hostname/IP, using the Windows Update Agent (WUA) for authoritative, full results (installed **and** missing), with a WMI fallback
- **Install missing updates** — check missing updates in the grid and click **Install Selected**, on the **local machine or a remote server**
- **Remote reboot** — restart a remote server after patching, with a 60-second warning + **Abort** countdown, then it waits for the server to come back and automatically re-scans
- **Sign in once** — credentials entered for a remote target are cached and reused for install, reboot, and re-scans
- **Update source & policy indicator** — shows each machine's update source (Microsoft Update or WSUS) and any restrictions (Group Policy block, automatic updates off, updates paused), and translates cryptic Windows Update error codes into plain English
- **Automatic TrustedHosts handling** — remote IP targets are trusted for the session and your original WinRM TrustedHosts list is restored on exit
- **Color-coded results** — green (Installed), red (Missing), plus a separate Reboot Pending indicator
- **Filters** — filter by Status and Severity (exports respect the active filters)
- **Detail popup** — double-click any row for full update details + Microsoft KB link
- **Export** — save results as CSV or self-contained HTML report
- Portable — a single PowerShell script, no external modules or installers

## Requirements

- Windows PowerShell 5.1 or later
- The tool **self-elevates at startup** (UAC prompt) — scanning and installing updates requires Administrator. If elevation is declined, it shows a message and exits.
- Remote operations (scan, install, reboot) require **WinRM enabled and administrator rights on the target server**

## Usage

```powershell
# Right-click and Run with PowerShell (elevation is requested automatically)
.\WinUpdateChecker.ps1
```

Or from a PowerShell prompt:

```powershell
powershell -ExecutionPolicy Bypass -File .\WinUpdateChecker.ps1
```

## Remote targets

Remote scan, install, and reboot talk to the target over **WinRM**. Enable it on the target server (as Administrator):

```powershell
winrm quickconfig
```

How remote scanning/installing works: Windows blocks the Windows Update Agent API from being called remotely, so the tool registers a one-shot scheduled task that runs the WUA scan/install **locally on the target as SYSTEM**, then reads the results back. If that path is unavailable (e.g. Task Scheduler disabled, or SYSTEM has no path to Windows Update), it falls back to a WMI query of installed hotfixes and shows a "partial data" banner.

**Targeting by IP address:** when you connect to an IP (rather than a hostname) with explicit credentials, WinRM requires the IP to be in the client's TrustedHosts list. **The tool manages this automatically** — it adds the target to TrustedHosts when you scan it and restores your original TrustedHosts when the app closes. (The original list is saved to a marker file so that if the app is killed before it can restore, the next launch cleans up the leftover entry.) Pre-existing TrustedHosts entries are always preserved.

If you'd rather manage it yourself, you can add the host manually (`Set-Item WSMan:\localhost\Client\TrustedHosts -Value '<target-ip>' -Concatenate -Force`) or configure WinRM over HTTPS (port 5986) on the target, which the tool auto-detects.

## Notes

- Exported CSV and HTML files are excluded from Git via `.gitignore`
- The WUA `Search()` call can take 30-120 seconds depending on the machine and network
- Only missing updates found via WUA can be checked for install; changing filters clears the selection
- After an install completes, a results dialog is shown and a fresh scan runs automatically. Updates that need a reboot may still show as Missing until the machine restarts
- Updates that require user input during installation are skipped automatically
- **Remote reboot** forces a restart (`/f`) after a 60-second delay that warns logged-on users; you can Abort during the countdown. The tool then waits up to 30 minutes for the server to return (or **Stop Waiting**) before re-scanning
