# Windows Update Checker — Project Roadmap

A portable PowerShell 5.1 WinForms tool for scanning, installing, and managing Windows updates on local and remote machines. This document tracks where the project is and where it's going.

_Last updated: 2026-06-13 — current release: **v1.0.0-beta.3**_

## Current state (shipped)

- **Local & remote scanning** via the Windows Update Agent (WUA), with a WMI fallback.
- **Install missing updates** on local or remote machines.
- **Remote reboot** with a 60s warning/abort countdown, down-then-up wait, and auto-rescan.
- Remote operations run WUA locally on the target via a one-shot **SYSTEM scheduled task** (the supported way around Windows' block on remote update API calls).
- Per-target **credential caching** (sign in once), color-coded results, status/severity filters, detail popup, CSV/HTML export, self-elevation.
- **Automatic TrustedHosts management** — remote IP targets are added to WinRM TrustedHosts on use and the original list is restored on exit (crash-resilient via a marker file).
- **Windows Update policy detection & error translation** — an always-on indicator shows each machine's update source (Microsoft Update / WSUS) and restrictions (policy block, no-auto-update, paused), and cryptic WUA error codes are translated to plain English in banners and failure messages.

## Design principles (keep future work consistent)

- **Portable**: a single `.ps1` script — no external modules or installers. Any persisted config (e.g. exclusion lists, logs) should be optional sibling files, not a required dependency.
- **Windows PowerShell 5.1 / .NET Framework** compatible.
- **Async via runspace + WinForms Timer**; background scripts are self-contained `[scriptblock]::Create('...')` literals (no apostrophes inside). `$script:ActiveOp` enforces mutual exclusion.
- **Remote = WinRM**, reusing the existing transport/credential/SSL plumbing and the SYSTEM-scheduled-task mechanism.
- Destructive/disruptive actions get explicit confirmation.

---

## Planned features

### 1. Fleet / multi-target operations  _(largest lift; highest value)_

Move from one-machine-at-a-time to managing many machines at once — the core MSP use case.

**Scope**
- **Target input**: multiline list / comma-separated, plus **CSV import**; later, optional Active Directory OU query.
- **Per-machine results view**: a master grid of machines (Target · Status · Installed/Missing counts · Reboot Pending · last result/error), with drill-down into a single machine's update grid (the current view becomes the detail view).
- **Batch actions**: scan / install-selected / reboot across selected machines.
- **Parallel execution** via a runspace pool with a concurrency cap; the existing single-target worker becomes the per-machine unit of work.
- **Error isolation**: one machine failing must not abort the batch; surface per-machine errors.

**Design considerations**
- Biggest architectural change (single-target GUI → fleet). Suggest phasing: **Phase A** sequential multi-target with the master grid; **Phase B** parallel execution.
- Mutual-exclusion model changes (N concurrent ops); credential caching is already per-target.
- Aggregate progress/status reporting.

### 4. Update selectivity & control  _(medium; mostly additive)_

More control over *which* updates install, instead of "all missing." Motivated in part by real-world bad-patch incidents (e.g. the April 2026 Windows Server 2025 reboot-loop updates) where admins need to avoid specific updates.

**Scope**
- Select/filter by **severity** (Critical/Important/Moderate/Low) and **category** (Security, Critical, Driver, Definition, Optional/Preview).
- Quick action: **install Critical/Security only**, skip drivers/optional.
- **KB exclusion / block list** (persisted) — never install specified KBs (e.g. known-bad updates).
- **Hide / block updates at the WUA level** — mark specific updates "hidden" so Windows stops offering them (the standard approach in comparable tools like WuMgr/WAU), distinct from install-time exclusion.
- **Dry-run / preview** of what would be installed.

**Design considerations**
- Capture update **category** in the scan result object (currently we capture severity but not category); WUA exposes `Categories`, `MsrcSeverity`, `BrowseOnly` (optional), `IsMandatory`, and supports hide via `IUpdate.IsHidden`.
- Exclusion/block list needs persistence — an optional sibling `.json` config (keep the portable-single-file principle: absent file = no exclusions).
- Largely independent of feature #1 — can ship first.

### Remote uninstall / rollback  _(new from landscape research; reuses the remote engine)_

Pull back a bad update — locally or on a remote server — when a patch breaks something. This addresses the most acute current pain point: patches that cause boot/reboot loops.

**Scope**
- Uninstall a selected installed update on local or remote targets.
- Pair with the existing **remote reboot** to complete the rollback.

**Design considerations**
- Run `wusa.exe /uninstall /kb:<id>` (or the WUA/DISM uninstall path) **locally on the target via the same SYSTEM scheduled-task engine** used for remote install — remote uninstall hits the same network-logon restriction as install.
- Not every update is uninstallable (some are permanent / superseded); detect and surface that clearly.
- Strong confirmation (removing a security update carries its own risk) and reuse of cached credentials.

### 5. Audit logging & multi-machine reporting  _(logging is small; reporting builds on #1)_

A compliance trail and client-facing reports.

**Scope**
- **Audit log**: append-only, timestamped log of actions (scan / install / reboot, target, result, outcome) in a structured + human-readable format.
- **Consolidated report**: extend the existing HTML/CSV export to cover **multiple machines** — a patch-compliance summary suitable to hand to a client.

**Design considerations**
- Log location: optional sibling file (portable); consider simple rotation; decide log verbosity.
- The multi-machine report depends on feature #1's data model; the **per-action audit log is independent and can ship early**.
- Reuse and extend `Export-ToHtml`.

---

## Suggested sequencing

1. **Update selectivity & control (#4, incl. hide/block KBs)**, **remote uninstall / rollback**, and the **per-action audit log (part of #5)** — all independent, additive, and ship-able without the big refactor; together they address the most urgent real-world pain (avoiding and removing bad patches).
2. **Fleet / multi-target (#1)** — the major architectural step (phase sequential, then parallel).
3. **Multi-machine reporting (rest of #5)** — rides on the fleet data model from #1.

## Farther future

- **Post-reboot health check** — after the remote reboot-wait returns, validate the server came back healthy (key services up, not in a boot loop) rather than just "WinRM answered."
- **Update history view** — a dedicated history view (dates, categories, installed/failed state) beyond the current scan snapshot, as comparable tools provide.

## Considered but deferred

- **ConnectWise Manage integration** — auto-post patch notes/time entries to the matching configuration or open/update a ticket. High workflow value given the MSP context; revisit after fleet support exists.
- **Scheduling & unattended mode** — a headless `-Unattended -Target -Install -Reboot` mode for off-hours maintenance windows via Task Scheduler, with email/report on completion.
- **Polish**: real per-update download/install progress (vs. marquee), pre-reboot "who's logged on" check, WSUS-vs-Microsoft-Update source selection, disk-space pre-check.
- **Remote worker via signed script file instead of `-EncodedCommand`** — bundle with code-signing. Base64 `-EncodedCommand` is a common EDR heuristic trigger; running a dropped, *signed* `.ps1` via `-File` is friendlier to EDRs that flag only the encoded form. (Note: this does **not** help behavioral AVs that sandbox the spawned process regardless of how it's launched — see Known deployment considerations.)

## Toward 1.0 (packaging & trust)

The current batch is the on-ramp to 1.0.0, not 1.0 itself. Graduating out of beta is gated on:
- **Code-signing** the EXE (the key legitimacy step — earns SmartScreen trust and enables clean publisher-based AV/EDR allowlisting). Ship the signed EXE *as* 1.0.0.
- **More real-world field testing** of the remote paths (HTTPS/5986 and the WMI fallback have not fired in the field; the credentialed-remote path was only verified to the auth boundary).
- Optionally the `-EncodedCommand`→signed-script change above.

## Known deployment considerations

From field testing (real client endpoints):
- **Aggressive endpoint AV/EDR can break remote scans.** Behavioral products such as **Webroot SecureAnywhere** sandbox/block the remote SYSTEM-task worker — the task runs (exit `0`) but its results never materialize, so the scan silently falls back to **partial WMI data** (installed hotfixes only, no "missing"). This is *not* fixable in the tool (no launch technique evades a process-sandboxing AV); it requires **allowlisting the tool in the AV/EDR console**. Code-signing makes that a clean publisher allowlist. Documented in the README's "Troubleshooting remote scans."
- **Multi-homed targets / addressing.** "WinRM unreachable" despite WinRM being enabled usually means the entered address (often the hostname) resolves to an IP the controller can't route to; use the target's IP on the controller's subnet. Also confirm a listener exists on 5985 (the WinRM *service* running is not sufficient).

## Non-goals

- **Third-party application patching** (browsers, Java, runtimes, etc.) — the headline differentiator of full RMM/patch suites, but a fundamentally different mechanism (winget/vendor catalogs, not the Windows Update Agent) and a major scope expansion. Intentionally out of scope to keep this a focused, portable WUA utility.
