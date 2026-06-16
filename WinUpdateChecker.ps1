#Requires -Version 5.1
<#
.SYNOPSIS
    Windows Update Verification Tool - checks installed vs missing updates locally or on a remote server.
.DESCRIPTION
    WinForms GUI that queries the Windows Update API (WUA COM) for update status,
    displays results color-coded in a DataGridView, and exports to CSV or HTML.
    Remote scanning and installation use a SYSTEM-context scheduled task on the target
    via CIM (no WinRM required for the scan itself; WinRM is used only to read back
    the result JSON and clean up).
    Remote Reboot: after a successful remote scan the user can schedule a forced reboot
    (shutdown.exe /r /f /t 60) on the target via WinRM/Invoke-Command, with countdown,
    abort, and auto-rescan after the target comes back online.
.NOTES
    Version: 1.0.0-beta.3
    Self-elevates at startup (installing updates requires Administrator).
    Remote scanning/installation requires WinRM on the target for result retrieval.
#>

param([switch]$NoElevate)   # -NoElevate: skip self-elevation (used by automated tests)

#region -- Hide Console Window --------------------------------------------------
# Hide the PowerShell console that hosts this GUI -- but ONLY if this process
# owns it (launched via double-click / "Run with PowerShell"). When started
# from an existing terminal, more than one process is attached to the console
# and hiding it would hide the user's terminal.
Add-Type -Name ConsoleWin -Namespace Native -MemberDefinition '
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]   public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("kernel32.dll")] public static extern uint GetConsoleProcessList(uint[] processList, uint processCount);
'
# uint32, not uint: the uint accelerator does not exist on PowerShell 5.1
$consoleProcs = New-Object uint32[] 2
if ([Native.ConsoleWin]::GetConsoleProcessList($consoleProcs, 2) -le 1) {
    [void][Native.ConsoleWin]::ShowWindow([Native.ConsoleWin]::GetConsoleWindow(), 0)  # 0 = SW_HIDE
}
#endregion

#region -- Self-Elevation -------------------------------------------------------
# Installing updates via the WUA API requires Administrator. Relaunch elevated
# if needed; if the UAC prompt is declined, explain and exit.
if (-not $NoElevate) {
    $wid  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $prin = New-Object Security.Principal.WindowsPrincipal($wid)
    if (-not $prin.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        try {
            Start-Process -FilePath "$PSHOME\powershell.exe" `
                -ArgumentList "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`"" `
                -Verb RunAs -ErrorAction Stop | Out-Null
        } catch {
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show(
                ('Administrator rights are required to scan for and install updates.' +
                 "`n`nThe application will now exit."),
                'Windows Update Checker',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        }
        exit
    }
}
#endregion

#region -- Imports ------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
#endregion

#region -- Shared Script State ------------------------------------------------
$script:ScanResults     = @()
$script:IsPartialData   = $false
$script:RebootPending   = $false
$script:ScannedOS       = ''     # Fix 11: captured from target during scan
$script:RunningPS       = $null  # Fix 1 / Fix 12: active PowerShell instance
$script:RunningRS       = $null  # Fix 1 / Fix 12: active Runspace instance
$script:PollHandle      = $null  # Fix 1: IAsyncResult from BeginInvoke
$script:PollTimer       = $null  # Fix 1 / Fix 12: System.Windows.Forms.Timer
$script:ActiveOp        = $null  # 'Scan' | 'Install' | $null (mutual exclusion)
$script:CanInstall      = $false # last scan succeeded (local WUA or remote SYSTEM-task WUA)
# Remote Reboot state machine
$script:CanReboot            = $false   # true only after a successful remote scan
$script:RebootPhase          = 'Idle'   # 'Idle' | 'Countdown' | 'Waiting'
$script:RebootMode           = $null    # 'Issue' | 'Abort' | 'Wait'  -- which bg op is running
$script:RebootCountdown      = 0        # seconds remaining in countdown
$script:RebootTimer          = $null    # System.Windows.Forms.Timer for 1-second countdown
$script:RebootWaitStart      = $null    # [datetime] when the Waiting phase began
$script:RebootDelaySec       = 60       # shutdown /t value (seconds before reboot)
$script:RebootWaitTimeoutSec = 1800     # 30-minute poll timeout
# Remote transport state -- stored at scan-launch; reused by Install
$script:LastWasRemote   = $false
$script:LastTarget      = ''
$script:LastCredential  = $null
$script:LastCredTarget  = ''    # target the cached credential was entered for (reuse key)
$script:LastUseSSL      = $false
$script:LastErrorText   = ''    # full text of the last error (for the details dialog)
# TrustedHosts auto-management state
$script:TrustedHostsManaged    = $false  # false if WSMan unavailable / non-admin
$script:OriginalTrustedHosts   = $null   # value at startup (baseline)
$script:AddedTrustedHosts      = @()     # hosts added this session (for restore logic)
$script:TrustedHostsMarkerPath = Join-Path $env:LOCALAPPDATA 'WinUpdateChecker\trustedhosts.json'
#endregion

#region -- Export Functions ---------------------------------------------------

function Get-FilteredResults {
    # Fix 8: shared filter logic used by Update-Grid AND both export buttons
    $statusFilter   = $cmbStatus.SelectedItem
    $severityFilter = $cmbSeverity.SelectedItem

    return $script:ScanResults | Where-Object {
        ($statusFilter   -eq 'All' -or $_.Status   -eq $statusFilter) -and
        ($severityFilter -eq 'All' -or $_.Severity -eq $severityFilter)
    }
}

function Export-ToCsv {
    param([object[]]$Records, [string]$Path)
    $Records | Select-Object KB, Title, Status, Severity, Date, SizeKB, Source |
        Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
}

function Export-ToHtml {
    # Fix 7: HTML-encode all interpolated values
    # Fix 11: OSVersion comes from scan result stored in script scope
    param([object[]]$Records, [string]$Path, [string]$TargetName, [string]$OSVersion)

    $timestamp  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $installed  = ($Records | Where-Object Status -eq 'Installed').Count
    $missing    = ($Records | Where-Object Status -eq 'Missing').Count

    # Fix 6: Reboot state is a separate flag, not a per-row status.
    $rebootText   = if ($script:RebootPending) { 'Yes' } else { 'No' }
    $rebootColor  = if ($script:RebootPending) { '#ffe6e6' } else { '#e6ffe6' }
    $rebootBorder = if ($script:RebootPending) { '#f44336' } else { '#4caf50' }

    # Fix 7: encode header-level values
    $safeTarget = [System.Net.WebUtility]::HtmlEncode($TargetName)
    $safeOS     = [System.Net.WebUtility]::HtmlEncode($OSVersion)

    $rows = $Records | ForEach-Object {
        $color = if ($_.Status -eq 'Installed') {
            '#e6ffe6'
        } elseif ($_.Status -eq 'Missing') {
            '#ffe6e6'
        } else {
            '#ffffff'
        }
        # Fix 7: encode per-row values
        $safeKB     = [System.Net.WebUtility]::HtmlEncode([string]$_.KB)
        $safeTitle  = [System.Net.WebUtility]::HtmlEncode([string]$_.Title)
        $safeStatus = [System.Net.WebUtility]::HtmlEncode([string]$_.Status)
        $safeSev    = [System.Net.WebUtility]::HtmlEncode([string]$_.Severity)
        $safeDate   = [System.Net.WebUtility]::HtmlEncode([string]$_.Date)
        $safeSize   = [System.Net.WebUtility]::HtmlEncode([string]$_.SizeKB)
        "<tr style='background:$color'><td>$safeKB</td><td>$safeTitle</td><td>$safeStatus</td><td>$safeSev</td><td>$safeDate</td><td>$safeSize</td></tr>"
    }

    $rowsJoined = $rows -join "`n"

    $html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Windows Update Report - $safeTarget</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; color: #222; }
  h1   { color: #0078d4; }
  .summary { display:flex; gap:20px; margin:16px 0; }
  .badge { padding:10px 18px; border-radius:6px; font-weight:bold; font-size:14px; }
  .green  { background:#e6ffe6; border:1px solid #4caf50; }
  .red    { background:#ffe6e6; border:1px solid #f44336; }
  table { border-collapse:collapse; width:100%; font-size:13px; }
  th    { background:#0078d4; color:#fff; padding:8px 10px; text-align:left; }
  td    { padding:6px 10px; border-bottom:1px solid #ddd; }
  tr:hover td { filter: brightness(0.96); }
</style>
</head>
<body>
<h1>Windows Update Report</h1>
<p><strong>Target:</strong> $safeTarget &nbsp;|&nbsp; <strong>OS:</strong> $safeOS &nbsp;|&nbsp; <strong>Scanned:</strong> $timestamp</p>
<div class="summary">
  <div class="badge green">&#10003; Installed: $installed</div>
  <div class="badge red">&#10007; Missing: $missing</div>
  <div class="badge" style="background:$rebootColor; border:1px solid $rebootBorder;">&#9888; Reboot Pending: $rebootText</div>
</div>
<table>
<thead><tr><th>KB</th><th>Title</th><th>Status</th><th>Severity</th><th>Date</th><th>Size (KB)</th></tr></thead>
<tbody>
$rowsJoined
</tbody>
</table>
</body>
</html>
"@
    $html | Out-File -FilePath $Path -Encoding UTF8
}

#endregion

#region -- WinForms UI --------------------------------------------------------

# -- Main Form --
$script:AppVersion  = '1.0.0-beta.3'
$form               = New-Object System.Windows.Forms.Form
$form.Text          = "Windows Update Checker v$script:AppVersion"
$form.Size          = New-Object System.Drawing.Size(900, 640)
$form.MinimumSize   = New-Object System.Drawing.Size(700, 500)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Segoe UI', 9)

# -- Target Panel --
$pnlTarget           = New-Object System.Windows.Forms.Panel
$pnlTarget.Dock      = 'Top'
$pnlTarget.Height    = 80
$pnlTarget.Padding   = New-Object System.Windows.Forms.Padding(10, 8, 10, 0)
$pnlTarget.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
# Size the panel to the form's client width BEFORE adding the right-anchored
# Scan button: anchor offsets are captured against the panel's width at add
# time (default 200px), which would otherwise push the button off-screen.
$pnlTarget.Width     = $form.ClientSize.Width

$lblTarget          = New-Object System.Windows.Forms.Label
$lblTarget.Text     = 'Target:'
$lblTarget.Location = New-Object System.Drawing.Point(10, 12)
$lblTarget.AutoSize = $true
$lblTarget.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

# -- Target sub-panel: wraps $rbLocal, $rbRemote, $txtHost in their own WinForms
# container so WinForms mutual-exclusion applies only to these two radio buttons.
# Positioned at (65, 8) inside $pnlTarget so that child coords offset correctly:
#   $rbLocal      Location (0,2)   -> absolute in pnlTarget: (65+0, 8+2) = (65,10)  [was (65,10)]
#   $rbRemote     Location (120,2) -> absolute: (65+120, 8+2) = (185,10) [was (185,10)]
#   $txtHost      Location (190,0) -> absolute: (65+190, 8+0) = (255,8)  [was (255,8)]
$pnlTargetGrp           = New-Object System.Windows.Forms.Panel
$pnlTargetGrp.Location  = New-Object System.Drawing.Point(65, 8)
$pnlTargetGrp.Size      = New-Object System.Drawing.Size(700, 30)
$pnlTargetGrp.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$pnlTargetGrp.BorderStyle = 'None'

$rbLocal          = New-Object System.Windows.Forms.RadioButton
$rbLocal.Text     = 'Local Machine'
$rbLocal.Location = New-Object System.Drawing.Point(0, 2)
$rbLocal.AutoSize = $true
$rbLocal.Checked  = $true

$rbRemote          = New-Object System.Windows.Forms.RadioButton
$rbRemote.Text     = 'Remote:'
$rbRemote.Location = New-Object System.Drawing.Point(120, 2)
$rbRemote.AutoSize = $true

$txtHost          = New-Object System.Windows.Forms.TextBox
$txtHost.Location = New-Object System.Drawing.Point(190, 0)
$txtHost.Width    = 200
$txtHost.Enabled  = $false
# Fix 2: PlaceholderText only exists on .NET Core / .NET 5+ WinForms, not on
# .NET Framework 4.x (used by Windows PowerShell 5.1). Guard the assignment.
if ($txtHost.PSObject.Properties['PlaceholderText']) {
    $txtHost.PlaceholderText = 'hostname or IP'
}

$pnlTargetGrp.Controls.AddRange(@($rbLocal, $rbRemote, $txtHost))

$lblAuth          = New-Object System.Windows.Forms.Label
$lblAuth.Text     = 'Auth:'
$lblAuth.Location = New-Object System.Drawing.Point(10, 44)
$lblAuth.AutoSize = $true
$lblAuth.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

# -- Auth sub-panel: wraps $rbCurrentUser, $rbCreds in their own WinForms
# container so this group is independent of the Target group above.
# Positioned at (65, 40) inside $pnlTarget:
#   $rbCurrentUser Location (0,2)   -> absolute in pnlTarget: (65+0, 40+2) = (65,42)  [was (65,42)]
#   $rbCreds       Location (120,2) -> absolute: (65+120, 40+2) = (185,42) [was (185,42)]
$pnlAuthGrp           = New-Object System.Windows.Forms.Panel
$pnlAuthGrp.Location  = New-Object System.Drawing.Point(65, 40)
$pnlAuthGrp.Size      = New-Object System.Drawing.Size(700, 30)
$pnlAuthGrp.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$pnlAuthGrp.BorderStyle = 'None'

$rbCurrentUser           = New-Object System.Windows.Forms.RadioButton
$rbCurrentUser.Text      = 'Current User'
$rbCurrentUser.Location  = New-Object System.Drawing.Point(0, 2)
$rbCurrentUser.AutoSize  = $true
$rbCurrentUser.Checked   = $true

$rbCreds          = New-Object System.Windows.Forms.RadioButton
$rbCreds.Text     = 'Specify Credentials'
$rbCreds.Location = New-Object System.Drawing.Point(120, 2)
$rbCreds.AutoSize = $true

$pnlAuthGrp.Controls.AddRange(@($rbCurrentUser, $rbCreds))

$btnScan            = New-Object System.Windows.Forms.Button
$btnScan.Text       = 'Scan Now'
$btnScan.Location   = New-Object System.Drawing.Point(750, 20)
$btnScan.Size       = New-Object System.Drawing.Size(100, 40)
$btnScan.BackColor  = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnScan.ForeColor  = [System.Drawing.Color]::White
$btnScan.FlatStyle  = 'Flat'
$btnScan.Font       = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$btnScan.Anchor     = 'Top,Right'

# $lblTarget, $lblAuth, $btnScan are direct children of $pnlTarget (labels are
# not radio buttons; $btnScan is anchored Top,Right and must remain here).
# $pnlTargetGrp and $pnlAuthGrp are the two invisible sub-panels.
$pnlTarget.Controls.AddRange(@($lblTarget, $lblAuth, $btnScan,
    $pnlTargetGrp, $pnlAuthGrp))

# -- Summary Panel --
$pnlSummary           = New-Object System.Windows.Forms.Panel
$pnlSummary.Dock      = 'Top'
$pnlSummary.Height    = 84
$pnlSummary.Padding   = New-Object System.Windows.Forms.Padding(10, 8, 10, 0)
$pnlSummary.BackColor = [System.Drawing.Color]::White

$lblInstalled           = New-Object System.Windows.Forms.Label
$lblInstalled.Text      = 'Installed: -'
$lblInstalled.Location  = New-Object System.Drawing.Point(10, 10)
$lblInstalled.AutoSize  = $true
$lblInstalled.ForeColor = [System.Drawing.Color]::FromArgb(46, 125, 50)
$lblInstalled.Font      = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)

$lblMissing           = New-Object System.Windows.Forms.Label
$lblMissing.Text      = 'Missing: -'
$lblMissing.Location  = New-Object System.Drawing.Point(160, 10)
$lblMissing.AutoSize  = $true
$lblMissing.ForeColor = [System.Drawing.Color]::FromArgb(198, 40, 40)
$lblMissing.Font      = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)

# Fix 6: label now shows "Reboot Pending: Yes/No" colored appropriately;
# it no longer counts rows with status 'Pending Reboot' (that status is not used).
$lblPending           = New-Object System.Windows.Forms.Label
$lblPending.Text      = 'Reboot Pending: -'
$lblPending.Location  = New-Object System.Drawing.Point(310, 10)
$lblPending.AutoSize  = $true
$lblPending.ForeColor = [System.Drawing.Color]::FromArgb(230, 81, 0)
$lblPending.Font      = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)

$lblScanTime           = New-Object System.Windows.Forms.Label
$lblScanTime.Text      = ''
$lblScanTime.Location  = New-Object System.Drawing.Point(10, 36)
$lblScanTime.AutoSize  = $true
$lblScanTime.ForeColor = [System.Drawing.Color]::Gray

# Own full-width row (the WSUS source/pause summary can be long, e.g. a full
# WSUS URL + pause date — placing it at x=10 on its own row avoids clipping
# off the right edge of the panel).
$lblPolicy           = New-Object System.Windows.Forms.Label
$lblPolicy.Text      = ''
$lblPolicy.Location  = New-Object System.Drawing.Point(10, 58)
$lblPolicy.AutoSize  = $true
$lblPolicy.ForeColor = [System.Drawing.Color]::Gray
$lblPolicy.Font      = New-Object System.Drawing.Font('Segoe UI', 9)

$pnlSummary.Controls.AddRange(@($lblInstalled, $lblMissing, $lblPending, $lblScanTime, $lblPolicy))

# -- Progress Panel --
$pnlProgress           = New-Object System.Windows.Forms.Panel
$pnlProgress.Dock      = 'Top'
$pnlProgress.Height    = 28
$pnlProgress.Padding   = New-Object System.Windows.Forms.Padding(10, 2, 10, 0)
$pnlProgress.Visible   = $false

$progressBar        = New-Object System.Windows.Forms.ProgressBar
$progressBar.Dock   = 'Fill'
$progressBar.Style  = 'Marquee'
$pnlProgress.Controls.Add($progressBar)

# -- Error Banner --
$pnlError           = New-Object System.Windows.Forms.Panel
$pnlError.Dock      = 'Top'
$pnlError.Height    = 30
$pnlError.BackColor = [System.Drawing.Color]::FromArgb(255, 230, 230)
$pnlError.Visible   = $false
$pnlError.Padding   = New-Object System.Windows.Forms.Padding(10, 5, 10, 5)

$lblError              = New-Object System.Windows.Forms.Label
$lblError.Dock         = 'Fill'
$lblError.ForeColor    = [System.Drawing.Color]::FromArgb(198, 40, 40)
$lblError.Font         = New-Object System.Drawing.Font('Segoe UI', 9)
# Wrap long (e.g. remote WinRM/CIM) messages; show an ellipsis if still clipped,
# and let the user click the banner for the full, copyable text.
$lblError.AutoEllipsis = $true
$lblError.Cursor       = [System.Windows.Forms.Cursors]::Hand
$pnlError.Controls.Add($lblError)

$errToolTip = New-Object System.Windows.Forms.ToolTip
$errToolTip.SetToolTip($lblError, 'Click to view the full error and copy it.')
$lblError.Add_Click({ Show-ErrorDetails })
$pnlError.Add_Click({ Show-ErrorDetails })

# -- Filter Panel --
$pnlFilter           = New-Object System.Windows.Forms.Panel
$pnlFilter.Dock      = 'Top'
$pnlFilter.Height    = 34
$pnlFilter.Padding   = New-Object System.Windows.Forms.Padding(10, 4, 10, 0)
$pnlFilter.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)

$lblFilterStatus          = New-Object System.Windows.Forms.Label
$lblFilterStatus.Text     = 'Status:'
$lblFilterStatus.Location = New-Object System.Drawing.Point(10, 7)
$lblFilterStatus.AutoSize = $true

$cmbStatus                = New-Object System.Windows.Forms.ComboBox
$cmbStatus.Location       = New-Object System.Drawing.Point(55, 4)
$cmbStatus.Width          = 130
$cmbStatus.DropDownStyle  = 'DropDownList'
# Fix 6: removed 'Pending Reboot' from filter - that status no longer appears on rows
$cmbStatus.Items.AddRange(@('All', 'Installed', 'Missing'))
$cmbStatus.SelectedIndex  = 0

$lblFilterSev          = New-Object System.Windows.Forms.Label
$lblFilterSev.Text     = 'Severity:'
$lblFilterSev.Location = New-Object System.Drawing.Point(200, 7)
$lblFilterSev.AutoSize = $true

$cmbSeverity               = New-Object System.Windows.Forms.ComboBox
$cmbSeverity.Location      = New-Object System.Drawing.Point(252, 4)
$cmbSeverity.Width         = 120
$cmbSeverity.DropDownStyle = 'DropDownList'
$cmbSeverity.Items.AddRange(@('All', 'Critical', 'Important', 'Moderate', 'Low', 'N/A'))
$cmbSeverity.SelectedIndex = 0

$pnlFilter.Controls.AddRange(@($lblFilterStatus, $cmbStatus, $lblFilterSev, $cmbSeverity))

# -- DataGridView --
$grid                         = New-Object System.Windows.Forms.DataGridView
$grid.Dock                    = 'Fill'
# Grid-level ReadOnly must be off so the Select checkbox column is editable;
# every other column is set ReadOnly individually below.
$grid.ReadOnly                = $false
$grid.AllowUserToAddRows      = $false
$grid.AllowUserToDeleteRows   = $false
$grid.RowHeadersVisible       = $false
$grid.AutoSizeColumnsMode     = 'Fill'
$grid.SelectionMode           = 'FullRowSelect'
$grid.BackgroundColor         = [System.Drawing.Color]::White
$grid.BorderStyle             = 'None'
$grid.ColumnHeadersHeightSizeMode = 'AutoSize'
$grid.Font                    = New-Object System.Drawing.Font('Segoe UI', 9)

# Checkbox column FIRST (column 0) -- fixed width so the Title Fill column
# sizing is unaffected
$colSelect              = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colSelect.Name         = 'Select'
$colSelect.HeaderText   = ''
$colSelect.Width        = 40
$colSelect.AutoSizeMode = 'None'
$grid.Columns.Add($colSelect) | Out-Null

$cols = @(
    @{ Name='KB';       Header='KB';        Width=90;  Fill=$false },
    @{ Name='Title';    Header='Title';     Width=350; Fill=$true  },
    @{ Name='Status';   Header='Status';    Width=110; Fill=$false },
    @{ Name='Severity'; Header='Severity';  Width=90;  Fill=$false },
    @{ Name='Date';     Header='Date';      Width=90;  Fill=$false },
    @{ Name='SizeKB';   Header='Size (KB)'; Width=80;  Fill=$false }
)
foreach ($c in $cols) {
    $col            = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name       = $c.Name
    $col.HeaderText = $c.Header
    $col.ReadOnly   = $true
    if ($c.Fill) {
        $col.AutoSizeMode = 'Fill'
    } else {
        $col.Width        = $c.Width
        $col.AutoSizeMode = 'None'
    }
    $grid.Columns.Add($col) | Out-Null
}

# Hidden column carrying the WUA UpdateID for each row (empty for WMI rows);
# used to re-acquire the IUpdate COM objects at install time
$colUid          = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colUid.Name     = 'UpdateID'
$colUid.Visible  = $false
$colUid.ReadOnly = $true
$grid.Columns.Add($colUid) | Out-Null

# -- Bottom Panel --
$pnlBottom           = New-Object System.Windows.Forms.Panel
$pnlBottom.Dock      = 'Bottom'
$pnlBottom.Height    = 40
$pnlBottom.Padding   = New-Object System.Windows.Forms.Padding(10, 5, 10, 5)
$pnlBottom.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)

$btnExportCsv          = New-Object System.Windows.Forms.Button
$btnExportCsv.Text     = 'Export CSV'
$btnExportCsv.Location = New-Object System.Drawing.Point(10, 5)
$btnExportCsv.Size     = New-Object System.Drawing.Size(100, 28)
$btnExportCsv.Enabled  = $false

$btnExportHtml          = New-Object System.Windows.Forms.Button
$btnExportHtml.Text     = 'Export HTML'
$btnExportHtml.Location = New-Object System.Drawing.Point(118, 5)
$btnExportHtml.Size     = New-Object System.Drawing.Size(100, 28)
$btnExportHtml.Enabled  = $false

$btnClear          = New-Object System.Windows.Forms.Button
$btnClear.Text     = 'Clear'
$btnClear.Location = New-Object System.Drawing.Point(226, 5)
$btnClear.Size     = New-Object System.Drawing.Size(80, 28)

$btnInstall           = New-Object System.Windows.Forms.Button
$btnInstall.Text      = 'Install Selected'
$btnInstall.Location  = New-Object System.Drawing.Point(314, 5)
$btnInstall.Size      = New-Object System.Drawing.Size(130, 28)
$btnInstall.Enabled   = $false
$btnInstall.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
$btnInstall.ForeColor = [System.Drawing.Color]::White
$btnInstall.FlatStyle = 'Flat'

# Reboot Remote button -- destructive styling; enabled only after a successful remote scan
$btnReboot           = New-Object System.Windows.Forms.Button
$btnReboot.Text      = 'Reboot Remote'
$btnReboot.Location  = New-Object System.Drawing.Point(452, 5)
$btnReboot.Size      = New-Object System.Drawing.Size(130, 28)
$btnReboot.Enabled   = $false
$btnReboot.BackColor = [System.Drawing.Color]::FromArgb(180, 30, 30)
$btnReboot.ForeColor = [System.Drawing.Color]::White
$btnReboot.FlatStyle = 'Flat'
$btnReboot.Font      = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

# Tooltip for the Install button (plan item 11: dropped "(local scans only)")
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($btnInstall,
    'Download and install the checked missing updates on the target machine.')
$toolTip.SetToolTip($btnReboot,
    'Schedule a forced reboot of the remote target (60-second delay, force-close apps).')

$lblStatus           = New-Object System.Windows.Forms.Label
$lblStatus.Text      = 'Ready'
$lblStatus.Location  = New-Object System.Drawing.Point(594, 10)
$lblStatus.AutoSize  = $true
$lblStatus.ForeColor = [System.Drawing.Color]::Gray

$pnlBottom.Controls.AddRange(@($btnExportCsv, $btnExportHtml, $btnClear, $btnInstall, $btnReboot, $lblStatus))

# -- Assemble Form --
$form.Controls.AddRange(@($grid, $pnlBottom, $pnlFilter, $pnlError,
    $pnlProgress, $pnlSummary, $pnlTarget))

#endregion

#region -- Helper Functions ---------------------------------------------------

function Update-Grid {
    # Fix 8: uses Get-FilteredResults (shared with exports)
    # Note: rebuilding the rows resets any checked Select boxes (documented).
    $filtered = Get-FilteredResults
    $grid.Rows.Clear()
    foreach ($u in $filtered) {
        $idx = $grid.Rows.Add($false, $u.KB, $u.Title, $u.Status, $u.Severity,
            $u.Date, $u.SizeKB, [string]$u.UpdateID)
        # Only Missing updates with a WUA identity are installable
        $installable = ($u.Status -eq 'Missing') -and
            -not [string]::IsNullOrEmpty([string]$u.UpdateID)
        if (-not $installable) {
            $grid.Rows[$idx].Cells['Select'].ReadOnly = $true
        }
    }
    Update-InstallButton
    Update-RebootButton
}

function Get-CheckedUpdateIDs {
    $ids = @()
    foreach ($row in $grid.Rows) {
        if ($row.Cells['Select'].Value -eq $true) {
            $id = [string]$row.Cells['UpdateID'].Value
            if ($id) { $ids += $id }
        }
    }
    return ,$ids
}

function Update-InstallButton {
    $checked = (Get-CheckedUpdateIDs).Count
    $btnInstall.Text = if ($checked -gt 0) {
        "Install Selected ($checked)"
    } else {
        'Install Selected'
    }
    $btnInstall.Enabled = ($null -eq $script:ActiveOp) -and
        $script:CanInstall -and ($checked -gt 0)
}

function Update-RebootButton {
    # Only governs Idle baseline; Countdown/Waiting phases set button text/enabled explicitly.
    if ($script:RebootPhase -ne 'Idle') { return }
    $btnReboot.Text    = 'Reboot Remote'
    $btnReboot.Enabled = ($null -eq $script:ActiveOp) -and
                         $script:CanReboot -and
                         $script:LastWasRemote
}

function Reset-RebootState {
    # Stop and dispose the countdown timer, then return all reboot state to Idle.
    # Called from: Abort-done, Stop-Waiting click, Wait-done, error paths.
    if ($null -ne $script:RebootTimer) {
        $script:RebootTimer.Stop()
        $script:RebootTimer.Dispose()
        $script:RebootTimer = $null
    }
    $script:RebootPhase  = 'Idle'
    $script:RebootMode   = $null
    $script:ActiveOp     = $null
    $btnScan.Enabled     = $true
    $pnlProgress.Visible = $false
    Update-InstallButton
    Update-RebootButton
}

function Show-Error {
    param([string]$Message)
    $script:LastErrorText = $Message
    $lblError.Text        = $Message
    # Grow the banner to fit the wrapped message, capped at ~5 lines so it never
    # dominates the window; the full text is always available via click.
    $w = $pnlError.ClientSize.Width - 20
    if ($w -lt 50) { $w = 600 }   # fallback before the panel has been laid out
    $sz = [System.Windows.Forms.TextRenderer]::MeasureText(
        $Message, $lblError.Font,
        (New-Object System.Drawing.Size($w, 0)),
        [System.Windows.Forms.TextFormatFlags]::WordBreak)
    $h = $sz.Height + 12
    if ($h -lt 30) { $h = 30 }
    if ($h -gt 96) { $h = 96 }
    $pnlError.Height  = $h
    $pnlError.Visible = $true
}

function Show-ErrorDetails {
    # Full, selectable, copyable error text — for messages too long for the banner.
    if ([string]::IsNullOrEmpty($script:LastErrorText)) { return }
    $dlg                 = New-Object System.Windows.Forms.Form
    $dlg.Text            = 'Error Details'
    $dlg.Size            = New-Object System.Drawing.Size(640, 360)
    $dlg.StartPosition   = 'CenterParent'
    $dlg.MinimizeBox     = $false
    $dlg.MaximizeBox     = $false

    $txt            = New-Object System.Windows.Forms.TextBox
    $txt.Multiline  = $true
    $txt.ReadOnly   = $true
    $txt.ScrollBars = 'Vertical'
    $txt.Dock       = 'Fill'
    $txt.Font       = New-Object System.Drawing.Font('Consolas', 9)
    $txt.Text       = $script:LastErrorText

    $pnl         = New-Object System.Windows.Forms.Panel
    $pnl.Dock    = 'Bottom'
    $pnl.Height  = 40
    $pnl.Padding = New-Object System.Windows.Forms.Padding(10, 6, 10, 6)

    $btnCopy      = New-Object System.Windows.Forms.Button
    $btnCopy.Text = 'Copy'
    $btnCopy.Dock = 'Left'
    $btnCopy.Width = 90
    $btnCopy.Add_Click({ [System.Windows.Forms.Clipboard]::SetText($script:LastErrorText) })

    $btnCloseDet      = New-Object System.Windows.Forms.Button
    $btnCloseDet.Text = 'Close'
    $btnCloseDet.Dock = 'Right'
    $btnCloseDet.Width = 90
    $btnCloseDet.Add_Click({ $dlg.Close() })

    $pnl.Controls.AddRange(@($btnCopy, $btnCloseDet))
    $dlg.Controls.AddRange(@($txt, $pnl))
    $dlg.ShowDialog($form) | Out-Null
}

function Hide-Error {
    $pnlError.Visible = $false
    $pnlError.Height  = 30
}

function Update-Summary {
    $installed = ($script:ScanResults | Where-Object Status -eq 'Installed').Count
    $missing   = ($script:ScanResults | Where-Object Status -eq 'Missing').Count

    $lblInstalled.Text = "Installed: $installed"
    $lblMissing.Text   = "Missing: $missing"

    # Fix 6: show reboot state as Yes/No colored label, not a count of rows
    if ($script:RebootPending) {
        $lblPending.Text      = 'Reboot Pending: Yes'
        $lblPending.ForeColor = [System.Drawing.Color]::FromArgb(198, 40, 40)
    } else {
        $lblPending.Text      = 'Reboot Pending: No'
        $lblPending.ForeColor = [System.Drawing.Color]::FromArgb(46, 125, 50)
    }

    $timeStr = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    if ($script:IsPartialData) {
        $lblScanTime.Text = "Last scan: $timeStr  [Partial data - WMI hotfixes only]"
    } else {
        $lblScanTime.Text = "Last scan: $timeStr"
    }
}

function Translate-WuaError {
    # Scans a message string for known WUA hex error codes and returns a
    # plain-English explanation appended in parentheses.  Safe on any input;
    # returns the original string unchanged when no known code is found.
    param([string]$Message)
    if ([string]::IsNullOrEmpty($Message)) { return $Message }

    $knownCodes = @{
        '0x8024002E' = 'Windows Update is disabled by Group Policy (access removed).'
        '0x8024000B' = 'The operation was cancelled by policy.'
        '0x8024402C' = 'Could not connect to the configured WSUS server.'
        '0x80244022' = 'The update server returned HTTP 503 (service unavailable).'
        '0x8024401C' = 'The request to the update server timed out.'
        '0x80240044' = 'Update install was denied (per-machine policy / WSUS-managed).'
        '0x80072EE2' = 'Could not reach the update server (network timeout/refused).'
        '0x80072EFD' = 'Could not reach the update server (network timeout/refused).'
        '0x800705B4' = 'Could not reach the update server (network timeout/refused).'
    }

    foreach ($code in $knownCodes.Keys) {
        if ($Message -match [regex]::Escape($code)) {
            return "$Message  ($($knownCodes[$code]))"
        }
    }
    return $Message
}

function Stop-AsyncCleanup {
    # Fix 12: dispose runspace/PS/poll-timer cleanly when the active background
    # operation (scan, install, or reboot sub-op) completes, aborts, or the form closes.
    # NOTE: Stop-AsyncCleanup intentionally does NOT abort a detached remote
    # SYSTEM scheduled task -- orphan-cleanup-on-connect is the safety net.
    # NOTE: Stop-AsyncCleanup does NOT touch $script:RebootTimer or $script:RebootPhase;
    # those are managed exclusively by Reset-RebootState so that entering the Countdown
    # phase (after the Issue op's finally block runs this cleanup) leaves the countdown intact.
    if ($null -ne $script:PollTimer) {
        $script:PollTimer.Stop()
        $script:PollTimer.Dispose()
        $script:PollTimer = $null
    }
    if ($null -ne $script:RunningPS) {
        try { $script:RunningPS.Stop() } catch { }
        $script:RunningPS.Dispose()
        $script:RunningPS = $null
    }
    if ($null -ne $script:RunningRS) {
        $script:RunningRS.Close()
        $script:RunningRS.Dispose()
        $script:RunningRS = $null
    }
    $script:PollHandle = $null
    # Clear ActiveOp unconditionally. For Scan/Install this returns the mutex to idle.
    # For Reboot ops: the reboot tick handler re-asserts ActiveOp='Reboot' immediately
    # after calling this (before any other UI-thread code can run), so clearing here is safe.
    $script:ActiveOp = $null
}

#endregion

#region -- TrustedHosts Auto-Management ---------------------------------------
#
# Two thin WSMan wrappers make all real logic unit-testable: tests simply
# redefine Get-CurrentTrustedHosts / Set-CurrentTrustedHosts to operate on an
# in-memory variable instead of the real WSMan provider.
#
# Design constraints:
#   - PS 5.1 / .NET Framework only.
#   - Non-admin / WSMan-unavailable -> TrustedHostsManaged=$false, all no-ops.
#   - Never remove pre-existing entries; baseline-restore handles full revert.
#   - Handles '*' wildcard (skip adding; already globally trusted).
#   - Empty baseline round-trips correctly (clearing TrustedHosts to '' is valid).

function Get-CurrentTrustedHosts {
    # Returns the current TrustedHosts string value, or $null on failure.
    try {
        return (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction Stop).Value
    } catch {
        return $null
    }
}

function Set-CurrentTrustedHosts {
    param([string]$Value)
    # Sets TrustedHosts to $Value (empty string clears the list). Returns $true on
    # success, $false on failure.
    try {
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value $Value -Force -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Initialize-TrustedHostsManagement {
    # Called once at startup.
    # 1. If a marker file exists (previous crash): restore baseline and delete it.
    # 2. Capture current value as baseline, write fresh marker.
    # No-ops gracefully if WSMan is unavailable or this process lacks admin.

    # Ensure the directory exists (no-op if already present).
    $markerDir = Split-Path $script:TrustedHostsMarkerPath -Parent
    if (-not (Test-Path $markerDir)) {
        try { New-Item -ItemType Directory -Path $markerDir -Force -ErrorAction Stop | Out-Null }
        catch { $script:TrustedHostsManaged = $false; return }
    }

    # Crash-sweep: if a stale marker exists, restore the baseline it contains.
    if (Test-Path $script:TrustedHostsMarkerPath) {
        try {
            $stale = Get-Content -Raw -LiteralPath $script:TrustedHostsMarkerPath -ErrorAction Stop |
                     ConvertFrom-Json -ErrorAction Stop
            $staleBaseline = [string]$stale.Baseline   # may be empty string
            Set-CurrentTrustedHosts -Value $staleBaseline | Out-Null
        } catch {
            # If we cannot read/parse the marker or restore, remove it and continue.
        }
        try { Remove-Item -LiteralPath $script:TrustedHostsMarkerPath -Force -ErrorAction SilentlyContinue } catch {}
    }

    # Probe WSMan availability (also validates admin rights).
    $currentValue = Get-CurrentTrustedHosts
    if ($null -eq $currentValue) {
        # WSMan unavailable or access denied -- no-op for the entire session.
        $script:TrustedHostsManaged  = $false
        $script:OriginalTrustedHosts = $null
        return
    }

    # Capture baseline and write marker so a crash leaves the next launch able to restore.
    $script:OriginalTrustedHosts = $currentValue
    $script:TrustedHostsManaged  = $true
    $script:AddedTrustedHosts    = @()

    try {
        $marker = [PSCustomObject]@{ Baseline = $currentValue } | ConvertTo-Json -Compress
        [System.IO.File]::WriteAllText($script:TrustedHostsMarkerPath, $marker,
            [System.Text.Encoding]::UTF8)
    } catch {
        # Cannot write marker -- disable management to avoid partial state.
        $script:TrustedHostsManaged  = $false
        $script:OriginalTrustedHosts = $null
    }
}

function Add-TrustedHostForTarget {
    param([string]$Target)
    # Ensures $Target is in TrustedHosts for the duration of this session.
    # Idempotent: no-op if already present or if '*' is set.
    # Records added hosts in $script:AddedTrustedHosts so Restore-TrustedHosts
    # can report what was cleaned up.

    if (-not $script:TrustedHostsManaged) { return }
    if ([string]::IsNullOrWhiteSpace($Target)) { return }

    try {
        $current = Get-CurrentTrustedHosts
        if ($null -eq $current) { return }

        # '*' means all hosts are already trusted -- nothing to add.
        if ($current -eq '*') { return }

        # Case-insensitive check: split on comma, trim each entry.
        $entries = $current -split ',' | ForEach-Object { $_.Trim() } |
                   Where-Object { $_ -ne '' }
        foreach ($e in $entries) {
            if ($e -ieq $Target) { return }   # already present; not ours to track
        }

        # Append target.
        $newValue = if ($current -eq '') { $Target } else { "$current,$Target" }
        $ok = Set-CurrentTrustedHosts -Value $newValue
        if ($ok) {
            $script:AddedTrustedHosts += $Target
            $lblStatus.Text = "Added $Target to TrustedHosts (removed on exit)."
        } else {
            $lblStatus.Text = "Could not update TrustedHosts for $Target (will proceed)."
        }
    } catch {
        $lblStatus.Text = "Could not update TrustedHosts: $($_.Exception.Message)"
        # Non-fatal -- scan continues regardless.
    }
}

function Restore-TrustedHosts {
    # Called from FormClosing (after all user-cancel checks).
    # Restores TrustedHosts to exactly the baseline captured at startup,
    # then deletes the marker file. Wrapped in try/catch (non-fatal).
    if (-not $script:TrustedHostsManaged) { return }
    try {
        Set-CurrentTrustedHosts -Value $script:OriginalTrustedHosts | Out-Null
    } catch { }
    try {
        if (Test-Path $script:TrustedHostsMarkerPath) {
            Remove-Item -LiteralPath $script:TrustedHostsMarkerPath -Force -ErrorAction SilentlyContinue
        }
    } catch { }
    $script:TrustedHostsManaged = $false
}

#endregion

#region -- Remote SYSTEM-Task Engine ------------------------------------------
#
# Worker source strings used by Invoke-RemoteSystemTask (defined inside each
# background scriptblock -- runspaces cannot see main-scope vars).
# Rules: single-quoted outer string means NO apostrophes anywhere inside;
# all inner strings must use double-quotes.  Placeholder tokens are replaced
# before Base64-encoding.
#
# $script:RemoteScanWorkerSrc  -- runs on the remote machine as SYSTEM
# $script:RemoteInstallWorkerSrc -- runs on the remote machine as SYSTEM

$script:RemoteScanWorkerSrc = '
$resultPath = "__RESULTPATH__"

function Get-WuPolicyInfo {
    $wuKey  = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
    $auKey  = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"

    $wuProps = Get-ItemProperty -Path $wuKey -ErrorAction SilentlyContinue
    $auProps = Get-ItemProperty -Path $auKey -ErrorAction SilentlyContinue

    $disabled      = [int]($wuProps.DisableWindowsUpdateAccess) -eq 1
    $noInternet    = [int]($wuProps.DoNotConnectToWindowsUpdateInternetLocations) -eq 1
    $useWsus       = [int]($auProps.UseWUServer) -eq 1
    $wsusUrl       = if ($wuProps.WUServer) { [string]$wuProps.WUServer } else { "" }
    $noAutoUpdate  = [int]($auProps.NoAutoUpdate) -eq 1

    $pauseExp   = if ($wuProps.PauseUpdatesExpiryTime)        { [string]$wuProps.PauseUpdatesExpiryTime }        else { "" }
    $pauseQual  = if ($wuProps.PauseQualityUpdatesStartTime)  { [string]$wuProps.PauseQualityUpdatesStartTime }  else { "" }
    $pauseFeat  = if ($wuProps.PauseFeatureUpdatesStartTime)  { [string]$wuProps.PauseFeatureUpdatesStartTime }  else { "" }
    $isPaused   = ($pauseExp -ne "") -or ($pauseQual -ne "") -or ($pauseFeat -ne "")

    $blocked = $disabled
    $summary = ""
    $detail  = ""

    if ($disabled) {
        $summary = "Windows Update access removed by Group Policy"
        $detail  = "DisableWindowsUpdateAccess=1 removes access to Windows Update on this machine."
    } elseif ($useWsus -and $wsusUrl -ne "") {
        $base = "Source: WSUS (" + $wsusUrl + ")"
        if ($isPaused) {
            $expiryNote = if ($pauseExp -ne "") { " until " + $pauseExp } else { "" }
            $summary = $base + " - updates paused" + $expiryNote
            $detail  = "WSUS source: " + $wsusUrl + ". Updates are paused" + $expiryNote + "."
        } else {
            $summary = $base
            if ($noInternet) { $summary = $summary + " (internet WU blocked)" }
            $detail  = ""
        }
    } else {
        $base = "Source: Microsoft Update"
        if ($isPaused) {
            $expiryNote = if ($pauseExp -ne "") { " until " + $pauseExp } else { "" }
            $summary = $base + "; updates paused" + $expiryNote
            $detail  = "Updates are paused" + $expiryNote + "."
        } elseif ($noInternet) {
            $summary = $base + "; internet Windows Update blocked by policy"
            $detail  = "DoNotConnectToWindowsUpdateInternetLocations=1 prevents connecting to Microsoft Update."
        } else {
            $summary = $base + "; no restrictions"
            $detail  = ""
        }
    }

    if ($noAutoUpdate -and -not $disabled) {
        $detail = ($detail + " NoAutoUpdate=1 (automatic updates off; manual install still allowed).").TrimStart()
    }

    return [PSCustomObject]@{
        PolicySummary = $summary
        PolicyBlocked = $blocked
        PolicyDetail  = $detail
    }
}

$out = [PSCustomObject]@{
    Updates         = @()
    IsPartial       = $false
    RebootPending   = $false
    OSVersion       = "Unknown"
    WuaSearchFailed = $false
    WuaSearchError  = ""
    PolicySummary   = ""
    PolicyBlocked   = $false
    PolicyDetail    = ""
    ErrorKind       = $null
    ErrorMessage    = $null
}
try {
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    if ($os) { $out.OSVersion = $os.Caption }

    $wuaOk = $false
    try {
        $session  = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $searcher.Online = $true
        $sr = $searcher.Search("IsInstalled=0 OR IsInstalled=1")
        $updates = @()
        foreach ($u in $sr.Updates) {
            $kb = if ($u.KBArticleIDs.Count -gt 0) { "KB" + $u.KBArticleIDs[0] } else { "N/A" }
            $uid = ""
            try { $uid = $u.Identity.UpdateID } catch {}
            $updates += [PSCustomObject]@{
                KB       = $kb
                Title    = $u.Title
                Status   = if ($u.IsInstalled) { "Installed" } else { "Missing" }
                Severity = if ($u.MsrcSeverity) { $u.MsrcSeverity } else { "N/A" }
                Date     = if ($u.LastDeploymentChangeTime) { $u.LastDeploymentChangeTime.ToString("yyyy-MM-dd") } else { "N/A" }
                SizeKB   = [math]::Round($u.MaxDownloadSize / 1KB, 0)
                Source   = "WUA"
                UpdateID = $uid
            }
        }
        $out.Updates = $updates
        $wuaOk = $true
    } catch {
        $out.WuaSearchFailed = $true
        $out.WuaSearchError  = $_.ToString()
    }

    if (-not $wuaOk) {
        $out.IsPartial = $true
        $hf = Get-CimInstance -ClassName Win32_QuickFixEngineering -ErrorAction Stop
        $out.Updates = @($hf | ForEach-Object {
            [PSCustomObject]@{
                KB       = $_.HotFixID
                Title    = $_.Description
                Status   = "Installed"
                Severity = "N/A"
                Date     = if ($_.InstalledOn) { $_.InstalledOn.ToString("yyyy-MM-dd") } else { "N/A" }
                SizeKB   = 0
                Source   = "WMI"
                UpdateID = ""
            }
        })
    }

    $keys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
    )
    $pr = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" `
        -Name PendingFileRenameOperations -ErrorAction SilentlyContinue).PendingFileRenameOperations
    foreach ($k in $keys) { if (Test-Path $k) { $out.RebootPending = $true } }
    if ($pr) { $out.RebootPending = $true }

    try {
        $pol = Get-WuPolicyInfo
        $out.PolicySummary = $pol.PolicySummary
        $out.PolicyBlocked = $pol.PolicyBlocked
        $out.PolicyDetail  = $pol.PolicyDetail
    } catch {}
} catch {
    $out.ErrorKind    = "SCAN_FAILED"
    $out.ErrorMessage = $_.ToString()
}
$out | ConvertTo-Json -Depth 5 | Out-File -FilePath $resultPath -Encoding UTF8 -Force
'

$script:RemoteInstallWorkerSrc = '
$resultPath = "__RESULTPATH__"
$UpdateIDs  = @(__UPDATEIDS__)
$out = [PSCustomObject]@{
    Items          = @()
    SkippedTitles  = @()
    NotFoundIDs    = @()
    RebootRequired = $false
    ErrorKind      = $null
    ErrorMessage   = $null
}
try {
    $session  = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()
    $searcher.Online = $true
    $search = $searcher.Search("IsInstalled=0")

    $toInstall = New-Object -ComObject Microsoft.Update.UpdateColl
    $found = @{}
    foreach ($u in $search.Updates) {
        $uid = ""
        try { $uid = $u.Identity.UpdateID } catch {}
        if ($UpdateIDs -contains $uid) {
            $found[$uid] = $true
            if ($u.InstallationBehavior.CanRequestUserInput) {
                $out.SkippedTitles += ($u.Title + " [requires user input]")
                continue
            }
            if (-not $u.EulaAccepted) {
                try { $u.AcceptEula() } catch {
                    $out.SkippedTitles += ($u.Title + " [EULA could not be accepted]")
                    continue
                }
            }
            [void]$toInstall.Add($u)
        }
    }
    foreach ($id in $UpdateIDs) {
        if (-not $found.ContainsKey($id)) { $out.NotFoundIDs += $id }
    }

    if ($toInstall.Count -eq 0) {
        $out.ErrorKind    = "NOTHING_TO_INSTALL"
        $out.ErrorMessage = "None of the selected updates are currently installable. Rescan and try again."
        $out | ConvertTo-Json -Depth 5 | Out-File -FilePath $resultPath -Encoding UTF8 -Force
        return
    }

    $downloader = $session.CreateUpdateDownloader()
    $downloader.Updates = $toInstall
    $dl = $downloader.Download()
    if ($dl.ResultCode -ne 2 -and $dl.ResultCode -ne 3) {
        $out.ErrorKind    = "DOWNLOAD_FAILED"
        $out.ErrorMessage = ("Download failed - result code {0}, HResult 0x{1:X8}" -f $dl.ResultCode, $dl.HResult)
        $out | ConvertTo-Json -Depth 5 | Out-File -FilePath $resultPath -Encoding UTF8 -Force
        return
    }

    $installer = $session.CreateUpdateInstaller()
    $installer.Updates = $toInstall
    $inst = $installer.Install()
    $out.RebootRequired = [bool]$inst.RebootRequired

    for ($i = 0; $i -lt $toInstall.Count; $i++) {
        $ur = $inst.GetUpdateResult($i)
        if     ($ur.ResultCode -eq 2) { $outcome = "Succeeded" }
        elseif ($ur.ResultCode -eq 3) { $outcome = "Succeeded with errors" }
        elseif ($ur.ResultCode -eq 4) { $outcome = "Failed" }
        elseif ($ur.ResultCode -eq 5) { $outcome = "Aborted" }
        else                          { $outcome = "Unknown ($($ur.ResultCode))" }
        $item = $toInstall.Item($i)
        $kb = if ($item.KBArticleIDs.Count -gt 0) { "KB" + $item.KBArticleIDs[0] } else { "N/A" }
        $out.Items += [PSCustomObject]@{
            KB      = $kb
            Title   = $item.Title
            Outcome = $outcome
            HResult = ("0x{0:X8}" -f $ur.HResult)
        }
    }
} catch {
    $out.ErrorKind    = "INSTALL_FAILED"
    $out.ErrorMessage = $_.ToString()
}
$out | ConvertTo-Json -Depth 5 | Out-File -FilePath $resultPath -Encoding UTF8 -Force
'

#endregion

#region -- Poll Timer Tick Handler --------------------------------------------
# Fix 1: Tick handler is defined at script scope so it can reference
# script-scoped UI controls and script variables without closure issues.

$scanPollTick = {
    if (-not $script:PollHandle.IsCompleted) { return }

    # Scan finished -- stop polling first
    $script:PollTimer.Stop()

    # Collect result before cleanup disposes the PS instance
    $scanResult = $null
    $psErrors   = $null
    $hadErrors  = $false
    try {
        $scanResult = $script:RunningPS.EndInvoke($script:PollHandle)
        $psErrors   = $script:RunningPS.Streams.Error
        $hadErrors  = $script:RunningPS.HadErrors
    } catch {
        $hadErrors = $true
        $psErrors  = @($_)
    } finally {
        Stop-AsyncCleanup
    }

    # Restore UI
    $pnlProgress.Visible = $false
    $btnScan.Enabled     = $true
    $script:CanInstall   = $false
    $script:CanReboot    = $false
    Update-InstallButton
    Update-RebootButton

    # Handle PowerShell stream-level errors
    if ($hadErrors -and ($null -eq $scanResult -or $scanResult.Count -eq 0)) {
        $errTxt = if ($psErrors -and $psErrors.Count -gt 0) {
            $psErrors[0].ToString()
        } else {
            'Unknown error'
        }
        Show-Error "Scan error: $errTxt"
        $lblStatus.Text = 'Scan failed.'
        return
    }

    # EndInvoke returns a collection; grab first element
    $res = if ($scanResult -is [System.Collections.IList]) { $scanResult[0] } else { $scanResult }

    if ($null -eq $res) {
        Show-Error 'Scan returned no data.'
        $lblStatus.Text = 'Scan failed.'
        return
    }

    # Application-level error flags returned in the result object
    if ($res.ErrorKind -eq 'UNREACHABLE') {
        Show-Error "WinRM unreachable on '$($res.ErrorMessage)' - enable it with: winrm quickconfig"
        $lblStatus.Text = 'Scan failed.'
        return
    }
    if ($res.ErrorKind -eq 'SCAN_FAILED') {
        Show-Error "Scan failed: $($res.ErrorMessage)"
        $lblStatus.Text = 'Scan failed.'
        # If the failure looks credential-related, drop the cached credential so
        # the next scan re-prompts instead of silently reusing bad credentials.
        if ($res.ErrorMessage -match 'Access is denied|logon|authentication|user name or password|password is incorrect|credential') {
            $script:LastCredential = $null
            $script:LastCredTarget = ''
        }
        return
    }

    # Success path
    $script:ScanResults   = @($res.Updates)
    $script:IsPartialData = [bool]$res.IsPartial
    $script:RebootPending = [bool]$res.RebootPending
    $script:ScannedOS     = $res.OSVersion
    # Fix SSL: store the SSL flag actually used during scan so install reuses it correctly
    if ($res.IsRemote) { $script:LastUseSSL = [bool]$res.EffectiveUseSSL }

    # Plan item 7a: CanInstall is true for any non-partial scan (local or remote full WUA)
    $script:CanInstall = (-not $res.IsPartial) -and ($null -eq $res.ErrorKind)
    # CanReboot: enabled for any successful remote scan (even partial WMI-only), no error
    $script:CanReboot  = [bool]$res.IsRemote -and ($null -eq $res.ErrorKind)

    # Policy indicator label
    $polSummary = if ($res.PolicySummary) { [string]$res.PolicySummary } else { '' }
    $polBlocked = [bool]$res.PolicyBlocked
    $lblPolicy.Text = $polSummary
    if ($polBlocked) {
        $lblPolicy.ForeColor = [System.Drawing.Color]::FromArgb(198, 40, 40)
    } elseif ($polSummary -match 'paused|blocked|restricted') {
        $lblPolicy.ForeColor = [System.Drawing.Color]::FromArgb(230, 81, 0)
    } else {
        $lblPolicy.ForeColor = [System.Drawing.Color]::Gray
    }

    # Plan item 7b/7c: banner logic
    if ($script:IsPartialData -and (-not $res.IsRemote)) {
        # Local scan fell back to WMI
        $bannerMsg = 'WUA COM unavailable locally - showing installed hotfixes from WMI only.'
        $polDetail = if ($res.PolicyDetail) { [string]$res.PolicyDetail } else { '' }
        if ($polDetail -ne '') { $bannerMsg = $bannerMsg + '  Policy: ' + $polDetail }
        Show-Error $bannerMsg
    } elseif ($script:IsPartialData -and $res.IsRemote) {
        # Plan item 7b: remote partial - rewording to not say WUA over WinRM is impossible
        $bannerMsg = 'Accurate WUA scan unavailable on the remote target (Task Scheduler/agent path failed) - showing installed hotfixes from WMI only.'
        if ([bool]$res.WuaSearchFailed) {
            $wuaErr = if ($res.WuaSearchError) { Translate-WuaError ([string]$res.WuaSearchError) } else { '' }
            $bannerMsg += ' The remote machine ran the scan but Windows Update returned an error (commonly a per-user proxy or unreachable WSUS under the SYSTEM account).'
            if ($wuaErr -ne '') { $bannerMsg += '  Detail: ' + $wuaErr }
        }
        $polDetail = if ($res.PolicyDetail) { [string]$res.PolicyDetail } else { '' }
        if ($polDetail -ne '') { $bannerMsg = $bannerMsg + '  Policy: ' + $polDetail }
        Show-Error $bannerMsg
    } elseif ($polBlocked) {
        $polDetail = if ($res.PolicyDetail) { [string]$res.PolicyDetail } else { '' }
        $bannerMsg = if ($polDetail -ne '') { $polDetail } else { [string]$res.PolicySummary }
        Show-Error $bannerMsg
    } else {
        Hide-Error
    }

    Update-Summary
    Update-Grid    # calls Update-InstallButton and Update-RebootButton

    $btnExportCsv.Enabled  = ($script:ScanResults.Count -gt 0)
    $btnExportHtml.Enabled = ($script:ScanResults.Count -gt 0)
    $countStr = $script:ScanResults.Count

    # Plan item 7c: remove the "Install is unavailable for remote targets" message
    $lblStatus.Text = "Scan complete - $countStr updates found."
}

# Poll tick for the install operation; same structure as $scanPollTick.
$installPollTick = {
    if (-not $script:PollHandle.IsCompleted) { return }

    $script:PollTimer.Stop()

    $installResult = $null
    $psErrors      = $null
    $hadErrors     = $false
    try {
        $installResult = $script:RunningPS.EndInvoke($script:PollHandle)
        $psErrors      = $script:RunningPS.Streams.Error
        $hadErrors     = $script:RunningPS.HadErrors
    } catch {
        $hadErrors = $true
        $psErrors  = @($_)
    } finally {
        Stop-AsyncCleanup
    }

    # Restore UI
    $pnlProgress.Visible = $false
    $btnScan.Enabled     = $true
    Update-InstallButton
    Update-RebootButton

    if ($hadErrors -and ($null -eq $installResult -or $installResult.Count -eq 0)) {
        $errTxt = if ($psErrors -and $psErrors.Count -gt 0) {
            $psErrors[0].ToString()
        } else {
            'Unknown error'
        }
        Show-Error "Install error: $(Translate-WuaError $errTxt)"
        $lblStatus.Text = 'Install failed.'
        return
    }

    $res = if ($installResult -is [System.Collections.IList]) { $installResult[0] } else { $installResult }

    if ($null -eq $res) {
        Show-Error 'Install returned no data.'
        $lblStatus.Text = 'Install failed.'
        return
    }

    if ($res.ErrorKind) {
        Show-Error "Install failed: $(Translate-WuaError ([string]$res.ErrorMessage))"
        $lblStatus.Text = 'Install failed.'
        return
    }

    # Success path: tally outcomes and surface reboot state
    $succeeded = @($res.Items | Where-Object { $_.Outcome -like 'Succeeded*' }).Count
    $failed    = @($res.Items | Where-Object { $_.Outcome -notlike 'Succeeded*' }).Count
    $skipped   = @($res.SkippedTitles).Count

    $script:RebootPending = $script:RebootPending -or [bool]$res.RebootRequired
    Update-Summary

    if ($failed -gt 0) {
        Show-Error "$failed of $($res.Items.Count) updates failed to install - see the results dialog."
    } elseif ($res.RebootRequired) {
        Show-Error 'Reboot required to complete installation.'
    }
    $lblStatus.Text = "Install complete: $succeeded succeeded, $failed failed, $skipped skipped."

    # Build the results report
    $lines = @("Install Results", ("-" * 60))
    foreach ($item in $res.Items) {
        $line = "$($item.Outcome.PadRight(22)) $($item.KB.PadRight(11)) $($item.Title)"
        if ($item.Outcome -notlike 'Succeeded*') {
            $hresultMsg = Translate-WuaError ([string]$item.HResult)
            $line += "  [HResult $hresultMsg]"
        }
        $lines += $line
    }
    foreach ($s in $res.SkippedTitles) { $lines += "Skipped                $s" }
    if (@($res.NotFoundIDs).Count -gt 0) {
        $lines += ''
        $lines += "$(@($res.NotFoundIDs).Count) selected update(s) were no longer offered by Windows Update" +
            ' (superseded, expired, or already installed) and were not processed.'
    }
    if ($res.RebootRequired) {
        $lines += ''
        $lines += 'A reboot is required to complete the installation. Updates may still' +
            ' show as Missing until the machine restarts.'
    }

    $dlg                 = New-Object System.Windows.Forms.Form
    $dlg.Text            = 'Install Results'
    $dlg.Size            = New-Object System.Drawing.Size(640, 360)
    $dlg.StartPosition   = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false

    $txt          = New-Object System.Windows.Forms.RichTextBox
    $txt.Dock     = 'Fill'
    $txt.ReadOnly = $true
    $txt.Font     = New-Object System.Drawing.Font('Consolas', 9)
    $txt.Text     = $lines -join "`r`n"

    $btnCloseDlg      = New-Object System.Windows.Forms.Button
    $btnCloseDlg.Text = 'Close'
    $btnCloseDlg.Dock = 'Bottom'
    $btnCloseDlg.Add_Click({ $dlg.Close() })

    $dlg.Controls.AddRange(@($txt, $btnCloseDlg))
    $dlg.ShowDialog($form) | Out-Null

    # Plan item 8: auto-rescan uses PerformClick so it rescans whatever target
    # (local or remote) is currently selected -- do NOT force $rbLocal.Checked
    $btnScan.PerformClick()
}

#endregion

#region -- Async Scan (Runspace + BeginInvoke + WinForms Timer) ---------------
#
# Fix 1: Replaces BackgroundWorker entirely.
# A dedicated runspace is created per scan. All needed functions are defined
# INSIDE the background scriptblock -- the new runspace cannot see functions
# from the main script scope.  A WinForms Timer polls the IAsyncResult on the
# UI thread (no cross-thread UI access).
#
# Remote path: Invoke-RemoteSystemTask is defined INSIDE the scriptblock
# (duplicate body) so the background runspace can call it.  It registers a
# SYSTEM scheduled task on the target via CIM, waits for completion, reads
# back a JSON result file over Invoke-Command (WinRM), and cleans up.

function Start-ScanAsync {
    param(
        [string]  $Target,
        [bool]    $IsRemote,
        [System.Management.Automation.PSCredential]$Credential,
        [bool]    $UseSSL,
        [string]  $ScanWorkerSrc,
        [string]  $InstallWorkerSrc
    )

    # Background script -- entirely self-contained; no access to main scope.
    # Single-quoted: NO apostrophes inside; all inner strings use double-quotes.
    $bgScript = [scriptblock]::Create('
        param($Target, $IsRemote, $Credential, $UseSSL, $ScanWorkerSrc, $InstallWorkerSrc)

        #-- Invoke-RemoteSystemTask: registers+runs+polls+reads a SYSTEM task --
        # This function body is duplicated in the install bgScript (runspaces
        # cannot see main-scope functions).
        function Invoke-RemoteSystemTask {
            param(
                [string]$Target,
                [System.Management.Automation.PSCredential]$Credential,
                [bool]$UseSSL,
                [string]$Mode,
                [string]$WorkerSrc,
                [string[]]$UpdateIDs,
                [int]$TimeoutSec
            )

            $cimOpts   = $null
            $cimSess   = $null
            $icParams  = @{ ComputerName = $Target; ErrorAction = "Stop" }

            try {
                if ($UseSSL) {
                    $cimOpts = New-CimSessionOption -UseSsl
                }
                $cimSessParams = @{ ComputerName = $Target; ErrorAction = "Stop" }
                if ($Credential) { $cimSessParams.Credential = $Credential }
                if ($UseSSL)     { $cimSessParams.SessionOption = $cimOpts }
                $cimSess = New-CimSession @cimSessParams

                if ($Credential) { $icParams.Credential = $Credential }
                if ($UseSSL)     { $icParams.UseSSL     = $true }

                # Generate unique names
                $guid     = [guid]::NewGuid().ToString("N")
                $taskName = "WinUpdateChecker_" + $Mode + "_" + $guid
                $resPath  = "C:\Windows\Temp\WUC_" + $guid + ".json"

                # Orphan cleanup: remove stale WinUpdateChecker tasks and result files
                try {
                    $staleTasks = Get-ScheduledTask -CimSession $cimSess -TaskName "WinUpdateChecker_*" -ErrorAction SilentlyContinue
                    foreach ($t in $staleTasks) {
                        Unregister-ScheduledTask -CimSession $cimSess -TaskName $t.TaskName -Confirm:$false -ErrorAction SilentlyContinue
                    }
                } catch {}
                try {
                    $cleanBlock = { Get-ChildItem "C:\Windows\Temp\WUC_*.json" -ErrorAction SilentlyContinue |
                        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-1) } |
                        Remove-Item -Force -ErrorAction SilentlyContinue }
                    $icParams.ScriptBlock = $cleanBlock
                    Invoke-Command @icParams | Out-Null
                } catch {}

                # Bake placeholders into the worker source
                $baked = $WorkerSrc.Replace("__RESULTPATH__", $resPath)
                if ($UpdateIDs -and $UpdateIDs.Count -gt 0) {
                    $quotedIDs = ($UpdateIDs | ForEach-Object { """$_""" }) -join ","
                    $baked = $baked.Replace("__UPDATEIDS__", $quotedIDs)
                }

                # Base64-encode for -EncodedCommand (UTF-16LE)
                $bytes  = [System.Text.Encoding]::Unicode.GetBytes($baked)
                $b64    = [Convert]::ToBase64String($bytes)
                $cmdArg = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand " + $b64

                # Register and start the scheduled task on the target via CIM
                $action    = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $cmdArg
                $principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -RunLevel Highest -LogonType ServiceAccount
                $execLimit = New-TimeSpan -Seconds ($TimeoutSec + 60)
                $settings  = New-ScheduledTaskSettingsSet -ExecutionTimeLimit $execLimit -MultipleInstances IgnoreNew

                Register-ScheduledTask -CimSession $cimSess -TaskName $taskName `
                    -Action $action -Principal $principal -Settings $settings `
                    -Force -ErrorAction Stop | Out-Null
                Start-ScheduledTask -CimSession $cimSess -TaskName $taskName -ErrorAction Stop

                # Poll until task finishes or deadline exceeded
                $deadline = (Get-Date).AddSeconds($TimeoutSec)
                do {
                    Start-Sleep -Seconds 2
                    if ((Get-Date) -gt $deadline) {
                        throw "Remote task timed out after " + $TimeoutSec + " seconds."
                    }
                    $info  = Get-ScheduledTaskInfo -CimSession $cimSess -TaskName $taskName -ErrorAction SilentlyContinue
                    $state = (Get-ScheduledTask -CimSession $cimSess -TaskName $taskName -ErrorAction SilentlyContinue).State
                } while ($state -eq "Running" -or $info.LastTaskResult -eq 267009)

                # Read result JSON back over WinRM
                $rawJson = $null
                try {
                    $icParams.ScriptBlock = [scriptblock]::Create(
                        "if (Test-Path """ + $resPath + """) { Get-Content -Raw -LiteralPath """ + $resPath + """ } else { `$null }"
                    )
                    $rawJson = Invoke-Command @icParams
                } catch {}

                if (-not $rawJson) {
                    $lastResult = if ($info) { $info.LastTaskResult } else { "unknown" }
                    throw "Remote task produced no output file. LastTaskResult=" + $lastResult
                }

                return ($rawJson | ConvertFrom-Json)

            } finally {
                # Cleanup: unregister task, remove result file, close CIM session
                if ($cimSess) {
                    try { Unregister-ScheduledTask -CimSession $cimSess -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue } catch {}
                    try { Remove-CimSession $cimSess -ErrorAction SilentlyContinue } catch {}
                }
                try {
                    $icParams.ScriptBlock = [scriptblock]::Create(
                        "Remove-Item -LiteralPath """ + $resPath + """ -Force -ErrorAction SilentlyContinue"
                    )
                    Invoke-Command @icParams | Out-Null
                } catch {}
            }
        }

        #-- inner helpers for local path --

        function Local-TestPendingReboot {
            $keys = @(
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
            )
            $pendingRename = (Get-ItemProperty `
                -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" `
                -Name PendingFileRenameOperations `
                -ErrorAction SilentlyContinue).PendingFileRenameOperations
            foreach ($key in $keys) { if (Test-Path $key) { return $true } }
            if ($pendingRename) { return $true }
            return $false
        }

        function Local-GetWuaUpdates {
            $updates = @()
            $session  = New-Object -ComObject Microsoft.Update.Session
            $searcher = $session.CreateUpdateSearcher()
            $searcher.Online = $true
            $result = $searcher.Search("IsInstalled=0 OR IsInstalled=1")
            foreach ($u in $result.Updates) {
                $kb = if ($u.KBArticleIDs.Count -gt 0) { "KB$($u.KBArticleIDs[0])" } else { "N/A" }
                $updates += [PSCustomObject]@{
                    KB       = $kb
                    Title    = $u.Title
                    Status   = if ($u.IsInstalled) { "Installed" } else { "Missing" }
                    Severity = if ($u.MsrcSeverity) { $u.MsrcSeverity } else { "N/A" }
                    Date     = if ($u.LastDeploymentChangeTime) { $u.LastDeploymentChangeTime.ToString("yyyy-MM-dd") } else { "N/A" }
                    SizeKB   = [math]::Round($u.MaxDownloadSize / 1KB, 0)
                    Source   = "WUA"
                    UpdateID = $u.Identity.UpdateID
                }
            }
            return $updates
        }

        function Local-GetWmiUpdates {
            $hotfixes = Get-CimInstance -ClassName Win32_QuickFixEngineering -ErrorAction Stop
            return $hotfixes | ForEach-Object {
                [PSCustomObject]@{
                    KB       = $_.HotFixID
                    Title    = $_.Description
                    Status   = "Installed"
                    Severity = "N/A"
                    Date     = if ($_.InstalledOn) { $_.InstalledOn.ToString("yyyy-MM-dd") } else { "N/A" }
                    SizeKB   = 0
                    Source   = "WMI"
                    UpdateID = ""
                }
            }
        }

        function Local-GetWuPolicyInfo {
            $wuKey  = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
            $auKey  = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"

            $wuProps = Get-ItemProperty -Path $wuKey -ErrorAction SilentlyContinue
            $auProps = Get-ItemProperty -Path $auKey -ErrorAction SilentlyContinue

            $disabled     = [int]($wuProps.DisableWindowsUpdateAccess) -eq 1
            $noInternet   = [int]($wuProps.DoNotConnectToWindowsUpdateInternetLocations) -eq 1
            $useWsus      = [int]($auProps.UseWUServer) -eq 1
            $wsusUrl      = if ($wuProps.WUServer) { [string]$wuProps.WUServer } else { "" }
            $noAutoUpdate = [int]($auProps.NoAutoUpdate) -eq 1

            $pauseExp  = if ($wuProps.PauseUpdatesExpiryTime)        { [string]$wuProps.PauseUpdatesExpiryTime }        else { "" }
            $pauseQual = if ($wuProps.PauseQualityUpdatesStartTime)  { [string]$wuProps.PauseQualityUpdatesStartTime }  else { "" }
            $pauseFeat = if ($wuProps.PauseFeatureUpdatesStartTime)  { [string]$wuProps.PauseFeatureUpdatesStartTime }  else { "" }
            $isPaused  = ($pauseExp -ne "") -or ($pauseQual -ne "") -or ($pauseFeat -ne "")

            $blocked = $disabled
            $summary = ""
            $detail  = ""

            if ($disabled) {
                $summary = "Windows Update access removed by Group Policy"
                $detail  = "DisableWindowsUpdateAccess=1 removes access to Windows Update on this machine."
            } elseif ($useWsus -and $wsusUrl -ne "") {
                $base = "Source: WSUS (" + $wsusUrl + ")"
                if ($isPaused) {
                    $expiryNote = if ($pauseExp -ne "") { " until " + $pauseExp } else { "" }
                    $summary = $base + " - updates paused" + $expiryNote
                    $detail  = "WSUS source: " + $wsusUrl + ". Updates are paused" + $expiryNote + "."
                } else {
                    $summary = $base
                    if ($noInternet) { $summary = $summary + " (internet WU blocked)" }
                    $detail  = ""
                }
            } else {
                $base = "Source: Microsoft Update"
                if ($isPaused) {
                    $expiryNote = if ($pauseExp -ne "") { " until " + $pauseExp } else { "" }
                    $summary = $base + "; updates paused" + $expiryNote
                    $detail  = "Updates are paused" + $expiryNote + "."
                } elseif ($noInternet) {
                    $summary = $base + "; internet Windows Update blocked by policy"
                    $detail  = "DoNotConnectToWindowsUpdateInternetLocations=1 prevents connecting to Microsoft Update."
                } else {
                    $summary = $base + "; no restrictions"
                    $detail  = ""
                }
            }

            if ($noAutoUpdate -and -not $disabled) {
                $detail = ($detail + " NoAutoUpdate=1 (automatic updates off; manual install still allowed).").TrimStart()
            }

            return [PSCustomObject]@{
                PolicySummary = $summary
                PolicyBlocked = $blocked
                PolicyDetail  = $detail
            }
        }

        #-- result skeleton --
        $result = [PSCustomObject]@{
            Updates         = @()
            IsPartial       = $false
            IsRemote        = $IsRemote
            RebootPending   = $false
            OSVersion       = "Unknown"
            WuaSearchFailed = $false
            EffectiveUseSSL = $false    # Fix SSL: probed value returned so install can reuse it
            PolicySummary   = ""
            PolicyBlocked   = $false
            PolicyDetail    = ""
            ErrorKind       = $null
            ErrorMessage    = $null
        }

        try {
            if ($IsRemote) {
                # Probe 5985 then 5986
                $port5985 = $false
                $port5986 = $false
                try {
                    $t = Test-NetConnection -ComputerName $Target -Port 5985 `
                        -WarningAction SilentlyContinue -ErrorAction Stop
                    $port5985 = $t.TcpTestSucceeded
                } catch { $port5985 = $false }

                if (-not $port5985) {
                    try {
                        $t = Test-NetConnection -ComputerName $Target -Port 5986 `
                            -WarningAction SilentlyContinue -ErrorAction Stop
                        $port5986 = $t.TcpTestSucceeded
                    } catch { $port5986 = $false }
                }

                if (-not $port5985 -and -not $port5986) {
                    $result.ErrorKind    = "UNREACHABLE"
                    $result.ErrorMessage = $Target
                    return $result
                }

                $useSSLFlag = ($port5986 -and -not $port5985) -or $UseSSL
                $result.EffectiveUseSSL = $useSSLFlag   # Fix SSL: carry probed value back to caller

                # Try the SYSTEM scheduled task WUA path
                try {
                    $tr = Invoke-RemoteSystemTask -Target $Target -Credential $Credential `
                        -UseSSL $useSSLFlag -Mode "Scan" -WorkerSrc $ScanWorkerSrc `
                        -TimeoutSec 300

                    $result.OSVersion     = if ($tr.OSVersion)     { $tr.OSVersion }     else { "Unknown" }
                    $result.RebootPending = [bool]$tr.RebootPending
                    $result.Updates       = @($tr.Updates)
                    $result.PolicySummary = if ($tr.PolicySummary) { [string]$tr.PolicySummary } else { "" }
                    $result.PolicyBlocked = [bool]$tr.PolicyBlocked
                    $result.PolicyDetail  = if ($tr.PolicyDetail)  { [string]$tr.PolicyDetail }  else { "" }

                    if ([bool]$tr.WuaSearchFailed) {
                        # WUA ran under SYSTEM but the search itself failed (proxy/WSUS)
                        # Fall into the WMI partial path; stamp the flag so the banner
                        # can explain the specific reason.
                        $result.IsPartial       = $true
                        $result.WuaSearchFailed = $true
                        # tr.Updates already contains WMI fallback rows from the worker
                    } else {
                        $result.IsPartial = [bool]$tr.IsPartial
                    }
                } catch {
                    # Invoke-RemoteSystemTask failed entirely (CIM unreachable, task error, etc.)
                    # Fall back to WinRM WMI
                    $result.IsPartial = $true

                    $icParams = @{ ComputerName = $Target; ErrorAction = "Stop" }
                    if ($useSSLFlag) { $icParams.UseSSL     = $true }
                    if ($Credential) { $icParams.Credential = $Credential }

                    # OS caption
                    try {
                        $icParams.ScriptBlock = {
                            (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
                        }
                        $osCaption = Invoke-Command @icParams
                        $result.OSVersion = if ($osCaption) { $osCaption } else { "Unknown" }
                    } catch {
                        $result.OSVersion = "Unknown"
                    }

                    # WMI hotfix fallback
                    try {
                        $icParams.ScriptBlock = {
                            $hotfixes = Get-CimInstance -ClassName Win32_QuickFixEngineering -ErrorAction Stop
                            $hotfixes | ForEach-Object {
                                [PSCustomObject]@{
                                    KB       = $_.HotFixID
                                    Title    = $_.Description
                                    Status   = "Installed"
                                    Severity = "N/A"
                                    Date     = if ($_.InstalledOn) { $_.InstalledOn.ToString("yyyy-MM-dd") } else { "N/A" }
                                    SizeKB   = 0
                                    Source   = "WMI"
                                    UpdateID = ""
                                }
                            }
                        }
                        $result.Updates = @(Invoke-Command @icParams)
                    } catch {
                        $result.ErrorKind    = "SCAN_FAILED"
                        $result.ErrorMessage = $_.ToString()
                        return $result
                    }

                    # Remote reboot check
                    try {
                        $icParams.ScriptBlock = {
                            $keys = @(
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
                            )
                            $pr = (Get-ItemProperty `
                                -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" `
                                -Name PendingFileRenameOperations `
                                -ErrorAction SilentlyContinue).PendingFileRenameOperations
                            foreach ($k in $keys) { if (Test-Path $k) { return $true } }
                            if ($pr) { return $true }
                            return $false
                        }
                        $result.RebootPending = Invoke-Command @icParams
                    } catch {
                        $result.RebootPending = $false
                    }
                }

            } else {
                # LOCAL scan
                $result.OSVersion = (Get-CimInstance Win32_OperatingSystem `
                    -ErrorAction SilentlyContinue).Caption
                if (-not $result.OSVersion) { $result.OSVersion = "Unknown" }

                # Try WUA COM locally
                try {
                    $result.Updates   = @(Local-GetWuaUpdates)
                    $result.IsPartial = $false
                } catch {
                    # Fall back to WMI
                    $result.IsPartial = $true
                    try {
                        $result.Updates = @(Local-GetWmiUpdates)
                    } catch {
                        $result.ErrorKind    = "SCAN_FAILED"
                        $result.ErrorMessage = $_.ToString()
                        return $result
                    }
                }

                $result.RebootPending = Local-TestPendingReboot

                try {
                    $pol = Local-GetWuPolicyInfo
                    $result.PolicySummary = $pol.PolicySummary
                    $result.PolicyBlocked = $pol.PolicyBlocked
                    $result.PolicyDetail  = $pol.PolicyDetail
                } catch {}
            }
        } catch {
            $result.ErrorKind    = "SCAN_FAILED"
            $result.ErrorMessage = $_.ToString()
        }

        return $result
    ')

    # Create runspace + PowerShell instance
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.Open()

    $ps           = [powershell]::Create()
    $ps.Runspace  = $rs

    [void]$ps.AddScript($bgScript)
    [void]$ps.AddArgument($Target)
    [void]$ps.AddArgument($IsRemote)
    [void]$ps.AddArgument($Credential)
    [void]$ps.AddArgument($UseSSL)
    [void]$ps.AddArgument($ScanWorkerSrc)
    [void]$ps.AddArgument($InstallWorkerSrc)

    $script:ActiveOp   = 'Scan'
    $script:RunningPS  = $ps
    $script:RunningRS  = $rs
    $script:PollHandle = $ps.BeginInvoke()

    # WinForms Timer polls on the UI thread (no cross-thread access)
    $timer          = New-Object System.Windows.Forms.Timer
    $timer.Interval = 500
    $timer.Add_Tick($scanPollTick)
    $script:PollTimer = $timer
    $timer.Start()
}

function Start-InstallAsync {
    param(
        [string[]]$UpdateIDs,
        [bool]    $IsRemote,
        [string]  $Target,
        [System.Management.Automation.PSCredential]$Credential,
        [bool]    $UseSSL,
        [string]  $ScanWorkerSrc,
        [string]  $InstallWorkerSrc
    )

    # Background install script -- self-contained, same pattern as the scan.
    # Single-quoted: NO apostrophes inside; all inner strings use double-quotes.
    $bgScript = [scriptblock]::Create('
        param($UpdateIDs, $IsRemote, $Target, $Credential, $UseSSL, $ScanWorkerSrc, $InstallWorkerSrc)

        #-- Invoke-RemoteSystemTask: duplicated body (runspace cannot see main scope) --
        function Invoke-RemoteSystemTask {
            param(
                [string]$Target,
                [System.Management.Automation.PSCredential]$Credential,
                [bool]$UseSSL,
                [string]$Mode,
                [string]$WorkerSrc,
                [string[]]$UpdateIDs,
                [int]$TimeoutSec
            )

            $cimOpts   = $null
            $cimSess   = $null
            $icParams  = @{ ComputerName = $Target; ErrorAction = "Stop" }

            try {
                if ($UseSSL) {
                    $cimOpts = New-CimSessionOption -UseSsl
                }
                $cimSessParams = @{ ComputerName = $Target; ErrorAction = "Stop" }
                if ($Credential) { $cimSessParams.Credential = $Credential }
                if ($UseSSL)     { $cimSessParams.SessionOption = $cimOpts }
                $cimSess = New-CimSession @cimSessParams

                if ($Credential) { $icParams.Credential = $Credential }
                if ($UseSSL)     { $icParams.UseSSL     = $true }

                $guid     = [guid]::NewGuid().ToString("N")
                $taskName = "WinUpdateChecker_" + $Mode + "_" + $guid
                $resPath  = "C:\Windows\Temp\WUC_" + $guid + ".json"

                # Orphan cleanup
                try {
                    $staleTasks = Get-ScheduledTask -CimSession $cimSess -TaskName "WinUpdateChecker_*" -ErrorAction SilentlyContinue
                    foreach ($t in $staleTasks) {
                        Unregister-ScheduledTask -CimSession $cimSess -TaskName $t.TaskName -Confirm:$false -ErrorAction SilentlyContinue
                    }
                } catch {}
                try {
                    $cleanBlock = { Get-ChildItem "C:\Windows\Temp\WUC_*.json" -ErrorAction SilentlyContinue |
                        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-1) } |
                        Remove-Item -Force -ErrorAction SilentlyContinue }
                    $icParams.ScriptBlock = $cleanBlock
                    Invoke-Command @icParams | Out-Null
                } catch {}

                # Bake placeholders
                $baked = $WorkerSrc.Replace("__RESULTPATH__", $resPath)
                if ($UpdateIDs -and $UpdateIDs.Count -gt 0) {
                    $quotedIDs = ($UpdateIDs | ForEach-Object { """$_""" }) -join ","
                    $baked = $baked.Replace("__UPDATEIDS__", $quotedIDs)
                }

                $bytes  = [System.Text.Encoding]::Unicode.GetBytes($baked)
                $b64    = [Convert]::ToBase64String($bytes)
                $cmdArg = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand " + $b64

                $action    = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $cmdArg
                $principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -RunLevel Highest -LogonType ServiceAccount
                $execLimit = New-TimeSpan -Seconds ($TimeoutSec + 60)
                $settings  = New-ScheduledTaskSettingsSet -ExecutionTimeLimit $execLimit -MultipleInstances IgnoreNew

                Register-ScheduledTask -CimSession $cimSess -TaskName $taskName `
                    -Action $action -Principal $principal -Settings $settings `
                    -Force -ErrorAction Stop | Out-Null
                Start-ScheduledTask -CimSession $cimSess -TaskName $taskName -ErrorAction Stop

                # Poll
                $deadline = (Get-Date).AddSeconds($TimeoutSec)
                do {
                    Start-Sleep -Seconds 2
                    if ((Get-Date) -gt $deadline) {
                        throw "Remote task timed out after " + $TimeoutSec + " seconds."
                    }
                    $info  = Get-ScheduledTaskInfo -CimSession $cimSess -TaskName $taskName -ErrorAction SilentlyContinue
                    $state = (Get-ScheduledTask -CimSession $cimSess -TaskName $taskName -ErrorAction SilentlyContinue).State
                } while ($state -eq "Running" -or $info.LastTaskResult -eq 267009)

                $rawJson = $null
                try {
                    $icParams.ScriptBlock = [scriptblock]::Create(
                        "if (Test-Path """ + $resPath + """) { Get-Content -Raw -LiteralPath """ + $resPath + """ } else { `$null }"
                    )
                    $rawJson = Invoke-Command @icParams
                } catch {}

                if (-not $rawJson) {
                    $lastResult = if ($info) { $info.LastTaskResult } else { "unknown" }
                    throw "Remote task produced no output file. LastTaskResult=" + $lastResult
                }

                return ($rawJson | ConvertFrom-Json)

            } finally {
                if ($cimSess) {
                    try { Unregister-ScheduledTask -CimSession $cimSess -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue } catch {}
                    try { Remove-CimSession $cimSess -ErrorAction SilentlyContinue } catch {}
                }
                try {
                    $icParams.ScriptBlock = [scriptblock]::Create(
                        "Remove-Item -LiteralPath """ + $resPath + """ -Force -ErrorAction SilentlyContinue"
                    )
                    Invoke-Command @icParams | Out-Null
                } catch {}
            }
        }

        #-- result skeleton --
        $result = [PSCustomObject]@{
            Items          = @()
            SkippedTitles  = @()
            NotFoundIDs    = @()
            RebootRequired = $false
            ErrorKind      = $null
            ErrorMessage   = $null
        }

        if ($IsRemote) {
            # Remote install via SYSTEM scheduled task
            try {
                $tr = Invoke-RemoteSystemTask -Target $Target -Credential $Credential `
                    -UseSSL $UseSSL -Mode "Install" -WorkerSrc $InstallWorkerSrc `
                    -UpdateIDs $UpdateIDs -TimeoutSec 1800

                $result.Items          = @($tr.Items)
                $result.SkippedTitles  = @($tr.SkippedTitles)
                $result.NotFoundIDs    = @($tr.NotFoundIDs)
                $result.RebootRequired = [bool]$tr.RebootRequired
                $result.ErrorKind      = $tr.ErrorKind
                $result.ErrorMessage   = $tr.ErrorMessage
            } catch {
                $result.ErrorKind    = "INSTALL_FAILED"
                $result.ErrorMessage = $_.ToString()
            }
            return $result
        }

        # LOCAL install path (existing logic)
        try {
            $session  = New-Object -ComObject Microsoft.Update.Session
            $searcher = $session.CreateUpdateSearcher()
            $searcher.Online = $true
            $search = $searcher.Search("IsInstalled=0")

            $toInstall = New-Object -ComObject Microsoft.Update.UpdateColl
            $found = @{}
            foreach ($u in $search.Updates) {
                $uid = $u.Identity.UpdateID
                if ($UpdateIDs -contains $uid) {
                    $found[$uid] = $true
                    if ($u.InstallationBehavior.CanRequestUserInput) {
                        $result.SkippedTitles += ($u.Title + " [requires user input]")
                        continue
                    }
                    if (-not $u.EulaAccepted) {
                        try { $u.AcceptEula() } catch {
                            $result.SkippedTitles += ($u.Title + " [EULA could not be accepted]")
                            continue
                        }
                    }
                    [void]$toInstall.Add($u)
                }
            }
            foreach ($id in $UpdateIDs) {
                if (-not $found.ContainsKey($id)) { $result.NotFoundIDs += $id }
            }

            if ($toInstall.Count -eq 0) {
                $result.ErrorKind    = "NOTHING_TO_INSTALL"
                $result.ErrorMessage = "None of the selected updates are currently installable. Rescan and try again."
                return $result
            }

            $downloader = $session.CreateUpdateDownloader()
            $downloader.Updates = $toInstall
            $dl = $downloader.Download()
            if ($dl.ResultCode -ne 2 -and $dl.ResultCode -ne 3) {
                $result.ErrorKind    = "DOWNLOAD_FAILED"
                $result.ErrorMessage = ("Download failed - result code {0}, HResult 0x{1:X8}" -f $dl.ResultCode, $dl.HResult)
                return $result
            }

            $installer = $session.CreateUpdateInstaller()
            $installer.Updates = $toInstall
            $inst = $installer.Install()
            $result.RebootRequired = [bool]$inst.RebootRequired

            for ($i = 0; $i -lt $toInstall.Count; $i++) {
                $ur = $inst.GetUpdateResult($i)
                if     ($ur.ResultCode -eq 2) { $outcome = "Succeeded" }
                elseif ($ur.ResultCode -eq 3) { $outcome = "Succeeded with errors" }
                elseif ($ur.ResultCode -eq 4) { $outcome = "Failed" }
                elseif ($ur.ResultCode -eq 5) { $outcome = "Aborted" }
                else                          { $outcome = "Unknown ($($ur.ResultCode))" }
                $item = $toInstall.Item($i)
                $kb = if ($item.KBArticleIDs.Count -gt 0) { "KB" + $item.KBArticleIDs[0] } else { "N/A" }
                $result.Items += [PSCustomObject]@{
                    KB      = $kb
                    Title   = $item.Title
                    Outcome = $outcome
                    HResult = ("0x{0:X8}" -f $ur.HResult)
                }
            }
        } catch {
            $result.ErrorKind    = "INSTALL_FAILED"
            $result.ErrorMessage = $_.ToString()
        }
        return $result
    ')

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.Open()

    $ps          = [powershell]::Create()
    $ps.Runspace = $rs

    [void]$ps.AddScript($bgScript)
    [void]$ps.AddArgument($UpdateIDs)
    [void]$ps.AddArgument($IsRemote)
    [void]$ps.AddArgument($Target)
    [void]$ps.AddArgument($Credential)
    [void]$ps.AddArgument($UseSSL)
    [void]$ps.AddArgument($ScanWorkerSrc)
    [void]$ps.AddArgument($InstallWorkerSrc)

    $script:ActiveOp   = 'Install'
    $script:RunningPS  = $ps
    $script:RunningRS  = $rs
    $script:PollHandle = $ps.BeginInvoke()

    $timer          = New-Object System.Windows.Forms.Timer
    $timer.Interval = 500
    $timer.Add_Tick($installPollTick)
    $script:PollTimer = $timer
    $timer.Start()
}

# ---------------------------------------------------------------------------
# Reboot poll-tick handler (UI-thread; polls background reboot sub-ops)
# ---------------------------------------------------------------------------
$rebootPollTick = {
    # Elapsed status update for Wait mode -- do this BEFORE the early-return so
    # the user sees progress even though the bg op has not completed yet.
    if ($script:RebootMode -eq 'Wait' -and $null -ne $script:RebootWaitStart) {
        $elapsedSec = [int]([datetime]::Now - $script:RebootWaitStart).TotalSeconds
        $lblStatus.Text = "Waiting for $($script:LastTarget) to come back online... ($($elapsedSec)s elapsed)"
    }

    if (-not $script:PollHandle.IsCompleted) { return }

    $script:PollTimer.Stop()

    $rebootResult = $null
    $psErrors     = $null
    $hadErrors    = $false
    $capturedMode = $script:RebootMode   # capture before cleanup clears it
    try {
        $rebootResult = $script:RunningPS.EndInvoke($script:PollHandle)
        $psErrors     = $script:RunningPS.Streams.Error
        $hadErrors    = $script:RunningPS.HadErrors
    } catch {
        $hadErrors = $true
        $psErrors  = @($_)
    } finally {
        Stop-AsyncCleanup   # disposes PS/RS/PollTimer; does NOT touch RebootTimer/RebootPhase
    }

    $res = if ($rebootResult -is [System.Collections.IList]) { $rebootResult[0] } else { $rebootResult }

    if ($capturedMode -eq 'Issue') {
        # ---- Issue done ----
        if ($hadErrors -and ($null -eq $res -or $res.Ok -ne $true)) {
            $errTxt = if ($psErrors -and $psErrors.Count -gt 0) {
                $psErrors[0].ToString()
            } else {
                'Unknown error launching remote reboot.'
            }
            Show-Error "Reboot error: $errTxt"
            $lblStatus.Text = 'Remote reboot failed.'
            # Drop cached cred if auth-related
            if ($errTxt -match 'Access is denied|logon|authentication|user name or password|password is incorrect|credential') {
                $script:LastCredential = $null
                $script:LastCredTarget = ''
            }
            Reset-RebootState
            return
        }

        if ($null -eq $res -or $res.ErrorKind) {
            $ek = if ($res) { $res.ErrorKind } else { 'null result' }
            $em = if ($res) { $res.ErrorMessage } else { 'No result from reboot operation.' }
            if ($ek -eq 'UNREACHABLE') {
                Show-Error "Cannot reach $($script:LastTarget) to schedule reboot - WinRM unreachable."
            } else {
                Show-Error "Reboot failed ($ek): $em"
            }
            $lblStatus.Text = 'Remote reboot failed.'
            Reset-RebootState
            return
        }

        # Success -- enter COUNTDOWN phase
        $script:RebootPhase     = 'Countdown'
        $script:RebootCountdown = $script:RebootDelaySec
        $script:ActiveOp        = 'Reboot'   # re-assert after Stop-AsyncCleanup
        $btnReboot.Text         = 'Abort Reboot'
        $btnReboot.Enabled      = $true
        $btnScan.Enabled        = $false
        $btnInstall.Enabled     = $false
        $lblStatus.Text         = "Rebooting $($script:LastTarget) in $($script:RebootCountdown)s - click Abort to cancel"

        $cTimer          = New-Object System.Windows.Forms.Timer
        $cTimer.Interval = 1000
        $cTimer.Add_Tick($rebootCountdownTick)
        $script:RebootTimer = $cTimer
        $cTimer.Start()

    } elseif ($capturedMode -eq 'Abort') {
        # ---- Abort done ----
        $lblStatus.Text = 'Reboot aborted.'
        Reset-RebootState

    } elseif ($capturedMode -eq 'Wait') {
        # ---- Wait done ----
        if ($null -ne $res -and $res.Back -eq $true) {
            $lblStatus.Text = "$($script:LastTarget) back online - re-scanning..."
            Reset-RebootState
            $btnScan.PerformClick()
        } elseif ($null -ne $res -and $res.TimedOut -eq $true) {
            $lblStatus.Text = "Timed out waiting for $($script:LastTarget); re-scan manually when it is back."
            Reset-RebootState
        } else {
            $errTxt = if ($null -ne $res -and $res.ErrorMessage) { $res.ErrorMessage } else { 'Unknown wait error.' }
            Show-Error "Wait error: $errTxt"
            $lblStatus.Text = "Error waiting for $($script:LastTarget); re-scan manually."
            Reset-RebootState
        }
    }
}

# ---------------------------------------------------------------------------
# Reboot countdown tick handler (1-second timer; decrements countdown)
# ---------------------------------------------------------------------------
$rebootCountdownTick = {
    $script:RebootCountdown--
    if ($script:RebootCountdown -gt 0) {
        $lblStatus.Text = "Rebooting $($script:LastTarget) in $($script:RebootCountdown)s - click Abort to cancel"
        return
    }

    # Countdown reached zero -- server is now rebooting
    $script:RebootTimer.Stop()
    $script:RebootTimer.Dispose()
    $script:RebootTimer = $null

    $script:RebootPhase    = 'Waiting'
    $script:RebootWaitStart = [datetime]::Now
    $btnReboot.Text        = 'Stop Waiting'
    $btnReboot.Enabled     = $true
    $lblStatus.Text        = "Waiting for $($script:LastTarget) to come back online... (0s elapsed)"

    Start-RebootOpAsync -Mode 'Wait' `
        -Target     $script:LastTarget `
        -Credential $script:LastCredential `
        -UseSSL     $script:LastUseSSL `
        -TimeoutSec $script:RebootWaitTimeoutSec
}

function Start-RebootOpAsync {
    param(
        [string]  $Mode,       # 'Issue' | 'Abort' | 'Wait'
        [string]  $Target,
        [System.Management.Automation.PSCredential]$Credential,
        [bool]    $UseSSL,
        [int]     $DelaySec  = 60,
        [string]  $Message   = '',
        [int]     $TimeoutSec = 1800
    )

    # ------------------------------------------------------------------
    # Background scriptblocks -- single-quoted [scriptblock]::Create('...')
    # NO apostrophes anywhere inside; all inner strings use double-quotes.
    # ------------------------------------------------------------------

    if ($Mode -eq 'Issue') {
        # Build shutdown command string (no apostrophes, no stray chars)
        $shutdownMsg = if ($Message -ne '') { $Message } else {
            "WinUpdateChecker: System reboot scheduled by administrator. Save your work now."
        }
        # We pass DelaySec and shutdownMsg as arguments to keep the bgScript clean.
        $bgScript = [scriptblock]::Create('
            param($Target, $Credential, $UseSSL, $DelaySec, $ShutdownMsg)

            $result = [PSCustomObject]@{ Ok=$false; ErrorKind=$null; ErrorMessage=$null }

            $icParams = @{ ComputerName=$Target; ErrorAction="Stop" }
            if ($Credential) { $icParams.Credential = $Credential }
            if ($UseSSL)     { $icParams.UseSSL     = $true }

            # Optional reachability probe
            $port5985 = $false
            $port5986 = $false
            try {
                $t = Test-NetConnection -ComputerName $Target -Port 5985 -WarningAction SilentlyContinue -ErrorAction Stop
                $port5985 = $t.TcpTestSucceeded
            } catch { $port5985 = $false }
            if (-not $port5985) {
                try {
                    $t = Test-NetConnection -ComputerName $Target -Port 5986 -WarningAction SilentlyContinue -ErrorAction Stop
                    $port5986 = $t.TcpTestSucceeded
                } catch { $port5986 = $false }
            }
            if (-not $port5985 -and -not $port5986) {
                $result.ErrorKind    = "UNREACHABLE"
                $result.ErrorMessage = "WinRM not reachable on $Target (ports 5985/5986)"
                return $result
            }

            try {
                $shutdownCmd = "shutdown.exe /r /f /t " + $DelaySec + " /c """ + $ShutdownMsg + """"
                $icParams.ScriptBlock = [scriptblock]::Create($shutdownCmd)
                Invoke-Command @icParams | Out-Null
                $result.Ok = $true
            } catch {
                $errStr = $_.ToString()
                # A connection-terminated exception after /t 60 is effectively success
                # (shutdown.exe returned 0 and the session dropped) -- classify as ok.
                if ($errStr -match "pipeline has been stopped|The pipeline was stopped|connection.*reset|forcibly closed|network path was not found") {
                    $result.Ok = $true
                } else {
                    $result.ErrorKind    = "REBOOT_FAILED"
                    $result.ErrorMessage = $errStr
                }
            }
            return $result
        ')

        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.Open()
        $ps          = [powershell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript($bgScript)
        [void]$ps.AddArgument($Target)
        [void]$ps.AddArgument($Credential)
        [void]$ps.AddArgument($UseSSL)
        [void]$ps.AddArgument($DelaySec)
        [void]$ps.AddArgument($shutdownMsg)

        $script:ActiveOp   = 'Reboot'
        $script:RebootMode = 'Issue'
        $script:RunningPS  = $ps
        $script:RunningRS  = $rs
        $script:PollHandle = $ps.BeginInvoke()

        $timer          = New-Object System.Windows.Forms.Timer
        $timer.Interval = 500
        $timer.Add_Tick($rebootPollTick)
        $script:PollTimer = $timer
        $timer.Start()

    } elseif ($Mode -eq 'Abort') {
        $bgScript = [scriptblock]::Create('
            param($Target, $Credential, $UseSSL)

            $result = [PSCustomObject]@{ Ok=$false; ErrorKind=$null; ErrorMessage=$null }

            $icParams = @{ ComputerName=$Target; ErrorAction="Stop" }
            if ($Credential) { $icParams.Credential = $Credential }
            if ($UseSSL)     { $icParams.UseSSL     = $true }

            try {
                $icParams.ScriptBlock = [scriptblock]::Create("shutdown.exe /a")
                Invoke-Command @icParams | Out-Null
                $result.Ok = $true
            } catch {
                # "no shutdown was in progress" is exit code 1116; treat as ok-enough
                $result.Ok = $true
            }
            return $result
        ')

        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.Open()
        $ps          = [powershell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript($bgScript)
        [void]$ps.AddArgument($Target)
        [void]$ps.AddArgument($Credential)
        [void]$ps.AddArgument($UseSSL)

        $script:ActiveOp   = 'Reboot'
        $script:RebootMode = 'Abort'
        $script:RunningPS  = $ps
        $script:RunningRS  = $rs
        $script:PollHandle = $ps.BeginInvoke()

        $timer          = New-Object System.Windows.Forms.Timer
        $timer.Interval = 500
        $timer.Add_Tick($rebootPollTick)
        $script:PollTimer = $timer
        $timer.Start()

    } elseif ($Mode -eq 'Wait') {
        $bgScript = [scriptblock]::Create('
            param($Target, $Credential, $UseSSL, $TimeoutSec)

            $result = [PSCustomObject]@{ Back=$false; TimedOut=$false; ErrorKind=$null; ErrorMessage=$null }

            $deadline = (Get-Date).AddSeconds($TimeoutSec)

            # Phase 1: wait until target stops responding (confirm it actually went down).
            # We poll on a 5-second interval for up to the full deadline.
            # We MUST observe the target offline before we can report it came back.
            # (Without this gate, a still-up target would be mis-reported as "Back" once
            # Phase 2 probes it and finds it reachable.)
            $wentDown = $false
            while ((Get-Date) -lt $deadline) {
                Start-Sleep -Seconds 5
                $up = $false
                try {
                    $t = Test-NetConnection -ComputerName $Target -Port 5985 -WarningAction SilentlyContinue -ErrorAction Stop
                    $up = $t.TcpTestSucceeded
                } catch { $up = $false }
                if (-not $up) {
                    try {
                        $t = Test-NetConnection -ComputerName $Target -Port 5986 -WarningAction SilentlyContinue -ErrorAction Stop
                        $up = $t.TcpTestSucceeded
                    } catch { $up = $false }
                }
                if (-not $up) { $wentDown = $true; break }
            }

            # If the target never went offline before the deadline, it never actually rebooted.
            if (-not $wentDown) {
                $result.TimedOut = $true
                return $result
            }

            # Phase 2: target confirmed offline; wait until it responds again
            while ((Get-Date) -lt $deadline) {
                Start-Sleep -Seconds 8
                $up = $false
                try {
                    $t = Test-NetConnection -ComputerName $Target -Port 5985 -WarningAction SilentlyContinue -ErrorAction Stop
                    $up = $t.TcpTestSucceeded
                } catch { $up = $false }
                if (-not $up) {
                    try {
                        $t = Test-NetConnection -ComputerName $Target -Port 5986 -WarningAction SilentlyContinue -ErrorAction Stop
                        $up = $t.TcpTestSucceeded
                    } catch { $up = $false }
                }
                if ($up) { $result.Back = $true; return $result }
            }

            $result.TimedOut = $true
            return $result
        ')

        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.Open()
        $ps          = [powershell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript($bgScript)
        [void]$ps.AddArgument($Target)
        [void]$ps.AddArgument($Credential)
        [void]$ps.AddArgument($UseSSL)
        [void]$ps.AddArgument($TimeoutSec)

        $script:ActiveOp   = 'Reboot'
        $script:RebootMode = 'Wait'
        $script:RunningPS  = $ps
        $script:RunningRS  = $rs
        $script:PollHandle = $ps.BeginInvoke()

        $timer          = New-Object System.Windows.Forms.Timer
        $timer.Interval = 2000   # poll every 2s for elapsed-status updates
        $timer.Add_Tick($rebootPollTick)
        $script:PollTimer = $timer
        $timer.Start()
    }
}

#endregion

#region -- Event Handlers -----------------------------------------------------

# Fix 4: disable Auth radio buttons for Local Machine; re-enable for Remote
$rbLocal.Add_CheckedChanged({
    if ($rbLocal.Checked) {
        $txtHost.Enabled       = $false
        $rbCurrentUser.Enabled = $false
        $rbCreds.Enabled       = $false
        # NOTE: do NOT force $rbCurrentUser.Checked here -- with independent sub-panels
        # WinForms handles mutual exclusion within each group correctly; forcing
        # the auth radio from the target handler was the root cause of the old bug.
    }
})

$rbRemote.Add_CheckedChanged({
    if ($rbRemote.Checked) {
        $txtHost.Enabled       = $true
        $rbCurrentUser.Enabled = $true
        $rbCreds.Enabled       = $true
    }
})

# Plan item 9: store effective transport in $script:Last* at scan launch
$btnScan.Add_Click({
    Hide-Error
    $grid.Rows.Clear()
    $script:ScanResults   = @()
    $script:RebootPending = $false
    $script:CanInstall    = $false
    $script:CanReboot     = $false
    Update-RebootButton
    $btnInstall.Enabled    = $false
    $btnExportCsv.Enabled  = $false
    $btnExportHtml.Enabled = $false
    $lblInstalled.Text    = 'Installed: -'
    $lblMissing.Text      = 'Missing: -'
    $lblPending.Text      = 'Reboot Pending: -'
    $lblPending.ForeColor = [System.Drawing.Color]::FromArgb(230, 81, 0)
    $lblScanTime.Text     = ''
    $lblPolicy.Text       = ''
    $lblPolicy.ForeColor  = [System.Drawing.Color]::Gray

    $isRemote = $rbRemote.Checked
    $target   = $txtHost.Text.Trim()

    if ($isRemote -and [string]::IsNullOrWhiteSpace($target)) {
        Show-Error 'Please enter a hostname or IP for remote scan.'
        return
    }

    $cred = $null
    if ($rbCreds.Checked -and $rbCreds.Enabled) {
        $credTarget = if ($isRemote) { $target } else { 'local scan' }
        if ($script:LastCredential -and $script:LastCredTarget -eq $credTarget) {
            # Reuse the credential already entered for this target this session
            # (so install + the post-install auto-rescan don't re-prompt).
            $cred = $script:LastCredential
        } else {
            $cred = Get-Credential -Message "Enter credentials for $credTarget"
            if (-not $cred) { return }
            $script:LastCredTarget = $credTarget
        }
    }

    # Plan item 9: capture effective transport for later install
    $script:LastWasRemote  = $isRemote
    $script:LastTarget     = $target
    $script:LastCredential = $cred
    $script:LastUseSSL     = $false   # UseSSL auto-detected inside bgScript; default false here

    # Auto-manage TrustedHosts: ensure the remote target is trusted before connecting.
    if ($isRemote) {
        Add-TrustedHostForTarget -Target $target
    }

    $btnScan.Enabled     = $false
    $pnlProgress.Visible = $true
    if ($isRemote) {
        $lblStatus.Text = "Scanning $target..."
    } else {
        $lblStatus.Text = 'Scanning...'
    }

    Start-ScanAsync -Target $target -IsRemote $isRemote -Credential $cred -UseSSL $false `
        -ScanWorkerSrc $script:RemoteScanWorkerSrc `
        -InstallWorkerSrc $script:RemoteInstallWorkerSrc
})

# Fix 10: color ENTIRE row background by status (not just the Status cell).
# Uses if/elseif rather than switch: inside a switch block PowerShell rebinds
# $_ to the switch input value, clobbering the CellFormatting event args.
# Non-installable Select cells get a gray background instead; this must live
# here (not as cell.Style in Update-Grid) or the repaint overwrites it.
$grid.Add_CellFormatting({
    if ($_.RowIndex -ge 0) {
        if ($grid.Columns[$_.ColumnIndex].Name -eq 'Select' -and
            $grid.Rows[$_.RowIndex].Cells['Select'].ReadOnly) {
            $_.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(225, 225, 225)
        } else {
            $statusVal = $grid.Rows[$_.RowIndex].Cells['Status'].Value
            if ($statusVal -eq 'Installed') {
                $_.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 255, 230)
            } elseif ($statusVal -eq 'Missing') {
                $_.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 230, 230)
            }
        }
    }
})

# Commit checkbox toggles immediately (otherwise the value lands only when
# the cell loses focus) so the Install button count stays live
$grid.Add_CurrentCellDirtyStateChanged({
    if ($grid.IsCurrentCellDirty -and
        $grid.CurrentCell -is [System.Windows.Forms.DataGridViewCheckBoxCell]) {
        $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})

$grid.Add_CellValueChanged({
    if ($_.RowIndex -ge 0 -and $grid.Columns[$_.ColumnIndex].Name -eq 'Select') {
        Update-InstallButton
    }
})

# Fix 9: KB URL construction uses anchored replace '^KB' so only the leading
# "KB" prefix is stripped (e.g. "KB1234" -> "1234"), not all occurrences.
$grid.Add_CellDoubleClick({
    if ($_.RowIndex -ge 0 -and $grid.Columns[$_.ColumnIndex].Name -ne 'Select') {
        $rowKb     = $grid.Rows[$_.RowIndex].Cells['KB'].Value
        $rowTitle  = $grid.Rows[$_.RowIndex].Cells['Title'].Value
        $rowStatus = $grid.Rows[$_.RowIndex].Cells['Status'].Value
        $rowSev    = $grid.Rows[$_.RowIndex].Cells['Severity'].Value
        $rowDate   = $grid.Rows[$_.RowIndex].Cells['Date'].Value
        $rowSize   = $grid.Rows[$_.RowIndex].Cells['SizeKB'].Value

        $detail                 = New-Object System.Windows.Forms.Form
        $detail.Text            = "Update Detail - $rowKb"
        $detail.Size            = New-Object System.Drawing.Size(500, 280)
        $detail.StartPosition   = 'CenterParent'
        $detail.FormBorderStyle = 'FixedDialog'
        $detail.MaximizeBox     = $false

        $txt          = New-Object System.Windows.Forms.RichTextBox
        $txt.Dock     = 'Fill'
        $txt.ReadOnly = $true
        $txt.Font     = New-Object System.Drawing.Font('Segoe UI', 9)

        # Fix 9: anchored replace strips only the leading "KB" prefix
        $kbNum    = $rowKb -replace '^KB', ''
        $txt.Text = "KB Article : $rowKb`r`nTitle      : $rowTitle`r`nStatus     : $rowStatus`r`nSeverity   : $rowSev`r`nDate       : $rowDate`r`nSize       : $rowSize KB`r`n`r`nMicrosoft Support URL:`r`nhttps://support.microsoft.com/kb/$kbNum"

        $btnClose = New-Object System.Windows.Forms.Button
        $btnClose.Text = 'Close'
        $btnClose.Dock = 'Bottom'
        $btnClose.Add_Click({ $detail.Close() })

        $detail.Controls.AddRange(@($txt, $btnClose))
        $detail.ShowDialog($form) | Out-Null
    }
})

# Filters -- Fix 8: Update-Grid delegates to Get-FilteredResults
$cmbStatus.Add_SelectedIndexChanged({ Update-Grid })
$cmbSeverity.Add_SelectedIndexChanged({ Update-Grid })

# Export CSV -- Fix 8: export the currently filtered view
$btnExportCsv.Add_Click({
    $dlg            = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter     = 'CSV files (*.csv)|*.csv'
    $dlg.FileName   = "WindowsUpdates_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($dlg.ShowDialog() -eq 'OK') {
        $toExport = Get-FilteredResults
        Export-ToCsv -Records $toExport -Path $dlg.FileName
        $lblStatus.Text = "Exported CSV to $($dlg.FileName)"
    }
})

# Export HTML -- Fix 8: export filtered view; Fix 11: use $script:ScannedOS
$btnExportHtml.Add_Click({
    $dlg            = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter     = 'HTML files (*.html)|*.html'
    $dlg.FileName   = "WindowsUpdates_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    if ($dlg.ShowDialog() -eq 'OK') {
        $targetName = if ($rbRemote.Checked) { $txtHost.Text } else { $env:COMPUTERNAME }
        $toExport   = Get-FilteredResults
        Export-ToHtml -Records $toExport -Path $dlg.FileName `
            -TargetName $targetName -OSVersion $script:ScannedOS
        $lblStatus.Text = "Exported HTML to $($dlg.FileName)"
        Start-Process $dlg.FileName
    }
})

# Plan item 10: pass remote context to Start-InstallAsync using stored $script:Last* values
# Plan item 10: name the target in the confirm dialog when remote
$btnInstall.Add_Click({
    $ids = Get-CheckedUpdateIDs
    if ($ids.Count -eq 0) {
        Show-Error 'No installable updates are checked.'
        return
    }

    $confirmMsg = if ($script:LastWasRemote) {
        "Download and install $($ids.Count) update(s) on $($script:LastTarget)?"
    } else {
        "Download and install $($ids.Count) update(s) on this machine?"
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        $confirmMsg,
        'Confirm Install',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    Hide-Error
    $btnScan.Enabled     = $false
    $btnInstall.Enabled  = $false
    $pnlProgress.Visible = $true

    $targetDesc = if ($script:LastWasRemote) { $script:LastTarget } else { 'this machine' }
    $lblStatus.Text = "Downloading and installing $($ids.Count) update(s) on $targetDesc... this can take several minutes."

    Start-InstallAsync -UpdateIDs $ids `
        -IsRemote    $script:LastWasRemote `
        -Target      $script:LastTarget `
        -Credential  $script:LastCredential `
        -UseSSL      $script:LastUseSSL `
        -ScanWorkerSrc    $script:RemoteScanWorkerSrc `
        -InstallWorkerSrc $script:RemoteInstallWorkerSrc
})

# Reboot Remote button -- three-phase click handler (if/elseif; NEVER switch in handler)
$btnReboot.Add_Click({
    if ($script:RebootPhase -eq 'Idle') {
        # Guard: must be remote
        if (-not $script:LastWasRemote) { return }

        $confirmMsg = (
            "You are about to schedule a FORCED REBOOT of $($script:LastTarget)." +
            "`r`n`r`n" +
            "- The reboot will occur in $($script:RebootDelaySec) seconds." + "`r`n" +
            "- All applications on the remote server will be force-closed." + "`r`n" +
            "- Any unsaved work on that server may be lost." + "`r`n" +
            "- This tool will wait for $($script:LastTarget) to come back online and then re-scan automatically." +
            "`r`n`r`n" +
            "Proceed?"
        )
        $answer = [System.Windows.Forms.MessageBox]::Show(
            $confirmMsg,
            "Confirm Remote Reboot - $($script:LastTarget)",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        Hide-Error
        $btnScan.Enabled    = $false
        $btnInstall.Enabled = $false
        $btnReboot.Enabled  = $false
        $lblStatus.Text     = "Scheduling reboot on $($script:LastTarget)..."

        $rebootWarning = "WinUpdateChecker: System reboot scheduled by administrator. Save your work now. Rebooting in $($script:RebootDelaySec) seconds."

        Start-RebootOpAsync -Mode 'Issue' `
            -Target     $script:LastTarget `
            -Credential $script:LastCredential `
            -UseSSL     $script:LastUseSSL `
            -DelaySec   $script:RebootDelaySec `
            -Message    $rebootWarning

    } elseif ($script:RebootPhase -eq 'Countdown') {
        # Abort the scheduled reboot
        if ($null -ne $script:RebootTimer) {
            $script:RebootTimer.Stop()
            $script:RebootTimer.Dispose()
            $script:RebootTimer = $null
        }
        $btnReboot.Enabled = $false
        $lblStatus.Text    = "Aborting reboot on $($script:LastTarget)..."

        Start-RebootOpAsync -Mode 'Abort' `
            -Target     $script:LastTarget `
            -Credential $script:LastCredential `
            -UseSSL     $script:LastUseSSL

    } elseif ($script:RebootPhase -eq 'Waiting') {
        # Stop waiting -- the reboot is committed; user will re-scan manually
        Stop-AsyncCleanup   # clears ActiveOp, disposes bg runspace/PS/PollTimer
        $lblStatus.Text = "Stopped waiting - re-scan manually when $($script:LastTarget) is back."
        Reset-RebootState   # stops RebootTimer (none in Wait), phase=Idle, re-enables controls
    }
})

# Clear
$btnClear.Add_Click({
    $grid.Rows.Clear()
    $script:ScanResults   = @()
    $script:RebootPending = $false
    $script:ScannedOS     = ''
    $script:CanInstall    = $false
    $script:CanReboot     = $false
    # Forget any cached credential so the next scan prompts fresh
    $script:LastCredential = $null
    $script:LastCredTarget = ''
    Update-InstallButton
    Update-RebootButton
    $btnExportCsv.Enabled  = $false
    $btnExportHtml.Enabled = $false
    $lblInstalled.Text    = 'Installed: -'
    $lblMissing.Text      = 'Missing: -'
    $lblPending.Text      = 'Reboot Pending: -'
    $lblPending.ForeColor = [System.Drawing.Color]::FromArgb(230, 81, 0)
    $lblScanTime.Text     = ''
    $lblPolicy.Text       = ''
    $lblPolicy.ForeColor  = [System.Drawing.Color]::Gray
    $lblStatus.Text       = 'Ready'
    Hide-Error
})

# Plan item 11: FormClosing warning includes note about remote task continuing.
# Fix 12: clean up timer/runspace if the form closes during an active operation.
# Closing mid-install gets a warning first: a synchronous WUA Install() cannot
# be aborted promptly and killing it can leave the update agent busy.
# Note: if a remote SYSTEM task is running, it will continue on the target;
# orphan-cleanup-on-next-connect is the safety net for that scenario.
$form.Add_FormClosing({
    # Reboot Countdown: reboot is already scheduled; closing won't cancel it
    if ($script:RebootPhase -eq 'Countdown') {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            ("A reboot of $($script:LastTarget) has been scheduled and will occur in approximately " +
             "$($script:RebootCountdown) seconds. CLOSING THIS APP WILL NOT CANCEL IT." +
             "`r`n`r`nTo cancel the reboot, click 'No' and then click 'Abort Reboot'." +
             "`r`n`r`nClose anyway (the scheduled reboot on $($script:LastTarget) will proceed)?"),
            'Reboot Scheduled - Closing Will Not Cancel',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            $_.Cancel = $true
            return
        }
    }
    # Reboot Waiting: safe to close; inform and proceed
    if ($script:RebootPhase -eq 'Waiting') {
        [System.Windows.Forms.MessageBox]::Show(
            ("$($script:LastTarget) is rebooting. Closing this app will stop the " +
             "online check but will not affect the reboot." +
             "`r`n`r`nRe-scan manually once the server is back online."),
            'Waiting for Server - Closing',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }
    if ($script:ActiveOp -eq 'Install') {
        $remoteNote = if ($script:LastWasRemote) {
            " Note: the remote install task on $($script:LastTarget) will continue running until it finishes."
        } else {
            ''
        }
        $answer = [System.Windows.Forms.MessageBox]::Show(
            ('An update installation is in progress. Closing now will not roll it back' +
             ' and may leave Windows Update busy until it finishes.' +
             $remoteNote +
             "`n`nClose anyway?"),
            'Installation In Progress',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            $_.Cancel = $true
            return
        }
    }
    # Clean up reboot timer if still running
    if ($null -ne $script:RebootTimer) {
        $script:RebootTimer.Stop()
        $script:RebootTimer.Dispose()
        $script:RebootTimer = $null
    }
    # Restore TrustedHosts to baseline (removes session-added entries, keeps pre-existing).
    Restore-TrustedHosts
    Stop-AsyncCleanup
})

#endregion

#region -- Entry Point --------------------------------------------------------
# Fix 4: Auth controls start disabled because Local Machine is the default
$rbCurrentUser.Enabled = $false
$rbCreds.Enabled       = $false

# Initialize TrustedHosts auto-management (crash sweep + baseline capture).
# Must run before ShowDialog so any stale marker from a prior crash is healed
# before the user can trigger a remote scan.
Initialize-TrustedHostsManagement

[void]$form.ShowDialog()
#endregion
