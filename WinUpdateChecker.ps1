#Requires -Version 5.1
<#
.SYNOPSIS
    Windows Update Verification Tool - checks installed vs missing updates locally or on a remote server.
.DESCRIPTION
    WinForms GUI that queries the Windows Update API (WUA COM) for update status,
    displays results color-coded in a DataGridView, and exports to CSV or HTML.
.NOTES
    Run as Administrator for best results. Remote scanning requires WinRM on the target.
#>

#region -- Imports ------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
#endregion

#region -- Shared Script State ------------------------------------------------
$script:ScanResults   = @()
$script:IsPartialData = $false
$script:RebootPending = $false
$script:ScannedOS     = ''     # Fix 11: captured from target during scan
$script:RunningPS     = $null  # Fix 1 / Fix 12: active PowerShell instance
$script:RunningRS     = $null  # Fix 1 / Fix 12: active Runspace instance
$script:PollHandle    = $null  # Fix 1: IAsyncResult from BeginInvoke
$script:PollTimer     = $null  # Fix 1 / Fix 12: System.Windows.Forms.Timer
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
        $color = switch ($_.Status) {
            'Installed' { '#e6ffe6' }
            'Missing'   { '#ffe6e6' }
            default     { '#ffffff' }
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
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Windows Update Checker'
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

$rbLocal          = New-Object System.Windows.Forms.RadioButton
$rbLocal.Text     = 'Local Machine'
$rbLocal.Location = New-Object System.Drawing.Point(65, 10)
$rbLocal.AutoSize = $true
$rbLocal.Checked  = $true

$rbRemote          = New-Object System.Windows.Forms.RadioButton
$rbRemote.Text     = 'Remote:'
$rbRemote.Location = New-Object System.Drawing.Point(185, 10)
$rbRemote.AutoSize = $true

$txtHost          = New-Object System.Windows.Forms.TextBox
$txtHost.Location = New-Object System.Drawing.Point(255, 8)
$txtHost.Width    = 200
$txtHost.Enabled  = $false
# Fix 2: PlaceholderText only exists on .NET Core / .NET 5+ WinForms, not on
# .NET Framework 4.x (used by Windows PowerShell 5.1). Guard the assignment.
if ($txtHost.PSObject.Properties['PlaceholderText']) {
    $txtHost.PlaceholderText = 'hostname or IP'
}

$lblAuth          = New-Object System.Windows.Forms.Label
$lblAuth.Text     = 'Auth:'
$lblAuth.Location = New-Object System.Drawing.Point(10, 44)
$lblAuth.AutoSize = $true
$lblAuth.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

$rbCurrentUser           = New-Object System.Windows.Forms.RadioButton
$rbCurrentUser.Text      = 'Current User'
$rbCurrentUser.Location  = New-Object System.Drawing.Point(65, 42)
$rbCurrentUser.AutoSize  = $true
$rbCurrentUser.Checked   = $true

$rbCreds          = New-Object System.Windows.Forms.RadioButton
$rbCreds.Text     = 'Specify Credentials'
$rbCreds.Location = New-Object System.Drawing.Point(185, 42)
$rbCreds.AutoSize = $true

$btnScan            = New-Object System.Windows.Forms.Button
$btnScan.Text       = 'Scan Now'
$btnScan.Location   = New-Object System.Drawing.Point(750, 20)
$btnScan.Size       = New-Object System.Drawing.Size(100, 40)
$btnScan.BackColor  = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnScan.ForeColor  = [System.Drawing.Color]::White
$btnScan.FlatStyle  = 'Flat'
$btnScan.Font       = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$btnScan.Anchor     = 'Top,Right'

$pnlTarget.Controls.AddRange(@($lblTarget, $rbLocal, $rbRemote, $txtHost,
    $lblAuth, $rbCurrentUser, $rbCreds, $btnScan))

# -- Summary Panel --
$pnlSummary           = New-Object System.Windows.Forms.Panel
$pnlSummary.Dock      = 'Top'
$pnlSummary.Height    = 60
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

$pnlSummary.Controls.AddRange(@($lblInstalled, $lblMissing, $lblPending, $lblScanTime))

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
$pnlError.Padding   = New-Object System.Windows.Forms.Padding(10, 5, 10, 0)

$lblError           = New-Object System.Windows.Forms.Label
$lblError.Dock      = 'Fill'
$lblError.ForeColor = [System.Drawing.Color]::FromArgb(198, 40, 40)
$lblError.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
$pnlError.Controls.Add($lblError)

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
# Fix 6: removed 'Pending Reboot' from filter — that status no longer appears on rows
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
$grid.ReadOnly                = $true
$grid.AllowUserToAddRows      = $false
$grid.AllowUserToDeleteRows   = $false
$grid.RowHeadersVisible       = $false
$grid.AutoSizeColumnsMode     = 'Fill'
$grid.SelectionMode           = 'FullRowSelect'
$grid.BackgroundColor         = [System.Drawing.Color]::White
$grid.BorderStyle             = 'None'
$grid.ColumnHeadersHeightSizeMode = 'AutoSize'
$grid.Font                    = New-Object System.Drawing.Font('Segoe UI', 9)

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
    if ($c.Fill) {
        $col.AutoSizeMode = 'Fill'
    } else {
        $col.Width        = $c.Width
        $col.AutoSizeMode = 'None'
    }
    $grid.Columns.Add($col) | Out-Null
}

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

$lblStatus           = New-Object System.Windows.Forms.Label
$lblStatus.Text      = 'Ready'
$lblStatus.Location  = New-Object System.Drawing.Point(320, 10)
$lblStatus.AutoSize  = $true
$lblStatus.ForeColor = [System.Drawing.Color]::Gray

$pnlBottom.Controls.AddRange(@($btnExportCsv, $btnExportHtml, $btnClear, $lblStatus))

# -- Assemble Form --
$form.Controls.AddRange(@($grid, $pnlBottom, $pnlFilter, $pnlError,
    $pnlProgress, $pnlSummary, $pnlTarget))

#endregion

#region -- Helper Functions ---------------------------------------------------

function Update-Grid {
    # Fix 8: uses Get-FilteredResults (shared with exports)
    $filtered = Get-FilteredResults
    $grid.Rows.Clear()
    foreach ($u in $filtered) {
        $grid.Rows.Add($u.KB, $u.Title, $u.Status, $u.Severity, $u.Date, $u.SizeKB) | Out-Null
    }
}

function Show-Error {
    param([string]$Message)
    $lblError.Text    = $Message
    $pnlError.Visible = $true
}

function Hide-Error {
    $pnlError.Visible = $false
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

function Stop-ScanCleanup {
    # Fix 12: dispose runspace/PS/timer cleanly when scan aborts or form closes
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
}

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
        Stop-ScanCleanup
    }

    # Restore UI
    $pnlProgress.Visible = $false
    $btnScan.Enabled     = $true

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
        return
    }

    # Success path
    $script:ScanResults   = $res.Updates
    $script:IsPartialData = $res.IsPartial
    $script:RebootPending = $res.RebootPending
    $script:ScannedOS     = $res.OSVersion   # Fix 11: store target OS for HTML export

    # Fix 3: banner for WMI-only results
    if ($script:IsPartialData -and (-not $res.IsRemote)) {
        # Local scan fell back to WMI
        Show-Error 'WUA COM unavailable locally - showing installed hotfixes from WMI only.'
    } elseif ($script:IsPartialData) {
        # Remote scan always uses WMI (Fix 3: clear explanation)
        Show-Error ('Remote WUA scanning is not supported by Windows (access denied over WinRM) ' +
            '- showing installed hotfixes from WMI only; missing updates cannot be detected remotely.')
    } else {
        Hide-Error
    }

    Update-Summary
    Update-Grid

    $btnExportCsv.Enabled  = ($script:ScanResults.Count -gt 0)
    $btnExportHtml.Enabled = ($script:ScanResults.Count -gt 0)
    $countStr              = $script:ScanResults.Count
    $lblStatus.Text        = "Scan complete - $countStr updates found."
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
# Fix 3: Remote WUA fallback banner explains WUA-over-WinRM is access-denied.
# Fix 5: Probes 5985 then 5986; threads UseSSL through Invoke-Command.
# Fix 6: RebootPending returned as a separate flag; row Status is not modified.
# Fix 11: OSVersion gathered from target and returned in result object.

function Start-ScanAsync {
    param(
        [string]  $Target,
        [bool]    $IsRemote,
        [System.Management.Automation.PSCredential]$Credential,
        [bool]    $UseSSL
    )

    # Background script -- entirely self-contained; no access to main scope.
    $bgScript = [scriptblock]::Create('
        param($Target, $IsRemote, $Credential, $UseSSL)

        #-- inner helpers: must be defined here ---------------------

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
                }
            }
        }

        #-- result skeleton -----------------------------------------
        $result = [PSCustomObject]@{
            Updates       = @()
            IsPartial     = $false
            IsRemote      = $IsRemote
            RebootPending = $false
            OSVersion     = "Unknown"
            ErrorKind     = $null
            ErrorMessage  = $null
        }

        try {
            if ($IsRemote) {
                # Fix 5: probe 5985 then 5986
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

                # Fix 5: use SSL when only 5986 is available, or caller requested it
                $useSSLFlag = ($port5986 -and -not $port5985) -or $UseSSL

                # Build Invoke-Command param table
                $icParams = @{
                    ComputerName = $Target
                    ErrorAction  = "Stop"
                }
                if ($useSSLFlag) { $icParams.UseSSL     = $true }
                if ($Credential) { $icParams.Credential = $Credential }

                # Gather OS caption from target (Fix 11)
                try {
                    $icParams.ScriptBlock = {
                        (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
                    }
                    $osCaption = Invoke-Command @icParams
                    $result.OSVersion = if ($osCaption) { $osCaption } else { "Unknown" }
                } catch {
                    $result.OSVersion = "Unknown"
                }

                # Fix 3: Remote WUA (Microsoft.Update.Session over WinRM) is blocked by
                # Windows under network logon (E_ACCESSDENIED).  Go directly to WMI.
                $result.IsPartial = $true

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
                            }
                        }
                    }
                    $result.Updates = @(Invoke-Command @icParams)
                } catch {
                    $result.ErrorKind    = "SCAN_FAILED"
                    $result.ErrorMessage = $_.ToString()
                    return $result
                }

                # Pending reboot check on remote machine
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
            }
        } catch {
            $result.ErrorKind    = "SCAN_FAILED"
            $result.ErrorMessage = $_.ToString()
        }

        return $result
    ')

    # Create runspace + PowerShell instance (Fix 1: proper async pattern)
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.Open()

    $ps           = [powershell]::Create()
    $ps.Runspace  = $rs

    [void]$ps.AddScript($bgScript)
    [void]$ps.AddArgument($Target)
    [void]$ps.AddArgument($IsRemote)
    [void]$ps.AddArgument($Credential)
    [void]$ps.AddArgument($UseSSL)

    $script:RunningPS  = $ps
    $script:RunningRS  = $rs
    $script:PollHandle = $ps.BeginInvoke()

    # Fix 1: WinForms Timer polls on the UI thread (no cross-thread access)
    $timer          = New-Object System.Windows.Forms.Timer
    $timer.Interval = 500
    $timer.Add_Tick($scanPollTick)
    $script:PollTimer = $timer
    $timer.Start()
}

#endregion

#region -- Event Handlers -----------------------------------------------------

# Fix 4: disable Auth radio buttons for Local Machine; re-enable for Remote
$rbLocal.Add_CheckedChanged({
    if ($rbLocal.Checked) {
        $txtHost.Enabled       = $false
        $rbCurrentUser.Enabled = $false
        $rbCreds.Enabled       = $false
        $rbCurrentUser.Checked = $true
    }
})

$rbRemote.Add_CheckedChanged({
    if ($rbRemote.Checked) {
        $txtHost.Enabled       = $true
        $rbCurrentUser.Enabled = $true
        $rbCreds.Enabled       = $true
    }
})

$btnScan.Add_Click({
    Hide-Error
    $grid.Rows.Clear()
    $script:ScanResults   = @()
    $script:RebootPending = $false
    $btnExportCsv.Enabled  = $false
    $btnExportHtml.Enabled = $false
    $lblInstalled.Text    = 'Installed: -'
    $lblMissing.Text      = 'Missing: -'
    $lblPending.Text      = 'Reboot Pending: -'
    $lblPending.ForeColor = [System.Drawing.Color]::FromArgb(230, 81, 0)
    $lblScanTime.Text     = ''

    $isRemote = $rbRemote.Checked
    $target   = $txtHost.Text.Trim()

    if ($isRemote -and [string]::IsNullOrWhiteSpace($target)) {
        Show-Error 'Please enter a hostname or IP for remote scan.'
        return
    }

    $cred = $null
    if ($rbCreds.Checked) {
        $credTarget = if ($isRemote) { $target } else { 'local scan' }
        $cred = Get-Credential -Message "Enter credentials for $credTarget"
        if (-not $cred) { return }
    }

    $btnScan.Enabled     = $false
    $pnlProgress.Visible = $true
    if ($isRemote) {
        $lblStatus.Text = "Scanning $target..."
    } else {
        $lblStatus.Text = 'Scanning...'
    }

    Start-ScanAsync -Target $target -IsRemote $isRemote -Credential $cred -UseSSL $false
})

# Fix 10: color ENTIRE row background by status (not just the Status cell)
$grid.Add_CellFormatting({
    if ($_.RowIndex -ge 0) {
        $statusVal = $grid.Rows[$_.RowIndex].Cells['Status'].Value
        switch ($statusVal) {
            'Installed' { $_.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 255, 230) }
            'Missing'   { $_.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 230, 230) }
        }
    }
})

# Fix 9: KB URL construction uses anchored replace '^KB' so only the leading
# "KB" prefix is stripped (e.g. "KB1234" -> "1234"), not all occurrences.
$grid.Add_CellDoubleClick({
    if ($_.RowIndex -ge 0) {
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

# Clear
$btnClear.Add_Click({
    $grid.Rows.Clear()
    $script:ScanResults   = @()
    $script:RebootPending = $false
    $script:ScannedOS     = ''
    $btnExportCsv.Enabled  = $false
    $btnExportHtml.Enabled = $false
    $lblInstalled.Text    = 'Installed: -'
    $lblMissing.Text      = 'Missing: -'
    $lblPending.Text      = 'Reboot Pending: -'
    $lblPending.ForeColor = [System.Drawing.Color]::FromArgb(230, 81, 0)
    $lblScanTime.Text     = ''
    $lblStatus.Text       = 'Ready'
    Hide-Error
})

# Fix 12: clean up timer/runspace if the form closes during an active scan
$form.Add_FormClosing({
    Stop-ScanCleanup
})

#endregion

#region -- Entry Point --------------------------------------------------------
# Fix 4: Auth controls start disabled because Local Machine is the default
$rbCurrentUser.Enabled = $false
$rbCreds.Enabled       = $false

[void]$form.ShowDialog()
#endregion
