#Requires -Version 5.1
<#
.SYNOPSIS
    Windows Update Verification Tool — checks installed vs missing updates locally or on a remote server.
.DESCRIPTION
    WinForms GUI that queries the Windows Update API (WUA COM) for update status,
    displays results color-coded in a DataGridView, and exports to CSV or HTML.
.NOTES
    Run as Administrator for best results. Remote scanning requires WinRM on the target.
#>

#region ── Imports ────────────────────────────────────────────────────────────
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.ComponentModel
[System.Windows.Forms.Application]::EnableVisualStyles()
#endregion

#region ── Update Query Functions ────────────────────────────────────────────

function Test-WinRMReachable {
    param([string]$Target)
    try {
        $result = Test-NetConnection -ComputerName $Target -Port 5985 -WarningAction SilentlyContinue -ErrorAction Stop
        return $result.TcpTestSucceeded
    } catch {
        return $false
    }
}

function Test-PendingReboot {
    param([string]$Target = $null, [System.Management.Automation.PSCredential]$Credential = $null)
    $scriptBlock = {
        $keys = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
        )
        $pendingRename = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' `
            -Name PendingFileRenameOperations -ErrorAction SilentlyContinue).PendingFileRenameOperations
        foreach ($key in $keys) {
            if (Test-Path $key) { return $true }
        }
        if ($pendingRename) { return $true }
        return $false
    }
    try {
        if ($Target) {
            $params = @{ ComputerName = $Target; ScriptBlock = $scriptBlock; ErrorAction = 'Stop' }
            if ($Credential) { $params.Credential = $Credential }
            return Invoke-Command @params
        } else {
            return & $scriptBlock
        }
    } catch {
        return $false
    }
}

function Get-UpdatesViaWuaCom {
    param([string]$Target = $null, [System.Management.Automation.PSCredential]$Credential = $null)

    $scriptBlock = {
        $updates = @()
        try {
            $session  = New-Object -ComObject Microsoft.Update.Session
            $searcher = $session.CreateUpdateSearcher()
            $searcher.Online = $true

            $result = $searcher.Search("IsInstalled=0 OR IsInstalled=1")

            foreach ($u in $result.Updates) {
                $kb = if ($u.KBArticleIDs.Count -gt 0) { "KB$($u.KBArticleIDs[0])" } else { "N/A" }
                $updates += [PSCustomObject]@{
                    KB        = $kb
                    Title     = $u.Title
                    Status    = if ($u.IsInstalled) { "Installed" } else { "Missing" }
                    Severity  = if ($u.MsrcSeverity) { $u.MsrcSeverity } else { "N/A" }
                    Date      = if ($u.LastDeploymentChangeTime) { $u.LastDeploymentChangeTime.ToString("yyyy-MM-dd") } else { "N/A" }
                    SizeKB    = [math]::Round($u.MaxDownloadSize / 1KB, 0)
                    Source    = "WUA"
                }
            }
        } catch {
            throw "WUA COM Error: $_"
        }
        return $updates
    }

    if ($Target) {
        $params = @{ ComputerName = $Target; ScriptBlock = $scriptBlock; ErrorAction = 'Stop' }
        if ($Credential) { $params.Credential = $Credential }
        return Invoke-Command @params
    } else {
        return & $scriptBlock
    }
}

function Get-UpdatesFallback {
    param([string]$Target = $null, [System.Management.Automation.PSCredential]$Credential = $null)

    $scriptBlock = {
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

    if ($Target) {
        $params = @{ ComputerName = $Target; ScriptBlock = $scriptBlock; ErrorAction = 'Stop' }
        if ($Credential) { $params.Credential = $Credential }
        return Invoke-Command @params
    } else {
        return & $scriptBlock
    }
}

#endregion

#region ── Export Functions ──────────────────────────────────────────────────

function Export-ToCsv {
    param([object[]]$Records, [string]$Path)
    $Records | Select-Object KB, Title, Status, Severity, Date, SizeKB, Source |
        Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
}

function Export-ToHtml {
    param([object[]]$Records, [string]$Path, [string]$TargetName, [string]$OSVersion)

    $timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $installed  = ($Records | Where-Object Status -eq 'Installed').Count
    $missing    = ($Records | Where-Object Status -eq 'Missing').Count
    $pending    = ($Records | Where-Object Status -eq 'Pending Reboot').Count

    $rows = $Records | ForEach-Object {
        $color = switch ($_.Status) {
            'Installed'     { '#e6ffe6' }
            'Missing'       { '#ffe6e6' }
            'Pending Reboot'{ '#fff8d2' }
            default         { '#ffffff' }
        }
        "<tr style='background:$color'><td>$($_.KB)</td><td>$($_.Title)</td><td>$($_.Status)</td><td>$($_.Severity)</td><td>$($_.Date)</td><td>$($_.SizeKB)</td></tr>"
    }

    $html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Windows Update Report — $TargetName</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; color: #222; }
  h1   { color: #0078d4; }
  .summary { display:flex; gap:20px; margin:16px 0; }
  .badge { padding:10px 18px; border-radius:6px; font-weight:bold; font-size:14px; }
  .green  { background:#e6ffe6; border:1px solid #4caf50; }
  .red    { background:#ffe6e6; border:1px solid #f44336; }
  .orange { background:#fff8d2; border:1px solid #ff9800; }
  table { border-collapse:collapse; width:100%; font-size:13px; }
  th    { background:#0078d4; color:#fff; padding:8px 10px; text-align:left; }
  td    { padding:6px 10px; border-bottom:1px solid #ddd; }
  tr:hover td { filter: brightness(0.96); }
</style>
</head>
<body>
<h1>Windows Update Report</h1>
<p><strong>Target:</strong> $TargetName &nbsp;|&nbsp; <strong>OS:</strong> $OSVersion &nbsp;|&nbsp; <strong>Scanned:</strong> $timestamp</p>
<div class="summary">
  <div class="badge green">&#10003; Installed: $installed</div>
  <div class="badge red">&#10007; Missing: $missing</div>
  <div class="badge orange">&#9888; Pending Reboot: $pending</div>
</div>
<table>
<thead><tr><th>KB</th><th>Title</th><th>Status</th><th>Severity</th><th>Date</th><th>Size (KB)</th></tr></thead>
<tbody>
$($rows -join "`n")
</tbody>
</table>
</body>
</html>
"@
    $html | Out-File -FilePath $Path -Encoding UTF8
}

#endregion

#region ── WinForms UI ────────────────────────────────────────────────────────

# ── Shared state ──
$script:ScanResults = @()
$script:IsPartialData = $false

# ── Main Form ──
$form                  = New-Object System.Windows.Forms.Form
$form.Text             = "Windows Update Checker"
$form.Size             = New-Object System.Drawing.Size(900, 640)
$form.MinimumSize      = New-Object System.Drawing.Size(700, 500)
$form.StartPosition    = "CenterScreen"
$form.Font             = New-Object System.Drawing.Font("Segoe UI", 9)

# ── Target Panel ──
$pnlTarget             = New-Object System.Windows.Forms.Panel
$pnlTarget.Dock        = "Top"
$pnlTarget.Height      = 80
$pnlTarget.Padding     = New-Object System.Windows.Forms.Padding(10, 8, 10, 0)
$pnlTarget.BackColor   = [System.Drawing.Color]::FromArgb(240, 240, 240)

$lblTarget             = New-Object System.Windows.Forms.Label
$lblTarget.Text        = "Target:"
$lblTarget.Location    = New-Object System.Drawing.Point(10, 12)
$lblTarget.AutoSize    = $true
$lblTarget.Font        = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$rbLocal               = New-Object System.Windows.Forms.RadioButton
$rbLocal.Text          = "Local Machine"
$rbLocal.Location      = New-Object System.Drawing.Point(65, 10)
$rbLocal.AutoSize      = $true
$rbLocal.Checked       = $true

$rbRemote              = New-Object System.Windows.Forms.RadioButton
$rbRemote.Text         = "Remote:"
$rbRemote.Location     = New-Object System.Drawing.Point(185, 10)
$rbRemote.AutoSize     = $true

$txtHost               = New-Object System.Windows.Forms.TextBox
$txtHost.Location      = New-Object System.Drawing.Point(255, 8)
$txtHost.Width         = 200
$txtHost.Enabled       = $false
$txtHost.PlaceholderText = "hostname or IP"

$lblAuth               = New-Object System.Windows.Forms.Label
$lblAuth.Text          = "Auth:"
$lblAuth.Location      = New-Object System.Drawing.Point(10, 44)
$lblAuth.AutoSize      = $true
$lblAuth.Font          = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$rbCurrentUser         = New-Object System.Windows.Forms.RadioButton
$rbCurrentUser.Text    = "Current User"
$rbCurrentUser.Location= New-Object System.Drawing.Point(65, 42)
$rbCurrentUser.AutoSize= $true
$rbCurrentUser.Checked = $true

$rbCreds               = New-Object System.Windows.Forms.RadioButton
$rbCreds.Text          = "Specify Credentials"
$rbCreds.Location      = New-Object System.Drawing.Point(185, 42)
$rbCreds.AutoSize      = $true

$btnScan               = New-Object System.Windows.Forms.Button
$btnScan.Text          = "Scan Now"
$btnScan.Location      = New-Object System.Drawing.Point(750, 20)
$btnScan.Size          = New-Object System.Drawing.Size(100, 40)
$btnScan.BackColor     = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnScan.ForeColor     = [System.Drawing.Color]::White
$btnScan.FlatStyle     = "Flat"
$btnScan.Font          = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnScan.Anchor        = "Top,Right"

$pnlTarget.Controls.AddRange(@($lblTarget, $rbLocal, $rbRemote, $txtHost, $lblAuth, $rbCurrentUser, $rbCreds, $btnScan))

# ── Summary Panel ──
$pnlSummary            = New-Object System.Windows.Forms.Panel
$pnlSummary.Dock       = "Top"
$pnlSummary.Height     = 60
$pnlSummary.Padding    = New-Object System.Windows.Forms.Padding(10, 8, 10, 0)
$pnlSummary.BackColor  = [System.Drawing.Color]::White

$lblInstalled          = New-Object System.Windows.Forms.Label
$lblInstalled.Text     = "Installed: —"
$lblInstalled.Location = New-Object System.Drawing.Point(10, 10)
$lblInstalled.AutoSize = $true
$lblInstalled.ForeColor= [System.Drawing.Color]::FromArgb(46, 125, 50)
$lblInstalled.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

$lblMissing            = New-Object System.Windows.Forms.Label
$lblMissing.Text       = "Missing: —"
$lblMissing.Location   = New-Object System.Drawing.Point(160, 10)
$lblMissing.AutoSize   = $true
$lblMissing.ForeColor  = [System.Drawing.Color]::FromArgb(198, 40, 40)
$lblMissing.Font       = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

$lblPending            = New-Object System.Windows.Forms.Label
$lblPending.Text       = "Pending Reboot: —"
$lblPending.Location   = New-Object System.Drawing.Point(310, 10)
$lblPending.AutoSize   = $true
$lblPending.ForeColor  = [System.Drawing.Color]::FromArgb(230, 81, 0)
$lblPending.Font       = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

$lblScanTime           = New-Object System.Windows.Forms.Label
$lblScanTime.Text      = ""
$lblScanTime.Location  = New-Object System.Drawing.Point(10, 36)
$lblScanTime.AutoSize  = $true
$lblScanTime.ForeColor = [System.Drawing.Color]::Gray

$pnlSummary.Controls.AddRange(@($lblInstalled, $lblMissing, $lblPending, $lblScanTime))

# ── Progress Panel ──
$pnlProgress           = New-Object System.Windows.Forms.Panel
$pnlProgress.Dock      = "Top"
$pnlProgress.Height    = 28
$pnlProgress.Padding   = New-Object System.Windows.Forms.Padding(10, 2, 10, 0)
$pnlProgress.Visible   = $false

$progressBar           = New-Object System.Windows.Forms.ProgressBar
$progressBar.Dock      = "Fill"
$progressBar.Style     = "Marquee"
$pnlProgress.Controls.Add($progressBar)

# ── Error Banner ──
$pnlError              = New-Object System.Windows.Forms.Panel
$pnlError.Dock         = "Top"
$pnlError.Height       = 30
$pnlError.BackColor    = [System.Drawing.Color]::FromArgb(255, 230, 230)
$pnlError.Visible      = $false
$pnlError.Padding      = New-Object System.Windows.Forms.Padding(10, 5, 10, 0)

$lblError              = New-Object System.Windows.Forms.Label
$lblError.Dock         = "Fill"
$lblError.ForeColor    = [System.Drawing.Color]::FromArgb(198, 40, 40)
$lblError.Font         = New-Object System.Drawing.Font("Segoe UI", 9)
$pnlError.Controls.Add($lblError)

# ── Filter Panel ──
$pnlFilter             = New-Object System.Windows.Forms.Panel
$pnlFilter.Dock        = "Top"
$pnlFilter.Height      = 34
$pnlFilter.Padding     = New-Object System.Windows.Forms.Padding(10, 4, 10, 0)
$pnlFilter.BackColor   = [System.Drawing.Color]::FromArgb(250, 250, 250)

$lblFilterStatus       = New-Object System.Windows.Forms.Label
$lblFilterStatus.Text  = "Status:"
$lblFilterStatus.Location = New-Object System.Drawing.Point(10, 7)
$lblFilterStatus.AutoSize = $true

$cmbStatus             = New-Object System.Windows.Forms.ComboBox
$cmbStatus.Location    = New-Object System.Drawing.Point(55, 4)
$cmbStatus.Width       = 130
$cmbStatus.DropDownStyle = "DropDownList"
$cmbStatus.Items.AddRange(@("All", "Installed", "Missing", "Pending Reboot"))
$cmbStatus.SelectedIndex = 0

$lblFilterSev          = New-Object System.Windows.Forms.Label
$lblFilterSev.Text     = "Severity:"
$lblFilterSev.Location = New-Object System.Drawing.Point(200, 7)
$lblFilterSev.AutoSize = $true

$cmbSeverity           = New-Object System.Windows.Forms.ComboBox
$cmbSeverity.Location  = New-Object System.Drawing.Point(252, 4)
$cmbSeverity.Width     = 120
$cmbSeverity.DropDownStyle = "DropDownList"
$cmbSeverity.Items.AddRange(@("All", "Critical", "Important", "Moderate", "Low", "N/A"))
$cmbSeverity.SelectedIndex = 0

$pnlFilter.Controls.AddRange(@($lblFilterStatus, $cmbStatus, $lblFilterSev, $cmbSeverity))

# ── DataGridView ──
$grid                  = New-Object System.Windows.Forms.DataGridView
$grid.Dock             = "Fill"
$grid.ReadOnly         = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.RowHeadersVisible = $false
$grid.AutoSizeColumnsMode = "Fill"
$grid.SelectionMode    = "FullRowSelect"
$grid.BackgroundColor  = [System.Drawing.Color]::White
$grid.BorderStyle      = "None"
$grid.ColumnHeadersHeightSizeMode = "AutoSize"
$grid.Font             = New-Object System.Drawing.Font("Segoe UI", 9)

$cols = @(
    @{ Name="KB";       Header="KB";         Width=90;  Fill=$false },
    @{ Name="Title";    Header="Title";      Width=350; Fill=$true  },
    @{ Name="Status";   Header="Status";     Width=110; Fill=$false },
    @{ Name="Severity"; Header="Severity";   Width=90;  Fill=$false },
    @{ Name="Date";     Header="Date";       Width=90;  Fill=$false },
    @{ Name="SizeKB";   Header="Size (KB)";  Width=80;  Fill=$false }
)
foreach ($c in $cols) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name       = $c.Name
    $col.HeaderText = $c.Header
    if ($c.Fill) {
        $col.AutoSizeMode = "Fill"
    } else {
        $col.Width = $c.Width
        $col.AutoSizeMode = "None"
    }
    $grid.Columns.Add($col) | Out-Null
}

# ── Bottom Panel ──
$pnlBottom             = New-Object System.Windows.Forms.Panel
$pnlBottom.Dock        = "Bottom"
$pnlBottom.Height      = 40
$pnlBottom.Padding     = New-Object System.Windows.Forms.Padding(10, 5, 10, 5)
$pnlBottom.BackColor   = [System.Drawing.Color]::FromArgb(240, 240, 240)

$btnExportCsv          = New-Object System.Windows.Forms.Button
$btnExportCsv.Text     = "Export CSV"
$btnExportCsv.Location = New-Object System.Drawing.Point(10, 5)
$btnExportCsv.Size     = New-Object System.Drawing.Size(100, 28)
$btnExportCsv.Enabled  = $false

$btnExportHtml         = New-Object System.Windows.Forms.Button
$btnExportHtml.Text    = "Export HTML"
$btnExportHtml.Location= New-Object System.Drawing.Point(118, 5)
$btnExportHtml.Size    = New-Object System.Drawing.Size(100, 28)
$btnExportHtml.Enabled = $false

$btnClear              = New-Object System.Windows.Forms.Button
$btnClear.Text         = "Clear"
$btnClear.Location     = New-Object System.Drawing.Point(226, 5)
$btnClear.Size         = New-Object System.Drawing.Size(80, 28)

$lblStatus             = New-Object System.Windows.Forms.Label
$lblStatus.Text        = "Ready"
$lblStatus.Location    = New-Object System.Drawing.Point(320, 10)
$lblStatus.AutoSize    = $true
$lblStatus.ForeColor   = [System.Drawing.Color]::Gray

$pnlBottom.Controls.AddRange(@($btnExportCsv, $btnExportHtml, $btnClear, $lblStatus))

# ── Assemble Form ──
$form.Controls.AddRange(@($grid, $pnlBottom, $pnlFilter, $pnlError, $pnlProgress, $pnlSummary, $pnlTarget))

#endregion

#region ── Helper: Apply Filters ─────────────────────────────────────────────

function Update-Grid {
    $statusFilter   = $cmbStatus.SelectedItem
    $severityFilter = $cmbSeverity.SelectedItem

    $filtered = $script:ScanResults | Where-Object {
        ($statusFilter   -eq "All" -or $_.Status   -eq $statusFilter) -and
        ($severityFilter -eq "All" -or $_.Severity -eq $severityFilter)
    }

    $grid.Rows.Clear()
    foreach ($u in $filtered) {
        $grid.Rows.Add($u.KB, $u.Title, $u.Status, $u.Severity, $u.Date, $u.SizeKB) | Out-Null
    }
}

function Show-Error {
    param([string]$Message)
    $lblError.Text   = $Message
    $pnlError.Visible = $true
}

function Hide-Error {
    $pnlError.Visible = $false
}

function Update-Summary {
    $installed = ($script:ScanResults | Where-Object Status -eq 'Installed').Count
    $missing   = ($script:ScanResults | Where-Object Status -eq 'Missing').Count
    $pending   = ($script:ScanResults | Where-Object Status -eq 'Pending Reboot').Count

    $lblInstalled.Text = "Installed: $installed"
    $lblMissing.Text   = "Missing: $missing"
    $lblPending.Text   = "Pending Reboot: $pending"
    $lblScanTime.Text  = "Last scan: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" +
                         $(if ($script:IsPartialData) { "  [Partial data — WUA COM unavailable, showing WMI results]" } else { "" })
}

#endregion

#region ── BackgroundWorker ──────────────────────────────────────────────────

$bgWorker = New-Object System.ComponentModel.BackgroundWorker
$bgWorker.WorkerReportsProgress      = $true
$bgWorker.WorkerSupportsCancellation = $true

$bgWorker.Add_DoWork({
    param($sender, $e)
    $args      = $e.Argument
    $target    = $args.Target
    $cred      = $args.Credential
    $isRemote  = $args.IsRemote

    $updates = @()
    $partial  = $false

    # WinRM check for remote
    if ($isRemote) {
        $reachable = Test-WinRMReachable -Target $target
        if (-not $reachable) {
            throw "WinRM_UNREACHABLE:$target"
        }
    }

    # Try WUA COM
    try {
        $updates = Get-UpdatesViaWuaCom -Target $(if ($isRemote) { $target } else { $null }) -Credential $cred
    } catch {
        # Fall back to WMI
        $partial = $true
        try {
            $updates = Get-UpdatesFallback -Target $(if ($isRemote) { $target } else { $null }) -Credential $cred
        } catch {
            throw "SCAN_FAILED:$_"
        }
    }

    # Check pending reboot and upgrade status
    $rebootPending = Test-PendingReboot -Target $(if ($isRemote) { $target } else { $null }) -Credential $cred
    if ($rebootPending) {
        $updates = $updates | ForEach-Object {
            if ($_.Status -eq 'Installed') { $_.Status = 'Pending Reboot' }
            $_
        }
    }

    $e.Result = [PSCustomObject]@{
        Updates     = $updates
        IsPartial   = $partial
        RebootPending = $rebootPending
    }
})

$bgWorker.Add_RunWorkerCompleted({
    param($sender, $e)
    $pnlProgress.Visible = $false
    $btnScan.Enabled     = $true

    if ($e.Error) {
        $msg = $e.Error.Message
        if ($msg -match "WinRM_UNREACHABLE:(.+)") {
            Show-Error "WinRM unreachable on '$($Matches[1])' — enable it with: winrm quickconfig"
        } elseif ($msg -match "SCAN_FAILED:(.+)") {
            Show-Error "Scan failed: $($Matches[1])"
        } else {
            Show-Error "Error: $msg"
        }
        $lblStatus.Text = "Scan failed."
        return
    }

    $script:ScanResults   = $e.Result.Updates
    $script:IsPartialData = $e.Result.IsPartial

    if ($script:IsPartialData) {
        Show-Error "Partial data — WUA COM unavailable, showing WMI (Get-HotFix) results only."
    } else {
        Hide-Error
    }

    Update-Summary
    Update-Grid

    $btnExportCsv.Enabled  = ($script:ScanResults.Count -gt 0)
    $btnExportHtml.Enabled = ($script:ScanResults.Count -gt 0)
    $lblStatus.Text        = "Scan complete — $($script:ScanResults.Count) updates found."
})

#endregion

#region ── Event Handlers ────────────────────────────────────────────────────

$rbRemote.Add_CheckedChanged({
    $txtHost.Enabled = $rbRemote.Checked
})

$rbLocal.Add_CheckedChanged({
    $txtHost.Enabled = $rbRemote.Checked
})

$btnScan.Add_Click({
    Hide-Error
    $grid.Rows.Clear()
    $script:ScanResults = @()
    $btnExportCsv.Enabled  = $false
    $btnExportHtml.Enabled = $false
    $lblInstalled.Text = "Installed: —"
    $lblMissing.Text   = "Missing: —"
    $lblPending.Text   = "Pending Reboot: —"
    $lblScanTime.Text  = ""

    $isRemote = $rbRemote.Checked
    $target   = $txtHost.Text.Trim()

    if ($isRemote -and [string]::IsNullOrWhiteSpace($target)) {
        Show-Error "Please enter a hostname or IP for remote scan."
        return
    }

    $cred = $null
    if ($rbCreds.Checked) {
        $cred = Get-Credential -Message "Enter credentials for $( if ($isRemote) { $target } else { 'local scan' })"
        if (-not $cred) { return }
    }

    $btnScan.Enabled     = $false
    $pnlProgress.Visible = $true
    $lblStatus.Text      = "Scanning$(if ($isRemote) { " $target" })..."

    $bgWorker.RunWorkerAsync([PSCustomObject]@{
        Target    = $target
        Credential = $cred
        IsRemote  = $isRemote
    })
})

# Color coding
$grid.Add_CellFormatting({
    if ($_.ColumnIndex -eq $grid.Columns['Status'].Index -and $_.RowIndex -ge 0) {
        switch ($_.Value) {
            'Installed'      { $_.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 255, 230) }
            'Missing'        { $_.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 230, 230) }
            'Pending Reboot' { $_.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 210) }
        }
    }
})

# Detail popup on double-click
$grid.Add_CellDoubleClick({
    if ($_.RowIndex -ge 0) {
        $kb    = $grid.Rows[$_.RowIndex].Cells['KB'].Value
        $title = $grid.Rows[$_.RowIndex].Cells['Title'].Value
        $status= $grid.Rows[$_.RowIndex].Cells['Status'].Value
        $sev   = $grid.Rows[$_.RowIndex].Cells['Severity'].Value
        $date  = $grid.Rows[$_.RowIndex].Cells['Date'].Value
        $size  = $grid.Rows[$_.RowIndex].Cells['SizeKB'].Value

        $detail = New-Object System.Windows.Forms.Form
        $detail.Text = "Update Detail — $kb"
        $detail.Size = New-Object System.Drawing.Size(500, 280)
        $detail.StartPosition = "CenterParent"
        $detail.FormBorderStyle = "FixedDialog"
        $detail.MaximizeBox = $false

        $txt = New-Object System.Windows.Forms.RichTextBox
        $txt.Dock = "Fill"
        $txt.ReadOnly = $true
        $txt.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $txt.Text = "KB Article : $kb`r`nTitle      : $title`r`nStatus     : $status`r`nSeverity   : $sev`r`nDate       : $date`r`nSize       : $size KB`r`n`r`nMicrosoft Support URL:`r`nhttps://support.microsoft.com/kb/$($kb -replace 'KB','')"

        $btnClose = New-Object System.Windows.Forms.Button
        $btnClose.Text = "Close"
        $btnClose.Dock = "Bottom"
        $btnClose.Add_Click({ $detail.Close() })

        $detail.Controls.AddRange(@($txt, $btnClose))
        $detail.ShowDialog($form) | Out-Null
    }
})

# Filters
$cmbStatus.Add_SelectedIndexChanged({ Update-Grid })
$cmbSeverity.Add_SelectedIndexChanged({ Update-Grid })

# Export CSV
$btnExportCsv.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = "CSV files (*.csv)|*.csv"
    $dlg.FileName = "WindowsUpdates_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($dlg.ShowDialog() -eq "OK") {
        Export-ToCsv -Records $script:ScanResults -Path $dlg.FileName
        $lblStatus.Text = "Exported CSV to $($dlg.FileName)"
    }
})

# Export HTML
$btnExportHtml.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = "HTML files (*.html)|*.html"
    $dlg.FileName = "WindowsUpdates_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    if ($dlg.ShowDialog() -eq "OK") {
        $targetName = if ($rbRemote.Checked) { $txtHost.Text } else { $env:COMPUTERNAME }
        $osVer = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
        Export-ToHtml -Records $script:ScanResults -Path $dlg.FileName -TargetName $targetName -OSVersion $osVer
        $lblStatus.Text = "Exported HTML to $($dlg.FileName)"
        Start-Process $dlg.FileName
    }
})

# Clear
$btnClear.Add_Click({
    $grid.Rows.Clear()
    $script:ScanResults    = @()
    $btnExportCsv.Enabled  = $false
    $btnExportHtml.Enabled = $false
    $lblInstalled.Text     = "Installed: —"
    $lblMissing.Text       = "Missing: —"
    $lblPending.Text       = "Pending Reboot: —"
    $lblScanTime.Text      = ""
    $lblStatus.Text        = "Ready"
    Hide-Error
})

#endregion

#region ── Entry Point ───────────────────────────────────────────────────────
[void]$form.ShowDialog()
#endregion
