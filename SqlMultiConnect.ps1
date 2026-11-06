<#
.SYNOPSIS
    SQL Server multi-connection tester (Windows Forms).

.DESCRIPTION
    - Test a single connection and see server details.
    - Open N connections from one input value (sequential or parallel, optional ramp delay).
    - Hold the connections open, each tagged with its own Application Name and SPID.
    - Run a query on the selected connection or on ALL open connections at once.
    - View results in a grid and export to CSV.

    Run with Windows PowerShell 5.1:
        powershell.exe -STA -ExecutionPolicy Bypass -File .\SqlMultiConnect.ps1

    NOTE: Opening many connections puts real load on the target server.
          Use only on servers you are authorized to test, and agree on counts first for production.
#>

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    & powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -File $MyInvocation.MyCommand.Path
    return
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data
[System.Windows.Forms.Application]::EnableVisualStyles()

# ------------------------------------------------------------------ state
$script:Conns     = New-Object System.Collections.ArrayList   # hashtables: Id, App, Conn, Item
$script:NextId    = 0
$script:Cancel    = $false
$script:LastTable = $null

# ------------------------------------------------------------------ helpers
function Pump { [System.Windows.Forms.Application]::DoEvents() }

function Wait-Ms([int]$Ms) {
    $end = [DateTime]::UtcNow.AddMilliseconds($Ms)

    while ([DateTime]::UtcNow -lt $end) { Pump; Start-Sleep -Milliseconds 10 }
}

function Write-Log([string]$Message) {
    $txtLog.AppendText(("[{0:HH:mm:ss.fff}] {1}`r`n" -f (Get-Date), $Message))
}

function New-ConnString([string]$AppName) {
    $b = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $b['Data Source']             = $txtServer.Text.Trim()
    if ($txtDb.Text.Trim()) { $b['Initial Catalog'] = $txtDb.Text.Trim() }
    $b['Connect Timeout']         = [int]$numConnTimeout.Value
    $b['Encrypt']                 = [bool]$chkEncrypt.Checked
    $b['TrustServerCertificate']  = [bool]$chkTrust.Checked
    $b['Pooling']                 = -not $chkNoPool.Checked
    $b['Application Name']        = $AppName
    if ($cboAuth.SelectedIndex -eq 0) {
        $b['Integrated Security'] = $true
    } else {
        $b['User ID']  = $txtUser.Text
        $b['Password'] = $txtPwd.Text
    }
    return $b.ConnectionString
}

function Update-Count {
    $open = 0
    foreach ($e in $script:Conns) {
        if ($e.Conn -and $e.Conn.State -eq 'Open') { $open++ }
    }
    $lblCount.Text = "Open: $open  |  Total listed: $($script:Conns.Count)"
}

function Set-Row {
    param($E, $Status, $Spid, $OpenMs, $Query, $Msg, $Color)
    $s = $E.Item.SubItems
    if ($PSBoundParameters.ContainsKey('Status'))  { $s[2].Text = [string]$Status }
    if ($PSBoundParameters.ContainsKey('Spid'))    { $s[3].Text = [string]$Spid }
    if ($PSBoundParameters.ContainsKey('OpenMs'))  { $s[4].Text = [string]$OpenMs }
    if ($PSBoundParameters.ContainsKey('Query'))   { $s[5].Text = [string]$Query }
    if ($PSBoundParameters.ContainsKey('Msg'))     { $s[6].Text = [string]$Msg }
    if ($PSBoundParameters.ContainsKey('Color'))   { $E.Item.ForeColor = $Color }
}

function Set-Busy([bool]$Busy) {
    foreach ($c in @($btnOpen, $btnCloseSel, $btnCloseAll, $btnTest, $btnRunSel, $btnRunAll)) { $c.Enabled = -not $Busy }
    $btnCancel.Enabled = $Busy
    $form.UseWaitCursor = $false
}

function Complete-Open($E, [long]$Ms) {
    $spid = ''
    try {
        $cmd = $E.Conn.CreateCommand()
        $cmd.CommandText = 'SELECT @@SPID'
        $spid = $cmd.ExecuteScalar()
    } catch { $spid = '?' }
    Set-Row -E $E -Status 'Open' -Spid $spid -OpenMs $Ms -Msg '' -Color ([System.Drawing.Color]::DarkGreen)
}

function Fail-Open($E, [long]$Ms, [string]$Message) {
    Set-Row -E $E -Status 'Failed' -OpenMs $Ms -Msg $Message -Color ([System.Drawing.Color]::Firebrick)
    Write-Log ("Connection {0} FAILED: {1}" -f $E.Id, $Message)
}

# ------------------------------------------------------------------ actions
function Test-OneConnection {
    if (-not $txtServer.Text.Trim()) { [void][System.Windows.Forms.MessageBox]::Show('Enter a server name.'); return }
    Set-Busy $true
    $conn = $null
    try {
        $app  = '{0}-TEST' -f $txtApp.Text.Trim()
        $conn = New-Object System.Data.SqlClient.SqlConnection -ArgumentList (New-ConnString $app)
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $conn.Open()
        $openMs = $sw.ElapsedMilliseconds
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = @"
SELECT @@SERVERNAME AS ServerName,
       CONVERT(nvarchar(128), SERVERPROPERTY('ProductVersion')) AS ProductVersion,
       CONVERT(nvarchar(128), SERVERPROPERTY('Edition')) AS Edition,
       DB_NAME() AS CurrentDatabase,
       SUSER_SNAME() AS LoginName,
       @@SPID AS SPID,
       CONVERT(nvarchar(60), CONNECTIONPROPERTY('net_transport')) AS NetTransport,
       CONVERT(nvarchar(60), CONNECTIONPROPERTY('auth_scheme')) AS AuthScheme,
       SYSDATETIME() AS ServerTime;
"@
        $dt = New-Object System.Data.DataTable
        $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        [void]$da.Fill($dt)
        $grid.DataSource = $dt
        $script:LastTable = $dt
        $tabs.SelectedTab = $tabResults
        Write-Log ("TEST OK in {0} ms | {1} | v{2}" -f $openMs, $dt.Rows[0].ServerName, $dt.Rows[0].ProductVersion)
        $lblStatus.Text = "Test connection succeeded ($openMs ms)"
    } catch {
        Write-Log ("TEST FAILED: {0}" -f $_.Exception.Message)
        $lblStatus.Text = 'Test connection FAILED - see Log tab'
        $tabs.SelectedTab = $tabLog
    } finally {
        if ($conn) { $conn.Dispose() }
        Set-Busy $false
    }
}

function Start-OpenConnections {
    if (-not $txtServer.Text.Trim()) { [void][System.Windows.Forms.MessageBox]::Show('Enter a server name.'); return }
    $n     = [int]$numCount.Value
    $delay = [int]$numDelay.Value
    $par   = $chkParallel.Checked
    $script:Cancel = $false
    Set-Busy $true
    Write-Log ("Opening {0} connection(s) to {1} | parallel={2} | delay={3} ms | pooling={4}" -f $n, $txtServer.Text.Trim(), $par, $delay, (-not $chkNoPool.Checked))
    $pending = New-Object System.Collections.ArrayList
    $batchSw = [Diagnostics.Stopwatch]::StartNew()
    try {
        for ($k = 0; $k -lt $n; $k++) {
            if ($script:Cancel) { Write-Log "Cancelled after starting $k connection(s)."; break }
            $script:NextId++
            $id  = $script:NextId
            $app = '{0}-{1:D3}' -f $txtApp.Text.Trim(), $id
            $conn = New-Object System.Data.SqlClient.SqlConnection -ArgumentList (New-ConnString $app)
            $item = New-Object System.Windows.Forms.ListViewItem -ArgumentList (, [string[]]@("$id", $app, 'Opening', '', '', '', ''))
            $entry = @{ Id = $id; App = $app; Conn = $conn; Item = $item }
            [void]$script:Conns.Add($entry)
            [void]$lv.Items.Add($item)
            $sw = [Diagnostics.Stopwatch]::StartNew()
            if ($par) {
                [void]$pending.Add(@{ Entry = $entry; Task = $conn.OpenAsync(); Sw = $sw })
            } else {
                try {
                    $conn.Open(); $sw.Stop()
                    Complete-Open $entry $sw.ElapsedMilliseconds
                } catch {
                    $sw.Stop()
                    Fail-Open $entry $sw.ElapsedMilliseconds $_.Exception.Message
                }
            }
            $lblStatus.Text = "Started $($k + 1) of $n"
            if ($delay -gt 0 -and $k -lt ($n - 1)) { Wait-Ms $delay } else { Pump }
        }

        while ($pending.Count -gt 0) {
            Pump
            foreach ($p in @($pending)) {
                if ($p.Task.IsCompleted) {
                    $p.Sw.Stop()
                    if ($p.Task.IsFaulted) {
                        Fail-Open $p.Entry $p.Sw.ElapsedMilliseconds $p.Task.Exception.GetBaseException().Message
                    } elseif ($p.Task.IsCanceled) {
                        Fail-Open $p.Entry $p.Sw.ElapsedMilliseconds 'Canceled'
                    } else {
                        Complete-Open $p.Entry $p.Sw.ElapsedMilliseconds
                    }
                    $pending.Remove($p)
                }
            }
            Start-Sleep -Milliseconds 25
        }
    } catch {
        Write-Log ("Open error: {0}" -f $_.Exception.Message)
    } finally {
        $batchSw.Stop()
        Update-Count
        Write-Log ("Open batch finished in {0} ms. {1}" -f $batchSw.ElapsedMilliseconds, $lblCount.Text)
        $lblStatus.Text = "Done. $($lblCount.Text)"
        Set-Busy $false
    }
}

function Get-SelectedEntries {
    $ids = @()
    foreach ($i in $lv.SelectedItems) { $ids += [int]$i.Text }
    return @($script:Conns | Where-Object { $ids -contains $_.Id })
}

function Close-Entries($Entries) {
    foreach ($e in @($Entries)) {
        try { if ($e.Conn) { $e.Conn.Dispose() } } catch { }
        $lv.Items.Remove($e.Item)
        [void]$script:Conns.Remove($e)
    }
    Update-Count
}

function Invoke-Query([string]$Mode) {
    $sql = $txtQuery.Text
    if (-not $sql.Trim()) { return }
    if ($Mode -eq 'Selected') {
        $targets = @(Get-SelectedEntries | Where-Object { $_.Conn.State -eq 'Open' })
        if ($targets.Count -eq 0) {
            $targets = @($script:Conns | Where-Object { $_.Conn.State -eq 'Open' } | Select-Object -First 1)
        }
    } else {
        $targets = @($script:Conns | Where-Object { $_.Conn.State -eq 'Open' })
    }
    if ($targets.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('No open connections. Use "Open N Connections" first.')
        return
    }

    Set-Busy $true
    $jobs = New-Object System.Collections.ArrayList
    $cmdTimeout = [int]$numCmdTimeout.Value
    try {
        Write-Log ("Running query on {0} connection(s)." -f $targets.Count)
        foreach ($t in $targets) {
            try {
                Set-Row -E $t -Query 'Running...' -Msg ''
                $cmd = $t.Conn.CreateCommand()
                $cmd.CommandText    = $sql
                $cmd.CommandTimeout = $cmdTimeout
                $sw = [Diagnostics.Stopwatch]::StartNew()
                [void]$jobs.Add(@{ Entry = $t; Task = $cmd.ExecuteReaderAsync(); Sw = $sw; Done = $false })
            } catch {
                Set-Row -E $t -Query 'Error' -Msg $_.Exception.Message -Color ([System.Drawing.Color]::Firebrick)
            }
        }

        # wait for all to finish while keeping UI alive
        while (@($jobs | Where-Object { -not $_.Done }).Count -gt 0) {
            Pump
            foreach ($j in $jobs) {
                if (-not $j.Done -and $j.Task.IsCompleted) { $j.Sw.Stop(); $j.Done = $true }
            }
            if ($script:Cancel) { break }
            Start-Sleep -Milliseconds 25
        }

        $shown = $false
        foreach ($j in $jobs) {
            if (-not $j.Done) { continue }
            $e = $j.Entry
            if ($j.Task.IsFaulted) {
                $msg = $j.Task.Exception.GetBaseException().Message
                Set-Row -E $e -Query ("{0} ms" -f $j.Sw.ElapsedMilliseconds) -Msg $msg -Color ([System.Drawing.Color]::Firebrick)
                Write-Log ("Conn {0} query error: {1}" -f $e.Id, $msg)
                continue
            }
            $reader = $j.Task.Result
            $dt = New-Object System.Data.DataTable
            try { $dt.Load($reader) } finally { $reader.Dispose() }
            Set-Row -E $e -Query ("{0} ms / {1} rows" -f $j.Sw.ElapsedMilliseconds, $dt.Rows.Count) -Msg '' -Color ([System.Drawing.Color]::DarkGreen)
            if (-not $shown) {
                $grid.DataSource = $dt
                $script:LastTable = $dt
                $tabs.SelectedTab = $tabResults
                $shown = $true
            }
        }
        $lblStatus.Text = 'Query complete. Grid shows the first successful result.'
    } catch {
        Write-Log ("Query error: {0}" -f $_.Exception.Message)
    } finally {
        Update-Count
        Set-Busy $false
    }
}

function Export-Results {
    if (-not $script:LastTable -or $script:LastTable.Rows.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('No results to export.'); return
    }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = 'CSV (*.csv)|*.csv'
    $dlg.FileName = 'SqlResults.csv'
    if ($dlg.ShowDialog() -eq 'OK') {
        $cols = @($script:LastTable.Columns | ForEach-Object { $_.ColumnName })
        $script:LastTable | Select-Object -Property $cols | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
        Write-Log "Exported results to $($dlg.FileName)"
    }
}

# ------------------------------------------------------------------ UI
$font = New-Object System.Drawing.Font('Segoe UI', 9)
$form = New-Object System.Windows.Forms.Form
$form.Text          = 'SQL Server Multi-Connection Tester'
$form.Size          = New-Object System.Drawing.Size(1150, 820)
$form.MinimumSize   = New-Object System.Drawing.Size(980, 700)
$form.StartPosition = 'CenterScreen'
$form.Font          = $font

function New-Label($text, $x, $y) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text; $l.AutoSize = $true; $l.Location = New-Object System.Drawing.Point($x, $y)
    return $l
}
function New-TextBox($x, $y, $w, $text = '') {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point($x, $y); $t.Width = $w; $t.Text = $text
    return $t
}
function New-Numeric($x, $y, $w, $min, $max, $val) {
    $n = New-Object System.Windows.Forms.NumericUpDown
    $n.Location = New-Object System.Drawing.Point($x, $y); $n.Width = $w
    $n.Minimum = $min; $n.Maximum = $max; $n.Value = $val
    return $n
}
function New-Button($text, $x, $y, $w) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text; $b.Location = New-Object System.Drawing.Point($x, $y); $b.Size = New-Object System.Drawing.Size($w, 28)
    return $b
}
function New-Check($text, $x, $y, $checked) {
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text = $text; $c.AutoSize = $true; $c.Location = New-Object System.Drawing.Point($x, $y); $c.Checked = $checked
    return $c
}

# ---- Connection group
$grpConn = New-Object System.Windows.Forms.GroupBox
$grpConn.Text = 'Connection'; $grpConn.Dock = 'Top'; $grpConn.Height = 135

$txtServer = New-TextBox 75 24 260 'localhost'
$txtDb     = New-TextBox 420 24 150 'master'
$cboAuth   = New-Object System.Windows.Forms.ComboBox
$cboAuth.DropDownStyle = 'DropDownList'; $cboAuth.Location = New-Object System.Drawing.Point(660, 24); $cboAuth.Width = 170
[void]$cboAuth.Items.AddRange(@('Windows Authentication', 'SQL Server Authentication'))
$cboAuth.SelectedIndex = 0
$txtUser = New-TextBox 75 58 200
$txtPwd  = New-TextBox 360 58 200
$txtPwd.UseSystemPasswordChar = $true
$txtUser.Enabled = $false; $txtPwd.Enabled = $false
$cboAuth.Add_SelectedIndexChanged({
    $sql = ($cboAuth.SelectedIndex -eq 1)
    $txtUser.Enabled = $sql; $txtPwd.Enabled = $sql
})

$chkEncrypt = New-Check 'Encrypt' 15 95 $true
$chkTrust   = New-Check 'Trust server certificate' 95 95 $true
$chkNoPool  = New-Check 'Disable pooling' 270 95 $true
$numConnTimeout = New-Numeric 520 92 55 1 600 15
$numCmdTimeout  = New-Numeric 700 92 55 0 36000 60
$txtApp = New-TextBox 880 92 150 'SqlMultiConn'

$grpConn.Controls.AddRange(@(
    (New-Label 'Server' 15 27), $txtServer,
    (New-Label 'Database' 350 27), $txtDb,
    (New-Label 'Auth' 620 27), $cboAuth,
    (New-Label 'User' 15 61), $txtUser,
    (New-Label 'Password' 295 61), $txtPwd,
    $chkEncrypt, $chkTrust, $chkNoPool,
    (New-Label 'Connect timeout (s)' 400 95), $numConnTimeout,
    (New-Label 'Command timeout (s)' 580 95), $numCmdTimeout,
    (New-Label 'App name prefix' 785 95), $txtApp
))

# ---- Multi-connection group
$grpMulti = New-Object System.Windows.Forms.GroupBox
$grpMulti.Text = 'Multiple connections'; $grpMulti.Dock = 'Top'; $grpMulti.Height = 85

$numCount = New-Numeric 140 24 70 1 5000 10
$numDelay = New-Numeric 395 24 70 0 60000 0
$chkParallel = New-Check 'Open in parallel' 485 26 $false
$btnTest     = New-Button 'Test Connection' 15 52 120
$btnOpen     = New-Button 'Open N Connections' 145 52 150
$btnCloseSel = New-Button 'Close Selected' 305 52 110
$btnCloseAll = New-Button 'Close All' 425 52 90
$btnCancel   = New-Button 'Cancel' 525 52 80
$btnCancel.Enabled = $false
$lblCount = New-Label 'Open: 0  |  Total listed: 0' 640 58
$lblCount.Font = New-Object System.Drawing.Font('Aptos', 9, [System.Drawing.FontStyle]::Bold)

$grpMulti.Controls.AddRange(@(
    (New-Label 'Number of connections' 5 25), $numCount,
    (New-Label 'Delay between opens (ms)' 235 27), $numDelay,
    $chkParallel, $btnTest, $btnOpen, $btnCloseSel, $btnCloseAll, $btnCancel, $lblCount
))

# ---- Status bar
$status = New-Object System.Windows.Forms.StatusStrip
$lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$lblStatus.Text = 'Ready.'
[void]$status.Items.Add($lblStatus)

# ---- Split: connection list / tabs
$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock = 'Fill'; $split.Orientation = 'Horizontal'; $split.SplitterDistance = 220

$lv = New-Object System.Windows.Forms.ListView
$lv.Dock = 'Fill'; $lv.View = 'Details'; $lv.FullRowSelect = $true; $lv.GridLines = $true; $lv.MultiSelect = $true
[void]$lv.Columns.Add('#', 45)
[void]$lv.Columns.Add('Application Name', 160)
[void]$lv.Columns.Add('Status', 80)
[void]$lv.Columns.Add('SPID', 60)
[void]$lv.Columns.Add('Open (ms)', 80)
[void]$lv.Columns.Add('Last query', 140)
[void]$lv.Columns.Add('Message', 520)
$split.Panel1.Controls.Add($lv)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$tabQuery   = New-Object System.Windows.Forms.TabPage; $tabQuery.Text   = 'Query'
$tabResults = New-Object System.Windows.Forms.TabPage; $tabResults.Text = 'Results'
$tabLog     = New-Object System.Windows.Forms.TabPage; $tabLog.Text     = 'Log'
$tabs.TabPages.AddRange(@($tabQuery, $tabResults, $tabLog))
$split.Panel2.Controls.Add($tabs)

# Query tab
$txtQuery = New-Object System.Windows.Forms.TextBox
$txtQuery.Multiline = $true; $txtQuery.ScrollBars = 'Both'; $txtQuery.WordWrap = $false
$txtQuery.Dock = 'Fill'; $txtQuery.Font = New-Object System.Drawing.Font('Consolas', 10)
$txtQuery.Text = "SELECT @@SERVERNAME AS ServerName, @@SPID AS SPID, SUSER_SNAME() AS LoginName, SYSDATETIME() AS ServerTime;"
$qPanel = New-Object System.Windows.Forms.Panel
$qPanel.Dock = 'Bottom'; $qPanel.Height = 40
$btnRunSel = New-Button 'Run on Selected (or first open)' 8 6 210
$btnRunAll = New-Button 'Run on ALL Open Connections' 228 6 210
$btnExport = New-Button 'Export Results CSV' 448 6 150
$qPanel.Controls.AddRange(@($btnRunSel, $btnRunAll, $btnExport))
$tabQuery.Controls.AddRange(@($txtQuery, $qPanel))

# Results tab
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Fill'; $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
$grid.AutoSizeColumnsMode = 'AllCells'; $grid.BackgroundColor = [System.Drawing.Color]::White
$tabResults.Controls.Add($grid)

# Log tab
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true; $txtLog.ReadOnly = $true; $txtLog.ScrollBars = 'Both'
$txtLog.Dock = 'Fill'; $txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$tabLog.Controls.Add($txtLog)

# Fill first, then Top controls, then status strip (docking order)
$form.Controls.AddRange(@($split, $grpMulti, $grpConn, $status))

# ------------------------------------------------------------------ events
$btnTest.Add_Click({ Test-OneConnection })
$btnOpen.Add_Click({ Start-OpenConnections })
$btnCancel.Add_Click({ $script:Cancel = $true; Write-Log 'Cancel requested.' })
$btnCloseSel.Add_Click({ Close-Entries (Get-SelectedEntries); $lblStatus.Text = 'Closed selected connection(s).' })
$btnCloseAll.Add_Click({
    $n = $script:Conns.Count
    Close-Entries @($script:Conns)
    Write-Log "Closed all connections ($n)."
    $lblStatus.Text = 'All connections closed.'
})
$btnRunSel.Add_Click({ Invoke-Query 'Selected' })
$btnRunAll.Add_Click({ Invoke-Query 'All' })
$btnExport.Add_Click({ Export-Results })
$form.Add_FormClosing({
    foreach ($e in @($script:Conns)) { try { $e.Conn.Dispose() } catch { } }
})

Write-Log 'Ready. Enter server details, then Test Connection or Open N Connections.'
[void]$form.ShowDialog()
