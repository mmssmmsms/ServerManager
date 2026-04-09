Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class DwmHelper {
    [DllImport("dwmapi.dll")]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    public static void SetDarkTitleBar(IntPtr handle) {
        int val = 1;
        DwmSetWindowAttribute(handle, 20, ref val, sizeof(int));
        DwmSetWindowAttribute(handle, 19, ref val, sizeof(int));
    }
}
"@

# Colors
$script:cBgDark      = [System.Drawing.Color]::FromArgb(18, 18, 24)
$script:cBgPanel     = [System.Drawing.Color]::FromArgb(28, 28, 36)
$script:cBgControl   = [System.Drawing.Color]::FromArgb(38, 38, 50)
$script:cBgHover     = [System.Drawing.Color]::FromArgb(50, 50, 65)
$script:cAccent      = [System.Drawing.Color]::FromArgb(99, 102, 241)
$script:cAccentHover = [System.Drawing.Color]::FromArgb(129, 132, 255)
$script:cSuccess     = [System.Drawing.Color]::FromArgb(52, 211, 153)
$script:cWarning     = [System.Drawing.Color]::FromArgb(251, 191, 36)
$script:cDanger      = [System.Drawing.Color]::FromArgb(248, 113, 113)
$script:cTextPrimary = [System.Drawing.Color]::FromArgb(240, 240, 255)
$script:cTextMuted   = [System.Drawing.Color]::FromArgb(140, 140, 170)
$script:cTextLabel   = [System.Drawing.Color]::FromArgb(99, 102, 241)
$script:cBorder      = [System.Drawing.Color]::FromArgb(55, 55, 75)
$script:cGreen       = [System.Drawing.Color]::FromArgb(30, 80, 50)
$script:cGreenHover  = [System.Drawing.Color]::FromArgb(40, 110, 70)
$script:cRed         = [System.Drawing.Color]::FromArgb(80, 30, 30)
$script:cRedHover    = [System.Drawing.Color]::FromArgb(110, 40, 40)
$script:cActBtn      = [System.Drawing.Color]::FromArgb(45, 45, 65)

# Fonts
$script:fUI    = [System.Drawing.Font]::new("Segoe UI", 9)
$script:fBold  = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$script:fMono  = [System.Drawing.Font]::new("Consolas", 9)
$script:fTitle = [System.Drawing.Font]::new("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$script:fSmall = [System.Drawing.Font]::new("Segoe UI", 8)

# State
$script:Results = [System.Collections.Generic.List[PSObject]]::new()
$script:AltCred = $null

# Styled button helper
function New-Btn {
    param(
        [string]$Text,
        [int]$X, [int]$Y, [int]$W, [int]$H,
        [System.Drawing.Color]$Bg,
        [System.Drawing.Color]$Hover,
        [System.Drawing.Font]$Font = $script:fUI
    )
    $b = [System.Windows.Forms.Button]::new()
    $b.Text      = $Text
    $b.Location  = [System.Drawing.Point]::new($X, $Y)
    $b.Size      = [System.Drawing.Size]::new($W, $H)
    $b.Font      = $Font
    $b.ForeColor = $script:cTextPrimary
    $b.BackColor = $Bg
    $b.FlatStyle = "Flat"
    $b.FlatAppearance.BorderSize         = 1
    $b.FlatAppearance.BorderColor        = $script:cBorder
    $b.FlatAppearance.MouseOverBackColor = $Hover
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    return $b
}

# Output helper - uses $script: controls
function AppendLine {
    param([string]$Text, [System.Drawing.Color]$Col)
    if (-not $Col) { $Col = $script:cTextPrimary }
    $script:txtOutput.SelectionStart  = $script:txtOutput.TextLength
    $script:txtOutput.SelectionLength = 0
    $script:txtOutput.SelectionColor  = $Col
    $script:txtOutput.AppendText("$Text`n")
    $script:txtOutput.ScrollToCaret()
}

# Remote execution helper
function Run-OnServer {
    param([string]$Server, [scriptblock]$SB)
    $isLocal = ($Server -eq $env:COMPUTERNAME -or
                $Server -eq "localhost" -or
                $Server -eq "127.0.0.1")
    if ($isLocal) { return (& $SB) }
    $p = @{ ComputerName=$Server; ScriptBlock=$SB; ErrorAction="Stop" }
    if ($script:rdoAlt.Checked -and $script:AltCred) {
        $p["Credential"] = $script:AltCred
    }
    return Invoke-Command @p
}

# Main form
$script:form               = [System.Windows.Forms.Form]::new()
$script:form.Text          = "Server Command Center"
$script:form.Size          = [System.Drawing.Size]::new(1160, 780)
$script:form.MinimumSize   = [System.Drawing.Size]::new(900, 600)
$script:form.StartPosition = "CenterScreen"
$script:form.BackColor     = $script:cBgDark
$script:form.ForeColor     = $script:cTextPrimary
$script:form.Font          = $script:fUI
$script:form.Add_Shown({ [DwmHelper]::SetDarkTitleBar($script:form.Handle) })

# Title bar
$script:pnlTitle           = [System.Windows.Forms.Panel]::new()
$script:pnlTitle.Dock      = "Top"
$script:pnlTitle.Height    = 48
$script:pnlTitle.BackColor = $script:cBgPanel

$script:lblTitle           = [System.Windows.Forms.Label]::new()
$script:lblTitle.Text      = "Server Command Center"
$script:lblTitle.Font      = $script:fTitle
$script:lblTitle.ForeColor = $script:cTextPrimary
$script:lblTitle.AutoSize  = $true
$script:lblTitle.Location  = [System.Drawing.Point]::new(14, 13)

$script:lblUser            = [System.Windows.Forms.Label]::new()
$script:lblUser.Text       = "$env:USERDOMAIN\$env:USERNAME"
$script:lblUser.Font       = $script:fSmall
$script:lblUser.ForeColor  = $script:cTextMuted
$script:lblUser.AutoSize   = $true
$script:lblUser.Location   = [System.Drawing.Point]::new(700, 16)

$script:pnlTitle.Controls.AddRange(@($script:lblTitle, $script:lblUser))

# Left panel
$script:pnlLeft            = [System.Windows.Forms.Panel]::new()
$script:pnlLeft.Size       = [System.Drawing.Size]::new(210, 690)
$script:pnlLeft.Location   = [System.Drawing.Point]::new(8, 8)
$script:pnlLeft.BackColor  = $script:cBgPanel

$script:lblSrv             = [System.Windows.Forms.Label]::new()
$script:lblSrv.Text        = "SERVERS"
$script:lblSrv.Font        = $script:fBold
$script:lblSrv.ForeColor   = $script:cTextLabel
$script:lblSrv.Location    = [System.Drawing.Point]::new(10, 12)
$script:lblSrv.AutoSize    = $true

$script:sep1               = [System.Windows.Forms.Panel]::new()
$script:sep1.Size          = [System.Drawing.Size]::new(190, 1)
$script:sep1.Location      = [System.Drawing.Point]::new(10, 32)
$script:sep1.BackColor     = $script:cBorder

$script:txtServer          = [System.Windows.Forms.TextBox]::new()
$script:txtServer.Size     = [System.Drawing.Size]::new(140, 26)
$script:txtServer.Location = [System.Drawing.Point]::new(10, 42)
$script:txtServer.BackColor   = $script:cBgControl
$script:txtServer.ForeColor   = $script:cTextMuted
$script:txtServer.BorderStyle = "FixedSingle"
$script:txtServer.Text        = "hostname or IP"
$script:txtServer.Font        = $script:fUI

$script:txtServer.Add_Enter({
    if ($script:txtServer.Text -eq "hostname or IP") {
        $script:txtServer.Text      = ""
        $script:txtServer.ForeColor = $script:cTextPrimary
    }
})
$script:txtServer.Add_Leave({
    if ($script:txtServer.Text.Trim() -eq "") {
        $script:txtServer.Text      = "hostname or IP"
        $script:txtServer.ForeColor = $script:cTextMuted
    }
})

$script:btnAdd = New-Btn -Text "Add" -X 155 -Y 41 -W 45 -H 26 `
    -Bg $script:cAccent -Hover $script:cAccentHover -Font $script:fBold

$script:btnLoad = New-Btn -Text "Load from File" -X 10 -Y 74 -W 190 -H 28 `
    -Bg $script:cBgControl -Hover $script:cBgHover

$script:clbServers              = [System.Windows.Forms.CheckedListBox]::new()
$script:clbServers.Location     = [System.Drawing.Point]::new(10, 110)
$script:clbServers.Size         = [System.Drawing.Size]::new(190, 370)
$script:clbServers.BackColor    = $script:cBgControl
$script:clbServers.ForeColor    = $script:cTextPrimary
$script:clbServers.BorderStyle  = "FixedSingle"
$script:clbServers.CheckOnClick = $true
$script:clbServers.Font         = $script:fUI

$script:btnSelAll = New-Btn -Text "All"  -X 10  -Y 488 -W 92  -H 26 `
    -Bg $script:cBgControl -Hover $script:cBgHover
$script:btnNone   = New-Btn -Text "None" -X 108 -Y 488 -W 92  -H 26 `
    -Bg $script:cBgControl -Hover $script:cBgHover
$script:btnRemove = New-Btn -Text "Remove Checked" -X 10 -Y 520 -W 190 -H 28 `
    -Bg $script:cRed -Hover $script:cRedHover

$script:sep2           = [System.Windows.Forms.Panel]::new()
$script:sep2.Size      = [System.Drawing.Size]::new(190, 1)
$script:sep2.Location  = [System.Drawing.Point]::new(10, 558)
$script:sep2.BackColor = $script:cBorder

$script:lblCred           = [System.Windows.Forms.Label]::new()
$script:lblCred.Text      = "CREDENTIALS"
$script:lblCred.Font      = $script:fBold
$script:lblCred.ForeColor = $script:cTextLabel
$script:lblCred.Location  = [System.Drawing.Point]::new(10, 565)
$script:lblCred.AutoSize  = $true

$script:rdoCurrent           = [System.Windows.Forms.RadioButton]::new()
$script:rdoCurrent.Text      = "Current account"
$script:rdoCurrent.Location  = [System.Drawing.Point]::new(10, 585)
$script:rdoCurrent.Size      = [System.Drawing.Size]::new(190, 20)
$script:rdoCurrent.Checked   = $true
$script:rdoCurrent.ForeColor = $script:cTextPrimary
$script:rdoCurrent.BackColor = $script:cBgPanel

$script:rdoAlt               = [System.Windows.Forms.RadioButton]::new()
$script:rdoAlt.Text          = "Alternate credentials"
$script:rdoAlt.Location      = [System.Drawing.Point]::new(10, 607)
$script:rdoAlt.Size          = [System.Drawing.Size]::new(190, 20)
$script:rdoAlt.ForeColor     = $script:cTextPrimary
$script:rdoAlt.BackColor     = $script:cBgPanel

$script:btnSetCred         = New-Btn -Text "Set Credentials" -X 10 -Y 630 -W 190 -H 26 `
    -Bg $script:cBgControl -Hover $script:cBgHover
$script:btnSetCred.Enabled = $false

$script:pnlLeft.Controls.AddRange(@(
    $script:lblSrv, $script:sep1,
    $script:txtServer, $script:btnAdd, $script:btnLoad,
    $script:clbServers, $script:btnSelAll, $script:btnNone, $script:btnRemove,
    $script:sep2, $script:lblCred,
    $script:rdoCurrent, $script:rdoAlt, $script:btnSetCred
))

# Center panel
$script:pnlCenter           = [System.Windows.Forms.Panel]::new()
$script:pnlCenter.Size      = [System.Drawing.Size]::new(210, 690)
$script:pnlCenter.Location  = [System.Drawing.Point]::new(226, 8)
$script:pnlCenter.BackColor = $script:cBgPanel

$script:lblAct              = [System.Windows.Forms.Label]::new()
$script:lblAct.Text         = "ACTIONS"
$script:lblAct.Font         = $script:fBold
$script:lblAct.ForeColor    = $script:cTextLabel
$script:lblAct.Location     = [System.Drawing.Point]::new(10, 12)
$script:lblAct.AutoSize     = $true

$script:sep3                = [System.Windows.Forms.Panel]::new()
$script:sep3.Size           = [System.Drawing.Size]::new(190, 1)
$script:sep3.Location       = [System.Drawing.Point]::new(10, 32)
$script:sep3.BackColor      = $script:cBorder

$script:pnlCenter.Controls.AddRange(@($script:lblAct, $script:sep3))

# About button at bottom of Actions panel
$script:btnAbout = New-Btn -Text "About" -X 10 -Y 640 -W 190 -H 30 `
    -Bg $script:cBgControl -Hover $script:cBgHover
$script:pnlCenter.Controls.Add($script:btnAbout)

# Right panel
$script:pnlRight            = [System.Windows.Forms.Panel]::new()
$script:pnlRight.Size       = [System.Drawing.Size]::new(690, 690)
$script:pnlRight.Location   = [System.Drawing.Point]::new(444, 8)
$script:pnlRight.BackColor  = $script:cBgPanel

$script:lblOut              = [System.Windows.Forms.Label]::new()
$script:lblOut.Text         = "OUTPUT"
$script:lblOut.Font         = $script:fBold
$script:lblOut.ForeColor    = $script:cTextLabel
$script:lblOut.Location     = [System.Drawing.Point]::new(10, 12)
$script:lblOut.AutoSize     = $true

$script:sep4                = [System.Windows.Forms.Panel]::new()
$script:sep4.Size           = [System.Drawing.Size]::new(670, 1)
$script:sep4.Location       = [System.Drawing.Point]::new(10, 32)
$script:sep4.BackColor      = $script:cBorder

$script:txtOutput           = [System.Windows.Forms.RichTextBox]::new()
$script:txtOutput.Location  = [System.Drawing.Point]::new(10, 40)
$script:txtOutput.Size      = [System.Drawing.Size]::new(670, 590)
$script:txtOutput.BackColor = $script:cBgDark
$script:txtOutput.ForeColor = $script:cTextPrimary
$script:txtOutput.Font      = $script:fMono
$script:txtOutput.ReadOnly  = $true
$script:txtOutput.ScrollBars   = "Vertical"
$script:txtOutput.BorderStyle  = "None"

$script:btnExportCSV = New-Btn -Text "Export CSV"   -X 10  -Y 638 -W 120 -H 32 `
    -Bg $script:cGreen -Hover $script:cGreenHover
$script:btnExportXL  = New-Btn -Text "Export Excel" -X 138 -Y 638 -W 130 -H 32 `
    -Bg $script:cGreen -Hover $script:cGreenHover
$script:btnClearOut  = New-Btn -Text "Clear Output" -X 276 -Y 638 -W 120 -H 32 `
    -Bg $script:cBgControl -Hover $script:cBgHover
$script:btnCopyOut   = New-Btn -Text "Copy All"     -X 404 -Y 638 -W 120 -H 32 `
    -Bg $script:cBgControl -Hover $script:cBgHover

$script:pnlRight.Controls.AddRange(@(
    $script:lblOut, $script:sep4, $script:txtOutput,
    $script:btnExportCSV, $script:btnExportXL,
    $script:btnClearOut, $script:btnCopyOut
))

# Status bar
$script:statusBar             = [System.Windows.Forms.StatusStrip]::new()
$script:statusBar.BackColor   = $script:cBgPanel

$script:statusLabel           = [System.Windows.Forms.ToolStripStatusLabel]::new()
$script:statusLabel.Text      = "Ready"
$script:statusLabel.ForeColor = $script:cTextMuted
$script:statusLabel.Font      = $script:fSmall

$script:statusSep             = [System.Windows.Forms.ToolStripSeparator]::new()

$script:statusRight           = [System.Windows.Forms.ToolStripStatusLabel]::new()
$script:statusRight.Text      = ""
$script:statusRight.ForeColor = $script:cSuccess
$script:statusRight.Font      = $script:fSmall
$script:statusRight.Spring    = $true
$script:statusRight.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight

$script:statusBar.Items.AddRange(@(
    $script:statusLabel, $script:statusSep, $script:statusRight
))

# Action button factory with rich Tag for dynamic discovery
function New-ActionBtn {
    param([string]$Label, [int]$Y, [scriptblock]$Action)

    $b = New-Btn -Text $Label -X 10 -Y $Y -W 190 -H 36 `
        -Bg $script:cActBtn -Hover $script:cAccent
    $b.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $b.Padding   = [System.Windows.Forms.Padding]::new(8, 0, 0, 0)

    # Keep ActionMap in case you still want it
    $safeKey = "Act_" + ($Label -replace '\W','_')
    $script:ActionMap = if ($script:ActionMap) { $script:ActionMap } else { @{} }
    $script:ActionMap[$safeKey]           = $Action
    $script:ActionMap[$safeKey + "_Label"] = $Label

    # Tag for dynamic discovery by Get Many
    $b.Tag = [PSCustomObject]@{
        Type   = "Action"
        Key    = $safeKey
        Label  = $Label
        Action = $Action
    }

    $b.Add_Click({
        $meta   = $this.Tag
        $key    = $meta.Key
        $ca     = $meta.Action
        $cl     = $meta.Label
        $servers = @($script:clbServers.CheckedItems)

        if ($servers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Check at least one server.",
                "No Server Selected",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $script:txtOutput.Clear()
        $script:Results.Clear()

        $ts      = (Get-Date).ToString("HH:mm:ss")
        $cm      = if ($script:rdoCurrent.Checked) { "current account" } else { "alt credentials" }
        $svrList = $servers -join ", "

        $script:txtOutput.SelectionStart  = $script:txtOutput.TextLength
        $script:txtOutput.SelectionLength = 0
        $script:txtOutput.SelectionColor  = $script:cTextMuted
        $script:txtOutput.AppendText("[$ts]  $cl  |  $svrList  |  $cm`n")

        $script:txtOutput.SelectionStart  = $script:txtOutput.TextLength
        $script:txtOutput.SelectionLength = 0
        $script:txtOutput.SelectionColor  = $script:cBorder
        $script:txtOutput.AppendText(("-" * 80) + "`n")

        foreach ($sv in $servers) {
            $script:statusLabel.Text      = "Querying $sv ..."
            $script:statusLabel.ForeColor = $script:cWarning

            $script:txtOutput.SelectionStart  = $script:txtOutput.TextLength
            $script:txtOutput.SelectionLength = 0
            $script:txtOutput.SelectionColor  = $script:cAccent
            $script:txtOutput.AppendText("  >> $sv`n")
            $script:txtOutput.ScrollToCaret()

            try {
                $r = & $ca -ServerName $sv
                if ($r) {
                    foreach ($item in $r) {
                        if ($item -is [PSCustomObject]) {
                            $script:Results.Add($item)
                        }
                    }
                }
            }
            catch {
                $em = $_.ToString()
                $script:txtOutput.SelectionStart  = $script:txtOutput.TextLength
                $script:txtOutput.SelectionLength = 0
                $script:txtOutput.SelectionColor  = $script:cDanger
                $script:txtOutput.AppendText("   ERROR: $em`n")
                $script:txtOutput.ScrollToCaret()
                $script:Results.Add([PSCustomObject]@{ Server=$sv; Error=$em })
            }
        }

        $script:txtOutput.SelectionStart  = $script:txtOutput.TextLength
        $script:txtOutput.SelectionLength = 0
        $script:txtOutput.SelectionColor  = $script:cBorder
        $script:txtOutput.AppendText(("-" * 80) + "`n")

        $rc = $script:Results.Count
        $sc = $servers.Count

        $script:txtOutput.SelectionStart  = $script:txtOutput.TextLength
        $script:txtOutput.SelectionLength = 0
        $script:txtOutput.SelectionColor  = $script:cSuccess
        $script:txtOutput.AppendText("Done  |  $sc server(s)  |  $rc row(s)`n")
        $script:txtOutput.ScrollToCaret()

        $script:statusLabel.Text      = "Done - $rc result(s)"
        $script:statusLabel.ForeColor = $script:cSuccess
        $script:statusRight.Text      = "Last run: $((Get-Date).ToString('HH:mm:ss'))"
    })

    return $b
}

# Actions
$actions = @(
    @{
        Label = "Ping / Connectivity"
        Action = {
            param($ServerName)
            $ping = Test-Connection -ComputerName $ServerName -Count 2 -ErrorAction SilentlyContinue
            if ($ping) {
                $avg = [math]::Round(($ping | Measure-Object -Property ResponseTime -Average).Average, 1)
                AppendLine "   Online   RTT avg: $avg ms" $script:cSuccess
                return [PSCustomObject]@{ Server=$ServerName; Status="Online"; AvgRTT_ms=$avg }
            } else {
                AppendLine "   Offline / unreachable" $script:cDanger
                return [PSCustomObject]@{ Server=$ServerName; Status="Offline"; AvgRTT_ms="N/A" }
            }
        }
    },
    @{
        Label = "OS and Build Info"
        Action = {
            param($ServerName)
            $os = Run-OnServer $ServerName { Get-CimInstance Win32_OperatingSystem }
            AppendLine "   OS:    $($os.Caption)" $script:cTextPrimary
            AppendLine "   Build: $($os.BuildNumber)   Last Boot: $($os.LastBootUpTime)" $script:cTextMuted
            return [PSCustomObject]@{
                Server=$ServerName; OS=$os.Caption
                Build=$os.BuildNumber; LastBoot=$os.LastBootUpTime
            }
        }
    },
    @{
        Label = "Disk Space"
        Action = {
            param($ServerName)
            $disks = Run-OnServer $ServerName {
                Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3"
            }
            $out = foreach ($d in $disks) {
                $free  = [math]::Round($d.FreeSpace / 1GB, 1)
                $total = [math]::Round($d.Size / 1GB, 1)
                $pct   = [math]::Round(($free / $total) * 100, 0)
                $bar   = ("X" * [math]::Floor($pct / 10)).PadRight(10, "-")
                $col   = if ($pct -lt 15) { $script:cDanger } elseif ($pct -lt 30) { $script:cWarning } else { $script:cSuccess }
                AppendLine "   $($d.DeviceID)  [$bar] $pct% free   $free GB / $total GB" $col
                [PSCustomObject]@{
                    Server=$ServerName; Drive=$d.DeviceID
                    FreeGB=$free; TotalGB=$total; FreePct=$pct
                }
            }
            return $out
        }
    },
    @{
        Label = "CPU and Memory"
        Action = {
            param($ServerName)
            $cpu   = Run-OnServer $ServerName {
                (Get-CimInstance Win32_Processor | Measure-Object LoadPercentage -Average).Average
            }
            $cs    = Run-OnServer $ServerName { Get-CimInstance Win32_ComputerSystem }
            $os    = Run-OnServer $ServerName { Get-CimInstance Win32_OperatingSystem }
            $total = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
            $free  = [math]::Round($os.FreePhysicalMemory / 1MB, 1)
            $used  = [math]::Round($total - $free, 1)
            $cpuCol = if ($cpu -gt 85) { $script:cDanger } elseif ($cpu -gt 60) { $script:cWarning } else { $script:cSuccess }
            AppendLine "   CPU  : $cpu%" $cpuCol
            AppendLine "   RAM  : $used GB used / $total GB total" $script:cTextPrimary
            return [PSCustomObject]@{
                Server=$ServerName; CPU_Pct=$cpu
                UsedRAM_GB=$used; TotalRAM_GB=$total
            }
        }
    },
    @{
        Label = "Uptime"
        Action = {
            param($ServerName)
            $os = Run-OnServer $ServerName { Get-CimInstance Win32_OperatingSystem }
            $up = (Get-Date) - $os.LastBootUpTime
            AppendLine "   Uptime: $($up.Days)d $($up.Hours)h $($up.Minutes)m" $script:cTextPrimary
            return [PSCustomObject]@{
                Server=$ServerName; Days=$up.Days
                Hours=$up.Hours; Minutes=$up.Minutes
            }
        }
    },
    @{
        Label = "Running Services"
        Action = {
            param($ServerName)
            $svcs = Run-OnServer $ServerName {
                Get-Service | Where-Object { $_.Status -eq "Running" } | Sort-Object DisplayName
            }
            AppendLine "   $($svcs.Count) services running" $script:cTextMuted
            foreach ($s in $svcs) { AppendLine "   - $($s.DisplayName)" $script:cTextPrimary }
            return $svcs | ForEach-Object {
                [PSCustomObject]@{ Server=$ServerName; Service=$_.DisplayName; Status=$_.Status }
            }
        }
    },
    @{
        Label = "WinRM Status"
        Action = {
            param($ServerName)
            $ok  = Test-WSMan -ComputerName $ServerName -ErrorAction SilentlyContinue
            $col = if ($ok) { $script:cSuccess } else { $script:cDanger }
            $msg = if ($ok) { "WinRM OK  (Version: $($ok.ProductVersion))" } else { "WinRM NOT responding" }
            AppendLine "   $msg" $col
            return [PSCustomObject]@{
                Server=$ServerName
                WinRM=if ($ok) { "OK" } else { "FAIL" }
                Version=$ok.ProductVersion
            }
        }
    },
    @{
        Label = "Recent Hotfixes"
        Action = {
            param($ServerName)
            $hf = Run-OnServer $ServerName {
                Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 10
            }
            foreach ($h in $hf) {
                AppendLine "   $($h.HotFixID)   $($h.InstalledOn)   $($h.Description)" $script:cTextPrimary
            }
            return $hf | ForEach-Object {
                [PSCustomObject]@{
                    Server=$ServerName; HotFixID=$_.HotFixID
                    Installed=$_.InstalledOn; Type=$_.Description
                }
            }
        }
    },
    @{
        Label = "Hardware Info"
        Action = {
            param($ServerName)
            $cs  = Run-OnServer $ServerName { Get-CimInstance Win32_ComputerSystem }
            $bb  = Run-OnServer $ServerName { Get-CimInstance Win32_BaseBoard }
            $gpu = Run-OnServer $ServerName { Get-CimInstance Win32_VideoController }

            $make  = $cs.Manufacturer
            $model = $cs.Model
            $board = "$($bb.Manufacturer) $($bb.Product)"

            AppendLine "   Manufacturer : $make" $script:cTextPrimary
            AppendLine "   Model        : $model" $script:cTextPrimary
            AppendLine "   Baseboard    : $board" $script:cTextMuted

            $results = [System.Collections.Generic.List[PSObject]]::new()
            $results.Add([PSCustomObject]@{
                Server       = $ServerName
                Manufacturer = $make
                Model        = $model
                Baseboard    = $board
            })

            foreach ($g in $gpu) {
                AppendLine "   GPU          : $($g.Name)  [$($g.AdapterRAM / 1MB) MB]" $script:cTextPrimary
                $results.Add([PSCustomObject]@{
                    Server       = $ServerName
                    Manufacturer = "GPU"
                    Model        = $g.Name
                    Baseboard    = "$([math]::Round($g.AdapterRAM / 1MB)) MB VRAM"
                })
            }

            return $results
        }
    },
    @{
        Label  = "Get Network"
        Action = {
            param($ServerName)

            $results = [System.Collections.Generic.List[PSObject]]::new()

            # IP and NIC information
            $ipConfig = Run-OnServer $ServerName {
                Get-NetIPConfiguration | Where-Object { $_.NetAdapter.Status -eq "Up" }
            }

            foreach ($cfg in $ipConfig) {
                $nic    = $cfg.NetAdapter
                $ifName = $nic.InterfaceAlias
                $ifDesc = $nic.InterfaceDescription
                $mac    = $nic.MacAddress

                # LinkSpeed is usually a string like "50 Gbps" or "1 Gbps"
                $speedStr = "$($nic.LinkSpeed)"
                $speedGb  = $null
                if ($speedStr -match '(\d+(\.\d+)?)') {
                    $speedGb = [double]$Matches[1]
                }

                AppendLine "   NIC: $ifName  ($ifDesc)" $script:cTextPrimary
                AppendLine "        MAC:   $mac   Speed: $speedStr" $script:cTextMuted

                foreach ($addr in $cfg.IPv4Address) {
                    AppendLine "        IPv4:  $($addr.IPAddress)/$($addr.PrefixLength)" $script:cTextPrimary
                }
                foreach ($addr in $cfg.IPv6Address) {
                    AppendLine "        IPv6:  $($addr.IPAddress)/$($addr.PrefixLength)" $script:cTextMuted
                }
                foreach ($gw in $cfg.IPv4DefaultGateway) {
                    AppendLine "        GWv4:  $($gw.NextHop)" $script:cTextMuted
                }
                foreach ($gw in $cfg.IPv6DefaultGateway) {
                    AppendLine "        GWv6:  $($gw.NextHop)" $script:cTextMuted
                }
                AppendLine "" $script:cTextMuted

                # Summary row for this NIC
                $ipv4 = ($cfg.IPv4Address | Select-Object -First 1).IPAddress
                $gw4  = ($cfg.IPv4DefaultGateway | Select-Object -First 1).NextHop

                $results.Add([PSCustomObject]@{
                    Server   = $ServerName
                    Type     = "NIC"
                    Name     = $nic.InterfaceAlias
                    MAC      = $nic.MacAddress
                    SpeedGb  = $speedGb
                    SpeedRaw = $speedStr
                    IPv4     = $ipv4
                    GWv4     = $gw4
                    TeamName = $null
                })
            }

            # NIC Teaming (LBFO teams)
            $teams = Run-OnServer $ServerName {
                if (Get-Command Get-NetLbfoTeam -ErrorAction SilentlyContinue) {
                    Get-NetLbfoTeam
                }
            }

            if ($teams) {
                AppendLine "   NIC Teams:" $script:cTextPrimary
                foreach ($t in $teams) {
                    AppendLine "      Team: $($t.Name)  Status: $($t.Status)  Mode: $($t.TeamingMode)  LB: $($t.LoadBalancingAlgorithm)" $script:cTextPrimary
                    AppendLine "            Members: $($t.Members -join ', ')" $script:cTextMuted

                    $results.Add([PSCustomObject]@{
                        Server      = $ServerName
                        Type        = "Team"
                        Name        = $t.Name
                        Status      = $t.Status
                        Mode        = $t.TeamingMode
                        LBAlgorithm = $t.LoadBalancingAlgorithm
                        Members     = ($t.Members -join ", ")
                        SpeedGb     = $null
                        SpeedRaw    = $null
                        IPv4        = $null
                        GWv4        = $null
                    })
                }
            } else {
                AppendLine "   NIC Teams: none or NetLbfo not available" $script:cTextMuted
            }

            return $results
        }
    },
    @{
        Label  = "Run Remote"
        Action = {
            param($ServerName)

            # Build a small modal dialog to capture the script text
            $dlg        = New-Object System.Windows.Forms.Form
            $dlg.Text   = "Run Remote on $ServerName"
            $dlg.Size   = [System.Drawing.Size]::new(600, 400)
            $dlg.StartPosition = "CenterParent"
            $dlg.BackColor     = $script:cBgPanel
            $dlg.ForeColor     = $script:cTextPrimary
            $dlg.Font          = $script:fUI
            $dlg.TopMost       = $true

            $lblInfo           = New-Object System.Windows.Forms.Label
            $lblInfo.Text      = "Enter PowerShell to run on the remote server:"
            $lblInfo.AutoSize  = $true
            $lblInfo.Location  = [System.Drawing.Point]::new(10, 10)
            $lblInfo.ForeColor = $script:cTextPrimary

            $txtScript         = New-Object System.Windows.Forms.TextBox
            $txtScript.Multiline  = $true
            $txtScript.ScrollBars = "Vertical"
            $txtScript.Location   = [System.Drawing.Point]::new(10, 30)
            $txtScript.Size       = [System.Drawing.Size]::new(560, 280)
            $txtScript.BackColor  = $script:cBgDark
            $txtScript.ForeColor  = $script:cTextPrimary
            $txtScript.Font       = $script:fMono
            # Example template
            $txtScript.Text       = '# Example:' + "`r`n" +
                                    'Get-Service | Where-Object Status -eq "Running" | Select-Object -First 10'

            $btnOK              = New-Btn -Text "Run" -X 380 -Y 320 -W 90 -H 30 `
                                       -Bg $script:cAccent -Hover $script:cAccentHover
            $btnCancel          = New-Btn -Text "Cancel" -X 480 -Y 320 -W 90 -H 30 `
                                       -Bg $script:cBgControl -Hover $script:cBgHover

            $btnOK.Add_Click({
                if ([string]::IsNullOrWhiteSpace($txtScript.Text)) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Please enter a PowerShell command or script block.",
                        "No Script",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Warning
                    ) | Out-Null
                } else {
                    $dlg.Tag = $txtScript.Text
                    $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    $dlg.Close()
                }
            })
            $btnCancel.Add_Click({
                $dlg.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $dlg.Close()
            })

            $dlg.Controls.AddRange(@($lblInfo, $txtScript, $btnOK, $btnCancel))

            $result = $dlg.ShowDialog()
            if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
                AppendLine "   Remote run cancelled for $ServerName" $script:cTextMuted
                return
            }

            $code = $dlg.Tag
            AppendLine "   Running remote script on $ServerName ..." $script:cTextMuted

            # Run the entered code on the remote server
            $scriptBlock = [scriptblock]::Create($code)
            $output = Run-OnServer $ServerName $scriptBlock

            if ($output -eq $null) {
                AppendLine "   (no output)" $script:cTextMuted
                return
            }

            foreach ($line in $output) {
                AppendLine ("   " + ($line | Out-String).TrimEnd()) $script:cTextPrimary
                $script:Results.Add([PSCustomObject]@{
                    Server = $ServerName
                    Output = ($line | Out-String).TrimEnd()
                })
            }
        }
    },
        @{
        Label  = "Get Many"
        Action = {
            param($ServerName)

            # Discover available actions from buttons in the Actions panel at runtime
            $availableActions = @()
            foreach ($ctrl in $script:pnlCenter.Controls) {
                if ($ctrl -is [System.Windows.Forms.Button] -and $ctrl.Tag) {
                    $tag = $ctrl.Tag
                    if ($tag.Type -eq "Action" -and $tag.Label -ne "Get Many") {
                        $availableActions += $tag
                    }
                }
            }

            if ($availableActions.Count -eq 0) {
                AppendLine "   Get Many: no actions discovered in the UI." $script:cDanger
                return
            }

            # Build a lookup: label -> action metadata
            $actionLookup = @{}
            foreach ($act in $availableActions) {
                # In case of duplicate labels, last one wins (labels are unique in your UI)
                $actionLookup[$act.Label] = $act
            }

            # Build selection dialog
            $dlg        = New-Object System.Windows.Forms.Form
            $dlg.Text   = "Select Actions to Run"
            $dlg.Size   = [System.Drawing.Size]::new(400, 420)
            $dlg.StartPosition = "CenterParent"
            $dlg.BackColor     = $script:cBgPanel
            $dlg.ForeColor     = $script:cTextPrimary
            $dlg.Font          = $script:fUI
            $dlg.TopMost       = $true

            $lbl              = New-Object System.Windows.Forms.Label
            $lbl.Text         = "Select one or more actions to run on the selected servers:"
            $lbl.AutoSize     = $true
            $lbl.Location     = [System.Drawing.Point]::new(10, 10)
            $lbl.ForeColor    = $script:cTextPrimary

            $lst              = New-Object System.Windows.Forms.CheckedListBox
            $lst.Location     = [System.Drawing.Point]::new(10, 35)
            $lst.Size         = [System.Drawing.Size]::new(360, 300)
            $lst.BackColor    = $script:cBgControl
            $lst.ForeColor    = $script:cTextPrimary
            $lst.BorderStyle  = "FixedSingle"
            $lst.CheckOnClick = $true

            # Populate list with labels only; we will map back via $actionLookup
            foreach ($act in $availableActions) {
                [void]$lst.Items.Add($act.Label)
            }

            $btnOK     = New-Btn -Text "Run"    -X 200 -Y 340 -W 80 -H 30 -Bg $script:cAccent    -Hover $script:cAccentHover
            $btnCancel = New-Btn -Text "Cancel" -X 290 -Y 340 -W 80 -H 30 -Bg $script:cBgControl -Hover $script:cBgHover

            $btnOK.Add_Click({
                if ($lst.CheckedItems.Count -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Check at least one action.",
                        "No Actions Selected",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Warning
                    ) | Out-Null
                    return
                }
                $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $dlg.Close()
            })
            $btnCancel.Add_Click({
                $dlg.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $dlg.Close()
            })

            $dlg.Controls.AddRange(@($lbl, $lst, $btnOK, $btnCancel))

            $result = $dlg.ShowDialog()
            if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
                AppendLine "   Get Many cancelled for $ServerName" $script:cTextMuted
                return
            }

            # Map selected labels back to actions
            $selected = @()
            foreach ($label in $lst.CheckedItems) {
                if ($actionLookup.ContainsKey($label)) {
                    $selected += $actionLookup[$label]
                }
            }

            if ($selected.Count -eq 0) {
                AppendLine "   Get Many: no actions selected after dialog." $script:cTextMuted
                return
            }

            AppendLine "   Get Many: running $($selected.Count) actions on $ServerName..." $script:cTextMuted

            foreach ($act in $selected) {
                AppendLine "     -> $($act.Label)" $script:cTextPrimary
                try {
                    $res = & $act.Action -ServerName $ServerName
                    if ($res) {
                        foreach ($item in $res) {
                            if ($item -is [PSCustomObject]) {
                                $script:Results.Add($item)
                            }
                        }
                    }
                } catch {
                    AppendLine "       ERROR in '$($act.Label)': $_" $script:cDanger
                    $script:Results.Add([PSCustomObject]@{
                        Server = $ServerName
                        Action = $act.Label
                        Error  = $_.ToString()
                    })
                }
            }
        }
    }

        ,@{
        Label  = "DownAdmin On/Off"
        Action = {
            param($ServerName)

            # Ask what the user wants to do
            $choice = [System.Windows.Forms.MessageBox]::Show(
                "Do you want monitoring ON (suppress for 1 hour) or OFF (resume monitoring) for $ServerName?" + "`r`n`r`n" +
                "Yes = Turn monitoring ON (suppress for 1 hour)" + "`r`n" +
                "No  = Turn monitoring OFF (resume)" + "`r`n" +
                "Cancel = Do nothing",
                "DownAdmin On/Off",
                [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )

            if ($choice -eq [System.Windows.Forms.DialogResult]::Cancel) {
                AppendLine "   DownAdmin: cancelled for $ServerName" $script:cTextMuted
                return
            }

            if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
                # Monitoring ON (suppress for 1 hour)
                $cmd = "downadmin -m $ServerName -hours 1 -comment `"Turning Monitoring Off`""
                $mode = "ON (suppress 1h)"
            } else {
                # Monitoring OFF (resume)
                $cmd = "downadmin -m $ServerName -u"
                $mode = "OFF (resume)"
            }

            AppendLine "   DownAdmin: $mode for $ServerName" $script:cTextPrimary
            AppendLine "      Command: $cmd" $script:cTextMuted

            try {
                # Run downadmin locally, targeting the selected host
                $output = & downadmin -m $ServerName @(
                    if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) { "-hours"; "1"; "-comment"; "Turning Monitoring Off" }
                    else { "-u" }
                )

                if ($output) {
                    foreach ($line in $output) {
                        AppendLine ("      " + ($line | Out-String).TrimEnd()) $script:cTextPrimary
                        $script:Results.Add([PSCustomObject]@{
                            Server   = $ServerName
                            Action   = "DownAdmin"
                            Mode     = $mode
                            Output   = ($line | Out-String).TrimEnd()
                        })
                    }
                } else {
                    AppendLine "      (no output from downadmin)" $script:cTextMuted
                    $script:Results.Add([PSCustomObject]@{
                        Server = $ServerName
                        Action = "DownAdmin"
                        Mode   = $mode
                        Output = "(no output)"
                    })
                }
            }
            catch {
                AppendLine "      ERROR running downadmin: $_" $script:cDanger
                $script:Results.Add([PSCustomObject]@{
                    Server = $ServerName
                    Action = "DownAdmin"
                    Mode   = $mode
                    Error  = $_.ToString()
                })
            }
        }
    }

)

$yPos = 42
foreach ($a in $actions) {
    $ab = New-ActionBtn -Label $a.Label -Y $yPos -Action $a.Action
    $script:pnlCenter.Controls.Add($ab)
    $yPos += 44
}

# Event handlers
$script:btnAdd.Add_Click({
    $name = $script:txtServer.Text.Trim()
    if ($name -and $name -ne "hostname or IP" -and $script:clbServers.Items -notcontains $name) {
        $script:clbServers.Items.Add($name, $true) | Out-Null
        $script:txtServer.Text      = "hostname or IP"
        $script:txtServer.ForeColor = $script:cTextMuted
        $script:statusLabel.Text      = "Added: $name"
        $script:statusLabel.ForeColor = $script:cSuccess
        $script:statusRight.Text = "$($script:clbServers.Items.Count) server(s)"
    }
})

$script:btnAbout.Add_Click({
    $msg = "Property of Intel IT - Engineering Computing`r`n" +
           "Developed by Mirza Saracevic 2026"
    [System.Windows.Forms.MessageBox]::Show(
        $msg,
        "About Server Command Center",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
})


$script:txtServer.Add_KeyDown({
    if ($_.KeyCode -eq "Return") { $script:btnAdd.PerformClick() }
})

$script:btnLoad.Add_Click({
    $dlg = [System.Windows.Forms.OpenFileDialog]::new()
    $dlg.Filter = "Text/CSV (*.txt;*.csv)|*.txt;*.csv|All Files|*.*"
    $dlg.Title  = "Load Server List"
    if ($dlg.ShowDialog() -eq "OK") {
        $lines = Get-Content $dlg.FileName | Where-Object { $_.Trim() -ne "" }
        $added = 0
        foreach ($line in $lines) {
            $s = $line.Trim().Split(',')[0].Trim()
            if ($s -and $script:clbServers.Items -notcontains $s) {
                $script:clbServers.Items.Add($s, $true) | Out-Null
                $added++
            }
        }
        AppendLine "Loaded $added server(s) from: $($dlg.FileName)" $script:cTextMuted
        $script:statusLabel.Text      = "Loaded $added server(s)"
        $script:statusLabel.ForeColor = $script:cSuccess
        $script:statusRight.Text = "$($script:clbServers.Items.Count) server(s)"
    }
})

$script:btnSelAll.Add_Click({
    for ($i = 0; $i -lt $script:clbServers.Items.Count; $i++) {
        $script:clbServers.SetItemChecked($i, $true)
    }
})

$script:btnNone.Add_Click({
    for ($i = 0; $i -lt $script:clbServers.Items.Count; $i++) {
        $script:clbServers.SetItemChecked($i, $false)
    }
})

$script:btnRemove.Add_Click({
    @($script:clbServers.CheckedItems) | ForEach-Object {
        $script:clbServers.Items.Remove($_)
    }
    $script:statusRight.Text = "$($script:clbServers.Items.Count) server(s)"
})

$script:rdoAlt.Add_CheckedChanged({
    $script:btnSetCred.Enabled = $script:rdoAlt.Checked
})

$script:btnSetCred.Add_Click({
    $cred = Get-Credential -Message "Enter alternate credentials for remote connections"
    if ($cred) {
        $script:AltCred               = $cred
        $un = $cred.UserName.Split('\')[-1]
        $script:btnSetCred.Text       = "Cred: $un"
        $script:statusLabel.Text      = "Credentials set: $($cred.UserName)"
        $script:statusLabel.ForeColor = $script:cWarning
    }
})

$script:btnClearOut.Add_Click({
    $script:txtOutput.Clear()
    $script:Results.Clear()
    $script:statusLabel.Text      = "Output cleared"
    $script:statusLabel.ForeColor = $script:cTextMuted
})

$script:btnCopyOut.Add_Click({
    [System.Windows.Forms.Clipboard]::SetText($script:txtOutput.Text)
    $script:statusLabel.Text      = "Copied to clipboard"
    $script:statusLabel.ForeColor = $script:cSuccess
})

$script:btnExportCSV.Add_Click({
    if ($script:Results.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Run an action first.", "No Data")
        return
    }
    $dlg = [System.Windows.Forms.SaveFileDialog]::new()
    $dlg.Filter   = "CSV (*.csv)|*.csv"
    $dlg.FileName = "ServerResults_$(Get-Date -f yyyyMMdd_HHmmss).csv"
    if ($dlg.ShowDialog() -eq "OK") {
        $script:Results | Export-Csv -Path $dlg.FileName -NoTypeInformation
        $script:statusLabel.Text      = "Exported: $($dlg.FileName)"
        $script:statusLabel.ForeColor = $script:cSuccess
    }
})

$script:btnExportXL.Add_Click({
    if ($script:Results.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Run an action first.", "No Data")
        return
    }
    $dlg = [System.Windows.Forms.SaveFileDialog]::new()
    $dlg.Filter   = "Excel (*.xlsx)|*.xlsx"
    $dlg.FileName = "ServerResults_$(Get-Date -f yyyyMMdd_HHmmss).xlsx"
    if ($dlg.ShowDialog() -eq "OK") {
        if (Get-Command Export-Excel -ErrorAction SilentlyContinue) {
            $script:Results | Export-Excel -Path $dlg.FileName `
                -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
            $script:statusLabel.Text      = "Exported: $($dlg.FileName)"
            $script:statusLabel.ForeColor = $script:cSuccess
        } else {
            try {
                $xl = New-Object -ComObject Excel.Application
                $xl.Visible = $false
                $wb = $xl.Workbooks.Add()
                $ws = $wb.Worksheets.Item(1)
                $headers = $script:Results[0].PSObject.Properties.Name
                for ($c = 0; $c -lt $headers.Count; $c++) {
                    $ws.Cells.Item(1, $c + 1)           = $headers[$c]
                    $ws.Cells.Item(1, $c + 1).Font.Bold = $true
                }
                for ($r = 0; $r -lt $script:Results.Count; $r++) {
                    for ($c = 0; $c -lt $headers.Count; $c++) {
                        $ws.Cells.Item($r + 2, $c + 1) = $script:Results[$r].($headers[$c])
                    }
                }
                $ws.UsedRange.Columns.AutoFit() | Out-Null
                $wb.SaveAs($dlg.FileName, 51)
                $wb.Close($false)
                $xl.Quit()
                [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
                $script:statusLabel.Text      = "Exported: $($dlg.FileName)"
                $script:statusLabel.ForeColor = $script:cSuccess
            } catch {
                AppendLine "Excel error: $_" $script:cDanger
                $script:statusLabel.Text      = "Excel failed. Try: Install-Module ImportExcel"
                $script:statusLabel.ForeColor = $script:cDanger
            }
        }
    }
})

# Assemble and launch
$script:pnlMain           = [System.Windows.Forms.Panel]::new()
$script:pnlMain.Dock      = "Fill"
$script:pnlMain.BackColor = $script:cBgDark
$script:pnlMain.Controls.AddRange(@(
    $script:pnlLeft, $script:pnlCenter, $script:pnlRight
))

$script:form.Controls.AddRange(@(
    $script:pnlTitle, $script:pnlMain, $script:statusBar
))

$script:statusLabel.Text      = "Ready  -  $env:USERDOMAIN\$env:USERNAME"
$script:statusLabel.ForeColor = $script:cTextMuted
AppendLine "Server Command Center  |  $env:USERDOMAIN\$env:USERNAME  |  $env:COMPUTERNAME" $script:cTextMuted
AppendLine "Add servers on the left, check them, then click an action." $script:cTextMuted
AppendLine "" $script:cTextMuted

[System.Windows.Forms.Application]::Run($script:form)
