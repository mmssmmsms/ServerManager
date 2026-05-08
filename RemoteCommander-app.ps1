<#
.SYNOPSIS
    GUI-based multi-server administration tool for running common queries and actions against remote Windows servers.

.DESCRIPTION
    Launches a WinForms "Remote Commander" window with a dark-themed UI that lets you
    manage a list of servers and run predefined administrative actions against one or more of
    them in parallel. Results are displayed in a color-coded output pane and can be exported
    to CSV or Excel.

    Available actions:
      - Ping / Connectivity    - Test reachability and round-trip time
      - OS and Build Info      - OS caption, build number, and last boot time
      - Disk Space             - Logical disk free/total with visual bar
      - CPU and Memory         - Current CPU load and RAM usage
      - Uptime                 - Days, hours, minutes since last boot
      - Running Services       - List all running services
      - WinRM Status           - Verify WinRM connectivity and version
      - Recent Hotfixes        - Last 10 installed hotfixes
      - Hardware Info          - Manufacturer, model, baseboard, and GPU details
      - Get Network            - NIC configuration, IP addresses, and LBFO teaming
      - Run Remote             - Execute arbitrary PowerShell on the remote server
      - Multi-Action           - Run multiple actions in a single pass
      - Reboot                 - Graceful or forced restart of the remote server
      - DownAdmin              - Toggle SCOM monitoring suppression via downadmin

    Servers can be added manually by hostname/IP or loaded from a text/CSV file (one server
    per line). The tool supports both current-user and alternate credential authentication
    for remote connections via WinRM (Invoke-Command). Local server queries bypass remoting.

.EXAMPLE
    .\RemoteCommander-app.ps1
    Launches the Remote Commander GUI.

.NOTES
    Requires PowerShell 5.1 or later.
    Requires WinRM enabled on target servers for remote actions.
    The "Export Excel" button uses the ImportExcel module if available, otherwise falls back
    to Excel COM automation.
    If launched from PowerShell 7/Core (pwsh), the script automatically re-launches
    itself under Windows PowerShell (powershell.exe).
#>
#requires -version 5.1

#region Re-launch under Windows PowerShell if running in pwsh
if ($PSVersionTable.PSEdition -eq 'Core') {
    $wpExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    if (Test-Path $wpExe) {
        & $wpExe -NoProfile -ExecutionPolicy Bypass -STA -File $MyInvocation.MyCommand.Path
        return
    } else {
        Write-Warning "Windows PowerShell not found at $wpExe - this script requires Windows PowerShell."
        return
    }
}
#endregion Re-launch under Windows PowerShell if running in pwsh

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# DwmHelper - dark title bar support (may already exist from previous run in same session)
if (-not ('DwmHelper' -as [type])) {
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
}

# TextBoxHelper - native placeholder/watermark text for text boxes
if (-not ('TextBoxHelper' -as [type])) {
    Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class TextBoxHelper {
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);
        private const int EM_SETCUEBANNER = 0x1501;
        public static void SetPlaceholder(IntPtr handle, string text) {
            SendMessage(handle, EM_SETCUEBANNER, (IntPtr)1, text);
        }
    }
"@
}

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
$script:AppTitle = "Remote Commander"
$script:Results  = [System.Collections.Generic.List[PSObject]]::new()
$script:AltCred  = $null

#region Event Log helper
$script:EventLogSource = 'RemoteCommander'
# Register the event source if it does not already exist (requires elevation on first run)
try {
    if (-not [System.Diagnostics.EventLog]::SourceExists($script:EventLogSource)) {
        [System.Diagnostics.EventLog]::CreateEventSource($script:EventLogSource, 'Application')
    }
} catch {
    # Non-admin session - source may already exist from a prior elevated run
}

function Write-AuditLog {
    param(
        [string]$Message,
        [System.Diagnostics.EventLogEntryType]$EntryType = 'Information',
        [int]$EventId = 1000
    )
    # Event log messages are limited to ~31,839 characters.
    # If the message exceeds the safe limit, split it across multiple entries.
    $maxLen = 31000
    try {
        if ($Message.Length -le $maxLen) {
            Write-EventLog -LogName Application -Source $script:EventLogSource `
                -EventId $EventId -EntryType $EntryType -Message $Message
        } else {
            # Calculate number of parts needed
            $totalParts = [math]::Ceiling($Message.Length / $maxLen)
            for ($i = 0; $i -lt $totalParts; $i++) {
                $start  = $i * $maxLen
                $length = [math]::Min($maxLen, $Message.Length - $start)
                $chunk  = $Message.Substring($start, $length)
                $header = "[Part $($i + 1) of $totalParts]`n"
                Write-EventLog -LogName Application -Source $script:EventLogSource `
                    -EventId $EventId -EntryType $EntryType -Message ($header + $chunk)
            }
        }
    } catch {
        # Silently continue if event log write fails (source not registered yet)
    }
}
#endregion Event Log helper

#region Concurrent execution infrastructure
# Thread-safe output queue for runspace workers to communicate with the UI thread.
# Items are PSCustomObjects with Type (Line/Result/Error/Done), Server, Text, ColorName, Data.
$script:OutputQueue = [System.Collections.Concurrent.ConcurrentQueue[PSObject]]::new()

# Color name-to-Color mapping for runspace workers (they cannot access $script:c* variables)
$script:ColorMap = @{
    TextPrimary = $script:cTextPrimary
    TextMuted   = $script:cTextMuted
    Success     = $script:cSuccess
    Warning     = $script:cWarning
    Danger      = $script:cDanger
    Accent      = $script:cAccent
    Border      = $script:cBorder
}

# Tracking object for in-flight concurrent work
$script:ConcurrentRun = $null

function Start-ConcurrentActions {
    <#
    .SYNOPSIS
        Fans out an action scriptblock across multiple servers using a runspace pool.
    .DESCRIPTION
        Creates a RunspacePool, injects helper functions (Run-OnServer variant and
        an AppendLine shim that enqueues to the ConcurrentQueue), then launches one
        PowerShell pipeline per server. Returns a tracking object consumed by the
        UI timer to drain output and detect completion.
    #>
    param(
        [scriptblock]$Action,
        [string[]]$Servers,
        [bool]$UseAltCred,
        [PSCredential]$Credential,
        [System.Collections.Concurrent.ConcurrentQueue[PSObject]]$Queue,
        [int]$ThrottleLimit = 32,
        [hashtable]$Context = @{}
    )

    # Build initial session state with helper functions available inside each runspace
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

    # Create the runspace pool
    $pool = [runspacefactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
    $pool.Open()

    $pipelines = [System.Collections.Generic.List[PSObject]]::new()

    foreach ($server in $Servers) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool

        # The wrapper script:
        # - Redefines AppendLine to enqueue messages
        # - Defines Run-OnServer with credential parameter
        # - Runs the actual action
        # - Catches errors and enqueues them
        # - Enqueues results and a Done marker
        [void]$ps.AddScript({
            param($ServerName, $ActionScript, $UseAltCred, $Credential, $Queue, $Context)

            # Expand PreAction context values as variables for the action scriptblock
            if ($Context) {
                foreach ($ctxKey in $Context.Keys) {
                    Set-Variable -Name $ctxKey -Value $Context[$ctxKey]
                }
            }

            # Shim AppendLine: enqueue instead of touching GUI
            function AppendLine {
                param([string]$Text, $Col)
                $colorName = 'TextPrimary'
                if ($Col) {
                    # Map known Color objects to names by ARGB comparison
                    # Workers pass color variables from the outer scope - we store as-is
                    # and let the UI thread resolve. Store the Color object directly.
                }
                $Queue.Enqueue([PSCustomObject]@{
                    Type      = 'Line'
                    Server    = $ServerName
                    Text      = $Text
                    Color     = $Col
                    Data      = $null
                })
            }

            # Shim Run-OnServer: uses passed credential, no GUI access
            function Run-OnServer {
                param([string]$Server, [scriptblock]$SB)
                $isLocal = ($Server -eq $env:COMPUTERNAME -or
                            $Server -eq "localhost" -or
                            $Server -eq "127.0.0.1")
                if ($isLocal) { return (& $SB) }
                $p = @{ ComputerName = $Server; ScriptBlock = $SB; ErrorAction = "Stop" }
                if ($UseAltCred -and $Credential) {
                    $p["Credential"] = $Credential
                }
                return Invoke-Command @p
            }

            try {
                # Enqueue per-server header
                $Queue.Enqueue([PSCustomObject]@{
                    Type   = 'Line'
                    Server = $ServerName
                    Text   = "  >> $ServerName"
                    Color  = $null  # will use Accent color
                    Data   = 'ServerHeader'
                })

                $actionSB = [scriptblock]::Create($ActionScript)
                $result = & $actionSB -ServerName $ServerName

                if ($result) {
                    foreach ($item in $result) {
                        if ($item -is [PSCustomObject]) {
                            $Queue.Enqueue([PSCustomObject]@{
                                Type   = 'Result'
                                Server = $ServerName
                                Text   = $null
                                Color  = $null
                                Data   = $item
                            })
                        }
                    }
                }
            }
            catch {
                $Queue.Enqueue([PSCustomObject]@{
                    Type   = 'Error'
                    Server = $ServerName
                    Text   = "   ERROR: $_"
                    Color  = $null
                    Data   = [PSCustomObject]@{ Server = $ServerName; Error = $_.ToString() }
                })
            }

            # Signal this server is done
            $Queue.Enqueue([PSCustomObject]@{
                Type   = 'Done'
                Server = $ServerName
                Text   = $null
                Color  = $null
                Data   = $null
            })
        })

        [void]$ps.AddParameter('ServerName', $server)
        [void]$ps.AddParameter('ActionScript', $Action.ToString())
        [void]$ps.AddParameter('UseAltCred', $UseAltCred)
        [void]$ps.AddParameter('Credential', $Credential)
        [void]$ps.AddParameter('Queue', $Queue)
        [void]$ps.AddParameter('Context', $Context)

        $handle = $ps.BeginInvoke()
        $pipelines.Add([PSCustomObject]@{
            Pipeline = $ps
            Handle   = $handle
            Server   = $server
        })
    }

    return [PSCustomObject]@{
        Pool        = $pool
        Pipelines   = $pipelines
        Queue       = $Queue
        ServerCount = $Servers.Count
        StartTime   = Get-Date
    }
}
#endregion Concurrent execution infrastructure

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

#region TrustedHosts management for alternate credentials
$script:_SavedTrustedHosts = $null

function Save-AndAddTrustedHosts {
    param([string[]]$Servers)
    # Filter out local server names - they do not need TrustedHosts entries
    $remote = @($Servers | Where-Object {
        $_ -ne $env:COMPUTERNAME -and $_ -ne 'localhost' -and $_ -ne '127.0.0.1'
    })
    if ($remote.Count -eq 0) { return }
    try {
        $current = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction Stop).Value
        $script:_SavedTrustedHosts = $current
        # If already wildcard, nothing to add
        if ($current -eq '*') { return }
        $existing = if ($current) { $current -split ',' | ForEach-Object { $_.Trim() } } else { @() }
        $toAdd = @($remote | Where-Object { $_ -notin $existing })
        if ($toAdd.Count -eq 0) { return }
        $newValue = (($existing + $toAdd) | Where-Object { $_ }) -join ','
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value $newValue -Force -ErrorAction Stop
        AppendLine "TrustedHosts: temporarily added $($toAdd -join ', ')" $script:cTextMuted
    } catch {
        $script:_SavedTrustedHosts = $null
        AppendLine "WARNING: Could not update TrustedHosts (admin required): $_" $script:cWarning
    }
}

function Restore-TrustedHosts {
    if ($null -eq $script:_SavedTrustedHosts) { return }
    try {
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value $script:_SavedTrustedHosts -Force -ErrorAction Stop
    } catch {
        AppendLine "WARNING: Could not restore TrustedHosts: $_" $script:cWarning
    }
    $script:_SavedTrustedHosts = $null
}
#endregion TrustedHosts management for alternate credentials

# Main form
$script:form               = [System.Windows.Forms.Form]::new()
$script:form.Text          = $script:AppTitle
$script:form.Size          = [System.Drawing.Size]::new(1160, 870)
$script:form.MinimumSize   = [System.Drawing.Size]::new(1000, 870)
$script:form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$script:form.MaximizeBox   = $true
$script:form.ShowIcon      = $false
$script:form.StartPosition = "CenterScreen"
$script:form.BackColor     = $script:cBgDark
$script:form.ForeColor     = $script:cTextPrimary
$script:form.Font          = $script:fUI
$script:form.Add_Shown({ [DwmHelper]::SetDarkTitleBar($script:form.Handle) })
$script:form.Add_FormClosing({
    # Ensure TrustedHosts is restored if the form is closed mid-run
    Restore-TrustedHosts
    # Clean up any in-flight concurrent run
    if ($script:ConcurrentRun) {
        $script:pollTimer.Stop()
        foreach ($pi in $script:ConcurrentRun.Pipelines) {
            try { $pi.Pipeline.Stop() } catch {}
            try { $pi.Pipeline.Dispose() } catch {}
        }
        try { $script:ConcurrentRun.Pool.Close() } catch {}
        try { $script:ConcurrentRun.Pool.Dispose() } catch {}
        $script:ConcurrentRun = $null
    }
})

# Title bar
$script:pnlTitle           = [System.Windows.Forms.Panel]::new()
$script:pnlTitle.Dock      = "Top"
$script:pnlTitle.Height    = 48
$script:pnlTitle.BackColor = $script:cBgPanel

$script:lblTitle           = [System.Windows.Forms.Label]::new()
$script:lblTitle.Text      = $script:AppTitle
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
$script:txtServer.Text        = ""
$script:txtServer.ForeColor   = $script:cTextPrimary
$script:txtServer.Font        = $script:fUI

# Use native Windows cue banner for placeholder text - the OS handles
# showing/hiding it automatically when the textbox is empty vs. has input.
$script:txtServer.Add_HandleCreated({
    [TextBoxHelper]::SetPlaceholder($script:txtServer.Handle, "hostname or IP")
})

$script:btnAdd = New-Btn -Text "Add" -X 155 -Y 41 -W 45 -H 26 `
    -Bg $script:cAccent -Hover $script:cAccentHover -Font $script:fBold

$script:btnLoad = New-Btn -Text "Load" -X 10 -Y 74 -W 60 -H 28 `
    -Bg $script:cBgControl -Hover $script:cBgHover

$script:btnSave = New-Btn -Text "Save" -X 75 -Y 74 -W 60 -H 28 `
    -Bg $script:cBgControl -Hover $script:cBgHover

$script:btnAddLocal = New-Btn -Text "+ Local" -X 140 -Y 74 -W 60 -H 28 `
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
    $script:txtServer, $script:btnAdd, $script:btnLoad, $script:btnSave, $script:btnAddLocal,
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

# Cancel button - hidden by default, shown during concurrent runs
$script:btnCancelRun = New-Btn -Text "Cancel Run" -X 10 -Y 600 -W 190 -H 32 `
    -Bg $script:cDanger -Hover $script:cRedHover -Font $script:fBold
$script:btnCancelRun.Visible = $false
$script:btnCancelRun.Add_Click({
    $run = $script:ConcurrentRun
    if (-not $run) { return }

    # Stop all in-flight pipelines
    foreach ($pi in $run.Pipelines) {
        try { $pi.Pipeline.Stop() } catch {}
        try { $pi.Pipeline.Dispose() } catch {}
    }
    try { $run.Pool.Close() } catch {}
    try { $run.Pool.Dispose() } catch {}

    $script:pollTimer.Stop()

    AppendLine "" $script:cTextPrimary
    AppendLine "--- Run cancelled by user ---" $script:cWarning

    $script:statusLabel.Text      = "Cancelled"
    $script:statusLabel.ForeColor = $script:cWarning
    $script:statusRight.Text      = "Cancelled: $((Get-Date).ToString('HH:mm:ss'))"

    # Re-enable action buttons
    foreach ($ctrl in $script:pnlCenter.Controls) {
        if ($ctrl -is [System.Windows.Forms.Button]) {
            $ctrl.Enabled = $true
        }
    }
    $script:btnCancelRun.Visible = $false
    Restore-TrustedHosts
    $script:ConcurrentRun = $null
})
$script:pnlCenter.Controls.Add($script:btnCancelRun)

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

#region Concurrent execution timer
# Timer that drains the output queue from worker runspaces and updates the GUI.
# Runs on the UI thread so all control access is safe.
$script:pollTimer = [System.Windows.Forms.Timer]::new()
$script:pollTimer.Interval = 100
$script:pollTimer.Add_Tick({
    $run = $script:ConcurrentRun
    if (-not $run) { return }

    # Drain the output queue
    $item = $null
    $doneCount = 0
    while ($run.Queue.TryDequeue([ref]$item)) {
        switch ($item.Type) {
            'Line' {
                # Buffer every line for grouped replay on completion
                $script:_OutputBuffer.Add($item)
                if ($item.Data -eq 'ServerHeader') {
                    AppendLine $item.Text $script:cAccent
                } elseif ($item.Color) {
                    AppendLine "[$($item.Server)] $($item.Text)" $item.Color
                } else {
                    AppendLine "[$($item.Server)] $($item.Text)" $script:cTextPrimary
                }
            }
            'Result' {
                if ($item.Data) {
                    $script:Results.Add($item.Data)
                }
            }
            'Error' {
                # Buffer errors for grouped replay as well
                $script:_OutputBuffer.Add($item)
                AppendLine $item.Text $script:cDanger
                if ($item.Data) {
                    $script:Results.Add($item.Data)
                }
            }
            'Done' {
                # Track completed servers
                $run._CompletedCount++
            }
        }
    }

    # Check if all servers are done
    if ($run._CompletedCount -ge $run.ServerCount) {
        # Stop the timer
        $script:pollTimer.Stop()

        # Clean up pipelines
        foreach ($pi in $run.Pipelines) {
            try {
                $pi.Pipeline.EndInvoke($pi.Handle)
                $pi.Pipeline.Dispose()
            } catch {}
        }
        $run.Pool.Close()
        $run.Pool.Dispose()

        #region Grouped replay - replace interleaved output with server-grouped view
        $script:txtOutput.Clear()

        # Re-emit header
        if ($script:_RunHeaderText) {
            $script:txtOutput.SelectionStart  = $script:txtOutput.TextLength
            $script:txtOutput.SelectionLength = 0
            $script:txtOutput.SelectionColor  = $script:cTextMuted
            $script:txtOutput.AppendText("$($script:_RunHeaderText)`n")

            $script:txtOutput.SelectionStart  = $script:txtOutput.TextLength
            $script:txtOutput.SelectionLength = 0
            $script:txtOutput.SelectionColor  = $script:cBorder
            $script:txtOutput.AppendText(("-" * 80) + "`n")
        }

        # Group buffered lines by server, preserving order of first appearance
        $seenServers = [System.Collections.Generic.List[string]]::new()
        $groupedLines = @{}
        foreach ($entry in $script:_OutputBuffer) {
            $svr = $entry.Server
            if (-not $groupedLines.ContainsKey($svr)) {
                $seenServers.Add($svr)
                $groupedLines[$svr] = [System.Collections.Generic.List[PSObject]]::new()
            }
            $groupedLines[$svr].Add($entry)
        }

        foreach ($svr in $seenServers) {
            # Server header
            AppendLine "  >> $svr" $script:cAccent

            foreach ($entry in $groupedLines[$svr]) {
                $txt = $entry.Text -replace '^\s+', ''
                if ($entry.Type -eq 'Error') {
                    AppendLine $txt $script:cDanger
                } elseif ($entry.Data -eq 'ServerHeader') {
                    # Skip - we already wrote the grouped header above
                } elseif ($entry.Color) {
                    AppendLine $txt $entry.Color
                } else {
                    AppendLine $txt $script:cTextPrimary
                }
            }
        }

        $script:_OutputBuffer = $null
        #endregion Grouped replay

        # Summary footer
        $script:txtOutput.SelectionStart  = $script:txtOutput.TextLength
        $script:txtOutput.SelectionLength = 0
        $script:txtOutput.SelectionColor  = $script:cBorder
        $script:txtOutput.AppendText(("-" * 80) + "`n")

        $rc = $script:Results.Count
        $sc = $run.ServerCount

        $script:txtOutput.SelectionStart  = $script:txtOutput.TextLength
        $script:txtOutput.SelectionLength = 0
        $script:txtOutput.SelectionColor  = $script:cSuccess
        $script:txtOutput.AppendText("Done  |  $sc server(s)  |  $rc row(s)`n")
        $script:txtOutput.ScrollToCaret()

        $script:statusLabel.Text      = "Done - $rc result(s)"
        $script:statusLabel.ForeColor = $script:cSuccess
        $script:statusRight.Text      = "Last run: $((Get-Date).ToString('HH:mm:ss'))"

        # Re-enable action buttons and cancel
        foreach ($ctrl in $script:pnlCenter.Controls) {
            if ($ctrl -is [System.Windows.Forms.Button]) {
                $ctrl.Enabled = $true
            }
        }
        $script:btnCancelRun.Visible = $false

        Restore-TrustedHosts
        $script:ConcurrentRun = $null
    }
})
#endregion Concurrent execution timer

# Action button factory with rich Tag for dynamic discovery
function New-ActionBtn {
    param([string]$Label, [int]$Y, [scriptblock]$Action, [switch]$Interactive, [scriptblock]$PreAction)

    $b = New-Btn -Text $Label -X 10 -Y $Y -W 190 -H 36 `
        -Bg $script:cActBtn -Hover $script:cAccent
    $b.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $b.Padding   = [System.Windows.Forms.Padding]::new(8, 0, 0, 0)

    # Keep ActionMap in case you still want it
    $safeKey = "Act_" + ($Label -replace '\W','_')
    $script:ActionMap = if ($script:ActionMap) { $script:ActionMap } else { @{} }
    $script:ActionMap[$safeKey]           = $Action
    $script:ActionMap[$safeKey + "_Label"] = $Label

    # Tag for dynamic discovery by Multi-Action
    $b.Tag = [PSCustomObject]@{
        Type        = "Action"
        Key         = $safeKey
        Label       = $Label
        Action      = $Action
        Interactive = [bool]$Interactive
        PreAction   = $PreAction
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

        # Prevent re-entry while a concurrent run is active
        if ($script:ConcurrentRun) {
            [System.Windows.Forms.MessageBox]::Show(
                "An action is already running. Cancel it first or wait for it to finish.",
                "Busy",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $script:txtOutput.Clear()
        $script:Results.Clear()
        $script:_OutputBuffer = [System.Collections.Generic.List[PSObject]]::new()
        $script:_RunHeaderText = $null  # saved for grouped replay

        $ts      = (Get-Date).ToString("HH:mm:ss")
        $cm      = if ($script:rdoCurrent.Checked) { "current account" } else { "alt credentials" }
        $svrList = $servers -join ", "

        # Audit log every action execution
        Write-AuditLog -Message ("Action: $cl`nUser: $env:USERDOMAIN\$env:USERNAME`nCredentials: $cm`nServers: $svrList") -EventId 1000

        $script:txtOutput.SelectionStart  = $script:txtOutput.TextLength
        $script:txtOutput.SelectionLength = 0
        $script:txtOutput.SelectionColor  = $script:cTextMuted
        $script:txtOutput.AppendText("[$ts]  $cl  |  $svrList  |  $cm`n")

        $script:txtOutput.SelectionStart  = $script:txtOutput.TextLength
        $script:txtOutput.SelectionLength = 0
        $script:txtOutput.SelectionColor  = $script:cBorder
        $script:txtOutput.AppendText(("-" * 80) + "`n")

        # Determine execution mode:
        # - PreAction: prompt on UI thread (dialogs), then run concurrently with context
        # - Interactive (no PreAction): sequential per-server (for per-server UI interaction)
        # - Default: concurrent via runspace pool
        $actionContext = @{}
        if ($meta.PreAction) {
            $actionContext = & $meta.PreAction
            if (-not $actionContext) {
                $script:statusLabel.Text      = "Cancelled"
                $script:statusLabel.ForeColor = $script:cTextMuted
                return
            }
        }

        # Temporarily add remote servers to TrustedHosts when using alternate credentials
        if ($script:rdoAlt.Checked -and $script:AltCred) {
            Save-AndAddTrustedHosts -Servers $servers
        }

        if ($meta.Interactive -and -not $meta.PreAction) {
            # Sequential path - for actions that need per-server UI interaction
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

            # Write summary footer for sequential path
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
            Restore-TrustedHosts
        } else {
            # Concurrent path - disable buttons, launch runspace pool, start timer
            $useAlt = $script:rdoAlt.Checked
            $cred   = $script:AltCred

            # Disable action buttons during run
            foreach ($ctrl in $script:pnlCenter.Controls) {
                if ($ctrl -is [System.Windows.Forms.Button] -and $ctrl -ne $script:btnCancelRun) {
                    $ctrl.Enabled = $false
                }
            }
            $script:btnCancelRun.Visible = $true

            $script:statusLabel.Text      = "Running $cl on $($servers.Count) server(s) concurrently..."
            $script:statusLabel.ForeColor = $script:cWarning

            # Save header text so grouped replay can re-emit it
            $script:_RunHeaderText = "[$ts]  $cl  |  $svrList  |  $cm"

            # Clear and prepare the output queue
            $item = $null
            while ($script:OutputQueue.TryDequeue([ref]$item)) {}

            $script:ConcurrentRun = Start-ConcurrentActions `
                -Action $ca `
                -Servers $servers `
                -UseAltCred $useAlt `
                -Credential $cred `
                -Queue $script:OutputQueue `
                -Context $actionContext

            # Add tracking counter for completed servers
            $script:ConcurrentRun | Add-Member -NotePropertyName '_CompletedCount' -NotePropertyValue 0

            # Start the polling timer - the timer tick handler drains the queue
            $script:pollTimer.Start()
        }
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
        Label  = "Run Remote Command"
        PreAction = {
            # Runs on the UI thread to show dialogs before concurrent dispatch.
            # Returns @{ RemoteCode = $code } for the concurrent action, or $null to cancel.
                $allServers = @($script:clbServers.CheckedItems)

                # Build a small modal dialog to capture the script text
                $dlg        = New-Object System.Windows.Forms.Form
                $dlg.Text   = "Run Remote on $($allServers.Count) server(s)"
                $dlg.Size   = [System.Drawing.Size]::new(600, 400)
                $dlg.StartPosition = "CenterParent"
                $dlg.BackColor     = $script:cBgPanel
                $dlg.ForeColor     = $script:cTextPrimary
                $dlg.Font          = $script:fUI
                $dlg.TopMost       = $true

                $lblInfo           = New-Object System.Windows.Forms.Label
                $lblInfo.Text      = "Enter PowerShell to run on $($allServers.Count) server(s):"
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
                # Populate with last-used command, or example template on first use
                if ($script:_LastRemoteCode) {
                    $txtScript.Text = $script:_LastRemoteCode
                } else {
                    $txtScript.Text = '# Example:' + "`r`n" +
                                     'Get-Service | Where-Object Status -eq "Running" | Select-Object -First 10'
                }

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
                    AppendLine "   Remote run cancelled" $script:cTextMuted
                    return $null
                }

                $code = $dlg.Tag

                # Scrollable confirmation dialog showing code + full server list
                $cfmDlg              = New-Object System.Windows.Forms.Form
                $cfmDlg.Text         = "Confirm Remote Execution"
                $cfmDlg.Size         = [System.Drawing.Size]::new(620, 480)
                $cfmDlg.StartPosition = "CenterParent"
                $cfmDlg.BackColor    = $script:cBgPanel
                $cfmDlg.ForeColor    = $script:cTextPrimary
                $cfmDlg.Font         = $script:fUI
                $cfmDlg.TopMost      = $true

                $cfmLabel            = New-Object System.Windows.Forms.Label
                $cfmLabel.Text       = "Review and confirm the code and target servers:"
                $cfmLabel.AutoSize   = $true
                $cfmLabel.Location   = [System.Drawing.Point]::new(10, 10)
                $cfmLabel.ForeColor  = $script:cWarning

                $svrListText = $allServers -join "`r`n"
                $cfmBody = "CODE TO EXECUTE:`r`n" +
                           ("-" * 50) + "`r`n" +
                           $code + "`r`n`r`n" +
                           "TARGET SERVERS ($($allServers.Count)):`r`n" +
                           ("-" * 50) + "`r`n" +
                           $svrListText

                $cfmText             = New-Object System.Windows.Forms.TextBox
                $cfmText.Multiline   = $true
                $cfmText.ScrollBars  = "Both"
                $cfmText.ReadOnly    = $true
                $cfmText.WordWrap    = $false
                $cfmText.Location    = [System.Drawing.Point]::new(10, 35)
                $cfmText.Size        = [System.Drawing.Size]::new(580, 350)
                $cfmText.BackColor   = $script:cBgDark
                $cfmText.ForeColor   = $script:cTextPrimary
                $cfmText.Font        = $script:fMono
                $cfmText.Text        = $cfmBody
                $cfmText.TabStop     = $false

                $cfmYes  = New-Btn -Text "Execute" -X 390 -Y 395 -W 95 -H 32 `
                               -Bg $script:cDanger -Hover $script:cRedHover -Font $script:fBold
                $cfmNo   = New-Btn -Text "Cancel"  -X 495 -Y 395 -W 95 -H 32 `
                               -Bg $script:cBgControl -Hover $script:cBgHover

                $cfmYes.Add_Click({
                    $cfmDlg.DialogResult = [System.Windows.Forms.DialogResult]::Yes
                    $cfmDlg.Close()
                })
                $cfmNo.Add_Click({
                    $cfmDlg.DialogResult = [System.Windows.Forms.DialogResult]::No
                    $cfmDlg.Close()
                })

                $cfmDlg.Controls.AddRange(@($cfmLabel, $cfmText, $cfmYes, $cfmNo))
                $cfmResult = $cfmDlg.ShowDialog()

                if ($cfmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
                    AppendLine "   Remote run aborted by user" $script:cTextMuted
                    return $null
                }

                # Audit log the remote code execution (once for the entire batch)
                $credMode = if ($script:rdoCurrent.Checked) { "current account" } else { "alt credentials" }
                Write-AuditLog -Message ("Run Remote executed`nUser: $env:USERDOMAIN\$env:USERNAME`nCredentials: $credMode`nServers ($($allServers.Count)): $($allServers -join ', ')`nCode:`n$code") -EventId 1001 -EntryType Warning

                # Return context hashtable for the concurrent action scriptblock.
                # The key 'RemoteCode' becomes a $RemoteCode variable inside each runspace.
                $script:_LastRemoteCode = $code
                return @{ RemoteCode = $code }
        }
        Action = {
            param($ServerName)
            # $RemoteCode is injected from the PreAction context by the runspace wrapper
            AppendLine "   Running remote script on $ServerName ..." $null

            $scriptBlock = [scriptblock]::Create($RemoteCode)
            $output = Run-OnServer $ServerName $scriptBlock

            if ($null -eq $output) {
                AppendLine "   (no output)" $null
                return [PSCustomObject]@{ Server = $ServerName; Output = "(no output)" }
            }

            $results = @()
            foreach ($line in $output) {
                $text = ($line | Out-String).TrimEnd()
                AppendLine ("   $text") $null
                $results += [PSCustomObject]@{
                    Server = $ServerName
                    Output = $text
                }
            }
            return $results
        }
    },
        @{
        Label  = "Multi-Action"
        PreAction = {
            # Runs on the UI thread to show the action selection dialog once.
            # Returns @{ SelectedActions = @( @{Label=...; ActionScript=...}, ... ) } or $null to cancel.

            # Discover available actions from buttons in the Actions panel at runtime
            $availableActions = @()
            foreach ($ctrl in $script:pnlCenter.Controls) {
                if ($ctrl -is [System.Windows.Forms.Button] -and $ctrl.Tag) {
                    $tag = $ctrl.Tag
                    if ($tag.Type -eq "Action" -and $tag.Label -ne "Multi-Action" -and -not $tag.PreAction) {
                        $availableActions += $tag
                    }
                }
            }

            if ($availableActions.Count -eq 0) {
                AppendLine "   Multi-Action: no actions discovered in the UI." $script:cDanger
                return $null
            }

            # Build a lookup: label -> action metadata
            $actionLookup = @{}
            foreach ($act in $availableActions) {
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

            foreach ($act in $availableActions) {
                $isChecked = $script:_LastGetManySelections -and $script:_LastGetManySelections -contains $act.Label
                [void]$lst.Items.Add($act.Label, $isChecked)
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
                AppendLine "   Multi-Action cancelled" $script:cTextMuted
                return $null
            }

            # Serialize selected actions as label + scriptblock text pairs
            # (scriptblocks cannot cross runspace boundaries, so we pass the text)
            $selected = @()
            $script:_LastGetManySelections = @($lst.CheckedItems)
            foreach ($label in $lst.CheckedItems) {
                if ($actionLookup.ContainsKey($label)) {
                    $selected += @{
                        Label        = $label
                        ActionScript = $actionLookup[$label].Action.ToString()
                    }
                }
            }

            if ($selected.Count -eq 0) {
                AppendLine "   Multi-Action: no actions selected after dialog." $script:cTextMuted
                return $null
            }

            return @{ SelectedActions = $selected }
        }
        Action = {
            param($ServerName)
            # $SelectedActions is injected from the PreAction context by the runspace wrapper.
            # Each entry has Label and ActionScript keys.

            AppendLine "   Multi-Action: running $($SelectedActions.Count) action(s) on $ServerName..." $null

            foreach ($act in $SelectedActions) {
                AppendLine "     -> $($act.Label)" $null
                try {
                    $actionSB = [scriptblock]::Create($act.ActionScript)
                    $res = & $actionSB -ServerName $ServerName
                    if ($res) {
                        foreach ($item in $res) {
                            if ($item -is [PSCustomObject]) {
                                $Queue.Enqueue([PSCustomObject]@{
                                    Type   = 'Result'
                                    Server = $ServerName
                                    Text   = $null
                                    Color  = $null
                                    Data   = $item
                                })
                            }
                        }
                    }
                } catch {
                    AppendLine "       ERROR in '$($act.Label)': $_" $null
                    $Queue.Enqueue([PSCustomObject]@{
                        Type   = 'Result'
                        Server = $ServerName
                        Text   = $null
                        Color  = $null
                        Data   = [PSCustomObject]@{
                            Server = $ServerName
                            Action = $act.Label
                            Error  = $_.ToString()
                        }
                    })
                }
            }
        }
    }

        ,@{
        Label     = "Reboot"
        PreAction = {
            # Confirmation dialog before rebooting servers
            $allServers = @($script:clbServers.CheckedItems)

            # Determine credential identity for audit trail
            $credIdentity = if ($script:rdoAlt.Checked -and $script:AltCred) {
                $script:AltCred.UserName
            } else {
                "$env:USERDOMAIN\$env:USERNAME"
            }

            # Scrollable confirmation dialog matching Run Remote style
            $cfmDlg              = New-Object System.Windows.Forms.Form
            $cfmDlg.Text         = "Confirm Reboot - $($allServers.Count) server(s)"
            $cfmDlg.Size         = [System.Drawing.Size]::new(620, 480)
            $cfmDlg.StartPosition = "CenterParent"
            $cfmDlg.BackColor    = $script:cBgPanel
            $cfmDlg.ForeColor    = $script:cTextPrimary
            $cfmDlg.Font         = $script:fUI
            $cfmDlg.FormBorderStyle = "FixedDialog"
            $cfmDlg.MaximizeBox  = $false
            $cfmDlg.MinimizeBox  = $false
            $cfmDlg.TopMost      = $true
            try { [DwmHelper]::SetDarkTitleBar($cfmDlg.Handle) } catch {}

            $cfmLabel            = New-Object System.Windows.Forms.Label
            $cfmLabel.Text       = "Review and confirm the reboot of the following servers:"
            $cfmLabel.AutoSize   = $true
            $cfmLabel.Location   = [System.Drawing.Point]::new(10, 10)
            $cfmLabel.ForeColor  = $script:cWarning

            $svrListText = $allServers -join "`r`n"
            $cfmBody = "ACTION: Force Reboot`r`n" +
                       "CREDENTIAL: $credIdentity`r`n" +
                       "TIMEOUT: 15 minutes per server`r`n" +
                       "`r`n" +
                       "Monitoring will be suppressed before reboot and`r`n" +
                       "resumed automatically for servers that come back online.`r`n" +
                       "`r`n" +
                       "TARGET SERVERS ($($allServers.Count)):`r`n" +
                       ("-" * 50) + "`r`n" +
                       $svrListText

            $cfmText             = New-Object System.Windows.Forms.TextBox
            $cfmText.Multiline   = $true
            $cfmText.ScrollBars  = "Both"
            $cfmText.ReadOnly    = $true
            $cfmText.WordWrap    = $false
            $cfmText.Location    = [System.Drawing.Point]::new(10, 35)
            $cfmText.Size        = [System.Drawing.Size]::new(580, 350)
            $cfmText.BackColor   = $script:cBgDark
            $cfmText.ForeColor   = $script:cTextPrimary
            $cfmText.Font        = $script:fMono
            $cfmText.Text        = $cfmBody
            $cfmText.TabStop     = $false

            $cfmYes  = New-Btn -Text "Reboot" -X 390 -Y 395 -W 95 -H 32 `
                           -Bg $script:cDanger -Hover $script:cRedHover -Font $script:fBold
            $cfmNo   = New-Btn -Text "Cancel" -X 495 -Y 395 -W 95 -H 32 `
                           -Bg $script:cBgControl -Hover $script:cBgHover

            $cfmYes.Add_Click({
                $cfmDlg.DialogResult = [System.Windows.Forms.DialogResult]::Yes
                $cfmDlg.Close()
            })
            $cfmNo.Add_Click({
                $cfmDlg.DialogResult = [System.Windows.Forms.DialogResult]::No
                $cfmDlg.Close()
            })

            $cfmDlg.Controls.AddRange(@($cfmLabel, $cfmText, $cfmYes, $cfmNo))
            $cfmResult = $cfmDlg.ShowDialog()

            if ($cfmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
                AppendLine "   Reboot cancelled by user" $script:cTextMuted
                return $null
            }

            # Resolve downadmin path for use in per-server concurrent actions
            $daPath = Join-Path $env:SystemDrive "iPosh\DownAdmin.ps1"
            if (-not (Test-Path $daPath)) {
                $daPath = (Get-Command downadmin -ErrorAction SilentlyContinue).Source
            }
            if (-not $daPath -or -not (Test-Path $daPath)) {
                AppendLine "   WARNING: downadmin not found - monitoring will not be suppressed" $script:cWarning
            }

            return @{ Confirmed = $true; DAPath = $daPath; CredIdentity = $credIdentity }
        }
        Action = {
            param($ServerName)
            $timeoutSec = 900  # 15 minutes
            $isLocal = ($ServerName -eq $env:COMPUTERNAME -or
                        $ServerName -eq 'localhost' -or
                        $ServerName -eq '127.0.0.1')

            # Build credential params for direct cmdlet calls
            $credParam = @{}
            if ($UseAltCred -and $Credential) {
                $credParam['Credential'] = $Credential
            }

            # Suppress monitoring before rebooting
            if ($DAPath -and (Test-Path $DAPath)) {
                try {
                    AppendLine "   Suppressing monitoring for $ServerName" $script:cTextMuted
                    & $DAPath -m $ServerName -hours 1 -comment "Reboot from Remote Commander by $CredIdentity" -PipelineOutput | Out-Null
                    AppendLine "   Monitoring suppressed for $ServerName" $script:cSuccess
                } catch {
                    AppendLine "   WARNING: Could not suppress monitoring for $ServerName : $_" $script:cWarning
                }
            }

            # Issue the restart
            try {
                if ($isLocal) {
                    Restart-Computer -Force -ErrorAction Stop
                } else {
                    Restart-Computer -ComputerName $ServerName -Force @credParam -ErrorAction Stop
                }
            } catch {
                AppendLine "   FAILED to initiate reboot: $_" $script:cDanger
                return [PSCustomObject]@{ Server=$ServerName; Action="Reboot"; Status="Failed"; Detail=$_.ToString() }
            }

            AppendLine "   Reboot initiated - waiting for $ServerName to come back (up to 15 min)..." $script:cWarning

            # Wait for the server to go offline first (max 60 seconds)
            $offlineTimer = [System.Diagnostics.Stopwatch]::StartNew()
            $wentOffline = $false
            while ($offlineTimer.Elapsed.TotalSeconds -lt 60) {
                Start-Sleep -Seconds 3
                $ping = Test-Connection -ComputerName $ServerName -Count 1 -Quiet -ErrorAction SilentlyContinue
                if (-not $ping) {
                    $wentOffline = $true
                    AppendLine "   $ServerName is offline - waiting for it to come back..." $script:cTextMuted
                    break
                }
            }
            if (-not $wentOffline) {
                AppendLine "   WARNING: $ServerName did not go offline within 60s - may not have rebooted" $script:cWarning
            }

            # Poll until back online or timeout - progress every 15 seconds
            $pollTimer = [System.Diagnostics.Stopwatch]::StartNew()
            $online = $false
            $lastReport = 0
            while ($pollTimer.Elapsed.TotalSeconds -lt $timeoutSec) {
                Start-Sleep -Seconds 5
                # Use Test-WSMan for a more reliable check that WinRM is ready
                $wsm = $null
                try {
                    $wsmParam = @{ ComputerName = $ServerName; ErrorAction = 'Stop' }
                    if ($credParam.Count -gt 0) { $wsmParam += $credParam }
                    $wsm = Test-WSMan @wsmParam
                } catch {}
                if ($wsm) {
                    $online = $true
                    break
                }
                # Progress update every 15 seconds with time remaining
                $elapsed = [math]::Round($pollTimer.Elapsed.TotalSeconds)
                if (($elapsed - $lastReport) -ge 15) {
                    $remaining = [math]::Round(($timeoutSec - $elapsed) / 60, 1)
                    AppendLine "   Waiting for $ServerName... ${elapsed}s elapsed, ${remaining} min remaining" $script:cTextMuted
                    $lastReport = $elapsed
                }
            }

            if ($online) {
                $elapsed = [math]::Round($pollTimer.Elapsed.TotalSeconds)
                AppendLine "   $ServerName is back ONLINE after $elapsed seconds" $script:cSuccess

                # Resume monitoring now that the server is back
                if ($DAPath -and (Test-Path $DAPath)) {
                    try {
                        AppendLine "   Resuming monitoring for $ServerName" $script:cTextMuted
                        & $DAPath -m $ServerName -unlock -PipelineOutput | Out-Null
                        AppendLine "   Monitoring resumed for $ServerName" $script:cSuccess
                    } catch {
                        AppendLine "   WARNING: Could not resume monitoring for $ServerName : $_" $script:cWarning
                    }
                }

                return [PSCustomObject]@{ Server=$ServerName; Action="Reboot"; Status="Online"; RecoveryTime_s=$elapsed; Monitoring="Resumed" }
            } else {
                AppendLine "   TIMEOUT: $ServerName did not come back within 15 minutes" $script:cDanger
                AppendLine "   Monitoring remains SUPPRESSED for $ServerName - resume manually when server is back" $script:cWarning
                return [PSCustomObject]@{ Server=$ServerName; Action="Reboot"; Status="Timeout"; RecoveryTime_s="N/A"; Monitoring="Suppressed" }
            }
        }
    }

        ,@{
        Label       = "DownAdmin"
        PreAction = {
            # Runs on the UI thread once before dispatching to all selected servers.
            # Returns @{ DAChoice = 'On'|'Off' } for the concurrent action, or $null to cancel.
            $allServers = @($script:clbServers.CheckedItems)

            # Build a dark-themed dialog matching the main form style
            $dlg              = New-Object System.Windows.Forms.Form
            $dlg.Text         = "DownAdmin - Suppress or Resume Monitoring"
            $dlg.Size         = [System.Drawing.Size]::new(420, 200)
            $dlg.StartPosition = "CenterParent"
            $dlg.BackColor    = $script:cBgPanel
            $dlg.ForeColor    = $script:cTextPrimary
            $dlg.Font         = $script:fUI
            $dlg.FormBorderStyle = "FixedDialog"
            $dlg.MaximizeBox  = $false
            $dlg.MinimizeBox  = $false
            $dlg.TopMost      = $true
            try { [DwmHelper]::SetDarkTitleBar($dlg.Handle) } catch {}

            $lblInfo          = New-Object System.Windows.Forms.Label
            $lblInfo.Text     = "Choose an action for $($allServers.Count) selected server(s):"
            $lblInfo.AutoSize = $true
            $lblInfo.Location = [System.Drawing.Point]::new(20, 20)
            $lblInfo.Font     = $script:fBold
            $lblInfo.ForeColor = $script:cTextPrimary

            $btnSuppress      = New-Btn -Text "Suppress Monitoring" -X 20 -Y 60 -W 170 -H 36 `
                                        -Bg $script:cRed -Hover $script:cRedHover
            $btnResume        = New-Btn -Text "Resume Monitoring" -X 200 -Y 60 -W 170 -H 36 `
                                        -Bg $script:cGreen -Hover $script:cGreenHover
            $btnCancel        = New-Btn -Text "Cancel" -X 140 -Y 110 -W 120 -H 30 `
                                        -Bg $script:cBgControl -Hover $script:cBgHover

            $btnSuppress.Add_Click({
                $dlg.Tag = 'Off'
                $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $dlg.Close()
            })
            $btnResume.Add_Click({
                $dlg.Tag = 'On'
                $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $dlg.Close()
            })
            $btnCancel.Add_Click({
                $dlg.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $dlg.Close()
            })

            $dlg.Controls.AddRange(@($lblInfo, $btnSuppress, $btnResume, $btnCancel))

            $result = $dlg.ShowDialog()
            if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
                return $null
            }

            return @{ DAChoice = $dlg.Tag }
        }
        Action = {
            param($ServerName)

            # $DAChoice is injected from the PreAction context by the runspace wrapper.
            # downadmin is always run locally - it handles remote communication
            # to BigPanda/EMCLI internally via the -m parameter.
            $daPath = Join-Path $env:SystemDrive "iPosh\DownAdmin.ps1"
            if (-not (Test-Path $daPath)) {
                $daPath = (Get-Command downadmin -ErrorAction SilentlyContinue).Source
            }

            if ($DAChoice -eq 'Off') {
                $mode = "Suppress (1h)"
                $daArgs = @{ m = $ServerName; hours = 1; comment = "Turning Monitoring Off"; PipelineOutput = $true }
            } else {
                $mode = "Resume"
                $daArgs = @{ m = $ServerName; unlock = $true; PipelineOutput = $true }
            }

            AppendLine "   DownAdmin: $mode for $ServerName" $script:cTextPrimary
            AppendLine "      Command: downadmin -m $ServerName $(if ($DAChoice -eq 'Off') { '-hours 1 -comment "Turning Monitoring Off"' } else { '-unlock' })" $script:cTextMuted

            $results = @()
            try {
                # downadmin runs locally and targets the host via -m (no remoting)
                $output = & $daPath @daArgs

                if ($output) {
                    foreach ($line in $output) {
                        AppendLine ("      " + ($line | Out-String).TrimEnd()) $script:cTextPrimary
                        $results += [PSCustomObject]@{
                            Server   = $ServerName
                            Action   = "DownAdmin"
                            Mode     = $mode
                            Output   = ($line | Out-String).TrimEnd()
                        }
                    }
                } else {
                    AppendLine "      (no output from downadmin)" $script:cTextMuted
                    $results += [PSCustomObject]@{
                        Server = $ServerName
                        Action = "DownAdmin"
                        Mode   = $mode
                        Output = "(no output)"
                    }
                }
            }
            catch {
                AppendLine "      ERROR running downadmin: $_" $script:cDanger
                $results += [PSCustomObject]@{
                    Server = $ServerName
                    Action = "DownAdmin"
                    Mode   = $mode
                    Error  = $_.ToString()
                }
            }
            return $results
        }
    }

)

$yPos = 42
foreach ($a in $actions) {
    $splat = @{
        Label       = $a.Label
        Y           = $yPos
        Action      = $a.Action
        Interactive = ($a.Interactive -eq $true)
    }
    if ($a.PreAction) { $splat['PreAction'] = $a.PreAction }
    $ab = New-ActionBtn @splat
    $script:pnlCenter.Controls.Add($ab)
    $yPos += 44
}

# Event handlers
$script:btnAdd.Add_Click({
    $name = $script:txtServer.Text.Trim()
    if ($name -and $script:clbServers.Items -notcontains $name) {
        $script:clbServers.Items.Add($name, $true) | Out-Null
        $script:txtServer.Text = ""
        $script:txtServer.Focus()
        $script:statusLabel.Text      = "Added: $name"
        $script:statusLabel.ForeColor = $script:cSuccess
        $script:statusRight.Text = "$($script:clbServers.Items.Count) server(s)"
    }
})

$script:btnAbout.Add_Click({
    $dlg              = New-Object System.Windows.Forms.Form
    $dlg.Text         = "About $($script:AppTitle)"
    $dlg.Size         = [System.Drawing.Size]::new(360, 160)
    $dlg.StartPosition = "CenterParent"
    $dlg.BackColor    = $script:cBgPanel
    $dlg.ForeColor    = $script:cTextPrimary
    $dlg.Font         = $script:fUI
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox  = $false
    $dlg.MinimizeBox  = $false
    $dlg.TopMost      = $true
    try { [DwmHelper]::SetDarkTitleBar($dlg.Handle) } catch {}

    $lbl              = New-Object System.Windows.Forms.Label
    $lbl.Text         = "Property of Intel IT - Engineering Computing"
    $lbl.AutoSize     = $false
    $lbl.Size         = [System.Drawing.Size]::new(320, 40)
    $lbl.Location     = [System.Drawing.Point]::new(20, 25)
    $lbl.ForeColor    = $script:cTextPrimary
    $lbl.Font         = $script:fBold
    $lbl.TextAlign    = [System.Drawing.ContentAlignment]::MiddleCenter

    $btnOK            = New-Btn -Text "OK" -X 135 -Y 80 -W 80 -H 30 `
                                -Bg $script:cAccent -Hover $script:cAccentHover
    $btnOK.Add_Click({
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dlg.Close()
    })

    $dlg.Controls.AddRange(@($lbl, $btnOK))
    $dlg.ShowDialog() | Out-Null
})


$script:btnAddLocal.Add_Click({
    $name = $env:COMPUTERNAME
    if ($name -and $script:clbServers.Items -notcontains $name) {
        $script:clbServers.Items.Add($name, $true) | Out-Null
        $script:statusLabel.Text      = "Added: $name"
        $script:statusLabel.ForeColor = $script:cSuccess
        $script:statusRight.Text = "$($script:clbServers.Items.Count) server(s)"
    }
})

$script:txtServer.Add_KeyDown({
    if ($_.KeyCode -eq "Return") { $script:btnAdd.PerformClick() }
})

$script:btnSave.Add_Click({
    if ($script:clbServers.Items.Count -eq 0) {
        $script:statusLabel.Text      = "No servers to save"
        $script:statusLabel.ForeColor = $script:cWarning
        return
    }
    $dlg = [System.Windows.Forms.SaveFileDialog]::new()
    $dlg.Filter   = "Text File (*.txt)|*.txt|CSV File (*.csv)|*.csv"
    $dlg.Title    = "Save Server List"
    $dlg.FileName = "servers.txt"
    if ($dlg.ShowDialog() -eq "OK") {
        $script:clbServers.Items | Out-File -FilePath $dlg.FileName -Encoding ASCII
        AppendLine "Saved $($script:clbServers.Items.Count) server(s) to: $($dlg.FileName)" $script:cTextMuted
        $script:statusLabel.Text      = "Saved $($script:clbServers.Items.Count) server(s)"
        $script:statusLabel.ForeColor = $script:cSuccess
    }
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
            $xl = $null; $wb = $null; $ws = $null
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
                $script:statusLabel.Text      = "Exported: $($dlg.FileName)"
                $script:statusLabel.ForeColor = $script:cSuccess
            } catch {
                AppendLine "Excel error: $_" $script:cDanger
                $script:statusLabel.Text      = "Excel failed. Try: Install-Module ImportExcel"
                $script:statusLabel.ForeColor = $script:cDanger
            } finally {
                # Release all COM objects to prevent orphaned EXCEL.EXE processes
                foreach ($comObj in @($ws, $wb, $xl)) {
                    if ($comObj) {
                        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($comObj) | Out-Null } catch {}
                    }
                }
            }
        }
    }
})

# Assemble and launch
$script:pnlMain           = [System.Windows.Forms.Panel]::new()
$script:pnlMain.Dock      = "Fill"
$script:pnlMain.BackColor = $script:cBgPanel
$script:pnlMain.Controls.AddRange(@(
    $script:pnlLeft, $script:pnlCenter, $script:pnlRight
))

#region Resize handlers - dynamically position controls when form is resized
# Anchoring does not work here because pnlMain is Dock=Fill and its initial
# size (200x100 default) is wrong when children are added, causing anchor
# math to produce negative distances and broken layout. Instead, we
# manually position everything in Resize events.

$script:pnlMain.Add_Resize({
    $pw = $script:pnlMain.ClientSize.Width
    $ph = $script:pnlMain.ClientSize.Height
    if ($pw -lt 10 -or $ph -lt 10) { return }  # ignore degenerate sizes during init

    # Three-column layout: left and center fixed width, right fills remainder
    $script:pnlLeft.SetBounds(8, 8, 210, ($ph - 8))
    $script:pnlCenter.SetBounds(226, 8, 210, ($ph - 8))
    $script:pnlRight.SetBounds(444, 8, ($pw - 452), ($ph - 8))
})

$script:pnlLeft.Add_Resize({
    $lh = $script:pnlLeft.ClientSize.Height
    if ($lh -lt 10) { return }

    # Server checklist fills between top controls (Y=110) and bottom section
    $script:clbServers.Size = [System.Drawing.Size]::new(190, [Math]::Max(60, $lh - 320))

    # Bottom section: All/None, Remove, separator, Credentials - pinned to bottom
    $script:btnSelAll.Location   = [System.Drawing.Point]::new(10,  ($lh - 202))
    $script:btnNone.Location     = [System.Drawing.Point]::new(108, ($lh - 202))
    $script:btnRemove.Location   = [System.Drawing.Point]::new(10,  ($lh - 170))
    $script:sep2.Location        = [System.Drawing.Point]::new(10,  ($lh - 132))
    $script:lblCred.Location     = [System.Drawing.Point]::new(10,  ($lh - 125))
    $script:rdoCurrent.Location  = [System.Drawing.Point]::new(10,  ($lh - 105))
    $script:rdoAlt.Location      = [System.Drawing.Point]::new(10,  ($lh - 83))
    $script:btnSetCred.Location  = [System.Drawing.Point]::new(10,  ($lh - 60))
})

$script:pnlCenter.Add_Resize({
    $ch = $script:pnlCenter.ClientSize.Height
    if ($ch -lt 10) { return }

    $script:btnCancelRun.Location = [System.Drawing.Point]::new(10, ($ch - 90))
    $script:btnAbout.Location     = [System.Drawing.Point]::new(10, ($ch - 50))
})

$script:pnlRight.Add_Resize({
    $rw = $script:pnlRight.ClientSize.Width
    $rh = $script:pnlRight.ClientSize.Height
    if ($rw -lt 10 -or $rh -lt 10) { return }

    $script:sep4.Size             = [System.Drawing.Size]::new(($rw - 20), 1)
    $script:txtOutput.Size        = [System.Drawing.Size]::new(($rw - 20), [Math]::Max(60, $rh - 100))
    $script:btnExportCSV.Location = [System.Drawing.Point]::new(10,  ($rh - 52))
    $script:btnExportXL.Location  = [System.Drawing.Point]::new(138, ($rh - 52))
    $script:btnClearOut.Location  = [System.Drawing.Point]::new(276, ($rh - 52))
    $script:btnCopyOut.Location   = [System.Drawing.Point]::new(404, ($rh - 52))
})

$script:pnlTitle.Add_Resize({
    $tw = $script:pnlTitle.ClientSize.Width
    if ($tw -lt 10) { return }
    $script:lblUser.Location = [System.Drawing.Point]::new(($tw - $script:lblUser.Width - 14), 16)
})
#endregion Resize handlers

$script:form.Controls.AddRange(@(
    $script:pnlTitle, $script:pnlMain, $script:statusBar
))

$script:statusLabel.Text      = "Ready  -  $env:USERDOMAIN\$env:USERNAME"
$script:statusLabel.ForeColor = $script:cTextMuted
AppendLine "$($script:AppTitle)  |  $env:USERDOMAIN\$env:USERNAME  |  $env:COMPUTERNAME" $script:cTextMuted
AppendLine "Add servers on the left, check them, then click an action." $script:cTextMuted
AppendLine "" $script:cTextMuted

[System.Windows.Forms.Application]::Run($script:form)
