#Requires -Version 5.1

<#
.SYNOPSIS
    Records low-overhead health metrics on a Windows PC to CSV logs, so crashes,
    slowdowns and freezes can be analysed after they happen.

.DESCRIPTION
    DesktopMonitor takes a sample every few seconds and appends one row per sample to
    a daily file in -LogDirectory:

      metrics_<PC>_<yyyyMMdd>.csv    One row per sample: CPU (also the busiest core),
                                     memory and kernel pool, disk, network, Wi-Fi, ping,
                                     connectivity (DNS, domain controller, chosen servers,
                                     web access through the proxy, mapped drives, VPN,
                                     Windows' own view of the network), GPU, temperature,
                                     power, user session, responsiveness (stalls, slow apps
                                     and taskbar), security and monitoring agents' load and
                                     file access time, top processes, and the monitor's own
                                     overhead.
      processes_<PC>_<yyyyMMdd>.csv  Top processes by CPU, I/O and hard page faults (each
                                     sample, when busy); by memory, handles, GDI and USER
                                     objects, TCP and UDP (once a minute); 32-bit processes
                                     with the least free address space and security /
                                     monitoring agents (once a minute); every process start
                                     and exit. Service host processes (svchost.exe) list
                                     the services they run.
      events_<PC>_<yyyyMMdd>.csv     Notable Windows event log entries with a Tag column:
                                     BlueScreen, UnexpectedReboot, UnexpectedShutdown,
                                     AppCrash, AppHang, ErrorReport, GpuDriverReset,
                                     GpuError, DiskError, HardwareError, LowMemory,
                                     CpuThrottled, ServiceCrash, PortExhaustion, Boot, Shutdown,
                                     RestartRequested, Sleep, Wake, Critical, Error,
                                     Warning, plus MonitorStart / MonitorStop rows and
                                     UserReportedFreeze rows from -Mark.
      inventory_<PC>_<yyyyMMdd-HHmmss>.csv
                                     A one-off hardware, OS and driver snapshot each time
                                     the monitor starts.
      monitor_<PC>_<user>.log        The monitor's own diagnostic messages.
      DesktopMonitor_<PC>_<yyyyMMdd>.zip
                                     The files above (except the .log) for one finished
                                     day, zipped. The report reads them as they are.

    Every file is CSV (or TSV) with a header row, so it opens in Excel, Power BI, pandas
    or Import-Csv. The metrics columns are the same on every PC: a group that is excluded
    or unsupported is left blank rather than removed. Files roll over daily and when they
    reach -MaxFileSizeMB. After midnight the finished day is zipped (to about a fifth of
    the size). Days are deleted after -RetentionDays, or sooner, oldest first, when this
    PC's logs would take more than -MaxTotalSizeMB.

    What to look for after an incident:
      Freezes    MaxStallMs of 1000 or more, or a few StallCount (the whole PC stopped);
                 ForegroundResponseMs in the thousands (that app stopped responding,
                 named in ForegroundResponseApp); ShellResponseMs in the thousands (the
                 taskbar / Explorer stopped responding); GapSec much bigger than the
                 interval (the monitor itself was stalled), CollectMs in the thousands,
                 HungAppCount above 0; UserReportedFreeze rows from -Mark.
      Stutter    CpuMaxCoreDpcPct above ~20: a driver is keeping one CPU core busy with
                 interrupts (mouse lag, audio crackle), which the whole-PC CpuDpcPct
                 averages away.
      Slowness   CpuPct near 100 with a high CpuQueue; CpuPerfPct well below 100 or
                 ThermalLimitPct below 100 (throttling); CommitPct above 90 or high
                 PageInPerSec (paging to disk; TopHardFaultProcess is the program
                 waiting on it); DiskBusyPct near 100 with DiskReadMs or DiskWriteMs above
                 ~50; GpuPct near 100 (TopGpuProcess); PoolNonPagedMB or Handles growing
                 for hours (leaks); Ping failures or high TcpRetransPerSec (network).
      Resources  TopGdiObjects or TopUserObjects near 10,000 (the per-process limit:
                 windows stop drawing, menus fail, the app hangs or crashes), GdiObjects
                 near 65,536 for the whole session, TopHandles in the hundreds of
                 thousands (a handle leak), TopPagedPoolMB or TopNonPagedPoolMB in the
                 hundreds (a driver or handle leak, named by its pool tag), TcpPortsInUse
                 or UdpEndpoints in the tens of thousands with TopTcpProcess or
                 TopUdpProcess (port exhaustion: connections start failing),
                 LargestFreeBlockMB in the low hundreds for Top32BitProcess (that 32-bit
                 program is running out of address space: "out of memory" errors, blank
                 windows or a crash, however much RAM is free).
      Network    DnsLookupMs in the hundreds or DnsFailed (name lookups slow or failing);
                 DcConnectMs high or DcFailed, DomainController in another site (sign-in,
                 Group Policy and file shares wait on it); NetworkCategory not Domain on the
                 office network (wrong firewall profile); HttpFailed or HttpMs in the
                 thousands (proxy, PAC/WPAD or the internet link); MappedDrivesDown;
                 ServerFailed; ConnectivityErrors says why. Event tags DnsTimeout,
                 NoDomainController, SecureChannel, GroupPolicyFailed, Kerberos, TimeSync,
                 NetworkAdapter, WifiDisconnected, NetworkDisconnected, IpAddressConflict.
      Agents     AgentCpuPct high (antivirus, EDR, DLP, monitoring or management agents
                 busy; TopAgentProcess names it), FileCheckMs in the hundreds (their file
                 filters slowing every program that opens files), DefenderScanStarted
                 events. The inventory lists the security products, file-system filter
                 drivers and the DLLs they inject into programs.
      Crashes    BlueScreen / UnexpectedReboot / AppCrash tags in the events file, and a
                 MonitorStart row saying the PC restarted since the last sample.

    Freezes shorter than the interval are caught by two kinds of background thread:
      * A stall watch wakes every 100 ms at the highest priority and records how late it
        was (MaxStallMs, StallCount). A busy CPU doesn't delay it; being late means
        nothing could run at all: a driver, firmware or hardware problem.
      * Probes send a no-op message (WM_NULL) every second to the foreground window and
        to the taskbar, and time the reply (ForegroundResponseMs, ShellResponseMs): how
        long that app, or Explorer, would have ignored a click or a key.

    Low overhead: performance counters are opened once and re-read, the sampling loop
    uses in-process .NET and Win32 calls (no WMI), slow checks run less often, the
    process runs at BelowNormal priority, and each sample is one small append. The
    MonitorCpuPct and MonitorMemMB columns record the monitor's own cost.

    Each write goes straight to disk (write-through), so everything up to the last
    sample survives a blue screen or power cut.

.PARAMETER IntervalSeconds
    Seconds between samples. Default 10.

.PARAMETER LogDirectory
    Folder for the log files. Default C:\ProgramData\DesktopMonitor\Logs.

.PARAMETER Format
    Csv (comma-separated, the default) or Tsv (tab-separated).

.PARAMETER DurationMinutes
    Stop after this many minutes. 0 (the default) runs until stopped.

.PARAMETER RetentionDays
    Delete logs from more than this many days ago. Default 30.

.PARAMETER MaxFileSizeMB
    Start a new file (name_1.csv, name_2.csv, ...) when a log reaches this size. Default 50.

.PARAMETER MaxTotalSizeMB
    The most space this PC's logs may take in -LogDirectory, zips included. Past it, the
    oldest days are deleted first. Today's files are never deleted. Checked every hour.
    Default 500.

.PARAMETER NoZip
    Leave finished days as plain CSV files instead of zipping them.

.PARAMETER TopProcessCount
    How many top processes to record for each ranking. 0 turns off the process tracking.
    Default 10.

.PARAMETER Exclude
    Metric groups to skip: Cpu, Memory, Disk, Network, Wifi, Ping, Connectivity, Gpu,
    Thermal, Power, Session, Responsiveness, Security, Processes, Events. Their columns stay
    in the file but are left blank.

.PARAMETER PingTarget
    One or two hosts to ping each sample (columns Ping1Ms and Ping2Ms). 'gateway' means
    the current default gateway. Default: gateway. Example: -PingTarget gateway,8.8.8.8

.PARAMETER PingTimeoutMs
    Ping timeout in milliseconds. Default 1000.

.PARAMETER DnsName
    The name looked up each sample to time the DNS servers (bypassing the PC's DNS cache).
    Default: on a domain-joined PC, the records the PC uses to find a domain controller
    (_ldap._tcp.dc._msdcs.<domain>); otherwise www.msftconnecttest.com.

.PARAMETER ServerTarget
    One or two servers to time a connection to each sample, as host:port, for example a
    file server (fileserver01:445), a proxy (proxy.corp.local:8080) or a business
    application's server. Many networks block ping to servers; a connection still works.

.PARAMETER HttpTarget
    A web address fetched once a minute through the user's proxy settings. Default
    http://www.msftconnecttest.com/connecttest.txt (the page Windows itself uses to check
    for internet access). Use an intranet address to test that instead, or 'none'.

.PARAMETER AgentName
    Extra process names (wildcards allowed) to count as security or monitoring agents, on top
    of the ones recognised by their publisher. Example: -AgentName 'acme*.exe','tracker.exe'

.PARAMETER EventLookbackHours
    On the first start, how many hours of the event log to read. Later starts carry on
    from the last sample in the existing logs (going back at most -RetentionDays), so
    nothing logged while the monitor was stopped is missed. Default 24.

.PARAMETER Priority
    CPU priority of the monitor: Idle, BelowNormal (the default) or Normal.

.PARAMETER Live
    Also print a one-line summary of each sample to the console.

.PARAMETER InventoryOnly
    Write the hardware/OS inventory file to -LogDirectory and exit. (The monitor runs this
    in a separate short-lived process at start-up, so the modules it loads don't stay in
    the monitor's memory.)

.PARAMETER Install
    Make the monitor run permanently for the current user, and start it. No admin rights
    are needed. It:
      * copies this script to -InstallDirectory;
      * creates a scheduled task named DesktopMonitor-<username> that starts the monitor
        hidden at logon, checks every 5 minutes and restarts it if it has stopped;
      * if the organisation stops users creating scheduled tasks, adds the monitor to the
        user's startup programs (the HKCU Run key) instead. That starts it at each logon,
        but nothing restarts it if it stops.
    It runs only while that user is logged on.

    The other parameters given with -Install become the monitor's settings. Installing
    again replaces them all: a setting left out goes back to its default. Installing
    again with the same settings and the same script changes nothing (it only starts the
    monitor if it isn't running), so it is safe to run from a logon script.

.PARAMETER InstallDirectory
    Used with -Install and -Uninstall. Where the script is copied.
    Default %LOCALAPPDATA%\DesktopMonitor.

.PARAMETER FreezeShortcut
    Used with -Install. Also put a 'Report a freeze' shortcut on the user's desktop. It
    runs -Mark with no window and confirms with a small popup. Installing again without
    -FreezeShortcut removes it, like any other setting.

.PARAMETER Uninstall
    Stop the installed monitor and remove its scheduled task (or startup entry), the
    installed copy of the script and the 'Report a freeze' shortcut. The logs are kept.

.PARAMETER Mark
    Record that the PC froze or was slow just now, as a UserReportedFreeze row in the
    events log, so the logs can be checked at exactly that time. Give the same
    -LogDirectory as the monitor if it isn't the default. If the monitor is running, it
    adds the row within one interval; if not, the row is written straight away.

.PARAMETER Note
    Used with -Mark. A few words about what happened, saved with the row.

.PARAMETER Popup
    Used with -Mark. Confirm in a small window that closes itself, instead of only in the
    console. The 'Report a freeze' shortcut uses this.

.EXAMPLE
    .\DesktopMonitor.ps1 -Live
    Runs in this console with the defaults and prints each sample. Ctrl+C stops it.

.EXAMPLE
    .\DesktopMonitor.ps1 -IntervalSeconds 5 -DurationMinutes 120 -LogDirectory D:\Diag
    Samples every 5 seconds for two hours, logging to D:\Diag.

.EXAMPLE
    .\DesktopMonitor.ps1 -Install -IntervalSeconds 15 -PingTarget gateway,1.1.1.1
    Runs permanently for the current user, sampling every 15 seconds and pinging the
    router and 1.1.1.1.

.EXAMPLE
    .\DesktopMonitor.ps1 -Install -IntervalSeconds 15 -PingTarget gateway,1.1.1.1 -Exclude Gpu
    Changes the settings of an installed monitor. The whole set is given again, because
    installing replaces the settings rather than adding to them.

.EXAMPLE
    .\DesktopMonitor.ps1 -Install -PingTarget gateway,8.8.8.8 -FreezeShortcut
    Runs permanently, and adds a 'Report a freeze' shortcut to the user's desktop.

.EXAMPLE
    .\DesktopMonitor.ps1 -Install -FreezeShortcut -ServerTarget fileserver01:445,proxy.corp.local:8080 -HttpTarget https://intranet.corp.local/
    Also times a connection to the file server and the proxy every sample, and fetches the
    intranet home page once a minute through the user's proxy settings.

.EXAMPLE
    .\DesktopMonitor.ps1 -Mark -Note 'Outlook froze for about 10 seconds'
    Records a freeze report at the current time.

.EXAMPLE
    .\DesktopMonitor.ps1 -Uninstall
#>
[CmdletBinding()]
param(
    [ValidateRange(1, 3600)] [int] $IntervalSeconds = 10,
    [ValidateNotNullOrEmpty()] [string] $LogDirectory = (Join-Path $env:ProgramData 'DesktopMonitor\Logs'),
    [ValidateSet('Csv', 'Tsv')] [string] $Format = 'Csv',
    [ValidateRange(0, 5256000)] [int] $DurationMinutes = 0,
    [ValidateRange(1, 3650)] [int] $RetentionDays = 30,
    [ValidateRange(1, 4096)] [int] $MaxFileSizeMB = 50,
    [ValidateRange(10, 1048576)] [int] $MaxTotalSizeMB = 500,
    [switch] $NoZip,
    [ValidateRange(0, 50)] [int] $TopProcessCount = 10,
    [ValidateSet('Cpu', 'Memory', 'Disk', 'Network', 'Wifi', 'Ping', 'Connectivity', 'Gpu', 'Thermal', 'Power', 'Session', 'Responsiveness', 'Security', 'Processes', 'Events')]
    [string[]] $Exclude = @(),
    [ValidateCount(1, 2)] [string[]] $PingTarget = @('gateway'),
    [ValidateRange(100, 5000)] [int] $PingTimeoutMs = 1000,
    [string] $DnsName,
    [ValidateCount(0, 2)] [string[]] $ServerTarget = @(),
    [string] $HttpTarget = 'http://www.msftconnecttest.com/connecttest.txt',
    [string[]] $AgentName = @(),
    [ValidateRange(0, 8760)] [int] $EventLookbackHours = 24,
    [ValidateSet('Idle', 'BelowNormal', 'Normal')] [string] $Priority = 'BelowNormal',
    [switch] $Live,
    [switch] $InventoryOnly,
    [switch] $Install,
    [ValidateNotNullOrEmpty()] [string] $InstallDirectory = (Join-Path $env:LOCALAPPDATA 'DesktopMonitor'),
    [switch] $FreezeShortcut,
    [switch] $Uninstall,
    [switch] $Mark,
    [string] $Note,
    [switch] $Popup
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# AppLocker / WDAC policies can force PowerShell into Constrained Language mode, which
# blocks the .NET and Win32 calls this script needs. Allowing the script in the policy
# (for example by signing it) runs it in Full Language mode again.
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    throw "PowerShell is running in $($ExecutionContext.SessionState.LanguageMode) mode (an AppLocker or WDAC policy), which blocks what DesktopMonitor needs. Ask IT to allow this script, for example by signing it with a trusted certificate."
}

# ---------------------------------------------------------------------------------------
# Settings and column layout
# ---------------------------------------------------------------------------------------

$script:Version    = '1.3.0'
$script:TaskName   = "DesktopMonitor-$env:USERNAME"
$script:RunKey     = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$script:Inv       = [Globalization.CultureInfo]::InvariantCulture
$script:Sep        = if ($Format -eq 'Tsv') { "`t" } else { ',' }
$script:Ext        = $Format.ToLowerInvariant()
$script:PC         = $env:COMPUTERNAME
$script:Cores      = [Environment]::ProcessorCount
$script:Self       = [Diagnostics.Process]::GetCurrentProcess()
$script:Utf8       = New-Object Text.UTF8Encoding($false)
$script:DiagLog    = $null
$script:WarnedAt   = @{}
$script:SampleNo   = 0
$script:Clock      = [Diagnostics.Stopwatch]::StartNew()
# Slower-changing things (Wi-Fi signal, top memory users, logged-on user) are read once a
# minute, whatever the interval.
$script:EveryMinute = [int][math]::Max(1, [math]::Round(60 / $IntervalSeconds))

$script:MetricColumns = @(
    'Timestamp', 'GapSec', 'CollectMs', 'UptimeHours',
    'CpuPct', 'CpuPrivilegedPct', 'CpuInterruptPct', 'CpuDpcPct', 'CpuMaxCorePct', 'CpuMaxCoreDpcPct', 'CpuPerfPct', 'CpuMHz',
    'CpuQueue', 'ContextSwitchesPerSec', 'Processes', 'Threads', 'Handles', 'GdiObjects', 'UserObjects',
    'MemUsedPct', 'MemAvailableMB', 'MemModifiedMB', 'CommitPct', 'CommitGB', 'CommitLimitGB',
    'PageInPerSec', 'PageOutPerSec', 'PageFileUsedPct', 'PoolNonPagedMB', 'PoolPagedMB',
    'TopNonPagedPoolTag', 'TopNonPagedPoolMB', 'TopPagedPoolTag', 'TopPagedPoolMB',
    'DiskBusyPct', 'DiskQueue', 'DiskReadMs', 'DiskWriteMs', 'DiskReadMBps', 'DiskWriteMBps',
    'DiskIops', 'SysDriveFreeGB', 'SysDriveFreePct', 'OtherDrivesFree',
    'NetRecvKBps', 'NetSentKBps', 'NetErrors', 'TcpConnections', 'TcpRetransPerSec', 'TcpPortsInUse', 'TcpTimeWait', 'UdpEndpoints',
    'NetAdapters', 'WifiSignalPct', 'Ping1Ms', 'Ping2Ms', 'PingFailed',
    'DnsLookupMs', 'DnsFailed', 'DomainController', 'DcConnectMs', 'DcFailed', 'Server1Ms', 'Server2Ms', 'ServerFailed',
    'HttpMs', 'HttpFailed', 'MappedDrivesDown', 'NetworkCategory', 'WindowsConnectivity', 'Vpn', 'DnsServers', 'NetChanges',
    'ConnectivityErrors',
    'GpuPct', 'Gpu3DPct', 'GpuDedicatedMB', 'GpuSharedMB', 'TopGpuProcess', 'TopGpuPct',
    'ThermalMaxC', 'ThermalLimitPct',
    'OnAC', 'BatteryPct', 'BatterySaver',
    'ConsoleUser', 'IdleSec', 'ForegroundApp', 'HungAppCount', 'HungApps',
    'MaxStallMs', 'StallCount', 'ForegroundResponseMs', 'ForegroundResponseApp', 'ShellResponseMs',
    'AgentCpuPct', 'AgentIoKBps', 'AgentMemMB', 'TopAgentProcess', 'TopAgentCpuPct', 'FileCheckMs', 'FileCheckFailed',
    'TopCpuProcess', 'TopCpuPct', 'TopMemProcess', 'TopMemMB', 'TopIoProcess', 'TopIoKBps',
    'TopHardFaultProcess', 'TopHardFaultsPerSec',
    'TopHandleProcess', 'TopHandles', 'TopGdiProcess', 'TopGdiObjects', 'TopUserObjProcess', 'TopUserObjects',
    'TopTcpProcess', 'TopTcpConnections', 'TopUdpProcess', 'TopUdpEndpoints',
    'Top32BitProcess', 'LargestFreeBlockMB', 'FreeAddressSpaceMB',
    'EventCount', 'EventTags',
    'MonitorCpuPct', 'MonitorMemMB'
)
$script:ProcessColumns = @(
    'Timestamp', 'Kind', 'Rank', 'Process', 'PID', 'ParentPID', 'CpuPct', 'IoKBps', 'HardFaultsPerSec',
    'PrivateMB', 'WorkingSetMB', 'Handles', 'GdiObjects', 'UserObjects', 'TcpConnections', 'UdpEndpoints', 'Threads',
    'LargestFreeBlockMB', 'FreeAddressSpaceMB', 'SessionId', 'StartedAt', 'Services', 'Agent'
)

# Publishers whose programs count as security, monitoring, management, VPN or web-filtering
# agents (matched against the CompanyName of a program file), names of Microsoft ones
# (whose publisher alone says nothing), and programs from those publishers that aren't
# agents. Extend with -AgentName.
$script:AgentCompanies = 'CrowdStrike|SentinelOne|Sentinel Labs|Sophos|Symantec|Broadcom|McAfee|Trellix|Musarubra|Trend Micro|' +
    'Palo Alto Networks|Cylance|BlackBerry|ESET|Kaspersky|Bitdefender|Carbon Black|Cybereason|Elastic|Malwarebytes|' +
    'WatchGuard|Panda Security|F-Secure|WithSecure|Check Point|Fortinet|Forcepoint|Digital Guardian|Fidelis|Zscaler|' +
    'Netskope|Cisco|Tanium|Qualys|Rapid7|Tenable|Arctic Wolf|Huntress|Webroot|Avast|AVG Technologies|Avira|NortonLifeLock|' +
    'Gen Digital|Comodo|Cynet|Deep Instinct|Secureworks|Darktrace|Absolute Software|Ivanti|BeyondTrust|CyberArk|' +
    'Thycotic|Delinea|Varonis|Code42|Proofpoint|Mimecast|Teramind|ActivTrak|Birch Grove|Veriato|Awareness Technologies|' +
    'Hubstaff|Nexthink|Lakeside Software|ControlUp|Netwrix|ManageEngine|Zoho Corporation|Kaseya|ConnectWise|N-able|' +
    'SolarWinds|Datto|NinjaOne|NinjaRMM|Splunk|Wazuh|Lansweeper|Menlo Security|iboss|Lookout|Skyhigh|Automox|Action1|' +
    'ThreatLocker|Airlock Digital|Morphisec|Bromium|HP Wolf|Imprivata|Okta|Duo Security|Akamai|Illumio|Guardicore|' +
    'Island Technology|Securly|GoGuardian|Lightspeed Systems|Smoothwall|Cloudflare'
$script:AgentNames = '^(MsMpEng|MpDefenderCoreService|NisSrv|MpCmdRun|MsSense|Sense[A-Za-z]*|SecurityHealthService|' +
    'smartscreen|CcmExec|CcmRepair|IntuneManagementExtension|AgentExecutor|Microsoft\.Management\.Services\.IntuneWindowsAgent)\.exe$'
$script:AgentNotAgents = 'Webex|Jabber|Collaboration|Meeting|Bluetooth|Wireless|WLAN|Ethernet'
$script:EventColumns = @('Timestamp', 'Tag', 'Level', 'Log', 'Provider', 'EventId', 'Message')

# ---------------------------------------------------------------------------------------
# Helpers: formatting, file writing, diagnostics
# ---------------------------------------------------------------------------------------

function Test-Group([string] $Name) { $Exclude -notcontains $Name }

# True on the 1st sample and then every $Samples samples.
function Test-Every([int] $Samples) { (($script:SampleNo - 1) % $Samples) -eq 0 }

function ConvertTo-Field($Value) {
    if ($null -eq $Value) { return '' }
    if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal]) {
        $d = [double] $Value
        if ([double]::IsNaN($d) -or [double]::IsInfinity($d)) { return '' }
        return $d.ToString('0.##', $script:Inv)
    }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm:ss', $script:Inv) }
    if ($Value -is [bool]) { if ($Value) { return '1' } else { return '0' } }
    $s = ([string] $Value) -replace '[\r\n\t]+', ' '
    # Text from outside (event messages, process names) that starts like a formula would
    # run as one when the file is opened in Excel. Any program can write such an event.
    if ($Value -is [string] -and $s -match '^[=+\-@]') { $s = "'" + $s }
    if ($s.IndexOf($script:Sep[0]) -ge 0 -or $s.IndexOf('"') -ge 0 -or $s -ne $s.Trim()) {
        $s = '"' + $s.Replace('"', '""') + '"'
    }
    return $s
}

# Appends lines with write-through so they reach the disk before a crash can lose them.
# Other programs (Excel, Notepad) can read the file meanwhile. Returns the new file size.
function Add-FileLines([string] $Path, [string[]] $Lines, [switch] $Bom) {
    $bytes = $script:Utf8.GetBytes([string]::Join("`r`n", $Lines) + "`r`n")
    $fs = New-Object IO.FileStream($Path, [IO.FileMode]::Append, [IO.FileAccess]::Write,
        [IO.FileShare]'ReadWrite, Delete', 4096, [IO.FileOptions]::WriteThrough)
    try {
        # A byte-order mark makes Excel read the file as UTF-8.
        if ($Bom -and $fs.Length -eq 0) { $fs.Write([byte[]] (0xEF, 0xBB, 0xBF), 0, 3) }
        $fs.Write($bytes, 0, $bytes.Length)
        return $fs.Length
    } finally { $fs.Dispose() }
}

function Write-MonitorLog([string] $Message, [string] $Level = 'INFO') {
    $line = '{0} {1,-5} {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', $script:Inv), $Level, $Message
    if ($Level -eq 'INFO') { Write-Host $line } else { Write-Host $line -ForegroundColor Yellow }
    if (-not $script:DiagLog) { return }
    try {
        $file = Get-Item -LiteralPath $script:DiagLog -ErrorAction SilentlyContinue
        if ($file -and $file.Length -gt 5MB) {
            Move-Item -LiteralPath $script:DiagLog -Destination "$($script:DiagLog).old" -Force
        }
        [void] (Add-FileLines $script:DiagLog @($line))
    } catch { }
}

# Logs a repeating problem at most once an hour.
function Write-Throttled([string] $Key, [string] $Message) {
    $now = Get-Date
    if (-not $script:WarnedAt.ContainsKey($Key) -or ($now - $script:WarnedAt[$Key]).TotalMinutes -ge 60) {
        $script:WarnedAt[$Key] = $now
        Write-MonitorLog $Message 'WARN'
    }
}

function New-LogTable([string] $Name, [string[]] $Columns) {
    [pscustomobject] @{
        Name      = $Name
        Columns   = $Columns
        Header    = [string]::Join($script:Sep, [string[]] @($Columns | ForEach-Object { ConvertTo-Field $_ }))
        Day       = $null
        Path      = $null
        FailCount = 0
        Pending   = New-Object 'System.Collections.Generic.List[string]'
    }
}

# Today's file for a table. Moves on to name_1, name_2, ... when a file is full or has a
# different header (for example from an older version of this script).
function Select-TableFile($Table) {
    $base = Join-Path $LogDirectory ('{0}_{1}_{2}' -f $Table.Name, $script:PC, $Table.Day)
    for ($n = 0; $n -lt 1000; $n++) {
        $path = if ($n -eq 0) { "$base.$($script:Ext)" } else { "${base}_$n.$($script:Ext)" }
        $file = Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
        if (-not $file -or $file.Length -eq 0) { return $path }
        if ($file.Length -ge $MaxFileSizeMB * 1MB) { continue }
        $first = $null
        try {
            # Opening for writing also skips files we can't add to: another user's file on a
            # shared PC, or one that Excel has locked.
            $fs = New-Object IO.FileStream($path, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]'ReadWrite, Delete')
            try { $first = (New-Object IO.StreamReader($fs, $script:Utf8, $true)).ReadLine() } finally { $fs.Dispose() }
        } catch { continue }
        if ($first -eq $Table.Header) { return $path }
    }
    throw "No free file name for $base"
}

# Rows are hashtables keyed by column name. If the file can't be written (say Excel has
# locked it), rows are kept in memory and the next attempt picks a file again, which moves
# on to name_1 (and so on) if the problem persists.
function Write-TableRows($Table, [object[]] $Rows) {
    foreach ($r in $Rows) {
        $fields = foreach ($c in $Table.Columns) { ConvertTo-Field $r[$c] }
        $Table.Pending.Add([string]::Join($script:Sep, [string[]] $fields))
    }
    if ($Table.Pending.Count -eq 0) { return }
    try {
        $day = (Get-Date).ToString('yyyyMMdd', $script:Inv)
        if ($Table.Day -ne $day) { $Table.Day = $day; $Table.Path = $null }
        if (-not $Table.Path) { $Table.Path = Select-TableFile $Table }
        $lines = $Table.Pending.ToArray()
        $isNew = -not (Test-Path -LiteralPath $Table.Path)
        if (-not $isNew) { $isNew = ((Get-Item -LiteralPath $Table.Path).Length -eq 0) }
        if ($isNew) { $lines = @($Table.Header) + $lines }
        $size = Add-FileLines $Table.Path $lines -Bom
        $Table.Pending.Clear()
        if ($Table.FailCount -gt 0) {
            Write-MonitorLog "Writing to $($Table.Path) works again."
            $Table.FailCount = 0
        }
        if ($size -ge $MaxFileSizeMB * 1MB) { $Table.Path = $null }
    } catch {
        $Table.FailCount++
        $Table.Path = $null
        if ($Table.FailCount -eq 1) {
            Write-MonitorLog "Cannot write the $($Table.Name) log ($($_.Exception.Message)). Keeping rows in memory and retrying." 'WARN'
        }
        $over = $Table.Pending.Count - 20000
        if ($over -gt 0) { $Table.Pending.RemoveRange(0, $over) }
    }
}

# Zipping finished days, -RetentionDays and -MaxTotalSizeMB, on a background thread (see
# [DesktopMonitor.Housekeeping]). Its messages go to the diagnostic log at the next sample.
function Start-LogHousekeeping {
    [void] [DesktopMonitor.Housekeeping]::Start($LogDirectory, $script:PC, (-not $NoZip), $RetentionDays, [long] $MaxTotalSizeMB * 1MB)
}

function Write-HousekeepingMessages {
    foreach ($m in [DesktopMonitor.Housekeeping]::TakeMessages()) {
        $level, $text = $m -split ' ', 2
        Write-MonitorLog $text $level
    }
}

# ---------------------------------------------------------------------------------------
# Win32 calls that have no cheap PowerShell/.NET equivalent
# ---------------------------------------------------------------------------------------

$script:NativeSource = @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.Win32;

namespace DesktopMonitor {

public class ProcInfo {
    public int Id, ParentId, Handles, Threads, SessionId, TcpConnections, UdpEndpoints;
    public int GdiObjects = -1, UserObjects = -1;   // -1 = not readable / not in the desktop session
    public string Name, Services, ServicesShort;     // Services: what a service host runs, or null
    public string Agent;                             // the security / monitoring product it belongs to, or null
    public long CreateTime, IoBytes, HardFaults;
    public double CpuSeconds, CpuPct, IoKBps, HardFaultsPerSec, PrivateMB, WorkingSetMB;
    public double LargestFreeBlockMB = -1, FreeAddressSpaceMB = -1;   // -1 = not a 32-bit process, or not measured

    // The name, plus the services for svchost.exe, which on its own says nothing.
    public string Label() {
        if (ServicesShort == null || !string.Equals(Name, "svchost.exe", StringComparison.OrdinalIgnoreCase)) return Name;
        return Name + " (" + ServicesShort + ")";
    }
}

public class ProcSample {
    public List<ProcInfo> All = new List<ProcInfo>();
    public List<ProcInfo> Started = new List<ProcInfo>();
    public List<ProcInfo> Exited = new List<ProcInfo>();
    public long TotalHandles;
    public int GdiTotal = -1, UserTotal = -1, TcpTotal, TcpTimeWait, UdpTotal;
    public bool HasDelta;
}

public class PoolTop {
    public string PagedTag, NonPagedTag;
    public double PagedMB, NonPagedMB;
}

public static class Native {
    [StructLayout(LayoutKind.Sequential)]
    struct LASTINPUTINFO { public uint cbSize; public uint dwTime; }

    [StructLayout(LayoutKind.Sequential)]
    class MEMORYSTATUSEX {
        public uint dwLength = 64; public uint dwMemoryLoad;
        public ulong ullTotalPhys, ullAvailPhys, ullTotalPageFile, ullAvailPageFile,
                     ullTotalVirtual, ullAvailVirtual, ullAvailExtendedVirtual;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct SYSTEM_POWER_STATUS {
        public byte ACLineStatus, BatteryFlag, BatteryLifePercent, SystemStatusFlag;
        public int BatteryLifeTime, BatteryFullLifeTime;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct UNICODE_STRING { public ushort Length, MaximumLength; public IntPtr Buffer; }

    // SYSTEM_PROCESS_INFORMATION. Natural alignment gives the right layout on x86 and x64.
    [StructLayout(LayoutKind.Sequential)]
    struct SPI {
        public uint NextEntryOffset, NumberOfThreads;
        public long WorkingSetPrivateSize;
        public uint HardFaultCount, NumberOfThreadsHighWatermark;
        public ulong CycleTime;
        public long CreateTime, UserTime, KernelTime;
        public UNICODE_STRING ImageName;
        public int BasePriority;
        public IntPtr UniqueProcessId, InheritedFromUniqueProcessId;
        public uint HandleCount, SessionId;
        public UIntPtr UniqueProcessKey, PeakVirtualSize, VirtualSize;
        public uint PageFaultCount;
        public UIntPtr PeakWorkingSetSize, WorkingSetSize, QuotaPeakPagedPoolUsage,
                       QuotaPagedPoolUsage, QuotaPeakNonPagedPoolUsage, QuotaNonPagedPoolUsage,
                       PagefileUsage, PeakPagefileUsage, PrivatePageCount;
        public long ReadOperationCount, WriteOperationCount, OtherOperationCount,
                    ReadTransferCount, WriteTransferCount, OtherTransferCount;
    }

    // Natural alignment gives the right layout for both 32-bit and 64-bit callers.
    [StructLayout(LayoutKind.Sequential)]
    struct MEMORY_BASIC_INFORMATION {
        public IntPtr BaseAddress, AllocationBase;
        public uint AllocationProtect;
        public UIntPtr RegionSize;
        public uint State, Protect, Type;
    }

    delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")] static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
    [DllImport("user32.dll")] static extern bool IsHungAppWindow(IntPtr hWnd);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] static extern IntPtr GetWindow(IntPtr hWnd, uint cmd);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lParam);
    [DllImport("kernel32.dll")] static extern bool GlobalMemoryStatusEx([In, Out] MEMORYSTATUSEX m);
    [DllImport("kernel32.dll")] static extern bool GetSystemPowerStatus(out SYSTEM_POWER_STATUS s);
    [DllImport("kernel32.dll")] public static extern ulong GetTickCount64();
    [DllImport("ntdll.dll")] static extern int NtQuerySystemInformation(int cls, IntPtr buf, int len, out int retLen);
    [DllImport("user32.dll")] static extern uint GetGuiResources(IntPtr hProcess, uint flags);
    [DllImport("kernel32.dll")] static extern IntPtr OpenProcess(uint access, bool inherit, int pid);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
    [DllImport("iphlpapi.dll")] static extern uint GetExtendedTcpTable(IntPtr table, ref int size, bool order, int af, int tableClass, uint reserved);
    [DllImport("iphlpapi.dll")] static extern uint GetExtendedUdpTable(IntPtr table, ref int size, bool order, int af, int tableClass, uint reserved);
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode)] static extern IntPtr OpenSCManager(string machine, string database, uint access);
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode)] static extern bool EnumServicesStatusEx(IntPtr scm, int level, uint type, uint state, IntPtr buf, int size, out int needed, out int count, ref int resume, string group);
    [DllImport("advapi32.dll")] static extern bool CloseServiceHandle(IntPtr h);
    [DllImport("kernel32.dll")] static extern UIntPtr VirtualQueryEx(IntPtr h, IntPtr address, out MEMORY_BASIC_INFORMATION info, UIntPtr length);
    [DllImport("kernel32.dll")] static extern bool IsWow64Process(IntPtr h, out bool wow64);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] static extern bool QueryFullProcessImageName(IntPtr h, int flags, StringBuilder name, ref int size);

    static readonly int ownSession = Process.GetCurrentProcess().SessionId;

    // Adds one address family's TCP connections (2 = IPv4, 23 = IPv6) to counts by owning
    // process; listening sockets are skipped. Each table is a 4-byte row count followed by
    // 24-byte (IPv4) or 56-byte (IPv6) rows holding the state and the owner's PID. False if
    // the table couldn't be read: it can grow quickly exactly when ports are running out,
    // and a silent undercount then would hide the problem.
    static bool CountTcp(int af, Dictionary<int, int> byPid, ref int total, ref int timeWait) {
        int size = 0;
        GetExtendedTcpTable(IntPtr.Zero, ref size, false, af, 5 /* TCP_TABLE_OWNER_PID_ALL */, 0);
        for (int attempt = 0; attempt < 5; attempt++) {
            size += size / 2 + 16384;
            IntPtr t = Marshal.AllocHGlobal(size);
            try {
                uint r = GetExtendedTcpTable(t, ref size, false, af, 5, 0);
                if (r == 122 /* ERROR_INSUFFICIENT_BUFFER: the table grew */) continue;
                if (r != 0) return false;
                int n = Marshal.ReadInt32(t);
                int rowSize = af == 2 ? 24 : 56, stateAt = af == 2 ? 0 : 48, pidAt = af == 2 ? 20 : 52;
                for (int i = 0; i < n; i++) {
                    int row = 4 + i * rowSize;
                    int state = Marshal.ReadInt32(t, row + stateAt);
                    if (state == 2 /* LISTEN */) continue;
                    int pid = Marshal.ReadInt32(t, row + pidAt);
                    total++;
                    if (state == 11 /* TIME_WAIT */) timeWait++;
                    int c; byPid.TryGetValue(pid, out c); byPid[pid] = c + 1;
                }
                return true;
            } finally { Marshal.FreeHGlobal(t); }
        }
        return false;
    }

    // Adds one address family's UDP endpoints to counts by owning process. Each row is 12
    // bytes (IPv4) or 28 bytes (IPv6), ending with the owner's PID. Every endpoint holds a
    // port, so a program that leaks them runs Windows out of UDP ports (DNS lookups fail).
    static bool CountUdp(int af, Dictionary<int, int> byPid, ref int total) {
        int size = 0;
        GetExtendedUdpTable(IntPtr.Zero, ref size, false, af, 1 /* UDP_TABLE_OWNER_PID */, 0);
        for (int attempt = 0; attempt < 5; attempt++) {
            size += size / 2 + 16384;
            IntPtr t = Marshal.AllocHGlobal(size);
            try {
                uint r = GetExtendedUdpTable(t, ref size, false, af, 1, 0);
                if (r == 122 /* ERROR_INSUFFICIENT_BUFFER */) continue;
                if (r != 0) return false;
                int n = Marshal.ReadInt32(t);
                int rowSize = af == 2 ? 12 : 28;
                for (int i = 0; i < n; i++) {
                    int pid = Marshal.ReadInt32(t, 4 + i * rowSize + rowSize - 4);
                    total++;
                    int c; byPid.TryGetValue(pid, out c); byPid[pid] = c + 1;
                }
                return true;
            } finally { Marshal.FreeHGlobal(t); }
        }
        return false;
    }

    // ---- Who is at the screen ----
    [DllImport("kernel32.dll")] static extern uint WTSGetActiveConsoleSessionId();
    [DllImport("wtsapi32.dll", CharSet = CharSet.Unicode)] static extern bool WTSQuerySessionInformation(IntPtr server, uint session, int infoClass, out IntPtr buf, out int bytes);
    [DllImport("wtsapi32.dll")] static extern void WTSFreeMemory(IntPtr p);

    static string SessionText(uint session, int infoClass) {
        IntPtr b; int n;
        if (!WTSQuerySessionInformation(IntPtr.Zero, session, infoClass, out b, out n)) return null;
        try { return Marshal.PtrToStringUni(b); } finally { WTSFreeMemory(b); }
    }

    // DOMAIN\user signed in at the PC's own screen (not over Remote Desktop), "" at the
    // sign-in screen, null if unknown. On a shared PC it can differ from the account the
    // monitor runs as.
    public static string ConsoleUser() {
        uint s = WTSGetActiveConsoleSessionId();
        if (s == 0xFFFFFFFF) return "";
        string user = SessionText(s, 5 /* WTSUserName */);
        if (user == null) return null;
        if (user.Length == 0) return "";
        string domain = SessionText(s, 7 /* WTSDomainName */);
        return string.IsNullOrEmpty(domain) ? user : domain + "\\" + user;
    }

    // Which services each process runs, refreshed once a minute (and when a new service
    // host starts). Many services share each svchost.exe, so its name alone says nothing.
    static Dictionary<int, string[]> services = new Dictionary<int, string[]>();
    static long servicesAt;

    static void RefreshServices() {
        servicesAt = Stopwatch.GetTimestamp();
        IntPtr scm = OpenSCManager(null, null, 0x0004 /* SC_MANAGER_ENUMERATE_SERVICE */);
        if (scm == IntPtr.Zero) return;
        try {
            int needed, count, resume = 0;
            EnumServicesStatusEx(scm, 0 /* SC_ENUM_PROCESS_INFO */, 0x30 /* SERVICE_WIN32 */, 1 /* ACTIVE */, IntPtr.Zero, 0, out needed, out count, ref resume, null);
            int size = needed + 16384;
            IntPtr buf = Marshal.AllocHGlobal(size);
            try {
                resume = 0;
                if (!EnumServicesStatusEx(scm, 0, 0x30, 1, buf, size, out needed, out count, ref resume, null)) return;
                // ENUM_SERVICE_STATUS_PROCESS: two string pointers, then SERVICE_STATUS_PROCESS
                // (nine 4-byte fields, the process id 8th), padded to pointer alignment.
                int entry = (IntPtr.Size * 2 + 36 + IntPtr.Size - 1) / IntPtr.Size * IntPtr.Size;
                Dictionary<int, List<string>> byPid = new Dictionary<int, List<string>>();
                for (int i = 0; i < count; i++) {
                    IntPtr e = new IntPtr(buf.ToInt64() + (long)i * entry);
                    int pid = Marshal.ReadInt32(e, IntPtr.Size * 2 + 28);
                    if (pid == 0) continue;
                    List<string> l;
                    if (!byPid.TryGetValue(pid, out l)) { l = new List<string>(); byPid[pid] = l; }
                    l.Add(Marshal.PtrToStringUni(Marshal.ReadIntPtr(e)));
                }
                Dictionary<int, string[]> map = new Dictionary<int, string[]>();
                foreach (KeyValuePair<int, List<string>> kv in byPid) {
                    kv.Value.Sort(StringComparer.OrdinalIgnoreCase);
                    string all = string.Join(", ", kv.Value.ToArray());
                    string shortList = kv.Value.Count <= 2 ? all : kv.Value[0] + ", " + kv.Value[1] + " +" + (kv.Value.Count - 2);
                    map[kv.Key] = new string[] { all, shortList };
                }
                services = map;
            } finally { Marshal.FreeHGlobal(buf); }
        } finally { CloseServiceHandle(scm); }
    }

    // ---- Per-core CPU ----
    // The whole-PC counters average over every core, so one core stuck in a driver's
    // interrupts (mouse lag, audio crackle) barely shows on a PC with 16 threads.
    static IntPtr coreBuf = IntPtr.Zero;
    static long[] corePrev;

    // Since the previous call: { busiest core's busy %, highest % of one core's time in
    // interrupts and DPCs }. null on the first call.
    public static double[] CoreLoad() {
        const int entry = 48;   // SYSTEM_PROCESSOR_PERFORMANCE_INFORMATION, same size on x86 and x64
        if (coreBuf == IntPtr.Zero) coreBuf = Marshal.AllocHGlobal(64 * entry);   // a processor group holds at most 64
        int ret;
        if (NtQuerySystemInformation(8 /* SystemProcessorPerformanceInformation */, coreBuf, 64 * entry, out ret) != 0) return null;
        int n = ret / entry;
        long[] now = new long[n * 5];
        for (int i = 0; i < n * 5; i++) now[i] = Marshal.ReadInt64(coreBuf, (i / 5) * entry + (i % 5) * 8);
        double[] r = null;
        if (corePrev != null && corePrev.Length == now.Length) {
            double busy = 0, dpc = 0;
            for (int i = 0; i < n; i++) {
                int b = i * 5;   // idle, kernel (includes idle, DPC and interrupt time), user, DPC, interrupt
                double idle = now[b] - corePrev[b], total = (now[b + 1] - corePrev[b + 1]) + (now[b + 2] - corePrev[b + 2]);
                if (total <= 0) continue;
                busy = Math.Max(busy, 100.0 * (total - idle) / total);
                dpc = Math.Max(dpc, 100.0 * ((now[b + 3] - corePrev[b + 3]) + (now[b + 4] - corePrev[b + 4])) / total);
            }
            r = new double[] { Math.Min(100, Math.Max(0, busy)), Math.Min(100, Math.Max(0, dpc)) };
        }
        corePrev = now;
        return r;
    }

    // ---- Kernel pool by tag (what poolmon shows) ----
    // Drivers label their kernel memory with a 4-letter tag. When PoolPagedMB or
    // PoolNonPagedMB grows, the biggest tag usually names the leaking driver, or the kind of
    // object being leaked (Toke = security tokens, for example).
    static IntPtr poolBuf = IntPtr.Zero;
    static int poolSize = 0;

    public static PoolTop TopPoolTags() {
        if (poolBuf == IntPtr.Zero) { poolSize = 1 << 18; poolBuf = Marshal.AllocHGlobal(poolSize); }
        int needed, status;
        while ((status = NtQuerySystemInformation(22 /* SystemPoolTagInformation */, poolBuf, poolSize, out needed)) == unchecked((int)0xC0000004)) {
            Marshal.FreeHGlobal(poolBuf); poolBuf = IntPtr.Zero;
            poolSize = Math.Max(poolSize * 2, needed + 65536);
            poolBuf = Marshal.AllocHGlobal(poolSize);
        }
        if (status != 0) return null;
        // A count, then SYSTEM_POOLTAG entries: tag, paged allocs/frees, paged bytes used,
        // non-paged allocs/frees, non-paged bytes used (the sizes are pointer-sized).
        bool x64 = IntPtr.Size == 8;
        int n = Marshal.ReadInt32(poolBuf), start = x64 ? 8 : 4, size = x64 ? 40 : 28;
        int pagedAt = x64 ? 16 : 12, nonPagedAt = x64 ? 32 : 24;
        PoolTop top = new PoolTop();
        long paged = -1, nonPaged = -1;
        for (int i = 0; i < n; i++) {
            int e = start + i * size;
            long p = x64 ? Marshal.ReadInt64(poolBuf, e + pagedAt) : (uint)Marshal.ReadInt32(poolBuf, e + pagedAt);
            long np = x64 ? Marshal.ReadInt64(poolBuf, e + nonPagedAt) : (uint)Marshal.ReadInt32(poolBuf, e + nonPagedAt);
            if (p > paged) { paged = p; top.PagedTag = TagAt(e); }
            if (np > nonPaged) { nonPaged = np; top.NonPagedTag = TagAt(e); }
        }
        top.PagedMB = Math.Max(0, paged) / 1048576.0;
        top.NonPagedMB = Math.Max(0, nonPaged) / 1048576.0;
        return top;
    }

    static string TagAt(int offset) {
        char[] c = new char[4];
        for (int i = 0; i < 4; i++) {
            byte b = Marshal.ReadByte(poolBuf, offset + i);
            c[i] = b >= 32 && b < 127 ? (char)b : '?';
        }
        return new string(c).TrimEnd();
    }

    // Seconds since the last keyboard or mouse input in this session.
    public static double IdleSeconds() {
        LASTINPUTINFO l = new LASTINPUTINFO();
        l.cbSize = (uint)Marshal.SizeOf(l);
        if (!GetLastInputInfo(ref l)) return -1;
        return unchecked((uint)Environment.TickCount - l.dwTime) / 1000.0;
    }

    public static int ForegroundPid() {
        IntPtr h = GetForegroundWindow();
        if (h == IntPtr.Zero) return 0;
        uint pid; GetWindowThreadProcessId(h, out pid);
        return (int)pid;
    }

    // Processes owning a visible top-level window that Windows considers "Not Responding".
    // IsHungAppWindow returns at once, unlike Process.Responding which waits 5 seconds.
    public static int[] HungWindowPids() {
        List<int> pids = new List<int>();
        EnumWindows(delegate(IntPtr h, IntPtr p) {
            if (IsWindowVisible(h) && GetWindow(h, 4 /* GW_OWNER */) == IntPtr.Zero && IsHungAppWindow(h)) {
                uint pid; GetWindowThreadProcessId(h, out pid);
                if (!pids.Contains((int)pid)) pids.Add((int)pid);
            }
            return true;
        }, IntPtr.Zero);
        return pids.ToArray();
    }

    // Total RAM, available RAM, commit limit, available commit (bytes).
    public static ulong[] Memory() {
        MEMORYSTATUSEX m = new MEMORYSTATUSEX();
        if (!GlobalMemoryStatusEx(m)) return null;
        return new ulong[] { m.ullTotalPhys, m.ullAvailPhys, m.ullTotalPageFile, m.ullAvailPageFile };
    }

    // AC line status (0 off, 1 on, 255 unknown), battery %, battery flag (128 = none), saver.
    public static int[] Power() {
        SYSTEM_POWER_STATUS s;
        if (!GetSystemPowerStatus(out s)) return null;
        return new int[] { s.ACLineStatus, s.BatteryLifePercent, s.BatteryFlag, s.SystemStatusFlag };
    }

    // Kept for the life of the process: re-allocating ~2 MB every sample costs more than
    // the query itself.
    static IntPtr buf = IntPtr.Zero;
    static int bufferSize = 0;

    // Every process with its CPU time, memory, handles and I/O, in one system call. Unlike
    // Get-Process this needs no process handles, so it also covers SYSTEM and protected
    // processes (antivirus, for example) when running as a standard user.
    public static List<ProcInfo> QueryProcesses() {
        if (buf == IntPtr.Zero) { bufferSize = 1 << 21; buf = Marshal.AllocHGlobal(bufferSize); }
        while (true) {
            int needed;
            int status = NtQuerySystemInformation(5, buf, bufferSize, out needed);
            if (status == 0) break;
            if (status != unchecked((int)0xC0000004)) throw new InvalidOperationException("NtQuerySystemInformation failed: 0x" + status.ToString("X8"));
            Marshal.FreeHGlobal(buf); buf = IntPtr.Zero;
            bufferSize = Math.Max(bufferSize * 2, needed + 262144);
            buf = Marshal.AllocHGlobal(bufferSize);
        }
        {
            List<ProcInfo> list = new List<ProcInfo>();
            long offset = 0;
            while (true) {
                SPI s = (SPI)Marshal.PtrToStructure(new IntPtr(buf.ToInt64() + offset), typeof(SPI));
                ProcInfo i = new ProcInfo();
                i.Id = (int)s.UniqueProcessId.ToInt64();
                i.ParentId = (int)s.InheritedFromUniqueProcessId.ToInt64();
                i.Name = s.ImageName.Buffer == IntPtr.Zero ? (i.Id == 0 ? "Idle" : "?")
                       : Marshal.PtrToStringUni(s.ImageName.Buffer, s.ImageName.Length / 2);
                i.CreateTime = s.CreateTime;
                i.CpuSeconds = (s.UserTime + s.KernelTime) / 1e7;
                i.PrivateMB = s.PagefileUsage.ToUInt64() / 1048576.0;
                i.WorkingSetMB = s.WorkingSetSize.ToUInt64() / 1048576.0;
                i.Handles = (int)s.HandleCount;
                i.Threads = (int)s.NumberOfThreads;
                i.SessionId = (int)s.SessionId;
                i.IoBytes = s.ReadTransferCount + s.WriteTransferCount + s.OtherTransferCount;
                i.HardFaults = s.HardFaultCount;
                list.Add(i);
                if (s.NextEntryOffset == 0) break;
                offset += s.NextEntryOffset;
            }
            return list;
        }
    }

    static Dictionary<string, ProcInfo> previous;
    static Dictionary<int, string> names = new Dictionary<int, string>();
    static long previousTicks;

    // Processes with CPU % (of the whole machine), I/O rate and hard page faults (reads from
    // disk the process had to wait for) since the previous call, plus the processes that
    // started or exited in between.
    public static ProcSample SampleProcesses() {
        List<ProcInfo> list = QueryProcesses();
        long now = Stopwatch.GetTimestamp();
        double dt = previous == null ? 0 : (now - previousTicks) / (double)Stopwatch.Frequency;
        Dictionary<string, ProcInfo> map = new Dictionary<string, ProcInfo>(list.Count);
        Dictionary<int, string> newNames = new Dictionary<int, string>(list.Count);
        ProcSample r = new ProcSample();
        r.HasDelta = dt > 0;
        int cores = Environment.ProcessorCount;
        foreach (ProcInfo p in list) {
            string key = p.Id + "|" + p.CreateTime;
            map[key] = p;
            newNames[p.Id] = p.Name;
            r.All.Add(p);
            r.TotalHandles += p.Handles;
            if (previous == null) continue;
            ProcInfo old;
            if (previous.TryGetValue(key, out old)) {
                if (dt > 0) {
                    p.CpuPct = Math.Max(0, (p.CpuSeconds - old.CpuSeconds) / dt / cores * 100);
                    p.IoKBps = Math.Max(0, (p.IoBytes - old.IoBytes) / 1024.0 / dt);
                    p.HardFaultsPerSec = Math.Max(0, (p.HardFaults - old.HardFaults) / dt);
                }
            } else {
                r.Started.Add(p);
            }
        }

        bool newHost = false;
        foreach (ProcInfo p in r.Started) if (string.Equals(p.Name, "svchost.exe", StringComparison.OrdinalIgnoreCase)) newHost = true;
        if (servicesAt == 0 || newHost || (now - servicesAt) > 60 * Stopwatch.Frequency) {
            try { RefreshServices(); } catch { }
        }
        foreach (ProcInfo p in list) {
            string[] s;
            if (p.Id != 0 && services.TryGetValue(p.Id, out s)) { p.Services = s[0]; p.ServicesShort = s[1]; }
        }
        if (agentCompany != null || agentName != null) {
            Dictionary<string, string> agents = new Dictionary<string, string>(list.Count);
            foreach (ProcInfo p in list) {
                string key = p.Id + "|" + p.CreateTime, a;
                bool known = true;
                if (!agentByKey.TryGetValue(key, out a)) a = ClassifyAgent(p, out known) ?? "";
                if (known) agents[key] = a;   // otherwise try again next time (its service may not be listed yet)
                if (a.Length > 0) p.Agent = a;
            }
            agentByKey = agents;
        }
        if (previous != null)
            foreach (KeyValuePair<string, ProcInfo> kv in previous)
                if (!map.ContainsKey(kv.Key)) r.Exited.Add(kv.Value);
        previous = map; previousTicks = now; names = newNames;

        // GDI and USER objects (each process may hold 10,000 of each by default; past that
        // its windows stop drawing or it crashes). Only processes in this desktop session
        // share those limits, so only they are counted.
        foreach (ProcInfo p in list) {
            if (p.SessionId != ownSession || p.Id == 0) continue;
            IntPtr h = OpenProcess(0x1000 /* PROCESS_QUERY_LIMITED_INFORMATION */, false, p.Id);
            if (h == IntPtr.Zero) continue;   // e.g. a program run as administrator
            p.GdiObjects = (int)GetGuiResources(h, 0);
            p.UserObjects = (int)GetGuiResources(h, 1);
            CloseHandle(h);
        }
        // Whole-session totals, which also include the processes we can't open.
        r.GdiTotal = (int)GetGuiResources(new IntPtr(-2) /* GR_GLOBAL */, 0);
        r.UserTotal = (int)GetGuiResources(new IntPtr(-2), 1);

        // TCP connections and UDP endpoints per process, to find what is using up the ports.
        // A total of -1 means a table couldn't be read.
        Dictionary<int, int> tcp = new Dictionary<int, int>();
        bool tcpOk = CountTcp(2, tcp, ref r.TcpTotal, ref r.TcpTimeWait);
        tcpOk &= CountTcp(23, tcp, ref r.TcpTotal, ref r.TcpTimeWait);
        if (!tcpOk) { r.TcpTotal = -1; r.TcpTimeWait = -1; }
        Dictionary<int, int> udp = new Dictionary<int, int>();
        bool udpOk = CountUdp(2, udp, ref r.UdpTotal);
        udpOk &= CountUdp(23, udp, ref r.UdpTotal);
        if (!udpOk) r.UdpTotal = -1;
        foreach (ProcInfo p in list) {
            int c;
            if (p.Id == 0) continue;
            if (tcp.TryGetValue(p.Id, out c)) p.TcpConnections = c;
            if (udp.TryGetValue(p.Id, out c)) p.UdpEndpoints = c;
        }
        return r;
    }

    // ---- Security and monitoring agents ----
    // Antivirus, EDR, DLP, web filtering, VPN, monitoring and management agents. A process
    // counts as one if its name matches namePattern, or the publisher (CompanyName) of its
    // program file matches companyPattern and neither its name nor its description matches
    // excludePattern. Services are looked up by the program their service runs, so agents
    // running as SYSTEM are found too. Each process is checked once.
    static Regex agentCompany, agentName, agentExclude;
    static Dictionary<string, string> agentByKey = new Dictionary<string, string>();          // pid|start -> product, "" = not an agent
    static Dictionary<string, string[]> fileInfo = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);   // path -> company, description

    public static void SetAgentRules(string companyPattern, string namePattern, string excludePattern) {
        RegexOptions o = RegexOptions.IgnoreCase | RegexOptions.CultureInvariant;
        agentCompany = string.IsNullOrEmpty(companyPattern) ? null : new Regex(companyPattern, o);
        agentName = string.IsNullOrEmpty(namePattern) ? null : new Regex(namePattern, o);
        agentExclude = string.IsNullOrEmpty(excludePattern) ? null : new Regex(excludePattern, o);
        agentByKey = new Dictionary<string, string>();
    }

    // The product name, or null. known = false when its program file couldn't be found.
    static string ClassifyAgent(ProcInfo p, out bool known) {
        known = true;
        if (p.Id <= 4) return null;
        string path = null;
        string[] svc;
        if (services.TryGetValue(p.Id, out svc)) path = ServiceProgram(svc[0].Split(',')[0].Trim());
        if (path == null) {
            IntPtr h = OpenProcess(0x1000 /* PROCESS_QUERY_LIMITED_INFORMATION */, false, p.Id);
            if (h != IntPtr.Zero) {
                try { StringBuilder sb = new StringBuilder(1024); int n = sb.Capacity; if (QueryFullProcessImageName(h, 0, sb, ref n)) path = sb.ToString(); }
                finally { CloseHandle(h); }
            }
        }
        string[] info = path == null ? null : FileInfoOf(path);
        string company = info == null ? "" : info[0], description = info == null ? "" : info[1];
        bool byName = agentName != null && agentName.IsMatch(p.Name);
        known = path != null || byName;
        bool byCompany = agentCompany != null && company.Length > 0 && agentCompany.IsMatch(company);
        if (!byName && !byCompany) return null;
        if (!byName && agentExclude != null && (agentExclude.IsMatch(p.Name) || agentExclude.IsMatch(description))) return null;
        // "Microsoft Corporation" says nothing; its description (e.g. "Antimalware Service Executable") does.
        if (company.Length == 0 || company.StartsWith("Microsoft", StringComparison.OrdinalIgnoreCase))
            return description.Length > 0 ? description : p.Name;
        return company;
    }

    // The program a service runs, from its ImagePath in the registry.
    static string ServiceProgram(string service) {
        try {
            using (RegistryKey k = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Services\" + service)) {
                string cmd = k == null ? null : k.GetValue("ImagePath") as string;
                if (string.IsNullOrEmpty(cmd)) return null;
                cmd = Environment.ExpandEnvironmentVariables(cmd.Trim());
                if (cmd.StartsWith("\"")) { int end = cmd.IndexOf('"', 1); return end > 1 ? cmd.Substring(1, end - 1) : null; }
                int exe = cmd.IndexOf(".exe", StringComparison.OrdinalIgnoreCase);
                return exe > 0 ? cmd.Substring(0, exe + 4) : cmd;
            }
        } catch { return null; }
    }

    static string[] FileInfoOf(string path) {
        string[] r;
        if (fileInfo.TryGetValue(path, out r)) return r;
        try {
            FileVersionInfo v = FileVersionInfo.GetVersionInfo(path);
            r = new string[] { (v.CompanyName ?? "").Trim(), (v.FileDescription ?? "").Trim() };
        } catch { r = new string[] { "", "" }; }
        fileInfo[path] = r;
        return r;
    }

    // The agent processes, busiest first, and their totals: { CPU %, I/O KB/s, private MB }.
    public static List<ProcInfo> Agents(List<ProcInfo> all, out double[] totals) {
        List<ProcInfo> l = new List<ProcInfo>();
        totals = new double[3];
        foreach (ProcInfo p in all) {
            if (p.Agent == null) continue;
            l.Add(p); totals[0] += p.CpuPct; totals[1] += p.IoKBps; totals[2] += p.PrivateMB;
        }
        l.Sort(delegate(ProcInfo a, ProcInfo b) { return b.CpuPct.CompareTo(a.CpuPct); });
        return l;
    }

    public static string NameOf(int pid) {
        string n;
        if (names.TryGetValue(pid, out n)) return n;
        try { return Process.GetProcessById(pid).ProcessName + ".exe"; } catch { return "?"; }
    }

    public static double Measure(ProcInfo p, string by) {
        switch (by) {
            case "Cpu": return p.CpuPct;
            case "Io": return p.IoKBps;
            case "HardFaults": return p.HardFaultsPerSec;
            case "Handles": return p.Handles;
            case "Gdi": return p.GdiObjects;
            case "User": return p.UserObjects;
            case "Tcp": return p.TcpConnections;
            case "Udp": return p.UdpEndpoints;
            default: return p.PrivateMB;   // "Memory"
        }
    }

    // The top n processes (excluding Idle) by Cpu, Io, HardFaults, Memory, Handles, Gdi,
    // User, Tcp or Udp.
    public static List<ProcInfo> Top(List<ProcInfo> all, string by, int n) {
        List<ProcInfo> l = new List<ProcInfo>();
        foreach (ProcInfo p in all) if (p.Id != 0) l.Add(p);
        l.Sort(delegate(ProcInfo a, ProcInfo b) { return Measure(b, by).CompareTo(Measure(a, by)); });
        if (l.Count > n) l.RemoveRange(n, l.Count - n);
        return l;
    }

    // ---- Address space of 32-bit processes ----
    // A 32-bit program can address 2 GB (4 GB on 64-bit Windows if it is marked
    // large-address-aware), however much RAM the PC has. Over time that space fragments;
    // once no free block is big enough, its allocations fail ("out of memory", windows that
    // don't draw, a crash) while the PC still has plenty of memory. 64-bit programs have
    // 128 TB and never run short, so they are skipped.

    static readonly bool os64 = Environment.Is64BitOperatingSystem;
    static Dictionary<string, ulong> addressTop = new Dictionary<string, ulong>();   // pid|start -> top of usable space; 0 = skip
    static Dictionary<string, bool> largeAware = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

    // Measures the 32-bit processes in this session (the ones it can open: not those run as
    // administrator) and returns them, least room first. Each process's bitness is checked
    // once; the walk costs about a microsecond per memory region.
    public static List<ProcInfo> MeasureAddressSpace(List<ProcInfo> all) {
        List<ProcInfo> measured = new List<ProcInfo>();
        Dictionary<string, ulong> tops = new Dictionary<string, ulong>();
        foreach (ProcInfo p in all) {
            if (p.SessionId != ownSession || p.Id == 0) continue;
            string key = p.Id + "|" + p.CreateTime;
            ulong top;
            bool known = addressTop.TryGetValue(key, out top);
            if (known && top == 0) { tops[key] = 0; continue; }
            IntPtr h = OpenProcess(0x0400 /* PROCESS_QUERY_INFORMATION */, false, p.Id);
            if (h == IntPtr.Zero) { tops[key] = 0; continue; }
            try {
                if (!known) top = UsableTop(h);
                tops[key] = top;
                if (top == 0) continue;
                ulong largest, free;
                FreeAddressSpace(h, top, out largest, out free);
                p.LargestFreeBlockMB = largest / 1048576.0;
                p.FreeAddressSpaceMB = free / 1048576.0;
                measured.Add(p);
            } finally { CloseHandle(h); }
        }
        addressTop = tops;
        measured.Sort(delegate(ProcInfo a, ProcInfo b) { return a.LargestFreeBlockMB.CompareTo(b.LargestFreeBlockMB); });
        return measured;
    }

    // The end of the address space a 32-bit process can use, or 0 for a 64-bit process.
    static ulong UsableTop(IntPtr h) {
        bool wow64;
        if (os64 && !(IsWow64Process(h, out wow64) && wow64)) return 0;
        bool aware = false;
        StringBuilder path = new StringBuilder(1024);
        int size = path.Capacity;
        if (QueryFullProcessImageName(h, 0, path, ref size)) {
            string exe = path.ToString();
            if (!largeAware.TryGetValue(exe, out aware)) { aware = IsLargeAddressAware(exe); largeAware[exe] = aware; }
        }
        if (!aware) return 0x80000000UL;               // 2 GB
        return os64 ? 0x100000000UL : 0xC0000000UL;    // 4 GB; up to 3 GB on 32-bit Windows (the walk stops at the real end)
    }

    // The large-address-aware flag in the program file's PE header.
    static bool IsLargeAddressAware(string exe) {
        try {
            byte[] b = new byte[4096];
            int n;
            using (FileStream fs = new FileStream(exe, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) n = fs.Read(b, 0, b.Length);
            if (n < 64 || b[0] != 'M' || b[1] != 'Z') return false;
            int pe = BitConverter.ToInt32(b, 0x3C);
            if (pe < 0 || pe + 24 > n || b[pe] != 'P' || b[pe + 1] != 'E' || b[pe + 2] != 0 || b[pe + 3] != 0) return false;
            return (BitConverter.ToUInt16(b, pe + 4 + 18) & 0x0020 /* IMAGE_FILE_LARGE_ADDRESS_AWARE */) != 0;
        } catch { return false; }
    }

    // The largest free block and the total free space below top.
    static void FreeAddressSpace(IntPtr h, ulong top, out ulong largest, out ulong total) {
        largest = 0; total = 0;
        MEMORY_BASIC_INFORMATION m;
        UIntPtr len = new UIntPtr((uint)Marshal.SizeOf(typeof(MEMORY_BASIC_INFORMATION)));
        ulong a = 0x10000;   // the first 64 KB can never be allocated
        for (int i = 0; i < 1000000 && a < top; i++) {
            if (VirtualQueryEx(h, ToPtr(a), out m, len) == UIntPtr.Zero) break;   // past the end of the usable space
            ulong start = FromPtr(m.BaseAddress), end = start + m.RegionSize.ToUInt64();
            if (end <= a) break;
            if (m.State == 0x10000 /* MEM_FREE */) {
                ulong free = Math.Min(end, top) - Math.Max(start, a);
                total += free;
                if (free > largest) largest = free;
            }
            a = end;
        }
    }

    // Addresses up to 4 GB, also when this PowerShell is itself 32-bit.
    static IntPtr ToPtr(ulong a) { return IntPtr.Size == 8 ? new IntPtr((long)a) : new IntPtr(unchecked((int)(uint)a)); }
    static ulong FromPtr(IntPtr p) { return IntPtr.Size == 8 ? (ulong)p.ToInt64() : (uint)p.ToInt32(); }
}

// Background threads that catch freezes shorter than the sampling interval.
//  * The stall watch wakes every TickMs at the highest priority and measures how late it
//    woke. A busy CPU doesn't delay it, so being late means nothing could run at all for
//    that long: a driver, firmware or hardware problem, or the monitor's process being
//    suspended (which Windows does to desktop programs in Modern Standby).
//  * The probes send a no-op message (WM_NULL) every ProbeEveryMs to the foreground window
//    and to the taskbar, and time the reply: how long that app, or Explorer, would have
//    ignored a click or a key. Each target has its own thread, so one that hangs doesn't
//    hold up the other.
public static class Responsiveness {
    [DllImport("kernel32.dll")] static extern bool QueryUnbiasedInterruptTime(out ulong time);
    [DllImport("kernel32.dll")] static extern IntPtr GetCurrentThread();
    [DllImport("kernel32.dll")] static extern bool SetThreadPriority(IntPtr thread, int priority);
    [DllImport("user32.dll", SetLastError = true)] static extern IntPtr SendMessageTimeout(IntPtr h, uint msg, UIntPtr w, IntPtr l, uint flags, uint timeoutMs, out UIntPtr result);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr FindWindow(string cls, string title);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern bool IsHungAppWindow(IntPtr h);

    public const int TickMs = 100, StallMs = 250, ProbeEveryMs = 1000, ProbeTimeoutMs = 5000;
    const long Tick = TickMs * 10000L, Stall = StallMs * 10000L;   // in 100 ns units

    static readonly object gate = new object();   // the probes' values (never taken by the stall watch)
    static Thread stallThread;
    // The stall watch shares these with Read() through Interlocked only. With a lock, a busy
    // CPU could pause the low-priority main thread while it held the lock, and the stall
    // watch would wait for it and record a stall that never happened.
    static long lastTick;                         // when the stall watch last woke (100 ns units)
    static long maxLate; static int stalls;
    static double appMs = -1, shellMs = -1; static int appPid;
    static long appSince, shellSince; static int appSincePid;   // a probe still waiting (Stopwatch ticks); 0 = none

    // Time since boot in 100 ns units, not counting sleep or hibernation (so those aren't
    // stalls). It moves in clock ticks of ~16 ms, plenty for stalls of 250 ms or more.
    static long Now() { ulong t; QueryUnbiasedInterruptTime(out t); return (long)t; }

    public static void Start() {
        if (stallThread != null) return;
        Interlocked.Exchange(ref lastTick, Now());
        stallThread = new Thread(StallLoop, 262144);
        stallThread.IsBackground = true;
        stallThread.Start();
        foreach (bool app in new bool[] { true, false }) {
            Thread t = new Thread(ProbeLoop, 262144);
            t.IsBackground = true;
            t.Priority = ThreadPriority.Highest;   // so a busy CPU doesn't delay noticing the reply
            t.Start(app);
        }
    }

    // Total user-mode CPU time of all processors (100 ns units), or -1.
    [DllImport("ntdll.dll")] static extern int NtQuerySystemInformation(int cls, IntPtr buf, int len, out int retLen);
    static readonly IntPtr cpuBuf = Marshal.AllocHGlobal(64 * 48);
    static long UserTime() {
        int ret;
        if (NtQuerySystemInformation(8 /* SystemProcessorPerformanceInformation */, cpuBuf, 64 * 48, out ret) != 0) return -1;
        long sum = 0;
        for (int i = 0; i < ret / 48; i++) sum += Marshal.ReadInt64(cpuBuf, i * 48 + 16);
        return sum;
    }

    static void StallLoop() {
        SetThreadPriority(GetCurrentThread(), 15 /* THREAD_PRIORITY_TIME_CRITICAL */);
        int cores = Environment.ProcessorCount;
        long user = UserTime();
        while (true) {
            Thread.Sleep(TickMs);
            long now = Now();
            long late = now - Interlocked.Exchange(ref lastTick, now) - Tick;
            long userNow = UserTime();
            // Woke late, but other programs got plenty of CPU meanwhile: the PC wasn't stalled,
            // only this process was paused (.NET pauses all its threads for garbage collection,
            // which takes longer when the CPU is busy). Not counted.
            bool othersRan = late >= Stall && user >= 0 && userNow >= 0 && (userNow - user) * 4 > (late + Tick) * cores;
            user = userNow;
            if (othersRan) continue;
            long seen;
            while (late > (seen = Interlocked.Read(ref maxLate)) && Interlocked.CompareExchange(ref maxLate, late, seen) != seen) { }
            if (late >= Stall) Interlocked.Increment(ref stalls);
        }
    }

    static void ProbeLoop(object target) {
        bool app = (bool)target;
        while (true) {
            try {
                IntPtr h = app ? GetForegroundWindow() : FindWindow("Shell_TrayWnd", null);
                if (h != IntPtr.Zero) {
                    uint pid = 0;
                    if (app) GetWindowThreadProcessId(h, out pid);
                    long t0 = Stopwatch.GetTimestamp();
                    lock (gate) { if (app) { appSince = t0; appSincePid = (int)pid; } else shellSince = t0; }
                    double ms = Ask(h);
                    lock (gate) {
                        if (app) {
                            appSince = 0;
                            if (ms > appMs) { appMs = ms; appPid = (int)pid; }
                        } else {
                            shellSince = 0;
                            if (ms > shellMs) shellMs = ms;
                        }
                    }
                }
            } catch { }
            Thread.Sleep(ProbeEveryMs);
        }
    }

    // Milliseconds the window took to handle a WM_NULL message. ProbeTimeoutMs if it didn't
    // within that time, or Windows already shows it as Not Responding. -1 if it can't be
    // asked: it closed, or it runs as administrator and Windows blocks the message.
    public static double Ask(IntPtr h) {
        long t0 = Stopwatch.GetTimestamp();
        UIntPtr result;
        IntPtr ok = SendMessageTimeout(h, 0 /* WM_NULL */, UIntPtr.Zero, IntPtr.Zero, 2 /* SMTO_ABORTIFHUNG */, ProbeTimeoutMs, out result);
        int error = Marshal.GetLastWin32Error();
        double ms = (Stopwatch.GetTimestamp() - t0) * 1000.0 / Stopwatch.Frequency;
        if (ok != IntPtr.Zero) return ms;
        if (error == 1460 /* ERROR_TIMEOUT */ || IsHungAppWindow(h)) return Math.Max(ms, ProbeTimeoutMs);
        return -1;
    }

    // Since the previous call: { longest stall ms, stalls of StallMs or more, slowest
    // foreground reply ms (-1 = none), that window's process id, slowest taskbar reply ms
    // (-1 = none) }. A stall or probe still under way counts with its time so far.
    public static double[] Read() {
        long now = Now();
        long ticks = Stopwatch.GetTimestamp();
        double late = Math.Max(0, Math.Max(Interlocked.Exchange(ref maxLate, 0), now - Interlocked.Read(ref lastTick) - Tick)) / 10000.0;
        int stallCount = Interlocked.Exchange(ref stalls, 0);
        lock (gate) {
            double app = appMs, shell = shellMs;
            int pid = appPid;
            if (appSince != 0) {
                double ms = (ticks - appSince) * 1000.0 / Stopwatch.Frequency;
                if (ms > app) { app = ms; pid = appSincePid; }
            }
            if (shellSince != 0) shell = Math.Max(shell, (ticks - shellSince) * 1000.0 / Stopwatch.Frequency);
            double[] r = new double[] { late, stallCount, app, pid, shell };
            appMs = -1; shellMs = -1; appPid = 0;
            return r;
        }
    }
}

// One check, repeated on its own background thread, so a check that hangs (a DNS server
// that doesn't answer, an unreachable server) holds up neither the samples nor the other
// checks. The check returns null when it worked, or why it failed.
public class Probe {
    readonly object gate = new object();
    readonly Func<string> check;
    readonly int everyMs;
    double worst = -1; int failures, runs; long since; string why;

    public Probe(Func<string> check, int everyMs) {
        this.check = check; this.everyMs = everyMs;
        Thread t = new Thread(Loop, 262144);
        t.IsBackground = true;
        t.Priority = ThreadPriority.Highest;   // so a busy CPU doesn't add to the times measured
        t.Start();
    }

    void Loop() {
        while (true) {
            long t0 = Stopwatch.GetTimestamp();
            lock (gate) since = t0;
            string fail;
            try { fail = check(); } catch (Exception e) { fail = e.Message; }
            double ms = (Stopwatch.GetTimestamp() - t0) * 1000.0 / Stopwatch.Frequency;
            lock (gate) { since = 0; runs++; if (ms > worst) worst = ms; if (fail != null) { failures++; why = fail; } }
            Thread.Sleep(Math.Max(200, everyMs - (int)ms));
        }
    }

    // Since the previous call: { slowest run in ms (a run still going counts with its time so
    // far; -1 = none), runs that failed, why the last failure happened (or null), runs
    // finished }.
    public object[] Read() {
        long now = Stopwatch.GetTimestamp();
        lock (gate) {
            double w = worst;
            if (since != 0) w = Math.Max(w, (now - since) * 1000.0 / Stopwatch.Frequency);
            object[] r = new object[] { w, failures, why, runs };
            worst = -1; failures = 0; why = null; runs = 0;
            return r;
        }
    }
}

// Network and environment checks a desktop depends on: DNS, the domain controller, chosen
// servers, the web through the user's proxy, mapped drives, and how long a simple file
// write and read takes (antivirus and DLP filters sit in that path).
public static class Connectivity {
    [DllImport("dnsapi.dll", CharSet = CharSet.Unicode)] static extern int DnsQuery_W(string name, short type, int options, IntPtr extra, out IntPtr results, IntPtr reserved);
    [DllImport("dnsapi.dll")] static extern void DnsRecordListFree(IntPtr list, int freeType);
    [DllImport("netapi32.dll", CharSet = CharSet.Unicode)] static extern int DsGetDcName(string computer, string domain, IntPtr guid, string site, int flags, out IntPtr info);
    [DllImport("netapi32.dll", CharSet = CharSet.Unicode)] static extern int NetGetJoinInformation(string server, out IntPtr name, out int status);
    [DllImport("netapi32.dll")] static extern int NetApiBufferFree(IntPtr b);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct DOMAIN_CONTROLLER_INFO {
        public string DomainControllerName, DomainControllerAddress;
        public uint DomainControllerAddressType;
        public Guid DomainGuid;
        public string DomainName, DnsForestName;
        public uint Flags;
        public string DcSiteName, ClientSiteName;
    }

    public static Probe Dns, Dc, Http, Drives, File;
    public static Probe[] Servers = new Probe[0];
    public static volatile string DcText;   // e.g. "dc01.corp.local (site London)"
    static volatile string dcHost;
    static long dcFoundAt;
    static int changes;

    // The Active Directory domain this PC is joined to (NetBIOS name), or null.
    public static string JoinedDomain() {
        IntPtr name; int status;
        if (NetGetJoinInformation(null, out name, out status) != 0) return null;
        try { return status == 3 /* NetSetupDomainName */ ? Marshal.PtrToStringUni(name) : null; }
        finally { NetApiBufferFree(name); }
    }

    public static void Start(string dnsName, bool domain, string[] servers, string httpUrl, string httpExpect, int everyMs) {
        if (Dns != null) return;
        NetworkChange.NetworkAddressChanged += delegate { Interlocked.Increment(ref changes); };
        NetworkChange.NetworkAvailabilityChanged += delegate { Interlocked.Increment(ref changes); };
        if (!string.IsNullOrEmpty(dnsName)) Dns = new Probe(delegate { return CheckDns(dnsName); }, everyMs);
        if (domain) Dc = new Probe(CheckDc, everyMs);
        List<Probe> list = new List<Probe>();
        foreach (string s in servers) {
            string target = s;
            int colon = target.LastIndexOf(':');
            int port;
            if (colon < 1 || !int.TryParse(target.Substring(colon + 1), out port)) continue;
            string host = target.Substring(0, colon);
            list.Add(new Probe(delegate { return Connect(host, port, 3000); }, everyMs));
        }
        Servers = list.ToArray();
        if (!string.IsNullOrEmpty(httpUrl)) Http = new Probe(delegate { return CheckHttp(httpUrl, httpExpect); }, 60000);
        Drives = new Probe(CheckDrives, 60000);
    }

    public static void StartFileCheck(int everyMs) { if (File == null) File = new Probe(CheckFile, everyMs); }

    // Network address or availability changes (an adapter connecting or dropping, a new IP
    // address, a VPN connecting) since the previous call.
    public static int TakeChanges() { return Interlocked.Exchange(ref changes, 0); }

    // A fresh lookup from the DNS servers, not the PC's cache. A name starting with _ is
    // looked up as an SRV record (the records domain controllers are found by).
    static string CheckDns(string name) {
        IntPtr res;
        short type = name.StartsWith("_") ? (short)33 : (short)1;
        int rc = DnsQuery_W(name, type, 0x8 /* BYPASS_CACHE */ | 0x40 /* NO_HOSTS_FILE */, IntPtr.Zero, out res, IntPtr.Zero);
        if (res != IntPtr.Zero) DnsRecordListFree(res, 1);
        if (rc == 0) return null;
        return name + ": " + new System.ComponentModel.Win32Exception(rc).Message;
    }

    // Finds the domain controller this PC uses (refreshed every 5 minutes, or after a
    // failure), then times a connection to its LDAP port.
    static string CheckDc() {
        long now = Stopwatch.GetTimestamp();
        if (dcHost == null || now - dcFoundAt > 300L * Stopwatch.Frequency) {
            IntPtr p;
            int rc = DsGetDcName(null, null, IntPtr.Zero, null, 0x40000000 /* DS_RETURN_DNS_NAME */, out p);
            if (rc != 0) { dcHost = null; DcText = null; return "no domain controller found: " + new System.ComponentModel.Win32Exception(rc).Message; }
            try {
                DOMAIN_CONTROLLER_INFO i = (DOMAIN_CONTROLLER_INFO)Marshal.PtrToStructure(p, typeof(DOMAIN_CONTROLLER_INFO));
                dcHost = (i.DomainControllerName ?? "").TrimStart('\\');
                string text = dcHost;
                if (!string.IsNullOrEmpty(i.DcSiteName)) {
                    text += " (site " + i.DcSiteName;
                    if (!string.IsNullOrEmpty(i.ClientSiteName) && !string.Equals(i.ClientSiteName, i.DcSiteName, StringComparison.OrdinalIgnoreCase))
                        text += ", but this PC is in site " + i.ClientSiteName;
                    else if (string.IsNullOrEmpty(i.ClientSiteName)) text += ", this PC's site unknown";
                    text += ")";
                }
                DcText = text;
                dcFoundAt = now;
            } finally { NetApiBufferFree(p); }
        }
        string fail = Connect(dcHost, 389, 3000);
        if (fail != null) dcHost = null;   // find a domain controller again next time
        return fail;
    }

    // Times a TCP connection (the name lookup included).
    static string Connect(string host, int port, int timeoutMs) {
        TcpClient c = new TcpClient();
        try {
            IAsyncResult ar = c.BeginConnect(host, port, null, null);
            if (!ar.AsyncWaitHandle.WaitOne(timeoutMs)) return host + ":" + port + " didn't answer within " + (timeoutMs / 1000) + " s";
            c.EndConnect(ar);
            return null;
        } catch (Exception e) {
            return host + ":" + port + ": " + (e.InnerException ?? e).Message;
        } finally { c.Close(); }
    }

    // A web request through the proxy settings the user's programs use (including a PAC
    // script, WPAD and proxy sign-in), so a slow or broken proxy shows up here.
    static IWebProxy proxy; static long proxyAt;
    static string CheckHttp(string url, string expect) {
        long now = Stopwatch.GetTimestamp();
        if (proxy == null || now - proxyAt > 3600L * Stopwatch.Frequency) {
            proxy = WebRequest.GetSystemWebProxy();
            proxy.Credentials = CredentialCache.DefaultCredentials;
            proxyAt = now;
        }
        HttpWebRequest r = (HttpWebRequest)WebRequest.Create(url);
        r.Proxy = proxy;
        r.UseDefaultCredentials = true;
        r.KeepAlive = false;                    // a new connection each time, like a first visit
        r.AllowAutoRedirect = false;            // a redirect usually means a proxy or sign-in page
        r.Timeout = 15000; r.ReadWriteTimeout = 15000;
        r.UserAgent = "DesktopMonitor";
        try {
            using (HttpWebResponse resp = (HttpWebResponse)r.GetResponse())
            using (StreamReader sr = new StreamReader(resp.GetResponseStream())) {
                char[] buf = new char[4096];
                int n = sr.ReadBlock(buf, 0, buf.Length);
                int code = (int)resp.StatusCode;
                if (code >= 300 && code < 400) return "redirected to " + resp.Headers["Location"] + " (a proxy or sign-in page?)";
                if (code != 200) return "HTTP " + code + " " + resp.StatusDescription;
                if (!string.IsNullOrEmpty(expect) && new string(buf, 0, n).IndexOf(expect, StringComparison.Ordinal) < 0)
                    return "unexpected page (a proxy block or sign-in page?)";
                return null;
            }
        } catch (WebException e) {
            HttpWebResponse resp = e.Response as HttpWebResponse;
            if (resp != null) { using (resp) return "HTTP " + (int)resp.StatusCode + " " + resp.StatusDescription; }
            return e.Status + ": " + e.Message;
        }
    }

    // Mapped drives (persistent ones, in HKCU\Network) whose server doesn't answer on the
    // file sharing port. File Explorer and Office can hang for 30 s or more on such a drive.
    static string CheckDrives() {
        List<string> down = new List<string>();
        using (RegistryKey net = Registry.CurrentUser.OpenSubKey("Network")) {
            if (net == null) return null;
            foreach (string letter in net.GetSubKeyNames()) {
                string path;
                using (RegistryKey k = net.OpenSubKey(letter)) path = k == null ? null : k.GetValue("RemotePath") as string;
                if (path == null || !path.StartsWith(@"\\") || path.Contains("@")) continue;   // skip WebDAV
                string server = path.Substring(2).Split('\\')[0];
                if (Connect(server, 445, 3000) != null) down.Add(letter.ToUpperInvariant() + ": " + path);
            }
        }
        return down.Count == 0 ? null : string.Join("; ", down.ToArray());
    }

    // Writes, reads back and deletes a 64 KB file in the user's temp folder. Antivirus, EDR
    // and DLP products check files as they are written and opened; when one of them is
    // struggling, every program that touches files slows down, and this shows it.
    static readonly byte[] fileData = Encoding.ASCII.GetBytes(new string('x', 63) + "\n");
    static string CheckFile() {
        string path = Path.Combine(Path.GetTempPath(), "DesktopMonitor-check-" + Process.GetCurrentProcess().Id + ".txt");
        byte[] buf = new byte[65536];
        using (FileStream f = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
            for (int i = 0; i < 1024; i++) f.Write(fileData, 0, fileData.Length);
        using (FileStream f = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            while (f.Read(buf, 0, buf.Length) > 0) { }
        System.IO.File.Delete(path);
        return null;
    }
}

// Zips each finished day into DesktopMonitor_<PC>_<yyyyMMdd>.zip, deletes days past the
// retention period, and deletes the oldest days while this PC's logs take more than the
// size limit. It runs on a low-priority background thread so the samples aren't held up,
// and touches only this PC's files (several PCs can share a log folder).
//
// A file is renamed to *.zipping before it is zipped. That fails if another program has
// it open without allowing deletion (Excel, say) or if it belongs to another user, and the
// file is then left for the next pass. If the monitor stops part-way, the *.zipping files
// are zipped on the next pass; a zip still being written (*.zip.tmp) is discarded first.
public static class Housekeeping {
    static Thread worker;
    static readonly List<string> messages = new List<string>();
    static readonly HashSet<string> warned = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    static readonly CultureInfo inv = CultureInfo.InvariantCulture;

    class LogFile { public string Path, Name, Day; public long Size; public bool IsZip, IsTemp; }

    public static bool Busy { get { Thread w = worker; return w != null && w.IsAlive; } }

    public static bool Start(string dir, string pc, bool zip, int retentionDays, long maxBytes) {
        if (Busy) return false;
        worker = new Thread(delegate() {
            try { Run(dir, pc, zip, retentionDays, maxBytes); }
            catch (Exception e) { Say("WARN", "Log housekeeping failed: " + e.Message); }
        }, 1 << 20);
        worker.IsBackground = true;
        worker.Priority = ThreadPriority.Lowest;
        worker.Start();
        return true;
    }

    public static bool Wait(int ms) { Thread w = worker; return w == null || w.Join(ms); }

    // "LEVEL message" lines for the monitor's diagnostic log, since the last call.
    public static string[] TakeMessages() { lock (messages) { string[] r = messages.ToArray(); messages.Clear(); return r; } }
    static void Say(string level, string m) { lock (messages) messages.Add(level + " " + m); }
    static void SayOnce(string key, string m) { lock (warned) { if (!warned.Add(key)) return; } Say("WARN", m); }

    static List<LogFile> List(string dir, string pc) {
        string p = Regex.Escape(pc);
        Regex re = new Regex("^(?:(?:metrics|processes|events)_" + p + @"_(?<day>\d{8})(?:_\d+)?\.(?:csv|tsv)" +
                             "|inventory_" + p + @"_(?<day>\d{8})-\d{6}\.(?:csv|tsv)" +
                             "|DesktopMonitor_" + p + @"_(?<day>\d{8})(?:_\d+)?\.zip(?<tmp>\.tmp)?)(?<zipping>\.zipping)?$", RegexOptions.IgnoreCase);
        List<LogFile> list = new List<LogFile>();
        foreach (FileInfo f in new DirectoryInfo(dir).GetFiles()) {
            Match m = re.Match(f.Name);
            if (!m.Success) continue;
            LogFile l = new LogFile();
            l.Path = f.FullName; l.Size = f.Length; l.Day = m.Groups["day"].Value;
            l.Name = m.Groups["zipping"].Success ? f.Name.Substring(0, f.Name.Length - ".zipping".Length) : f.Name;
            l.IsTemp = m.Groups["tmp"].Success;
            l.IsZip = !l.IsTemp && l.Name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase);
            list.Add(l);
        }
        list.Sort(delegate(LogFile a, LogFile b) { return string.CompareOrdinal(a.Day + a.Name, b.Day + b.Name); });
        return list;
    }

    public static void Run(string dir, string pc, bool zip, int retentionDays, long maxBytes) {
        string today = DateTime.Today.ToString("yyyyMMdd", inv);
        string cutoff = DateTime.Today.AddDays(-retentionDays).ToString("yyyyMMdd", inv);
        List<LogFile> files = List(dir, pc);
        foreach (LogFile f in files) if (f.IsTemp) Delete(f);

        // Days past the retention period first, so they aren't zipped only to be deleted.
        foreach (KeyValuePair<string, List<LogFile>> d in ByDay(files)) {
            if (string.CompareOrdinal(d.Key, cutoff) >= 0) break;
            long freed; int deleted = DeleteDay(d.Value, out freed);
            if (deleted > 0) Say("INFO", string.Format(inv, "Deleted the logs from {0} ({1}), which were more than {2} days old.", d.Key, Files(deleted), retentionDays));
        }

        if (zip) {
            foreach (KeyValuePair<string, List<LogFile>> d in ByDay(files))
                if (string.CompareOrdinal(d.Key, cutoff) >= 0 && string.CompareOrdinal(d.Key, today) < 0) ZipDay(dir, pc, d.Key, d.Value);
        }

        // Then the oldest days until this PC's logs fit in maxBytes. Today's are kept.
        files = List(dir, pc);
        long total = 0;
        foreach (LogFile f in files) if (!f.IsTemp) total += f.Size;
        foreach (KeyValuePair<string, List<LogFile>> d in ByDay(files)) {
            if (total <= maxBytes || string.CompareOrdinal(d.Key, today) >= 0) break;
            long freed; int deleted = DeleteDay(d.Value, out freed);
            total -= freed;
            if (deleted > 0) Say("INFO", string.Format(inv, "Deleted the logs from {0} ({1}, {2}) to keep this PC's logs under {3} MB.", d.Key, Files(deleted), Size(freed), maxBytes / 1048576));
        }
        if (total > maxBytes)
            SayOnce("size|" + today, string.Format(inv, "This PC's logs take {0}, more than the {1} MB limit, but today's files are never deleted.", Size(total), maxBytes / 1048576));
    }

    static int DeleteDay(List<LogFile> dayFiles, out long freed) {
        freed = 0; int deleted = 0;
        foreach (LogFile f in dayFiles) if (!f.IsTemp && Delete(f)) { freed += f.Size; deleted++; }
        return deleted;
    }

    static string Files(int n) { return n == 1 ? "1 file" : n + " files"; }
    static string Size(long bytes) {
        return bytes < 1048576 ? string.Format(inv, "{0:0} KB", bytes / 1024.0) : string.Format(inv, "{0:0.#} MB", bytes / 1048576.0);
    }

    static SortedDictionary<string, List<LogFile>> ByDay(List<LogFile> files) {
        SortedDictionary<string, List<LogFile>> days = new SortedDictionary<string, List<LogFile>>(StringComparer.Ordinal);
        foreach (LogFile f in files) {
            List<LogFile> l;
            if (!days.TryGetValue(f.Day, out l)) { l = new List<LogFile>(); days[f.Day] = l; }
            l.Add(f);
        }
        return days;
    }

    static bool Delete(LogFile f) {
        try { File.Delete(f.Path); return true; }
        catch (Exception e) { SayOnce("delete|" + f.Path, "Could not delete " + f.Path + ": " + e.Message); return false; }
    }

    static void ZipDay(string dir, string pc, string day, List<LogFile> dayFiles) {
        // Names and sizes already in this day's zips: if the monitor stopped after a zip was
        // finished but before its files were deleted, they only need deleting.
        HashSet<string> zipped = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        List<LogFile> ready = new List<LogFile>();
        foreach (LogFile f in dayFiles) {
            if (f.IsTemp) continue;
            if (f.IsZip) {
                try {
                    using (ZipArchive z = new ZipArchive(new FileStream(f.Path, FileMode.Open, FileAccess.Read, FileShare.Read), ZipArchiveMode.Read))
                        foreach (ZipArchiveEntry e in z.Entries) zipped.Add(e.Name + "|" + e.Length);
                } catch { }
                continue;
            }
            if (!f.Path.EndsWith(".zipping", StringComparison.OrdinalIgnoreCase)) {
                try { File.Move(f.Path, f.Path + ".zipping"); f.Path += ".zipping"; }
                catch (Exception e) { SayOnce("zip|" + f.Path, "Could not zip " + f.Name + " yet (" + e.Message + "). It will be tried again every hour."); continue; }
            }
            ready.Add(f);
        }
        List<LogFile> todo = new List<LogFile>();
        foreach (LogFile f in ready) {
            if (zipped.Contains(f.Name + "|" + f.Size)) Delete(f); else todo.Add(f);
        }
        if (todo.Count == 0) return;

        string zipPath = Path.Combine(dir, "DesktopMonitor_" + pc + "_" + day + ".zip");
        for (int n = 2; File.Exists(zipPath); n++) zipPath = Path.Combine(dir, "DesktopMonitor_" + pc + "_" + day + "_" + n + ".zip");
        string tmp = zipPath + ".tmp";
        long before = 0;
        try {
            using (FileStream fs = new FileStream(tmp, FileMode.CreateNew, FileAccess.Write))
            using (ZipArchive z = new ZipArchive(fs, ZipArchiveMode.Create)) {
                foreach (LogFile f in todo) {
                    ZipArchiveEntry e = z.CreateEntry(f.Name, CompressionLevel.Optimal);
                    e.LastWriteTime = File.GetLastWriteTime(f.Path);
                    using (Stream src = new FileStream(f.Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                    using (Stream dst = e.Open()) src.CopyTo(dst);
                    before += f.Size;
                }
            }
            File.Move(tmp, zipPath);
        } catch (Exception e) {
            try { File.Delete(tmp); } catch { }
            foreach (LogFile f in ready) { try { File.Move(f.Path, f.Path.Substring(0, f.Path.Length - ".zipping".Length)); } catch { } }
            Say("WARN", "Could not zip the logs from " + day + ": " + e.Message);
            return;
        }
        foreach (LogFile f in todo) Delete(f);
        Say("INFO", string.Format(inv, "Zipped the logs from {0} ({1}) into {2}: {3} down to {4}.",
            day, Files(todo.Count), Path.GetFileName(zipPath), Size(before), Size(new FileInfo(zipPath).Length)));
    }
}
}
'@

# ---------------------------------------------------------------------------------------
# Performance counters. Each is opened once and re-read every sample, which is far
# cheaper than Get-Counter or WMI. .NET accepts the English counter names on Windows in
# any language.
# ---------------------------------------------------------------------------------------

$script:Counters = @{}

function Open-Counter([string] $Key, [string] $Category, [string] $Counter, [string] $Instance = '') {
    try {
        $pc = New-Object Diagnostics.PerformanceCounter($Category, $Counter, $Instance, $true)
        [void] $pc.NextValue()   # rate counters need a first reading to compare against
        $script:Counters[$Key] = $pc
    } catch {
        Write-MonitorLog "Counter \$Category($Instance)\$Counter is not available, so its column stays blank: $($_.Exception.Message)" 'WARN'
    }
}

function Read-Counter([string] $Key) {
    $pc = $script:Counters[$Key]
    if ($null -eq $pc) { return $null }
    try { return [double] $pc.NextValue() }
    catch { Write-Throttled "counter:$Key" "Reading counter $Key failed: $($_.Exception.Message)"; return $null }
}

$script:SystemCat = $null
$script:CtxPrev = $null

function Initialize-Counters {
    if (Test-Group 'Cpu') {
        Open-Counter CpuPct   'Processor' '% Processor Time' '_Total'
        Open-Counter CpuPriv  'Processor' '% Privileged Time' '_Total'
        Open-Counter CpuIntr  'Processor' '% Interrupt Time' '_Total'
        Open-Counter CpuDpc   'Processor' '% DPC Time' '_Total'
        Open-Counter CpuPerf  'Processor Information' '% Processor Performance' '_Total'
        Open-Counter CpuFreq  'Processor Information' 'Processor Frequency' '_Total'
        try { [void] [DesktopMonitor.Native]::CoreLoad() } catch { }   # first reading to compare against
        # Each read of a "System" counter costs ~10 ms (Windows counts every process and
        # thread to produce it), so read the whole category once per sample instead.
        try {
            $script:SystemCat = New-Object Diagnostics.PerformanceCounterCategory('System')
            $row = @{}; Read-SystemCounters $row
        } catch { $script:SystemCat = $null; Write-MonitorLog "Cannot read the System counters: $($_.Exception.Message)" 'WARN' }
    }
    if (Test-Group 'Memory') {
        Open-Counter PageIn   'Memory' 'Pages Input/sec'
        Open-Counter PageOut  'Memory' 'Pages Output/sec'
        Open-Counter PoolNP   'Memory' 'Pool Nonpaged Bytes'
        Open-Counter PoolP    'Memory' 'Pool Paged Bytes'
        Open-Counter Modified 'Memory' 'Modified Page List Bytes'
        Open-Counter PageFile 'Paging File' '% Usage' '_Total'
    }
    if (Test-Group 'Disk') {
        Open-Counter DiskIdle   'PhysicalDisk' '% Idle Time' '_Total'
        Open-Counter DiskQueue  'PhysicalDisk' 'Avg. Disk Queue Length' '_Total'
        Open-Counter DiskRead   'PhysicalDisk' 'Avg. Disk sec/Read' '_Total'
        Open-Counter DiskWrite  'PhysicalDisk' 'Avg. Disk sec/Write' '_Total'
        Open-Counter DiskReadB  'PhysicalDisk' 'Disk Read Bytes/sec' '_Total'
        Open-Counter DiskWriteB 'PhysicalDisk' 'Disk Write Bytes/sec' '_Total'
        Open-Counter DiskIops   'PhysicalDisk' 'Disk Transfers/sec' '_Total'
    }
    if (Test-Group 'Network') {
        # IPv4 and IPv6 are counted separately; many home and office networks use both.
        Open-Counter TcpConn     'TCPv4' 'Connections Established'
        Open-Counter TcpRetrans  'TCPv4' 'Segments Retransmitted/sec'
        Open-Counter TcpConn6    'TCPv6' 'Connections Established'
        Open-Counter TcpRetrans6 'TCPv6' 'Segments Retransmitted/sec'
    }
}

# The sum of the counters that could be read, or $null if none.
function Read-CounterSum([string[]] $Keys) {
    $sum = $null
    foreach ($k in $Keys) { $v = Read-Counter $k; if ($null -ne $v) { $sum = [double] $sum + $v } }
    return $sum
}

# ---------------------------------------------------------------------------------------
# Collectors. Each fills in its columns of $row; a failure blanks only its own columns.
# ---------------------------------------------------------------------------------------

function Get-CpuMetrics($row) {
    $row.CpuPct           = Read-Counter CpuPct
    $row.CpuPrivilegedPct = Read-Counter CpuPriv
    $row.CpuInterruptPct  = Read-Counter CpuIntr
    $row.CpuDpcPct        = Read-Counter CpuDpc
    $perf = Read-Counter CpuPerf
    $freq = Read-Counter CpuFreq
    $row.CpuPerfPct = $perf
    if ($null -ne $perf -and $null -ne $freq) { $row.CpuMHz = [math]::Round($freq * $perf / 100) }
    # The busiest single core, and the most time one core spent on drivers' interrupts and
    # DPCs: a driver hogging one core stutters the PC while the averages look fine.
    $cores = [DesktopMonitor.Native]::CoreLoad()
    if ($cores) { $row.CpuMaxCorePct = $cores[0]; $row.CpuMaxCoreDpcPct = $cores[1] }
    if ($script:SystemCat) { Read-SystemCounters $row }
}

function Read-SystemCounters($row) {
    $data = $script:SystemCat.ReadCategory()
    foreach ($pair in @(@('Processor Queue Length', 'CpuQueue'), @('Processes', 'Processes'), @('Threads', 'Threads'))) {
        $c = $data[$pair[0]]
        if ($c) { foreach ($inst in $c.Values) { $row[$pair[1]] = $inst.RawValue } }
    }
    $c = $data['Context Switches/sec']
    if ($c) {
        foreach ($inst in $c.Values) {
            if ($script:CtxPrev) { $row.ContextSwitchesPerSec = [Diagnostics.CounterSample]::Calculate($script:CtxPrev, $inst.Sample) }
            $script:CtxPrev = $inst.Sample
        }
    }
}

function Get-MemoryMetrics($row) {
    $m = [DesktopMonitor.Native]::Memory()
    if ($m) {
        $row.MemAvailableMB = $m[1] / 1MB
        if ($m[0] -gt 0) { $row.MemUsedPct = 100.0 * ($m[0] - $m[1]) / $m[0] }
        $row.CommitGB      = ($m[2] - $m[3]) / 1GB
        $row.CommitLimitGB = $m[2] / 1GB
        if ($m[2] -gt 0) { $row.CommitPct = 100.0 * ($m[2] - $m[3]) / $m[2] }
    }
    $row.PageInPerSec    = Read-Counter PageIn
    $row.PageOutPerSec   = Read-Counter PageOut
    $row.PageFileUsedPct = Read-Counter PageFile
    $v = Read-Counter PoolNP; if ($null -ne $v) { $row.PoolNonPagedMB = $v / 1MB }
    $v = Read-Counter PoolP;  if ($null -ne $v) { $row.PoolPagedMB = $v / 1MB }
    # Changed pages waiting to be written to disk. Gigabytes here, with a slow disk, make
    # programs wait for memory.
    $v = Read-Counter Modified; if ($null -ne $v) { $row.MemModifiedMB = $v / 1MB }

    # The biggest kernel pool tags, once a minute and repeated in between.
    if (Test-Every $script:EveryMinute) {
        try { $script:PoolTop = [DesktopMonitor.Native]::TopPoolTags() }
        catch { $script:PoolTop = $null; Write-Throttled 'poolTags' "Reading pool tags failed: $($_.Exception.Message)" }
    }
    if ($script:PoolTop) {
        $row.TopNonPagedPoolTag = $script:PoolTop.NonPagedTag; $row.TopNonPagedPoolMB = $script:PoolTop.NonPagedMB
        $row.TopPagedPoolTag = $script:PoolTop.PagedTag;       $row.TopPagedPoolMB = $script:PoolTop.PagedMB
    }
}
$script:PoolTop = $null

function Get-DiskMetrics($row) {
    $idle = Read-Counter DiskIdle
    if ($null -ne $idle) { $row.DiskBusyPct = [math]::Max(0, [math]::Min(100, 100 - $idle)) }
    $row.DiskQueue = Read-Counter DiskQueue
    $v = Read-Counter DiskRead;   if ($null -ne $v) { $row.DiskReadMs = $v * 1000 }
    $v = Read-Counter DiskWrite;  if ($null -ne $v) { $row.DiskWriteMs = $v * 1000 }
    $v = Read-Counter DiskReadB;  if ($null -ne $v) { $row.DiskReadMBps = $v / 1MB }
    $v = Read-Counter DiskWriteB; if ($null -ne $v) { $row.DiskWriteMBps = $v / 1MB }
    $row.DiskIops = Read-Counter DiskIops

    $others = New-Object 'System.Collections.Generic.List[string]'
    foreach ($d in [IO.DriveInfo]::GetDrives()) {
        if ($d.DriveType -ne 'Fixed' -or -not $d.IsReady -or $d.TotalSize -le 0) { continue }
        $freeGB = $d.AvailableFreeSpace / 1GB
        $pct = 100.0 * $d.AvailableFreeSpace / $d.TotalSize
        $letter = $d.Name.TrimEnd('\')
        if ($letter -ieq $env:SystemDrive) {
            $row.SysDriveFreeGB = $freeGB
            $row.SysDriveFreePct = $pct
        } else {
            $others.Add(('{0}{1}GB({2}%)' -f $letter, $freeGB.ToString('0.#', $script:Inv), $pct.ToString('0', $script:Inv)))
        }
    }
    if ($others.Count) { $row.OtherDrivesFree = [string]::Join(' ', $others) }
}

# Up, non-loopback adapters that have a link speed.
function Get-ActiveAdapters {
    foreach ($ni in [Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
        if ($ni.OperationalStatus -ne 'Up' -or $ni.Speed -le 0) { continue }
        if ($ni.NetworkInterfaceType -eq 'Loopback' -or $ni.NetworkInterfaceType -eq 'Tunnel') { continue }
        $ni
    }
}

# Traffic comes from the "Network Interface" counters (physical adapters only, so VPN and
# virtual adapters don't double-count, and ~1 ms versus ~8 ms for the .NET adapter list).
# The adapter list itself is refreshed once a minute.
$script:NetCat = $null
$script:NetPrev = @{}
$script:AdapterText = $null
function Get-NetworkMetrics($row) {
    if ($script:NetCat) {
        $data = $script:NetCat.ReadCategory()
        $rx = 0.0; $tx = 0.0; $err = 0L; $have = $false
        $current = @{}
        foreach ($pair in @(@('Bytes Received/sec', 'rx'), @('Bytes Sent/sec', 'tx'), @('Packets Received Errors', 'err'), @('Packets Outbound Errors', 'err'))) {
            $c = $data[$pair[0]]
            if (-not $c) { continue }
            foreach ($inst in $c.Values) {
                $key = $pair[0] + '|' + $inst.InstanceName
                $current[$key] = $inst.Sample
                $old = $script:NetPrev[$key]
                if ($null -eq $old) { continue }   # first reading, or a new adapter
                $have = $true
                if ($pair[1] -eq 'rx') { $rx += [Diagnostics.CounterSample]::Calculate($old, $inst.Sample) }
                elseif ($pair[1] -eq 'tx') { $tx += [Diagnostics.CounterSample]::Calculate($old, $inst.Sample) }
                elseif ($inst.RawValue -gt $old.RawValue) { $err += $inst.RawValue - $old.RawValue }
            }
        }
        $script:NetPrev = $current
        if ($have) { $row.NetRecvKBps = $rx / 1KB; $row.NetSentKBps = $tx / 1KB; $row.NetErrors = $err }
    }
    if ($null -eq $script:AdapterText -or (Test-Every $script:EveryMinute)) {
        $script:AdapterText = [string]::Join('; ', [string[]] @(Get-ActiveAdapters | ForEach-Object { '{0} {1}Mbps' -f $_.Name, [math]::Round($_.Speed / 1e6) }))
    }
    $row.NetAdapters = $script:AdapterText
    $row.TcpConnections   = Read-CounterSum TcpConn, TcpConn6
    $row.TcpRetransPerSec = Read-CounterSum TcpRetrans, TcpRetrans6
}

# netsh is a separate process, so Wi-Fi signal is read once a minute and repeated between.
$script:WifiSignal = $null
function Get-WifiMetrics($row) {
    if (Test-Every $script:EveryMinute) {
        $script:WifiSignal = $null
        if (@(Get-ActiveAdapters | Where-Object { $_.NetworkInterfaceType -eq 'Wireless80211' }).Count) {
            $out = & "$env:SystemRoot\System32\netsh.exe" wlan show interfaces 2>$null
            # The labels are translated ("Signal", "Señal", "Segnale", ...), but the signal is
            # the only value shown as a percentage, so match "<label> : <n>%" in any language.
            $m = $out | Select-String -Pattern '^\s*[^:]+:\s*(\d+)\s*%\s*$' | Select-Object -First 1
            if ($m) { $script:WifiSignal = [int] $m.Matches[0].Groups[1].Value }
        }
    }
    $row.WifiSignalPct = $script:WifiSignal
}

function Get-DefaultGateway {
    foreach ($ni in @(Get-ActiveAdapters)) {
        foreach ($g in $ni.GetIPProperties().GatewayAddresses) {
            if ($g.Address.AddressFamily -eq 'InterNetwork' -and $g.Address.ToString() -ne '0.0.0.0') {
                return $g.Address.ToString()
            }
        }
    }
    return $null
}

# Both targets are pinged at once, and the wait is capped, so a dead network or a DNS
# lookup that hangs can't hold up the sample. The wait is not counted in CollectMs, which
# is meant to show the monitor itself being held up (a freeze), not the network being down.
$script:Pingers = @((New-Object Net.NetworkInformation.Ping), (New-Object Net.NetworkInformation.Ping))
$script:PingTasks = @($null, $null)
$script:Gateway = $null
function Get-PingMetrics($row) {
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $failed = 0
    $sent = @{}
    for ($i = 0; $i -lt $PingTarget.Count; $i++) {
        $target = $PingTarget[$i]
        if ($target -eq 'gateway') {
            if (-not $script:Gateway -or (Test-Every $script:EveryMinute)) { $script:Gateway = Get-DefaultGateway }
            $target = $script:Gateway
        }
        if (-not $target) { $failed++; continue }   # no network connection at all
        # A ping still running from an earlier sample (a DNS lookup that hangs) counts as failed.
        $old = $script:PingTasks[$i]
        if ($old -and -not $old.IsCompleted) { $failed++; continue }
        try { $sent[$i] = $script:Pingers[$i].SendPingAsync($target, $PingTimeoutMs); $script:PingTasks[$i] = $sent[$i] }
        catch { $failed++ }
    }
    if ($sent.Count) {
        try { [void] [Threading.Tasks.Task]::WaitAll([Threading.Tasks.Task[]] @($sent.Values), $PingTimeoutMs + 500) } catch { }   # a failed ping throws here
    }
    foreach ($i in $sent.Keys) {
        $t = $sent[$i]
        if ($t.Status -eq 'RanToCompletion' -and $t.Result.Status -eq 'Success') { $row["Ping$($i + 1)Ms"] = $t.Result.RoundtripTime; continue }
        $failed++
        if ($PingTarget[$i] -eq 'gateway') { $script:Gateway = $null }   # look it up again next time
    }
    $row.PingFailed = $failed
    $script:WaitMs += $clock.Elapsed.TotalMilliseconds
}

# GPU use is per process and per engine. Like Task Manager, add up the processes on each
# engine and report the busiest engine; a process's own use is its busiest engine.
# ReadCategory gets every instance in one call.
$script:GpuEngineCat = $null
$script:GpuMemoryCat = $null
$script:GpuPrev = @{}
function Get-GpuMetrics($row) {
    if ($script:GpuEngineCat) {
        $util = $script:GpuEngineCat.ReadCategory()['Utilization Percentage']
        $current = @{}
        $engines = @{}
        $byPid = @{}
        if ($util) {
            foreach ($inst in $util.Values) {
                $current[$inst.InstanceName] = $inst.Sample
                $old = $script:GpuPrev[$inst.InstanceName]
                if ($null -eq $old) { continue }
                $v = [Diagnostics.CounterSample]::Calculate($old, $inst.Sample)
                # pid_123_luid_0x0_0x1_phys_0_eng_3_engtype_3D -> the process, and one key per engine
                if ($inst.InstanceName -notmatch '^pid_(\d+)_(.*)$') { continue }
                $engine = $Matches[2]
                $engines[$engine] = $engines[$engine] + $v
                $procId = [int] $Matches[1]
                if ($v -gt $byPid[$procId]) { $byPid[$procId] = $v }
            }
        }
        $script:GpuPrev = $current
        if ($engines.Count) {
            $max = 0.0; $max3d = 0.0
            foreach ($k in $engines.Keys) {
                $v = [math]::Min(100.0, $engines[$k])
                if ($v -gt $max) { $max = $v }
                if ($k -like '*engtype_3D' -and $v -gt $max3d) { $max3d = $v }
            }
            $row.GpuPct = $max
            $row.Gpu3DPct = $max3d
        }
        $topPid = 0; $topPct = 1.0   # ignore processes under 1%
        foreach ($k in $byPid.Keys) { if ($k -gt 0 -and $byPid[$k] -ge $topPct) { $topPid = $k; $topPct = $byPid[$k] } }
        if ($topPid) {
            $row.TopGpuProcess = [DesktopMonitor.Native]::NameOf($topPid)
            $row.TopGpuPct = [math]::Min(100.0, $topPct)
        }
    }
    if ($script:GpuMemoryCat) {
        $data = $script:GpuMemoryCat.ReadCategory()
        foreach ($pair in @(@('Dedicated Usage', 'GpuDedicatedMB'), @('Shared Usage', 'GpuSharedMB'))) {
            $c = $data[$pair[0]]
            if (-not $c) { continue }
            $sum = 0.0
            foreach ($inst in $c.Values) { $sum += $inst.RawValue }
            $row[$pair[1]] = $sum / 1MB
        }
    }
}

# ACPI thermal zones: temperature in Kelvin, and "% Passive Limit" below 100 means the
# zone is throttling the CPU to cool down. Many desktops don't report these.
$script:ThermalCat = $null
function Get-ThermalMetrics($row) {
    if (-not $script:ThermalCat) { return }
    $data = $script:ThermalCat.ReadCategory()
    $temp = $data['Temperature']
    if ($temp) {
        $max = $null
        foreach ($inst in $temp.Values) {
            if ($inst.RawValue -le 0) { continue }
            $c = $inst.RawValue - 273.15
            if ($null -eq $max -or $c -gt $max) { $max = $c }
        }
        $row.ThermalMaxC = $max
    }
    $limit = $data['% Passive Limit']
    if ($limit) {
        $min = $null
        foreach ($inst in $limit.Values) { if ($null -eq $min -or $inst.RawValue -lt $min) { $min = $inst.RawValue } }
        $row.ThermalLimitPct = $min
    }
}

function Initialize-CategoryReaders {
    if (Test-Group 'Gpu') {
        foreach ($pair in @(@('GPU Engine', 'GpuEngineCat'), @('GPU Adapter Memory', 'GpuMemoryCat'))) {
            try {
                if ([Diagnostics.PerformanceCounterCategory]::Exists($pair[0])) {
                    Set-Variable -Scope Script -Name $pair[1] -Value (New-Object Diagnostics.PerformanceCounterCategory($pair[0]))
                } else { Write-MonitorLog "Counter category '$($pair[0])' not found; GPU columns stay blank." 'WARN' }
            } catch { Write-MonitorLog "Cannot open '$($pair[0])': $($_.Exception.Message)" 'WARN' }
        }
        try { Get-GpuMetrics @{} } catch { }   # first reading to compare against
    }
    if (Test-Group 'Network') {
        try {
            $script:NetCat = New-Object Diagnostics.PerformanceCounterCategory('Network Interface')
            $row = @{}; Get-NetworkMetrics $row   # first reading to compare against
        } catch { $script:NetCat = $null; Write-MonitorLog "Cannot read the Network Interface counters: $($_.Exception.Message)" 'WARN' }
    }
    if (Test-Group 'Thermal') {
        try {
            if ([Diagnostics.PerformanceCounterCategory]::Exists('Thermal Zone Information')) {
                $script:ThermalCat = New-Object Diagnostics.PerformanceCounterCategory('Thermal Zone Information')
            } else { Write-MonitorLog 'This PC reports no thermal zones; temperature columns stay blank.' }
        } catch { Write-MonitorLog "Cannot open thermal zones: $($_.Exception.Message)" 'WARN' }
    }
}

function Get-PowerMetrics($row) {
    $p = [DesktopMonitor.Native]::Power()
    if (-not $p) { return }
    if ($p[0] -ne 255) { $row.OnAC = $p[0] }
    if ($p[2] -ne 128 -and $p[1] -ne 255) { $row.BatteryPct = $p[1] }
    $row.BatterySaver = $p[3]
}

# Who is at the screen, and the monitor's user's idle time, foreground app and "Not
# Responding" windows. On a shared PC the person at the screen can be someone else (fast
# user switching); the other columns then describe the monitor's user's background session.
function Get-SessionMetrics($row) {
    $row.ConsoleUser = [DesktopMonitor.Native]::ConsoleUser()
    $row.IdleSec = [math]::Round([DesktopMonitor.Native]::IdleSeconds())
    $fg = [DesktopMonitor.Native]::ForegroundPid()
    if ($fg -gt 0) { $row.ForegroundApp = [DesktopMonitor.Native]::NameOf($fg) }
    $hung = [DesktopMonitor.Native]::HungWindowPids()
    $row.HungAppCount = $hung.Count
    if ($hung.Count) {
        $row.HungApps = [string]::Join('; ', [string[]] @($hung | ForEach-Object { '{0}({1})' -f [DesktopMonitor.Native]::NameOf($_), $_ }))
    }
}

# Freezes shorter than the interval, from the background threads in
# [DesktopMonitor.Responsiveness]: the longest stall of the whole PC, and the slowest replies
# from the foreground window and the taskbar since the previous sample.
function Get-ResponsivenessMetrics($row) {
    $r = [DesktopMonitor.Responsiveness]::Read()
    $row.MaxStallMs = [math]::Round($r[0])
    $row.StallCount = [int] $r[1]
    if ($r[2] -ge 0) {
        $row.ForegroundResponseMs = [math]::Round($r[2])
        if ($r[3] -gt 0) { $row.ForegroundResponseApp = [DesktopMonitor.Native]::NameOf([int] $r[3]) }
    }
    if ($r[4] -ge 0) { $row.ShellResponseMs = [math]::Round($r[4]) }
}

# ---------------------------------------------------------------------------------------
# Connectivity: what a desktop depends on beyond its own network link. The checks run on
# background threads ([DesktopMonitor.Connectivity]); here their results since the previous
# sample are put in the row. Windows' own view of the network (from the Network List
# Manager, which also picks the firewall profile), VPN and DNS servers are read once a minute.
# ---------------------------------------------------------------------------------------

$script:VpnPattern = 'VPN|AnyConnect|Secure Client|GlobalProtect|PANGP|FortiClient|Fortinet|Pulse Secure|Ivanti|Juniper|' +
    'Check Point|Zscaler|WireGuard|OpenVPN|TAP-Windows|Wintun|SonicWall|NetExtender|Cloudflare WARP|Netskope|IP-HTTPS|IKEv2|Citrix Secure'
$script:NetState = $null    # NetworkCategory, WindowsConnectivity, Vpn, DnsServers
$script:DrivesDown = $null
$script:Nlm = $null

# A probe's results since the previous sample into a time column and a failures column.
# Failure reasons are added to $Errors. Returns the number of runs that finished.
function Read-Probe($Probe, $row, [string] $MsColumn, [string] $FailColumn, [string] $Label, $Errors) {
    if (-not $Probe) { return 0 }
    $r = $Probe.Read()
    if ($MsColumn -and $r[0] -ge 0) { $row[$MsColumn] = [math]::Round($r[0], 1) }
    if ($FailColumn -and ($r[3] -gt 0 -or $r[1] -gt 0)) { $row[$FailColumn] = $r[1] }
    if ($r[2]) { $Errors.Add("${Label}: $($r[2])") }
    return $r[3]
}

function Get-WindowsNetworkState {
    $state = @{}
    try {
        if (-not $script:Nlm) { $script:Nlm = [Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid] 'DCB00C01-570F-4A9B-8D69-199FDBA5723B')) }
        # NLM_CONNECTIVITY flags: 0x40/0x400 internet, 0x20/0x200 local network, 0x10/0x100 subnet only, 0x1/0x2 no traffic.
        $flags = [int] $script:Nlm.GetConnectivity()
        $state.WindowsConnectivity = if ($flags -band 0x440) { 'Internet' } elseif ($flags -band 0x330) { 'LocalNetwork' } elseif ($flags -band 0x3) { 'NoTraffic' } else { 'Disconnected' }
        # The category also picks the firewall profile. A domain PC on its office network
        # should be Domain; if it isn't, Windows didn't recognise the network, and the
        # Public or Private firewall rules apply instead.
        $cats = @(foreach ($n in $script:Nlm.GetNetworks(1)) { @('Public', 'Private', 'Domain')[[int] $n.GetCategory()] })
        $state.NetworkCategory = [string]::Join(', ', [string[]] @($cats | Sort-Object -Unique))
    } catch { Write-Throttled 'nlm' "Reading the Windows network state failed: $($_.Exception.Message)" }
    $vpn = New-Object 'System.Collections.Generic.List[string]'
    $dns = New-Object 'System.Collections.Generic.List[string]'
    foreach ($ni in [Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
        if ($ni.OperationalStatus -ne 'Up' -or $ni.NetworkInterfaceType -eq 'Loopback') { continue }
        $isVpn = $ni.NetworkInterfaceType -eq 'Ppp' -or $ni.Description -match $script:VpnPattern
        if ($isVpn) { $vpn.Add($ni.Description) }
        if (-not $isVpn -and ($ni.Speed -le 0 -or $ni.NetworkInterfaceType -eq 'Tunnel')) { continue }
        foreach ($a in $ni.GetIPProperties().DnsAddresses) {
            $text = $a.ToString()
            if ($text -like 'fec0:*' -or $dns.Contains($text)) { continue }   # fec0:... are unused IPv6 defaults
            $dns.Add($text)
        }
    }
    $state.Vpn = [string]::Join('; ', $vpn)
    $state.DnsServers = [string]::Join(' ', $dns)
    return $state
}

function Get-ConnectivityMetrics($row) {
    $errors = New-Object 'System.Collections.Generic.List[string]'
    $c = [DesktopMonitor.Connectivity]
    [void] (Read-Probe $c::Dns $row 'DnsLookupMs' 'DnsFailed' 'DNS' $errors)
    if ($c::Dc) {
        [void] (Read-Probe $c::Dc $row 'DcConnectMs' 'DcFailed' 'Domain controller' $errors)
        $row.DomainController = $c::DcText
    }
    $servers = $c::Servers
    $failed = 0
    for ($i = 0; $i -lt $servers.Count; $i++) {
        $r = $servers[$i].Read()
        if ($r[0] -ge 0) { $row["Server$($i + 1)Ms"] = [math]::Round($r[0], 1) }
        $failed += $r[1]
        if ($r[2]) { $errors.Add("Server: $($r[2])") }
    }
    if ($servers.Count) { $row.ServerFailed = $failed }
    [void] (Read-Probe $c::Http $row 'HttpMs' 'HttpFailed' 'Web' $errors)
    # The drive check runs once a minute; its result is repeated until the next one.
    if ($c::Drives) {
        $r = $c::Drives.Read()
        if ($r[3] -gt 0) { $script:DrivesDown = if ($r[1] -gt 0) { $r[2] } else { $null } }
        $row.MappedDrivesDown = $script:DrivesDown
    }
    $row.NetChanges = $c::TakeChanges()
    if ($null -eq $script:NetState -or (Test-Every $script:EveryMinute)) { $script:NetState = Get-WindowsNetworkState }
    foreach ($k in $script:NetState.Keys) { $row[$k] = $script:NetState[$k] }
    if ($errors.Count) { $row.ConnectivityErrors = [string]::Join('; ', $errors) }
}

# Security and monitoring agents (their CPU, I/O and memory come from the process snapshot
# in Get-ProcessMetrics), and how long a small file write and read takes.
function Get-SecurityMetrics($row) {
    $p = [DesktopMonitor.Connectivity]::File
    if (-not $p) { return }
    $r = $p.Read()
    if ($r[0] -ge 0) { $row.FileCheckMs = [math]::Round($r[0], 1) }
    if ($r[3] -gt 0 -or $r[1] -gt 0) { $row.FileCheckFailed = $r[1] }
    if ($r[2]) { Write-Throttled 'fileCheck' "The file access check failed: $($r[2])" }
}

function New-ProcessRow([datetime] $Time, [string] $Kind, $Rank, $p) {
    $started = $null
    if ($p.CreateTime -gt 0) { $started = [datetime]::FromFileTime($p.CreateTime) }
    @{
        Timestamp = $Time; Kind = $Kind; Rank = $Rank; Process = $p.Name; PID = $p.Id
        ParentPID = $p.ParentId; CpuPct = $p.CpuPct; IoKBps = $p.IoKBps; HardFaultsPerSec = $p.HardFaultsPerSec
        PrivateMB = $p.PrivateMB; WorkingSetMB = $p.WorkingSetMB; Handles = $p.Handles
        GdiObjects = $(if ($p.GdiObjects -ge 0) { $p.GdiObjects }); UserObjects = $(if ($p.UserObjects -ge 0) { $p.UserObjects })
        TcpConnections = $p.TcpConnections; UdpEndpoints = $p.UdpEndpoints; Threads = $p.Threads
        SessionId = $p.SessionId; StartedAt = $started; Services = $p.Services; Agent = $p.Agent
        LargestFreeBlockMB = $(if ($p.LargestFreeBlockMB -ge 0) { $p.LargestFreeBlockMB })
        FreeAddressSpaceMB = $(if ($p.FreeAddressSpaceMB -ge 0) { $p.FreeAddressSpaceMB })
    }
}

# Resources that leak or run out: ranking, metrics columns for the top process, and the
# ProcInfo field. The top process goes in every metrics row; the top N go in the
# processes file once a minute (Kind = the ranking name).
$script:ResourceTops = @(
    @('Memory',  'TopMemProcess',     'TopMemMB',          'PrivateMB'),
    @('Handles', 'TopHandleProcess',  'TopHandles',        'Handles'),
    @('Gdi',     'TopGdiProcess',     'TopGdiObjects',     'GdiObjects'),
    @('User',    'TopUserObjProcess', 'TopUserObjects',    'UserObjects'),
    @('Tcp',     'TopTcpProcess',     'TopTcpConnections', 'TcpConnections'),
    @('Udp',     'TopUdpProcess',     'TopUdpEndpoints',   'UdpEndpoints')
)

# Busy processes, written every sample: ranking, metrics columns for the top process, the
# ProcInfo field, and the least that counts as busy.
$script:BusyTops = @(
    @('Cpu',        'TopCpuProcess',       'TopCpuPct',           'CpuPct',           1),
    @('Io',         'TopIoProcess',        'TopIoKBps',           'IoKBps',           100),
    @('HardFaults', 'TopHardFaultProcess', 'TopHardFaultsPerSec', 'HardFaultsPerSec', 20)
)

# The 32-bit process with the least free address space, measured once a minute and
# repeated in the metrics rows in between.
$script:AddressSpaceWorst = $null

# Busy processes every sample (CPU >= 1%, I/O >= 100 KB/s or 20+ hard faults a second), the
# biggest users of each resource above, the 32-bit processes with the least address space
# left, and every start and exit seen between samples.
function Get-ProcessMetrics($row) {
    $s = [DesktopMonitor.Native]::SampleProcesses()
    $now = $row.Timestamp
    $row.Handles = $s.TotalHandles
    if ($s.GdiTotal -ge 0) { $row.GdiObjects = $s.GdiTotal; $row.UserObjects = $s.UserTotal }
    if ($s.TcpTotal -ge 0) { $row.TcpPortsInUse = $s.TcpTotal; $row.TcpTimeWait = $s.TcpTimeWait }
    if ($s.UdpTotal -ge 0) { $row.UdpEndpoints = $s.UdpTotal }

    $everyMinute = Test-Every $script:EveryMinute
    if ($everyMinute) {
        try {
            $measured = [DesktopMonitor.Native]::MeasureAddressSpace($s.All)
            $script:AddressSpaceWorst = $null
            if ($measured.Count) { $script:AddressSpaceWorst = $measured[0] }
            $rank = 0
            foreach ($p in $measured) {
                $rank++
                if ($rank -gt $TopProcessCount) { break }
                $script:ProcRows.Add((New-ProcessRow $now 'AddressSpace' $rank $p))
            }
        } catch { Write-Throttled 'addressSpace' "Measuring 32-bit address space failed: $($_.Exception.Message)" }
    }
    # Security and monitoring agents: their total load every sample, each one once a minute.
    if ((Test-Group 'Security') -and $s.HasDelta) {
        $totals = $null
        $agents = [DesktopMonitor.Native]::Agents($s.All, [ref] $totals)
        $row.AgentCpuPct = $totals[0]; $row.AgentIoKBps = $totals[1]; $row.AgentMemMB = $totals[2]
        if ($agents.Count) { $row.TopAgentProcess = '{0} ({1})' -f $agents[0].Name, $agents[0].Agent; $row.TopAgentCpuPct = $agents[0].CpuPct }
        if ($everyMinute) {
            $rank = 0
            foreach ($p in $agents) { $rank++; if ($rank -gt $TopProcessCount) { break }; $script:ProcRows.Add((New-ProcessRow $now 'Agent' $rank $p)) }
        }
    }

    if ($script:AddressSpaceWorst) {
        $row.Top32BitProcess = $script:AddressSpaceWorst.Name
        $row.LargestFreeBlockMB = $script:AddressSpaceWorst.LargestFreeBlockMB
        $row.FreeAddressSpaceMB = $script:AddressSpaceWorst.FreeAddressSpaceMB
    }

    foreach ($spec in $script:ResourceTops) {
        $top = [DesktopMonitor.Native]::Top($s.All, $spec[0], $TopProcessCount)
        if ($top.Count -and $top[0].($spec[3]) -gt 0) { $row[$spec[1]] = $top[0].Label(); $row[$spec[2]] = $top[0].($spec[3]) }
        if (-not $everyMinute) { continue }
        $rank = 0
        foreach ($p in $top) {
            if ($p.($spec[3]) -le 0) { break }
            $rank++
            $script:ProcRows.Add((New-ProcessRow $now $spec[0] $rank $p))
        }
    }
    if (-not $s.HasDelta) { return }

    foreach ($spec in $script:BusyTops) {
        $top = [DesktopMonitor.Native]::Top($s.All, $spec[0], $TopProcessCount)
        if ($top.Count -and $top[0].($spec[3]) -gt 0) { $row[$spec[1]] = $top[0].Label(); $row[$spec[2]] = $top[0].($spec[3]) }
        $rank = 0
        foreach ($p in $top) {
            if ($p.($spec[3]) -lt $spec[4]) { break }
            $rank++
            $script:ProcRows.Add((New-ProcessRow $now $spec[0] $rank $p))
        }
    }
    foreach ($p in $s.Started) { $script:ProcRows.Add((New-ProcessRow $now 'Started' $null $p)) }
    foreach ($p in $s.Exited)  { $script:ProcRows.Add((New-ProcessRow $now 'Exited' $null $p)) }
}

# ---------------------------------------------------------------------------------------
# Event log: all critical/error/warning entries from System, critical/error entries from
# Application, plus boot, shutdown, sleep, wake and Windows Error Reporting entries.
# ---------------------------------------------------------------------------------------

$script:EventTagById = @{
    'Microsoft-Windows-Kernel-Power|41'                   = 'UnexpectedReboot'
    'Microsoft-Windows-Kernel-Power|42'                   = 'Sleep'
    'Microsoft-Windows-Kernel-Power|107'                  = 'Wake'
    'Microsoft-Windows-Kernel-Power|506'                  = 'Sleep'   # entering Modern Standby
    'Microsoft-Windows-Kernel-Power|507'                  = 'Wake'    # leaving Modern Standby
    'Microsoft-Windows-Power-Troubleshooter|1'            = 'Wake'
    'EventLog|6008'                                       = 'UnexpectedShutdown'
    'Microsoft-Windows-Kernel-General|12'                 = 'Boot'
    'Microsoft-Windows-Kernel-General|13'                 = 'Shutdown'
    'User32|1074'                                         = 'RestartRequested'
    'Microsoft-Windows-WER-SystemErrorReporting|1001'     = 'BlueScreen'
    'Display|4101'                                        = 'GpuDriverReset'
    'Microsoft-Windows-Resource-Exhaustion-Detector|2004' = 'LowMemory'
    'Microsoft-Windows-Kernel-Processor-Power|37'         = 'CpuThrottled'
    'Application Error|1000'                              = 'AppCrash'
    'Application Hang|1002'                               = 'AppHang'
    '.NET Runtime|1026'                                   = 'AppCrash'
    'Windows Error Reporting|1001'                        = 'ErrorReport'
    'Service Control Manager|7031'                        = 'ServiceCrash'
    'Service Control Manager|7034'                        = 'ServiceCrash'
    'Tcpip|4231'                                          = 'PortExhaustion'
    'Tcpip|4266'                                          = 'PortExhaustion'
    'Tcpip|4199'                                          = 'IpAddressConflict'
    # Name resolution and Active Directory
    'Microsoft-Windows-DNS-Client|1014'                   = 'DnsTimeout'
    'NETLOGON|5719'                                       = 'NoDomainController'
    'NETLOGON|5783'                                       = 'SecureChannel'
    'NETLOGON|3210'                                       = 'SecureChannel'
    'NETLOGON|5722'                                       = 'SecureChannel'
    'NETLOGON|5723'                                       = 'SecureChannel'
    'NETLOGON|5805'                                       = 'SecureChannel'
    'Microsoft-Windows-GroupPolicy|1129'                  = 'GroupPolicyNoNetwork'
    'LsaSrv|40960'                                        = 'AuthFailure'
    'LsaSrv|40961'                                        = 'AuthFailure'
    # Other logs (read once a minute, see $script:ExtraLogs)
    'Microsoft-Windows-NetworkProfile|10000'              = 'NetworkConnected'
    'Microsoft-Windows-NetworkProfile|10001'              = 'NetworkDisconnected'
    'Microsoft-Windows-NCSI|4042'                         = 'ConnectivityChange'
    'Microsoft-Windows-WLAN-AutoConfig|8001'              = 'WifiConnected'
    'Microsoft-Windows-WLAN-AutoConfig|8002'              = 'WifiConnectFailed'
    'Microsoft-Windows-WLAN-AutoConfig|8003'              = 'WifiDisconnected'
    'Microsoft-Windows-Windows Defender|1000'             = 'DefenderScanStarted'
    'Microsoft-Windows-Windows Defender|1001'             = 'DefenderScanFinished'
    'Microsoft-Windows-Windows Defender|1002'             = 'DefenderScanStopped'
    'Microsoft-Windows-Windows Defender|1116'             = 'MalwareDetected'
    'Microsoft-Windows-Windows Defender|1117'             = 'MalwareAction'
    'Microsoft-Windows-Windows Defender|2000'             = 'DefenderUpdate'
    'Microsoft-Windows-Windows Defender|2001'             = 'DefenderUpdateFailed'
    'Microsoft-Windows-Windows Defender|3002'             = 'DefenderProblem'
    'Microsoft-Windows-Windows Defender|5001'             = 'DefenderOff'
    'RasClient|20227'                                     = 'VpnFailed'
}
$script:EventTagByProvider = @{
    'Microsoft-Windows-WHEA-Logger' = 'HardwareError'
    'disk' = 'DiskError'; 'Ntfs' = 'DiskError'; 'Microsoft-Windows-Ntfs' = 'DiskError'
    'volmgr' = 'DiskError'; 'stornvme' = 'DiskError'; 'storahci' = 'DiskError'
    'iaStorAC' = 'DiskError'; 'iaStorAVC' = 'DiskError'; 'Microsoft-Windows-StorPort' = 'DiskError'
    'nvlddmkm' = 'GpuError'; 'amdkmdag' = 'GpuError'; 'amdkmdap' = 'GpuError'; 'igfx' = 'GpuError'
    'Microsoft-Windows-GroupPolicy' = 'GroupPolicyFailed'
    'Microsoft-Windows-Time-Service' = 'TimeSync'
    'Microsoft-Windows-Security-Kerberos' = 'Kerberos'
    'Microsoft-Windows-Dhcp-Client' = 'Dhcp'
    'Microsoft-Windows-NDIS' = 'NetworkAdapter'
    'Microsoft-Windows-WLAN-AutoConfig' = 'Wifi'
}
# Network adapter drivers (Intel, Realtek, Broadcom, Qualcomm, MediaTek, ...) logging a
# link drop, reset or hang in the System log.
$script:NicDriverPattern = '^(e1\w*express|e1[a-z]\d+x64|e2f\w*|rt640x64|rtcx21\w*|rtwlan\w*|Netwtw\d+|Netwbw\d+|b57nd60a|bcmwl\w*|athw\w*|qcwlan\w*|mtkwl\w*|ixgbe\w*|i40e\w*|mlx\w*)$'

# Logs besides System and Application, read once a minute, with the event IDs kept (any
# level) or 'problems' (critical, error and warning only). A log a standard user can't read
# is skipped with a note in the monitor's .log.
$script:ExtraLogs = [ordered] @{
    'Microsoft-Windows-NetworkProfile/Operational'     = @(10000, 10001)
    'Microsoft-Windows-NCSI/Operational'               = @(4042)
    'Microsoft-Windows-WLAN-AutoConfig/Operational'    = @(8001, 8002, 8003)
    'Microsoft-Windows-Windows Defender/Operational'   = @(1000, 1001, 1002, 1116, 1117, 2000, 2001, 3002, 5001)
    'Microsoft-Windows-Dhcp-Client/Admin'              = 'problems'
}
$script:ExtraLogsOff = @{}
$script:EventLevels = @{ 0 = 'Information'; 1 = 'Critical'; 2 = 'Error'; 3 = 'Warning'; 4 = 'Information'; 5 = 'Verbose' }
$script:EventSince = $null   # on the first read, entries newer than this (UTC)
$script:EventLast = @{}      # log name -> newest RecordId already read

# Which entries to keep. This filtering is done here rather than in an XPath query
# because Windows has no index on event time: a filtered query scans the whole log, which
# costs 100+ ms every time even when nothing matches.
function Test-EventWanted([string] $Log, $e) {
    $level = [int] $e.Level
    $provider = $e.ProviderName
    if ($script:ExtraLogs.Contains($Log)) {
        $keep = $script:ExtraLogs[$Log]
        if ($keep -eq 'problems') { return ($level -ge 1 -and $level -le 3) }
        return ($keep -contains $e.Id)
    }
    if ($Log -eq 'System') {
        if ($level -ge 1 -and $level -le 3) {
            # DCOM permission warnings are constant noise on most PCs.
            return -not ($provider -eq 'Microsoft-Windows-DistributedCOM' -and ($e.Id -eq 10010 -or $e.Id -eq 10016))
        }
        if ($provider -eq 'Microsoft-Windows-Kernel-General') { return ($e.Id -eq 12 -or $e.Id -eq 13) }
        # Sleep and wake, including Modern Standby, so the report can tell sleep from a freeze.
        if ($provider -eq 'Microsoft-Windows-Kernel-Power') { return (@(42, 107, 506, 507) -contains $e.Id) }
        return ($provider -eq 'Microsoft-Windows-Power-Troubleshooter' -or $provider -eq 'User32')
    }
    if ($level -eq 1 -or $level -eq 2) { return $true }
    return (($provider -eq 'Windows Error Reporting' -and $e.Id -eq 1001) -or ($provider -eq 'RasClient' -and $e.Id -eq 20227))
}

# New wanted entries in one log, oldest first. Reads newest-to-oldest and stops at the
# first entry already seen, so each call only touches what was added since the last one.
function Read-NewEvents([string] $Log) {
    $found = New-Object 'System.Collections.Generic.List[object]'
    $query = New-Object Diagnostics.Eventing.Reader.EventLogQuery($Log, [Diagnostics.Eventing.Reader.PathType]::LogName)
    $query.ReverseDirection = $true
    $reader = New-Object Diagnostics.Eventing.Reader.EventLogReader($query)
    $last = $script:EventLast[$Log]
    $cutoff = $script:EventSince.ToLocalTime()
    $newest = $null
    try {
        while ($null -ne ($e = $reader.ReadEvent())) {
            if ($null -eq $newest) {
                $newest = $e.RecordId
                if ($null -ne $last -and $newest -lt $last) { $last = $null }   # the log was cleared
            }
            $stop = if ($null -ne $last) { $e.RecordId -le $last } else { $e.TimeCreated -lt $cutoff }
            if ($stop) { $e.Dispose(); break }
            if (Test-EventWanted $Log $e) { $found.Add($e) } else { $e.Dispose() }
        }
    } finally { $reader.Dispose() }
    if ($null -ne $newest) { $script:EventLast[$Log] = $newest }
    $found.Reverse()
    return , $found
}

function Get-EventTag($e) {
    $tag = $script:EventTagById["$($e.ProviderName)|$($e.Id)"]
    if (-not $tag) { $tag = $script:EventTagByProvider[$e.ProviderName] }
    if (-not $tag -and $e.ProviderName -match $script:NicDriverPattern) { $tag = 'NetworkAdapter' }
    if (-not $tag) {
        $tag = switch ([int] $e.Level) { 1 { 'Critical' } 2 { 'Error' } 3 { 'Warning' } default { 'Information' } }
    }
    return $tag
}

function Add-MonitorEvent([string] $Tag, [string] $Message) {
    $script:EventRows.Add(@{
        Timestamp = (Get-Date); Tag = $Tag; Level = 'Information'; Log = 'DesktopMonitor'
        Provider = 'DesktopMonitor'; EventId = 0; Message = $Message
    })
}

function Add-RowTag($row, [string] $Tag) {
    $tags = @(([string] $row['EventTags']) -split ';' | Where-Object { $_ })
    if ($tags -notcontains $Tag) { $row['EventTags'] = [string]::Join(';', [string[]] ($tags + $Tag)) }
}

# ---------------------------------------------------------------------------------------
# Freeze reports (-Mark). While the monitor runs, -Mark leaves a small mark_*.txt file in
# the log folder and the monitor turns it into an events row, so the monitor stays the only
# process writing to its logs. The file holds the time, the user and the note.
# ---------------------------------------------------------------------------------------

function New-MarkEvent([datetime] $Time, [string] $User, [string] $Text) {
    $msg = if ($Text) { "Reported by ${User}: $Text" } else { "Reported by $User." }
    @{ Timestamp = $Time; Tag = 'UserReportedFreeze'; Level = 'Information'; Log = 'DesktopMonitor'
       Provider = 'DesktopMonitor'; EventId = 0; Message = $msg }
}

$script:MarksSeen = New-Object 'System.Collections.Generic.HashSet[string]'
function Read-FreezeMarks($row) {
    foreach ($path in [IO.Directory]::GetFiles($LogDirectory, 'mark_*.txt')) {
        if (-not $script:MarksSeen.Add($path)) { continue }
        try {
            $lines = @([IO.File]::ReadAllLines($path, $script:Utf8)) + @('', '', '')
            $time = [datetime]::ParseExact($lines[0], 'yyyy-MM-dd HH:mm:ss', $script:Inv)
            $script:EventRows.Add((New-MarkEvent $time $lines[1] $lines[2]))
            Add-RowTag $row 'UserReportedFreeze'
            Write-MonitorLog "Freeze reported by $($lines[1]) at $($lines[0])."
        } catch { Write-Throttled 'mark' "Could not read the freeze report ${path}: $($_.Exception.Message)" }
        # On a shared PC, another user's file can't be deleted; MarksSeen stops it being logged twice.
        try { [IO.File]::Delete($path) } catch { }
    }
}

# A full path without a trailing \ (except for a drive root), however it was typed, so that
# C:\Logs, c:\logs\ and .\Logs (from C:\) all mean the same folder.
function ConvertTo-FullPath([string] $Path) {
    $full = [IO.Path]::GetFullPath([IO.Path]::Combine((Get-Location -PSProvider FileSystem).ProviderPath, $Path))
    if ($full.Length -gt 3) { $full = $full.TrimEnd('\') }
    return $full
}

# One name per log folder. Kernel object names are limited to 260 characters, so a long
# path is replaced by its hash.
function Get-MonitorMutexName {
    $key = (ConvertTo-FullPath $LogDirectory).ToLowerInvariant()
    if ($key.Length -gt 200) {
        $sha = [Security.Cryptography.SHA256]::Create()
        $key = -join @($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($key)) | ForEach-Object { $_.ToString('x2') })
    }
    'Global\DesktopMonitor_' + ($key -replace '[\\/:]', '_')
}

# On a shared PC each user's monitor, and each user's 'Report a freeze', writes to the same
# folder. A file belongs to whoever created it, so without this one user's monitor couldn't
# add to, zip or delete another user's files, or clear their freeze reports. Done only when
# this run created the folder or it is the default one, and only by its owner.
function Grant-LogFolderAccess([string] $Path) {
    try {
        $dir = New-Object IO.DirectoryInfo($Path)
        $acl = $dir.GetAccessControl([Security.AccessControl.AccessControlSections]::Access)
        $users = New-Object Security.Principal.SecurityIdentifier('S-1-5-32-545')   # BUILTIN\Users
        $inherit = [Security.AccessControl.InheritanceFlags] 'ContainerInherit, ObjectInherit'
        $modify = [Security.AccessControl.FileSystemRights]::Modify
        foreach ($r in $acl.GetAccessRules($true, $true, [Security.Principal.SecurityIdentifier])) {
            if ($r.IdentityReference -eq $users -and $r.AccessControlType -eq 'Allow' -and
                ($r.FileSystemRights -band $modify) -eq $modify -and ($r.InheritanceFlags -band $inherit) -eq $inherit) { return }
        }
        $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($users, $modify, $inherit, 'None', 'Allow')))
        $dir.SetAccessControl($acl)
        Write-MonitorLog "Let every user of this PC change the files in $Path, so each user's monitor can add to, zip and tidy them."
    } catch { }   # not the folder's owner: the owner's monitor does this when it next starts
}

function Test-MonitorRunning {
    $m = $null
    try {
        if ([Threading.Mutex]::TryOpenExisting((Get-MonitorMutexName), [ref] $m)) { $m.Dispose(); return $true }
        return $false
    } catch { return $true }   # it exists, but belongs to another user's monitor
}

function Add-FreezeMark {
    $now = Get-Date
    $user = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    $text = (([string] $Note) -replace '\s+', ' ').Trim()
    if ($text.Length -gt 300) { $text = $text.Substring(0, 300) }
    $when = $now.ToString('HH:mm:ss', $script:Inv)
    try {
        if (Test-MonitorRunning) {
            $name = 'mark_{0}_{1}' -f $now.ToString('yyyyMMdd-HHmmss-fff', $script:Inv), ($user -split '\\')[-1]
            $tmp = Join-Path $LogDirectory "$name.tmp"
            # Written under another name first, so the monitor never reads half a file.
            [IO.File]::WriteAllLines($tmp, [string[]] @($now.ToString('yyyy-MM-dd HH:mm:ss', $script:Inv), $user, $text), $script:Utf8)
            # The monitor may run as another user on a shared PC. It deletes the file once it
            # has logged it; if it couldn't, it would log it again each time it restarts.
            try {
                $acl = [IO.File]::GetAccessControl($tmp, [Security.AccessControl.AccessControlSections]::Access)
                $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule(
                    (New-Object Security.Principal.SecurityIdentifier('S-1-5-32-545')), 'Read, Delete', 'Allow')))
                [IO.File]::SetAccessControl($tmp, $acl)
            } catch { }
            Move-Item -LiteralPath $tmp -Destination (Join-Path $LogDirectory "$name.txt") -Force
            $msg = "Thanks. The freeze at $when has been recorded."
        } else {
            $table = New-LogTable 'events' $script:EventColumns
            Write-TableRows $table @(New-MarkEvent $now $user $text)
            if ($table.Pending.Count) { throw "Could not write to the events log in $LogDirectory." }
            $msg = "The freeze at $when has been recorded, but DesktopMonitor isn't running, so there are no measurements to go with it."
        }
        $icon = 64   # information
    } catch {
        $msg = "The freeze could not be recorded: $($_.Exception.Message)"
        $icon = 48   # warning
    }
    Write-Host $msg
    if ($Popup) { try { [void] (New-Object -ComObject WScript.Shell).Popup($msg, 8, 'DesktopMonitor', $icon) } catch { } }
}

function Get-EventMetrics($row) {
    if ($null -eq $script:EventSince) { $script:EventSince = [datetime]::UtcNow }
    # The first read catches up on everything since the monitor last ran, which can take
    # seconds. That isn't the monitor being held up, so it isn't counted in CollectMs.
    $clock = [Diagnostics.Stopwatch]::StartNew()
    try { Read-EventRows $row }
    finally { if ($script:SampleNo -eq 1) { $script:WaitMs += $clock.Elapsed.TotalMilliseconds } }
}

# Shutdowns (User32 1074, Kernel-General 13) and sleeps (Kernel-Power 42, 506) in the System
# log since a time. With Fast Startup on (the Windows default), Shut down hibernates the
# kernel instead of restarting it, so the boot time stays the same; these show it happened.
function Get-ShutdownsAndSleeps([datetime] $Since) {
    $utc = $Since.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.000Z', $script:Inv)
    $xpath = "*[System[TimeCreated[@SystemTime>='$utc'] and (" +
        "(Provider[@Name='User32'] and EventID=1074) or " +
        "(Provider[@Name='Microsoft-Windows-Kernel-General'] and EventID=13) or " +
        "(Provider[@Name='Microsoft-Windows-Kernel-Power'] and (EventID=42 or EventID=506)))]]"
    $query = New-Object Diagnostics.Eventing.Reader.EventLogQuery('System', [Diagnostics.Eventing.Reader.PathType]::LogName, $xpath)
    $reader = New-Object Diagnostics.Eventing.Reader.EventLogReader($query)
    try {
        while ($null -ne ($e = $reader.ReadEvent())) {
            [pscustomobject] @{ Shutdown = ($e.Id -eq 1074 -or $e.Id -eq 13); Time = $e.TimeCreated }
            $e.Dispose()
        }
    } finally { $reader.Dispose() }
}

function Read-EventRows($row) {
    $tags = New-Object 'System.Collections.Generic.List[string]'
    $count = 0
    $logs = @('System', 'Application')
    if (Test-Every $script:EveryMinute) { $logs += @($script:ExtraLogs.Keys | Where-Object { -not $script:ExtraLogsOff.ContainsKey($_) }) }
    foreach ($log in $logs) {
        $found = @()
        try { $found = Read-NewEvents $log }
        catch {
            if ($log -eq 'System' -or $log -eq 'Application') { throw }
            $script:ExtraLogsOff[$log] = $true   # e.g. only administrators may read it, or it doesn't exist here
            Write-MonitorLog "Not reading the $log event log: $($_.Exception.Message)"
            continue
        }
        foreach ($e in $found) {
            try {
                $tag = Get-EventTag $e
                $msg = $null
                try { $msg = $e.FormatDescription() } catch { }
                if (-not $msg) { $msg = [string]::Join(' | ', [string[]] @($e.Properties | ForEach-Object { $_.Value })) }
                $msg = ([string] $msg -replace '\s+', ' ').Trim()
                if ($msg.Length -gt 500) { $msg = $msg.Substring(0, 500) + '...' }
                $script:EventRows.Add(@{
                    Timestamp = $e.TimeCreated; Tag = $tag; Level = $script:EventLevels[[int] $e.Level]
                    Log = $log; Provider = $e.ProviderName; EventId = $e.Id; Message = $msg
                })
                $count++
                if (-not $tags.Contains($tag)) { $tags.Add($tag) }
            } finally { $e.Dispose() }
        }
    }
    $row.EventCount = $count
    if ($tags.Count) { $row.EventTags = [string]::Join(';', $tags) }
}

# ---------------------------------------------------------------------------------------
# One-off inventory, written each time the monitor starts. WMI is fine here.
# ---------------------------------------------------------------------------------------

function Write-Inventory([hashtable] $Bound) {
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $add = { param($Section, $Name, $Value) $rows.Add(@{ Section = $Section; Name = $Name; Value = $Value }) }
    $section = {
        param([string] $Name, [scriptblock] $Body)
        try { & $Body } catch { & $add $Name 'Error' $_.Exception.Message }
    }

    & $section 'Monitor' {
        & $add 'Monitor' 'Version' $script:Version
        & $add 'Monitor' 'Script' $PSCommandPath
        & $add 'Monitor' 'Parameters' ([string]::Join(' ', [string[]] @($Bound.GetEnumerator() | Where-Object { $_.Key -ne 'InventoryOnly' } | ForEach-Object { "-$($_.Key) $($_.Value -join ',')" })))
        & $add 'Monitor' 'IntervalSeconds' $IntervalSeconds
        & $add 'Monitor' 'PingTargets' ($PingTarget -join ', ')
        & $add 'Monitor' 'Excluded' ($Exclude -join ', ')
        & $add 'Monitor' 'RunningAs' ([Security.Principal.WindowsIdentity]::GetCurrent().Name)
        & $add 'Monitor' 'SessionId' $script:Self.SessionId
        & $add 'Monitor' 'PowerShell' $PSVersionTable.PSVersion.ToString()
        & $add 'Monitor' 'TimeZone' ([TimeZoneInfo]::Local.Id)
    }
    & $section 'OS' {
        $os = Get-CimInstance Win32_OperatingSystem
        $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
        # DisplayVersion (e.g. 23H2) only exists from Windows 10 20H2; older builds have ReleaseId.
        $release = foreach ($n in 'DisplayVersion', 'ReleaseId') { if ($cv.PSObject.Properties[$n]) { $cv.$n; break } }
        $ubr = if ($cv.PSObject.Properties['UBR']) { $cv.UBR } else { 0 }
        & $add 'OS' 'Name' $os.Caption
        & $add 'OS' 'Version' ('{0}.{1} ({2})' -f $os.Version, $ubr, $release)
        & $add 'OS' 'Architecture' $os.OSArchitecture
        & $add 'OS' 'Installed' $os.InstallDate
        & $add 'OS' 'LastBoot' $os.LastBootUpTime
        $hb = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' -ErrorAction SilentlyContinue)
        if ($hb -and ($hb.PSObject.Properties.Name -contains 'HiberbootEnabled')) { & $add 'OS' 'FastStartup' $hb.HiberbootEnabled }
        $pending = @()
        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $pending += 'ComponentServicing' }
        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $pending += 'WindowsUpdate' }
        & $add 'OS' 'PendingReboot' ($(if ($pending) { $pending -join ', ' } else { 'No' }))
    }
    & $section 'Hardware' {
        $cs = Get-CimInstance Win32_ComputerSystem
        $bios = Get-CimInstance Win32_BIOS
        & $add 'Hardware' 'Manufacturer' $cs.Manufacturer
        & $add 'Hardware' 'Model' $cs.Model
        & $add 'Hardware' 'SerialNumber' $bios.SerialNumber
        & $add 'Hardware' 'BIOS' ('{0} {1} ({2:yyyy-MM-dd})' -f $bios.Manufacturer, $bios.SMBIOSBIOSVersion, $bios.ReleaseDate)
        & $add 'Hardware' 'RAM_GB' ([math]::Round($cs.TotalPhysicalMemory / 1GB, 1))
        & $add 'Hardware' 'PageFileAutomatic' $cs.AutomaticManagedPagefile
        foreach ($cpu in @(Get-CimInstance Win32_Processor)) {
            & $add 'Hardware' 'CPU' ('{0} | {1} cores, {2} threads, {3} MHz' -f $cpu.Name.Trim(), $cpu.NumberOfCores, $cpu.NumberOfLogicalProcessors, $cpu.MaxClockSpeed)
        }
        foreach ($m in @(Get-CimInstance Win32_PhysicalMemory)) {
            & $add 'Hardware' 'MemoryModule' ('{0} GB {1} MHz {2} {3} ({4})' -f ($m.Capacity / 1GB), $m.Speed, $m.Manufacturer, ([string] $m.PartNumber).Trim(), $m.DeviceLocator)
        }
        foreach ($g in @(Get-CimInstance Win32_VideoController)) {
            & $add 'Hardware' 'GPU' ('{0} | driver {1} ({2:yyyy-MM-dd})' -f $g.Name, $g.DriverVersion, $g.DriverDate)
        }
        foreach ($pf in @(Get-CimInstance Win32_PageFileUsage)) {
            & $add 'Hardware' 'PageFile' ('{0} size {1} MB, peak use {2} MB' -f $pf.Name, $pf.AllocatedBaseSize, $pf.PeakUsage)
        }
    }
    & $section 'Disks' {
        foreach ($d in @(Get-PhysicalDisk -ErrorAction Stop)) {
            & $add 'Disks' 'PhysicalDisk' ('{0} | {1} {2} {3:0} GB | health {4}' -f $d.FriendlyName, $d.MediaType, $d.BusType, ($d.Size / 1GB), $d.HealthStatus)
        }
        foreach ($v in [IO.DriveInfo]::GetDrives()) {
            if ($v.DriveType -ne 'Fixed' -or -not $v.IsReady) { continue }
            & $add 'Disks' 'Volume' ('{0} {1} {2:0.0} GB free of {3:0.0} GB' -f $v.Name, $v.DriveFormat, ($v.AvailableFreeSpace / 1GB), ($v.TotalSize / 1GB))
        }
    }
    & $section 'Network' {
        foreach ($ni in @(Get-ActiveAdapters)) {
            $ip = $ni.GetIPProperties()
            $v4 = @($ip.UnicastAddresses | Where-Object { $_.Address.AddressFamily -eq 'InterNetwork' } | ForEach-Object { $_.Address.ToString() }) -join ' '
            $dhcp = try { $ip.GetIPv4Properties().IsDhcpEnabled } catch { '?' }
            & $add 'Network' 'Adapter' ('{0} | {1} | {2} | {3} Mbps | IP {4} | DHCP {5} | DNS {6} | suffix {7}' -f $ni.Name, $ni.Description,
                $ni.NetworkInterfaceType, [math]::Round($ni.Speed / 1e6), $v4, $dhcp, (@($ip.DnsAddresses | ForEach-Object { $_.ToString() }) -join ' '), $ip.DnsSuffix)
        }
        & $add 'Network' 'DefaultGateway' (Get-DefaultGateway)
        # Drivers and power saving: an old network driver, or "allow the computer to turn off
        # this device", are common causes of dropouts.
        foreach ($a in @(Get-NetAdapter -Physical -ErrorAction Stop)) {
            $pm = try { (Get-NetAdapterPowerManagement -Name $a.Name -ErrorAction Stop).AllowComputerToTurnOffDevice } catch { '?' }
            & $add 'Network' 'AdapterDriver' ('{0} | {1} | {2} {3} ({4:yyyy-MM-dd}) | status {5} | power saving {6}' -f $a.Name, $a.InterfaceDescription,
                $a.DriverProvider, $a.DriverVersionString, $a.DriverDate, $a.Status, $pm)
        }
        foreach ($ni in [Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
            if ($ni.NetworkInterfaceType -eq 'Ppp' -or $ni.Description -match $script:VpnPattern) {
                & $add 'Network' 'VpnAdapter' ('{0} | {1} | {2}' -f $ni.Name, $ni.Description, $ni.OperationalStatus)
            }
        }
    }
    & $section 'Proxy' {
        # WinHTTP is used by Windows services (updates, Defender, Office licensing); the user
        # settings by browsers and most apps. A PAC script or WPAD that is slow to load makes
        # every new connection slow.
        $winhttp = (& "$env:SystemRoot\System32\netsh.exe" winhttp show proxy) | Where-Object { $_ -match '\S' -and $_ -notmatch 'Current WinHTTP' }
        & $add 'Proxy' 'WinHttp' (($winhttp | ForEach-Object { $_.Trim() }) -join ' ')
        $is = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
        foreach ($n in 'ProxyEnable', 'ProxyServer', 'ProxyOverride', 'AutoConfigURL') {
            if ($is.PSObject.Properties[$n]) { & $add 'Proxy' "User$n" $is.$n }
        }
        $conn = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections' -ErrorAction SilentlyContinue
        if ($conn -and $conn.PSObject.Properties['DefaultConnectionSettings']) {
            & $add 'Proxy' 'UserAutoDetect' ([bool] ($conn.DefaultConnectionSettings[8] -band 8))   # "Automatically detect settings" (WPAD)
        }
        foreach ($k in @(Get-ChildItem 'HKCU:\Network' -ErrorAction SilentlyContinue)) {
            & $add 'Proxy' 'MappedDrive' ('{0}: {1}' -f $k.PSChildName.ToUpper(), (Get-ItemProperty $k.PSPath).RemotePath)
        }
    }
    & $section 'Domain' {
        $cs = Get-CimInstance Win32_ComputerSystem
        & $add 'Domain' 'PartOfDomain' $cs.PartOfDomain
        & $add 'Domain' 'Domain' $cs.Domain
        & $add 'Domain' 'LogonServer' $env:LOGONSERVER
        if ($cs.PartOfDomain) {
            # The domain controller and AD site in use (a DC in another site means sign-in,
            # Group Policy and file access cross the WAN).
            $dc = & "$env:SystemRoot\System32\nltest.exe" "/dsgetdc:$($cs.Domain)" 2>&1
            foreach ($l in $dc) { if ($l -match '^\s*(DC|Address|Dom Name|Dc Site Name|Our Site Name):\s*(.+)$') { & $add 'Domain' ($Matches[1] -replace ' ', '') $Matches[2].Trim() } }
        }
        # Entra ID (Azure AD) and hybrid join state.
        $ds = & "$env:SystemRoot\System32\dsregcmd.exe" /status 2>$null
        foreach ($l in $ds) { if ($l -match '^\s*(AzureAdJoined|DomainJoined|EnterpriseJoined|WorkplaceJoined|AzureAdPrt|TenantName)\s*:\s*(.+)$') { & $add 'Domain' $Matches[1] $Matches[2].Trim() } }
        $src = (@(& "$env:SystemRoot\System32\w32tm.exe" /query /source 2>$null) | Where-Object { $_ -match '\S' }) -join ' '
        if ($src -and $src -notmatch 'error|denied') { & $add 'Domain' 'TimeSource' $src.Trim() }   # needs admin on some PCs
    }
    & $section 'Security' {
        # What Windows Security Center knows: on/off and up to date for each product.
        foreach ($c in 'AntiVirusProduct', 'AntiSpywareProduct', 'FirewallProduct') {
            foreach ($p in @(Get-CimInstance -Namespace 'root\SecurityCenter2' -ClassName $c -ErrorAction SilentlyContinue)) {
                $s = [int] $p.productState
                & $add 'Security' $c ('{0} | {1} | {2}' -f $p.displayName, $(if ((($s -shr 12) -band 0xF) -eq 1) { 'on' } else { 'off' }),
                    $(if ((($s -shr 4) -band 0xF) -eq 0) { 'up to date' } else { 'OUT OF DATE' }))
            }
        }
        try {
            $mp = Get-MpComputerStatus -ErrorAction Stop
            & $add 'Security' 'Defender' ('mode {0} | real-time {1} | signatures {2:yyyy-MM-dd HH:mm} | last quick scan {3:yyyy-MM-dd HH:mm} | last full scan {4:yyyy-MM-dd HH:mm}' -f
                $mp.AMRunningMode, $mp.RealTimeProtectionEnabled, $mp.AntivirusSignatureLastUpdated, $mp.QuickScanEndTime, $mp.FullScanEndTime)
        } catch { }
        $company = @{}
        $companyOf = { param($path) if (-not $path) { return '' }; if (-not $company.ContainsKey($path)) { $company[$path] = try { [string] [Diagnostics.FileVersionInfo]::GetVersionInfo($path).CompanyName } catch { '' } }; $company[$path] }
        $exeOf = { param($cmd) $cmd = [Environment]::ExpandEnvironmentVariables([string] $cmd).Trim(); if ($cmd -match '^"([^"]+)"') { $Matches[1] } elseif ($cmd -match '^(.+?\.(exe|sys))\b') { $Matches[1] -replace '^\\SystemRoot', $env:SystemRoot -replace '^\\\?\?\\', '' } else { $cmd } }
        # Kernel drivers not from Microsoft. File-system filter drivers (with an altitude) see
        # every file opened on the PC: 320000-329999 antivirus, 360000-389999 activity
        # monitoring (EDR), 260000-269999 content screening (DLP), 140000-149999 encryption.
        foreach ($d in @(Get-CimInstance Win32_SystemDriver -Filter "State = 'Running'")) {
            $path = & $exeOf $d.PathName
            $co = & $companyOf $path
            if ($co -match 'Microsoft' -or -not $co) { continue }
            $alt = Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Services\$($d.Name)\Instances" -ErrorAction SilentlyContinue |
                ForEach-Object { (Get-ItemProperty $_.PSPath).Altitude } | Where-Object { $_ } | Select-Object -First 1
            if ($alt) { & $add 'Security' 'FilterDriver' ('{0} | {1} | {2} | altitude {3}' -f $d.Name, $d.DisplayName, $co, $alt) }
            else { & $add 'ThirdParty' 'Driver' ('{0} | {1} | {2}' -f $d.Name, $d.DisplayName, $co) }
        }
        # Every running service not from Microsoft (security, management and hardware agents).
        foreach ($sv in @(Get-CimInstance Win32_Service -Filter "State = 'Running'")) {
            $co = & $companyOf (& $exeOf $sv.PathName)
            if ($co -and $co -notmatch 'Microsoft') { & $add 'ThirdParty' 'Service' ('{0} | {1} | {2}' -f $sv.Name, $sv.DisplayName, $co) }
        }
        # Security products load their DLLs into every program to watch it; these are the
        # ones loaded into this PowerShell.
        foreach ($m in (Get-Process -Id $PID).Modules) {
            if ($m.Company -and $m.Company -notmatch 'Microsoft') { & $add 'Security' 'InjectedDll' ('{0} | {1} | {2}' -f $m.ModuleName, $m.Company, $m.FileVersion) }
        }
        # Antimalware Scan Interface providers check every PowerShell script and Office macro.
        foreach ($k in @(Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\AMSI\Providers' -ErrorAction SilentlyContinue)) {
            $dll = (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\CLSID\$($k.PSChildName)\InprocServer32" -ErrorAction SilentlyContinue).'(default)'
            & $add 'Security' 'AmsiProvider' ('{0} | {1}' -f $dll, (& $companyOf ($dll -replace '"', '')))
        }
    }
    & $section 'Software' {
        $plan = (& "$env:SystemRoot\System32\powercfg.exe" /getactivescheme 2>$null) -join ' '
        & $add 'Software' 'PowerPlan' ($plan -replace '^.*?:\s*', '')
        foreach ($av in @(Get-CimInstance -Namespace 'root\SecurityCenter2' -ClassName AntiVirusProduct -ErrorAction SilentlyContinue)) {
            & $add 'Software' 'Antivirus' $av.displayName
        }
        foreach ($hf in @(Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 5)) {
            & $add 'Software' 'RecentUpdate' ('{0} {1:yyyy-MM-dd}' -f $hf.HotFixID, $hf.InstalledOn)
        }
    }

    $path = Join-Path $LogDirectory ('inventory_{0}_{1}.{2}' -f $script:PC, (Get-Date).ToString('yyyyMMdd-HHmmss', $script:Inv), $script:Ext)
    $table = New-LogTable 'inventory' @('Section', 'Name', 'Value')
    $lines = @($table.Header) + @($rows | ForEach-Object { [string]::Join($script:Sep, [string[]] @((ConvertTo-Field $_.Section), (ConvertTo-Field $_.Name), (ConvertTo-Field $_.Value))) })
    [void] (Add-FileLines $path $lines -Bom)
    Write-MonitorLog "Wrote inventory to $path"
}

# ---------------------------------------------------------------------------------------
# Install / uninstall, for the current user (no admin rights)
# ---------------------------------------------------------------------------------------

# Turns bound parameters back into argument text for running this script again (the
# scheduled task, the inventory process). Run it with -Command, not -File, so that lists
# such as -Exclude Gpu,Thermal stay lists. Sorted, so the same settings always give the
# same text whatever order they were typed in.
function ConvertTo-ArgumentText([hashtable] $Bound, [string[]] $Skip) {
    $parts = foreach ($key in @($Bound.Keys | Sort-Object)) {
        if ($Skip -contains $key) { continue }
        $value = $Bound[$key]
        if ($value -is [switch]) { if ($value) { "-$key" } }
        elseif ($value -is [int]) { "-$key $value" }
        else { "-$key " + (@($value | ForEach-Object { "'" + ([string] $_).Replace("'", "''") + "'" }) -join ',') }
    }
    return [string]::Join(' ', [string[]] @($parts))
}

# What an earlier -Install set up for this user: the scheduled task's launch arguments,
# or failing that the Run key's command line. $null when it isn't installed.
function Get-InstalledLaunch {
    $task = Get-ScheduledTask -TaskName $script:TaskName -ErrorAction SilentlyContinue
    if ($task) { return [pscustomobject] @{ Via = 'Task'; Command = [string] $task.Actions[0].Arguments } }
    $run = Get-ItemProperty -Path $script:RunKey -Name 'DesktopMonitor' -ErrorAction SilentlyContinue
    if ($run) { return [pscustomobject] @{ Via = 'RunKey'; Command = [string] $run.DesktopMonitor } }
    return $null
}

# Splits a launch command into the installed script's path and its settings text.
function Split-Launch([string] $Command) {
    if ($Command -match "& '((?:[^']|'')+)'\s*(.*?)`"\s*$") { return @($Matches[1].Replace("''", "'"), $Matches[2]) }
    return @($null, '')
}

# Running monitors started from any of these script paths (not the inventory process).
function Get-MonitorProcesses([string[]] $Paths) {
    $Paths = @($Paths | Where-Object { $_ })
    @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" | Where-Object {
        $cmd = [string] $_.CommandLine
        $_.ProcessId -ne $PID -and $cmd -notlike '*-InventoryOnly*' -and $cmd -notlike '* -Mark*' -and
            @($Paths | Where-Object { $cmd.IndexOf($_, [StringComparison]::OrdinalIgnoreCase) -ge 0 }).Count -gt 0
    })
}

# Stops the installed monitor. A monitor started by hand from elsewhere is left alone.
function Stop-InstalledMonitor([string[]] $Paths) {
    if (Get-ScheduledTask -TaskName $script:TaskName -ErrorAction SilentlyContinue) {
        Stop-ScheduledTask -TaskName $script:TaskName -ErrorAction SilentlyContinue
    }
    # Stopping the task doesn't end the PowerShell it started under conhost, so end that too.
    foreach ($p in (Get-MonitorProcesses $Paths)) {
        try { Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop; Write-Host "  Stopped the running monitor (process $($p.ProcessId))." } catch { }
    }
}

function Start-InstalledMonitor([string] $Via, [string] $HiddenArgs) {
    if ($Via -eq 'Task') { Start-ScheduledTask -TaskName $script:TaskName }
    else { Start-Process -FilePath "$env:SystemRoot\System32\conhost.exe" -ArgumentList $HiddenArgs }
}

function Format-Settings([string] $Text) { if ($Text) { $Text } else { '(all defaults)' } }

# The 'Report a freeze' desktop shortcut runs -Mark with no window at all (like the
# monitor itself) and confirms with a popup. It passes the same settings as the monitor, so
# the report goes to the same log folder.
function Get-FreezeShortcutPath { Join-Path ([Environment]::GetFolderPath('Desktop')) 'Report a freeze.lnk' }

function Set-FreezeShortcut([string] $ScriptPath, [string] $SettingsText) {
    $path = Get-FreezeShortcutPath
    $existed = Test-Path -LiteralPath $path
    $ps = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    $command = ("& '{0}' -Mark -Popup {1}" -f $ScriptPath.Replace("'", "''"), $SettingsText).TrimEnd()
    $conhost = "$env:SystemRoot\System32\conhost.exe"
    $arguments = '--headless "{0}" -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -Command "{1}"' -f $ps, $command
    $link = (New-Object -ComObject WScript.Shell).CreateShortcut($path)   # loads the existing one, if any
    if ($existed -and $link.TargetPath -eq $conhost -and $link.Arguments -eq $arguments) { return }
    $link.TargetPath = $conhost
    $link.Arguments = $arguments
    $link.Description = 'Tell DesktopMonitor that the PC just froze or was slow'
    $link.IconLocation = "$env:SystemRoot\System32\perfmon.exe,0"
    $link.Save()
    if ($existed) { Write-Host "  Updated the 'Report a freeze' shortcut on the desktop." }
    else { Write-Host "  Added a 'Report a freeze' shortcut to the desktop: $path" }
}

# Removes the shortcut only if it is the one -Install made.
function Remove-FreezeShortcut {
    $path = Get-FreezeShortcutPath
    if (-not (Test-Path -LiteralPath $path)) { return }
    $link = (New-Object -ComObject WScript.Shell).CreateShortcut($path)
    if ($link.Arguments -like '*DesktopMonitor.ps1*-Mark*') {
        Remove-Item -LiteralPath $path -Force
        Write-Host "  Removed the 'Report a freeze' shortcut from the desktop."
    }
}

function Update-FreezeShortcut([string] $ScriptPath, [string] $SettingsText) {
    try {
        if ($FreezeShortcut) { Set-FreezeShortcut $ScriptPath $SettingsText } else { Remove-FreezeShortcut }
    } catch { Write-Host "  Could not update the 'Report a freeze' shortcut: $($_.Exception.Message)" -ForegroundColor Yellow }
}

# The settings are the parameters given with -Install, and a new -Install replaces them
# all. Installing again with the same settings and script changes nothing.
function Install-Monitor([hashtable] $Bound) {
    # Saved as full paths: the scheduled task and the shortcut don't start in this folder.
    $LogDirectory = ConvertTo-FullPath $LogDirectory
    if ($Bound.ContainsKey('LogDirectory')) { $Bound['LogDirectory'] = $LogDirectory }
    $created = -not (Test-Path -LiteralPath $LogDirectory)
    New-Item -ItemType Directory -Force -Path $InstallDirectory, $LogDirectory | Out-Null
    if ($created) { Grant-LogFolderAccess $LogDirectory }
    $target = Join-Path (Resolve-Path -LiteralPath $InstallDirectory).Path 'DesktopMonitor.ps1'
    $settingsText = ConvertTo-ArgumentText $Bound @('Install', 'Uninstall', 'InstallDirectory', 'Live', 'InventoryOnly', 'FreezeShortcut', 'Mark', 'Note', 'Popup')
    $ps = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    $command = ("& '{0}' {1}" -f $target.Replace("'", "''"), $settingsText).TrimEnd()
    $psArgs = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -Command "{0}"' -f $command
    # conhost --headless runs PowerShell in the user's session with no window at all.
    $conhost = "$env:SystemRoot\System32\conhost.exe"
    $hiddenArgs = "--headless `"$ps`" $psArgs"
    $runValue = "`"$conhost`" $hiddenArgs"

    $before = Get-InstalledLaunch
    $oldPath = $null; $oldSettings = ''
    if ($before) { $oldPath, $oldSettings = Split-Launch $before.Command }
    $sameScript = (Test-Path -LiteralPath $target) -and
        ((Get-FileHash -LiteralPath $target).Hash -eq (Get-FileHash -LiteralPath $PSCommandPath).Hash)
    $sameLaunch = $before -and (($before.Via -eq 'Task' -and $before.Command -eq $hiddenArgs) -or
                                ($before.Via -eq 'RunKey' -and $before.Command -eq $runValue))

    if ($sameScript -and $sameLaunch) {
        Write-Host "DesktopMonitor is already installed with these settings: $(Format-Settings $settingsText)"
        if (@(Get-MonitorProcesses @($target)).Count) { Write-Host '  It is running, so it was left alone.' }
        else { Start-InstalledMonitor $before.Via $hiddenArgs; Write-Host '  It was not running, so it has been started.' }
        # The shortcut isn't part of the monitor's launch settings, so it is checked here too.
        Update-FreezeShortcut $target $settingsText
        return
    }

    if ($before) {
        Write-Host 'Updating the installed DesktopMonitor.'
        if ($oldSettings -ne $settingsText) {
            Write-Host "  Settings were: $(Format-Settings $oldSettings)"
            Write-Host "  Settings now:  $(Format-Settings $settingsText)"
        }
        if (-not $sameScript) {
            if (Test-Path -LiteralPath $target) { Write-Host '  Copying in the new version of the script.' }
            else { Write-Host "  Copying the script to the new install folder." }
        }
    }
    # Stop the old monitor first; otherwise it keeps the log folder and the new one exits.
    Stop-InstalledMonitor @($oldPath, $target)
    if ((Resolve-Path -LiteralPath $PSCommandPath).Path -ne $target) { Copy-Item -LiteralPath $PSCommandPath -Destination $target -Force }
    if ($oldPath -and $oldPath -ne $target -and $oldPath -ne $PSCommandPath -and (Test-Path -LiteralPath $oldPath)) {
        Remove-Item -LiteralPath $oldPath -Force
        Write-Host "  Deleted the copy in the old install folder: $oldPath"
    }

    $who = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    $action = New-ScheduledTaskAction -Execute $conhost -Argument $hiddenArgs
    # Start at logon, and every 5 minutes start it again if it isn't running.
    $triggers = @(
        (New-ScheduledTaskTrigger -AtLogOn -User $who),
        (New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 5))
    )
    $principal = New-ScheduledTaskPrincipal -UserId $who -LogonType Interactive -RunLevel Limited
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
        -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) -MultipleInstances IgnoreNew
    $via = 'Task'
    try {
        # -Force replaces an existing task, so there is only ever one per user.
        Register-ScheduledTask -TaskName $script:TaskName -Action $action -Trigger $triggers -Principal $principal -Settings $settings `
            -Description 'DesktopMonitor: logs PC health metrics to help troubleshoot crashes, freezes and slowdowns.' -Force -ErrorAction Stop | Out-Null
        Remove-ItemProperty -Path $script:RunKey -Name 'DesktopMonitor' -ErrorAction SilentlyContinue
        Write-Host "Installed the '$($script:TaskName)' scheduled task: it starts at logon and is checked every 5 minutes."
    } catch {
        # Some organisations stop users creating scheduled tasks. The Run key needs no
        # permission and starts the monitor at each logon, but nothing restarts it if it stops.
        $via = 'RunKey'
        Write-Host "Could not create a scheduled task ($($_.Exception.Message))." -ForegroundColor Yellow
        Unregister-ScheduledTask -TaskName $script:TaskName -Confirm:$false -ErrorAction SilentlyContinue   # an old one would start the old settings
        New-ItemProperty -Path $script:RunKey -Name 'DesktopMonitor' -Value $runValue -PropertyType String -Force | Out-Null
        Write-Host 'Added DesktopMonitor to the startup programs instead (HKCU Run key). It starts at each logon, but nothing restarts it if it stops.'
    }
    Start-InstalledMonitor $via $hiddenArgs
    Write-Host "  Running with: $(Format-Settings $settingsText)"
    Write-Host "  Script:       $target"
    Write-Host "  Logs:         $LogDirectory"
    Update-FreezeShortcut $target $settingsText
}

function Uninstall-Monitor {
    $before = Get-InstalledLaunch
    $oldPath = $null; $oldSettings = ''
    if ($before) { $oldPath, $oldSettings = Split-Launch $before.Command }
    $target = Join-Path $InstallDirectory 'DesktopMonitor.ps1'
    if (-not $before) { Write-Host 'DesktopMonitor is not installed for this user.' }
    Stop-InstalledMonitor @($oldPath, $target)
    if (Get-ScheduledTask -TaskName $script:TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $script:TaskName -Confirm:$false
        Write-Host "Removed the '$($script:TaskName)' scheduled task."
    }
    if (Get-ItemProperty -Path $script:RunKey -Name 'DesktopMonitor' -ErrorAction SilentlyContinue) {
        Remove-ItemProperty -Path $script:RunKey -Name 'DesktopMonitor'
        Write-Host 'Removed DesktopMonitor from the startup programs (HKCU Run key).'
    }
    foreach ($copy in @(@($oldPath, $target) | Where-Object { $_ } | Select-Object -Unique)) {
        if ((Test-Path -LiteralPath $copy) -and (Resolve-Path -LiteralPath $copy).Path -ne $PSCommandPath) {
            Remove-Item -LiteralPath $copy -Force
            Write-Host "Deleted $copy"
        }
    }
    try { Remove-FreezeShortcut } catch { Write-Host "Could not remove the 'Report a freeze' shortcut: $($_.Exception.Message)" -ForegroundColor Yellow }
    $logs = if ($oldSettings -match "-LogDirectory '((?:[^']|'')+)'") { $Matches[1].Replace("''", "'") } else { $LogDirectory }
    Write-Host "The logs were kept in $logs"
}

# ---------------------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------------------

# Time of the newest row in the newest metrics file, to see what happened while stopped.
function Get-LastSampleTime {
    $latest = Get-ChildItem -LiteralPath $LogDirectory -Filter "metrics_$($script:PC)_*" -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $latest) { return $null }
    $fs = New-Object IO.FileStream($latest.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]'ReadWrite, Delete')
    try {
        $len = [int] [math]::Min($fs.Length, 16384)
        [void] $fs.Seek(-$len, [IO.SeekOrigin]::End)
        $buf = New-Object byte[] $len
        [void] $fs.Read($buf, 0, $len)
    } finally { $fs.Dispose() }
    $lines = @($script:Utf8.GetString($buf) -split "`r?`n" | Where-Object { $_ -match '^"?\d{4}-\d\d-\d\d \d\d:\d\d:\d\d' })
    if (-not $lines.Count) { return $null }
    return [datetime]::ParseExact(($lines[-1].TrimStart('"')).Substring(0, 19), 'yyyy-MM-dd HH:mm:ss', $script:Inv)
}

function Format-LiveLine($r) {
    $f = { param($v, $fmt = '0') if ($null -eq $v) { '-' } else { ([double] $v).ToString($fmt, $script:Inv) } }
    ('{0:HH:mm:ss}  cpu {1}%  mem {2}%  commit {3}%  disk {4}% ({5}/{6} ms)  net {7}/{8} KB/s  ping {9} ms  gpu {10}%  top {11}' +
     '  stall {12} ms  app {13} ms  taskbar {14} ms  gap {15}s  took {16} ms') -f `
        $r['Timestamp'], (& $f $r['CpuPct']), (& $f $r['MemUsedPct']), (& $f $r['CommitPct']), (& $f $r['DiskBusyPct']),
        (& $f $r['DiskReadMs'] '0.0'), (& $f $r['DiskWriteMs'] '0.0'), (& $f $r['NetRecvKBps']), (& $f $r['NetSentKBps']),
        (& $f $r['Ping1Ms']), (& $f $r['GpuPct']), $r['TopCpuProcess'],
        (& $f $r['MaxStallMs']), (& $f $r['ForegroundResponseMs']), (& $f $r['ShellResponseMs']),
        (& $f $r['GapSec'] '0.0'), (& $f $r['CollectMs'])
}

$bound = @{} + $PSBoundParameters
if ($Install -and $Uninstall) { throw 'Use -Install or -Uninstall, not both.' }
if ($Mark -and ($Install -or $Uninstall -or $InventoryOnly)) { throw 'Use -Mark on its own (with -Note, -Popup and -LogDirectory if needed).' }
if (($Note -or $Popup) -and -not $Mark) { throw '-Note and -Popup go with -Mark.' }
if ($FreezeShortcut -and -not $Install) { throw '-FreezeShortcut goes with -Install.' }
if ($Uninstall) { Uninstall-Monitor; return }
if ($Install) { Install-Monitor $bound; return }

# Through powershell.exe -File a list arrives as one comma-separated string.
$PingTarget = @($PingTarget | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -First 2)
if ($TopProcessCount -eq 0 -and $Exclude -notcontains 'Processes') { $Exclude += 'Processes' }
# Waiting for pings mustn't take most of the interval, or every sample would run late.
if ($PingTimeoutMs -gt $IntervalSeconds * 500) { $PingTimeoutMs = [int] [math]::Max(100, $IntervalSeconds * 500) }

$LogDirectory = ConvertTo-FullPath $LogDirectory
$createdLogDirectory = -not (Test-Path -LiteralPath $LogDirectory)
New-Item -ItemType Directory -Force -Path $LogDirectory | Out-Null
# Per account, so users sharing a PC don't collide on one file.
$script:DiagLog = Join-Path $LogDirectory ('monitor_{0}_{1}.log' -f $script:PC, ([Security.Principal.WindowsIdentity]::GetCurrent().Name -split '\\')[-1])
if ($createdLogDirectory -or $LogDirectory -eq (ConvertTo-FullPath (Join-Path $env:ProgramData 'DesktopMonitor\Logs'))) { Grant-LogFolderAccess $LogDirectory }
try { $script:Self.PriorityClass = [Diagnostics.ProcessPriorityClass] $Priority } catch { }

if ($InventoryOnly) { Write-Inventory $bound; return }
if ($Mark) { Add-FreezeMark; return }

# One monitor per log folder, so a manual run can't collide with the scheduled task.
$owned = $false
try {
    $script:Mutex = New-Object Threading.Mutex($false, (Get-MonitorMutexName))
    try { $owned = $script:Mutex.WaitOne(0) }
    catch { $owned = ($_.Exception.GetBaseException() -is [Threading.AbandonedMutexException]) }
} catch { $owned = $false }
if (-not $owned) { Write-MonitorLog "Another DesktopMonitor is already logging to $LogDirectory, so this one is exiting." 'WARN'; return }

if (-not ('DesktopMonitor.Native' -as [type])) {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -TypeDefinition $script:NativeSource -Language CSharp -ReferencedAssemblies ([IO.Compression.ZipArchive].Assembly.Location)
}

Write-MonitorLog "DesktopMonitor $($script:Version) starting: every $IntervalSeconds s, logging to $LogDirectory"
$script:EventRows = New-Object 'System.Collections.Generic.List[object]'
$script:ProcRows  = New-Object 'System.Collections.Generic.List[object]'

$previousSample = $null
try { $previousSample = Get-LastSampleTime } catch { Write-MonitorLog "Could not read the last sample time: $($_.Exception.Message)" 'WARN' }
# Via UTC so the boot time shows in the clock time of the day it happened (summer time).
$bootTime = [datetime]::UtcNow.AddMilliseconds(-1.0 * [DesktopMonitor.Native]::GetTickCount64()).ToLocalTime()
# On the first start, read -EventLookbackHours of the event log. Later starts carry on from
# the last sample, so nothing logged while the monitor was stopped is missed, but go back no
# further than the logs are kept.
$script:EventSince = [datetime]::UtcNow.AddHours(-$EventLookbackHours)
if ($previousSample) {
    $script:EventSince = $previousSample.ToUniversalTime().AddMinutes(-2)
    $oldest = [datetime]::UtcNow.AddDays(-$RetentionDays)
    if ($script:EventSince -lt $oldest) { $script:EventSince = $oldest }
}

$startMsg = 'Monitor started (version {0}, every {1} s, as {2}). Windows started at {3}.' -f $script:Version, $IntervalSeconds,
    [Security.Principal.WindowsIdentity]::GetCurrent().Name, $bootTime.ToString('yyyy-MM-dd HH:mm:ss', $script:Inv)
if ($previousSample) {
    $gap = '{0:0.#} minutes' -f ((Get-Date) - $previousSample).TotalMinutes
    $prevText = $previousSample.ToString('yyyy-MM-dd HH:mm:ss', $script:Inv)
    if ($bootTime -gt $previousSample) {
        $startMsg += " The PC restarted after the last sample ($prevText, $gap ago). If nobody restarted it on purpose, look for BlueScreen, UnexpectedReboot or UnexpectedShutdown events around then."
    } else {
        $between = @()
        try { $between = @(Get-ShutdownsAndSleeps $previousSample) } catch { }
        $startMsg += " The monitor was not running for $gap (last sample $prevText)."
        if (@($between | Where-Object Shutdown).Count) {
            $startMsg += ' Windows was shut down in between. Fast Startup is on, so that was not a full restart: Windows resumed where it was, like waking from hibernation.'
        } elseif ($between.Count) {
            $startMsg += ' The PC was asleep or in standby for some of that time.'
        } else {
            $startMsg += ' Windows kept running all that time.'
        }
    }
} else { $startMsg += ' No earlier samples in this folder.' }
Add-MonitorEvent 'MonitorStart' $startMsg
Write-MonitorLog $startMsg
# After Get-LastSampleTime, which reads the newest unzipped metrics file.
Start-LogHousekeeping

# The inventory loads WMI and storage modules (~30 MB). Running it in a separate,
# short-lived process keeps that memory out of the long-running monitor.
try {
    $invCommand = "& '{0}' -InventoryOnly -LogDirectory '{1}' {2}" -f $PSCommandPath.Replace("'", "''"), $LogDirectory.Replace("'", "''"),
        (ConvertTo-ArgumentText $bound @('Live', 'InventoryOnly', 'LogDirectory', 'Install', 'Uninstall', 'FreezeShortcut', 'Mark', 'Note', 'Popup'))
    Start-Process -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -WindowStyle Hidden `
        -ArgumentList ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -Command "{0}"' -f $invCommand)
} catch { Write-MonitorLog "Could not start the inventory: $($_.Exception.Message)" 'WARN' }

Initialize-Counters
Initialize-CategoryReaders
if (Test-Group 'Security') {
    $names = $script:AgentNames
    $extra = @($AgentName | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($extra.Count) { $names += '|^(' + (($extra | ForEach-Object { [regex]::Escape($_).Replace('\*', '.*').Replace('\?', '.') }) -join '|') + ')$' }
    [DesktopMonitor.Native]::SetAgentRules($script:AgentCompanies, $names, $script:AgentNotAgents)
}
if (Test-Group 'Processes') { [void] [DesktopMonitor.Native]::SampleProcesses() }
[GC]::Collect()

$metricsTable = New-LogTable 'metrics' $script:MetricColumns
$procTable    = New-LogTable 'processes' $script:ProcessColumns
$eventTable   = New-LogTable 'events' $script:EventColumns

# Group, function, run every N samples. Responsiveness comes first so its period ends at the
# row's Timestamp. Processes come before Session so that hung windows can be named from the
# same process snapshot.
$collectors = @(
    @('Responsiveness', 'Get-ResponsivenessMetrics', 1), @('Security', 'Get-SecurityMetrics', 1),
    @('Cpu', 'Get-CpuMetrics', 1), @('Memory', 'Get-MemoryMetrics', 1), @('Disk', 'Get-DiskMetrics', 1),
    @('Network', 'Get-NetworkMetrics', 1), @('Wifi', 'Get-WifiMetrics', 1), @('Ping', 'Get-PingMetrics', 1),
    @('Connectivity', 'Get-ConnectivityMetrics', 1),
    @('Gpu', 'Get-GpuMetrics', 1), @('Thermal', 'Get-ThermalMetrics', 1), @('Power', 'Get-PowerMetrics', 1),
    @('Processes', 'Get-ProcessMetrics', 1), @('Session', 'Get-SessionMetrics', 1),
    @('Events', 'Get-EventMetrics', 1)
)
$deadline   = if ($DurationMinutes -gt 0) { (Get-Date).AddMinutes($DurationMinutes) } else { [datetime]::MaxValue }
$intervalMs = $IntervalSeconds * 1000.0
$nextDue    = $script:Clock.Elapsed.TotalMilliseconds
$prevClock  = $null
$selfPrev   = $null
$script:WaitMs = 0.0
# The process runs at below-normal priority, but the sampling thread runs level with ordinary
# programs. Otherwise a busy CPU holds it up for many seconds, which looks just like a
# freeze in the logs. It only needs ~50 ms a sample.
if ($Priority -ne 'Normal') { try { [Threading.Thread]::CurrentThread.Priority = [Threading.ThreadPriority]::Highest } catch { } }
# Housekeeping runs again after midnight (to zip the day just finished) and every hour (to
# keep within -MaxTotalSizeMB, and to retry files that were open in another program).
$cleanupDay = (Get-Date).Date
$cleanupDue = (Get-Date).AddHours(1)
$stopReason = 'Monitor stopped (Ctrl+C or the window was closed).'
if (Test-Group 'Responsiveness') { [DesktopMonitor.Responsiveness]::Start() }
if (Test-Group 'Connectivity') {
    $domain = $null
    try { $domain = [DesktopMonitor.Connectivity]::JoinedDomain() } catch { }
    if (-not $DnsName) {
        # On a domain PC, the records used to find a domain controller (sign-in, Group Policy
        # and file shares depend on them); the domain's DNS name is the PC's primary DNS suffix.
        $DnsName = 'www.msftconnecttest.com'
        $tcpip = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters' -ErrorAction SilentlyContinue
        $suffix = foreach ($n in 'NV Domain', 'Domain') { if ($tcpip -and $tcpip.PSObject.Properties[$n] -and $tcpip.$n) { $tcpip.$n; break } }
        if ($domain -and $suffix) { $DnsName = "_ldap._tcp.dc._msdcs.$suffix" }
    }
    $servers = @($ServerTarget | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -First 2)
    $http = if ($HttpTarget -and $HttpTarget -ne 'none') { $HttpTarget } else { '' }
    $expect = if ($http -match 'msftconnecttest\.com/connecttest\.txt$') { 'Microsoft Connect Test' } else { '' }
    [DesktopMonitor.Connectivity]::Start($DnsName, [bool] $domain, [string[]] $servers, $http, $expect, [int] ($IntervalSeconds * 1000))
    Write-MonitorLog ("Connectivity checks: DNS name $DnsName" + $(if ($domain) { ", domain $domain" } else { ', not on a domain' }) +
        $(if ($servers.Count) { ", servers $($servers -join ', ')" }) + $(if ($http) { ", web $http" }) + '.')
}
if (Test-Group 'Security') { [DesktopMonitor.Connectivity]::StartFileCheck([int] ($IntervalSeconds * 1000)) }

try {
    while ($true) {
        $script:SampleNo++
        $script:WaitMs = 0.0
        $now = Get-Date
        $t0 = $script:Clock.Elapsed.TotalMilliseconds
        $row = @{ Timestamp = $now }
        # From the stopwatch, not the clock: a summer-time change or a clock correction would
        # otherwise look like an hour-long freeze, or a negative gap.
        if ($null -ne $prevClock) { $row.GapSec = ($t0 - $prevClock) / 1000.0 }
        $prevClock = $t0
        $row.UptimeHours = [DesktopMonitor.Native]::GetTickCount64() / 3600000.0

        $slow = New-Object 'System.Collections.Generic.List[string]'
        foreach ($c in $collectors) {
            if (-not (Test-Group $c[0]) -or -not (Test-Every $c[2])) { continue }
            $tc = $script:Clock.Elapsed.TotalMilliseconds
            try { & $c[1] $row }
            catch { Write-Throttled "collect:$($c[0])" "$($c[0]) metrics failed, columns left blank: $($_.Exception.Message)" }
            $took = $script:Clock.Elapsed.TotalMilliseconds - $tc
            if ($took -ge 500) { $slow.Add(('{0} {1:0} ms' -f $c[0], $took)) }
        }
        try { Read-FreezeMarks $row } catch { Write-Throttled 'marks' "Checking for freeze reports failed: $($_.Exception.Message)" }

        # The monitor's own cost, as a share of the whole machine.
        $script:Self.Refresh()
        $tNow = $script:Clock.Elapsed.TotalMilliseconds
        $cpuMs = $script:Self.TotalProcessorTime.TotalMilliseconds
        if ($selfPrev -and $tNow -gt $selfPrev[0]) { $row.MonitorCpuPct = 100.0 * ($cpuMs - $selfPrev[1]) / ($tNow - $selfPrev[0]) / $script:Cores }
        $selfPrev = @($tNow, $cpuMs)
        $row.MonitorMemMB = $script:Self.PrivateMemorySize64 / 1MB
        # Not counting time spent waiting for pings or the start-up event catch-up.
        $row.CollectMs = [math]::Max(0, $tNow - $t0 - $script:WaitMs)
        # Say which part of a slow sample was slow, so the monitor's own delays can be told
        # apart from the PC's.
        if ($row.CollectMs -ge 2000 -and $slow.Count) { Write-Throttled 'slowSample' "Sample took $([math]::Round($row.CollectMs)) ms: $($slow -join ', ')." }

        Write-TableRows $metricsTable @($row)
        if ($script:ProcRows.Count) { Write-TableRows $procTable $script:ProcRows.ToArray(); $script:ProcRows.Clear() }
        if ($script:EventRows.Count) { Write-TableRows $eventTable $script:EventRows.ToArray(); $script:EventRows.Clear() }
        if ($Live) { try { Write-Host (Format-LiveLine $row) } catch { Write-Throttled 'live' "Console summary failed: $($_.Exception.Message)" } }

        Write-HousekeepingMessages
        if ((Get-Date).Date -ne $cleanupDay -or (Get-Date) -ge $cleanupDue) {
            $cleanupDay = (Get-Date).Date; $cleanupDue = (Get-Date).AddHours(1)
            Start-LogHousekeeping
        }
        if ((Get-Date) -ge $deadline) { $stopReason = "Monitor stopped after $DurationMinutes minutes (-DurationMinutes)."; break }

        # Sleep until the next slot. If the PC stalled past it, skip ahead; GapSec shows the stall.
        $nextDue += $intervalMs
        $wait = $nextDue - $script:Clock.Elapsed.TotalMilliseconds
        if ($wait -lt 0) {
            $nextDue += [math]::Ceiling(-$wait / $intervalMs) * $intervalMs
            $wait = $nextDue - $script:Clock.Elapsed.TotalMilliseconds
        }
        Start-Sleep -Milliseconds ([int] [math]::Max(0, $wait))
    }
} catch {
    $stopReason = "Monitor stopped by an error: $($_.Exception.Message)"
    Write-MonitorLog $stopReason 'ERROR'
    throw   # non-zero exit, so the scheduled task restarts it
} finally {
    Add-MonitorEvent 'MonitorStop' $stopReason
    Write-TableRows $eventTable $script:EventRows.ToArray()
    # Let a zip that is under way finish (if it doesn't, the next start does it again).
    try { [void] [DesktopMonitor.Housekeeping]::Wait(10000); Write-HousekeepingMessages } catch { }
    Write-MonitorLog $stopReason
    if ($script:Mutex) { try { $script:Mutex.ReleaseMutex() } catch { }; $script:Mutex.Dispose() }
}
