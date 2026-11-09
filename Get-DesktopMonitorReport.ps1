#Requires -Version 5.1

<#
.SYNOPSIS
    Summarises DesktopMonitor logs: problem periods, peaks, slow leaks and notable events.

.DESCRIPTION
    Reads the metrics and events files written by DesktopMonitor.ps1 and prints, for each PC
    found in the folder:

      * Problem periods: freezes and stalls, apps or the taskbar not responding, hung apps,
        a driver hogging a CPU core, CPU / GPU / memory / disk saturation, heavy paging, CPU
        throttling, network loss, port exhaustion, low disk space and 32-bit programs
        running out of address space, each with when it started, how long it lasted, its
        worst value and the likely culprit. Gaps that overlap sleep or Modern Standby are
        left out.
      * Reported freezes: the samples around each time a user ran -Mark (or the 'Report a
        freeze' shortcut).
      * Notable event log entries: blue screens, unexpected restarts, app crashes and hangs,
        GPU driver resets, disk and hardware errors, and when the monitor itself stopped/started.
      * Peaks: average, 95th percentile and maximum of the key metrics.
      * Slow growth that suggests a leak (handles, kernel memory pools, commit, ports), the
        biggest kernel pool tags, and the processes that grew the most.
      * The 32-bit processes with the least free address space.

    Point it at the Logs folder on the PC, or a copy of it. Days the monitor has zipped
    (DesktopMonitor_<PC>_<date>.zip) are read straight from the zip.

.PARAMETER LogDirectory
    Folder with the log files. Default C:\ProgramData\DesktopMonitor\Logs.

.PARAMETER From
    Start of the period to report on. Default: 24 hours before -To.

.PARAMETER To
    End of the period. Default: the newest sample in the files.

.PARAMETER ComputerName
    Only report on this PC when the folder holds logs from several. Default: all of them.

.PARAMETER IncidentCsv
    Also save the problem periods to this CSV file.

.EXAMPLE
    .\Get-DesktopMonitorReport.ps1
    Reports on the last 24 hours of this PC's logs.

.EXAMPLE
    .\Get-DesktopMonitorReport.ps1 -LogDirectory '\\PC042\c$\ProgramData\DesktopMonitor\Logs' -From '2026-11-09 09:00' -To '2026-11-09 12:00'
#>
[CmdletBinding()]
param(
    [string] $LogDirectory = (Join-Path $env:ProgramData 'DesktopMonitor\Logs'),
    [datetime] $From,
    [datetime] $To,
    [string] $ComputerName = '*',
    [string] $IncidentCsv
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$inv = [Globalization.CultureInfo]::InvariantCulture

# Numeric columns used below. Blank values become NaN, which fails every comparison.
$numeric = @(
    'GapSec', 'CollectMs', 'CpuPct', 'CpuMaxCorePct', 'CpuMaxCoreDpcPct', 'CpuPerfPct', 'CpuQueue', 'Handles',
    'MemUsedPct', 'MemModifiedMB', 'CommitPct', 'CommitGB',
    'PageInPerSec', 'PoolNonPagedMB', 'PoolPagedMB', 'TopNonPagedPoolMB', 'TopPagedPoolMB',
    'DiskBusyPct', 'DiskReadMs', 'DiskWriteMs', 'DiskIops',
    'SysDriveFreePct', 'NetRecvKBps', 'NetSentKBps', 'TcpRetransPerSec', 'Ping1Ms', 'Ping2Ms', 'PingFailed',
    'GpuPct', 'TopGpuPct', 'ThermalMaxC', 'ThermalLimitPct', 'HungAppCount', 'MonitorCpuPct', 'MonitorMemMB',
    'GdiObjects', 'UserObjects', 'TopHandles', 'TopGdiObjects', 'TopUserObjects', 'TcpPortsInUse', 'TcpTimeWait', 'TopTcpConnections',
    'UdpEndpoints', 'TopUdpEndpoints', 'TopHardFaultsPerSec',
    'MaxStallMs', 'StallCount', 'ForegroundResponseMs', 'ShellResponseMs', 'LargestFreeBlockMB', 'FreeAddressSpaceMB',
    'DnsLookupMs', 'DnsFailed', 'DcConnectMs', 'DcFailed', 'Server1Ms', 'Server2Ms', 'ServerFailed', 'HttpMs', 'HttpFailed',
    'NetChanges', 'AgentCpuPct', 'AgentIoKBps', 'AgentMemMB', 'TopAgentCpuPct', 'FileCheckMs', 'FileCheckFailed', 'WifiSignalPct'
)
$text = @('TopCpuProcess', 'TopMemProcess', 'TopIoProcess', 'TopHardFaultProcess', 'TopHandleProcess', 'TopGdiProcess',
    'TopUserObjProcess', 'TopTcpProcess', 'TopUdpProcess', 'TopGpuProcess', 'TopNonPagedPoolTag', 'TopPagedPoolTag',
    'HungApps', 'ForegroundApp', 'NetAdapters', 'ForegroundResponseApp', 'Top32BitProcess', 'EventTags',
    'DomainController', 'MappedDrivesDown', 'NetworkCategory', 'WindowsConnectivity', 'Vpn', 'DnsServers', 'ConnectivityErrors',
    'TopAgentProcess')

# A column's value, or '' when the file predates that column.
function Get-Field($Row, [string] $Name) {
    if ($Row.PSObject.Properties[$Name]) { return $Row.$Name }
    return ''
}

function ConvertTo-Number([string] $s) {
    $d = 0.0
    if ([double]::TryParse($s, [Globalization.NumberStyles]::Float, $inv, [ref] $d)) { return $d }
    return [double]::NaN
}

# The time, or $null for a damaged line (a power cut can leave half a line, or zero bytes,
# at the end of a file).
function ConvertTo-Time([string] $s) {
    $t = [datetime]::MinValue
    if ([datetime]::TryParseExact($s, 'yyyy-MM-dd HH:mm:ss', $inv, [Globalization.DateTimeStyles]::None, [ref] $t)) { return $t }
    return $null
}
$script:BadRows = 0
$script:BadFiles = New-Object 'System.Collections.Generic.List[string]'

Add-Type -AssemblyName System.IO.Compression

# Log files of one kind for one PC: plain files (also *.zipping ones, which the monitor was
# about to zip when it stopped), and the entries inside the daily zips. Each comes back as
# @{ Day; Name; Zip; Entry } (Zip and Entry are blank for a plain file), oldest day first.
function Get-LogSources([string] $Kind, [string] $Pc) {
    $pcRe = [regex]::Escape($Pc)
    $found = foreach ($f in (Get-ChildItem -LiteralPath $LogDirectory -File)) {
        if ($f.Name -match "^$Kind`_${pcRe}_(\d{8})(_\d+)?\.(csv|tsv)(\.zipping)?$") {
            [pscustomobject] @{ Day = $Matches[1]; Name = $f.FullName; Zip = ''; Entry = '' }
        } elseif ($f.Name -match "^DesktopMonitor_${pcRe}_(\d{8})(_\d+)?\.zip$") {
            $day = $Matches[1]
            # A damaged or half-copied zip is skipped (and named at the end), not fatal.
            try {
                $zip = New-Object IO.Compression.ZipArchive(([IO.File]::OpenRead($f.FullName)), [IO.Compression.ZipArchiveMode]::Read)
                try {
                    foreach ($e in $zip.Entries) {
                        if ($e.Name -match "^$Kind`_${pcRe}_\d{8}(_\d+)?\.(csv|tsv)$") {
                            [pscustomobject] @{ Day = $day; Name = $e.Name; Zip = $f.FullName; Entry = $e.FullName }
                        }
                    }
                } finally { $zip.Dispose() }
            } catch { if (-not $script:BadFiles.Contains($f.Name)) { $script:BadFiles.Add($f.Name) } }
        }
    }
    @($found | Sort-Object Day, Name)
}

function Read-LogSource($Source) {
    $delimiter = if ($Source.Name -match '\.tsv(\.zipping)?$') { "`t" } else { ',' }
    try {
        if (-not $Source.Zip) { return Import-Csv -LiteralPath $Source.Name -Delimiter $delimiter }
        $zip = New-Object IO.Compression.ZipArchive(([IO.File]::OpenRead($Source.Zip)), [IO.Compression.ZipArchiveMode]::Read)
        try {
            $reader = New-Object IO.StreamReader($zip.GetEntry($Source.Entry).Open(), [Text.Encoding]::UTF8, $true)
            try { $lines = $reader.ReadToEnd() -split "`r?`n" } finally { $reader.Dispose() }
        } finally { $zip.Dispose() }
        $lines | Where-Object { $_ } | ConvertFrom-Csv -Delimiter $delimiter
    } catch {
        $name = if ($Source.Zip) { Split-Path $Source.Zip -Leaf } else { Split-Path $Source.Name -Leaf }
        if (-not $script:BadFiles.Contains($name)) { $script:BadFiles.Add($name) }
    }
}

# Rows from every file of one kind for one PC whose date is inside the period. Events can
# be written to a later day's file than the day they happened (the monitor catches up on
# the event log when it starts), so for them any later day is read too.
function Import-LogRows([string] $Kind, [string] $Pc, [datetime] $Start, [datetime] $End) {
    foreach ($s in (Get-LogSources $Kind $Pc)) {
        $day = [datetime]::ParseExact($s.Day, 'yyyyMMdd', $inv)
        if ($day -lt $Start.Date -or ($day -gt $End.Date -and $Kind -ne 'events')) { continue }
        foreach ($r in (Read-LogSource $s)) {
            $t = ConvertTo-Time (Get-Field $r 'Timestamp')
            if ($null -eq $t) { $script:BadRows++; continue }
            if ($t -ge $Start -and $t -le $End) { $r | Add-Member -NotePropertyName Time -NotePropertyValue $t -PassThru }
        }
    }
}

function Get-Percentile([double[]] $Values, [double] $P) {
    if (-not $Values.Count) { return [double]::NaN }
    $sorted = $Values | Sort-Object
    $sorted[[int][math]::Min($sorted.Count - 1, [math]::Floor($P * $sorted.Count))]
}

function Format-Number([double] $v, [string] $fmt = '0.#') {
    if ([double]::IsNaN($v)) { return '-' }
    return $v.ToString($fmt, $inv)
}

# The most frequent non-blank value of a column across rows.
function Get-MostCommon($Rows, [string] $Column) {
    $top = $Rows | Where-Object { $_.$Column } | Group-Object $Column | Sort-Object Count -Descending | Select-Object -First 1
    if ($top) { return '{0} ({1}x)' -f $top.Name, $top.Count }
    return ''
}

# Problem rules. A period is reported when the test holds for at least MinSamples samples
# in a row (short CPU or disk spikes are normal). Peak is the column whose worst value is
# shown (Low = smaller is worse); Who is the column naming the likely culprit (the first of
# several that has a value: older logs lack the newer columns), or Culprit is fixed text.
# SkipSleep leaves out periods that overlap sleep or Modern Standby, which look the same as
# a freeze (the monitor's process doesn't run).
$rules = @(
    # The monitor's own sample was late. With the CPU at 100% and no stall seen, that is the
    # monitor waiting for CPU time (reported as CpuSaturated), not a freeze.
    @{ Name = 'Freeze';         MinSamples = 1; Peak = 'GapSec';          Who = 'ForegroundApp'; SkipSleep = $true
       Test = { param($r, $i) ($r.GapSec -gt [math]::Max(2 * $i, $i + 5) -or $r.CollectMs -gt 3000) -and
                              -not ($r.CpuPct -ge 90 -and $r.MaxStallMs -lt 1000) } }
    # Nothing on the PC could run: the monitor's highest-priority thread woke late.
    @{ Name = 'Stall';          MinSamples = 1; Peak = 'MaxStallMs';      Who = 'EventTags'; SkipSleep = $true
       Test = { param($r, $i) $r.MaxStallMs -ge 1000 -or $r.StallCount -ge 3 } }
    @{ Name = 'AppNotResponding'; MinSamples = 1; Peak = 'ForegroundResponseMs'; Who = 'ForegroundResponseApp'; SkipSleep = $true
       Test = { param($r, $i) $r.ForegroundResponseMs -ge 2000 } }
    @{ Name = 'TaskbarNotResponding'; MinSamples = 1; Peak = 'ShellResponseMs'; Who = 'TopCpuProcess'; SkipSleep = $true
       Test = { param($r, $i) $r.ShellResponseMs -ge 1000 } }
    @{ Name = 'HungApp';        MinSamples = 1; Peak = 'HungAppCount';    Who = 'HungApps'
       Test = { param($r, $i) $r.HungAppCount -gt 0 } }
    # One core busy with a driver's interrupts and DPCs: mouse lag, audio crackle, stutter.
    @{ Name = 'DpcStorm';       MinSamples = 1; Peak = 'CpuMaxCoreDpcPct'; Culprit = 'a driver (LatencyMon names it)'
       Test = { param($r, $i) $r.CpuMaxCoreDpcPct -ge 20 } }
    @{ Name = 'CpuSaturated';   MinSamples = 3; Peak = 'CpuPct';          Who = 'TopCpuProcess'
       Test = { param($r, $i) $r.CpuPct -ge 90 } }
    @{ Name = 'GpuSaturated';   MinSamples = 3; Peak = 'GpuPct';          Who = 'TopGpuProcess'
       Test = { param($r, $i) $r.GpuPct -ge 95 } }
    @{ Name = 'MemoryPressure'; MinSamples = 1; Peak = 'CommitPct';       Who = 'TopMemProcess'
       Test = { param($r, $i) $r.CommitPct -ge 90 -or $r.MemUsedPct -ge 95 } }
    @{ Name = 'HeavyPaging';    MinSamples = 2; Peak = 'PageInPerSec';    Who = 'TopHardFaultProcess', 'TopIoProcess'
       Test = { param($r, $i) $r.PageInPerSec -ge 1000 } }
    @{ Name = 'DiskSlow';       MinSamples = 3; Peak = 'DiskBusyPct';     Who = 'TopIoProcess'
       Test = { param($r, $i) $r.DiskBusyPct -ge 90 -or (($r.DiskReadMs -ge 50 -or $r.DiskWriteMs -ge 50) -and $r.DiskIops -ge 5) } }
    @{ Name = 'CpuThrottled';   MinSamples = 3; Peak = 'CpuPerfPct';      Who = 'TopCpuProcess'; Low = $true
       Test = { param($r, $i) ($r.CpuPerfPct -lt 60 -and $r.CpuPct -ge 20) -or $r.ThermalLimitPct -lt 100 } }
    @{ Name = 'NetworkLoss';    MinSamples = 1; Peak = 'PingFailed';      Who = 'NetAdapters'
       Test = { param($r, $i) $r.PingFailed -gt 0 } }
    # Windows itself saw the network go (the Network List Manager).
    @{ Name = 'NetworkDown';    MinSamples = 1; Peak = '';                Who = 'NetAdapters'
       Test = { param($r, $i) $r.WindowsConnectivity -eq 'Disconnected' -or $r.WindowsConnectivity -eq 'NoTraffic' } }
    # Name lookups slow or failing: almost everything on a network waits on DNS first.
    @{ Name = 'DnsProblem';     MinSamples = 1; Peak = 'DnsLookupMs';     Who = 'ConnectivityErrors', 'DnsServers'
       Test = { param($r, $i) $r.DnsLookupMs -ge 1000 -or $r.DnsFailed -gt 0 } }
    # Sign-in, Group Policy, file shares and Kerberos tickets need a domain controller.
    @{ Name = 'DomainControllerProblem'; MinSamples = 1; Peak = 'DcConnectMs'; Who = 'ConnectivityErrors', 'DomainController'
       Test = { param($r, $i) $r.DcConnectMs -ge 1000 -or $r.DcFailed -gt 0 } }
    # The domain controller answers, but Windows hasn't recognised the network as the domain
    # network, so the Public or Private firewall profile applies, not the Domain one.
    @{ Name = 'DomainNetworkNotRecognised'; MinSamples = 3; Peak = 'DcConnectMs'; Who = 'NetworkCategory'
       Test = { param($r, $i) $r.DcConnectMs -lt 1000 -and $r.NetworkCategory -and $r.NetworkCategory -notmatch 'Domain' } }
    @{ Name = 'ServerUnreachable'; MinSamples = 1; Peak = 'ServerFailed'; Who = 'ConnectivityErrors'
       Test = { param($r, $i) $r.ServerFailed -gt 0 -or $r.Server1Ms -ge 1000 -or $r.Server2Ms -ge 1000 } }
    # The web through the user's proxy: a slow PAC/WPAD script, proxy sign-in or TLS
    # inspection shows up here.
    @{ Name = 'WebAccessProblem'; MinSamples = 1; Peak = 'HttpMs';        Who = 'ConnectivityErrors'
       Test = { param($r, $i) $r.HttpMs -ge 5000 -or $r.HttpFailed -gt 0 } }
    # File Explorer and Office can hang for 30 s or more on a drive whose server is gone.
    @{ Name = 'MappedDriveOffline'; MinSamples = 1; Peak = '';            Who = 'MappedDrivesDown'
       Test = { param($r, $i) [bool] $r.MappedDrivesDown } }
    # Security and monitoring agents using a lot of CPU (a scan, an update, a stuck agent).
    @{ Name = 'SecurityAgentBusy'; MinSamples = 3; Peak = 'AgentCpuPct';  Who = 'TopAgentProcess'
       Test = { param($r, $i) $r.AgentCpuPct -ge 25 } }
    # Writing and reading a 64 KB file normally takes a few ms. When antivirus, EDR or DLP
    # filters are struggling, every program that opens files waits like this.
    @{ Name = 'FileAccessSlow'; MinSamples = 1; Peak = 'FileCheckMs';     Who = 'TopAgentProcess', 'TopIoProcess'
       Test = { param($r, $i) $r.FileCheckMs -ge 500 -or $r.FileCheckFailed -gt 0 } }
    @{ Name = 'LowDiskSpace';   MinSamples = 1; Peak = 'SysDriveFreePct'; Who = 'TopIoProcess'; Low = $true
       Test = { param($r, $i) $r.SysDriveFreePct -lt 10 } }
    # 10,000 GDI or USER objects per process (and 65,536 GDI per session) are the default
    # limits; near them, windows stop drawing and apps hang or crash.
    @{ Name = 'GdiNearLimit';   MinSamples = 1; Peak = 'TopGdiObjects';   Who = 'TopGdiProcess'
       Test = { param($r, $i) $r.TopGdiObjects -ge 8000 -or $r.GdiObjects -ge 50000 } }
    @{ Name = 'UserObjNearLimit'; MinSamples = 1; Peak = 'TopUserObjects'; Who = 'TopUserObjProcess'
       Test = { param($r, $i) $r.TopUserObjects -ge 8000 } }
    @{ Name = 'HighHandleCount'; MinSamples = 1; Peak = 'TopHandles';     Who = 'TopHandleProcess'
       Test = { param($r, $i) $r.TopHandles -ge 100000 } }
    # Windows has 16,384 ephemeral ports by default; past that, new connections fail.
    @{ Name = 'TcpPortPressure'; MinSamples = 1; Peak = 'TcpPortsInUse';  Who = 'TopTcpProcess'
       Test = { param($r, $i) $r.TcpPortsInUse -ge 10000 } }
    @{ Name = 'UdpPortPressure'; MinSamples = 1; Peak = 'UdpEndpoints';   Who = 'TopUdpProcess'
       Test = { param($r, $i) $r.UdpEndpoints -ge 10000 } }
    # A 32-bit program whose address space is nearly used up or too fragmented: its next
    # large allocation may fail ("out of memory", blank windows, a crash).
    @{ Name = 'AddressSpaceLow'; MinSamples = 1; Peak = 'LargestFreeBlockMB'; Who = 'Top32BitProcess'; Low = $true
       Test = { param($r, $i) $r.LargestFreeBlockMB -lt 200 -or $r.FreeAddressSpaceMB -lt 500 } }
)

# Periods left out because they overlap sleep or standby (counted for the report).
$script:SkippedForSleep = 0

function Find-Incidents($Rows, [double] $Interval, [string] $Pc, [datetime[]] $SleepTimes) {
    foreach ($rule in $rules) {
        $run = New-Object 'System.Collections.Generic.List[object]'
        # A trailing $null closes the last run. A blank GapSec marks the first sample after the
        # monitor (re)started, so a period never spans time when it wasn't running.
        foreach ($r in @($Rows) + @($null)) {
            $restart = ($null -ne $r -and [double]::IsNaN($r.GapSec))
            $match = ($null -ne $r -and (& $rule.Test $r $Interval))
            if ($match -and -not ($restart -and $run.Count)) { $run.Add($r); continue }
            if ($run.Count -ge $rule.MinSamples) {
                $values = @(if ($rule.Peak) { $run | ForEach-Object { $_.($rule.Peak) } | Where-Object { -not [double]::IsNaN($_) } })
                $peak = if (-not $values.Count) { [double]::NaN }
                        elseif ($rule.ContainsKey('Low')) { ($values | Measure-Object -Minimum).Minimum }
                        else { ($values | Measure-Object -Maximum).Maximum }
                $end = $run[$run.Count - 1].Time
                # The first sample covers the time since the one before (the stall, for a freeze).
                $firstGap = if ([double]::IsNaN($run[0].GapSec)) { $Interval } else { $run[0].GapSec }
                $since = $run[0].Time.AddSeconds(-$firstGap - 60)
                $asleep = $rule.ContainsKey('SkipSleep') -and
                    @($SleepTimes | Where-Object { $_ -ge $since -and $_ -le $end.AddSeconds(60) }).Count -gt 0
                if ($asleep) {
                    $script:SkippedForSleep++
                } else {
                    $culprit = ''
                    if ($rule.ContainsKey('Culprit')) { $culprit = $rule.Culprit }
                    else { foreach ($w in @($rule.Who)) { $culprit = Get-MostCommon $run $w; if ($culprit) { break } } }
                    [pscustomobject] @{
                        Computer = $Pc; Problem = $rule.Name
                        Start    = $run[0].Time.ToString('yyyy-MM-dd HH:mm:ss', $inv)
                        End      = $end.ToString('yyyy-MM-dd HH:mm:ss', $inv)
                        Minutes  = [math]::Round((($end - $run[0].Time).TotalSeconds + $firstGap) / 60, 1)
                        Worst    = $(if ($rule.Peak) { '{0} {1}' -f $rule.Peak, (Format-Number $peak) } else { '' })
                        Culprit  = $culprit
                    }
                }
            }
            $run.Clear()
            # The row that broke the run because of a restart may itself start a new one.
            if ($match) { $run.Add($r) }
        }
    }
}

# ---------------------------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $LogDirectory)) { throw "Log folder not found: $LogDirectory" }
$pcs = @(Get-ChildItem -LiteralPath $LogDirectory -File | Where-Object { $_.Name -like "metrics_$ComputerName`_*" -or $_.Name -like "DesktopMonitor_$ComputerName`_*.zip" } |
    ForEach-Object { if ($_.Name -match '^(?:metrics_(.+)_\d{8}(_\d+)?\.(csv|tsv)(\.zipping)?|DesktopMonitor_(.+)_\d{8}(_\d+)?\.zip)$') { "$($Matches[1])$($Matches[5])" } } | Sort-Object -Unique)
if (-not $pcs.Count) { throw "No metrics files in $LogDirectory" }

$allIncidents = New-Object 'System.Collections.Generic.List[object]'
foreach ($pc in $pcs) {
    # Default period: the 24 hours up to this PC's newest sample.
    $end = $To; $start = $From
    if (-not $PSBoundParameters.ContainsKey('To')) { $end = [datetime]::MaxValue }
    if (-not $PSBoundParameters.ContainsKey('From')) { $start = [datetime]::MinValue }
    if (-not $PSBoundParameters.ContainsKey('To') -or -not $PSBoundParameters.ContainsKey('From')) {
        $sources = @(Get-LogSources 'metrics' $pc)
        $newestDay = if ($sources.Count) { $sources[-1].Day } else { '' }
        $newest = $null
        # The last line of a plain file is quick to read; a zipped day is read whole.
        foreach ($s in @($sources | Where-Object { $_.Day -eq $newestDay })) {
            $t = $null
            if (-not $s.Zip) {
                $last = Get-Content -LiteralPath $s.Name -Tail 5 | Where-Object { $_ -match '^"?(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d)' } | Select-Object -Last 1
                if ($last -and $last -match '^"?(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d)') { $t = ConvertTo-Time $Matches[1] }
            } else {
                $t = Read-LogSource $s | ForEach-Object { ConvertTo-Time (Get-Field $_ 'Timestamp') } | Where-Object { $_ } | Sort-Object | Select-Object -Last 1
            }
            if ($t -and (-not $newest -or $t -gt $newest)) { $newest = $t }
        }
        if (-not $newest) { $newest = Get-Date }
        if (-not $PSBoundParameters.ContainsKey('To')) { $end = $newest }
        if (-not $PSBoundParameters.ContainsKey('From')) { $start = $end.AddHours(-24) }
    }

    $rows = foreach ($r in (Import-LogRows 'metrics' $pc $start $end)) {
        $o = [ordered] @{ Time = $r.Time }
        foreach ($c in $text) { $o[$c] = Get-Field $r $c }
        foreach ($c in $numeric) { $o[$c] = ConvertTo-Number (Get-Field $r $c) }
        [pscustomobject] $o
    }
    # Sorted, because a locked file makes the monitor carry on in name_1.csv.
    $rows = @($rows | Sort-Object Time)
    $events = @(Import-LogRows 'events' $pc $start $end | Sort-Object Time)

    Write-Host ''
    Write-Host ('=' * 100)
    Write-Host ("{0}: {1:yyyy-MM-dd HH:mm} to {2:yyyy-MM-dd HH:mm}, {3} samples" -f $pc, $start, $end, $rows.Count) -ForegroundColor Cyan
    Write-Host ('=' * 100)
    if (-not $rows.Count) { Write-Host 'No samples in this period.'; continue }

    $gaps = @($rows | ForEach-Object { $_.GapSec } | Where-Object { -not [double]::IsNaN($_) })
    $interval = if ($gaps.Count) { Get-Percentile $gaps 0.5 } else { 10 }

    Write-Host "`nPROBLEM PERIODS" -ForegroundColor Yellow
    $sleepTimes = [datetime[]] @($events | Where-Object { $_.Tag -eq 'Sleep' -or $_.Tag -eq 'Wake' } | ForEach-Object { $_.Time })
    $script:SkippedForSleep = 0
    $incidents = @(Find-Incidents $rows $interval $pc $sleepTimes | Sort-Object Start)
    if ($incidents.Count) {
        $incidents | Format-Table Start, Problem, Minutes, Worst, Culprit -AutoSize -Wrap | Out-String -Width 160 | Write-Host
        $incidents | ForEach-Object { $allIncidents.Add($_) }
    } else { Write-Host '  None found.' }
    if ($script:SkippedForSleep -eq 1) {
        Write-Host '  (1 freeze or stall period overlapped sleep or Modern Standby and is left out.)'
    } elseif ($script:SkippedForSleep) {
        Write-Host "  ($($script:SkippedForSleep) freeze or stall periods overlapped sleep or Modern Standby and are left out.)"
    }

    # What the monitor saw just before and after each time a user reported a freeze.
    $marks = @($events | Where-Object { $_.Tag -eq 'UserReportedFreeze' })
    if ($marks.Count) {
        Write-Host "`nREPORTED FREEZES (samples from 2 minutes before to 30 seconds after each report)" -ForegroundColor Yellow
        if ($marks.Count -gt 10) { Write-Host "  The last 10 of $($marks.Count) reports:" }
        foreach ($m in ($marks | Select-Object -Last 10)) {
            Write-Host ("`n  {0:yyyy-MM-dd HH:mm:ss}  {1}" -f $m.Time, $m.Message)
            $near = @($rows | Where-Object { $_.Time -ge $m.Time.AddMinutes(-2) -and $_.Time -le $m.Time.AddSeconds(30) })
            if (-not $near.Count) { Write-Host '    No samples around this time (the monitor was not running).'; continue }
            $near | Format-Table @{ n = 'Time'; e = { $_.Time.ToString('HH:mm:ss') } },
                @{ n = 'StallMs'; e = { Format-Number $_.MaxStallMs '0' } },
                @{ n = 'AppMs'; e = { Format-Number $_.ForegroundResponseMs '0' } }, @{ n = 'App'; e = { $_.ForegroundResponseApp } },
                @{ n = 'TaskbarMs'; e = { Format-Number $_.ShellResponseMs '0' } }, @{ n = 'Cpu%'; e = { Format-Number $_.CpuPct '0' } },
                @{ n = 'DpcCore%'; e = { Format-Number $_.CpuMaxCoreDpcPct '0' } },
                @{ n = 'Commit%'; e = { Format-Number $_.CommitPct '0' } }, @{ n = 'Disk%'; e = { Format-Number $_.DiskBusyPct '0' } },
                @{ n = 'PageIn/s'; e = { Format-Number $_.PageInPerSec '0' } },
                @{ n = 'DnsMs'; e = { Format-Number $_.DnsLookupMs '0' } }, @{ n = 'DcMs'; e = { Format-Number $_.DcConnectMs '0' } },
                @{ n = 'Ping'; e = { Format-Number $_.Ping1Ms '0' } }, @{ n = 'FileMs'; e = { Format-Number $_.FileCheckMs '0' } },
                @{ n = 'Agent%'; e = { Format-Number $_.AgentCpuPct '0' } },
                @{ n = 'TopCpu'; e = { $_.TopCpuProcess } }, @{ n = 'Hung'; e = { $_.HungApps } } -AutoSize |
                Out-String -Width 240 | Write-Host
            # Event log entries around the same time (network drops, DNS timeouts, scans, ...).
            $around = @($events | Where-Object { $_.Time -ge $m.Time.AddMinutes(-5) -and $_.Time -le $m.Time.AddMinutes(1) -and $_.Tag -ne 'UserReportedFreeze' })
            foreach ($ev in $around) { Write-Host ('    event {0:HH:mm:ss}  {1,-22} {2}' -f $ev.Time, $ev.Tag, $(if ($ev.Message.Length -gt 100) { $ev.Message.Substring(0, 100) + '...' } else { $ev.Message })) }
            $errs = @($near | Where-Object { $_.ConnectivityErrors } | ForEach-Object { $_.ConnectivityErrors } | Select-Object -Unique)
            foreach ($e in $errs) { Write-Host "    connectivity: $e" }
        }
    }

    Write-Host "NOTABLE EVENTS" -ForegroundColor Yellow
    $quiet = @('Information', 'Warning', 'Error', 'Boot', 'Shutdown', 'Sleep', 'Wake', 'ErrorReport', 'RestartRequested',
        'NetworkConnected', 'WifiConnected', 'ConnectivityChange', 'DefenderScanFinished', 'DefenderUpdate')
    $notable = @($events | Where-Object { $quiet -notcontains $_.Tag -and -not ($_.Tag -eq 'MonitorStart' -and $_.Message -notmatch 'restarted') -and $_.Tag -ne 'MonitorStop' })
    if ($notable.Count) {
        $notable | Select-Object -Last 40 |
            Format-Table @{ n = 'Time'; e = { $_.Time.ToString('yyyy-MM-dd HH:mm:ss') } }, Tag, Provider, EventId,
                @{ n = 'Message'; e = { if ($_.Message.Length -gt 110) { $_.Message.Substring(0, 110) + '...' } else { $_.Message } } } -AutoSize |
            Out-String -Width 220 | Write-Host
        if ($notable.Count -gt 40) { Write-Host "  ($($notable.Count - 40) earlier entries not shown; see the events file.)" }
    } else { Write-Host '  None.' }
    if ($events.Count) {
        Write-Host ('  All event tags: ' + (($events | Group-Object Tag | Sort-Object Count -Descending | ForEach-Object { "$($_.Name) $($_.Count)" }) -join ', '))
    }

    Write-Host "`nPEAKS" -ForegroundColor Yellow
    $peakColumns = 'MaxStallMs', 'ForegroundResponseMs', 'ShellResponseMs',
        'DnsLookupMs', 'DcConnectMs', 'Server1Ms', 'Server2Ms', 'HttpMs', 'FileCheckMs', 'AgentCpuPct', 'AgentMemMB',
        'CpuPct', 'CpuMaxCorePct', 'CpuMaxCoreDpcPct', 'CpuQueue', 'MemUsedPct', 'CommitPct', 'MemModifiedMB', 'PageInPerSec',
        'TopHardFaultsPerSec', 'DiskBusyPct', 'DiskReadMs', 'DiskWriteMs',
        'NetRecvKBps', 'TcpRetransPerSec', 'TcpPortsInUse', 'UdpEndpoints', 'Ping1Ms', 'Ping2Ms', 'GpuPct', 'ThermalMaxC',
        'GdiObjects', 'TopGdiObjects', 'TopUserObjects', 'TopHandles', 'GapSec', 'MonitorCpuPct', 'MonitorMemMB'
    $peakColumns | ForEach-Object {
        $c = $_
        $valid = @($rows | Where-Object { -not [double]::IsNaN($_.$c) })
        if (-not $valid.Count) { return }
        $max = $valid | Sort-Object $c -Descending | Select-Object -First 1
        [pscustomobject] @{
            Metric = $c
            Average = Format-Number (($valid | Measure-Object $c -Average).Average)
            P95 = Format-Number (Get-Percentile ([double[]] @($valid | ForEach-Object { $_.$c })) 0.95)
            Max = Format-Number $max.$c
            At = $max.Time.ToString('MM-dd HH:mm:ss')
        }
    } | Format-Table -AutoSize | Out-String | Write-Host

    Write-Host "GROWTH (first hour vs last hour; steady growth over days suggests a leak)" -ForegroundColor Yellow
    $hours = ($rows[-1].Time - $rows[0].Time).TotalHours
    if ($hours -lt 2) { Write-Host '  Needs at least 2 hours of samples.' }
    else {
        $firstHour = @($rows | Where-Object { $_.Time -le $rows[0].Time.AddHours(1) })
        $lastHour  = @($rows | Where-Object { $_.Time -ge $rows[-1].Time.AddHours(-1) })
        'Handles', 'GdiObjects', 'UserObjects', 'PoolNonPagedMB', 'PoolPagedMB', 'CommitGB', 'TcpPortsInUse', 'UdpEndpoints' | ForEach-Object {
            $c = $_
            $a = ($firstHour | Where-Object { -not [double]::IsNaN($_.$c) } | Measure-Object $c -Average).Average
            $b = ($lastHour | Where-Object { -not [double]::IsNaN($_.$c) } | Measure-Object $c -Average).Average
            if ($null -eq $a -or $null -eq $b) { return }
            [pscustomobject] @{
                Metric = $c; FirstHour = Format-Number $a; LastHour = Format-Number $b
                Change = '{0}%' -f (Format-Number $(if ($a) { 100 * ($b - $a) / $a } else { [double]::NaN }) '+0;-0;0')
                PerHour = Format-Number (($b - $a) / ($hours - 1))
            }
        } | Format-Table -AutoSize | Out-String | Write-Host
    }

    # The biggest kernel pool tag at the start and end of the period. Pool is kernel memory;
    # a tag of hundreds of MB, or one that keeps growing, is a driver (or a kind of object,
    # such as Toke for security tokens) that is leaking.
    $pool = @($rows | Where-Object { $_.TopPagedPoolTag -or $_.TopNonPagedPoolTag })
    if ($pool.Count) {
        Write-Host "KERNEL POOL (the biggest tag; hundreds of MB, or steady growth, points to a leak)" -ForegroundColor Yellow
        $(foreach ($kind in 'Paged', 'NonPaged') {
            $tagCol = "Top$($kind)PoolTag"; $mbCol = "Top$($kind)PoolMB"
            $have = @($pool | Where-Object { $_.$tagCol })
            if (-not $have.Count) { continue }
            $first = $have[0]; $last = $have[-1]
            $peak = $have | Sort-Object $mbCol -Descending | Select-Object -First 1
            [pscustomobject] @{
                Pool = $kind
                Start = '{0} {1} MB' -f $first.$tagCol, (Format-Number $first.$mbCol '0')
                End = '{0} {1} MB' -f $last.$tagCol, (Format-Number $last.$mbCol '0')
                Peak = '{0} {1} MB at {2:MM-dd HH:mm}' -f $peak.$tagCol, (Format-Number $peak.$mbCol '0'), $peak.Time
            }
        }) | Format-Table -AutoSize | Out-String -Width 160 | Write-Host
    }

    # Which process is leaking: biggest rise per process in the once-a-minute rankings.
    Write-Host "PROCESS GROWTH (biggest rises, from the once-a-minute top lists)" -ForegroundColor Yellow
    $procRows = @(Import-LogRows 'processes' $pc $start $end)
    $ranked = @($procRows | Where-Object { @('Memory', 'Handles', 'Gdi', 'User', 'Tcp', 'Udp') -contains $_.Kind })
    $growth = foreach ($pair in @(@('Memory', 'PrivateMB'), @('Handles', 'Handles'), @('Gdi', 'GdiObjects'), @('User', 'UserObjects'),
                                  @('Tcp', 'TcpConnections'), @('Udp', 'UdpEndpoints'))) {
        $ranked | Where-Object { $_.Kind -eq $pair[0] } | Group-Object PID, Process | ForEach-Object {
            $series = @($_.Group | Sort-Object Time)
            if ($series.Count -lt 2) { return }
            $a = ConvertTo-Number (Get-Field $series[0] $pair[1])
            $b = ConvertTo-Number (Get-Field $series[-1] $pair[1])
            if ([double]::IsNaN($a) -or [double]::IsNaN($b) -or $b -le $a) { return }
            $hoursSpan = ($series[-1].Time - $series[0].Time).TotalHours
            $services = Get-Field $series[-1] 'Services'
            [pscustomobject] @{
                Resource = $pair[1]; PID = $series[0].PID
                Process = $(if ($services) { '{0} ({1})' -f $series[0].Process, $services } else { $series[0].Process })
                From = Format-Number $a; To = Format-Number $b; Rise = $b - $a
                PerHour = Format-Number $(if ($hoursSpan -gt 0) { ($b - $a) / $hoursSpan } else { [double]::NaN })
                Between = '{0:MM-dd HH:mm} - {1:MM-dd HH:mm}' -f $series[0].Time, $series[-1].Time
            }
        } | Sort-Object Rise -Descending | Select-Object -First 3
    }
    $growth = @($growth)
    if ($growth.Count) {
        $growth | Format-Table Resource, Process, PID, From, To, PerHour, Between -AutoSize | Out-String -Width 160 | Write-Host
    } else { Write-Host '  None (needs a few minutes of samples).' }

    # Each security / monitoring agent's load, from the once-a-minute Agent rows.
    $agentRows = @($procRows | Where-Object { $_.Kind -eq 'Agent' })
    if ($agentRows.Count) {
        Write-Host "`nSECURITY AND MONITORING AGENTS (once a minute; the busiest first)" -ForegroundColor Yellow
        $agentRows | Group-Object Process, Agent | ForEach-Object {
            $cpu = @($_.Group | ForEach-Object { ConvertTo-Number $_.CpuPct } | Where-Object { -not [double]::IsNaN($_) })
            $mem = @($_.Group | ForEach-Object { ConvertTo-Number $_.PrivateMB } | Where-Object { -not [double]::IsNaN($_) })
            $io  = @($_.Group | ForEach-Object { ConvertTo-Number $_.IoKBps } | Where-Object { -not [double]::IsNaN($_) })
            [pscustomobject] @{
                Process = $_.Group[0].Process; Product = Get-Field $_.Group[0] 'Agent'
                AvgCpuPct = ($cpu | Measure-Object -Average).Average; MaxCpuPct = Format-Number ($cpu | Measure-Object -Maximum).Maximum
                MaxMemMB = Format-Number ($mem | Measure-Object -Maximum).Maximum '0'; MaxIoKBps = Format-Number ($io | Measure-Object -Maximum).Maximum '0'
            }
        } | Sort-Object AvgCpuPct -Descending | Select-Object -First 15 |
            Format-Table Process, Product, @{ n = 'AvgCpu%'; e = { Format-Number $_.AvgCpuPct '0.##' } }, @{ n = 'MaxCpu%'; e = { $_.MaxCpuPct } },
                @{ n = 'MaxMemMB'; e = { $_.MaxMemMB } }, @{ n = 'MaxIoKB/s'; e = { $_.MaxIoKBps } } -AutoSize | Out-String -Width 200 | Write-Host
    }

    # A 32-bit program has 2 GB of address space (4 GB if large-address-aware, on 64-bit
    # Windows). When its largest free block gets small, its allocations start failing.
    $va = @($procRows | Where-Object { $_.Kind -eq 'AddressSpace' })
    if ($va.Count) {
        Write-Host "`n32-BIT ADDRESS SPACE (the least room each 32-bit process had; under ~200 MB in one block, allocations start failing)" -ForegroundColor Yellow
        $va | Group-Object PID, Process | ForEach-Object {
            $worst = $_.Group | Sort-Object { ConvertTo-Number $_.LargestFreeBlockMB } | Select-Object -First 1
            [pscustomobject] @{
                Process = $worst.Process; PID = $worst.PID
                Largest = ConvertTo-Number $worst.LargestFreeBlockMB
                LargestFreeBlockMB = Format-Number (ConvertTo-Number $worst.LargestFreeBlockMB) '0'
                FreeAddressSpaceMB = Format-Number (ConvertTo-Number $worst.FreeAddressSpaceMB) '0'
                PrivateMB = Format-Number (ConvertTo-Number $worst.PrivateMB) '0'
                At = $worst.Time.ToString('MM-dd HH:mm')
            }
        } | Sort-Object Largest | Select-Object -First 8 |
            Format-Table Process, PID, LargestFreeBlockMB, FreeAddressSpaceMB, PrivateMB, At -AutoSize | Out-String -Width 160 | Write-Host
    }
}

if ($script:BadRows) { Write-Host "Skipped $($script:BadRows) damaged lines (usually the end of a file cut short by a power cut or crash)." }
if ($script:BadFiles.Count) { Write-Host "Could not read these files, which may be damaged or half-copied: $($script:BadFiles -join ', ')" -ForegroundColor Yellow }

if ($IncidentCsv) {
    $allIncidents | Export-Csv -LiteralPath $IncidentCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Saved $($allIncidents.Count) problem periods to $IncidentCsv"
}

