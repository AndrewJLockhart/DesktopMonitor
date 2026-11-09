# DesktopMonitor

**Find out why a Windows PC freezes, slows down, loses the network or crashes, even when
it happens while you're not looking.**

DesktopMonitor runs quietly in the background on a user's PC. Every 10 seconds it writes
down what the PC is doing:

- how busy the processor, memory, disk and network are;
- whether programs and the PC itself are responding;
- whether DNS, the domain controller, file servers and the web can be reached;
- how hard the security software is working;
- anything notable in the Windows event log.

Later, a report script reads the logs and tells you what went wrong, when, and the likely
cause.

- Two PowerShell scripts, nothing to install system-wide.
- No admin rights needed.
- Light: about 0.15% of the CPU and 100 MB of memory.
- Windows 10 and 11, with the PowerShell 5.1 that comes with Windows.

This README has two parts:

- **[Part 1: Getting started](#part-1-getting-started)**: from nothing to your first report
  in about ten minutes.
- **[Part 2: Reference](#part-2-reference)**: everything it records, every report section
  and setting, and how to deploy it.

---

# Part 1: Getting started

## Step 1. Put the scripts on the PC

Copy `DesktopMonitor.ps1` and `Get-DesktopMonitorReport.ps1` into a folder, for example
`C:\Tools\DesktopMonitor`. Then open **Windows PowerShell** (from the Start menu, *not*
Command Prompt) and go to that folder:

```powershell
cd C:\Tools\DesktopMonitor
```

If the files came from a download or an email, Windows may block them. Unblock them once:

```powershell
Unblock-File .\DesktopMonitor.ps1, .\Get-DesktopMonitorReport.ps1
```

> If PowerShell says *"running scripts is disabled on this system"*, start each command
> with `powershell -ExecutionPolicy Bypass -File`, for example
> `powershell -ExecutionPolicy Bypass -File .\DesktopMonitor.ps1 -Live`.

## Step 2. Try it

```powershell
.\DesktopMonitor.ps1 -Live
```

After a few seconds it prints one line every 10 seconds. Press **Ctrl+C** to stop.

```text
23:42:51  cpu 16%  mem 77%  commit 93%  disk 4% (0.9/0.2 ms)  net 8/2 KB/s  ping 0 ms  gpu 2%  top MsMpEng.exe  stall 17 ms  app 0 ms  taskbar 475 ms  gap 10.0s  took 75 ms
23:43:01  cpu 7%  mem 76%  commit 92%  disk 5% (1.5/0.3 ms)  net 4/3 KB/s  ping 0 ms  gpu 2%  top dwm.exe  stall 18 ms  app 0 ms  taskbar 0 ms  gap 10.0s  took 57 ms
```

What the line shows, left to right:

- **the time**
- **processor use**
- **memory:** RAM in use, and *commit*, the share of RAM plus page file that programs have
  claimed
- **disk:** how busy it is, and how long a read and a write take
- **network traffic:** in and out
- **ping time to the router**
- **GPU use**
- **the busiest program**
- **how long the PC, the program in front and the taskbar were unresponsive**
- **the time since the last sample, and how long this sample took to collect**

That's only a summary; the log files hold much more. They're written to
`C:\ProgramData\DesktopMonitor\Logs` while it runs.

## Step 3. Leave it running on the user's PC

Sign in as the user whose PC has the problem, and run:

```powershell
.\DesktopMonitor.ps1 -Install -FreezeShortcut
```

```text
Installed the 'DesktopMonitor-alice' scheduled task: it starts at logon and is checked every 5 minutes.
  Running with: (all defaults)
  Script:       C:\Users\alice\AppData\Local\DesktopMonitor\DesktopMonitor.ps1
  Logs:         C:\ProgramData\DesktopMonitor\Logs
  Added a 'Report a freeze' shortcut to the desktop: C:\Users\alice\Desktop\Report a freeze.lnk
```

From now on it starts by itself whenever that user signs in, runs hidden, and restarts
itself within 5 minutes if it stops. It doesn't need admin rights.

## Step 4. Ask the user to click "Report a freeze"

`-FreezeShortcut` put a **Report a freeze** icon on the user's desktop. Ask them to
double-click it *whenever* the PC freezes, is slow, or loses the network. A small popup
confirms it:

> Thanks. The freeze at 10:42:07 has been recorded.

That exact time is the best clue you will get. You can also record one yourself:

```powershell
.\DesktopMonitor.ps1 -Mark -Note 'Outlook froze for about 10 seconds'
```

## Step 5. Read the report

After the problem has happened a few times, run the report on that PC:

```powershell
.\Get-DesktopMonitorReport.ps1
```

It covers the last 24 hours. For a particular time, or several days:

```powershell
.\Get-DesktopMonitorReport.ps1 -From '2026-11-09 09:00' -To '2026-11-09 12:00'
.\Get-DesktopMonitorReport.ps1 -From (Get-Date).AddDays(-3)
```

Start with **PROBLEM PERIODS**. Each line is one problem: when it started, how long it
lasted, how bad it got, and the likely culprit.

```text
PROBLEM PERIODS

Start               Problem           Minutes Worst               Culprit
2026-11-09 23:19:39 HighHandleCount       1.1 TopHandles 174124   svchost.exe (fdPHost) (13x)
2026-11-09 23:19:39 MemoryPressure        1.1 CommitPct 94.8      explorer.exe (13x)
2026-11-09 23:19:44 ServerUnreachable       1 ServerFailed 1      Server: 10.255.255.1:445 didn't answer within 3 s (12x)
2026-11-09 23:19:49 HeavyPaging           0.4 PageInPerSec 3967.7 Memory Compression (3x)
```

Each line reads as plain English. For example:

- *A Windows service called fdPHost is holding 174,124 handles.* It's leaking them.
- *Memory is 95% committed.* The PC is close to running out.
- *A server didn't answer on port 445 (file sharing).*

Then look at **REPORTED FREEZES**. For each click on *Report a freeze* it shows the samples
from two minutes before to 30 seconds after, and the event log entries and connection
errors around it. [How to read it](#reported-freezes).

You can also report from your own PC, through the admin share or on a copy of the folder:

```powershell
.\Get-DesktopMonitorReport.ps1 -LogDirectory '\\PC042\c$\ProgramData\DesktopMonitor\Logs'
```

## Step 6. Remove it when you're done

```powershell
.\DesktopMonitor.ps1 -Uninstall
```

This removes the scheduled task, the copied script and the desktop icon. The logs stay in
`C:\ProgramData\DesktopMonitor\Logs` until you delete them.

## Quick recipes

| The user says... | Install with | Then look in the report for |
|---|---|---|
| "It freezes a few times a day" | `-Install -FreezeShortcut` | REPORTED FREEZES; **Stall**, **AppNotResponding**, **Freeze** |
| "The network keeps dropping" | `-Install -FreezeShortcut -PingTarget gateway,8.8.8.8` | **NetworkLoss**, **NetworkDown**, `WifiDisconnected` and `NetworkAdapter` events |
| "Files on the server / Outlook are slow or disconnect" | `-Install -FreezeShortcut -ServerTarget fileserver01:445,mail.corp.local:443` | **ServerUnreachable**, **DnsProblem**, **MappedDriveOffline** |
| "Logon is slow" / "drives don't map" | `-Install -FreezeShortcut` (domain checks are automatic) | **DomainControllerProblem**, **DomainNetworkNotRecognised**, `GroupPolicyFailed` events |
| "Websites are slow, internal apps are fine" | `-Install -FreezeShortcut` (the web check is on by default) | **WebAccessProblem** |
| "Everything is slow since the new security software" | `-Install -FreezeShortcut` | **SecurityAgentBusy**, **FileAccessSlow**, the AGENTS table |
| "It gets slower until I restart" | `-Install`, then wait a few days | GROWTH, PROCESS GROWTH, KERNEL POOL |
| "It crashed / restarted by itself" | `-Install` | NOTABLE EVENTS: `BlueScreen`, `UnexpectedReboot` |

Remember that **each `-Install` replaces all the earlier settings**. Give the full set every
time ([details](#installing-again-and-changing-settings)).

## How it works, in short

```text
 User's PC                                                Anywhere
 ┌───────────────────────────────────────────────┐        ┌──────────────────────────────┐
 │ DesktopMonitor.ps1 (hidden, every 10 s)       │        │ Get-DesktopMonitorReport.ps1 │
 │  ├─ counters: CPU, memory, disk, network      │  logs  │  ├─ problem periods          │
 │  ├─ background checks: stalls, app replies,   │ ─────► │  ├─ reported freezes         │
 │  │  DNS, DC, servers, web, file access        │  CSV   │  ├─ notable events           │
 │  ├─ processes, security agents, event log     │  files │  ├─ peaks and growth         │
 │  └─ "Report a freeze" clicks                  │        │  └─ agents, 32-bit programs  │
 └───────────────────────────────────────────────┘        └──────────────────────────────┘
```

The logs are ordinary CSV files, so you can also open them in Excel.

---

# Part 2: Reference

- [The report](#the-report)
  - [Problem types](#problem-types)
  - [Reported freezes](#reported-freezes)
- [The log files](#the-log-files)
  - [Metrics file: every column](#metrics-file-every-column)
  - [Processes file](#processes-file)
  - [Events file](#events-file)
  - [Inventory file](#inventory-file)
- [Reading the logs: symptom by symptom](#reading-the-logs-symptom-by-symptom)
- [Freezes: how short freezes are caught](#freezes-how-short-freezes-are-caught)
- [Network and Active Directory checks](#network-and-active-directory-checks)
- [Security and monitoring software](#security-and-monitoring-software)
- [32-bit programs and address space](#32-bit-programs-and-address-space)
- [Next steps with other tools](#next-steps-with-other-tools)
- [Installing, deploying and checking](#installing-deploying-and-checking)
- [Log size and history](#log-size-and-history)
- [Impact on the PC](#impact-on-the-pc)
- [What it collects (privacy)](#what-it-collects-privacy)
- [All parameters](#all-parameters)

## The report

```powershell
.\Get-DesktopMonitorReport.ps1 [-LogDirectory <folder>] [-From <time>] [-To <time>] [-ComputerName <PC>] [-IncidentCsv <file>]
```

| Parameter | Default | Meaning |
|---|---|---|
| `-LogDirectory` | `C:\ProgramData\DesktopMonitor\Logs` | The log folder, a copy of it, or an admin share such as `\\PC042\c$\ProgramData\DesktopMonitor\Logs`. |
| `-From` | 24 hours before `-To` | Start of the period. |
| `-To` | the newest sample | End of the period. |
| `-ComputerName` | all | One PC, when the folder holds several PCs' logs. |
| `-IncidentCsv` | | Also save the problem periods to a CSV file. |

It reads zipped days as they are, and skips lines damaged by a power cut, saying how many.
It prints these sections for each PC:

| Section | Contents |
|---|---|
| PROBLEM PERIODS | Each problem found: start, problem type, minutes, worst value, likely culprit. See [Problem types](#problem-types). Freezes that overlap sleep or standby are left out, with a note. |
| REPORTED FREEZES | Around each *Report a freeze* click. See [Reported freezes](#reported-freezes). |
| NOTABLE EVENTS | The last 40 notable event log entries (crashes, blue screens, driver resets, disk errors, network and domain problems, security software events), and a count of every tag. |
| PEAKS | Average, 95th percentile and maximum of the key measurements, and when the maximum was. |
| GROWTH | First hour against last hour for handles, GDI/USER objects, kernel pools, commit and ports. Steady growth over days means a leak. Needs 2 hours of samples. |
| KERNEL POOL | The biggest kernel memory tag at the start, the end and the peak. |
| PROCESS GROWTH | The processes whose memory, handles, GDI/USER objects, TCP or UDP grew the most. |
| SECURITY AND MONITORING AGENTS | Each agent's average and peak CPU, peak memory and peak I/O. |
| 32-BIT ADDRESS SPACE | The 32-bit programs with the least memory room left. |

### Problem types

A problem is reported when its condition holds for at least the number of samples shown
(one sample is 10 seconds by default). The *Culprit* column is the most frequent value of
the column named, with how many samples it appeared in.

| Problem | Reported when | Samples | Culprit from |
|---|---|---|---|
| **Freeze** | The gap since the previous sample was more than twice the interval and more than the interval + 5 s, or the sample took over 3 s to collect. Not reported when the CPU was at 90%+ and no stall was seen: that is the monitor waiting for CPU time. | 1 | `ForegroundApp` |
| **Stall** | `MaxStallMs` ≥ 1000, or 3+ stalls of 250 ms in one sample. Nothing on the PC could run. | 1 | `EventTags` |
| **AppNotResponding** | The program in front took ≥ 2 s to respond. | 1 | `ForegroundResponseApp` |
| **TaskbarNotResponding** | The taskbar (File Explorer) took ≥ 1 s to respond. | 1 | `TopCpuProcess` |
| **HungApp** | Windows showed a program as *Not Responding*. | 1 | `HungApps` |
| **DpcStorm** | One CPU core spent ≥ 20% of its time on driver interrupts. | 1 | a driver |
| **CpuSaturated** | CPU ≥ 90%. | 3 | `TopCpuProcess` |
| **GpuSaturated** | GPU ≥ 95%. | 3 | `TopGpuProcess` |
| **MemoryPressure** | Commit ≥ 90% or RAM ≥ 95% used. | 1 | `TopMemProcess` |
| **HeavyPaging** | ≥ 1000 pages a second read from disk. | 2 | `TopHardFaultProcess` |
| **DiskSlow** | Disk ≥ 90% busy, or reads or writes taking ≥ 50 ms on average (with at least 5 a second). | 3 | `TopIoProcess` |
| **CpuThrottled** | The CPU ran below 60% of its base speed while at least 20% busy, or a thermal zone was limiting it. | 3 | `TopCpuProcess` |
| **NetworkLoss** | A ping failed. | 1 | `NetAdapters` |
| **NetworkDown** | Windows saw no network, or no traffic. | 1 | `NetAdapters` |
| **DnsProblem** | A DNS lookup took ≥ 1 s or failed. | 1 | `ConnectivityErrors` |
| **DomainControllerProblem** | Connecting to the domain controller took ≥ 1 s or failed. | 1 | `ConnectivityErrors` |
| **DomainNetworkNotRecognised** | The DC answers, but Windows didn't class the network as the domain network. | 3 | `NetworkCategory` |
| **ServerUnreachable** | A `-ServerTarget` took ≥ 1 s to connect or failed. | 1 | `ConnectivityErrors` |
| **WebAccessProblem** | The web check took ≥ 5 s or failed. | 1 | `ConnectivityErrors` |
| **MappedDriveOffline** | A mapped drive's server didn't answer. | 1 | `MappedDrivesDown` |
| **SecurityAgentBusy** | Security and monitoring agents together used ≥ 25% of the CPU. | 3 | `TopAgentProcess` |
| **FileAccessSlow** | Writing and reading a 64 KB test file took ≥ 500 ms, or failed. | 1 | `TopAgentProcess` |
| **LowDiskSpace** | Less than 10% free on the system drive. | 1 | `TopIoProcess` |
| **GdiNearLimit** | A process had ≥ 8000 GDI objects, or the session ≥ 50,000. | 1 | `TopGdiProcess` |
| **UserObjNearLimit** | A process had ≥ 8000 USER objects. | 1 | `TopUserObjProcess` |
| **HighHandleCount** | A process had ≥ 100,000 handles. | 1 | `TopHandleProcess` |
| **TcpPortPressure** | ≥ 10,000 TCP connections. | 1 | `TopTcpProcess` |
| **UdpPortPressure** | ≥ 10,000 UDP endpoints. | 1 | `TopUdpProcess` |
| **AddressSpaceLow** | A 32-bit program's largest free memory block < 200 MB, or < 500 MB free in all. | 1 | `Top32BitProcess` |

### Reported freezes

For each *Report a freeze* click (the last 10 in the period), the report shows a row per
sample from two minutes before to 30 seconds after. Read across the row at the time of the
click:

| Column | High values mean |
|---|---|
| `StallMs` | In the thousands: the whole PC stopped (a driver, the firmware or the hardware). |
| `AppMs`, `App` | In the thousands: that program stopped responding. |
| `TaskbarMs` | In the thousands: File Explorer and the taskbar stopped responding. |
| `Cpu%`, `DpcCore%` | Processor busy; a driver hogging one core. |
| `Commit%`, `PageIn/s` | Out of memory and swapping to disk. |
| `Disk%` | Disk busy. |
| `DnsMs`, `DcMs`, `Ping` | Waiting on DNS, the domain controller or the network. |
| `FileMs`, `Agent%` | Security software slowing file access, or busy itself. |
| `TopCpu`, `Hung` | The busiest program, and programs shown as *Not Responding*. |

Below the table it lists event log entries from five minutes before to one minute after,
such as `WifiDisconnected`, `DnsTimeout` or `DefenderScanStarted`. It also lists any
connection errors in those samples.

## The log files

Everything goes in `-LogDirectory`, `C:\ProgramData\DesktopMonitor\Logs` by default:

| File | Contents |
|---|---|
| `metrics_<PC>_<date>.csv` | One row per sample. [Every column](#metrics-file-every-column). |
| `processes_<PC>_<date>.csv` | Top processes in several rankings, agents, and process starts and exits. [Details](#processes-file). |
| `events_<PC>_<date>.csv` | Notable event log entries, the monitor's own start and stop, and freeze reports. [Details](#events-file). |
| `inventory_<PC>_<time>.csv` | Hardware, software, network, domain and security setup, written at every start. [Details](#inventory-file). |
| `monitor_<PC>_<user>.log` | The monitor's own messages: starts, stops, warnings, zipping, and which part of a sample was slow if one took over 2 seconds. It rolls over at 5 MB to `.old`. |
| `DesktopMonitor_<PC>_<date>.zip` | A finished day's CSV files, zipped after midnight. |
| `mark_*.txt` | A freeze report waiting to be picked up. It's deleted within one interval. |

All files are UTF-8 CSV with a header row (`-Format Tsv` gives tab-separated files). They
open in Excel, Power BI, pandas or `Import-Csv`. Every PC's files have the same columns; a
measurement a PC doesn't support, or a group switched off with `-Exclude`, is left blank. A
new file is started each day, when a file reaches `-MaxFileSizeMB`, and when a newer
version of the script changes the columns (as `name_1.csv`). The report reads them all.

### Metrics file: every column

Values are for the time since the previous sample unless marked *(now)*, or *(once a
minute)*, in which case the value is repeated in the rows between.

**Core**

| Column | Meaning |
|---|---|
| `Timestamp` | Local time of the sample. |
| `GapSec` | Seconds since the previous sample, timed with a stopwatch (clock changes don't affect it). Blank on the first sample after a start. Much bigger than the interval means the monitor itself was held up. |
| `CollectMs` | How long collecting the sample took, not counting waiting for pings or the start-up catch-up. |
| `UptimeHours` | Hours since Windows last fully started. With Fast Startup, a shutdown doesn't reset it. |

**Cpu**

| Column | Meaning |
|---|---|
| `CpuPct` | Processor use, % of the whole PC. |
| `CpuPrivilegedPct` | Of which in Windows itself and drivers. |
| `CpuInterruptPct`, `CpuDpcPct` | Of which handling hardware interrupts and deferred driver work, averaged over all cores. |
| `CpuMaxCorePct` | The busiest single core. |
| `CpuMaxCoreDpcPct` | The most time any one core spent on interrupts and deferred driver work. Over ~20%: a driver is hogging a core (mouse lag, crackling sound). |
| `CpuPerfPct`, `CpuMHz` | Actual speed as a % of base speed, and in MHz. Well below 100 under load: throttling. |
| `CpuQueue` | Threads waiting for a processor *(now)*. |
| `ContextSwitchesPerSec` | Thread switches per second. |
| `Processes`, `Threads` | How many are running *(now)*. |

**Memory**

| Column | Meaning |
|---|---|
| `MemUsedPct`, `MemAvailableMB` | RAM in use, and RAM available *(now)*. |
| `MemModifiedMB` | Changed memory waiting to be written to disk *(now)*. |
| `CommitPct`, `CommitGB`, `CommitLimitGB` | Memory promised to programs, against RAM + page file *(now)*. Near 100%, programs fail to get memory. |
| `PageInPerSec`, `PageOutPerSec` | Pages read from and written to disk per second. Thousands: the PC is swapping. |
| `PageFileUsedPct` | Page file in use *(now)*. |
| `PoolNonPagedMB`, `PoolPagedMB` | Windows kernel memory *(now)*. Growing for hours: a driver or handle leak. |
| `TopNonPagedPoolTag`, `TopNonPagedPoolMB`, `TopPagedPoolTag`, `TopPagedPoolMB` | The biggest kernel memory user of each kind, by its 4-letter tag *(once a minute)*. |

**Disk**

| Column | Meaning |
|---|---|
| `DiskBusyPct`, `DiskQueue` | How busy the physical disks were, and requests waiting. |
| `DiskReadMs`, `DiskWriteMs` | Average time per read and write. Over ~50 ms is slow. |
| `DiskReadMBps`, `DiskWriteMBps`, `DiskIops` | Throughput and operations per second. |
| `SysDriveFreeGB`, `SysDriveFreePct` | Free space on the Windows drive *(now)*. |
| `OtherDrivesFree` | Free space on other fixed drives, for example `D:120.5GB(40%)` *(now)*. |

**Network**

| Column | Meaning |
|---|---|
| `NetRecvKBps`, `NetSentKBps` | Traffic on the physical network adapters. |
| `NetErrors` | Packets received or sent with errors. |
| `TcpConnections` | Established TCP connections, IPv4 and IPv6 *(now)*. |
| `TcpRetransPerSec` | TCP segments sent again because they were lost. High: a poor network. |
| `NetAdapters` | Connected adapters and link speeds *(once a minute)*. |
| `WifiSignalPct` | Wi-Fi signal strength *(once a minute)*. |
| `Ping1Ms`, `Ping2Ms` | Ping time to each `-PingTarget`. Blank: no reply. |
| `PingFailed` | How many pings got no reply. |

**Connectivity**: see [Network and Active Directory checks](#network-and-active-directory-checks)

| Column | Meaning |
|---|---|
| `DnsLookupMs`, `DnsFailed` | Slowest DNS lookup, and failed lookups. |
| `DomainController` | The DC in use and its site, e.g. `dc01.corp.local (site London)`, or *"... but this PC is in site Paris"*. |
| `DcConnectMs`, `DcFailed` | Time to connect to the DC, and failures. |
| `Server1Ms`, `Server2Ms`, `ServerFailed` | Time to connect to each `-ServerTarget`, and failures. |
| `HttpMs`, `HttpFailed` | Time to fetch the `-HttpTarget` page through the proxy, and failures *(once a minute)*. |
| `MappedDrivesDown` | Mapped drives whose server didn't answer, e.g. `S: \\fs01\shared` *(once a minute)*. |
| `NetworkCategory` | How Windows classes the connected networks: `Domain`, `Private` or `Public` *(once a minute)*. |
| `WindowsConnectivity` | `Internet`, `LocalNetwork`, `NoTraffic` or `Disconnected`, as Windows sees it *(once a minute)*. |
| `Vpn` | Connected VPN adapters *(once a minute)*. |
| `DnsServers` | DNS servers in use *(once a minute)*. |
| `NetChanges` | Network changes: adapters connecting or dropping, addresses changing. |
| `ConnectivityErrors` | Why checks failed, in words. |

**Gpu, Thermal, Power**

| Column | Meaning |
|---|---|
| `GpuPct`, `Gpu3DPct` | The busiest GPU engine, and the busiest 3D engine. |
| `GpuDedicatedMB`, `GpuSharedMB` | GPU memory in use *(now)*. |
| `TopGpuProcess`, `TopGpuPct` | The program using the GPU most. |
| `ThermalMaxC`, `ThermalLimitPct` | Hottest ACPI thermal zone, and how far it is limiting the CPU (100 = not at all). Many desktops don't report these. |
| `OnAC`, `BatteryPct`, `BatterySaver` | On mains power (1/0), battery level, battery saver on (1/0) *(now)*. |

**Session**

| Column | Meaning |
|---|---|
| `ConsoleUser` | Who is signed in at the PC's own screen *(now)*. |
| `IdleSec` | Seconds since the user last used the keyboard or mouse *(now)*. |
| `ForegroundApp` | The program in front *(now)*. |
| `HungAppCount`, `HungApps` | Programs Windows shows as *Not Responding* *(now)*. |

**Responsiveness**: see [Freezes](#freezes-how-short-freezes-are-caught)

| Column | Meaning |
|---|---|
| `MaxStallMs`, `StallCount` | The longest time nothing on the PC could run, and how many stalls of 250 ms or more. Normally 10–30 ms. |
| `ForegroundResponseMs`, `ForegroundResponseApp` | Slowest reply from the program in front, and which program. 5000 = no reply in 5 s. |
| `ShellResponseMs` | Slowest reply from the taskbar. |

**Security**: see [Security and monitoring software](#security-and-monitoring-software)

| Column | Meaning |
|---|---|
| `AgentCpuPct`, `AgentIoKBps`, `AgentMemMB` | All security and monitoring agents together. |
| `TopAgentProcess`, `TopAgentCpuPct` | The busiest agent, with its product name. |
| `FileCheckMs`, `FileCheckFailed` | Time to write, read back and delete a 64 KB test file. Normally a few ms. |

**Processes**

| Column | Meaning |
|---|---|
| `Handles`, `GdiObjects`, `UserObjects` | Totals: kernel handles for the whole PC; GDI and USER objects for the user's session *(now)*. |
| `TcpPortsInUse`, `TcpTimeWait`, `UdpEndpoints` | TCP connections (not counting listeners), those in TIME_WAIT, and UDP endpoints *(now)*. Windows has 16,384 ports for each. |
| `TopCpuProcess`/`TopCpuPct`, `TopIoProcess`/`TopIoKBps`, `TopHardFaultProcess`/`TopHardFaultsPerSec` | The busiest process for CPU, I/O and reads from disk it had to wait for. |
| `TopMemProcess`/`TopMemMB`, `TopHandleProcess`/`TopHandles`, `TopGdiProcess`/`TopGdiObjects`, `TopUserObjProcess`/`TopUserObjects`, `TopTcpProcess`/`TopTcpConnections`, `TopUdpProcess`/`TopUdpEndpoints` | The biggest user of each resource *(now)*. An `svchost.exe` is shown with its services, e.g. `svchost.exe (Dnscache)`. |
| `Top32BitProcess`, `LargestFreeBlockMB`, `FreeAddressSpaceMB` | The 32-bit program with the least memory room *(once a minute)*. |

**Events and self**

| Column | Meaning |
|---|---|
| `EventCount`, `EventTags` | New event log entries since the previous sample, and their tags. |
| `MonitorCpuPct`, `MonitorMemMB` | The monitor's own CPU (% of the whole PC) and memory. |

### Processes file

One row per process per list. The `Kind` column says which list:

| Kind | When | Which processes |
|---|---|---|
| `Cpu`, `Io`, `HardFaults` | every sample | The top 10 using ≥ 1% CPU, ≥ 100 KB/s I/O, or ≥ 20 reads from disk a second. |
| `Memory`, `Handles`, `Gdi`, `User`, `Tcp`, `Udp` | once a minute | The top 10 users of each. |
| `AddressSpace` | once a minute | The 10 32-bit programs with the least memory room. |
| `Agent` | once a minute | The 10 busiest security and monitoring agents. |
| `Started`, `Exited` | every sample | Every process that started or ended. |

The columns are:

- `Timestamp`, `Kind`, `Rank`, `Process`, `PID` and `ParentPID`.
- `CpuPct`, `IoKBps`, `HardFaultsPerSec`, `PrivateMB`, `WorkingSetMB`, `Handles`, `GdiObjects`,
  `UserObjects`, `TcpConnections`, `UdpEndpoints`, `Threads`, `LargestFreeBlockMB`,
  `FreeAddressSpaceMB` and `SessionId`.
- `StartedAt`.
- `Services`: the Windows services in that process.
- `Agent`: the security or monitoring product, if it is one.

`-TopProcessCount` changes the 10.

### Events file

Columns: `Timestamp`, `Tag`, `Level`, `Log`, `Provider`, `EventId`, `Message`.

The monitor keeps these entries:

- **System log:** every critical, error and warning entry (except two common DCOM
  warnings), plus boot, shutdown, sleep and wake.
- **Application log:** every critical and error entry, plus Windows Error Reporting and
  Windows VPN failures.
- **Once a minute:** network connects and disconnects, Windows' connectivity changes, Wi-Fi
  connects and drops, Defender scans, updates and detections, and DHCP problems.

Each entry gets a `Tag`:

| Area | Tags |
|---|---|
| Crashes and restarts | `BlueScreen`, `UnexpectedReboot`, `UnexpectedShutdown`, `AppCrash`, `AppHang`, `ServiceCrash`, `ErrorReport`, `RestartRequested`, `Boot`, `Shutdown` |
| Sleep | `Sleep`, `Wake` (including Modern Standby) |
| Hardware | `HardwareError`, `DiskError`, `GpuError`, `GpuDriverReset`, `CpuThrottled`, `LowMemory` |
| Network | `NetworkAdapter`, `NetworkConnected`, `NetworkDisconnected`, `ConnectivityChange`, `Wifi`, `WifiConnected`, `WifiConnectFailed`, `WifiDisconnected`, `Dhcp`, `IpAddressConflict`, `PortExhaustion`, `DnsTimeout`, `VpnFailed` |
| Domain and sign-in | `NoDomainController`, `SecureChannel`, `GroupPolicyFailed`, `GroupPolicyNoNetwork`, `Kerberos`, `AuthFailure`, `TimeSync` |
| Security software | `DefenderScanStarted`, `DefenderScanFinished`, `DefenderScanStopped`, `DefenderUpdate`, `DefenderUpdateFailed`, `DefenderProblem`, `DefenderOff`, `MalwareDetected`, `MalwareAction` |
| The monitor | `MonitorStart` (says whether Windows restarted, was shut down, or slept since the last sample), `MonitorStop`, `UserReportedFreeze` |
| Anything else | `Critical`, `Error`, `Warning`, `Information` |

### Inventory file

Columns `Section`, `Name`, `Value`, written each time the monitor starts:

| Section | What it records |
|---|---|
| Monitor | Version, settings, the account it runs as. |
| OS | Windows version and build, install date, last full boot, Fast Startup, pending restart. |
| Hardware | Make, model, serial number, BIOS, CPU, memory modules, GPU and driver, page file. |
| Disks | Each disk's type and health, each volume's free space. |
| Network | Each adapter's IP, DHCP, DNS servers and suffix; network card drivers and power saving; VPN adapters; the default gateway. |
| Proxy | WinHTTP proxy, the user's proxy, PAC script and automatic detection (WPAD), mapped drives. |
| Domain | Domain membership, logon server, the domain controller and AD sites, the Entra ID (Azure AD) join state, the time source. |
| Security | Antivirus, antispyware and firewall products (on/off, up to date); Defender's mode and last scans; file-system filter drivers with their altitude; DLLs injected into programs; AMSI scanners. |
| ThirdParty | Every running driver and service not from Microsoft, with its publisher. |
| Software | Power plan, antivirus, the five most recent updates. |

## Reading the logs: symptom by symptom

| Symptom | Look for |
|---|---|
| The whole PC freezes for a moment | `MaxStallMs` ≥ 1000; `GpuDriverReset`, `DiskError`, `HardwareError` events at the same time |
| Mouse lag, crackling sound, stutter, CPU looks low | `CpuMaxCoreDpcPct` above ~20: a driver is hogging a core |
| One program freezes ("Not Responding") | `ForegroundResponseMs` in the thousands with `ForegroundResponseApp`; `HungApps`; `AppHang` events |
| Taskbar, Start or File Explorer freezes | `ShellResponseMs` in the thousands; `MappedDrivesDown` |
| Programs freeze waiting on the network | `DnsLookupMs`, `DcConnectMs`, `Server1Ms`/`Server2Ms` high; `ConnectivityErrors` |
| Network drops | `PingFailed`, `WindowsConnectivity`, `NetChanges`, `WifiSignalPct`; `NetworkAdapter`, `WifiDisconnected`, `Dhcp` events |
| Web slow, internal apps fine | `HttpMs`, `HttpFailed` (the proxy, PAC/WPAD, TLS inspection) |
| Slow logon, Group Policy or drive maps fail | `DcConnectMs`, `DomainController` site, `NetworkCategory`; `NoDomainController`, `SecureChannel`, `GroupPolicyFailed`, `TimeSync` events |
| Everything slow, nothing obvious in Task Manager | `FileCheckMs`, `AgentCpuPct`, `TopAgentProcess`; Defender scan events |
| General slowness | `CpuPct` ~100 (`TopCpuProcess`); `CommitPct` > 90 or high `PageInPerSec` (`TopHardFaultProcess`); `DiskBusyPct` ~100 with `DiskReadMs` > 50; `GpuPct` ~100 |
| Slow under load, fine when idle | `CpuPerfPct` well below 100, `ThermalLimitPct` < 100; `CpuThrottled` events |
| Worse over days | `Handles`, `GdiObjects`, `PoolPagedMB`, `PoolNonPagedMB`, `CommitGB` climbing |
| Kernel memory large | `TopPagedPoolTag`/`TopNonPagedPoolTag` in the hundreds of MB. `Toke`, `Proc` or `Thre` usually mean a program leaking handles (see `TopHandleProcess`). Any other tag is usually a driver: `findstr /m /l <tag> C:\Windows\System32\drivers\*.sys` |
| Windows don't redraw, blank menus | `TopGdiObjects` or `TopUserObjects` near 10,000 |
| New connections or DNS lookups fail, network fine | `TcpPortsInUse`/`UdpEndpoints` in the tens of thousands; `TopTcpProcess`/`TopUdpProcess`; `PortExhaustion` events |
| "Out of memory" with plenty free | `LargestFreeBlockMB` under ~200 for `Top32BitProcess` |
| Crash or restart | `BlueScreen`, `UnexpectedReboot`, `UnexpectedShutdown` events; `MonitorStart` rows |

## Freezes: how short freezes are caught

Most counters are averages over the 10-second sample, so a 2-second freeze barely shows in
them. Background threads catch freezes that short:

- **The stall watch** (`MaxStallMs`, `StallCount`). A top-priority thread wakes every 100 ms
  and records how late it woke. If it was late and no other program got the CPU either,
  nothing on the PC could run: a driver, the firmware or the hardware. If other programs
  did run meanwhile, only the monitor was paused, and it isn't counted.
- **Responsiveness probes** (`ForegroundResponseMs`, `ShellResponseMs`). Every second the
  monitor sends a do-nothing message to the window in front and to the taskbar, and times
  the reply. That's how long the program would have ignored a click or a key. Windows only
  calls a window *Not Responding* after 5 seconds; the probes catch shorter freezes too.
  Programs run as administrator can't be probed.

**Reading them together:**

- **A stall:** the whole PC stopped.
- **A slow program with no stall:** that program was waiting on something. That could be
  the disk, memory, the network (`DnsLookupMs`, `DcConnectMs`), security software
  (`FileCheckMs`), or its own work (`TopCpuProcess`).

**Sleep, standby and clock changes.** The monitor stops while the PC sleeps or is in
standby, and that looks like a freeze. The report recognises sleep from the event log and
leaves those periods out. The monitor's sampling thread runs level with normal programs,
so a busy PC doesn't delay it. Summer-time clock changes don't affect the timings.

**Reporting a freeze.** `-Mark` (or the *Report a freeze* icon) records the exact time, the
user and an optional `-Note`. While the monitor runs, `-Mark` leaves a small `mark_*.txt`
file that the monitor turns into a `UserReportedFreeze` event within one interval. If the
monitor isn't running, it writes the event directly, and says that there are no
measurements to go with it. `-Popup` shows the confirmation in a window, which is what the
icon uses. If the logs are somewhere other than the default, `-Mark` needs the same
`-LogDirectory` (the icon includes it).

## Network and Active Directory checks

Each check runs on its own background thread, so one that hangs never delays the samples.
`ConnectivityErrors` says why a check failed, for example
`DNS: _ldap._tcp.dc._msdcs.corp.local: This operation returned because the timeout period expired`.

| Check | How often | Catches |
|---|---|---|
| **Ping** `-PingTarget` (the router by default; add an internet address to compare) | every sample | Link drops and packet loss. A failing router ping means the PC's own link (cable, Wi-Fi, network card). Only the internet failing means the problem is further out. |
| **DNS**: a fresh lookup from the DNS servers, bypassing the PC's cache | every sample | Slow or failing DNS, which almost everything waits on. |
| **Domain controller**: found the way Windows finds it, then connected to (LDAP, port 389) | every sample, domain PCs only | A DC that is slow, unreachable, or in another site (WAN). |
| **Your servers** (`-ServerTarget host:port`, up to two) | every sample | File, mail, proxy or application servers not answering. Works where ping is blocked. |
| **Web page through the user's proxy settings** (`-HttpTarget`) | once a minute | A slow PAC script or WPAD, proxy sign-in problems, TLS inspection, block or sign-in pages. |
| **Mapped drives**: each drive's server, on port 445 | once a minute | A drive whose server is gone, which makes File Explorer and Office hang. |
| **Windows' view of the network** (Network List Manager) | once a minute | No internet as Windows sees it. A domain network classed as Public or Private means the wrong firewall rules apply. |
| **VPN and DNS servers** | once a minute | Whether a VPN was up, and which DNS servers were in use at the time. |
| **Network changes** | every sample | Adapters connecting or dropping, addresses changing. |

**Defaults.**

- **DNS name:** on a domain PC, the records Windows uses to find a domain controller
  (`_ldap._tcp.dc._msdcs.<domain>`), which sign-in, Group Policy and file shares depend on.
  Otherwise `www.msftconnecttest.com`.
- **Web page:** `http://www.msftconnecttest.com/connecttest.txt`, the page Windows itself
  uses to check for internet access. The check confirms the expected text comes back, not
  a proxy page.

```powershell
# Check an intranet site instead, plus a file server and a proxy
.\DesktopMonitor.ps1 -Install -FreezeShortcut -HttpTarget https://intranet.corp.local/ -ServerTarget fs01:445,proxy01:8080

# Look up a name of your choice; no web check
.\DesktopMonitor.ps1 -Install -FreezeShortcut -DnsName mail.corp.local -HttpTarget none
```

The event log tags for network and domain problems are listed under
[Events file](#events-file). The SMB client and Group Policy *operational* logs need admin
rights, so they aren't read. The Group Policy failures that matter also appear in the
System log, which the monitor does read.

## Security and monitoring software

Antivirus, EDR, DLP, web filtering, VPN, employee monitoring and management agents sit in
the path of almost everything a PC does. They check every file opened, every program
started and every connection. When one is busy or misbehaving, the whole PC slows down,
and Task Manager often doesn't show why.

**How agents are recognised.** By the **publisher of their program file**. The list includes
CrowdStrike, SentinelOne, Sophos, Symantec/Broadcom, McAfee/Trellix, Trend Micro, Palo Alto
Networks, Cylance, ESET, Kaspersky, Bitdefender, Carbon Black, Cybereason, Forcepoint,
Digital Guardian, Zscaler, Netskope, Cisco, Tanium, Qualys, Rapid7, Ivanti, BeyondTrust,
CyberArk, Varonis, Nexthink, Teramind, ActivTrak, Kaseya, ConnectWise, NinjaOne and
others. Microsoft Defender, Defender for Endpoint, Configuration Manager (SCCM) and
Intune are recognised by name. The program behind each Windows service is checked, so
agents running as SYSTEM are found too. To add one it doesn't know:

```powershell
.\DesktopMonitor.ps1 -Install -FreezeShortcut -AgentName 'acmeagent*.exe','tracker.exe'
```

**What is measured.**

| Where | What |
|---|---|
| `AgentCpuPct`, `AgentIoKBps`, `AgentMemMB`, `TopAgentProcess` | The agents' combined load, and the busiest one, every sample. |
| `FileCheckMs` | Writing, reading back and deleting a 64 KB file in the user's temp folder. Normally a few ms. Security products check files as they're written and opened; when they struggle, every program that touches files waits like this. |
| processes file, `Kind` = `Agent` | Each agent's CPU, memory, handles and I/O, once a minute. |
| events file | Defender scans, updates and detections. Scans and updates often line up with slow periods. |
| inventory file, `Security` section | The products Windows Security Center knows, whether each is on and up to date. Defender's mode (normal, passive, EDR block mode) and its last updates and scans. **File-system filter drivers** with their *altitude*, which says what kind they are: 320000–329999 antivirus, 360000–389999 activity monitoring (EDR), 260000–269999 content screening (DLP), 140000–149999 encryption. Every file opened on the PC passes through each of them. The **DLLs security products inject** into programs. The **AMSI scanners** that check every PowerShell script and Office macro. |

## 32-bit programs and address space

A 32-bit program can use at most 2 GB of memory (4 GB if it's marked large-address-aware),
however much RAM the PC has. Over time that space fragments. Once no free block is big
enough, the program fails with "out of memory", blank windows or a crash, while the PC has
plenty free. Once a minute the monitor measures every 32-bit process in the user's session.
Typical candidates are 32-bit Office (Outlook with a large mailbox), older business
applications and 32-bit Java. Programs run as administrator can't be measured. 64-bit
programs never run short.

## Next steps with other tools

Once the logs point somewhere, these dig deeper. Handle, Process Monitor and LatencyMon need
administrator rights.

| The logs show | Next step |
|---|---|
| `TopHandles` high or growing | `handle.exe -s -p <PID>` ([Sysinternals Handle](https://learn.microsoft.com/sysinternals/downloads/handle)) counts the handles by type. For an `svchost.exe`, the services in brackets say which service to restart or update. |
| `CpuMaxCoreDpcPct` high | [LatencyMon](https://www.resplendence.com/latencymon) names the driver. |
| A large pool tag | [PoolMon](https://learn.microsoft.com/windows-hardware/drivers/devtest/poolmon) shows every tag live. |
| A program freezes | `procdump.exe -h <name or PID>` ([ProcDump](https://learn.microsoft.com/sysinternals/downloads/procdump)) saves a dump when its window hangs, for the vendor. |
| Unexplained file or registry activity | [Process Monitor](https://learn.microsoft.com/sysinternals/downloads/procmon), for a few minutes. |
| DNS or DC problems | `nltest /dsgetdc:<domain>`, `Resolve-DnsName <name> -DnsOnly`, `Test-NetConnection <dc> -Port 389` |
| A busy security agent | The vendor's console: scan schedules, exclusions and policy changes at that time. |
| Slow at every logon | [Autoruns](https://learn.microsoft.com/sysinternals/downloads/autoruns) lists everything that starts with Windows. |

## Installing, deploying and checking

### What `-Install` does

Everything runs as the signed-in user; no admin rights are needed. `-Install`:

1. copies the script to `%LOCALAPPDATA%\DesktopMonitor` (`-InstallDirectory` changes it);
2. creates a scheduled task, `DesktopMonitor-<username>`, that starts the monitor hidden at
   logon and restarts it within 5 minutes if it stops. If the organisation stops users
   creating scheduled tasks, it adds the monitor to the user's startup programs instead (the
   HKCU `Run` key), which starts it at logon but doesn't restart it;
3. with `-FreezeShortcut`, adds the *Report a freeze* icon to the user's desktop;
4. starts the monitor.

It runs only while that user is signed in, so anything before logon is missed. Hangs from
the previous session still show up as `AppHang` events.

### Installing again and changing settings

There's only ever one install per user. Running `-Install` again compares itself with
what's there:

| Situation | What happens |
|---|---|
| Not installed | Installs and starts it. |
| Same settings and same script | Nothing changes; it's started if it isn't running. Safe to run at every logon. |
| Different settings | Shows the old and new settings, and restarts the monitor with the new ones. |
| A newer copy of the script | Copies it in and restarts the monitor. |
| A different `-InstallDirectory` | Moves the script there. |

**The settings given with `-Install` replace the previous ones completely. They aren't
merged.** A setting you leave out goes back to its default, so to change one, give the
whole set again:

```powershell
# First install
.\DesktopMonitor.ps1 -Install -FreezeShortcut -ServerTarget fs01:445
# Later, to add an internet ping: -FreezeShortcut and -ServerTarget must be given again
.\DesktopMonitor.ps1 -Install -FreezeShortcut -ServerTarget fs01:445 -PingTarget gateway,8.8.8.8
```

`-Uninstall` stops the monitor and removes the task (or startup entry), the copied script
and the icon. It keeps the logs.

### Deploying to many PCs

Because `-Install` needs no admin rights and does nothing when nothing has changed, it can
run as a **user logon script** (Group Policy, Intune or your RMM tool) from a file share:

```text
powershell.exe -NoProfile -ExecutionPolicy Bypass -File \\server\share\DesktopMonitor\DesktopMonitor.ps1 -Install -FreezeShortcut
```

To roll out a new version, replace the script on the share; each PC picks it up at the
next logon. Keep the logs on the PC (the default) rather than on a share, so logging
carries on when the network is the problem. Collect them afterwards, or report over the
admin share.

### Checking that it's running

```powershell
# The scheduled task and when it last ran
Get-ScheduledTask -TaskName "DesktopMonitor-$env:USERNAME" | Get-ScheduledTaskInfo

# The monitor's own messages
Get-Content "C:\ProgramData\DesktopMonitor\Logs\monitor_$($env:COMPUTERNAME)_$($env:USERNAME).log" -Tail 20

# The newest sample
Get-ChildItem C:\ProgramData\DesktopMonitor\Logs\metrics_* | Sort-Object LastWriteTime | Select-Object -Last 1
```

### Several users on one PC

Only one monitor writes to a log folder at a time. If two users are signed in at once, the
second user's monitor sees the first and exits. When the first signs out, the second
user's task starts it again within 5 minutes. `ConsoleUser` shows who is actually at the
screen. When the monitor creates the log folder, it lets every user of the PC change the
files in it, so each user's monitor can add to, zip and tidy them.

### What can block it

- **An execution policy set by Group Policy.** `-ExecutionPolicy Bypass` can't override it;
  the script needs signing with a certificate the PCs trust.
- **AppLocker or WDAC.** These force PowerShell into Constrained Language mode. The script
  stops with a message saying so; IT needs to allow or sign it.
- **Security software.** A script that reads performance data and event logs can be flagged
  by an EDR product. If so, allow it in that product, by path or by signature.

## Log size and history

- A new file starts each day, and when a file reaches `-MaxFileSizeMB` (50 MB).
- After midnight, the finished day's CSV files are zipped into
  `DesktopMonitor_<PC>_<date>.zip`, about a fifth of the size. `-NoZip` keeps them as CSV.
  To open a zipped day in Excel, unzip it first; the report reads zips directly.
- Days older than `-RetentionDays` (30) are deleted.
- If this PC's logs take more than `-MaxTotalSizeMB` (500 MB), zips included, the oldest
  days are deleted first. Today's files never are. This is checked every hour.
- A file open in Excel can't be zipped yet; it's tried again every hour.
- Several PCs can log to one folder: each tidies only its own files.

At the default settings a day takes about 20–25 MB, or 4–5 MB zipped.

## Impact on the PC

Measured on an i7-8700 (6 cores, 12 threads):

| | |
|---|---|
| CPU | About 50 ms per sample, about 140 ms for the once-a-minute work. That averaged **0.15% of total CPU** at 10 seconds. |
| Memory | About 100 MB. |
| Disk | One small append per sample, written straight to disk so it survives a crash or power cut. A 64 KB test file written and deleted in the user's temp folder each sample. About 20–25 MB of logs a day. |
| Network | Each sample: one ping per target, one DNS query, one connection to the domain controller (domain PCs only), one per `-ServerTarget`. Once a minute: one small web request, and one connection per mapped drive's server. |

It stays light because:

- it runs at below-normal priority;
- performance counters are opened once and re-read;
- no WMI is used while sampling;
- slow checks run on background threads or once a minute;
- the inventory runs in a separate short-lived process.

If a sample ever takes over 2 seconds, the monitor's `.log` says which part was slow.

## What it collects (privacy)

Everything stays on the PC, in the log folder, unless you copy it elsewhere.
`C:\ProgramData` can be read by every user of the PC.

**It records:**

- **Performance figures:** processor, memory, disk, network, GPU, battery and temperatures.
- **Processes:** names, IDs, resource use, and the names of the services they run. It does
  *not* record command lines or window titles.
- **Accounts:** the names of the account running the monitor, the user at the screen, and
  whoever reports a freeze, plus their note.
- **Event log entries:** the ones listed under [Events file](#events-file), with their
  message text. That text can contain user, computer and file names.
- **Network setup:** IP addresses, DNS servers, gateway, proxy settings, mapped drive paths,
  domain controller and site names.
- **Software:** hardware and software inventory, including security products and drivers.
- **Its own checks:** the times and errors of the connection checks. The only check that
  reaches outside the organisation is the web check (`-HttpTarget`, once a minute), which
  can be pointed at an intranet page or turned off with `-HttpTarget none`.

**It does not record:** keystrokes, screen contents, file contents or names (except where an
event log message contains them), browsing history, or email.

## All parameters

**Monitor (`DesktopMonitor.ps1`)**

| Parameter | Default | What it does |
|---|---|---|
| `-IntervalSeconds` | 10 | Seconds between samples (1–3600). |
| `-LogDirectory` | `C:\ProgramData\DesktopMonitor\Logs` | Where the logs go. |
| `-PingTarget` | `gateway` | One or two hosts to ping. `gateway` is the router. |
| `-PingTimeoutMs` | 1000 | Ping timeout; at most half the interval. |
| `-DnsName` | DC records, or `www.msftconnecttest.com` | The name looked up to time DNS. |
| `-ServerTarget` | none | One or two `host:port` servers to connect to each sample. |
| `-HttpTarget` | Windows' connectivity test page | A page fetched once a minute through the proxy; `none` turns it off. |
| `-AgentName` | none | Extra process names (wildcards allowed) to count as security or monitoring agents. |
| `-TopProcessCount` | 10 | Processes in each top list; 0 turns process tracking off. |
| `-Exclude` | none | Groups to switch off: `Cpu`, `Memory`, `Disk`, `Network`, `Wifi`, `Ping`, `Connectivity`, `Gpu`, `Thermal`, `Power`, `Session`, `Responsiveness`, `Security`, `Processes`, `Events`. Their columns stay, blank. |
| `-RetentionDays` | 30 | Days of logs kept. |
| `-MaxFileSizeMB` | 50 | Size at which a log carries on in a new file. |
| `-MaxTotalSizeMB` | 500 | The most space this PC's logs may take. |
| `-NoZip` | | Don't zip finished days. |
| `-Format` | `Csv` | `Csv` or `Tsv`. |
| `-EventLookbackHours` | 24 | Hours of event log read on the very first start. Later starts carry on from the last sample. |
| `-Priority` | `BelowNormal` | `Idle`, `BelowNormal` or `Normal`. |
| `-DurationMinutes` | 0 (forever) | Stop after this many minutes. |
| `-Live` | | Print a line per sample to the console. |
| `-Install` | | Install for the current user. See [Installing](#installing-deploying-and-checking). |
| `-InstallDirectory` | `%LOCALAPPDATA%\DesktopMonitor` | Where `-Install` copies the script. |
| `-FreezeShortcut` | | With `-Install`: add the *Report a freeze* icon. |
| `-Uninstall` | | Remove the install. The logs are kept. |
| `-Mark` | | Record a freeze report now. |
| `-Note` | | With `-Mark`: a few words about what happened. |
| `-Popup` | | With `-Mark`: confirm in a small window. |
| `-InventoryOnly` | | Write the inventory file and exit (used internally at start-up). |

**Report (`Get-DesktopMonitorReport.ps1`)**: `-LogDirectory`, `-From`, `-To`,
`-ComputerName`, `-IncidentCsv`. See [The report](#the-report).

`Get-Help .\DesktopMonitor.ps1 -Full` shows the same help in PowerShell.
