# Install-Exchange15.ps1

PowerShell script for fully unattended installation of Microsoft Exchange Server 2016, 2019, and Exchange SE — including prerequisites, Active Directory preparation, and post-configuration.

**Maintainer:** st03ps | **Original author:** Michel de Rooij (michel@eightwone.com) · [eightwone.com](http://eightwone.com)
**Version:** 5.84 (April 2026, last updated 2026-04-22)
**License:** As-Is, without warranty

---

## Supported Versions

| Exchange | Windows Server |
|---|---|
| Exchange 2016 CU23 | Windows Server 2016 |
| Exchange 2019 CU10–CU14 | Windows Server 2019, 2022 |
| Exchange 2019 CU15+ | Windows Server 2025 |
| Exchange SE RTM | Windows Server 2022, 2025 |

---

## Prerequisites

- PowerShell 5.1 or later
- Run as local Administrator (script auto-elevates via UAC if needed)
- Domain membership (except Edge role)
- Schema Admin + Enterprise Admin rights for AD preparation
- Static IP address (or Azure Guest Agent detected)
- Exchange setup files (ISO or extracted folder) accessible

---

## Quick Start

### Interactive mode (recommended)

Start the script without parameters to open the interactive installation menu:

```powershell
.\Install-Exchange15.ps1
```

The menu lets you select the installation mode and toggle all options. Press letter keys to toggle switches instantly; press Enter to start.

If a `config.psd1` file is found in the same folder as the script or `.exe`, you are asked whether to use it — allowing a fully pre-configured start without manual input.

### Command-line / unattended

```powershell
# Install Mailbox role — interactive (Copilot mode)
.\Install-Exchange15.ps1 -SourcePath D:\Exchange

# Fully unattended with Autopilot (automatic reboots through all phases)
.\Install-Exchange15.ps1 -SourcePath D:\Exchange -Autopilot

# Load all parameters from a .psd1 config file (skips interactive menu)
.\Install-Exchange15.ps1 -ConfigFile .\deploy-mbx01.psd1

# Install prerequisites only, skip Exchange setup
.\Install-Exchange15.ps1 -NoSetup -SourcePath D:\Exchange

# Edge Transport role
.\Install-Exchange15.ps1 -InstallEdge -SourcePath D:\Exchange

# Recover a server
.\Install-Exchange15.ps1 -Recover -SourcePath D:\Exchange

# Pre-flight report only (no installation)
.\Install-Exchange15.ps1 -SourcePath D:\Exchange -PreflightOnly

# Swing migration: copy config from source server, import PFX, join DAG
.\Install-Exchange15.ps1 -SourcePath D:\Exchange -Autopilot `
    -CopyServerConfig EX01 -CertificatePath D:\certs\mail.pfx -DAGName DAG01

# Install Recipient Management Tools on an admin workstation
.\Install-Exchange15.ps1 -InstallRecipientManagement -SourcePath D:\Exchange

# Install Exchange Management Tools only (Server OS)
.\Install-Exchange15.ps1 -InstallManagementTools -SourcePath D:\Exchange

# Run all post-install optimizations on an existing Exchange server (no setup required)
.\Install-Exchange15.ps1 -StandaloneOptimize -Namespace mail.contoso.com `
    -CertificatePath C:\certs\mail.pfx -LogRetentionDays 30 `
    -RelaySubnets '10.0.1.0/24' -ExternalRelaySubnets '10.0.2.5'

# Generate a Word document for the full organisation on an existing server (ad-hoc inventory)
.\Install-Exchange15.ps1 -StandaloneDocument -Language DE

# Generate a customer-ready Word document — full org + all servers, sensitive values redacted
.\Install-Exchange15.ps1 -StandaloneDocument -Language EN -CustomerDocument

# Document only org-wide configuration (no per-server hardware queries)
.\Install-Exchange15.ps1 -StandaloneDocument -DocumentScope Org -Language DE

# Document specific servers only in a large farm
.\Install-Exchange15.ps1 -StandaloneDocument -IncludeServers EX01,EX02 -Language DE
```

### Compile to .exe (optional)

Use `Build.ps1` to compile the script into a self-contained Windows executable via PS2Exe:

```powershell
.\Build.ps1
# Output: .\Install-Exchange15.exe (runs elevated, all parameters preserved)
```

---

## Key Parameters

| Parameter | Description |
|---|---|
| `-SourcePath` | Path to Exchange setup files or ISO |
| `-TargetPath` | Target folder for Exchange binaries (default: `C:\Program Files\Microsoft\Exchange Server\V15`) |
| `-InstallPath` | Working directory for state file, logs, and downloaded packages (default: `C:\Install`) |
| `-Autopilot` | Fully unattended mode — reboots and resumes automatically |
| `-Credentials` | Credentials for AutoPilot (prompted interactively if omitted) |
| `-Organization` | Exchange organization name (required for new deployments) |
| `-MDBName` | Name of the first mailbox database |
| `-MDBDBPath` | Path for database files (.edb) |
| `-MDBLogPath` | Path for transaction logs |
| `-ConfigFile` | Load all parameters from a `.psd1` file (skips interactive menu) |
| `-IncludeFixes` | Install latest Exchange Security Update (SU) after setup |
| `-InstallWindowsUpdates` | Install Windows security/critical updates in Phase 1 |
| `-DisableSSL3` | Disable SSL 3.0 (POODLE) |
| `-DisableRC4` | Disable RC4 encryption |
| `-EnableTLS12` | Enforce TLS 1.2 (disables TLS 1.0/1.1) |
| `-EnableTLS13` | Enable TLS 1.3 (WS2022+, Exchange 2019 CU15+ / Exchange SE) |
| `-EnableECC` | Enable ECC certificate support |
| `-EnableAMSI` | Enable AMSI body scanning for OWA, ECP, EWS, PowerShell |
| `-NoSetup` | Install prerequisites only, skip Exchange setup |
| `-Phase` | Start directly at a specific phase (0–6) |
| `-PreflightOnly` | Generate HTML pre-flight report and exit |
| `-CopyServerConfig` | Export config from source server and import after setup |
| `-CertificatePath` | Path to PFX file for certificate import (IIS + SMTP) |
| `-DAGName` | Join a Database Availability Group after setup |
| `-InstallEdge` | Install the Edge Transport role |
| `-Recover` | RecoverServer mode |
| `-InstallRecipientManagement` | Install Recipient Management Tools (Server or Client OS) |
| `-InstallManagementTools` | Install Exchange Management Tools only (Server OS) |
| `-SkipHealthCheck` | Skip CSS-Exchange HealthChecker run at end |
| `-NoCheckpoint` | Skip System Restore checkpoints |
| `-SkipWindowsUpdates` | Skip Windows Update check even when `-InstallWindowsUpdates` is set |
| `-RunEOMT` | Download and run CSS-Exchange Emergency Mitigation Tool (EOMT) in Phase 5 |
| `-WaitForADSync` | After PrepareAD, wait up to 6 minutes for AD replication to be error-free |
| `-LogRetentionDays` | Register a daily scheduled task to delete IIS + Exchange logs older than N days (1–365) |
| `-RelaySubnets` | IP ranges for anonymous internal relay (accepted domains only, no external relay right) |
| `-ExternalRelaySubnets` | IP ranges for anonymous external relay (any recipient); removes AnonymousUsers from Default Frontend on success |
| `-StandaloneOptimize` | Run all post-install tasks on an already-installed Exchange server without running the full install flow |
| `-StandaloneDocument` | Generate a Word installation document on an existing server (loads Exchange module, no install flow) |
| `-SkipInstallReport` | Skip the HTML/PDF installation report generated at Phase 6 completion |
| `-NoWordDoc` | Skip Word (.docx) installation document generation at Phase 6 |
| `-CustomerDocument` | Redact RFC1918 IPs, certificate thumbprints, and passwords in the Word document |
| `-Language` | Document language: `DE` (default) or `EN` |
| `-DocumentScope` | `All` (default): org + all servers + local. `Org`: org-wide only. `Local`: per-server only |
| `-IncludeServers` | Limit per-server documentation to specific server names (e.g. `EX01,EX02`) |
| `-Verbose` | Raise log tier to `VERBOSE` — adds detailed progress / decision entries to the log file |
| `-Debug` | Raise log tier to `DEBUG` — adds `VERBOSE` + `DEBUG` + `SUPPRESSED-ERROR` entries (reconstructs errors silently caught by `try/catch`); invaluable for diagnosing BITS / MSI / CIM failures |

See `deploy-example.psd1` for a fully documented configuration file template.

### Logging

A single log file is written for every run. Entries are tier-prefixed; the active tier is selected by the standard PowerShell switches on the `.ps1` call:

| Switch | Tiers written |
|---|---|
| _(none)_ | `INFO` / `WARNING` / `ERROR` / `EXE` |
| `-Verbose` | + `VERBOSE` |
| `-Debug` | + `DEBUG` + `SUPPRESSED-ERROR` |

`SUPPRESSED-ERROR` lines reconstruct exceptions swallowed by `try/catch` — without them, silent BITS failures, missing registry keys, or CIM timeouts leave no trace. The tier is preserved across Autopilot reboots (re-applied to the resumed process via `RunOnce`).

Log encoding is UTF-8 without BOM — so the file renders correctly in every viewer (PS 5.1 `Out-File` would default to UTF-16 LE, which caused rendering artefacts in earlier versions).

---

## Installation Phases

The script runs through 7 phases (0–6) and saves state in an XML file to
automatically resume after reboots:

```
Phase 0 → Preflight checks, pre-flight HTML report
Phase 1 → Windows features, Windows Updates (optional)
Phase 2 → .NET Framework 4.8/4.8.1, OS hotfixes, Visual C++ Runtimes, URL Rewrite
Phase 3 → UCMA Runtime, Active Directory preparation (PrepareAD/PrepareSchema), AD replication check
Phase 4 → Run Exchange Setup
Phase 5 → Post-configuration (security hardening, performance tuning, certificate import, Exchange SU)
Phase 6 → Start services, IIS health check, Virtual Directory URLs, DAG join, HealthChecker, HTML + PDF Installation Report, Word Installation Document, cleanup
```

Recipient Management / Management Tools modes use phases 0–2 only.

---

## Post-Configuration (Phase 5)

The following best-practice configurations are automatically applied after Exchange setup:

### Security Hardening
- **TLS 1.0 / 1.1 disabled, TLS 1.2 enforced** — SChannel + .NET Framework strong crypto; TLS 1.3 on WS2022+ with Exchange 2019 CU15+ / SE ([Exchange TLS Guide](https://techcommunity.microsoft.com/t5/exchange-team-blog/exchange-server-tls-guidance-part-1-getting-ready-for-tls-1-2/ba-p/607649))
- **SMBv1 disabled** — removes legacy protocol attack vector ([Microsoft Blog](https://techcommunity.microsoft.com/t5/storage-at-microsoft/stop-using-smb1/ba-p/425858))
- **WDigest credential caching disabled** — prevents cleartext password storage in LSASS ([MS Learn](https://learn.microsoft.com/en-us/windows-server/security/credentials-protection-and-management/configuring-additional-lsa-protection))
- **LAN Manager level 5** — NTLMv2 only, refuse LM and NTLM ([MS Learn](https://learn.microsoft.com/en-us/windows/security/threat-protection/security-policy-settings/network-security-lan-manager-authentication-level))
- **LSA Protection (RunAsPPL)** — enabled for Exchange 2019 CU12+ / Exchange SE; prevents credential theft from LSASS ([MS Learn](https://learn.microsoft.com/en-us/windows-server/security/credentials-protection-and-management/configuring-additional-lsa-protection))
- **Serialized Data Signing** — mitigates PowerShell deserialization attacks ([Exchange Blog](https://techcommunity.microsoft.com/blog/exchange/released-2022-h1-cumulative-updates-for-exchange-server/3285026))
- **Credential Guard disabled** — causes performance issues on Exchange; default-on in WS2025 ([Exchange Virtualization](https://learn.microsoft.com/en-us/exchange/plan-and-deploy/virtualization))
- **IPv4 over IPv6 preference** — `DisabledComponents = 0x20`; prefers IPv4, keeps IPv6 loopback (required by Exchange internal components) ([Exchange Blog](https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/configure-ipv6-in-windows))
- **NetBIOS over TCP/IP disabled** on all NICs — reduces LLMNR/NBT-NS attack surface; Exchange does not require NetBIOS ([MS Learn](https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/disable-netbios-tcp-ip-using-dhcp))
- **HTTP/2 disabled** — prevents known Exchange MAPI/RPC issues ([Exchange Blog](https://techcommunity.microsoft.com/blog/exchange/released-2022-h1-cumulative-updates-for-exchange-server/3285026))
- **Extended Protection** — CU14+/SE: validated via OWA VDir; pre-CU14: `ExchangeExtendedProtection.ps1` (CSS-Exchange) ([MS Learn](https://learn.microsoft.com/en-us/exchange/plan-and-deploy/post-installation-tasks/security-best-practices/exchange-extended-protection))
- **SSL offloading disabled** (Outlook Anywhere) — required for Extended Protection channel binding ([MS Learn](https://learn.microsoft.com/en-us/exchange/plan-and-deploy/post-installation-tasks/security-best-practices/exchange-extended-protection))
- **MRS Proxy disabled** (EWS VDir) — re-enable manually only for cross-forest migrations ([MS Learn](https://learn.microsoft.com/en-us/exchange/architecture/mailbox-servers/mrs-proxy-endpoint))
- **MAPI encryption required** (Mailbox role) — forces encrypted Outlook MAPI connections ([MS Learn](https://learn.microsoft.com/en-us/exchange/clients/mapi-over-http/configure-mapi-over-http))
- **Root certificate auto-update enabled** — required for Exchange Online / M365 hybrid connectivity; re-enables if disabled by policy ([MS Trusted Root](https://learn.microsoft.com/en-us/security/trusted-root/release-notes))
- **Windows Defender exclusions** — folder, process, and extension exclusions per Microsoft guidance
- **HSTS header** — `Strict-Transport-Security: max-age=31536000` on OWA and ECP; conditional on `-CertificatePath`
- **AMSI body scanning** — OWA, ECP, EWS, PowerShell (Exchange 2016/2019; Exchange SE: default-on since Aug 2025 SU)
- **Windows services disabled** (CIS / NSA / DISA STIG) — the following services are stopped and set to `Disabled`; none are required on Exchange and each increases the attack surface or interferes with automation:
  - `Spooler` (Print Spooler) — PrintNightmare CVE-2021-34527; no printing on mail servers
  - `Fax` — no use case on Exchange
  - `seclogon` (Secondary Logon) — pass-the-hash / privilege escalation vector (`runas`)
  - `SCardSvr` (Smart Card) — no smart card hardware on servers
  - `WSearch` (Windows Search) — Exchange uses its own content indexing engine
- **Shutdown Event Tracker disabled** — registry (`ShutdownReasonOn/UI = 0`); redundant with Event IDs 1074/6006/6008; dialog blocks unattended Autopilot reboots

### Performance Tuning
- **High Performance power plan** — activated automatically; Exchange must not run on Balanced
- **NIC power management disabled** — prevents adapter sleep / power state changes
- **Pagefile configured** — 25% RAM (Exchange 2019+) or RAM+10 MB (Exchange 2016), minimum 32 GB
- **TCP settings** — RPC minimum connection timeout 120 s; Keep-Alive 15 min
- **TCP Chimney and Task Offload disabled** — Microsoft recommendation for Exchange servers
- **HTTP/2 disabled** — see Security Hardening above
- **Windows Search service disabled** — see *Windows services disabled* in Security Hardening above
- **RSS enabled on all NICs** — ensures network traffic uses all CPU cores; receive queue count = physical core count
- **MaxConcurrentAPI configured** — Netlogon set to logical core count (min 10) to prevent 0xC000005E Kerberos errors
- **CtsProcessorAffinityPercentage = 0** — Exchange Search best practice (no CPU affinity limit)
- **NodeRunner memory limit = 0** — removes Exchange Search performance limiter
- **MAPI Front End Server GC** — enabled on systems with 20+ GB RAM
- **CRL check timeout configured** — 15 seconds; prevents Exchange startup delays on networks with slow CRL endpoints
- **Scheduled defragmentation disabled** — not needed on Exchange servers
- **Disk allocation unit size verification** — warns if volumes are not 64 KB formatted

### Exchange-Level Optimizations (interactive menu, all enabled by default)
- **Modern Authentication (OAuth2)** — required for Outlook 2016+, Teams, mobile clients, Hybrid
- **OWA/ECP session timeout** — 6 hours inactivity auto-logout (security compliance)
- **CEIP / telemetry disabled** — `CustomerFeedbackEnabled $false` (GDPR / privacy)
- **MAPI over HTTP** — explicit enablement; replaces RPC/HTTP; better failover and NAT behavior
- **Max message size: 150 MB** — org-wide send/receive limit; Frontend Receive Connectors updated consistently
- **Message expiration: 7 days** — delays NDR generation during multi-day outages
- **SMTP banner hardened** — generic `220 Mail Service` banner hides Exchange version from attackers (CIS / DISA STIG)
- **HTML Non-Delivery Reports** — improves end-user NDR readability
- **Shadow Redundancy** — prefer remote DAG member for in-flight message redundancy (DAG only)
- **Safety Net hold time: 2 days** — explicit redelivery hold time for post-failover message recovery

---

## What's New

### v5.84 — April 2026
- **Word Document: organisation-wide + all servers** — `New-InstallationDocument` now documents the entire Exchange organisation, not only the local server; new chapter 4 covers org-wide configuration (Org-Config, Accepted/Remote Domains, E-Mail-Address Policies, Transport Rules, Transport Config, Journal/DLP/Retention, Mobile/OWA Policies, all DAGs with DB copies, org-scoped Send Connectors, Federation/Hybrid/OAuth, AuthConfig); new chapter 5 enumerates every Exchange server with identity, hardware details, databases, virtual directories, receive connectors, certificates, and transport agents
- **Ad-hoc inventory mode** — run `-StandaloneDocument` on any existing Exchange server at any time to produce a full current-state document of the organisation, without a prior EXpress install
- **Remote hardware queries via CIM/WSMan** — hardware, pagefile, volume, and NIC data collected from all servers using WinRM (TCP 5985/5986, Kerberos); no WMI/DCOM, no dynamic RPC ports; Exchange EMS already requires WinRM so no extra infrastructure needed
- **Interactive retry/skip prompt** — if a remote server is unreachable, EXpress shows `[R] Retry / [S] Skip` with a 10-minute auto-skip countdown; silent skip in Autopilot/unattended mode; failure is noted inline in the document
- **`-DocumentScope All|Org|Local`** — control document depth: `All` = full org + all servers (default), `Org` = org-wide chapter only, `Local` = per-server sections only; useful for large farms or targeted documentation runs
- **`-IncludeServers`** — limit per-server documentation to specific Exchange servers in large environments
- **Pre-requisite tooling** — `tools/Enable-EXpressRemoteQuery.ps1` enables WinRM on target servers with one command; `docs/remote-query-setup.md` provides GPO equivalent, firewall matrix, and error guide
- **Titelblatt Szenario-Zeile** — document title page shows whether the run was a new environment, a server addition, or an ad-hoc inventory

### v5.83 — April 2026
- **Three-tier logging** — single log file with tier-prefixed entries (`INFO` / `WARNING` / `ERROR` / `EXE` / `VERBOSE` / `DEBUG`); activated via standard `-Verbose` / `-Debug` switches on the `.ps1` call. Debug tier additionally emits `[SUPPRESSED-ERROR]` lines that reconstruct errors silently swallowed by `try/catch` — invaluable for diagnosing BITS/MSI/CIM failures. Encoding pinned to UTF-8 without BOM so the log renders correctly in every viewer. Log tier survives Autopilot reboots.
- **Unified file nomenclature** — all generated artefacts (log, preflight, installation report, Word document, RBAC, exported config, log-cleanup) now follow `{PC}_{Tag}_{yyyyMMdd-HHmmss}.{ext}`; Preflight gains a timestamp so prior runs are preserved instead of overwritten
- **Credential prompt fix** — `Get-ValidatedCredentials` now picks GUI vs. Read-Host deterministically via `$env:SESSIONNAME` (console vs. RDP); eliminates silent `$null` fallback when `Get-Credential` is cancelled
- **Bootstrap order** — log initialisation runs before the menu; `$script:isFreshStart` snapshot prevents early state mutations from triggering a false Autopilot-resume path
- **Dev tools** — `Test-ScriptSanity` (14 structural checks), `Test-ScriptQuality`, `Fix-IfAsArg`, `Fix-PhaseNum`, `Parse-Check`

### v5.82 — April 2026
- **Word Installation Document** (`New-InstallationDocument`) — generated automatically at Phase 6 completion alongside the HTML report; pure-PowerShell OpenXML engine, no Office dependency; 15 chapters: installation parameters, system details, network & DNS, Active Directory, Exchange configuration (DBs, VDirs, connectors, certificates, DAG, transport), hardening & tuning, backup readiness, HealthChecker reference, monitoring, hybrid status, public folders, executed cmdlets, runbooks and open items
- **`-CustomerDocument`** — redacts RFC1918 IP addresses, certificate thumbprints, and passwords for sharing with external parties
- **`-Language`** — selects document language: `DE` (default) or `EN`
- **`-NoWordDoc`** — skip Word document generation when not needed
- **`-StandaloneDocument`** (menu mode 7) — generates the Word document on an existing Exchange server without running the install flow; just loads the Exchange module and writes the document
- **Konzept-/Freigabedokument templates** (`tools/Build-KonzeptTemplate.ps1`) — generates two static Word templates (`templates/Exchange-Konzept-Vorlage-DE.docx` + `-EN.docx`) covering 16 chapters: architecture, sizing, security, message hygiene, backup/DR, monitoring, migration, hybrid, public folders, compliance, mobile, questionnaire, and approval page; Exchange SE only (Exchange 2016/2019 are out-of-support since 14 October 2025)

### v5.81 — April 2026
- **Installation Report FormatException (complete fix)** — root cause was `New-HtmlSection`, `Format-Badge`, `Format-RefLink`, and all section assembly lines using `-f` with dynamic Exchange data; any user-defined value containing `{n}` (connector name, cert SAN, policy value) caused `String.Format` to throw; all formatting converted to string concatenation

### v5.80 — April 2026
- **Installation Report FormatException** — `$exContent` HERE-STRING with `-f` partially fixed; `New-HtmlSection` root cause not yet addressed
- **HealthChecker report name** — HC output now saved as `SERVER_HCExchangeServerReport-<timestamp>.html`

### v5.79 — April 2026
- **Installation Report crash** — transcript read with wrong encoding (UTF-8 instead of UTF-16 LE); log section now auto-detects encoding from BOM, is wrapped in try/catch, and capped at last 2 000 lines; report generation wrapped in try/catch so a crash no longer kills the entire script

### v5.78 — April 2026
- **Exchange SU reboot loop** — Exchange SU installer (`.exe`) may internally call `ExitWindowsEx` and reboot the machine before the script's phase-end logic runs; in Autopilot mode, `RunOnce` + state are now persisted **before** launching the installer so the script always auto-resumes; a per-KB flag in state prevents the SU from being reinstalled when phase 5 re-runs after the reboot

### v5.77 — April 2026
- **Exchange SU installer** — removed `/norestart` from arguments; Exchange SU `.exe` only accepts `/passive` and `/silent` — `/norestart` caused the installer to abort with "command line option not recognized"

### v5.76 — April 2026
- **Auth Certificate check** — `Test-AuthCertificate` no longer throws "null-valued expression" when `Get-AuthConfig` returns `$null` after an IIS restart; null-guard added before property access
- **External relay connector** — fixed race condition where `Add-ADPermission` failed immediately after `New-ReceiveConnector` because Exchange had not yet registered the object in AD; connector object now taken directly from the `New-ReceiveConnector` return value; 3-attempt/5 s retry fallback for edge cases

### v5.75 — April 2026
- **AD preparation** — `Initialize-Exchange` returns `$true`/`$false`; `Wait-ADReplication` only called when PrepareAD actually ran; progress label reflects conditional check
- **Edge Transport guards** — `Enable-AMSI`, `Set-MaxConcurrentAPI`, `Set-CtsProcessorAffinityPercentage`, `Set-NodeRunnerMemoryLimit` now skip silently for Edge role

### v5.74 — April 2026
- **AMSI body scanning** — Exchange SE exception removed; HealthChecker always checks for the `AmsiRequestBodyScanning` SettingOverride, so the override is now applied for all Exchange versions when `-EnableAMSI` is used
- **HealthChecker server membership note** — post-run note explains why "Exchange Server Membership" may show blank/failed in the same-session run (Kerberos token refresh requires reboot)

### v5.73 — April 2026
- **Antispam agent warnings** — `$PSDefaultParameterValues['*:WarningAction'] = 'Ignore'` applied before `Install-AntispamAgents.ps1` to defeat internal `$WarningPreference` resets; `Enable/Disable-TransportAgent` calls switched from `$null =` (Stream 1 only) to `*>&1 | Out-Null` with `-WarningAction Ignore`

### v5.72 — April 2026
- **HealthChecker HTML report** — `Invoke-HealthChecker` now calls `-BuildHtmlServersReport` after data collection so `ExchangeAllServersReport-*.html` is generated in `ReportsPath`
- **Installation report** — HealthChecker section re-added (Section 9) with iframe embed, direct link, and "skipped/not found" fallback messages; TOC entry added

### v5.71 — April 2026
- **SU install countdown** — checks `ConfigDriven` flag instead of `Autopilot`; countdown was incorrectly skipped in interactive (Copilot) sessions with auto-reboot enabled

### v5.70 — April 2026
- **Link fixes** — 6 broken/wrong-content URLs corrected in README and installation report: Extended Protection (404), SSL Offloading (404), 2022 H1 CU blog (wrong ba-p), TLS 1.2 Part 2 (wrong ba-p), TLS 1.3 (wrong ba-p), IPv6 (wrong ba-p); `docs.microsoft.com` → `learn.microsoft.com`

### v5.69 — April 2026
- **Mode label fixed** — `Mode: Copilot (interactive)` now shown correctly when starting via the interactive menu, even if the auto-reboot toggle is on; `Mode: Autopilot (fully automated)` is reserved for config-file (`-ConfigFile`) starts

### v5.68 — April 2026
- **Unnecessary services disabled** (`Disable-UnnecessaryServices`) — Print Spooler (PrintNightmare, CVE-2021-34527), Fax, Secondary Logon (pass-the-hash vector), Smart Card; per CIS/NSA/DISA STIG recommendations
- **Shutdown Event Tracker disabled** (`Disable-ShutdownEventTracker`) — redundant with Windows Event IDs 1074/6006/6008; dialog blocks unattended Autopilot reboots

### v5.66 — April 2026
- **IPv4 over IPv6 preference** (`Set-IPv4OverIPv6Preference`) — `DisabledComponents = 0x20`; Microsoft recommended setting for Exchange; keeps IPv6 loopback intact; flags `RebootRequired`
- **NetBIOS disabled on all NICs** (`Disable-NetBIOSOnAllNICs`) — `SetTcpipNetbios(2)` via WMI; Exchange does not require NetBIOS; reduces LLMNR/NBT-NS attack surface

### v5.65 — April 2026
- **SU download hint URL fixed** — `aka.ms/ExchangeSU` was a dead link (resolved to Bing); replaced with the KB-specific `https://support.microsoft.com/help/<number>` URL derived from the KB field; applied to both the manual placement hint and the post-install failure warning

### v5.64 — April 2026
- **Log cleanup path coverage expanded** — `Invoke-ExchangeLogCleanup.ps1` now cleans `V15\Logging\` and `V15\TransportRoles\Logs\` recursively instead of 6 specific sub-paths; covers EWS, OWA, HttpProxy, ECP, RpcClientAccess, MessageTracking, DSN, Connectivity, etc.
- **HTTPERR logs added** — `%SystemRoot%\System32\LogFiles\HTTPERR` cleaned with the same retention period
- **IIS log path dynamic** — resolved from IIS metabase via `Get-WebConfigurationProperty`; fallback to `inetpub\logs\LogFiles`

### v5.63 — April 2026 — Bugfixes
- **Antispam agent warnings** — `3>$null` replaced by `*>&1 | Out-Null`; implicit-remoting warnings from `Enable-TransportAgent` (PS 5.1) now fully suppressed
- **Exchange SE RTM SU (KB5074992)** — removed duplicate hardcoded install attempt; WU-catalog CAB URL cleared (not DISM-compatible); 5-minute interactive countdown prompt when EXE not found; download URL shown as `https://support.microsoft.com/help/5074992`

### v5.62 — April 2026
- **F13** `Disable-SSLOffloading` — prerequisite for Extended Protection channel binding
- **F6** `Enable-ExtendedProtection` — CU14+/SE: native validation via OWA VDir; pre-CU14: downloads and runs `ExchangeExtendedProtection.ps1` (CSS-Exchange)
- **F17** `Enable-RootCertificateAutoUpdate` — re-enables root cert auto-update when disabled by policy
- **F18** `Disable-MRSProxy` — `MRSProxyEnabled $false` on EWS VDir; Mailbox role only
- **F19** `Set-MAPIEncryptionRequired` — forces encrypted MAPI connections; Mailbox role only
- **F8** `Test-DAGReplicationHealth` — checks database copy status after DAG join; warns on non-Mounted/Healthy
- **F9** `Test-VSSWriters` — validates VSS writer state; broken VSS → failed Exchange backups
- **F10** `Test-EEMSStatus` — checks MSExchangeMitigation service and org config (CU11+ / SE)
- **F11** `Get-ModernAuthReport` — warns when OAuth2 is disabled in org config
- **P7** Compliance mapping — CIS / BSI control-ID column added to Installation Report Security section

### v5.61 — April 2026 (2026-04-20) — Bugfixes

- **Virtual directory `-Force` removed** — `@forceParam` removed from all six `Set-*VirtualDirectory` calls; OWA's `-Force` was ambiguous (matched `ForceSave*`/`ForceWac*` parameters) causing `ParameterBindingException` in Autopilot mode; all cmdlets now use `-Confirm:$false` only
- **External relay placeholder text** — warning message corrected to show `192.0.2.2/32` (was incorrectly showing `192.0.2.1/32`, the internal connector's address)
- **`Add-ADPermission` warning suppressed** — "access control entry already present" warning no longer shown when re-running Phase 6 on an existing external relay connector
- **Log cleanup prompt hang fixed** — when RawUI is unavailable, default folder is now accepted silently instead of blocking on `Read-Host`; prevented indefinite hang after 2-minute countdown
- **Countdown progress bars** — all timed prompts now show a `Write-Progress` countdown bar (Id 2): log cleanup folder (2 min), Windows Update per-update prompt (2 min), Autopilot resume (10 s), reboot countdown (10 s)
- **Installation report `if` syntax** — `-f (if ...)` caused `CommandNotFoundException` for `if` in PS 5.1; replaced with intermediate variable
- **HealthChecker report detection** — updated to match current HC output filename pattern (`ExchangeAllServersReport-*.html`); detected path stored in `$State['HCReportPath']` for reliable embedding in the Installation Report

### v5.6 — April 2026 (latest: 2026-04-19)

- **Reports subfolder** — all reports and logs are now written to `<InstallPath>\reports\` (default `C:\Install\reports\`); folder is created automatically on first run and on AutoPilot resume
- **RBAC Role Group Membership** in Installation Report — 10 standard Exchange role groups queried live via `Get-RoleGroupMember`; members shown with `RecipientType`; new section 8 in HTML report
- **Installation Log** in Installation Report — full transcript embedded as scrollable dark block; new section 9 in HTML report
- **Autodiscover SCP** moved into Virtual Directory URLs table — first row, queried via `Get-ClientAccessService`
- **UAC re-enabled before report** — `Enable-UAC` and `Enable-IEESC` now run before `New-InstallationReport`; report correctly shows UAC as Enabled
- **HealthChecker section** distinguishes `-SkipHealthCheck` (intentionally skipped with actionable hint) from HC failure (directs to installation log)
- **Relay connectors** — menu now always creates both internal **and** external relay connectors when `[Y]` selected; blank subnet entry uses RFC 5737 placeholder `192.0.2.1/32` (never routable); Default Frontend hardening skipped when only placeholders set — no risk of anonymous lockout
- **Log cleanup prompt** — interactive folder prompt suppressed in AutoPilot mode (uses default `C:\#service` silently)

### v5.51 — April 2026 (2026-04-19) — Bugfixes

- **Credential prompt** — `Get-Credential` PSObject cast fixed; `Read-Host` fallback added for PS2Exe/compiled-exe and console environments where `Get-Credential` returns `$null`
- **`Start-DisableMSExchangeAutodiscoverAppPoolJob`** — `Test-Path 'IIS:\AppPools\...'` replaces `Get-WebAppPoolState` (PathNotFound not suppressed by `-ErrorAction SilentlyContinue`)
- **Service restart warnings** — `-WarningAction SilentlyContinue` on all `Restart-Service` calls for W3SVC/WAS and MSExchangeTransport; antispam install script output redirected via `3>$null`; `Enable/Disable-TransportAgent` warnings suppressed
- **Virtual directory URL confirmation** — `-Confirm:$false` on all `Set-*VirtualDirectory` calls; suppresses "host can't be resolved, continue?" prompt during AutoPilot
- **Log cleanup input loop** — `FlushInputBuffer` moved to its own `try/catch`; a failure no longer aborts the 2-minute RawUI loop and jumps immediately to `Read-Host`
- **VC++ 2012 (v11.0)** — install condition extended to all Exchange versions (was Exchange 2016/Edge only); HealthChecker flags v11.0 as required for Exchange 2019/SE
- **KB5074992 download URL** — Windows Update Catalog CAB URL added; SU is now downloaded automatically without manual file placement
- **`Reconnect-ExchangeSession`** — new helper reconnects the Exchange implicit-remoting PS session after IIS restarts from `Enable-ECC`/`Enable-CBC`/`Enable-AMSI`; waits up to 90 s for the Exchange PS endpoint; called automatically before `Invoke-ExchangeOptimizations`

### v5.5 — April 2026 (latest: 2026-04-19)

- **Anti-spam agents** (`Install-AntispamAgents`) — runs `Install-AntispamAgents.ps1`, restarts transport, disables all agents except RecipientFilter Agent (Phase 6)
- **Send Connector integration** (`Add-ServerToSendConnectors`) — interactive prompt `[Y/N]` to add the new server to existing Send Connectors (Phase 6)
- **Log cleanup script** — interactive prompt for script folder (default `C:\#service`, 2-min timeout); generated script logs to `logs\` subfolder and cleans its own logs older than 30 days; covers IIS, Exchange transport, and Monitoring logs
- **Relay connector hardening** — relay connectors now use `-AuthMechanism Tls` (STARTTLS offered) and `-ProtocolLoggingLevel Verbose`
- **Certificate wildcard detection** — `Import-ExchangeCertificateFromPFX` detects wildcard vs. non-wildcard certs; non-wildcard certs additionally enable IMAP and POP services
- **Bugfix** — ISO was only remounted for phases 1–3 and dismounted at end of Phase 4; path is now consistent
- **Bugfix** — `Test-Preflight`: heavy checks (setup path, AD, FFL, roles) now correctly skipped for Phase ≥ 5
- **Bugfix** — `Set-VirtualDirectoryURLs`: MAPI `-InternalAuthenticationMethods` wrapped in separate `try/catch` (not all builds support it); OWA now sets `-LogonFormat UPN`
- **Bugfix** — `Get-RBACReport`: format string crash fixed (catch block uses string interpolation instead of `-f` operator)
- **Bugfix** — `Import-ExchangeModule`: no longer emits WARNING when module is already loaded; uses `Get-ExchangeServer` instead of `Connect-ExchangeServer`; Phase 6 loads module only once

### v5.4 — April 2026 (2026-04-18)

- **Installation Report** (`New-InstallationReport`) — comprehensive HTML report generated automatically at the end of Phase 6; 6 sections: Installation Parameters, System Information, Active Directory, Exchange Configuration (virtual directory URLs, mailbox databases, receive connectors, certificates), Security Settings, Performance & Tuning; status badges (green/orange/red) for every setting; sidebar navigation; print-friendly CSS
- **PDF export** — automatic via Microsoft Edge headless (`--print-to-pdf`) when Edge is installed on the server; graceful fallback message if not found
- **`-SkipInstallReport`** — switch to suppress report generation when not needed
- **Verbose logging** — verbose messages are always written to the log file; console output remains clean (`$VerbosePreference = SilentlyContinue`)

### v5.3 — April 2026

Code quality and robustness improvements; no new parameters.

- **`Add-BackgroundJob` helper** — prunes `Completed`/`Failed`/`Stopped` jobs from `$Global:BackgroundJobs` before each append; prevents unbounded list growth across phases
- **`New-LDAPSearch` helper** — encapsulates `DirectorySearcher` creation (SearchRoot + Filter); eliminates duplicated 3-line blocks in four functions (`Clear-AutodiscoverSCP`, `Set-AutodiscoverSCP`, `Test-ExistingExchangeServer`, `Get-ExchangeServerObjects`)
- **Registry idempotency** — `RunOnce`, `Disable/Enable-UAC`, `Enable-AutoLogon`, `Disable-OpenFileSecurityWarning`, `Set-NETFrameworkInstallBlock`, and `Disable-ServerManagerAtLogon` now all use `Set-RegistryValue`; write is skipped when the value is already correct
- **BSTR memory zeroing** — `ZeroFreeBSTR` called immediately after `PtrToStringAuto` in `Test-Credentials` and `Enable-AutoLogon`; eliminates plaintext password residue in process memory
- **Exit code checks** — `RUNDLL32.EXE` (`Clear-DesktopBackground`) and `powercfg.exe` (`Enable-HighPerformancePowerPlan`) now emit `Write-MyWarning` on non-zero exit codes
- **Pester tests** extended from 45 to 54 tests: `Set-RegistryValue` idempotency (5 cases) + `Add-BackgroundJob` pruning (4 cases); assertion fix for Exchange 2016 CU23 label

### v5.2 — April 2026

- **HSTS header** (`Set-HSTSHeader`) — configures `Strict-Transport-Security: max-age=31536000` on OWA and ECP; conditional on `-CertificatePath` to avoid browser lockout with self-signed certificates (Phase 5)
- **Emergency Mitigation Tool** (`-RunEOMT`) — downloads and runs CSS-Exchange EOMT for CVE mitigation; BITS download with PS 5.1 fallback; SHA256 logged (Phase 5)
- **AD replication check** (`-WaitForADSync`) — polls `repadmin /showrepl /errorsonly` after PrepareAD until error-free or 6-minute timeout (Phase 3)
- **Database/log path check** (`Test-DBLogPathSeparation`) — warns when DB and log share the same volume root; DAG-aware size guidance (Phase 6)
- **Log cleanup task** (`-LogRetentionDays`) — registers `\Exchange\Exchange Log Cleanup` scheduled task (daily 02:00, SYSTEM, 1h limit) for IIS, transport, and tracking logs (Phase 6)
- **Anonymous relay connectors** (`-RelaySubnets`, `-ExternalRelaySubnets`) — two-connector design: internal relay (no external relay right) and external relay (`Ms-Exch-SMTP-Accept-Any-Recipient`); `AnonymousUsers` removed from Default Frontend on success; account name resolved via SID S-1-5-7 (language-independent, works on DE/EN/FR/etc.) (Phase 6)
- **Standalone optimize mode** (`-StandaloneOptimize`) — single-phase run of all post-install tasks on an existing Exchange server; no setup flow required; shares `-Namespace`, `-CertificatePath`, `-DAGName`, `-SkipHealthCheck`, `-RelaySubnets`, `-ExternalRelaySubnets`, `-LogRetentionDays`
- **Pre-flight report** — added Exchange database sizing best-practices section (DAG vs. standalone limits, log/DB separation, 64 KB allocation unit, free space guidance)
- **Pester tests** extended — `Get-FullDomainAccount` edge cases, DB/log separation logic, HSTS header value validation (no `includeSubDomains`, min 1-year `max-age`)

**Bugfixes and quality improvements (2026-04-17):**
- **ValidatePattern regex** — removed inline `(?# ...)` comment from `-Organization` pattern that caused a parse error (`Too many )'s`) on script load
- **Windows Update count** — installed count now restricted to approved KBs; PSWindowsUpdate previously included already-installed updates in the `Installed` result set
- **`Disable-IEESC` and `Disable-ServerManagerAtLogon`** — moved from AutoPilot reboot preparation to Phase 1 (called once; registry changes persist across reboots)
- **Dead code removed** — `DisableSharedCacheServiceProbe` function (defined but never called)
- **Named constants** — `$ERR_SUS_NOT_APPLICABLE` and `$POWERPLAN_HIGH_PERFORMANCE` replace hardcoded magic values

### v5.1 — April 2026

- **Interactive installation menu** — start without parameters; numbered mode selection + letter toggles; RawUI.ReadKey for instant response (falls back to Read-Host for RDP/PS2Exe compatibility)
- **Credential validation loop** — `Get-ValidatedCredentials` with up to 3 retries (R=Retry / Q=Quit)
- **Recipient Management Tools** (`-InstallRecipientManagement`) — 3-phase install for dedicated admin workstations; works on Server and Client OS; includes RSAT-ADDS prereqs, Add-PermissionForEMT.ps1, and desktop shortcut
- **Exchange Management Tools** (`-InstallManagementTools`) — 3-phase install of `setup.exe /roles:ManagementTools` on Server OS
- **Config file support** (`-ConfigFile`) — load all parameters from a `.psd1` (see `deploy-example.psd1`); `config.psd1` in script folder is auto-detected on interactive start
- **Windows Update integration** (`-InstallWindowsUpdates`) — per-update prompt `[Y/N/A/S]` in interactive mode; download+install in background job with 60-minute timeout (Exchange install continues on timeout); WUA COM fallback
- **Exchange Security Update automation** — built-in `$ExchangeSUMap` with direct download URLs for all supported CUs; installed in Phase 5 when `-IncludeFixes` is set
- **Build.ps1** — compile script to self-contained `.exe` via PS2Exe (`-requireAdmin`, preserves all parameters, embeds version metadata); progress output visible via Write-MyOutput fallback
- **Write-Progress indicators** — overall phase progress (Id 0) and Phase 5 step counter (Id 1, ~25 steps); PS2Exe fallback via plain output
- **LSA Protection** — `Enable-LSAProtection` sets `RunAsPPL=1`; Exchange 2019 CU12+ / Exchange SE; takes effect after reboot
- **MaxConcurrentAPI** — `Set-MaxConcurrentAPI` configures Netlogon to prevent 0xC000005E under Exchange auth load
- **Performance fixes** — `Clear-DesktopBackground` no longer uses `Add-Type`/C# (3–10s faster per phase start); `Get-DetectedFileVersion` uses `FileVersionInfo` API instead of `Get-Command`
- **Bugfix** — `Zone.Identifier` ADS check skipped for mounted ISO sources (UDF/ISO9660 has no ADS support)
- **Bugfix** — Server Manager no longer appears on intermediate AutoPilot reboots (`Disable-ServerManagerAtLogon` moved to AutoPilot preparation block before every reboot; later refined to Phase 1 in v5.2)

### v5.01 — April 2026

- **Auto-elevation** — script re-launches elevated via UAC if not running as Administrator
- **Auto-unblock** — detects and removes `Zone.Identifier` on Exchange setup source files (prevents .NET sandbox errors from downloaded media)
- Fixed `Initialize-Exchange`: `$MinFFL`/`$MinDFL` now correctly set for new-org path
- Fixed `Initialize-Exchange`: `setup.exe /PrepareAD` exit code now checked and enforced
- Fixed FFL/DFL `$null` check in `Test-Preflight` (PowerShell: `$null -lt 17000` evaluates to `$true`)
- Pre-flight report now generated only on first phase (was repeated every phase)
- System Restore checkpoint: gracefully skipped on Windows Server (`Checkpoint-Computer` not available)

### v5.0 — March 2026

- **Pre-flight HTML report** (`-PreflightOnly`) — comprehensive validation before starting
- **Source server config export/import** (`-CopyServerConfig <ServerName>`) — copies Virtual Directory URLs, Transport settings, Receive Connectors, Outlook Anywhere, and Autodiscover SCP
- **PFX certificate import** (`-CertificatePath`) — imports and enables for IIS and SMTP
- **DAG join automation** (`-DAGName`) — automatically adds the server to a Database Availability Group
- **CSS-Exchange HealthChecker** — auto-downloaded and run at end of setup (`-SkipHealthCheck` to skip)
- **System Restore checkpoints** before each phase (`-NoCheckpoint` to skip)
- **Exchange Server SE RTM** support (build 15.02.2562.017) including OS compatibility check and Feb 2026 SU (KB5074992)

---

## Notes

- State file: `<InstallPath>\<ComputerName>_Install-Exchange15_state.xml` (default: `C:\Install\`)
- Log file: `<InstallPath>\reports\<ComputerName>_InstallLog_<yyyyMMdd-HHmmss>.log` — tier-prefixed entries (`INFO` / `WARNING` / `ERROR` / `EXE`; `VERBOSE` with `-Verbose`; `DEBUG` + `SUPPRESSED-ERROR` with `-Debug`); UTF-8 without BOM; see **Logging** section above
- Installation report (HTML): `<InstallPath>\reports\<ComputerName>_InstallationReport_<timestamp>.html` (+ `.pdf` if Edge available)
- Installation document (Word): `<InstallPath>\reports\<ComputerName>_InstallationDocument_<DE|EN>_<timestamp>.docx`
- All reports (Preflight, Installation, RBAC, HealthChecker) written to `<InstallPath>\reports\`
- With `-Autopilot`: AutoLogon is temporarily enabled and removed after next login
- All downloads use BITS with `Invoke-WebDownload` fallback (PS 5.1-compatible, handles certificate bypass)
- Pester tests (54 total): `Invoke-Pester .\Install-Exchange15.Tests.ps1 -Output Detailed` (requires Pester 5.x)

---

## References

- [Exchange Server Build Numbers and Release Dates](https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates)
- [Exchange 2019 Prerequisites](https://learn.microsoft.com/en-us/exchange/plan-and-deploy/prerequisites)
- [CSS-Exchange HealthChecker](https://github.com/microsoft/CSS-Exchange)
- [eightwone.com Blog](http://eightwone.com)
