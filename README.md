# EXpress

PowerShell script for fully unattended installation of Microsoft Exchange Server 2016, 2019, and Exchange SE — including prerequisites, Active Directory preparation, post-configuration, security hardening, and Word installation documentation.

**Maintainer:** st03ps | **Original author:** Michel de Rooij (michel@eightwone.com) · [eightwone.com](http://eightwone.com)
**Version:** 1.1.5 (April 2026)
**License:** As-Is, without warranty

**Versioning scheme:** `MAJOR.MINOR` = feature release · `MAJOR.MINOR.PATCH` = bugfix / maintenance release. Example: `1.1` introduces features, `1.1.1` contains only bugfixes on top of `1.1`.

---

## Supported Versions

Only the **latest CU** of each Exchange line is supported as an install target. Older CUs (Ex2019 CU10–CU14) are out of Microsoft SU support and rejected by the preflight check. Older CUs can still be **source servers** for `Export-SourceServerConfig` during migration.

| Exchange | Windows Server |
|---|---|
| Exchange 2016 CU23 (final) | Windows Server 2016 |
| Exchange 2019 CU15+ | Windows Server 2019, 2022, 2025 |
| Exchange Server SE (RTM+) | Windows Server 2019, 2022, 2025 |

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
.\EXpress.ps1
```

The menu lets you select the installation mode and toggle all options. Press letter keys to toggle switches instantly; press Enter to start.

If a `config.psd1` file is found in the same folder as the script or `.exe`, you are asked whether to use it — allowing a fully pre-configured start without manual input.

### Command-line / unattended

```powershell
# Install Mailbox role — interactive (Copilot mode)
.\EXpress.ps1 -SourcePath D:\Exchange

# Fully unattended with Autopilot (automatic reboots through all phases)
.\EXpress.ps1 -SourcePath D:\Exchange -Autopilot

# Load all parameters from a .psd1 config file (skips interactive menu)
.\EXpress.ps1 -ConfigFile .\deploy-mbx01.psd1

# Install prerequisites only, skip Exchange setup
.\EXpress.ps1 -NoSetup -SourcePath D:\Exchange

# Edge Transport role
.\EXpress.ps1 -InstallEdge -SourcePath D:\Exchange

# Recover a server
.\EXpress.ps1 -Recover -SourcePath D:\Exchange

# Pre-flight report only (no installation)
.\EXpress.ps1 -SourcePath D:\Exchange -PreflightOnly

# Swing migration: copy config from source server, import PFX, join DAG
.\EXpress.ps1 -SourcePath D:\Exchange -Autopilot `
    -CopyServerConfig EX01 -CertificatePath D:\certs\mail.pfx -DAGName DAG01

# Install Recipient Management Tools on an admin workstation
.\EXpress.ps1 -InstallRecipientManagement -SourcePath D:\Exchange

# Install Exchange Management Tools only (Server OS)
.\EXpress.ps1 -InstallManagementTools -SourcePath D:\Exchange

# Run all post-install optimizations on an existing Exchange server (no setup required)
.\EXpress.ps1 -StandaloneOptimize -Namespace mail.contoso.com `
    -CertificatePath C:\certs\mail.pfx -LogRetentionDays 30 `
    -RelaySubnets '10.0.1.0/24' -ExternalRelaySubnets '10.0.2.5'

# Generate a Word document for the full organisation on an existing server (ad-hoc inventory)
.\EXpress.ps1 -StandaloneDocument -Language DE

# Generate a customer-ready Word document — full org + all servers, sensitive values redacted
.\EXpress.ps1 -StandaloneDocument -Language EN -CustomerDocument

# Document only org-wide configuration (no per-server hardware queries)
.\EXpress.ps1 -StandaloneDocument -DocumentScope Org -Language DE

# Document specific servers only in a large farm
.\EXpress.ps1 -StandaloneDocument -IncludeServers EX01,EX02 -Language DE
```

### Compile to .exe (optional)

Use `Build.ps1` to compile the script into a self-contained Windows executable via PS2Exe:

```powershell
.\Build.ps1
# Output: .\EXpress.exe (runs elevated, all parameters preserved)
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

**Logo for Word documents** — place a `logo.png` (400×80 px recommended) in one of the following locations (first match wins): `<InstallPath>\sources\logo.png` (user-placed custom logo), alongside `EXpress.ps1`, or `assets\logo.png` (default sample shipped with the repo).

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

The script runs Phase 0 (preflight) once, then 6 install phases (1–6). State
is persisted to XML so Autopilot automatically resumes at the correct phase
after every reboot:

```
Phase 0 → Preflight checks, pre-flight HTML report
Phase 1 → Windows features, Windows Updates (optional)
Phase 2 → .NET Framework 4.8/4.8.1, OS hotfixes, Visual C++ Runtimes, URL Rewrite
Phase 3 → UCMA Runtime, Active Directory preparation (PrepareAD/PrepareSchema), AD replication check
Phase 4 → Run Exchange Setup
Phase 5 → Post-configuration (security hardening, performance tuning, certificate import, Exchange SU)
Phase 6 → Start services, IIS health check, Virtual Directory URLs, DAG join, connectors, HealthChecker, HTML + PDF Installation Report, Word Installation Document, cleanup
```

Reboots between phases are conditional in Autopilot: Phase 1→2 always reboots (Windows features); Phase 2→3 skips the reboot when `Test-RebootPending` reports nothing pending (typical on WS2025 + Exchange SE); Phase 5→6 skips the reboot unless the Exchange SU set `RebootRequired` (exit code 3010) or Windows signals a pending reboot.

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

See [RELEASE-NOTES.md](RELEASE-NOTES.md) for the full version history.

---

## Tools

Helper scripts in `tools/` — run standalone, no Exchange or EXpress state required.

| Script | Purpose |
|---|---|
| `tools/Get-EXpressDownloads.ps1` | Pre-stages all prerequisite downloads into a local `sources/` folder before deploying to air-gapped or proxy-restricted networks. Downloads: .NET 4.8 / 4.8.1, VC++ 2012 / 2013 Redistributables, UCMA 4.0, URL Rewrite 2.1, and the CSS-Exchange scripts (HealthChecker, EOMT, SetupAssist, SetupLogReviewer, ExchangeExtendedProtectionManagement, MonitorExchangeAuthCertificate). Idempotent — skips files already present. Use `-SkipDotNet` to skip the large .NET installers (75–116 MB) when running on WS2025 (where .NET 4.8.1 ships in-box). |
| `tools/Enable-EXpressRemoteQuery.ps1` | Enables WinRM / CIM over WSMan on a target Exchange server so EXpress can query hardware, pagefile, volume, and NIC data for the Word installation document via `Get-RemoteServerData`. Run locally on every server to be documented — or deploy the equivalent settings via GPO (see `docs/remote-query-setup.md`). Options: `-EnableHttps` adds a TCP 5986 HTTPS listener using the server's auth certificate; `-RestrictToGroup <ADGroup>` restricts PSSessionConfiguration to a specific AD group. |
| `tools/Build-InstallationTemplate.ps1` | Regenerates the default Word document templates (`templates/Exchange-installation-document-DE.docx` + `-EN.docx`). Run by the maintainer when the cover-page layout or token set changes. End users supply their own branded template via `-TemplatePath` (see F24). |
| `tools/Build-ConceptTemplate.ps1` | Generates Exchange concept / approval document templates (`templates/Exchange-concept-template-DE.docx` + `-EN.docx`) — 16 chapters covering architecture, sizing, security, migration, hybrid, and an acceptance page. Exchange SE–only scope (Exchange 2016 / 2019 are out of Microsoft support since October 2025). |
| `tools/Merge-Source.ps1` | Merges all `modules/*.ps1` files into `dist/EXpress.ps1`. Called automatically by `Build.ps1` before PS2Exe compilation; run manually after editing a module to verify the merged output. |

---

## Notes

- State file: `<InstallPath>\<ComputerName>_State.xml` (default: `C:\Install\`)
- Log file: `<InstallPath>\reports\<ComputerName>_InstallLog_<yyyyMMdd-HHmmss>.log` — tier-prefixed entries (`INFO` / `WARNING` / `ERROR` / `EXE`; `VERBOSE` with `-Verbose`; `DEBUG` + `SUPPRESSED-ERROR` with `-Debug`); UTF-8 without BOM; see **Logging** section above
- Installation report (HTML): `<InstallPath>\reports\<ComputerName>_InstallationReport_<timestamp>.html` (+ `.pdf` if Edge available)
- Installation document (Word): `<InstallPath>\reports\<ComputerName>_InstallationDocument_<DE|EN>_<timestamp>.docx`
- All reports (Preflight, Installation, RBAC, HealthChecker) written to `<InstallPath>\reports\`
- With `-Autopilot`: AutoLogon is temporarily enabled and removed after next login
- All downloads use BITS with `Invoke-WebDownload` fallback (PS 5.1-compatible, handles certificate bypass)
- Pester tests (54 total): `Invoke-Pester .\EXpress.Tests.ps1 -Output Detailed` (requires Pester 5.x)

---

## Acknowledgement

EXpress stands on the shoulders of giants. **Michel de Rooij's** [Install-Exchange15.ps1](http://eightwone.com) laid the groundwork — a solid, community-proven script that guided Exchange deployments for years. EXpress takes that legacy forward: full automation, security hardening, modular architecture, and a modern deployment experience. Thank you, [Michel](http://eightwone.com).

---

## References

- [Exchange Server Build Numbers and Release Dates](https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates)
- [Exchange 2019 Prerequisites](https://learn.microsoft.com/en-us/exchange/plan-and-deploy/prerequisites)
- [CSS-Exchange HealthChecker](https://github.com/microsoft/CSS-Exchange)
- [eightwone.com Blog](http://eightwone.com)
- [GitHub — st03psn/EXpress](https://github.com/st03psn/EXpress)
