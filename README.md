# Install-Exchange15.ps1

PowerShell script for fully unattended installation of Microsoft Exchange Server 2016, 2019, and Exchange SE — including prerequisites, Active Directory preparation, and post-configuration.

**Maintainer:** st03ps | **Original author:** Michel de Rooij (michel@eightwone.com) · [eightwone.com](http://eightwone.com)
**Version:** 5.4 (April 2026, last updated 2026-04-18)
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
# Install Mailbox role — interactive (prompts for credentials if -AutoPilot)
.\Install-Exchange15.ps1 -SourcePath D:\Exchange

# Fully unattended with AutoPilot (automatic reboots through all phases)
.\Install-Exchange15.ps1 -SourcePath D:\Exchange -AutoPilot

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
.\Install-Exchange15.ps1 -SourcePath D:\Exchange -AutoPilot `
    -CopyServerConfig EX01 -CertificatePath D:\certs\mail.pfx -DAGName DAG01

# Install Recipient Management Tools on an admin workstation
.\Install-Exchange15.ps1 -InstallRecipientManagement -SourcePath D:\Exchange

# Install Exchange Management Tools only (Server OS)
.\Install-Exchange15.ps1 -InstallManagementTools -SourcePath D:\Exchange

# Run all post-install optimizations on an existing Exchange server (no setup required)
.\Install-Exchange15.ps1 -StandaloneOptimize -Namespace mail.contoso.com `
    -CertificatePath C:\certs\mail.pfx -LogRetentionDays 30 `
    -RelaySubnets '10.0.1.0/24' -ExternalRelaySubnets '10.0.2.5'
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
| `-AutoPilot` | Fully unattended mode — reboots and resumes automatically |
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
| `-SkipInstallReport` | Skip the HTML/PDF installation report generated at Phase 6 completion |

See `deploy-example.psd1` for a fully documented configuration file template.

---

## Installation Phases

The script runs through 7 phases (0–6) and saves state in an XML file to
automatically resume after reboots:

```
Phase 0 → Preflight checks, AD preparation, pre-flight HTML report
Phase 1 → Windows features, .NET Framework, Windows Updates (optional)
Phase 2 → Visual C++ Redistributables, URL Rewrite, other prerequisites
Phase 3 → Hotfixes, additional packages
Phase 4 → Run Exchange Setup
Phase 5 → Post-configuration (security hardening, performance tuning, Exchange SU)
Phase 6 → Start services, IIS health check, DAG join, HealthChecker, cleanup
```

Recipient Management / Management Tools modes use phases 0–2 only.

---

## Post-Configuration (Phase 5)

The following best-practice configurations are automatically applied after Exchange setup:

### Security Hardening
- **HSTS header** — `Strict-Transport-Security: max-age=31536000` on OWA and ECP; only applied when `-CertificatePath` is set (avoids browser lockout with self-signed certificates)
- **Anonymous relay connectors** — dedicated internal (`-RelaySubnets`) and external (`-ExternalRelaySubnets`) relay connectors; `AnonymousUsers` removed from Default Frontend on success; account resolved via SID S-1-5-7 (language-independent)
- **Windows Defender exclusions** — folder, process, and extension exclusions per Microsoft guidance
- **SMBv1 disabled** — removes legacy protocol attack vector
- **SSL 3.0 disabled** (optional) — POODLE mitigation
- **RC4 disabled** (optional) — weak cipher removal
- **TLS 1.2/1.3 configuration** — SChannel and .NET Framework strong crypto (v4.0 + v2.0 paths)
- **ECC certificate support** (optional) — Elliptic Curve Cryptography
- **AES256-CBC encryption** — enabled by default
- **AMSI body scanning** (optional) — for OWA, ECP, EWS, PowerShell
- **WDigest credential caching disabled** — prevents cleartext password storage in LSASS
- **HTTP/2 disabled** — prevents known Exchange MAPI/RPC issues
- **Credential Guard disabled** — causes performance issues on Exchange (default on WS2025)
- **LAN Manager level 5** — NTLMv2 only, refuse LM and NTLM
- **Serialized Data Signing** — mitigates PowerShell deserialization attacks
- **LSA Protection (RunAsPPL)** — enabled for Exchange 2019 CU12+ / Exchange SE; takes effect after reboot

### Performance Tuning
- **High Performance power plan** — activated automatically
- **NIC power management disabled** — prevents adapter sleep
- **Pagefile configured** — 25% RAM (Ex2019+) or RAM+10MB (Ex2016)
- **TCP settings** — RPC timeout 120s, Keep-Alive 15 min
- **TCP Chimney and Task Offload disabled** — Microsoft recommendation
- **Windows Search service disabled** — Exchange uses its own content indexing
- **Scheduled defragmentation disabled** — not needed on Exchange servers
- **Disk allocation unit size verification** — warns if volumes are not 64KB formatted
- **CRL check timeout configured** — prevents Exchange startup delays
- **RSS enabled on all NICs** — ensures network traffic uses all CPU cores; sets receive queue count to physical core count
- **MaxConcurrentAPI configured** — Netlogon set to logical core count (min 10) to prevent 0xC000005E errors
- **CtsProcessorAffinityPercentage = 0** — Exchange Search best practice
- **NodeRunner memory limit = 0** — removes Search performance limiter
- **MAPI Front End Server GC** — enabled on systems with 20+ GB RAM

---

## What's New

### v5.4 — April 2026 (latest: 2026-04-18)

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
- Log file: `<InstallPath>\<ComputerName>_Install-Exchange15_<timestamp>.log` — always verbose; `[INFO]`, `[WARNING]`, `[ERROR]`, `[VERBOSE]` entries
- Installation report: `<InstallPath>\<ComputerName>_InstallationReport_<timestamp>.html` (+ `.pdf` if Edge available)
- With `-AutoPilot`: AutoLogon is temporarily enabled and removed after next login
- All downloads use BITS with `Invoke-WebDownload` fallback (PS 5.1-compatible, handles certificate bypass)
- Pester tests (54 total): `Invoke-Pester .\Install-Exchange15.Tests.ps1 -Output Detailed` (requires Pester 5.x)

---

## References

- [Exchange Server Build Numbers and Release Dates](https://docs.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates)
- [Exchange 2019 Prerequisites](https://docs.microsoft.com/en-us/exchange/plan-and-deploy/prerequisites)
- [CSS-Exchange HealthChecker](https://github.com/microsoft/CSS-Exchange)
- [eightwone.com Blog](http://eightwone.com)
