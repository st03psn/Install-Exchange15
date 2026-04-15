# Install-Exchange15.ps1

PowerShell script for fully unattended installation of Microsoft Exchange Server 2016, 2019, and Exchange SE — including prerequisites, Active Directory preparation, and post-configuration.

**Maintainer:** st03ps | **Original author:** Michel de Rooij (michel@eightwone.com) · [eightwone.com](http://eightwone.com)
**Version:** 5.1 (April 2026)
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

### v5.1 — April 2026

- **Interactive installation menu** — start without parameters; numbered mode selection + letter toggles; RawUI.ReadKey for instant response (falls back to Read-Host for RDP/PS2Exe compatibility)
- **Credential validation loop** — `Get-ValidatedCredentials` with up to 3 retries (R=Retry / Q=Quit)
- **Recipient Management Tools** (`-InstallRecipientManagement`) — 3-phase install for dedicated admin workstations; works on Server and Client OS; includes RSAT-ADDS prereqs, Add-PermissionForEMT.ps1, and desktop shortcut
- **Exchange Management Tools** (`-InstallManagementTools`) — 3-phase install of `setup.exe /roles:ManagementTools` on Server OS
- **Config file support** (`-ConfigFile`) — load all parameters from a `.psd1` (see `deploy-example.psd1`)
- **Windows Update integration** (`-InstallWindowsUpdates`) — installs security/critical updates in Phase 1 via PSWindowsUpdate module (auto-installed if needed) with WUA COM fallback
- **Exchange Security Update automation** — built-in `$ExchangeSUMap` with direct download URLs for all supported CUs; installed in Phase 5 when `-IncludeFixes` is set
- **Build.ps1** — compile script to self-contained `.exe` via PS2Exe (`-requireAdmin`, preserves all parameters, embeds version metadata)
- **Write-Progress indicators** — overall phase progress (Id 0) and Phase 5 step counter (Id 1, ~25 steps)
- **LSA Protection** — `Enable-LSAProtection` sets `RunAsPPL=1`; Exchange 2019 CU12+ / Exchange SE; takes effect after reboot
- **MaxConcurrentAPI** — `Set-MaxConcurrentAPI` configures Netlogon to prevent 0xC000005E under Exchange auth load
- **Performance fixes** — `Clear-DesktopBackground` no longer uses `Add-Type`/C# (3–10s faster per phase start); `Get-DetectedFileVersion` uses `FileVersionInfo` API instead of `Get-Command`

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
- Log file: `<InstallPath>\<ComputerName>_Install-Exchange15.log`
- With `-AutoPilot`: AutoLogon is temporarily enabled and removed after next login
- All downloads use BITS with `Invoke-WebDownload` fallback (PS 5.1-compatible, handles certificate bypass)
- Pester tests: `Invoke-Pester .\Install-Exchange15.Tests.ps1 -Output Detailed` (requires Pester 5.x)

---

## References

- [Exchange Server Build Numbers and Release Dates](https://docs.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates)
- [Exchange 2019 Prerequisites](https://docs.microsoft.com/en-us/exchange/plan-and-deploy/prerequisites)
- [CSS-Exchange HealthChecker](https://github.com/microsoft/CSS-Exchange)
- [eightwone.com Blog](http://eightwone.com)
