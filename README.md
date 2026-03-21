# Install-Exchange15.ps1

PowerShell script for fully unattended installation of Microsoft Exchange Server 2016, 2019, and Exchange SE — including prerequisites, Active Directory preparation, and post-configuration.

**Author:** Michel de Rooij (michel@eightwone.com) · [eightwone.com](http://eightwone.com)
**Version:** 4.30 (March 2026)
**License:** As-Is, without warranty

---

## Supported Versions

| Exchange | Windows Server |
|---|---|
| Exchange 2016 CU23 | Windows Server 2016 |
| Exchange 2019 CU10–CU14 | Windows Server 2019, 2022 |
| Exchange 2019 CU15+ / Exchange SE | Windows Server 2022, 2025 |

---

## Prerequisites

- PowerShell 5.1 or later
- Run as local Administrator
- Domain membership (except Edge role)
- Schema Admin + Enterprise Admin rights for AD preparation
- Static IP address (or Azure Guest Agent detected)
- Exchange setup files (ISO or extracted) accessible

---

## Usage

```powershell
# Install Mailbox role (interactive)
.\Install-Exchange15.ps1 -InstallMailbox -SourcePath D:\Exchange

# Fully unattended with AutoPilot (automatic reboots through all phases)
.\Install-Exchange15.ps1 -InstallMailbox -SourcePath D:\Exchange -AutoPilot -Credentials (Get-Credential)

# Install prerequisites only, skip Exchange setup
.\Install-Exchange15.ps1 -NoSetup

# Edge Transport role
.\Install-Exchange15.ps1 -InstallEdge -SourcePath D:\Exchange

# Recover a server
.\Install-Exchange15.ps1 -Recover -SourcePath D:\Exchange
```

### Key Parameters

| Parameter | Description |
|---|---|
| `-InstallMailbox` | Install the Mailbox role |
| `-InstallEdge` | Install the Edge Transport role |
| `-SourcePath` | Path to Exchange setup files or ISO |
| `-TargetPath` | Target folder for Exchange (default: `C:\Program Files\Microsoft\Exchange Server\V15`) |
| `-AutoPilot` | Fully unattended mode with automatic reboots |
| `-Credentials` | Credentials for AutoPilot |
| `-OrganizationName` | Exchange organization name (new installation) |
| `-InstallMDBName` | Name of the first mailbox database |
| `-InstallMDBDBPath` | Path for database files (.edb) |
| `-InstallMDBLogPath` | Path for transaction logs |
| `-IncludeFixes` | Install recommended security updates after setup |
| `-DisableSSL3` | Disable SSL 3.0 (POODLE) |
| `-DisableRC4` | Disable RC4 encryption |
| `-EnableTLS12` | Explicitly enable TLS 1.2 |
| `-EnableTLS13` | Enable TLS 1.3 (WS2022+, Exchange 2019 CU15+) |
| `-EnableECC` | Enable ECC certificates |
| `-EnableAMSI` | Enable AMSI body scanning |
| `-NoSetup` | Install prerequisites only, skip Exchange setup |
| `-Phase` | Start directly at a specific phase (0–6) |

---

## Installation Phases

The script runs through 7 phases (0–6) and saves state in an XML file to
automatically resume after reboots:

```
Phase 0 → Preflight checks, AD preparation
Phase 1 → Windows features, .NET Framework
Phase 2 → Visual C++ Redistributables, URL Rewrite, other prerequisites
Phase 3 → Hotfixes, additional packages
Phase 4 → Run Exchange Setup
Phase 5 → Post-configuration (security hardening, performance tuning)
Phase 6 → Start services, IIS health check, cleanup
```

---

## Post-Configuration (Phase 5)

The following best-practice configurations are automatically applied after Exchange setup:

### Security Hardening
- **Windows Defender exclusions** — folder, process, and extension exclusions per Microsoft guidance
- **SMBv1 disabled** — removes legacy protocol attack vector
- **SSL 3.0 disabled** (optional) — POODLE mitigation
- **RC4 disabled** (optional) — weak cipher removal
- **TLS 1.2/1.3 configuration** — SChannel and .NET Framework strong crypto
- **ECC certificate support** (optional) — Elliptic Curve Cryptography
- **AES256-CBC encryption** — enabled by default
- **AMSI body scanning** (optional) — for OWA, ECP, EWS, PowerShell
- **WDigest credential caching disabled** — prevents cleartext password storage in LSASS memory
- **HTTP/2 disabled** — prevents known Exchange MAPI/RPC issues

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

---

## Changes from Original (v4.30)

### v4.30 — Security Hardening & Performance (March 2026)

**New post-configuration features:**
- Disable SMBv1 protocol
- Disable Windows Search service
- Disable WDigest credential caching
- Disable HTTP/2 protocol for Exchange compatibility
- Disable TCP Chimney and Task Offload
- Verify 64KB disk allocation unit sizes
- Disable unnecessary scheduled tasks (defrag)
- Configure CRL check timeout (15 seconds)

### v4.22 — Bug Fixes & Modernization (March 2025)

**Bug fixes:**
- **`$WS2025_PREFULL`** corrected: `10.0.26100` (was incorrectly `10.0.20348` = WS2022)
- **`Get-WindowsFeature` check** corrected: now uses `.Installed` property instead of implicit boolean cast
- **Error message** in `Remove-NETFrameworkInstallBlock` corrected: "Unable to remove" instead of "Unable to set"
- **Stray console output** in `Enable-WindowsDefenderExclusions` removed
- **Infinite loops** in Autodiscover SCP background jobs: added retry limit of 30 x 10 sec
- **`$Error[0].ExceptionMessage`** replaced with `$_.Exception.Message` in all catch blocks
- **Typo** `'Wil run Setup'` to `'Will run Setup'`

**API Modernization:**

| Old | New |
|---|---|
| `Get-WmiObject` (all 9 occurrences) | `Get-CimInstance` |
| `$obj.psbase.Put()` | `Set-CimInstance -InputObject $obj -Property @{...}` |
| `New-Object Net.WebClient` + `ServerCertificateValidationCallback` | `Invoke-WebRequest -SkipCertificateCheck -UseBasicParsing` |
| `New-Object -com shell.application` (ZIP extraction) | `Expand-Archive` |
| `$PSHome\powershell.exe` (RunOnce) | `(Get-Process -Id $PID).Path` (PS 7 compatible) |
| `mkdir` | `New-Item -ItemType Directory` |

**Refactoring:**
- New `Write-ToTranscript` helper — eliminates 4x duplicated `Test-Path`/`Out-File` logic
- New `Set-SchannelProtocol` and `Set-NetFrameworkStrongCrypto` helpers — reduces `Set-TLSSettings` from ~90 to ~35 lines
- Central `$AUTODISCOVER_SCP_FILTER` constant (was hardcoded 4x)
- Function name `get-FullDomainAccount` to `Get-FullDomainAccount` (PS convention)
- `Test-RebootPending` — added third registry check for Windows Update

---

## Notes

- State file location: `%TEMP%\<ComputerName>_Install-Exchange15_state.xml`
- Log file: `%TEMP%\<ComputerName>_Install-Exchange15.log`
- With `-AutoPilot`: UAC is temporarily disabled and re-enabled after completion
- AutoLogon password is automatically removed from registry after next login
- `-SkipCertificateCheck` on `Invoke-WebRequest` requires PowerShell 6+; PS 5.1 may need a fallback

---

## References & Documentation

- [Exchange Server Build Numbers](https://docs.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates)
- [Exchange 2019 Prerequisites](https://docs.microsoft.com/en-us/exchange/plan-and-deploy/prerequisites)
- [eightwone.com Blog](http://eightwone.com)
