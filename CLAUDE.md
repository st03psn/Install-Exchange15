# Install-Exchange15.ps1 — Project Context for Claude Code

---

## What is this script?

`Install-Exchange15.ps1` is a PowerShell automation script for fully unattended installation
of Microsoft Exchange Server 2016, 2019, and Exchange SE (Subscription Edition),
including all prerequisites, Active Directory preparation, and post-configuration.

**Maintainer:** st03ps | **Original author:** Michel de Rooij (michel@eightwone.com)
**Current Version:** 5.1 (April 2026)
**PowerShell Requirement:** `#Requires -Version 5.1`
**Execution:** Must be run as Administrator

---

## Supported Environments

| Exchange Version | Windows Server |
|---|---|
| Exchange 2016 CU23 | Windows Server 2016 |
| Exchange 2019 CU10–CU14 | Windows Server 2019, 2022 |
| Exchange 2019 CU15+ | Windows Server 2025 |
| Exchange SE RTM | Windows Server 2022, 2025 |

---

## Architecture

### Phases (InstallPhase 0–6)

The script operates in phases. State is persisted in an XML file so that after
each reboot the script automatically resumes at the correct phase.

| Phase | Description |
|---|---|
| 0 | Initialization, preflight checks, AD preparation, pre-flight HTML report |
| 1 | Install Windows features, .NET Framework, export source server config |
| 2 | Wait for reboots, install prerequisites (VC++, URL Rewrite, etc.) |
| 3 | Additional prerequisites and hotfixes |
| 4 | Run Exchange Setup, set transport services to Manual |
| 5 | Post-configuration (Defender exclusions, TLS, security hardening, performance tuning), import server config, import PFX certificate |
| 6 | Restore services, IIS health check, DAG join, HealthChecker, cleanup |

### State Management

```powershell
$StateFile = "$InstallPath\${env:computerName}_Install-Exchange15_state.xml"
Save-State $State      # Export-Clixml
Restore-State          # Import-Clixml (returns empty hashtable if not found)
```

The state hashtable contains all passed parameters plus runtime variables
(phase, versions, paths, flags).

### AutoPilot Mode

With `-AutoPilot`, the script automatically reboots the server after each phase
and resumes itself via a `RunOnce` registry entry. Credentials are stored
encrypted in the state file.

```powershell
# RunOnce entry (optimized: dynamic PS interpreter path)
$PSExe = (Get-Process -Id $PID).Path   # powershell.exe or pwsh.exe
```

---

## Key Constants

```powershell
# OS versions (build prefix)
$WS2016_MAJOR   = '10.0'
$WS2019_PREFULL = '10.0.17709'   # WS2019 pre-RTM threshold (RTM = 10.0.17763)
$WS2022_PREFULL = '10.0.20348'
$WS2025_PREFULL = '10.0.26100'   # IMPORTANT: was incorrectly '10.0.20348' (= WS2022)

# Exchange Setup versions (ExSetup.exe)
$EX2016SETUPEXE_CU23 = '15.01.2507.006'
$EX2019SETUPEXE_CU10 = '15.02.0922.007'
$EX2019SETUPEXE_CU11 = '15.02.0986.005'
$EX2019SETUPEXE_CU12 = '15.02.1118.007'
$EX2019SETUPEXE_CU13 = '15.02.1258.012'
$EX2019SETUPEXE_CU14 = '15.02.1544.004'
$EX2019SETUPEXE_CU15 = '15.02.1748.008'
$EXSESETUPEXE_RTM    = '15.02.2562.017'

# .NET Framework
$NETVERSION_48  = 528040
$NETVERSION_481 = 533320

# Autodiscover SCP LDAP filter (central constant, used 4x)
$AUTODISCOVER_SCP_FILTER    = '(&(cn={0})(objectClass=serviceConnectionPoint)...)'
$AUTODISCOVER_SCP_MAX_RETRIES = 30   # 30 x 10 sec = 5 min timeout
```

---

## Function Overview

### Logging

```powershell
Write-ToTranscript $Level $Text   # Internal helper function (added during refactoring)
Write-MyOutput  $Text             # Write-Output + Transcript [INFO]
Write-MyWarning $Text             # Write-Warning + Transcript [WARNING]
Write-MyError   $Text             # Write-Error + Transcript [ERROR]
Write-MyVerbose $Text             # Write-Verbose + Transcript [VERBOSE]
```

### Preflight Checks (`Test-Preflight`)

Validates: admin rights, domain membership, OS version, Exchange version,
AD Forest/Domain level, static IP, roles, organization name, setup path.

### Package Installation

```powershell
Invoke-WebDownload -Uri $URL -OutFile $Path             # PS 5.1-compatible download (cert bypass)
Get-MyPackage   $Package $URL $FileName $InstallPath   # Download via BITS (Invoke-WebDownload fallback)
Install-MyPackage $PackageID $Package $FileName $URL $Arguments
Test-MyPackage  $PackageID                              # Registry + CIM check
Invoke-Process  $FilePath $FileName $ArgumentList       # MSU/MSI/MSP/EXE
Invoke-Extract  $FilePath $FileName                     # ZIP via Expand-Archive
```

### TLS/Cryptography

```powershell
Set-SchannelProtocol -Protocol 'TLS 1.2' -Enable $true/$false   # Helper
Set-NetFrameworkStrongCrypto                                      # Helper
Set-TLSSettings -TLS12 -TLS13                                    # Main function
Disable-SSL3
Disable-RC4
Enable-ECC
Enable-CBC
Enable-AMSI
```

### AD / Exchange

```powershell
Get-ForestRootNC / Get-RootNC / Get-ForestConfigurationNC
Get-ForestFunctionalLevel / Get-ExchangeForestLevel / Get-ExchangeDomainLevel
Get-ExchangeOrganization / Test-ExchangeOrganization
Test-ExistingExchangeServer $Name
Clear-AutodiscoverServiceConnectionPoint $Name [-Wait]
Set-AutodiscoverServiceConnectionPoint $Name $ServiceBinding [-Wait]
Initialize-Exchange          # PrepareAD / PrepareSchema
```

### Post-Configuration

```powershell
Enable-WindowsDefenderExclusions   # Folder and process exclusions
Enable-HighPerformancePowerPlan
Disable-NICPowerManagement
Set-Pagefile
Set-TCPSettings                    # RPC Timeout, Keep-Alive
Disable-SMBv1                      # Security: disable legacy protocol
Disable-WindowsSearchService       # Exchange has own content indexing
Disable-WDigestCredentialCaching   # Security: prevent cleartext creds in LSASS
Disable-HTTP2                      # Exchange MAPI/RPC compatibility
Disable-TCPOffload                 # Performance: disable chimney/offload
Test-DiskAllocationUnitSize        # Verify 64KB allocation units
Disable-UnnecessaryScheduledTasks  # Disable defrag etc.
Set-CRLCheckTimeout                # Prevent startup delays
Disable-CredentialGuard            # Performance: disable on Exchange servers
Set-LmCompatibilityLevel           # Security: NTLMv2 only (level 5)
Enable-RSSOnAllNICs                # Performance: enable RSS + set queues to physical core count
Set-MaxConcurrentAPI               # Performance/Security: Netlogon MaxConcurrentApi = logical core count (min 10)
Set-CtsProcessorAffinityPercentage # Search: set to 0 for best performance
Enable-SerializedDataSigning       # Security: mitigate serialization attacks
Set-NodeRunnerMemoryLimit          # Search: remove memory limit (set to 0)
Enable-MAPIFrontEndServerGC        # Performance: Server GC for 20+ GB RAM
Enable-LSAProtection               # Security: RunAsPPL=1 (Exchange 2019 CU12+/SE; reboot required)
```

### v5.0 Features

```powershell
New-PreflightReport                              # HTML report with all preflight checks (-PreflightOnly to exit after)
Export-SourceServerConfig $ServerName             # Export config from source Exchange server via Remote PS
Import-ServerConfig                              # Import Virtual Dirs, Transport, Receive Connectors from export
Import-ExchangeCertificateFromPFX                # Import PFX cert, enable for IIS+SMTP (-CertificatePath)
Join-DAG                                         # Join server to Database Availability Group (-DAGName)
Invoke-HealthChecker                             # Download and run CSS-Exchange HealthChecker (-SkipHealthCheck to skip)
# System Restore checkpoints created before each phase (-NoCheckpoint to skip)
# Set-RegistryValue now idempotent (skips if value already set)
```

### v5.1 Features

```powershell
Show-InstallationMenu                            # Interactive console menu (modes 1-5, letter toggles A-R; RawUI.ReadKey for instant toggle, falls back to Read-Host)
Get-ValidatedCredentials                         # Credential retry loop (max 3 attempts, R=Retry/Q=Quit)
Install-PendingWindowsUpdates                    # Windows security/critical updates via PSWindowsUpdate + WUA COM fallback
Get-LatestExchangeSecurityUpdate                 # Detect latest Exchange SU from built-in $ExchangeSUMap
Install-ExchangeSecurityUpdate                   # Download via BITS + install Exchange SU (Phase 5)
Disable-ServerManagerAtLogon                     # Machine-wide Server Manager disable (policy + default hive + scheduled task)
Get-OSVersionText $OSVersion                     # Friendly name for OS build (e.g. "Windows Server 2025 (build 10.0.26100)")
# -ConfigFile: load all parameters from .psd1 (Import-PowerShellDataFile), skips interactive menu
# Build.ps1: PS2Exe wrapper to compile script to .exe with version metadata
```

---

## Optimization History

### 2025-03-21 — Round 1: Critical Fixes
| # | Change | Lines (before) |
|---|---|---|
| Bug | `$WS2025_PREFULL` = `10.0.26100` (was `10.0.20348` = WS2022) | 645 |
| Refactor | `Write-ToTranscript` helper, simplified all 4 `Write-My*` functions | 709-739 |
| API | `Get-WmiObject` (MSExchangeServiceHost) to `Get-CimInstance` | 2797 |
| API | `WebClient`/`ServerCertificateValidationCallback` to `Invoke-WebRequest -SkipCertificateCheck` | 2838-2851 |
| Feature | `Test-RebootPending` added Windows Update registry key | 805-814 |

### 2025-03-21 — Round 2: Security & Code Quality
| # | Change |
|---|---|
| Constants | Introduced `$AUTODISCOVER_SCP_FILTER` + `$AUTODISCOVER_SCP_MAX_RETRIES` |
| Security | Added comment in `Enable-AutoLogon` about plaintext password risk |
| API | `Enable-RunOnce`: `$PSHome\powershell.exe` to `(Get-Process -Id $PID).Path` |
| API | `Invoke-Extract`: COM `shell.application` to `Expand-Archive` |
| Bug | Infinite loops in SCP background jobs: added retry limit + timeout |
| API | `Get-WmiObject win32_quickfixengineering` to `Get-CimInstance Win32_QuickFixEngineering` |
| Convention | `get-FullDomainAccount` to `Get-FullDomainAccount` |
| Typo | `'Wil run Setup'` to `'Will run Setup'` |
| Exception | `$Error[0].ExceptionMessage` to `$_.Exception.Message` in all catch blocks |

### 2025-03-21 — Round 3: WMI Migration & Bug Fixes
| # | Change |
|---|---|
| API | `mkdir` to `New-Item -ItemType Directory` |
| Bug | `Remove-NETFrameworkInstallBlock`: error message "set" to "remove" |
| Bug | Removed stray `$Location` output in `Enable-WindowsDefenderExclusions` |
| API | `$CS = Get-WmiObject Win32_ComputerSystem` + `.Put()` to CIM + `Set-CimInstance` |
| API | `Get-WmiObject Win32_NetworkAdapter` + `MSPower_DeviceEnable` + `psbase.Put()` to CIM |
| API | `Get-WmiObject Win32_ComputerSystem/Win32_NetworkAdapterConfiguration` to CIM |
| API | `Get-WmiObject Win32_PowerPlan` to CIM |
| Bug | `Get-WindowsFeature` check: `if (Get-WindowsFeature $x)` to `.Installed` |
| Refactor | `Set-TLSSettings`: 50 lines duplicate code to `Set-SchannelProtocol` + `Set-NetFrameworkStrongCrypto` |
| Cosmetic | `$Env:SystemRoot` without unnecessary string interpolation |

### 2026-03-21 — Round 4: Security Hardening & Performance
| # | Change |
|---|---|
| Security | `Disable-SMBv1`: disable legacy SMBv1 protocol |
| Security | `Disable-WDigestCredentialCaching`: prevent cleartext creds in LSASS |
| Security | `Disable-HTTP2`: disable HTTP/2 for Exchange MAPI/RPC compatibility |
| Security | `Set-CRLCheckTimeout`: prevent startup delays with unreachable CRL endpoints |
| Performance | `Disable-WindowsSearchService`: Exchange has own content indexing |
| Performance | `Disable-TCPOffload`: disable TCP Chimney and Task Offload |
| Performance | `Disable-UnnecessaryScheduledTasks`: disable defrag on Exchange servers |
| Validation | `Test-DiskAllocationUnitSize`: warn if volumes not using 64KB allocation units |

### 2026-03-21 — Round 5: CSS-Exchange HealthChecker Recommendations
| # | Change |
|---|---|
| Security | `Disable-CredentialGuard`: disable on Exchange servers (performance, default on WS2025) |
| Security | `Set-LmCompatibilityLevel`: NTLMv2 only (level 5) |
| Security | `Enable-SerializedDataSigning`: mitigate PowerShell serialization attacks |
| Performance | `Enable-RSSOnAllNICs`: enable Receive Side Scaling on all adapters |
| Performance | `Set-CtsProcessorAffinityPercentage`: set to 0 for Exchange Search |
| Performance | `Set-NodeRunnerMemoryLimit`: remove memory limit for Exchange Search |
| Performance | `Enable-MAPIFrontEndServerGC`: enable Server GC for MAPI FE (20+ GB RAM) |
| TLS | `Set-NetFrameworkStrongCrypto`: extended to v2.0 paths (HealthChecker requirement) |
| Bug | `Enable-MSExchangeAutodiscoverAppPool`: fixed `$Error[0]` to `$_.Exception.Message` |

### 2026-03-22 — Round 6: v5.0 Major Feature Release
| # | Change |
|---|---|
| Feature | `New-PreflightReport`: HTML pre-flight validation report (`-PreflightOnly`) |
| Feature | `Export-SourceServerConfig` / `Import-ServerConfig`: copy config from source server (`-CopyServerConfig`) |
| Feature | `Import-ExchangeCertificateFromPFX`: PFX certificate import with IIS+SMTP binding (`-CertificatePath`) |
| Feature | `Join-DAG`: automated DAG membership (`-DAGName`) |
| Feature | `Invoke-HealthChecker`: auto-download and run CSS-Exchange HealthChecker (`-SkipHealthCheck`) |
| Feature | System Restore checkpoints before each phase (`-NoCheckpoint` to skip) |
| Quality | `Set-RegistryValue`: idempotency guard (skip if value already set) |

### 2026-04-08 — Round 7: v5.01 Bugfixes & Robustness
| # | Change |
|---|---|
| Feature | Auto-elevation: script re-launches elevated via `Start-Process -Verb RunAs` |
| Feature | Auto-unblock: detect `Zone.Identifier` on source files, `Unblock-File` in preflight |
| Bug | `Initialize-Exchange`: `$MinFFL`/`$MinDFL` now set in new-org path (was `$null`) |
| Bug | `Initialize-Exchange`: `Invoke-Process` exit code now checked (was silently ignored) |
| Bug | `Test-Preflight` FFL/DFL check: distinguish `$null` (AD not prepared) from version too low |
| Quality | Pre-flight report: only generated on first phase (was every phase) |
| Quality | System Restore checkpoint: skip gracefully on Windows Server (no `Checkpoint-Computer`) |

### 2026-04-10 — Round 8: v5.1 Major Feature Release (Maintainer: st03ps)
| # | Change |
|---|---|
| Header | Author updated to `st03ps` with credit to Michel de Rooij; all docs/comments in English |
| Header | `$ScriptVersion = '5.1'`; revision history corrected to ascending order |
| Header | `.PARAMETER` block synchronized: added PreflightOnly, CopyServerConfig, CertificatePath, DAGName, SkipHealthCheck, NoCheckpoint, InstallRecipientManagement, InstallManagementTools, RecipientMgmtCleanup, ConfigFile, InstallWindowsUpdates, SkipWindowsUpdates |
| Feature | New ParameterSet `'R'` (`-InstallRecipientManagement`): 3-phase install of Exchange Management Tools on Server or Client OS, with RSAT-ADDS prereqs, Add-PermissionForEMT.ps1, and desktop shortcut |
| Feature | New ParameterSet `'T'` (`-InstallManagementTools`): 3-phase install of `setup.exe /roles:ManagementTools` on Server OS |
| Feature | `Show-InstallationMenu`: interactive console menu (mode 1-5, letter toggles A-R, greying of unavailable options, Read-Host based for RDP/Hyper-V/Terminal compatibility) |
| Feature | `Get-ValidatedCredentials`: interactive credential retry loop (max 3 attempts with R=Retry/Q=Quit); only for interactive sessions; single-validation kept for `-Credentials` CLI param |
| Feature | `Install-PendingWindowsUpdates`: Windows security/critical update install via PSWindowsUpdate module (auto-installs if missing) with WUA COM API fallback; integrated into Phase 1 |
| Feature | `Get-LatestExchangeSecurityUpdate` / `Install-ExchangeSecurityUpdate`: Exchange SU detection via built-in `$ExchangeSUMap` hashtable, download via BITS, install in Phase 5 |
| Feature | `Build.ps1`: compiles `Install-Exchange15.ps1` to `.exe` via PS2Exe (`-requireAdmin`, no `-noConsole`, embeds version/copyright metadata) |
| Quality | `$ScriptFullName` fallback to `[Diagnostics.Process]::GetCurrentProcess().MainModule.FileName` when `$MyInvocation.MyCommand.Path` is empty (PS2Exe compiled run) |
| Quality | `Enable-RunOnce`: detects `.exe` vs `.ps1` launch, sets RunOnce entry accordingly for AutoPilot compatibility in PS2Exe builds |
| Quality | MAX_PHASE = 3 for R/T modes (was always 3 for NoSetup, 6 otherwise) |
| New params | `-InstallWindowsUpdates`, `-SkipWindowsUpdates`, `-InstallRecipientManagement`, `-RecipientMgmtCleanup`, `-InstallManagementTools`, `-ConfigFile` |

### 2026-04-15 — Round 9: Pitfall Fixes & PS 5.1 Compatibility
| # | Change |
|---|---|
| Bug | `$Error[0].ExceptionMessage` → `$_.Exception.Message` in `Start-DisableMSExchangeAutodiscoverAppPoolJob` ScriptBlock |
| Bug | `Write-Host` → `Write-MyOutput` in `Enable-MSExchangeAutodiscoverAppPool` (was outside job scope, should use logging helper) |
| Bug | Missing `$InstallWindowsUpdates` mapping in menu result block — toggle R was read but never written to `$InstallWindowsUpdates` |
| Feature | `Invoke-WebDownload`: new PS 5.1-compatible download helper (replaces bare `Invoke-WebRequest` fallback calls) |
| Fix | IIS health check endpoint loop: same PS 5.1 / PS 6+ version split using `WebClient.DownloadString()` |

### 2026-04-15 — Round 10: Quality & Performance Improvements
| # | Change |
|---|---|
| Quality | `Get-SetupTextVersion`: direct hashtable lookup first (O(1) for exact CU builds), fallback sort only for SU builds |
| Performance | `Enable-RSSOnAllNICs`: additionally sets `NumberOfReceiveQueues` to physical core count (HealthChecker requirement) |
| Performance | `Set-MaxConcurrentAPI`: new function — Netlogon `MaxConcurrentApi` set to logical processor count (min 10) to prevent 0xC000005E errors under Exchange auth load |
| Security | `Clear-DesktopBackground`: new function — removes desktop wallpaper at each phase start via Registry + RUNDLL32 |
| Bug | `Cleanup`: `Get-WindowsFeature Bits` → `(Get-WindowsFeature -Name 'Bits').Installed` |

### 2026-04-15 — Round 11: Performance & Security Improvements
| # | Change |
|---|---|
| Performance | `Clear-DesktopBackground`: `Add-Type`/C#-Compiler removed — now uses Registry + `RUNDLL32 UpdatePerUserSystemParameters` (3–10s faster per phase start) |
| Performance | `Get-DetectedFileVersion`: `Get-Command` → `[System.Diagnostics.FileVersionInfo]::GetVersionInfo()` (avoids PATH/module discovery overhead on ISO paths) |
| Quality | Parameter block: `ValueFromPipelineByPropertyName = $false` removed from all `[parameter()]` attributes (it is the default; ~60 lines saved) |
| Security | `Enable-LSAProtection`: new function — sets `HKLM:\...\Lsa\RunAsPPL = 1`; called in Phase 5; effective after reboot; Exchange 2019 CU12+/SE compatible |
| UX | `Write-PhaseProgress`: new helper — Id 0 = overall phase progress (1-6), Id 1 = Phase 5 step counter (~25 steps) |
| UX | `Show-InstallationMenu`: RawUI.ReadKey with `KeyAvailable` probe; instant toggle without Enter; falls back to Read-Host if console unavailable |

---

## Open Items / Possible Next Steps

### v5.1 Remaining Work
- [x] **Config-Templates:** `-ConfigFile` loads all parameters from a .psd1 — see `deploy-example.psd1`
- [x] **Write-Progress** indicators per phase — `Write-PhaseProgress` helper (Id 0 = overall, Id 1 = Phase 5 steps)
- [x] **`Show-InstallationMenu`:** RawUI.ReadKey with `$host.UI.RawUI.KeyAvailable` probe; falls back to Read-Host if console unavailable (PS2Exe, redirected stdin)
- [x] **`$ExchangeSUMap`:** Fully rewritten — direct download.microsoft.com URLs, `.exe` filenames, all supported CU versions (SE RTM, 2019 CU13-15, 2016 CU23); `Install-ExchangeSecurityUpdate` bug `.msp`→`.exe` fixed
- [x] **Recipient Management / Management Tools:** Implementation complete (prereqs, setup.exe, EMT-Script, desktop shortcut, AD cleanup note)

### General Quality / Future
- [x] Reduce parameter block redundancy — `ValueFromPipelineByPropertyName = $false` (default) removed from all `[parameter()]` attributes
- [x] Enable LSA Protection (RunAsPPL) — `Enable-LSAProtection` added, Phase 5; Exchange 2019 CU12+/SE compatible, takes effect after reboot
- [x] Pester tests — `Install-Exchange15.Tests.ps1`: covers `Get-SetupTextVersion`, `Get-OSVersionText`, `Get-FullDomainAccount`, `ExchangeSUMap` structure validation. Run: `Invoke-Pester .\Install-Exchange15.Tests.ps1 -Output Detailed`
