# Install-Exchange15.ps1 — Project Context for Claude Code

---

## What is this script?

`Install-Exchange15.ps1` is a PowerShell automation script for fully unattended installation
of Microsoft Exchange Server 2016, 2019, and Exchange SE (Subscription Edition),
including all prerequisites, Active Directory preparation, and post-configuration.

**Author:** Michel de Rooij (michel@eightwone.com)
**Current Version:** 4.31 (March 2026)
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
| 0 | Initialization, preflight checks, AD preparation |
| 1 | Install Windows features, .NET Framework |
| 2 | Wait for reboots, install prerequisites (VC++, URL Rewrite, etc.) |
| 3 | Additional prerequisites and hotfixes |
| 4 | Run Exchange Setup, set transport services to Manual |
| 5 | Post-configuration (Defender exclusions, TLS, security hardening, performance tuning) |
| 6 | Restore services, IIS health check, cleanup |

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
$WS2019_PREFULL = '10.0.17709'
$WS2022_PREFULL = '10.0.20348'
$WS2025_PREFULL = '10.0.26100'   # IMPORTANT: was incorrectly '10.0.20348' (= WS2022)

# Exchange Setup versions (ExSetup.exe)
$EX2016SETUPEXE_CU23    = '15.01.2507.006'
$EX2019SETUPEXE_CU10-15 = '15.02.xxxx.xxx'
$EXSESETUPEXE_RTM       = '15.02.2562.017'

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
Get-MyPackage   $Package $URL $FileName $InstallPath   # Download via BITS
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
Enable-RSSOnAllNICs                # Performance: enable Receive Side Scaling
Set-CtsProcessorAffinityPercentage # Search: set to 0 for best performance
Enable-SerializedDataSigning       # Security: mitigate serialization attacks
Set-NodeRunnerMemoryLimit          # Search: remove memory limit (set to 0)
Enable-MAPIFrontEndServerGC        # Performance: Server GC for 20+ GB RAM
```

---

## Known Pitfalls & Design Decisions

### 1. CIM instead of WMI (fully migrated)
All `Get-WmiObject` calls have been replaced with `Get-CimInstance`.
For write operations: `Set-CimInstance -InputObject $obj -Property @{...}`
instead of `$obj.Property = ...; $obj.psbase.Put()`.

### 2. Get-WindowsFeature always checks `.Installed`
```powershell
# WRONG - always returns an object, even if not installed
if (Get-WindowsFeature 'Web-Server') { ... }

# CORRECT
if ((Get-WindowsFeature -Name 'Web-Server').Installed) { ... }
```

### 3. Autodiscover SCP Background Jobs
`Clear-` and `Set-AutodiscoverServiceConnectionPoint` start jobs with `do..while($true)`.
The `$AUTODISCOVER_SCP_MAX_RETRIES` counter prevents infinite loops.
The filter template is passed as a parameter because script scope is not available in jobs.

### 4. AutoLogon writes plaintext password
`Enable-AutoLogon` writes the password to `HKLM:\...\Winlogon\DefaultPassword`.
`Disable-AutoLogon` removes it on the next login. This is intentional by design.

### 5. `Invoke-WebRequest -SkipCertificateCheck`
Only available in PowerShell 6+. For PS 5.1, a fallback may be needed.

### 6. `$AUTODISCOVER_SCP_FILTER` as template
The filter contains `{0}` as a placeholder for the server name:
```powershell
$LDAPSearch.Filter = $AUTODISCOVER_SCP_FILTER -f $Name
```

### 7. Error handling in catch blocks
Always use `$_.Exception.Message`, not `$Error[0].ExceptionMessage`
(can be overwritten by concurrent errors).

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

---

## Open Items / Possible Next Steps

- [ ] `Invoke-WebRequest -SkipCertificateCheck` PS 5.1 fallback
- [ ] Audit all `$Error[0]` occurrences (outside of catch blocks)
- [ ] Pester tests for key helper functions
- [ ] Reduce parameter block redundancy (many parameters with 4x identical `[parameter()]` attributes)
- [ ] Make `Get-SetupTextVersion` more efficient (direct hashtable lookup)
- [ ] Configure RSS queues to match physical core count
- [ ] Set MaxConcurrentAPI for Kerberos authentication optimization
- [ ] Enable LSA Protection (RunAsPPL) — test for Exchange compatibility first
