---
name: v5.1 Feature Plan — Status
description: Implementation status v5.1 — completed features, rounds 9-11 improvements
type: project
originSessionId: f6bf005e-9747-4843-b86d-7b7dc84abbeb
---
## v5.1 — Complete as of 2026-04-15

All v5.1 open items resolved in sessions 2026-04-14 and 2026-04-15 (Rounds 9-11).

### Completed Features

| # | Feature | Status |
|---|---------|--------|
| A | ParameterSets R + T, State-Keys, MAX_PHASE logic, Phase-Switch branching | ✅ done |
| B | `Get-ValidatedCredentials` + Test-Preflight credential block replaced | ✅ done |
| C | `Show-InstallationMenu` (modes 1-5, toggles A-R, greying, RawUI.ReadKey + Read-Host fallback) | ✅ done |
| D | Recipient Management: `Install-RecipientManagementPrereqs`, `Install-RecipientManagement`, `New-RecipientManagementShortcut`, `Invoke-RecipientManagementADCleanup` | ✅ done |
| E | Management Tools: `Install-ManagementToolsPrereqs`, `Install-ManagementToolsRuntimePrereqs`, `Install-ManagementToolsOnly` | ✅ done |
| F | `Build.ps1` — PS2Exe wrapper, auto-install PS2Exe, version detection | ✅ done |
| G | Header + docs — author st03ps, credits Michel de Rooij, `$ScriptVersion = '5.1'`, `.PARAMETER` block synchronized | ✅ done |
| H | `Install-PendingWindowsUpdates` (PSWindowsUpdate + WUA COM fallback) + `$ExchangeSUMap` + `Get-/Install-ExchangeSecurityUpdate` | ✅ done |
| I | `-ConfigFile`: loads all parameters from .psd1 (Import-PowerShellDataFile); interactive menu skipped | ✅ done |
| J | `Write-PhaseProgress`: Id 0 = overall phase (1-6), Id 1 = Phase 5 step counter (~25 steps) | ✅ done |
| K | `Enable-LSAProtection`: RunAsPPL=1, Exchange 2019 CU12+/SE, takes effect after reboot | ✅ done |
| L | `Set-MaxConcurrentAPI`: Netlogon MaxConcurrentApi = logical core count (min 10) | ✅ done |
| M | `Enable-RSSOnAllNICs`: NumberOfReceiveQueues = physical core count (HealthChecker req.) | ✅ done |
| N | `Invoke-WebDownload`: PS 5.1-compatible download helper (WebClient fallback + cert bypass) | ✅ done |
| O | Pester tests: `Install-Exchange15.Tests.ps1` covering Get-SetupTextVersion, Get-OSVersionText, Get-FullDomainAccount, ExchangeSUMap structure | ✅ done |
| P | Parameter block: `ValueFromPipelineByPropertyName = $false` removed (~60 lines saved) | ✅ done |
| Q | `$ExchangeSUMap`: direct download.microsoft.com URLs, .exe filenames, all supported CUs (SE RTM, 2019 CU13-15, 2016 CU23); `Install-ExchangeSecurityUpdate` .msp→.exe fixed | ✅ done |

### Round 9 Bugfixes (2026-04-15)

| Bug | Fix |
|-----|-----|
| `$Error[0].ExceptionMessage` in autodiscover background job catch blocks | `$_.Exception.Message` |
| `Write-Host` in `Enable-MSExchangeAutodiscoverAppPool` | `Write-MyOutput` |
| Missing `$InstallWindowsUpdates` mapping from menu result block | Added assignment after `$NoNet481` |
| IIS health check Invoke-WebRequest without PS 5.1 compat | PS 5.1/6+ split using WebClient.DownloadString |
| `Get-WindowsFeature Bits` without .Installed in Cleanup | `(Get-WindowsFeature -Name 'Bits').Installed` |

### Round 10-11 Performance/Security (2026-04-15)

| Change | Detail |
|--------|--------|
| `Clear-DesktopBackground` | Replaced Add-Type/C# with Registry + RUNDLL32 (3-10s faster per phase) |
| `Get-DetectedFileVersion` | Replaced Get-Command with FileVersionInfo API (no PATH/module overhead) |
| `Get-SetupTextVersion` | Direct hashtable lookup first (O(1) for exact CU builds) |

**Why:** v5.1 extended the script from pure Exchange installer to comprehensive Exchange management toolkit with interactive menu, multiple install modes, Windows Update + Exchange SU automation, and Build.ps1 for .exe compilation.
**How to apply:** All existing parameter-based calls behave identically to v5.01. New features activate via new parameters or interactive menu.
