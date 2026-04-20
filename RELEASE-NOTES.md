# EXpress — Release Notes

Full optimization and feature history. See `README.md` for user-facing changelog.

---

## v5.69 (2026-04-20)

- `$State['ConfigDriven']` added to state: `$true` only when started via `-ConfigFile` (headless); `$false` for interactive menu start
- Mode display (`Mode: Autopilot (fully automated)` / `Mode: Copilot (interactive)`) now driven by `ConfigDriven` instead of `Autopilot`; fixes incorrect "Autopilot" label when the auto-reboot toggle was on in the interactive menu

---

## v5.68 (2026-04-20)

- `Disable-UnnecessaryServices` — disables Print Spooler (Spooler), Fax, Secondary Logon (seclogon), Smart Card (SCardSvr); per CIS Benchmark / NSA / DISA STIG; Phase 5 hardening
- `Disable-ShutdownEventTracker` — sets `ShutdownReasonOn`/`ShutdownReasonUI = 0` under `HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Reliability`; redundant with Event IDs 1074/6006/6008; avoids blocking Autopilot reboots

---

## v5.6 (2026-04-19)

**Bugfixes (v5.51):**
- `Get-ValidatedCredentials`: PSObject cast fix + `Read-Host` fallback for PS2Exe
- `Start-DisableMSExchangeAutodiscoverAppPoolJob`: `Test-Path` replaces `Get-WebAppPoolState` (PathNotFound in background job)
- `Restart-Service` W3SVC/WAS/Transport: `-WarningAction SilentlyContinue` (suppressed repetitive warnings)
- `Install-AntispamAgents`: `3>$null` on install script; `-WarningAction SilentlyContinue` on agent cmdlets
- `Set-VirtualDirectoryURLs`: `-Confirm:$false` on all `Set-*VirtualDirectory` calls
- `Register-ExchangeLogCleanup`: `FlushInputBuffer` isolated in own `try/catch`
- VC++ 2012 install condition extended to all Exchange versions
- `ExchangeSUMap` KB5074992: Windows Update Catalog CAB URL added

**Features (v5.6):**
- `Reconnect-ExchangeSession`: reconnects Exchange PS session after IIS restart (waits up to 90s)
- `New-AnonymousRelayConnector`: both internal AND external connector always created; RFC 5737 placeholder `192.0.2.1/32` when no IPs given
- `New-InstallationReport`: Section 8 (RBAC) + Section 9 (Installation Log); Autodiscover SCP in vdir table; UAC re-enabled before report; HC "skipped" vs "not found" distinction
- Reports subfolder: all reports and logs moved to `<InstallPath>\reports\`
- `Register-ExchangeLogCleanup`: prompt skipped in AutoPilot mode

---

## v5.5 (2026-04-19)

- ISO remounted only for phases 1–3; dismounted at end of phase 4
- `Test-Preflight`: heavy checks skipped for phase ≥ 5
- `Set-VirtualDirectoryURLs`: MAPI `-InternalAuthenticationMethods` wrapped in separate try/catch; OWA `-LogonFormat UPN`
- `Get-RBACReport`: format string crash fixed in catch block
- `Import-ExchangeModule`: no longer prints "already loaded" warning
- `Install-AntispamAgents`: Phase 6
- `Add-ServerToSendConnectors`: interactive Y/N prompt; Phase 6
- `Register-ExchangeLogCleanup`: interactive folder prompt (default `C:\#service`, 2-min timeout)
- `New-AnonymousRelayConnector`: `-AuthMechanism Tls`, `-ProtocolLoggingLevel Verbose`
- `Import-ExchangeCertificateFromPFX`: wildcard detection; non-wildcard also enables IMAP/POP

---

## v5.4 (2026-04-18)

- `New-InstallationReport`: 6-section HTML report (Parameters, System, AD, Exchange, Security, Performance); sidebar nav, status badges, print CSS
- PDF export via Edge headless (`--print-to-pdf`)
- `-SkipInstallReport` switch
- `$VerbosePreference = 'Continue'` unconditionally

---

## v5.3 (2026-04-17)

- `Add-BackgroundJob`: prunes Completed/Failed/Stopped before append
- `New-LDAPSearch`: helper replacing 4 duplicated DirectorySearcher blocks
- Registry idempotency: 7 call sites via `Set-RegistryValue`
- BSTR zeroing: `ZeroFreeBSTR` in `Test-Credentials` and `Enable-AutoLogon`
- Exit code checks: RUNDLL32 + powercfg warn on non-zero
- Pester: 45 → 54 tests

---

## v5.2 (2026-04-16–17)

- `Set-HSTSHeader`: HSTS on OWA/ECP (Phase 5, conditional on CertificatePath)
- `Test-DBLogPathSeparation`: warn when DB and log share a volume (Phase 6)
- DB sizing guidance in pre-flight HTML report
- `Invoke-EOMT`: EOMT CVE mitigation, `-RunEOMT`, Phase 5
- `Register-ExchangeLogCleanup`: daily scheduled task, `-LogRetentionDays`
- `Wait-ADReplication`: `-WaitForADSync`, repadmin, Phase 3
- `-StandaloneOptimize` ParameterSet 'O': single-phase post-install run
- `New-AnonymousRelayConnector`: `-RelaySubnets`, `-ExternalRelaySubnets`, Phase 6
- `Get-OptimizationCatalog` / `Invoke-ExchangeOptimizations`: 10-entry data-driven menu (Phase 5)
- `Invoke-SetupAssist`: Phase 4 failure handler
- `Set-VirtualDirectoryURLs`: `-Namespace` parameter
- `Get-RBACReport`: 10 role groups → UTF-8 file
- ValidatePattern regex fixes, named constants, dead code removed

---

## v5.1 (2026-04-10–15)

- `Show-InstallationMenu`: interactive console menu (modes 1–5, letter toggles, RawUI.ReadKey + Read-Host fallback)
- `Get-ValidatedCredentials`: credential retry loop (max 3 attempts)
- `Install-PendingWindowsUpdates`: PSWindowsUpdate + WUA COM fallback; per-update Y/N/A/S prompt; background job with polling + [C] cancel
- `Get-LatestExchangeSecurityUpdate` / `Install-ExchangeSecurityUpdate`: built-in `$ExchangeSUMap`, BITS download
- `Build.ps1`: PS2Exe compilation to .exe
- `-ConfigFile`: load all parameters from .psd1; `config.psd1` auto-detection
- `Write-PhaseProgress`: Write-Progress wrapper with PS2Exe fallback
- `Reconnect-ExchangeSession`: reconnect after IIS restart
- ParameterSet 'R' (`-InstallRecipientManagement`): EMT install on Server or Client OS
- ParameterSet 'T' (`-InstallManagementTools`): `setup.exe /roles:ManagementTools`

---

## v5.0 (2026-03-22)

- `New-PreflightReport`: HTML pre-flight report (`-PreflightOnly`)
- `Export-SourceServerConfig` / `Import-ServerConfig`: swing migration config copy
- `Import-ExchangeCertificateFromPFX`: PFX import, IIS+SMTP binding
- `Join-DAG`: automated DAG membership
- `Invoke-HealthChecker`: CSS-Exchange HealthChecker integration
- System Restore checkpoints before each phase (`-NoCheckpoint`)
- `Set-RegistryValue`: idempotency guard

---

## Rounds 1–7 (2025-03-21 – 2026-04-08) — Code Quality & Modernisation

- WMI → CIM migration (`Get-WmiObject` → `Get-CimInstance`, `.Put()` → `Set-CimInstance`)
- `Write-ToTranscript` helper; unified `Write-My*` functions
- `WebClient`/`ServerCertificateValidationCallback` → `Invoke-WebRequest -SkipCertificateCheck`
- `Invoke-Extract`: COM shell.application → `Expand-Archive`
- `Enable-RunOnce`: `$PSHome\powershell.exe` → `(Get-Process -Id $PID).Path`
- Autodiscover SCP background jobs: retry limit + timeout (was infinite loop)
- `$WS2025_PREFULL` corrected to `10.0.26100` (was `10.0.20348` = WS2022)
- Security: `Disable-SMBv1`, `Disable-WDigestCredentialCaching`, `Disable-HTTP2`, `Set-CRLCheckTimeout`
- Security: `Disable-CredentialGuard`, `Set-LmCompatibilityLevel`, `Enable-SerializedDataSigning`, `Enable-LSAProtection`
- Performance: `Enable-RSSOnAllNICs` + `NumberOfReceiveQueues`, `Set-MaxConcurrentAPI`, `Disable-TCPOffload`
- Auto-elevation via `Start-Process -Verb RunAs`; auto-unblock Zone.Identifier on source files
- `$AUTODISCOVER_SCP_FILTER`, `$AUTODISCOVER_SCP_MAX_RETRIES`, `$ERR_SUS_NOT_APPLICABLE`, `$POWERPLAN_HIGH_PERFORMANCE` constants introduced
