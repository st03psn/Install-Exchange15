# Modularization + EXpress Rename (next release / v1.0)

**Motivation:** The script has reached 10,253 lines. Every change forces editor/review/Claude to load the entire monolith. Modularization reduces load cost and mirrors the already-logical groupings (Logging, Package, AD, Hardening, Reports, OpenXML, RemoteQuery, …) in the file structure.

**Why:** Development and review become drastically easier; the release artifact remains a single `.ps1` (download link unchanged) and stays PS2Exe-compatible.

**How to apply:** Ship together with the EXpress rename (v1.0) — both are invasive structural changes, bundling avoids duplicate migration work. The jump to 1.0 justifies both changes.

## Core Idea: Hybrid (Dev Dot-Source + Release Merge)

- `src/*.ps1` = single source of truth (numeric prefixes enforce load order).
- Entry script `EXpress.ps1` contains a `process{}` block with a `#region SOURCE-LOADER` that dot-sources `src/*.ps1` in dev mode.
- `tools/Merge-Source.ps1` replaces that loader with concatenated content → produces release `EXpress.ps1` (PS2Exe-compatible, single-file).
- `Build.ps1` calls Merge-Source before PS2Exe.
- The release `EXpress.ps1` is committed (raw-GitHub-download compatibility preserved).

## Module Layout (src/*.ps1)

| File | Functions |
|---|---|
| `00-Constants.ps1` | `$ScriptVersion`, `$ERR_*`, `$WS2016_MAJOR`, `$WS2019_PREFULL`, `$WS2022_PREFULL`, `$WS2025_PREFULL`, `$EX*_SETUPEXE_*`, `$NETVERSION_48/481`, `$AUTODISCOVER_SCP_*`, `$POWERPLAN_*`, `$WU_DOWNLOAD_TIMEOUT_SEC`, `$ERR_SUS_NOT_APPLICABLE` |
| `05-State.ps1` | Save-State, Restore-State |
| `10-Logging.ps1` | Write-ToTranscript, Write-MyOutput/Warning/Error/Verbose/Debug, Write-PhaseProgress |
| `15-Helpers.ps1` | Set-RegistryValue, Invoke-NativeCommand, Get-OSVersionText, Get-SetupTextVersion, Get-DetectedFileVersion, Get-PSExecutionPolicy, Add-BackgroundJob, Stop-BackgroundJobs |
| `20-Package.ps1` | Get/Install/Test-MyPackage, Invoke-Process/Extract/WebDownload, Resolve-SourcePath, Get-VCRuntime, Set/Remove-NETFrameworkInstallBlock |
| `25-AD.ps1` | Get-Forest*/RootNC/ConfigurationNC/FunctionalLevel, Get-ExchangeOrganization/DAGNames/ForestLevel/DomainLevel, Test-DomainNativeMode, Test-ExchangeOrganization, New-LDAPSearch, Get-ADSite |
| `30-Credentials.ps1` | Test-Admin, Test-ADGroupMember (+Schema/Enterprise), Test-ServerCore, Test-RebootPending, Get-FullDomainAccount, Get-CurrentUserName, Test-Local/Credentials, Get-ValidatedCredentials, Enable/Disable-AutoLogon, Enable-RunOnce, Enable/Disable-UAC, Enable/Disable-IEESC, Disable/Enable-OpenFileSecurityWarning |
| `35-Exchange.ps1` | Import-ExchangeModule, Reconnect-ExchangeSession, Initialize-Exchange, Install-Exchange15_, Clear/Set-AutodiscoverServiceConnectionPoint, Test-ExistingExchangeServer, Get-LocalFQDNHostname, Get-ExchangeServerObjects, Set-EdgeDNSSuffix |
| `40-Preflight.ps1` | Test-Preflight, New-PreflightReport, Install-WindowsFeatures, Get-FFLText, Get-NetVersionText, Get-NETVersion |
| `45-ServerConfig.ps1` | Export-SourceServerConfig, Import-ServerConfig, Test-DBLogPathSeparation, Wait-ADReplication, Register-ExchangeLogCleanup, Write-Log |
| `50-Connectors.ps1` | Add-ServerToSendConnectors, Install-AntispamAgents, New-AnonymousRelayConnector |
| `55-Security.ps1` | Invoke-EOMT, Set-HSTSHeader, Import-ExchangeCertificateFromPFX, Set-VirtualDirectoryURLs |
| `60-DAG-HC.ps1` | Join-DAG, Invoke-HealthChecker, Invoke-SetupAssist, Test-AuthCertificate, Test-DAGReplicationHealth, Test-VSSWriters, Test-EEMSStatus, Get-ModernAuthReport |
| `65-RemoteQuery.ps1` | Get-RemoteServerData, Invoke-RemoteQueryWithPrompt |
| `70-ReportData.ps1` | Get-OrganizationReportData, Get-ServerReportData, Get-InstallationReportData |
| `72-ReportHtml.ps1` | New-InstallationReport (+ Format-Badge, New-HtmlSection, Get-SecRegVal, Format-RefLink) |
| `74-OpenXml.ps1` | Invoke-XmlEscape, New-Wd{Heading,Paragraph,Bullet,Code,Table,DocumentXml,File}; **F24:** template-merge helpers (Open-WdTemplate, Set-WdContentControl, Merge-WdBodyAnchor, Test-WdTemplate) |
| `76-InstallDoc.ps1` | New-InstallationDocument (F22) + local helpers (Mask-Ip, Mask-Val, SafeVal, L, Lc, Get-SecReg, Format-RemoteSysRows); **F24:** `-TemplatePath` parameter, body-anchor injection loop |
| `78-RBAC.ps1` | Get-RBACReport |
| `80-Optimizations.ps1` | Get-OptimizationCatalog, Invoke-Single/ExchangeOptimizations + Draw-OptimizationMenu; all hardening functions: Enable-HighPerformancePowerPlan, Disable-NICPowerManagement, Set-Pagefile, Set-TCPSettings, Disable-SMBv1/WindowsSearchService/UnnecessaryServices/ShutdownEventTracker/WDigestCredentialCaching/HTTP2/TCPOffload/UnnecessaryScheduledTasks/ServerManagerAtLogon/CredentialGuard/SSL3/RC4/MRSProxy/NetBIOSOnAllNICs/SSLOffloading, Test-DiskAllocationUnitSize, Set-CRLCheckTimeout/LmCompatibilityLevel/MaxConcurrentAPI/CtsProcessorAffinityPercentage/NodeRunnerMemoryLimit/IPv4OverIPv6Preference/MAPIEncryptionRequired/SchannelProtocol/NetFrameworkStrongCrypto/TLSSettings, Enable-RSSOnAllNICs/LSAProtection/SerializedDataSigning/MAPIFrontEndServerGC/ECC/CBC/AMSI/IanaTimeZoneMappings/ExtendedProtection/RootCertificateAutoUpdate/WindowsDefenderExclusions, Start-DisableMSExchangeAutodiscoverAppPoolJob, Enable-MSExchangeAutodiscoverAppPool |
| `85-WU-SU.ps1` | Install-PendingWindowsUpdates, Get-LatestExchangeSecurityUpdate, Get-InstalledExchangeBuild, Get-LatestSUBuildFromHC, Install-ExchangeSecurityUpdate |
| `88-RecipientMgmt.ps1` | Test-IsClientOS, Install-RecipientManagementPrereqs/RecipientManagement, New-RecipientManagementShortcut, Invoke-RecipientManagementADCleanup, Install-ManagementToolsPrereqs/RuntimePrereqs/Only |
| `90-Cleanup.ps1` | Cleanup, LockScreen, Clear-DesktopBackground |
| `95-Menu.ps1` | Show-InstallationMenu, Draw-Menu, Read-MenuInput, Get-DynamicDisabled, Write-MenuLine, Get-CfgValue |
| `99-Main.ps1` | Phase orchestration 0–6 (everything between the last function definition and `exit $ERR_OK`), including Step-P5 helper |

## Tools

- `tools/Merge-Source.ps1` (new): reads entry script, finds `#region SOURCE-LOADER` … `#endregion`, replaces with concatenated `src/*.ps1` content (sorted), writes UTF-8 without BOM, 4-space indentation for `process{}` context. Output: `dist/EXpress.ps1` (or root `EXpress.ps1` directly).
- `Build.ps1` (extended): new `-SkipMerge` switch; otherwise calls Merge before PS2Exe.
- CI guard: pre-commit hook or GitHub Action running merge and `git diff --exit-code` on the release `.ps1` — prevents committing a `src/` change without updating the merged release build.

## Migration Sequence

1. Create `src/` + `tools/Merge-Source.ps1`, insert `#region SOURCE-LOADER` into entry script. Hash verification: run `tools/Merge-Source.ps1`, SHA256 of output must be identical to current `Install-Exchange15.ps1` (byte-exact — no functional change).
2. Extract module by module (1 commit per module). After each step: merge + `tools/Parse-Check.ps1` + `tools/Test-ScriptSanity.ps1` + Autopilot smoketest.
3. Rename entry script: `Install-Exchange15.ps1` → `EXpress.ps1` (v1.0 rename checklist).
4. Switch all generated filenames to `{PC}_EXpress_{Tag}_...` (State, Preflight, Report, Document, RBAC, Config, LogCleanup).
5. `$ScriptVersion = '1.0'`, README/CLAUDE/docblock updates, `docs/index.html` announcement banner.
6. Enable CI guard, decide on `dist/` policy (committed vs. release asset — recommendation: committed for raw.githubusercontent downloads).

## F24 — Installation-Document Template (bundled with v1.0)

Convert F22 from fully code-driven OpenXML to a **style-shell + dynamic body** hybrid. Lands naturally with modularization because the OpenXML engine is already being extracted into `src/74-OpenXml.ps1`.

**New template files (repo-committed):**
- `templates/Exchange-installation-document-DE.docx`
- `templates/Exchange-installation-document-EN.docx`

Contents of each template:
- Cover page with `[LOGO]` placeholder and SDT content controls for `org_name`, `server_name`, `install_date`, `scenario`
- Header/footer with classification stamp (default `INTERN`, overridable via SDT)
- `word/styles.xml` with Heading1–3, Code, Callout, Table styles
- `word/theme1.xml` with corporate colour palette
- `word/numbering.xml` for bullet / numbered lists
- Empty body-anchor SDTs with deterministic tags: `body_chapter_1` … `body_chapter_16`

**Generator:** extend `tools/Build-ConceptTemplate.ps1` (rename to `tools/Build-DocumentTemplates.ps1` to cover both F23 concept + F24 installation templates) or add a parallel `tools/Build-InstallationTemplate.ps1`.

**Runtime flow (`New-InstallationDocument` post-F24):**
1. Load template ZIP (or `-TemplatePath` override)
2. Populate cover-page SDTs with scope data from `Get-InstallationReportData`
3. For each chapter: generate the chapter body XML via existing `New-Wd*` helpers, inject into the matching `body_chapter_N` anchor SDT
4. Write output ZIP

**New helpers in `src/74-OpenXml.ps1`:**
- `Open-WdTemplate -Path <path>` — loads template as ZIP, returns handle
- `Set-WdContentControl -Handle -Tag -Value` — replace SDT content
- `Merge-WdBodyAnchor -Handle -AnchorTag -BodyXml` — inject generated XML into anchor SDT
- `Test-WdTemplate -Path` — validates all expected tags present, returns pass/fail + missing list (called before generation; fail fast with clear error)

**New parameter (`EXpress.ps1`):** `-TemplatePath <path>` — customer-specific template override.

**Risks:**
- Template-schema brittleness: user editing the .docx could break SDT tags → `Test-WdTemplate` mandatory before every run
- Two languages = two templates to keep in sync on structure changes → `Build-*Template.ps1` generator must be the source of truth; hand-edits in .docx should be limited to styling, not structure
- Styles in template must not conflict with runtime-generated content → runtime only references styles, never defines new ones

**Effort:** 1–2 days total. Splits cleanly:
1. Template generator + committed .docx files (0.5 d)
2. Engine helpers `Open-WdTemplate` / `Set-WdContentControl` / `Merge-WdBodyAnchor` / `Test-WdTemplate` (0.5 d)
3. `New-InstallationDocument` refactor: generate chapter XML fragments instead of full doc, inject into anchors (0.5 d)
4. `-TemplatePath` param + validation + testing with a customized template (0.25 d)

**Sequencing with v1.0 release:** do F24 *after* the src/*.ps1 extraction is complete (so the engine already lives in its own module) but *before* the v1.0 tag. Commit order:
1. Modularization complete → merge verification green
2. F24 template generator + committed templates
3. F24 engine helpers
4. F24 `New-InstallationDocument` refactor
5. EXpress rename (all file paths, nomenclature, docblock)
6. Cut v1.0

## Open Decisions

- **Release `.ps1` path:** root (`EXpress.ps1`) or `dist/EXpress.ps1`? → recommend root, so existing downloads keep working. `src/` sits alongside it in the repo root.
- **Indentation in `src/*.ps1`:** write modules without leading indent; the merge step indents by 4 spaces. Easier to edit, Parse-Check works on both forms.
- **`99-Main.ps1` boundary:** clean split between "function defined" and "phase-N code". If Main grows too large, switch to one file per phase (`99-Phase0.ps1` … `99-Phase6.ps1`).

## Risks

- **Scope:** functions rely on being defined inside the `process{}` scope of the entry script (access to `$State`, `$Credentials`, constants). Dot-source preserves that semantics; the merged variant is identical. No scope-logic refactoring required.
- **Line numbers:** all existing line-number references in CLAUDE.md / commit messages become invalid. The first post-split commit should document that those references point to the pre-split version.
- **`tools/Fix-*.ps1`:** operate on the merged file — still compatible. Optionally adapt them to work on `src/*.ps1` as well (nice-to-have, not blocking).
