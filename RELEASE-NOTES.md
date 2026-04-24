# EXpress — Release Notes

Full version history for EXpress. See [README.md](README.md) for overview and quick start.

**Versioning scheme:** `MAJOR.MINOR` = feature release; `MAJOR.MINOR.PATCH` = bugfix / maintenance release on top of the matching feature version.

---

## v1.1.8 (2026-04-24) — bugfix / enhancement release

- **Word Installation Document: phantom certificates** — Certificates with `NotAfter = DateTime.MinValue` (year 0001) leaked through into the Word doc certificate table as in the HTML report. Filtered at source in `Get-ServerReportData` (Thumbprint empty or `NotAfter ≤ 1970-01-01`) and belt-and-suspenders in `New-InstallationDocument`.
- **Word Installation Document: readable Exchange version** — Sections 5.x.1 (per-server identity) and Section 7 (installation table) now show the friendly version string (e.g. "Exchange Server 2019 Cumulative Update 15") alongside the build number. Evaluation/Trial editions show a prominent warning note.
- **Word Installation Document: Extended Protection per VDir** — Section 5.x.4 (Virtual Directories) table now includes an EP column with `None`/`Allow`/`Require` per service (integer values 0/1/2 normalized to enum name).
- **Word Installation Document: hardware type and time zone** — Section 5.x.2 (System Details) now shows hardware type (VMware/Hyper-V/KVM/Physical with manufacturer+model), time zone (StandardName + Caption), and uptime in days alongside the last boot timestamp.
- **Word Installation Document: NIC driver details** — Section 5.x.2 now includes NIC driver version, driver date, and link speed rows from `Get-NetAdapter`.
- **Word Installation Document: VC++ runtime table** — Section 7 now includes a dedicated table of installed Visual C++ Redistributables (package name, version, install date) for HC compliance verification.
- **Word Installation Document: TLS cipher suite inventory** — Section 8.1 now includes a table of active TLS cipher suites from `Get-TlsCipherSuite` (suite name, key exchange, hash, algorithm).
- **Word Installation Document: HSTS and Download Domains** — Section 8.4 now includes HSTS (`Strict-Transport-Security` IIS header value) and CVE-2021-1730 Download Domains status (`EnableDownloadDomains` + configured domain).
- **HTML report: hardware type, time zone, readable Exchange version** — Server detail section now shows hardware type (VMware/Hyper-V/Physical), time zone, and friendly Exchange version string alongside the build number.
- **HTML report: per-VDir Extended Protection table** — Security section now shows a dedicated inline table with EP status per virtual directory (OWA/ECP/EWS/OAB/EAS/MAPI/Autodiscover).
- **HTML report: HSTS and Download Domains** — Security section now shows the IIS `Strict-Transport-Security` header value and CVE-2021-1730 `EnableDownloadDomains` org flag + configured domain.

---

## v1.1.7 (2026-04-24) — bugfix release

- **HTML report: phantom certificates** — `Get-ExchangeCertificate` returns entries with `NotAfter = DateTime.MinValue` (year 0001) and empty Subject/Thumbprint for internal/orphan certificate objects. These were rendered as "Expires -739729d!" rows with blank cells and inflated the "expiring within 90 days" summary count. Phantom entries are now filtered before processing.
- **HTML report: certificate expiry used `.Days` instead of `.TotalDays`** — `TimeSpan.Days` is the days-component of the timespan, not the total number of days. A certificate with 1799 days remaining showed e.g. 14 days if the month/year components were non-zero. Fixed to `[Math]::Floor(.TotalDays)`, consistent with the Auth cert and InstallDoc calculations.
- **HTML report: Root CA Auto-Update shows empty value** — When the `DisableRootAutoUpdate` policy registry key is absent (factory default — auto-update allowed), `$rootAU` is `$null` and the cell rendered as "DisableRootAutoUpdate = ". Now shows `(not set — default enabled)` when the key is absent.
- **HTML report: NetBIOS count null in PS 5.1** — `(pipeline | Where-Object {}).Count` returns `$null` (not 0) in PowerShell 5.1 when zero items match, causing the "X of Y NICs disabled" cell to show " of 1 NICs disabled". Fixed in the registry-fallback rewrite from v1.1.6 which uses `$nbDisabled = 0` + increment and is not affected by this PS 5.1 edge case.
- **IANA timezone log entry misleading** — `Register-ExecutedCommand` was called unconditionally before `Enable-IanaTimeZoneMappings`, so `Set-OrganizationConfig -UseIanaTimeZoneId $true` appeared in the install log even when the function skipped it (property already true or not available). Command registration moved inside the function, only logged when actually executed.

---

## v1.1.6 (2026-04-24) — bugfix release

- **CVE-2021-1730 Download Domains incomplete** — `Set-OrganizationConfig -EnableDownloadDomains $true` was never called; only the OWA VDir hostnames were set. Without the org-level flag Exchange ignores the VDir setting entirely, so HealthChecker correctly flagged the mitigation as missing even when a Download Domain was configured.
- **PowerShell VDir InternalUrl** — `Set-VirtualDirectoryURLs` was setting both InternalUrl and ExternalUrl to `https://`. Exchange internal service-to-service communication relies on the http InternalUrl by default; changing it to https can break services. Now only ExternalUrl is set.
- **Installation report: NetBIOS shown as not disabled** — `Disable-NetBIOSOnAllNICs` calls `SetTcpipNetbios(2)` which may return code 1 (pending reboot), leaving the live WMI/CIM value stale. The HTML report now cross-checks the registry (`NetBT\Parameters\Interfaces\Tcpip_<GUID>\NetbiosOptions`) as fallback, so pending-reboot changes are reflected correctly.
- **Installation report: OWA Extended Protection shown as not set** — `ExtendedProtectionTokenChecking` can be deserialized from AD as an integer (0/1/2) rather than the enum string (None/Allow/Require), causing the badge check to fall through. Integer values are now normalized before comparison.
- **Installation document: certificate expiry wrong (e.g. 14d instead of 1799d)** — `(NotAfter - Get-Date).Days` returns only the days-component of the TimeSpan (e.g. 14 for a cert expiring in 4 years 0 months 14 days), not the total number of days. Fixed to `[Math]::Floor(.TotalDays)`, consistent with the Auth cert expiry calculation.

---

## v1.1.5 (2026-04-24) — docs / maintenance release

- **Docs: menu screenshots** — Five terminal screenshots of the Copilot menu (main menu, mode selection, Advanced Configuration pages 1/3, 2/3, 3/3) embedded as base64 data URIs in `docs/index.html` in a new tabbed "Interactive Copilot" section.
- **Docs: Word doc mockup** — "HealthChecker" nav item corrected to "Open Items"; HealthChecker belongs in the HTML report section, not the Word document mockup.

---

## v1.1.4 (2026-04-24) — bugfix release

- **Windows Updates Autopilot** — Security/Critical updates are no longer auto-approved in Autopilot mode. Without explicit opt-in, pending updates are listed and skipped with a warning. New Advanced Configuration toggle `AutoApproveWindowsUpdates` (default off) enables the previous auto-approve behaviour when deliberately set.

---

## v1.1.3 (2026-04-24) — bugfix release

- **Windows Updates** — `[A]=all` option removed from per-update confirmation prompt; each Security/Critical update must be confirmed individually. Autopilot (non-interactive) still approves all automatically.

---

## v1.1.2 (2026-04-24) — bugfix release

- **NuGet auto-install** — `Install-Module -ForceBootstrap` added so the NuGet provider prompt is answered automatically when `Install-PackageProvider` cannot reach its index URI but internet is otherwise available.
- **Autopilot RunOnce path** — `$MyInvocation.MyCommand.Path` inside a dot-sourced module resolves to that module file, not to `EXpress.ps1`. `Enable-RunOnce` therefore wrote the registry key pointing to `modules\99-Main.ps1`; after reboot Windows ran the module directly (all other modules missing → silent failure). `EXpress.ps1` now captures `$EXpressEntryScript` before the dot-source loop; `99-Main.ps1` prefers it.
- **Exchange source default path** — prompt order swapped (Working folder first); default ISO derived from `<InstallPath>\sources\` instead of the script parent dir (which resolved to `modules\` when dot-sourced).
- **Module parse errors / dot-source** — `45-ServerConfig.ps1` here-string was truncated; leaked body removed from `50-Connectors.ps1`; `Show-InstallationMenu` closing `}` moved from `99-Main.ps1` to `95-Menu.ps1` so all 21 modules are independently parseable.
- **PS 5.1 `(if ...)` crashes** — `-ValidateMessage (if ...)` in menu field-edit and `-f ..., (if ...)` in hardening EP command builder replaced with pre-assigned variables.

---

## v1.1 (2026-04-24) — feature release

### Source layout & CI
- `src/` renamed to `modules/` — clearer naming; `Merge-Source.ps1`, `Build.ps1`, `EXpress.ps1` source-loader updated.
- `.github/workflows/merge-guard.yml` — new CI guard: merge + parse + `git diff --exit-code dist/` + Pester on every push/PR touching `modules/`, `EXpress.ps1`, or `dist/`.

### Install-target matrix
- Exchange 2019 CU10–CU14 rejected by preflight (CU15+ required); all dead version-gate and pre-CU14 EP code removed.
- Exchange 2016 CU23 restricted to WS2016 only in OS check.

### Centralized downloads (`SourcesPath`)
- All package downloads (`Get-MyPackage` call sites in 40-Preflight, 50-Connectors, 55-Security, 60-VDir-DAG, 85-WU-SU, 88-RecipientMgmt, 90-Hardening) switched from `InstallPath` to `SourcesPath` (`<InstallPath>\sources\`).
- `SourcesPath` initialized in Phase 0 (preflight); directory created if absent.
- `tools/Get-EXpressDownloads.ps1` — new tool; pre-stages all downloads for air-gapped / proxy-restricted networks; idempotent; `-SkipDotNet` switch.
- `sources/` added to `.gitignore`; `logo.png` moved from `sources/` to `assets/`.
- Logo probe in `New-InstallationDocument` now checks three paths in order: `sources\logo.png` → beside script → `assets\logo.png`.
- `Get-MyPackage` emits actionable offline hint when BITS + WebClient both fail.
- SU interactive countdown and log message now correctly show `SourcesPath` (was `InstallPath`).

### F26: Access Namespace mail config
- New `-MailDomain` parameter (auto-derived from `-Namespace` when omitted).
- `Enable-AccessNamespaceMailConfig` registers root domain as Authoritative Accepted Domain; updates default Email Address Policy primary SMTP; removes `.local`/`.lan` templates.
- Advanced catalog knob `AccessNamespaceMail` (default on when Namespace set) controls the feature.
- `deploy-example.psd1`, menu (mode 1 + mode 6), and `95-Menu.ps1` edit-fields list updated.

### Menu back/edit
- Confirmation screen (Step 4) shows all entered fields; new `E` = edit option re-prompts individual fields with current value as default — no restart required to fix a typo.

### Tools & templates
- `tools/Build-InstallationTemplate.ps1` — fixed stale `src/` → `modules/` path; installation document templates regenerated.
- `tools/Build-ConceptTemplate.ps1` — SE matrix now includes WS2019; EXpress version example updated to `1.1`.
- README: new Tools section documents all `tools/*.ps1` helper scripts.
- `docs/index.html` — versions strip corrected; script names / GitHub links updated; four new feature cards.

---

## v1.0 (2026-04-24) — major release: EXpress rename + modularization

### EXpress rename (formerly Install-Exchange15.ps1)

- Script renamed `Install-Exchange15.ps1` → `EXpress.ps1`; tests renamed accordingly.
- GitHub repository renamed: `st03psn/Install-Exchange15` → `st03psn/EXpress`.
- `$ScriptVersion` jumps to `1.0` (new identity — not a continuation of v5.x numbering).
- State file: `{PC}_State.xml` → `{PC}_EXpress_State.xml` (prevent conflict with legacy state files).
- All generated output files gain `EXpress` as second filename segment: `{PC}_EXpress_{Tag}_...` (Install log, Preflight, Report, RBAC, Config, InstallDoc).
- Menu headers, HTML/DOCX footers, and report branding updated to `EXpress v1.0`.
- `Build.ps1`: now runs `Merge-Source.ps1` automatically before PS2Exe; new `-SkipMerge` switch; output is `EXpress.exe`.
- README: renamed, URL updated, Michel de Rooij acknowledgement section added.

### Modularization

- Source split into 21 `modules/*.ps1` files with numeric load-order prefixes (`00-Constants.ps1` … `99-Main.ps1`).
- Entry script `EXpress.ps1` uses `#region SOURCE-LOADER` for dev-mode dot-sourcing; `tools/Merge-Source.ps1` produces `dist/EXpress.ps1` for release.
- `tools/Parse-Check.ps1` — AST syntax gate; runs against `dist/EXpress.ps1`.
- `dist/EXpress.ps1` is byte-identical to the pre-split monolith (SHA256 verified).
- GitHub Actions workflow `.github/workflows/merge-guard.yml` enforces `modules/*.ps1` ↔ `dist/EXpress.ps1` sync on every push/PR.

### Centralized downloads and tightened install target matrix

- All prerequisite packages (.NET, VC++, UCMA, URL Rewrite, hotfixes, Exchange SUs) and CSS-Exchange scripts (HealthChecker, EOMT, SetupAssist, SetupLogReviewer, ExchangeExtendedProtection, MonitorExchangeAuthCertificate, Add-PermissionForEMT) now land in `<InstallPath>\sources\` instead of `<InstallPath>\`. Pre-staging works automatically: any file already present is reused and not re-downloaded — enables air-gapped / proxy-restricted installs without code changes.
- Install targets tightened to the latest CU per Exchange line. Preflight rejects older Ex2019 CUs (CU10–CU14) as out-of-Microsoft-SU-support. Older CUs remain valid as migration **source servers** via `Export-SourceServerConfig`.
  - Ex2016 CU23 (final) → Windows Server 2016
  - Ex2019 CU15+         → Windows Server 2019 / 2022 / 2025
  - Exchange Server SE   → Windows Server 2019 / 2022 / 2025

---


## v0.8 (2026-04-24) — Advanced Configuration menu + document templates

### F25 — Advanced Configuration menu (`v5.95`, `v5.95.1`)

~55 hardening / tuning / policy switches moved into a new paginated Advanced Configuration menu,
prompted with a 60-second auto-skip after the main menu.

- **Main menu** shrunk to five toggles (A/B/N/R/U/V); advanced switches separated into six
  pages: Security / TLS · Security / Hardening · Performance / Tuning · Exchange Org Policy ·
  Post-Config / Integration · Install-Flow / Debug.
- **`Get-AdvancedFeatureCatalog`** — ordered catalog keyed by feature name; each entry has
  `Category`, `Label`, `Description`, `Default`, optional `Condition` scriptblock (TLS 1.3
  pre-WS2022, Shadow Redundancy without DAG, AnonymousRelay without `RelaySubnets`, etc.).
- **`Test-Feature -Name`** — single enforcement gate; condition evaluated before stored/default
  value — prevents config-file bypass of conditional features. Precedence:
  `$State['AdvancedFeatures']` > catalog default. Unknown names return `$false`.
- **`Invoke-AdvancedConfigurationPrompt`** — 60-second countdown (`Write-Progress -Id 2`);
  silently skipped in Autopilot or when `$State['SuppressAdvancedPrompt']` is set.
- **Config-file parity** — `AdvancedFeatures = @{...}` block in `deploy-example.psd1`; legacy
  top-level keys and CLI switches merged into `$State['AdvancedFeatures']` at startup.
- Bugfix: Windows Update countdown label `Xs remaining` changed to `auto-abort in Xs`. Menu
  language toggle label clarified to show English default.

### F24 — Installation-Document Template (`v5.96`)

Optional custom DOCX template for the Word installation document.

- **`-TemplatePath`** parameter — cover page and header/footer sourced from template; all 18
  chapter sections injected as generated XML.
- **`{{token}}` placeholder replacement** — tokens replaced in every XML part; `{{document_body}}`
  replaces its entire anchor paragraph with the generated chapter XML. Token values are
  XML-escaped before substitution.
- **`Test-WdTemplate -Path -RequiredTags`** — validates that a template DOCX contains all required
  `{{token}}` placeholders; missing required tokens trigger fallback to built-in cover page.
- **`Write-WdFromTemplate -TemplatePath -OutputPath -Tokens`** — copies the template ZIP and
  performs token replacement in all XML parts.
- **`tools\Build-InstallationTemplate.ps1`** — generates starter templates
  `templates\Exchange-installation-document-EN.docx` / `-DE.docx` for customisation.
- Config-file support: `TemplatePath` key in `.psd1` deployment files.
- Fallback to built-in cover page when template is omitted or fails validation — zero regression.

---

## v0.7 (2026-04-23) — Language reform + MEAC hybrid + audit-readiness

### Word document readability (`v5.90`)

- Anti-spam filter table now shows the effective transport-agent pipeline state separately from
  the org-config feature switch, with an explanatory paragraph clarifying why the two differ on
  gateway-fronted Mailbox servers.
- Exchange Online / M365 content promoted from inside Organisation chapter to top-level Section 15
  (placed before Operational Runbooks).
- Wide tables (Receive Connectors, Certificates) use new `-Compact` switch (8pt, ~40% more
  horizontal characters per line). Receive Connector table split into Network and Security halves.

### Default language English; `-German` switch (`v5.91`)

- **Default output: English.** All generated output (Word document, prompts, help) is English
  unless `-German` is set.
- Previous `-Language DE|EN` parameter removed; replaced by single `-German` switch.
- Config-file back-compat: legacy `Language = 'DE'` entries still honoured; map to `-German`.
- `$State['Language']` remains `'EN'`/`'DE'` internally.

### Plain-text credentials in config file (`v5.92`)

- `.psd1` config files may carry `AdminUser` / `AdminPassword` in plain text for zero-touch
  pipelines. Load triggers a box-framed **SECURITY WARNING** to the transcript on every run.
- CLI `-Credentials` takes precedence. Config file must be deleted after install.
- `tools/Check.ps1` quality-gate wrapper moved under `tools/`.

### Hybrid-aware MEAC + AD Split Permissions (`v5.93`)

- **Hybrid coexistence**: `Register-AuthCertificateRenewal` probes `Get-HybridConfiguration`; in
  hybrid mode without `-MEACIgnoreHybridConfig`, task registered in hybrid-safe mode with advisory
  logged. `-MEACNotificationEmail` recommended escape hatch.
- **Auto-generated automation-account password**: 32-char RNG password passed to MEAC at task
  registration; transient — not persisted to state or log.
- **AD Split Permissions**: `-MEACPrepareADOnly` / `-MEACADAccountDomain` parameter set runs
  DA-side prep standalone; `-MEACAutomationCredential` passes Exchange-side credential;
  DPAPI-encrypted in state across Autopilot reboots.
- New CLI: `-MEACAutomationCredential`, `-MEACIgnoreHybridConfig`, `-MEACIgnoreUnreachableServers`,
  `-MEACNotificationEmail`, `-MEACPrepareADOnly`, `-MEACADAccountDomain`.
- New config keys: `MEACIgnoreHybridConfig`, `MEACIgnoreUnreachableServers`, `MEACNotificationEmail`.
- New error code: `ERR_MEACPREPAREAD = 1038`. New helper: `Get-MEACAutomationCredentialFromState`.

### Word document audit-readiness (`v5.94`, `v5.94.1`)

Nine new sections so the Word document doubles as an audit / handover record:

- **Section 1.1** Change-Management placeholder table (change-request, approver, sign-off).
- **Section 4.17** Service-Accounts / RBAC: live `Get-RoleGroupMember` for six Exchange role groups.
- **Section 6.2** Ports / Firewall: 13-row static reference (SMTP, HTTPS, RPC, SMB, GC, WinRM, MAPI).
- **Section 7.2** Security-Update status: Exchange build, Windows build, last boot, SU version.
- **Section 8.8** Compliance mapping: 14 hardening measures vs. CIS Benchmark + BSI IT-Grundschutz.
- **Section 8.9** GDPR checklist: 8 rows (Art. 4 Nr. 2, AVV, TOM, deletion concept, etc.).
- **Section 10.4** Backup evidence placeholder (product, last full/incremental, RPO/RTO).
- **Section 12.2** Monitoring readiness: 9-row go-live checklist.
- **Chapter 16** Acceptance tests: 12 test cases, OWA/ECP/EWS/Autodiscover URLs auto-populated.
- **Section 8.6 / 8.8** SIEM/forensics guidance and recommended source channels (`v5.94.1`).
- **Section 4.7** Retention tags rendered per tag (Type / AgeLimitForRetention / RetentionAction) (`v5.94.1`).
- `New-WdTable`: multi-line cells via `<w:br/>`.

---

## v0.6 (2026-04-23) — Security hardening + MEAC basic + Word doc enrichment

### Defender + network hardening (`v5.86`, `v5.68`)

- `Disable-DefenderRealtimeMonitoring` / `Enable-DefenderRealtimeMonitoring`: disabled at start
  of Phase 1 (before prereq installs), re-enabled at start of Phase 6. Tamper Protection
  (MDE/Intune) registry flip is best-effort; verified via `Get-MpComputerStatus.IsTamperProtected`.
  `$State['DefenderTPPrev']` captures original registry state for exact restore.
- `Disable-LLMNR` and `Disable-MDNS`: closes Responder-class name-spoofing / NTLM-hash-capture
  vectors (CIS L1 §18.5.4.2).
- `Disable-UnnecessaryServices`: Print Spooler, Fax, Secondary Logon, Smart Card disabled
  (CIS Benchmark / NSA / DISA STIG). `Disable-ShutdownEventTracker`: avoids blocking Autopilot reboots.

### MEAC — Auth Certificate auto-renewal (`v5.86`)

`Register-AuthCertificateRenewal` downloads CSS-Exchange `MonitorExchangeAuthCertificate.ps1`
and registers a daily scheduled task. Task runs as SYSTEM; renews 60 days before expiry. Runs
in Phase 6 after `Test-AuthCertificate`; skipped on Edge and management-only installs.

Other Phase 6 improvements (`v5.86`): `New-AnonymousRelayConnector` sets SMTP banner at connector
creation. OWA Extended Protection filtered to Frontend VDir. `Enable-UAC` / `Enable-IEESC`
moved before report generation so security state in reports reflects final hardened config.

### Word document enrichment (`v5.87`, `v5.88`)

- **Section 8 Hardening**: binary registry values rendered as localised `enabled` / `disabled`
  text (`Format-RegBool`); TLS table shows semantic state (active / disabled — no double-negation);
  AMSI and LM Compatibility Level clearly labelled.
- **Section 5.x IMAP/POP3**: `Get-ImapSettings` / `Get-PopSettings` per server.
- **Section 5.x Connectors**: `RequireTLS`, `FQDN`, max size added; Receive Connector table split
  into Network and Security sub-tables. Send Connector extended with same fields.
- **Section 6.1 DNS**: replaced dynamic `Resolve-DnsName` lookup with static template (MX / SPF /
  DKIM / DMARC / Autodiscover placeholders — to be filled post-go-live).
- **Section 8.4 Autodiscover AppPool**: live `Get-WebAppPoolState` replaces configuration intent flag.
- **Section 4.16 Admin Audit Log**: `Get-AdminAuditLogConfig` — retention, mailbox, cmdlets,
  exclusions, log level.
- **Section 9.1 Anti-Spam configuration**: `Get-ContentFilterConfig` / `Get-SenderFilterConfig` /
  `Get-RecipientFilterConfig` / `Get-SenderIdConfig` — four sub-tables.
- **Section 12.1 Crimson Event Log channels**: `Get-WinEvent -ListLog "Microsoft-Exchange*"`.
- Installing user recorded: `$State['InstallingUser']` from `WindowsIdentity.GetCurrent()`.

### Bugfixes (`v5.86.1`, `v5.86.2`, `v5.88.1`–`v5.88.3`)

- **Phase 5→6 spurious reboot** (`v5.86.1`): `Set-IPv4OverIPv6Preference` and
  `Disable-NetBIOSOnAllNICs` no longer set `$State['RebootRequired']`; conditional Phase 5→6
  skip from v0.5 now works as intended.
- **Antispam output** (`v5.86.1`): per-agent table replaced with compact summary list; spurious
  restart-required warnings filtered; pipeline visibility promoted to `Write-MyOutput`.
- **Word doc — nested-array flattening** (`v5.86.2`): `New-WdTable` auto-detects PS 5.1 flattened
  `@(@('a','b'),@('c','d'))` shape and reshapes in place.
- **Word doc — Auth Certificate validity** (`v5.86.2`): `Get-ExchangeCertificate -Thumbprint`
  lookup (+ `Cert:\LocalMachine\My` fallback) replaces missing `NotAfter` property.
- **Word doc — Transport Agents empty Name** (`v5.86.2`): fallback to `.Identity`; all four
  transport scopes iterated (deduped by name).
- **Word doc — TLS semantics** (`v5.86.2`): raw registry values replaced with semantic state strings;
  `ACHTUNG` suffix on detected hardening gaps.
- **Word doc — Serialized Data Signing typo** (`v5.86.2`): reader key corrected to match writer
  (`EnableSerializationDataSigning`).
- **MEAC fixes** (`v5.86.1`, `v5.86.2`, `v5.88.1`): `-Url` changed to `-Uri` in download call;
  parameter `ConfigureScriptViaScheduledTask` renamed to `ConfigureScriptToRunViaScheduledTask`;
  task registration verified via `Get-ScheduledTask`; MEAC task search broadened to cover naming
  variants.
- **PS 5.1 `(if ...)` crashes** (`v5.88.2`): six more occurrences fixed in Word doc;
  `ParenExpressionAst` walker widened in `Test-ScriptQuality.ps1`.
- **`$state` / `$State` shadowing** (`v5.88.3`): `$logState` / `$optState` renames; `SingletonShadow`
  detector added to `Test-ScriptQuality.ps1`.
- Edge Transport guards added to `Enable-AMSI`, `Set-MaxConcurrentAPI`,
  `Set-CtsProcessorAffinityPercentage`, `Set-NodeRunnerMemoryLimit` (not domain-joined).
- `$State['ConfigDriven']` mode-label fix — Autopilot label was shown incorrectly in interactive
  Copilot sessions with auto-reboot toggle on.

---

## v0.5 (2026-04-22) — Org-wide documentation + conditional reboots

### F22 scope expansion — org-wide + remote hardware (`v5.84`)

`New-InstallationDocument` documents the entire Exchange organisation. Three scenario labels:
new server / server addition / ad-hoc inventory. New chapters: Chapter 4 (org-wide config) and
Chapter 5 (all servers with per-server hardware via CIM/WSMan).

New data helpers: `Get-OrganizationReportData`, `Get-ServerReportData`, `Get-InstallationReportData`.
Remote query: `Get-RemoteServerData` uses CIM over WSMan exclusively (no WMI/DCOM).
`Invoke-RemoteQueryWithPrompt`: `[R] Retry / [S] Skip` with 10-minute auto-skip; Autopilot = silent skip.
New parameters: `-DocumentScope All|Org|Local`, `-IncludeServers <Name[]>`.
New tool: `tools/Enable-EXpressRemoteQuery.ps1`.

Three-tier logging (`Write-ToTranscript`), unified file naming `{PC}_{Tag}_{yyyyMMdd-HHmmss}.{ext}`,
credential GUI / `Read-Host` decision via `$env:SESSIONNAME`, log bootstrap before menu draw, and
`SUPPRESSED-ERROR` debug mode delivered alongside as `v5.83`. New dev tools: `Test-ScriptSanity.ps1`,
`Test-ScriptQuality.ps1`, `Fix-IfAsArg.ps1`, `Fix-PhaseNum.ps1`.

### Conditional phase reboots + stability (`v5.85`)

- **Phase 2→3** reboot skipped when `Test-RebootPending` reports nothing pending. Saves a reboot
  on WS2025 + Exchange SE (VC++ / URL Rewrite do not set reboot flags).
- **Phase 5→6** reboot skipped unless `$State['RebootRequired']` set (Exchange SU exit 3010) or
  `Test-RebootPending` signals pending.
- `Test-RebootPending` inspects CBS, WU, `PendingFileRenameOperations`, pending rename, CCM ClientSDK.
- **HealthChecker on Domain Controllers**: `DomainRole -ge 4` detected; clarifying warning emitted
  (local SAM groups absent on DCs — HC finding not actionable).
- **VC++ 2013 URL** updated to `https://aka.ms/highdpimfc2013x64enu` (delivers 12.0.40664;
  old CDN URL delivered 12.0.40660 — flagged outdated by HealthChecker).
- **Antispam install output**: stream-level record routing replaces `Out-Null` + preference gymnastics.
- **VERBOSE console spam after Autopilot resume**: `$VerbosePreference = 'SilentlyContinue'`
  pinned before the first `Get-CimInstance` call.

---

## v0.4 (2026-04-21) — Word Installation Document

### F22 — `New-InstallationDocument` (`v5.82`)

Pure-PowerShell OpenXML engine (no Office/COM required); 15 chapters covering installation
parameters, system details, network, AD, Exchange configuration, hardening, backup readiness,
HealthChecker, monitoring, hybrid, public folders, executed cmdlets, and runbooks.
`CustomerDocument` mode redacts RFC1918 IPs, certificate thumbprints, and passwords.

New parameters: `-NoWordDoc`, `-StandaloneDocument`, `-CustomerDocument`, `-Language`.
Mode 7 ("Standalone Document") generates a document on existing servers without a full install.
`tools/Build-ConceptTemplate.ps1` generates DE + EN concept / approval document templates
(16 chapters: architecture, sizing, security, migration, hybrid, compliance, questionnaire,
approval page).

Bugfixes for the `New-InstallationReport` HTML report (`v5.79`–`v5.81`): `FormatException` on
curly-brace content fixed in all seven affected call sites; transcript encoding auto-detected
(UTF-16 LE / UTF-8); log capped to 2000 lines; call site wrapped in try/catch; HC report renamed
and all three known output prefixes detected.

---

## v0.3 (2026-04-18–21) — Installation reports + post-config features

### HTML Installation Report + PDF (`v5.4`)

- `New-InstallationReport`: 6-section HTML report (Parameters, System, AD, Exchange, Security,
  Performance); sidebar nav, status badges, print CSS.
- PDF export via Edge headless. `-SkipInstallReport` switch.

### Post-config features (`v5.5`, `v5.6`)

- ISO remounted only for phases 1–3; dismounted at end of phase 4.
- `Install-AntispamAgents` and `Add-ServerToSendConnectors` in Phase 6.
- `Register-ExchangeLogCleanup`: interactive folder prompt (2-min timeout); daily scheduled task.
- `Import-ExchangeCertificateFromPFX`: wildcard detection; non-wildcard enables IMAP/POP.
- `Reconnect-ExchangeSession`: reconnects Exchange session after IIS restart (waits up to 90s).
- `New-AnonymousRelayConnector`: both internal AND external connectors created; RFC 5737 placeholder IP.
- `New-InstallationReport` enriched: Section 8 (RBAC) + Section 9 (Installation Log); Autodiscover
  SCP in VDir table; HC "skipped" vs "not found" distinction. Reports moved to `reports\` subfolder.
- **Bugfixes (v5.51)**: `Get-ValidatedCredentials` PSObject cast + `Read-Host` fallback;
  Autodiscover AppPool `Test-Path` replaces `Get-WebAppPoolState`; `Set-*VirtualDirectory`
  `-Confirm:$false`; VC++ 2012 condition extended to all Exchange versions.

### Maintenance releases

**v0.3.1** (2026-04-20):
- `Disable-UnnecessaryServices` (Print Spooler / Fax / Secondary Logon / Smart Card) and
  `Disable-ShutdownEventTracker` added (Phase 5, CIS Benchmark / DISA STIG).
- `$State['ConfigDriven']`: fixes incorrect Autopilot mode label in interactive sessions.
- Broken / stale link fixes in README and HTML report.

**v0.3.2** (2026-04-20–21):
- `Invoke-HealthChecker`: `-BuildHtmlServersReport` call; HC HTML re-added to installation report.
- `Install-AntispamAgents`: `$PSDefaultParameterValues` + `*>&1 | Out-Null` stream suppression.
- `Enable-AMSI`: Exchange SE exception removed.
- `Initialize-Exchange`: returns `$true`/`$false`; `Wait-ADReplication` conditional on PrepareAD.

**v0.3.3** (2026-04-21):
- Edge Transport guards: `Enable-AMSI`, `Set-MaxConcurrentAPI`, `Set-CtsProcessorAffinityPercentage`,
  `Set-NodeRunnerMemoryLimit` (Edge is not domain-joined).
- Exchange SU: `/norestart` removed (not a supported argument — caused immediate abort).
- `Test-AuthCertificate`: null-guard for `$authConfig`.
- `New-AnonymousRelayConnector`: race condition fix; 3-attempt retry fallback.

**v0.3.4** (2026-04-21):
- `Install-ExchangeSecurityUpdate`: `RunOnce` + state persisted before installer launch; per-KB
  skip flag avoids double-install after installer-triggered reboot.
- `New-InstallationReport`: transcript encoding auto-detect; log capped to 2000 lines; call site
  try/catch; `FormatException` on curly-brace content fixed; HC report detection updated.

---

## v0.2 (2026-04-16–17) — Hardening + connector framework

### Hardening + VDir URLs (`v5.2`)

- `Set-HSTSHeader`: HSTS on OWA/ECP (Phase 5, conditional on `CertificatePath`).
- `Test-DBLogPathSeparation`: warn when DB and log share a volume (Phase 6).
- DB sizing guidance in pre-flight HTML report.
- `Invoke-EOMT`: EOMT CVE mitigation; `-RunEOMT`; Phase 5.
- `Register-ExchangeLogCleanup`: daily scheduled task; `-LogRetentionDays`.
- `Wait-ADReplication`: `-WaitForADSync`; `repadmin`; Phase 3.
- `-StandaloneOptimize` ParameterSet `O`: single-phase post-install run.
- `New-AnonymousRelayConnector`: `-RelaySubnets`, `-ExternalRelaySubnets`; Phase 6.
- `Get-OptimizationCatalog` / `Invoke-ExchangeOptimizations`: 10-entry data-driven menu (Phase 5).
- `Invoke-SetupAssist`: Phase 4 failure handler.
- `Set-VirtualDirectoryURLs`: `-Namespace` parameter.
- `Get-RBACReport`: 10 role groups exported to UTF-8 file.
- `ValidatePattern` regex fixes, named constants, dead code removed.

### Code quality (`v5.3`)

- `Add-BackgroundJob`: prunes Completed/Failed/Stopped before append.
- `New-LDAPSearch`: helper replacing 4 duplicated `DirectorySearcher` blocks.
- Registry idempotency: 7 call sites via `Set-RegistryValue`.
- BSTR zeroing: `ZeroFreeBSTR` in `Test-Credentials` + `Enable-AutoLogon`.
- Exit code checks: RUNDLL32 + powercfg warn on non-zero.
- Pester: 45 to 54 tests.

---

## v0.1 (2025-03-21 – 2026-04-15) — Foundation

### Code quality and modernisation (Rounds 1–7, 2025-03-21 – 2026-04-08)

Major WMI-to-CIM migration, logging framework, and security baseline:

- WMI to CIM: `Get-WmiObject` replaced with `Get-CimInstance`, `.Put()` with `Set-CimInstance`
- `Write-ToTranscript` helper; unified `Write-My*` logging functions
- `WebClient`/`ServerCertificateValidationCallback` replaced with `Invoke-WebRequest -SkipCertificateCheck`
- `Invoke-Extract`: COM `shell.application` replaced with `Expand-Archive`
- `Enable-RunOnce`: `$PSHome\powershell.exe` replaced with `(Get-Process -Id $PID).Path`
  (works for both `powershell.exe` and `pwsh.exe`)
- Autodiscover SCP background jobs: retry limit + timeout (was infinite loop)
- `$WS2025_PREFULL` corrected to `10.0.26100` (was `10.0.20348` = WS2022 — critical fix)
- Security: `Disable-SMBv1`, `Disable-WDigestCredentialCaching`, `Disable-HTTP2`,
  `Set-CRLCheckTimeout`, `Disable-CredentialGuard`, `Set-LmCompatibilityLevel`,
  `Enable-SerializedDataSigning`, `Enable-LSAProtection`
- Performance: `Enable-RSSOnAllNICs` + `NumberOfReceiveQueues`, `Set-MaxConcurrentAPI`,
  `Disable-TCPOffload`
- Auto-elevation via `Start-Process -Verb RunAs`; auto-unblock `Zone.Identifier` on source files
- `$AUTODISCOVER_SCP_FILTER`, `$AUTODISCOVER_SCP_MAX_RETRIES`, `$ERR_SUS_NOT_APPLICABLE`,
  `$POWERPLAN_HIGH_PERFORMANCE` constants introduced

### Pre-flight report + source migration (`v5.0`, 2026-03-22)

- `New-PreflightReport`: HTML pre-flight report (`-PreflightOnly`)
- `Export-SourceServerConfig` / `Import-ServerConfig`: swing-migration config copy
- `Import-ExchangeCertificateFromPFX`: PFX import, IIS + SMTP binding
- `Join-DAG`: automated DAG membership
- `Invoke-HealthChecker`: CSS-Exchange HealthChecker integration
- System Restore checkpoints before each phase (`-NoCheckpoint`)
- `Set-RegistryValue`: idempotency guard

### Interactive menu + Autopilot (`v5.1`, 2026-04-10–15)

- `Show-InstallationMenu`: interactive console menu (modes 1–5, letter toggles,
  `RawUI.ReadKey` + `Read-Host` fallback)
- `Get-ValidatedCredentials`: credential retry loop (max 3 attempts)
- `Install-PendingWindowsUpdates`: PSWindowsUpdate + WUA COM fallback; per-update Y/N/A/S prompt
- `Get-LatestExchangeSecurityUpdate` / `Install-ExchangeSecurityUpdate`: built-in `$ExchangeSUMap`,
  BITS download
- `Build.ps1`: PS2Exe compilation to `.exe`
- `-ConfigFile`: load all parameters from `.psd1`; `config.psd1` auto-detection
- `Write-PhaseProgress`: `Write-Progress` wrapper with PS2Exe fallback
- ParameterSet `R` (`-InstallRecipientManagement`): EMT install on Server or Client OS
- ParameterSet `T` (`-InstallManagementTools`): `setup.exe /roles:ManagementTools`

---

## Original Script History — Install-Exchange15.ps1 by Michel de Rooij

> The following changelog covers the original [Install-Exchange15.ps1](http://eightwone.com) by
> Michel de Rooij, on which EXpress is based. Reproduced here for traceability.
> EXpress was forked after version 4.23.

### 4.23
- Fixed Edge installation (no need to check for Exchange 2013 in AD)

### 4.22
- Corrected VC++ 2013 runtime download URL (shortcut was unavailable)

### 4.21
- Disabling MSExchangeAutodiscoverAppPool during setup to prevent responding to requests during setup and post-config

### 4.20
- Clearing/setting SCP is now a background job during install (asynchronous)

### 4.13
- Fixed race issue when installing from ISO and restarting installation
- Tested with Exchange SE ISO

### 4.12
- Fixed feature installation (`Web-W-Auth` → `Web-Windows-Auth`)
- Using ADSI for Exchange 2013 detection

### 4.11
- Fixed feature installation for WS2022/WS2025 Core

### 4.10
- Added support for Exchange Server SE (Subscription Edition)

### 4.01
- Removed obsolete TLS 1.3 setup detection

### 4.0
- Added support for Exchange 2019 CU15 and Windows Server 2025 (CU15+)
- Removed Exchange 2013 support; removed Exchange 2016 CU1–CU22; removed Exchange 2019 RTM–CU9; removed WS2012R2
- Added removal of obsolete MSMQ feature when installed
- Added `EnableECC`, `NoCBC`, `EnableAMSI`, `EnableTLS12`, `EnableTLS13` switches
- Removed `InstallMailbox`, `InstallCAS`, `InstallMultiRole`; removed `NoNet461/471/472/48`; removed `UseWMF3`
- Added Exchange 2013 detection (cannot coexist with CU15+)
- Set minimum required PowerShell version to 5.1; functions now use approved verbs; code cleanup

### 3.9
- Added support for Exchange 2019 CU14
- Added .NET Framework 4.8.1 support; added `NONET481` switch
- Added `DoNotEnableEP` and `DoNotEnableEP_FEEWS` switches (Exchange 2019 CU14+ Extended Protection)
- Added AUG2023 SU deployment for CU13/CU12/CU23 when `-IncludeFixes` specified
- Fixed detection of source path when ISO already mounted without drive letter

### 3.8
- Added support for Exchange 2019 CU13

### 3.71
- Updated recommended Defender AV inclusions/exclusions

### 3.7
- Added support for Windows Server 2022
- Fixed IIS URL Rewrite module install logic (CU22+/CU11+)
- Fixed `/IAcceptExchangeServerLicenseTerms_DiagnosticData*` switch detection

### 3.62
- Added support for Exchange 2019 CU12 and Exchange 2016 CU23

### 3.61
- Added mention of Exchange 2019 in output

### 3.6
- Added support for Exchange 2019 CU9–CU11 and Exchange 2016 CU20–CU22
- Added IIS URL Rewrite prerequisite for CU11+/CU22+
- Added `DiagnosticData` switch (initial DataCollectionEnabled mode)
- Added KB2999226 support for WS2012R2

### 3.5
- Added support for Exchange 2019 CU8–CU9 and Exchange 2016 CU19–CU20
- Added KB5003435 (2019CU8/9, 2016CU19/20, 2013CU23) and KB5000871 (older CUs)
- Added Interim Update installation and detection
- Updated .NET 4.8, VC++ 2012, and VC++ 2013 download links
- Fixed High Performance Power Plan setting for recent Windows builds

### 3.4
- Added support for Exchange 2019 CU8 and Exchange 2016 CU19
- Script allows non-static IP when Azure Guest Agent, Network Agent, or Telemetry Service is present

### 3.3
- Added support for Exchange 2019 CU7 and Exchange 2016 CU18

### 3.2.6
- Added support for Exchange 2019 CU6 and Exchange 2016 CU17
- Added VC++ 2012 Redistributable requirement for Exchange 2019

### 3.2.5
- Fixed typo in Exchange build enumeration
- Fixed specified vs. used `MDBLogPath` (was incorrectly appending `\Log`)

### 3.2.4
- Added support for Exchange 2019 CU4 + CU5 and Exchange 2016 CU15 + CU16

### 3.2.3
- Fixed typo in Exchange 2019 CU3 detection

### 3.2.2
- Added support for Exchange 2019 CU3 and Exchange 2016 CU14

### 3.2.1
- Updated pagefile configuration for Exchange 2019 (25% of memory size)

### 3.2
- Added support for Exchange 2019 CU2, Exchange 2016 CU13, Exchange 2013 CU23
- Added .NET Framework 4.8 support; added `NoNET48` switch
- Added disabling Server Manager during installation
- Removed WS2008R2 and WS2012 support; removed `UseWMF3` switch

### 3.1.1
- Fixed detection of Windows Defender

### 3.1.0
- Added support for Exchange 2019 CU1, Exchange 2016 CU12, Exchange 2013 CU22
- Fixed KB3041832 URL; fixed NoSetup mode/EmptyRoles problem
- Added skip for Health Monitor checks on Edge installations

### 3.0.0 – 3.0.4
- Added Exchange 2019 support; rewritten VC++ detection
- Integrated Exchange 2019 RTM cipher correction
- Replaced filename constructs with `Join-Path`; fixed various bugs

### 2.99 – 2.99.9
- Added Windows Defender exclusions (Exchange 2016 on WS2016)
- Added support for Exchange 2016 CU8–CU11 and Exchange 2013 CU19–CU21
- Added .NET 4.7, 4.7.1, 4.7.2 blocking switches; added upgrade mode detection
- Added Exchange 2019 Public Preview support (WS2016 and WS2019)
- Added post-setup HealthCheck / IIS Warmup; added `SkipRolesCheck` switch
- Fixed Recover mode phase sequencing; fixed `InstallMDBDBPath` location check
- Script aborts on non-static IP (unless Azure/Network Agent present)

### 2.8 – 2.98
- Added `DisableRC4` switch (KB2868725); added NIC Power Management disable
- Added WS2016 support (Exchange 2016 CU3+); added Exchange 2016 CU4–CU9
- Added Exchange 2013 CU14–CU20 support; added FFL 2008R2 checks
- Added blocking of .NET 4.7 / 4.7.1 / 4.7.2 Preview; added `NONET471` switch
- Various cosmetics, code cleanup, and minor bug fixes

### 2.5 – 2.7
- Added recommended hotfixes: KB3146717, KB2985459, KB3041832, KB3004383
- Added AD Site logging; added computer name to filenames
- Changed credential prompting to use current account
- Fixed KeepAlive timeout; added Schema/Enterprise Admin checks
- Added `Recover` switch for RecoverServer mode
- Script aborts on non-zero ExSetup exit code; ignores SUS_E_NOT_APPLICABLE
- Fixed SCP parameter handling; fixed NoSetup `.NET 4.6.1` logic

### 2.3 – 2.42
- Added Exchange 2013 CU12–CU13 and Exchange 2016 CU1–CU2 support
- Added .NET 4.6.1 support with OS-dependent hotfixes; added `NONET461` switch
- Added `DisableSSL3` switch (KB187498); added Keep-Alive and RPC timeout settings
- Switched version detection to ExSetup (follows build number)

### 2.0 – 2.2
- Renamed script to `Install-Exchange15`
- Added Exchange 2016 Preview and CU9 support; fixed registry/GPO checks
- Added `ClearSCP` / `SCP` parameter; added `Lock` switch
- Added `Import-ExchangeModule()` for EMS post-configuration
- Added multi-forest AD support; added access checks for paths

### 1.5 – 1.9
- Added WS2008R2 support (.NET 4.5, WMF3, IEESC toggling, hotfixes)
- Added Exchange 2013 SP1, WS2012R2 support; added .NET 4.51/4.52/4.6.1/4.7.2
- Added `DisableRetStructPinning`, `InstallFilterPack`, `UseWMF3`, `NONET461` switches
- Added Exchange 2016 CU1–CU9, Exchange 2013 CU12–CU20 support
- Fixed AutoPilot RunOnce; fixed OS version comparison for localized OS

### 1.0 – 1.1
- Initial community release
- Added AD preparation logic; added domain/forest functional level checks
- RSAT-ADDS-Tools not uninstalled when used for AD preparation
- Pending reboot detection — AutoPilot reboots and restarts phase
- Installs `Server-Media-Foundation` feature (UCMA 4.0 requirement)
- Validates provided credentials for AutoPilot
