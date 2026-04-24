# EXpress — Release Notes

Full optimization and feature history. See `README.md` for user-facing changelog.

**Versioning scheme:** `MAJOR.MINOR` = feature release; `MAJOR.MINOR.PATCH` = bugfix / maintenance release on top of the matching feature version.

---

## v5.96 (2026-04-24) — feature

### F24 — Installation-Document Template (hybrid)

Optional custom DOCX template for the Word installation document.

- New parameter **`-TemplatePath <path>`** — path to a customer DOCX template. When supplied, the cover page and header/footer come from the template; the script injects all 18 chapter sections as generated XML.
- **`{{token}}` placeholder replacement** — template must contain `{{document_body}}` (required) plus any combination of `{{Organization}}`, `{{ServerName}}`, `{{Scenario}}`, `{{InstallMode}}`, `{{Version}}`, `{{DateLong}}`, `{{Author}}`, `{{Company}}`, `{{Classification}}`, `{{HeaderLabel}}`, `{{DocTitle}}`, `{{CoverSub}}`. Tokens are replaced in every XML part (document, header, footer).
- **`Test-WdTemplate -Path -RequiredTags`** — validates that a template DOCX contains all required `{{token}}` placeholders before use; missing required tokens trigger an automatic fallback to the built-in cover page with a warning.
- **`Write-WdFromTemplate -TemplatePath -OutputPath -Tokens`** — copies the template ZIP, replaces tokens in all XML parts. `{{document_body}}` replaces its entire anchor paragraph `<w:p><w:r><w:t>{{document_body}}</w:t></w:r></w:p>` with the generated chapter XML.
- **`tools\Build-InstallationTemplate.ps1`** — maintainer tool to generate starter templates `templates\Exchange-installation-document-EN.docx` and `-DE.docx`. Both ship in the repository as customization starting points.
- Config-file support: `TemplatePath` key in `.psd1` deployment files.
- Fallback: if `-TemplatePath` is omitted or the template fails validation, the existing built-in cover page is used unchanged (zero regression).

---

## v5.95.1 (2026-04-24) — bugfix

- Windows Update progress bar: label changed from `Xs remaining` to `auto-abort in Xs` to clarify it is the timeout countdown, not the estimated completion time.
- Main menu language toggle: label updated to `Language:  DE (default EN)` to make the English default explicit.

---

## v5.95 (2026-04-24) — feature

### F25 — Advanced Configuration menu

Main menu shrunk to the five installation-flow toggles; ~55 hardening / tuning / policy switches moved into a new paginated Advanced Configuration menu, prompted with a 60-second auto-skip (default = keep current v5.x behaviour).

- **Main menu toggles now A / B / N / R / U / V** — Autopilot, Install Exchange SU, Preflight-only, Install Windows Updates, Generate Installation Document, document language (German).
- **Advanced menu** — six pages (Security / TLS · Security / Hardening · Performance / Tuning · Exchange Org Policy · Post-Config / Integration · Install-Flow / Debug). Two-column layout, letter-keys to toggle, `N`=Next, `P`=Prev, `A`=Apply on last page, `S`=Skip-all, `ESC`=Cancel. `Read-Host` fallback when `RawUI.KeyAvailable` is unavailable (PS2Exe / redirected host).
- **`Get-AdvancedFeatureCatalog`** — ordered hashtable keyed by feature Name (`DisableSSL3`, `LSAProtection`, `MaxConcurrentAPI`, …); each entry carries `Category`, `Label`, `Description`, `Default` and an optional `Condition` scriptblock (TLS 1.3 hidden pre-WS2022, Shadow Redundancy hidden without DAG, AnonymousRelay hidden without `RelaySubnets`, etc.).
- **`Test-Feature -Name`** — single gate used wherever the script reads `$State['<name>']`. Precedence: `$State['AdvancedFeatures'][Name]` > catalog default. Unknown names return `$false` + verbose warning.
- **`Invoke-AdvancedConfigurationPrompt`** — 60-second countdown offer (`Write-Progress -Id 2`) shown after the main menu. Silently skipped in Autopilot or when `$State['SuppressAdvancedPrompt']` is set. Result persisted in `$State['AdvancedFeatures']` via `Save-State`.
- **Config-file parity** — new nested `AdvancedFeatures = @{ … }` block in `deploy-example.psd1`. Precedence: nested block > legacy top-level key > cmdline `-<Name>` switch > catalog default. Existing `.psd1` files and `-DisableSSL3` / `-EnableECC` / … cmdline switches keep working unchanged; they're merged into `$State['AdvancedFeatures']` at startup.

### Note
- Always-on (non-negotiable): Defender Exclusions + transient realtime disable, Page-File fixed size, VDir URLs when namespace set, Exchange log-cleanup scheduled task, High-Performance power plan. These are **not** exposed as toggles.

---

## v5.94.1 (2026-04-24) — enhancement

### Word installation document — SIEM guidance and retention-tag detail

Follow-up on the v5.94 audit-readiness work. Three additions clarifying what EXpress *does not* configure (long-term log retention) and surfacing previously-missing org data.

- **§8.6 — SIEM/forensics paragraph** — second paragraph after the existing logging/cleanup intro. Spells out that the scheduled cleanup (see §7.1) is purely volume-protection and not a substitute for tamper-evident long-term retention. Recommends SIEM integration (Splunk, Elastic Security, Microsoft Sentinel, Wazuh, IBM QRadar) via NXLog / WEF-WEC / Filebeat / Azure Monitor Agent, with typical retention guidance (12 months hot, 7 years archive). Cites BSI APP.5.2, GDPR accountability, GoBD.
- **§8.8 — SIEM-context paragraph + two mapping rows** — explanatory paragraph before the compliance table covering the SIEM value proposition (central correlation, anomaly alerting, tamper-evident retention, audit evidence) and recommended source channels (Windows Security/System/Application, IIS W3C, MessageTracking, HttpProxy, Managed Availability, Search-AdminAuditLog / Search-MailboxAuditLog). Two new table rows: *SIEM integration* (CIS Control 8 / BSI OPS.1.1.5 / APP.5.2 A13 — Out of scope, organisation-wide planning) and *Local log cleanup* (Implemented — scheduled task per §7.1).
- **§4.7 — Retention Tags rendered in detail** — `Get-OrganizationReportData` now also collects `Get-RetentionPolicyTag` into `$org.RetentionPolicyTags`. New sub-table after the existing Retention Policies table renders each tag's `Name`, `Type`, `AgeLimitForRetention` (in days), `RetentionAction` and `RetentionEnabled` state, sorted by Type then Name. Previously only policy-to-tag links were shown — actual tag behaviour (move-to-archive vs. delete vs. mark) was invisible to auditors.

### Verified — no code change

- **Admin Audit Log** — confirmed `AdminAuditLogEnabled=$true` is the Exchange 2013+ default; `Get-AdminAuditLogConfig` query in `Get-OrganizationReportData` and rendering in §4.16 already cover all relevant properties (enabled, age limit, mailbox, cmdlets, exclusions). No change needed.

---

## v5.94 (2026-04-24) — feature

### Word installation document — audit-readiness sections

Nine new sections added to `New-InstallationDocument` so the generated Word file doubles as an audit/handover record (ISO 27001, BSI IT-Grundschutz, DSGVO) without manual post-processing:

- **§1.1 Change-Management** — placeholder 2-column table (Change-Request-Nr., approver, approval date, sign-off, sign-off date, remarks) for the operator to complete after go-live.
- **§4.17 Service-Accounts / RBAC-Rollengruppen** — live `Get-RoleGroupMember` query for six Exchange role groups (Organization / Server / Recipient / Hygiene / Compliance Management, View-Only Organization Management). Multi-line cells via `<w:br/>` so each member is on its own line.
- **§6.2 Ports/Firewall** — 13-row static reference (SMTP 25/587, HTTPS 443, HTTP 80, IMAPS 993, POP3S 995, RPC 135, SMB 445, GC 3268/3269, WinRM 5985/5986, MAPI-over-HTTP 64327).
- **§7.2 Security-Update-Status** — Exchange build, Windows build, last boot, `$State['ExchangeSUVersion']` when present.
- **§8.8 Compliance-Mapping** — 14 hardening measures cross-referenced to CIS Benchmark and BSI IT-Grundschutz module IDs.
- **§8.9 DSGVO/Datenschutz-Hinweise** — 8-row GDPR checklist (Art. 4 Nr. 2 classification, AVV, TOM, Löschkonzept, Auskunftsrecht, Verletzungsmeldung, DSFA, Betriebsrat).
- **§10.4 Backup-Nachweis** — placeholder for backup product, last full/incremental, restore-test cadence, RPO/RTO.
- **§12.2 Monitoring-Checkliste** — 9-row go-live readiness list (crimson channels, HC schedule, DAG health, certificate expiry, queue depth, disk space, backup success, AD replication, security updates).
- **Chapter 16 Abnahmetest** — 12 acceptance test cases with auto-populated OWA/ECP/EWS/Autodiscover URLs when `$State['Namespace']` is set; otherwise `<Namespace>` placeholder.

### Chapter renumbering

- Operative Runbooks: §16 → §17 (subsections 16.1–16.6 → 17.1–17.6).
- Offene Punkte: §17 → §18.

### `New-WdTable` multi-line cells

Cell text containing `` `n `` is emitted with `<w:br/>` between runs so each line renders on its own row inside the cell (used by §4.17 Service-Accounts role-group membership list). `-Compact` (8pt) is now reserved for tables with 4+ columns that would otherwise wrap; 2–3 column tables stay at the default 11pt for legibility.

### F25 Advanced Configuration Menu — plan absorbed into master

The WIP on `feature/advanced-menu` (4 scaffolding functions, 319-line diff, stash commit `787b221`) is preserved as a reference. The design lives in Claude memory (not in the repo). No code on master changed — avoids maintaining two branches for a feature that is not yet scheduled.

---

## v5.93 (2026-04-23) — feature

### Hybrid-aware MEAC task registration + AD Split-Permissions prep

`Register-AuthCertificateRenewal` now drives CSS-Exchange `MonitorExchangeAuthCertificate.ps1` (MEAC) with the two scenarios the upstream documentation explicitly supports beyond the default: hybrid coexistence and AD Split Permissions. MEAC itself continues to self-provision the `SystemMailbox{b963af59-3975-4f92-9d58-ad0b1fe3a1a3}` automation account, its RBAC role group, and the scheduled task — EXpress only wires the operator-supplied knobs through.

#### Auto-generated automation-account password (standard deployments)

MEAC provisions the `SystemMailbox{b963af59-…}` user account itself, but Task Scheduler needs a password at registration time in order to register a task that runs AS that user — without one, MEAC logs "Please provide a password for the automation account" and exits without creating the task. Observed symptom: `Get-ScheduledTask -TaskName 'Daily Auth Certificate Check'` returns nothing even though MEAC reported no error.

EXpress now generates a strong 32-character random password inline (cryptographic RNG, mixed alpha+digits+symbols) and passes it via MEAC `-Password`. The password is transient:

- MEAC sets it on the `SystemMailbox{b963af59-…}` account AND registers the scheduled task atomically;
- Windows Task Scheduler then stores the credential in its (DPAPI-protected) local credential store;
- EXpress never needs the password again — not persisted to state, not logged, not returned. If re-registration ever becomes necessary (password policy rotation, task deleted), re-running `Register-AuthCertificateRenewal` generates a fresh password and MEAC resets both the account and the task atomically.

When `-MEACAutomationCredential` is supplied (Split-Permissions path), the CLI credential takes precedence and no password is generated.

#### Hybrid coexistence (transparent)

Per MEAC docs, replacing the Auth Certificate in a hybrid org requires a Hybrid Configuration Wizard rerun afterwards; doing it silently breaks federation with Exchange Online. MEAC therefore refuses to renew by default when it detects hybrid — the operator must pass `-IgnoreHybridConfig` to authorise.

EXpress now mirrors that safety model without surprising the operator:

1. `Register-AuthCertificateRenewal` probes `Get-HybridConfiguration`. If hybrid is detected and `-MEACIgnoreHybridConfig` was not supplied, the task is registered in hybrid-safe mode and a multi-line advisory is logged at registration time (not buried in `-Verbose`).
2. `-MEACIgnoreHybridConfig` is a pure passthrough. When present, it is forwarded to MEAC and a warning restates the HCW-rerun obligation.
3. `-MEACNotificationEmail <addr>` is forwarded to MEAC `-SendEmailNotificationTo`. In hybrid mode it's the recommended escape hatch: daily checks still run, the task doesn't renew, but the admin gets an email 60 days before expiry with enough time to schedule an HCW rerun.
4. `-MEACIgnoreUnreachableServers` is forwarded for multi-server orgs where partial outages are expected.

All four knobs have config-file equivalents (`MEACIgnoreHybridConfig`, `MEACIgnoreUnreachableServers`, `MEACNotificationEmail`) — the credential one does not (see below).

#### AD Split Permissions

Per MEAC docs, under AD Split Permissions the Exchange admin cannot create AD users, so the `SystemMailbox{b963af59-3975-4f92-9d58-ad0b1fe3a1a3}@<domain>` automation account must be pre-created by a Domain Admin on a non-Exchange box via `-PrepareADForAutomationOnly -ADAccountDomain <domain>`. The resulting credential is then passed to the Exchange-side run via `-AutomationAccountCredential`.

EXpress now supports both sides of that split:

- **DA side — `-MEACPrepareADOnly -MEACADAccountDomain contoso.local`**: a new standalone parameter set dispatched immediately after the logging bootstrap, before any Exchange prereq check. EXpress downloads MEAC into `$env:TEMP`, invokes it with `-PrepareADForAutomationOnly -ADAccountDomain $MEACADAccountDomain -Confirm:$false`, logs each MEAC line at INFO tier, and exits with `ERR_OK` (or the new `ERR_MEACPREPAREAD = 1038` on failure). No state, no phase loop, no reboot. Safe to run on a domain-controller or admin workstation with only the ActiveDirectory module.
- **Exchange side — `-MEACAutomationCredential (Get-Credential)`**: the standard install run receives the DA-created credential. EXpress persists it at Phase 0 intake (`$State['MEACAutomationUser']` + DPAPI-encrypted `$State['MEACAutomationPW']`) so it survives the Autopilot reboot chain, and `Register-AuthCertificateRenewal` rehydrates it via the new helper `Get-MEACAutomationCredentialFromState` and forwards it to MEAC as `-AutomationAccountCredential`. No bind-test, no validation, no prompts — it's a pure passthrough. If the credential is wrong, MEAC reports it; EXpress does not re-invent that layer.

No plain-text config key for the automation password — deliberately. Split-Permissions deployments where a config file can carry the password are rare enough that the security cost of the plain-text path isn't justified. `-MEACAutomationCredential` is CLI-only (or picked up from a per-deployment secret store via `$env:`-assembled PSCredential in a wrapper).

### Scope

- **New CLI:** `-MEACAutomationCredential`, `-MEACIgnoreHybridConfig`, `-MEACIgnoreUnreachableServers`, `-MEACNotificationEmail` (available in sets M / NoSetup / Recover); `-MEACPrepareADOnly` + `-MEACADAccountDomain` (new parameter set `MEACPrepareAD`).
- **New config keys:** `MEACIgnoreHybridConfig`, `MEACIgnoreUnreachableServers`, `MEACNotificationEmail`.
- **New state keys:** `MEACAutomationUser`, `MEACAutomationPW` (DPAPI, user+machine bound). Populated only when `-MEACAutomationCredential` is supplied.
- **New error code:** `ERR_MEACPREPAREAD = 1038`.
- **New helper:** `Get-MEACAutomationCredentialFromState` (13-line rehydrate-only helper; no validation).
- **`Register-AuthCertificateRenewal`:** hybrid probe + passthrough assembly. `$meacParams` hashtable fills incrementally based on what the operator supplied.

### Not in scope

The feature budget is roughly 130 lines of new code. No OU picker, no bind test, no 3-way A/M/S prompt, no Phase-0 credential cascade, no RBAC probe, no batch-logon probe, no plain-text MEAC password in config — MEAC handles all of those concerns itself when it provisions the automation account. EXpress stays a thin wrapper.

---

## v5.92 (2026-04-23) — feature

### Plain-text install-admin credentials in config file (fully unattended)

For true zero-touch deployment pipelines, the `-ConfigFile` .psd1 may now carry plain-text install-admin credentials alongside every other parameter:

```powershell
@{
    # ...existing keys...
    AdminUser     = 'CONTOSO\svc-exchadmin'
    AdminPassword = 'R3plac3-M3-Immediately!'
}
```

On load, EXpress converts the plain strings into a `PSCredential` object and emits a loud, box-framed **SECURITY WARNING** to the transcript every run:

```
################################################################
##  SECURITY WARNING: PLAIN-TEXT CREDENTIALS IN CONFIG FILE   ##
################################################################
```

These keys are an opt-in convenience for short-lived, unattended installation runs only. **The config file must be deleted or scrubbed immediately after install completes.** The state file's DPAPI encryption is user+machine bound; the config file is not — it is a plain artefact on disk, in backups, and in any version-control system it touches.

CLI `-Credentials` takes precedence over the config-file keys (command line wins).

### `Check.ps1` quality gate moved under `tools/`

Call sites now use `.\tools\Check.ps1 -SkipAnalyzer`, matching the rest of the quality-gate toolbelt (`Test-ScriptQuality.ps1`, `Test-ScriptSanity.ps1`, `Parse-Check.ps1`).

---

## v5.91 (2026-04-23) — feature

### Default output language flipped to English; single `-German` switch

The HTML install report and pre-flight report have always been English-only. The Word installation document was the outlier: it defaulted to German via `-Language DE`. This made the script's default output language inconsistent and awkward for non-German audiences.

- **Default: English.** All output (Word document, menu prompts, help) is English unless `-German` is set.
- **`-German` switch** is the single opt-in for German output. The previous `-Language DE|EN` parameter has been removed — a two-value parameter plus a default is strictly more surface area than a boolean switch for the same thing.
- **Menu prompt** (mode 7 "Standalone Document"): replaced "Document language [DE/EN]" with "German output? [y/N] (default: English)".
- **Config-file back-compat:** legacy `Language=DE` entries in config files are still honoured and map to `-German`. New config files can use `German=true` directly.
- `$State['Language']` remains `'EN'`/`'DE'` internally — `New-InstallationDocument`'s `L` helper was not touched.

Behavioural contract: omit `-German` → English. Set `-German` → German. Nothing in between.

---

## v5.90 (2026-04-23) — feature

### Word document: three readability + correctness fixes

**(1) Anti-spam filter display now reflects reality.** §9.1 previously showed `Get-*FilterConfig.Enabled` as "Enabled", which is misleading — that is only the organisation-wide feature switch. EXpress disables the underlying transport agents (`Disable-TransportAgent`) for every filter except Recipient Filter, so "org config Enabled = True" typically coexists with "agent Disabled" → the filter does not fire.

Each filter table (Content / Sender / Recipient / Sender ID) now has two rows:
- **Effective status (transport agent)** — the actual pipeline state from `Get-TransportAgent`:
  - `Enabled — agent runs in transport pipeline`
  - `Inactive — transport agent is disabled, filter does not fire (org switch is only a feature flag)`
  - `Not installed` (agent not present on this server)
- **Org config Enabled (feature flag)** — `Get-*FilterConfig.Enabled` (the old row), clearly labelled as the feature switch.

A paragraph above §9.1 explains the distinction so operators don't mis-read the table.

**Why the other agents are disabled at all:** on Mailbox-role servers behind EOP/gateway, Sender ID rejects practically every inbound because the visible source IP is the gateway (SPF fails on Neutral/Fail), Content Filter double-scores already-classified mail (unpredictable FP rate), and Connection/Sender Filters add no value without curated lists. Recipient Filter stays on for Directory Harvest Attack Protection — that single agent is always worth running.

**(2) Exchange Online / M365 promoted to top-level section.** Moved from §4.17 (inside "Organisation — Global Configuration") to §15 "Exchange Online and Microsoft 365", placed immediately before §16 "Operational Runbooks". Hybrid considerations (CMT, Free/Busy, Move Request, EOP, namespace, licensing) belong with day-2 operations, not org-config telemetry. "Open Items" renumbered §16 → §17.

**(3) Wide tables now fit the page.** `New-WdTable` gets a `-Compact` switch that emits runs at 8pt (`<w:sz w:val="16"/>`) instead of 11pt — ~40% more horizontal characters per line, turning cascading wraps on every column into at most one break per cell. Applied to:
- **Receive Connectors** — the 8-column table (Name / Bindings / Remote IPs / Auth / Permissions / TLS / FQDN / Max size) is also split into two 4/5-column tables (Network + Security and limits) sharing the connector name as the join key.
- **Certificates per server** — 5 columns (Subject / Expiry / Remaining / Services / Thumbprint); Subject and Thumbprint are long, Compact keeps them on one line.

---

## v5.88.3 (2026-04-23) — bugfix

### Word document: `$state` shadows `$State` hashtable — "Unable to index into System.String"

PowerShell variable lookup is **case-insensitive**: `$state` and `$State` are the same variable. Two locations inside `New-InstallationDocument` used `$state` as a local loop/condition variable, silently overwriting the script-level `$State` hashtable with a string, causing "Unable to index into an object of type System.String" at `$State['WordDocPath'] = $docPath` (line 7378).

- §12.1 Crimson Event Log loop: `$state = if ($log.IsEnabled) ...` → renamed `$logState`
- `Draw-OptimizationMenu`: `$state = if ($sel[$LastKey]) ...` → renamed `$optState`

`Test-ScriptQuality.ps1` now has a **SingletonShadow** detector (section 3e): walks all `AssignmentStatementAst` nodes, finds any local variable whose name (lowercased) matches a known script-level singleton (`$State`, `$StateFile`), and flags those not in the legitimate owner allowlist (`Restore-State`, `<top-level>`).

---

## v5.88.2 (2026-04-23) — bugfix

### Word document: fix remaining `(if ...)` runtime crashes (PS 5.1)

Six more occurrences of a control-flow statement inside plain grouping parens `(if ...)` — a known PS 5.1 pitfall documented in CLAUDE.md that produces "The term 'if' is not recognized as the name of a cmdlet" at runtime:

- §2 Installation Parameters — `TLS 1.0 / TLS 1.1` row (`.Add(@('...', (if ...)))`)
- §4.16 Admin Audit Log — `AdminAuditLogCmdlets` and `AdminAuditLogExclusions` rows (`SafeVal (if ...)`)
- §9.1 Sender Filter — `BlockedSenders` and `BlockedDomains` rows (`SafeVal (if ...)`)
- §9.1 Recipient Filter — `BlockedRecipients` row (`SafeVal (if ...)`)

All fixed by assigning the `if`-result to a temporary variable first. The `Test-ScriptQuality.ps1` detector (section 3a) was widened from a `CommandAst`-only scan to a global `ParenExpressionAst` walk — it now catches the pattern in all contexts (method arguments, array literals, binary operands). `Test-ScriptSanity.ps1` check 3 updated accordingly. New `Check.ps1` root wrapper runs both suites with a single command.

---

## v5.88.1 (2026-04-23) — bugfix

### MEAC: verify scheduled-task registration actually succeeded

`Register-AuthCertificateRenewal` no longer trusts MEAC's own exit signal. The success message is now conditional on `Get-ScheduledTask` confirming the `"Daily Auth Certificate Check"` task exists after the MEAC invocation. If not, a warning with the MEAC log path is emitted instead of a false-positive success line, so registration failures surface at install time rather than five years later when the Auth Certificate silently expires.

---

## v5.88 (2026-04-23) — feature

### Word installation document — §4.16 Admin-Auditprotokoll-Konfiguration

New org-wide section collects `Get-AdminAuditLogConfig` and documents Admin Audit Log enabled state, retention period, log mailbox, logged cmdlets, exclusions, test-cmdlet logging, and log level. Basis for compliance evidence (ISO 27001, BSI-Grundschutz, DSGVO-Rechenschaftspflicht).

### Word installation document — §9.1 Anti-Spam-Filter-Konfiguration

New sub-section after the transport agent table collects `Get-ContentFilterConfig`, `Get-SenderFilterConfig`, `Get-RecipientFilterConfig`, and `Get-SenderIdConfig`. Rendered as four separate sub-tables (Content Filter, Sender Filter, Recipient Filter, Sender ID) showing enabled state, SCL thresholds, quarantine mailbox, block lists, and spoof/temp-error actions. Only appears when anti-spam agents are installed (safe to call — caught if not present).

### Word installation document — §12.1 Exchange Crimson Event Log Channels

New sub-section in the Monitoring-Readiness chapter queries `Get-WinEvent -ListLog "Microsoft-Exchange*"` and lists all enabled or populated `/Operational` and `/Admin` crimson channels with current state, maximum log size, and record count. Helps document which monitoring channels are available and whether they have been configured for external forwarding. Closing paragraph names the four most important channels for production monitoring.

---

## v5.87 (2026-04-23) — feature

### Word installation document — binary registry values → localised text

All hardening registry values (0/1) in Sections 8.1–8.4 are now rendered as localised human-readable text (`aktiviert` / `deaktiviert` in German, `enabled` / `disabled` in English) via a new `Format-RegBool` helper instead of raw integers. Affected rows: `.NET Strong Crypto v4/v2`, `WDigest UseLogonCredential`, `LSA RunAsPPL`, `Credential Guard (VBS)`, `HTTP/2 Cleartext`, `SMBv1`, `Serialized Data Signing`, `AMSI Body Scanning`, `ECC Certificate Support`. AMSI has dedicated handling because its registry value is a *Disable* flag (0 = AMSI active, 1 = AMSI off). LM Compatibility Level now shows `Level N` instead of a plain integer.

### Word installation document — TLS protocol state: remove `(explizit konfiguriert)` annotation

`Get-TlsProtocolState` no longer appends `(explizit konfiguriert)` / `(explicitly configured)` to the state string. The annotation was redundant — the OS-default case is already distinguished by the `(OS-Standard)` suffix. The function also renames `aktiv` → `aktiviert` for consistency with all other hardening rows.

### Word installation document — IMAP/POP3 configuration section

`Get-ServerReportData` now collects `Get-ImapSettings` / `Get-PopSettings` for the local server. A new sub-section `5.x.6 IMAP/POP3-Konfiguration` in the per-server chapter shows external/internal connection settings, X.509 certificate name, and login type. Fields left blank prompt `(bitte manuell ergänzen)` if the namespace has not been configured yet.

### Word installation document — receive and send connector detail

Receive connector table extended with `RequireTLS`, `FQDN`, and `Max. Größe` columns (Sections 5.x.5). Send connector table extended with the same three fields plus the column order was adjusted to put FQDN/TLS/size next to the routing information (Section 4.10).

### Word installation document — DNS section replaced with manual template

The split-DNS lookup (Sections 6.1) was removed: `Resolve-DnsName` from inside the server returns the internal DNS view, SOA fallback objects rendered as their type name, and external records often do not exist at installation time. Replaced by a static template table pre-populated with accepted domain names and placeholder `(bitte manuell ergänzen)` values for MX, SPF, DKIM, DMARC, and Autodiscover A/CNAME — to be filled in after go-live.

### Word installation document — Autodiscover AppPool: show live IIS state

Section 8.4 previously showed `$State['DisableAutodiscoverAppPool']` (a configuration intent flag). Replaced with a live `Get-WebAppPoolState 'MSExchangeAutodiscoverAppPool'` query so the document reflects the actual current pool state (Started/Stopped) rather than what was configured during installation.

### Word installation document — EWS Max Concurrency: handle null (not-set) gracefully

`Get-ThrottlingPolicy` returns `$null` for `EwsMaxConcurrency` on fresh organisations where the value was never explicitly configured. The row now shows `(nicht gesetzt — Standard: 27)` / `(not set — default: 27)` instead of an empty string or runtime exception.

### Word installation document — installing user recorded

`$State['InstallingUser']` is now captured at Phase 0 start using `[Security.Principal.WindowsIdentity]::GetCurrent().Name`. Section 1 (Document Properties) shows the value as `Installiert durch` / `Installed by`.

### Word installation document — MEAC scheduled-task search broadened

`Get-OrganizationReportData` now uses a two-pass approach: first a direct `Get-ScheduledTask -TaskName` lookup for each known name, then a broad `Where-Object { $_.TaskName -match 'MonitorExchangeAuth|ExchangeLogClean|EXpressLog' }` pass as fallback. This handles CSS-Exchange naming variants and ensures the MEAC task appears in Section 7.1 regardless of the exact task name used by different CSS-Exchange releases.

---

## v5.86.2 (2026-04-23) — bugfix

### Word installation document — PS 5.1 nested-array flattening in `New-WdTable`

Literal `@( @('a','b'), @('c','d') )` row arrays passed to `New-WdTable -Rows` were flattened by PowerShell 5.1's `@()` operator *before* the parameter binder saw them, so the function received `@('a','b','c','d')` and bound it to `[object[][]]` as four rows × one cell each. Visible symptom: the cover-page *Versionshistorie*, the Section 1 *Dokumenteigenschaften*, the Section 4.12 *Auth-Zertifikat* and every other literal-array table rendered with only the first column filled and each original cell on its own row. `New-WdTable` now auto-detects the flattened shape (all elements scalar AND total length a multiple of the header column count) and reshapes in place before emitting rows. Call sites using `List[object[]].ToArray()` or `,@(…)` per row are unaffected.

### Word installation document — Section 4.12 Auth Certificate "Valid until" empty

`Get-AuthConfig` does not expose `NotAfter` or `CurrentCertificateLifetime`; the previous code read the non-existent property and emitted an empty cell. The validity cell is now computed by looking up the cert via `Get-ExchangeCertificate -Thumbprint … -Server` (with a `Cert:\LocalMachine\My` fallback) and printing `yyyy-MM-dd (N days remaining)`. Missing Next/Previous thumbprints and empty Realm/ServiceName now render as `(nicht gesetzt)` / `(leer — Default)` instead of blanks.

### Word installation document — Transport Agents table empty Name column

Implicit-remoting deserialization of `TransportAgent` objects sometimes leaves `.Name` blank while `.Identity` carries the display name, so Sections 5.x *Transport Agents* and Chapter 9 *Transport-Agents und Anti-Spam* showed status+priority with no agent name. Both call sites now fall back to `[string]$ta.Identity` when `$ta.Name` is empty. Chapter 9 additionally now iterates all four transport scopes (`TransportService`, `FrontendTransport`, `MailboxSubmission`, `MailboxDelivery`) and deduplicates by name — previously only the default `HubTransport` scope was queried, hiding agents such as the Front-End SMTP Receive agents.

### Word installation document — Section 8.1 TLS semantics ambiguous

The TLS table printed raw registry values with labels like `TLS 1.0 Server Disabled = 0`, which reads as a double-negation and is easily misread as "TLS 1.0 is active". Replaced the raw-value column with a semantic state (`aktiv (explizit konfiguriert)` / `deaktiviert (OS-Default)` / etc.) derived from both `Enabled` and `DisabledByDefault` values, with an explicit `— ACHTUNG: …` suffix when a hardening gap is detected (e.g. TLS 1.0 actually enabled or TLS 1.2 actually disabled).

### Word installation document — Serialized Data Signing always "(not set)"

`New-InstallationDocument` Section 8.4 read registry value `EnableSerializedDataSigning`, but `Enable-SerializedDataSigning` writes to `EnableSerializationDataSigning` (Microsoft's actual spelling under `HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Diagnostics`). The row therefore always rendered as `(not set)` even after the hardening had been applied. Corrected the reader to match the writer.

### MEAC scheduled-task parameter rename

`Register-AuthCertificateRenewal` invoked the downloaded `MonitorExchangeAuthCertificate.ps1` with `-ConfigureScriptViaScheduledTask`, but the current CSS-Exchange release (BuildVersion 26.03.06.1531) renamed the parameter to `-ConfigureScriptToRunViaScheduledTask`. The MEAC call failed with `"A parameter cannot be found that matches parameter name 'ConfigureScriptViaScheduledTask'"`, the Auth-Cert auto-renewal scheduled task was therefore never registered, and Chapter 7.1 of the generated Word document consequently listed only the EXpress Log Cleanup task. Parameter name updated to match the upstream script.

---

## v5.86.1 (2026-04-23) — bugfix

### Phase 5 → Phase 6 spurious reboot

`Set-IPv4OverIPv6Preference` and `Disable-NetBIOSOnAllNICs` no longer set `$State['RebootRequired'] = $true`. Both tweaks are activated at the next boot, and the end-of-Phase-6 reboot (or the next natural reboot) covers them; forcing a mid-install reboot from Phase 5 only triggered the `Autopilot` resume path unnecessarily — especially on Exchange SE on WS2025 where no SU is currently published and Phase 5 otherwise has nothing reboot-worthy to do. The conditional skip-logic at the Phase 5→6 boundary (v5.85) now works as originally intended.

### Antispam agent install output

`Install-AntispamAgents` no longer emits five separate tables (one per `TransportAgent` object returned by the Exchange-shipped `Install-AntispamAgents.ps1`). Agent records are collected and rendered as a single compact summary list (`Identity / Priority / Enabled`). The well-known "Please restart the Microsoft Exchange Transport service for changes to take effect" warning is filtered (the function restarts the service itself immediately after). Remaining pipeline/Verbose/Debug/Information records are demoted from `Write-MyOutput` to `Write-MyDebug` so the standard transcript stays clean.

### Word installation document — TransportConfig null-ref

`Get-TransportConfig` returns `MaxSendSize` / `MaxReceiveSize` as `Unlimited` on fresh organisations, where `.Value` is `$null`. Both `New-InstallationDocument` (OpenXML Transport-Configuration table) and `New-InstallationReport` (HTML Exchange Optimizations table) now null-guard the `.Value.ToBytes()` access and fall back to `Unlimited / not set`. Without the guard the Word-document generation aborted with a null-valued-expression error just before the document finalization.

### Word installation document — Format-RemoteSysRows collection unwrap

`Format-RemoteSysRows` returned a `List[object[]]`, but PowerShell unwraps collections when they cross the function-return boundary, leaving the caller with a scalar `object[]` (single-row error-path case) or a plain `object[][]` — neither exposes `.ToArray()`, and `New-InstallationDocument` blew up at `$sysDetailRows.ToArray()` in section 5.x.2 "System details" for every server. Both `return` sites now prefix with the comma operator (`return ,$rows`) to preserve the List wrapper intact.

### MEAC download — wrong parameter name

`Register-AuthCertificateRenewal` called `Invoke-WebDownload -Url` but the function declares `-Uri`. PowerShell 5.1 silently ignores the unmatched `-Url` (the target parameter stays `$null`), so `System.Net.WebClient.DownloadFile('', $meacPath)` threw `"The path is not of a legal form"` from the empty URL argument. Changed the caller to `-Uri`. The CSS-Exchange `MonitorExchangeAuthCertificate.ps1` auto-renewal task now installs as intended.

### Antispam WARN filter + PSSnapin autoload noise

The WARN filter in `Install-AntispamAgents` was regexed against `"restart the Microsoft Exchange Transport"` but the actual Exchange message is `"restart is required for the change(s) to take effect : MSExchangeTransport"` — no match, so six "restart required" warnings plus three "Please exit Windows PowerShell to complete the installation." PSSnapin-autoload warnings always leaked to the console. Broadened to `'(restart is required|restart the Microsoft Exchange Transport|exit Windows PowerShell to complete)'`.

### Service-restart visibility (ECC/CBC/AMSI + antispam)

Batched W3SVC/WAS restart (after `Enable-ECC` / `Enable-CBC` / `Enable-AMSI` SettingOverride changes) and both MSExchangeTransport restarts in `Install-AntispamAgents` were logged only at `Write-MyVerbose` tier, so on a default run the console sat idle for 30–60s with no indication why. Promoted to `Write-MyOutput` with explicit "may take ~30/60s" hints, plus a matching "restarted" confirmation line. The corresponding `Enable-ECC` / `Enable-CBC` / `Enable-AMSI` "Enabling ..." headers were also promoted so the operator can see what triggered the restart.

---

## v5.86 (2026-04-23)

### Defender realtime monitoring + Tamper Protection

New `Disable-DefenderRealtimeMonitoring` / `Enable-DefenderRealtimeMonitoring` pair plus a best-effort `Disable-DefenderTamperProtection` / `Enable-DefenderTamperProtection`. Realtime scanning is disabled at the start of Phase 1 (before prerequisite installs) and re-enabled at the start of Phase 6 (after Exchange setup + SU). Setup generates massive file I/O (ECP/OWA .config unpacking, assembly ngen, transport agents) that Defender scans inline, which can stall or randomly fail setup with file-lock errors.

Tamper Protection (MDE/Intune) blocks `Set-MpPreference` silently. The function flips `HKLM:\SOFTWARE\Microsoft\Windows Defender\Features\TamperProtection = 0` as best-effort, verifies via `Get-MpComputerStatus.IsTamperProtected` whether it took effect, and warns the operator if Intune/MDE still enforces. Original registry value is captured in `$State['DefenderTPPrev']` and restored on re-enable. `$State['DefenderRealtimeDisabledByEXpress']` ensures re-enable only happens when we were the ones who disabled it; a GPO/Intune tamper-back is detected silently.

Accepted trade-off: a GPO/Intune policy may re-enable realtime during the install window. On MDE-managed hosts Tamper Protection cannot be programmatically cleared — operator must disable via Intune policy first.

### Network-layer hardening: LLMNR + mDNS

`Disable-LLMNR` sets `HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient\EnableMulticast = 0` (CIS L1 §18.5.4.2). `Disable-MDNS` sets `HKLM:\SYSTEM\CurrentControlSet\Services\Dnscache\Parameters\EnableMDNS = 0` (WS2022+ default-on). Both close the Responder-class name-spoofing / NTLM-hash-capture vectors, alongside the existing NetBIOS-over-TCP/IP disable in Phase 5.

### MEAC — auto-renewal of Exchange Auth Certificate

`Register-AuthCertificateRenewal` downloads CSS-Exchange `MonitorExchangeAuthCertificate.ps1` and registers the daily scheduled task via `-ConfigureScriptViaScheduledTask`. Task runs as SYSTEM, renews the Auth Certificate 60 days before expiry. Without MEAC, Auth Cert expiry causes a full outage (OAuth, Hybrid, EWS push subscriptions). Runs in Phase 6 after `Test-AuthCertificate`; skipped on Edge and management-only installs.

### Menu toggle [U]: Generate Installation Document

Word Installation Document (F22) is no longer generated by default from the menu. Toggle **[U] Generate Installation Document** defaults to off; enable it only when the document is needed. Mode 7 (Installation Document only) still always generates the doc regardless of the toggle. CLI parameter `-NoWordDoc` and config-file behavior unchanged.

### SMTP Banner on anonymous relay connectors

`New-AnonymousRelayConnector` now sets `-Banner '220 Mail Service'` on both the internal and external relay connectors at creation/update time. Before, Phase 5's banner-hardening step ran before the relay connectors existed, leaving them with the default Exchange-version-disclosing banner. The Installation Report now correctly shows 5/5 hardened Frontend connectors on a typical install.

### Extended Protection (OWA): filter to Frontend VDir

`New-InstallationReport` was picking whichever OWA virtual directory came first from `Get-OwaVirtualDirectory` — on typical Exchange SE installs that's the Back End site, which has `ExtendedProtectionTokenChecking = None` by design. The report now explicitly filters to the frontend VDir (`WebSiteName -notlike '*Back End*'`) and renders the site name plus EP value in the Security table.

### Enable-UAC / Enable-IEESC moved before report generation

Previously ran at the very end of Phase 6, after `New-InstallationReport` and `New-InstallationDocument`, so reports captured UAC/IE ESC as disabled. Moved to the top of the report block so the captured security state reflects the final hardened configuration.

### Improved Word-document error diagnostics

Phase 6 `New-InstallationDocument` catch block now writes the failing script line number and a stack-trace verbose line, mirroring the behavior already present for `New-InstallationReport`. Makes Word-doc aborts actionable instead of a single-line "Word document failed: {message}".

### Autodiscover AppPool job: suppressed table output

`Start-DisableMSExchangeAutodiscoverAppPoolJob` returns the Job object from `Start-Job`; the caller now discards it with `$null = …` to suppress the default Id/Name/State/HasMoreData/Location/Command table that PowerShell otherwise writes to the console.

### Start-EXpress.cmd quickstart launcher

New batch file in the repo root that opens an elevated PowerShell (`-NoExit -ExecutionPolicy Bypass -NoProfile -Debug`), changes to `C:\install`, and invokes `Install-Exchange15.ps1`. Additional arguments are forwarded. Window stays open after script exit for inspection.

---

## v5.85 (2026-04-22)

### Conditional Phase 2→3 reboot

New helper `Test-RebootPending` inspects Windows' standard pending-reboot signals:

- `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending`
- `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired`
- `HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations`
- Pending computer rename (`ActiveComputerName` ≠ `ComputerName`)
- CCM ClientSDK `DetermineIfRebootPending` (SCCM-managed hosts)

After Phase 2 completes in Autopilot, the main phase loop calls `Test-RebootPending`. If nothing is pending, `InstallPhase` is advanced to 3 and `Save-State` persists the advance; the loop re-enters the `switch` and runs Phase 3 in the same process — no reboot, no RunOnce hop.

**When this saves a reboot:** WS2025 + Exchange SE (or any combination where .NET 4.8.1 is already present). VC++ 2012/2013 and URL Rewrite don't set reboot flags. On WS2016/WS2019/WS2022 where .NET 4.8 / 4.8.1 is actually installed, CBS or PendingFileRenameOperations will be set and the existing reboot path runs unchanged.

**Crash recovery:** because `LastSuccessfulPhase=2` and `InstallPhase=3` are saved before Phase 3 runs, a Phase 3 crash still resumes correctly at Phase 3 on the next start.

### Conditional Phase 5→6 reboot

Same mechanism applied to the Phase 5 → Phase 6 boundary: the reboot runs only when `$State['RebootRequired']` is set (Exchange SU installer returned exit code 3010) **or** `Test-RebootPending` reports a pending reboot from any other source. Phase 5 otherwise only writes registry values and IIS settings that don't require a reboot, so HealthChecker / Installation Report / Word Document (Phase 6) can run immediately in the same process.

### Antispam install output routing

`Install-AntispamAgents` no longer redirects all streams from `Install-AntispamAgents.ps1` into `Out-Null`. The previous approach (`*>&1 | Out-Null` + `$WarningPreference = 'Ignore'` + `$PSDefaultParameterValues['*:WarningAction']`) had two problems: it suppressed the legitimate output — the agent-configuration table we want to keep — and it did not suppress warnings emitted via `$host.UI.WriteWarningLine` (which bypass the pipeline).

The new approach captures all streams into a single merged record list and routes by record type:

- `WarningRecord` / `VerboseRecord` / `DebugRecord` → `Write-MyVerbose` (log only, no console)
- `ErrorRecord` → `Write-MyWarning`
- `InformationRecord` → `Write-MyOutput`
- Everything else (the agent-configuration table, success messages) → `Write-MyOutput` with `[antispam]` prefix

The `$WarningPreference` / `$PSDefaultParameterValues` gymnastics are removed — record-level filtering is deterministic and does not rely on stream preferences that Exchange scripts reset locally.

### Phase 7 merged back into Phase 6

`MAX_PHASE` reverted to 6; HealthChecker, Installation Report, and Word Installation Document all run at the end of Phase 6 again. The short-lived Phase 7 split (introduced for the "Exchange Server Membership" HC finding on same-session runs) is removed — the real cause is Domain-Controller specific and is now surfaced through a clarifying warning instead.

### HealthChecker on Domain Controllers

`Invoke-HealthChecker` now queries `Win32_ComputerSystem.DomainRole` and emits an explicit warning when the install target is a Domain Controller (role 4/5). HC's "Exchange Server Membership" check enumerates `Win32_GroupUser` for the local "Exchange Servers" / "Exchange Trusted Subsystem" groups, which don't exist on DCs — the SAM database is replaced by AD. The finding is therefore not actionable on DCs and the warning makes that explicit.

The previous `klist purge` workaround (attempted token refresh) is removed; it never addressed the real cause.

### VC++ 2013 update URL

`Install-MyPackage` for VC++ 2013 now uses `https://aka.ms/highdpimfc2013x64enu` which delivers 12.0.40664 (High-DPI aware MFC). The legacy GUID CDN URL (`CC2DF5F8-4454-44B4-802D-5EA68D086676/vcredist_x64.exe`) delivered 12.0.40660, which HealthChecker flags as outdated per `aka.ms/HC-LatestVC`. `Get-VCRuntime -version '12.0' -MinBuild '12.0.40664'` gates the install.

### Word document — `New-WdTable` hardening

PS 5.1 flattens literal `@(@('a','b'), @('c','d'))` to a single flat array when bound to `[object[][]]`. `New-WdTable` now:

- Wraps each `$row` in `@($row)` so both `object[][]` (correct) and flattened-scalar (buggy) input produce valid rows
- Pads short rows to the header width with empty cells

The "Offene Punkte" callsite in `New-InstallationDocument` additionally uses the `,@(...)` prefix on each row literal.

### VERBOSE console spam after Autopilot resume

`$VerbosePreference = 'SilentlyContinue'` and `$DebugPreference = 'SilentlyContinue'` are now pinned at the very top of `process{}`, before any `Get-CimInstance` call. Previously the first `Get-CimInstance Win32_OperatingSystem` ran with the default `$VerbosePreference='Continue'` (set by PowerShell when `-Verbose` was on the command line — preserved across RunOnce relaunches), which spammed `VERBOSE: Perform operation 'Enumerate CimInstances' ...` to the console but not to the tiered log.

Custom tier-aware logging (`Write-MyVerbose` / `Write-MyDebug`) is independent of the stream preferences — driven by `$State['LogVerbose']` / `$State['LogDebug']` flags set a few lines further down.

### Menu confirmation

Download Domain value is now logged next to Namespace in the post-menu summary, so an empty or mistyped entry is visible in the log for troubleshooting HC's Download Domains finding.

---

## v5.84 (2026-04-22)

### F22 — Installation Documentation: org-wide + all servers (remote query)

**Scope expansion:** `New-InstallationDocument` now documents the entire Exchange organisation, not just the local server. The document title page labels the run with one of three scenarios:

1. **New Exchange environment** — a single server, organisation freshly created
2. **Server addition** — new server joining an existing organisation; all servers documented plus the newly installed one
3. **Ad-hoc inventory** (`-StandaloneDocument` without a prior setup run) — pure stock-take; chapters 2 / 7 / 14 are omitted

**New chapter structure (16 chapters):**

- Chapter 4 — "Organisation — cross-cutting configuration" (new): Org-Config, Accepted/Remote Domains, email address policies, transport rules, transport configuration, journal/DLP/retention, mobile/OWA policies, DAGs (all, with database copies), send connectors, federation/hybrid/OAuth, AuthConfig
- Chapter 5 — "Servers in the organisation" (new): loop over `Get-ExchangeServer`; per server: identity, system details, databases, virtual directories, receive connectors, certificates, transport agents; local server marked "← newly installed"
- Chapters 6–16: network/DNS (local), Exchange installation (local), hardening, agents, backup/DR, HealthChecker, monitoring, public folders, cmdlets, runbooks, open items

**Remote query standard (CIM/WSMan):** `Get-RemoteServerData` gathers hardware / OS / pagefile / volume / NIC data from remote Exchange servers via CIM over WSMan (WinRM TCP 5985/5986, Kerberos). WMI/DCOM is intentionally avoided. Local server: direct CIM, no WinRM prompt.

**Interactive prompt on failure:** `Invoke-RemoteQueryWithPrompt` — on WinRM failure, shows a remediation hint and offers `[R] Retry / [S] Skip` with a 10-minute auto-skip countdown (`Write-Progress -Id 2`); under Autopilot or non-interactive sessions: silent skip. Skipped servers are flagged inline in the document.

**New data helpers:**

- `Get-OrganizationReportData` — org-wide settings (queried once)
- `Get-ServerReportData -ServerName` — per-server Exchange cmdlets + remote CIM
- `Get-InstallationReportData -Scope -IncludeServers` — aggregator consumed by both report functions

**New parameters:**

- `-DocumentScope All|Org|Local` (default `All`) — document depth selector
- `-IncludeServers <Name[]>` — filter to a subset of servers in large farms

**New files:**

- `tools/Enable-EXpressRemoteQuery.ps1` — idempotent prerequisite script for target servers (HTTP/HTTPS listener, optional AD-group ACL)
- `docs/remote-query-setup.md` — GPO walkthrough, firewall matrix, failure modes, hardening recommendations

---

## v5.83 (2026-04-22)

### Three-tier logging (single log file, tier-controlled via `-Verbose` / `-Debug`)

`Write-ToTranscript` writes a single log file with three tiers; each line is tagged with its tier prefix (`INFO` / `WARNING` / `ERROR` / `EXE` / `VERBOSE` / `DEBUG`):

| Invocation | Tiers written |
|---|---|
| `.\Install-Exchange15.ps1` (default) | `INFO`, `WARNING`, `ERROR`, `EXE` |
| `.\Install-Exchange15.ps1 -Verbose` | + `VERBOSE` |
| `.\Install-Exchange15.ps1 -Debug` | + `DEBUG` + `SUPPRESSED-ERROR` lines |

**`SUPPRESSED-ERROR` in debug mode:** `Write-ToTranscript` snapshots `$Error.Count` and reconstructs every error that appeared since the previous call — including those silently swallowed by `try/catch`. Line format: `[SUPPRESSED-ERROR] (Exception.Type) at line N: <offending line> :: <message>`. Essential for diagnosing BITS / MSI / CIM failures that would otherwise leave no trace.

**Encoding fix:** PS 5.1 `Out-File` defaults to UTF-16 LE with BOM, which combined with the UTF-8 header produced "strange font" rendering in viewers. Now `[System.IO.File]::AppendAllText(..., UTF8Encoding($false))` — UTF-8 without BOM, every line written with the same encoding.

Tiers are activated via the standard PowerShell `-Verbose` / `-Debug` switches on the `.ps1` invocation, tracked internally in `$State['LogVerbose']` / `$State['LogDebug']`. On Autopilot resume via `RunOnce`, the flags are passed through to the resumed process (`$logFlags`), so the selected log tier survives reboots.

New helper `Write-MyDebug` — console stays silent; the log line only appears when the debug tier is active.

### Unified file nomenclature

All artefacts generated by the script now follow the `{PC}_{Tag}_{yyyyMMdd-HHmmss}.{ext}` schema:

- Transcript / log (including `{PC}_InstallLog_*.log`)
- Preflight report (now timestamped — previously overwritten on each run)
- Installation report (HTML)
- Installation document (Word)
- RBAC report
- Exported server config (XML)
- Log cleanup protocol

Consistent sorting in Explorer; all artefacts from a single installation run are identifiable at a glance by their shared timestamp.

### Credential prompt fix

`Get-ValidatedCredentials`: deterministic GUI-vs-Read-Host decision via `$env:SESSIONNAME` check (console vs. RDP). Previously `Get-Credential` could return `$null` in rare cases and the script would silently continue — now the input modality is chosen up front, and the prompt is re-issued on cancel.

### Bootstrap order

- Log initialisation now runs **before** the menu is drawn — menu interactions are captured in the log
- `$script:isFreshStart` snapshot prevents early state mutations from being misclassified as an Autopilot resume

### Dev tools

- `tools/Test-ScriptSanity.ps1` — 14 structural sanity checks (param block, process block, encoding, ternary syntax, function count, …)
- `tools/Test-ScriptQuality.ps1` — qualitative checks
- `tools/Fix-IfAsArg.ps1` — fixer for the PS 5.1 `-f (if …)` syntax trap
- `tools/Fix-PhaseNum.ps1` — phase-number consistency enforcer
- `tools/Parse-Check.ps1` — fast parser check without execution

---

## v5.82 (2026-04-21)

- F22: `New-InstallationDocument` — generates a Word (.docx) installation report after Phase 6 using a pure-PowerShell OpenXML engine (no Office/COM); 15 chapters covering installation parameters, system details, network, AD, Exchange configuration, hardening, backup readiness, HealthChecker, monitoring, hybrid, public folders, executed cmdlets, and runbooks; CustomerDocument mode redacts RFC1918 IPs, certificate thumbprints, and passwords
- F22: New parameters `–NoWordDoc`, `–StandaloneDocument`, `–CustomerDocument`, `–Language` (DE/EN)
- F22: `–StandaloneDocument` mode (menu mode 7) runs Phase 1-only: loads Exchange module and generates document on existing servers without full install
- F23: `tools/Build-ConceptTemplate.ps1` — pure-PowerShell OpenXML generator for DE + EN concept / approval document templates; 16 chapters (architecture, sizing, security, migration, hybrid, compliance, questionnaire, approval page); output: `templates/Exchange-concept-template-DE.docx` + `…-EN.docx`; Exchange SE only (2016/2019 out-of-support since 14.10.2025)

---

## v5.81 (2026-04-21)

- `New-InstallationReport` (B17 complete): root cause identified — `New-HtmlSection`, `Format-Badge`, `Format-RefLink`, and all `$*Content` assembly lines used `-f` operator with dynamic HTML content; `String.Format` throws `FormatException` whenever any Exchange data value (connector name, cert SAN, policy value, etc.) contains a `{n}` sequence; all seven affected call sites converted to string concatenation — report is now fully immune to user-defined data containing curly braces

---

## v5.80 (2026-04-21)

- `New-InstallationReport` (B17): `$exContent` HERE-STRING with `-f` threw `FormatException` (`String.Format` index out of range) when any collected HTML row contained curly-brace patterns (e.g. CSS `{color:...}` or Exchange policy values); replaced `-f` string formatting with direct string concatenation — immune to content containing `{n}` sequences
- `Invoke-HealthChecker`: HC report renamed to `SERVER_HCExchangeServerReport-<timestamp>.html` (was `SERVER_ExchangeAllServersReport-...`); `Where-Object` filter and `Rename-Item` logic updated; all three known HC output prefixes (`ExchangeAllServersReport`, `HealthChecker`, `HCExchangeServerReport`) detected to handle existing reports and future HC versions

---

## v5.79 (2026-04-21)

- `New-InstallationReport` (B16): four defects fixed:
  1. **Wrong encoding** — transcript file was read with explicit `UTF-8`; PS 5.1 writes transcripts as UTF-16 LE; removed explicit encoding so .NET auto-detects the BOM (handles both UTF-8 and UTF-16 transcripts)
  2. **No try/catch on log reading** — an `IOException` (large/locked file) propagated to the global `trap { break }` and killed the entire script during Phase 6 report generation; log section now wrapped in `try/catch`
  3. **Log size not capped** — a transcript accumulated over multiple reboots (30 h+ install with 3× Phase 5 loops) could be several MB; the full content embedded in `<pre>` made the HTML report unusably large; capped to last 2 000 lines with truncation notice and full path
  4. **Call site unprotected** — `New-InstallationReport` called without try/catch in Phase 6; any uncaught exception inside the function killed the script before "We're good to go" and phase-end logic ran; call site now wrapped in `try/catch` that logs a warning and continues

---

## v5.78 (2026-04-21)

- `Install-ExchangeSecurityUpdate` (B15): Exchange SU installer (`.exe`) may call `ExitWindowsEx` internally and reboot the machine before the script's phase-end logic runs (`LastSuccessfulPhase` update + `Enable-RunOnce`); in Autopilot mode, `RunOnce` + state are now persisted **before** launching the installer so the script always auto-resumes after an installer-triggered reboot; a per-KB flag `ExchangeSUInstalled_<KB>` is stored in state after successful install (rc 0/3010) so phase-5 re-entry skips the SU even when `Get-InstalledExchangeBuild` still returns the pre-SU build number (service binary cache not yet flushed after reboot)

---

## v5.77 (2026-04-21)

- `Install-ExchangeSecurityUpdate` (B14): removed `/norestart` from Exchange SU installer arguments; Exchange SU `.exe` only supports `/passive` and `/silent` — `/norestart` caused the installer to abort immediately with "The following command line options are not recognized: /norestart"; exit code 3010 (reboot required) is already handled correctly

---

## v5.76 (2026-04-21)

- `Enable-AMSI`: added Edge Transport guard — `New-SettingOverride` requires an Exchange org connection; Edge is standalone/not domain-joined, so AMSI body scanning via SettingOverride is not applicable
- `Set-MaxConcurrentAPI`: added Edge Transport guard — Netlogon `MaxConcurrentApi` is a domain-authentication optimization; Edge is not domain-joined
- `Set-CtsProcessorAffinityPercentage`: added Edge Transport guard — Exchange Search registry path does not exist on Edge Transport
- `Set-NodeRunnerMemoryLimit`: added Edge Transport guard — NodeRunner (Exchange Search) does not run on Edge Transport
- `Test-AuthCertificate` (B12): added null-guard for `$authConfig` before `.CurrentCertificateThumbprint` access; `Get-AuthConfig` can return `$null` when Exchange PS session is not fully initialized (observed after IIS restart in Phase 6), previously causing a "You cannot call a method on a null-valued expression" error in the catch block
- `New-AnonymousRelayConnector` (B13): fixed race condition — `Get-ReceiveConnector` called immediately after `New-ReceiveConnector` failed because Exchange had not yet registered the connector object; connector is now taken directly from the `New-ReceiveConnector` return value; 3-attempt × 5 s retry fallback added for the edge case where the return value is null

---

## v5.75 (2026-04-21)

- `Initialize-Exchange`: returns `$true`/`$false` to indicate whether `setup.exe /PrepareAD` actually ran; exits early with `$false` when org exists and both FFL and DFL already meet the minimum — no setup.exe invoked, no wait
- Phase 3: `Wait-ADReplication` only called when `Initialize-Exchange` returns `$true`; progress label changed from "Preparing Active Directory" to "Checking Active Directory" to reflect the conditional nature

---

## v5.74 (2026-04-21)

- `Enable-AMSI`: removed Exchange SE exception — HealthChecker always checks for the `AmsiRequestBodyScanning` SettingOverride regardless of version defaults; SettingOverride is now applied for all Exchange versions when `-EnableAMSI` is used
- `Invoke-HealthChecker`: added NOTE output after HC run explaining that "Exchange Server Membership" may show blank/failed results in the same-session run due to Kerberos token not yet including the new "Exchange Servers" group membership — accurate after next reboot

---

## v5.73 (2026-04-21)

- `Install-AntispamAgents`: replace `$WarningPreference = 'SilentlyContinue'` with `$PSDefaultParameterValues['*:WarningAction'] = 'Ignore'` for the install-script call — this has higher precedence than `$WarningPreference` and cannot be overridden by the called script's own preference reset; also upgrade to `'Ignore'`
- `Install-AntispamAgents`: `Enable/Disable-TransportAgent` calls changed from `$null = ...` (Stream 1 only) to `*>&1 | Out-Null` with `-WarningAction Ignore` — properly suppresses Stream 3 warnings

---

## v5.72 (2026-04-20)

- `Invoke-HealthChecker`: added `-BuildHtmlServersReport` call after data collection so `ExchangeAllServersReport-*.html` is generated in `ReportsPath`
- `New-InstallationReport`: HealthChecker section re-added (iframe + direct link to HC HTML; shows "skipped" or "not found" messages when appropriate); TOC entry added

---

## v5.71 (2026-04-20)

- `Install-ExchangeSecurityUpdate`: SU file-placement countdown now checks `ConfigDriven` instead of `Autopilot`; countdown was incorrectly skipped when the auto-reboot toggle was on in an interactive (Copilot) session — user is present and needs to place the EXE

---

## v5.70 (2026-04-20)

Broken / stale link fixes in README.md and `New-InstallationReport` HTML output:
- Extended Protection URL (404) → `exchange-extended-protection` (MS Learn)
- SSL Offloading URL (404) → same `exchange-extended-protection` page (topic merged)
- 2022 H1 CU blog post `ba-p/3285209` (wrong content) → corrected `ba-p/3285026`
- TLS 1.2 guide Part 2 `ba-p/607646` (wrong content) → corrected `ba-p/607761`
- TLS 1.3 blog `ba-p/3777803` (wrong content) → `support.microsoft.com` KB article
- IPv6 TechCommunity post `ba-p/594506` (wrong content) → MS Learn `configure-ipv6-in-windows`
- `docs.microsoft.com` → `learn.microsoft.com` (redirect cleanup, README only)

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
