# EXpress Roadmap

Active bugs, v1.0 rename + modularization checklist, and feature backlog.

## v1.0 — EXpress Rename + Modularization (planned)

**Approved announcement text:**
> EXpress stands on the shoulders of giants. Michel de Rooij's Install-Exchange15.ps1 laid the groundwork — a solid, community-proven script that guided Exchange deployments for years. EXpress takes that legacy forward: full automation, security hardening, and a modern deployment experience. Thank you, [Michel](http://eightwone.com).

**Versioning:** Jump to 1.0 (new identity + architectural split — not a continuation of v5.x).

**Scope:** Bundles two invasive structural changes in one release — the rename justifies the version jump, and the modularization benefits from the rename (new filename structure already affects all generated files, so file/path churn happens once).

### Rename Checklist
- [ ] GitHub repo rename: `Install-Exchange15` → `EXpress`
- [ ] `git remote set-url origin https://github.com/st03psn/EXpress.git`
- [ ] `git mv Install-Exchange15.ps1 EXpress.ps1`
- [ ] `git mv Install-Exchange15.Tests.ps1 EXpress.Tests.ps1`
- [ ] `$StateFile`: `_State.xml` → `_EXpress_State.xml`
- [ ] All generated filenames: add `EXpress` as second segment — `{PC}_EXpress_{Tag}_...` (Install, Preflight, Report, Document, RBAC, Config, LogCleanup)
- [ ] `$ScriptVersion = '1.0'`
- [ ] `.SYNOPSIS` / `.DESCRIPTION` / `.EXAMPLE` docblock: update filename + EXpress branding
- [ ] `Build.ps1`: input/output filename + `-SkipMerge` switch (see modularization)
- [ ] `README.md`: all filename + URL references; add Michel de Rooij acknowledgement section
- [ ] `CLAUDE.md`: filename references + document `src/` structure
- [ ] `deploy-example.psd1`: filename in comments
- [ ] `docs/index.html`: download link, GitHub URL, version badge, footer updated to EXpress v1.0 (April 2026) ✅
- [ ] `docs/plan-F22-F23-word-documentation.md`: path references to `EXpress.ps1` and new `src/` structure

### F24 Checklist — Installation-Document Template (see `plan-modularization.md` for full detail)
- [ ] Template generator (extend `tools/Build-ConceptTemplate.ps1` or new `Build-InstallationTemplate.ps1`)
- [ ] `templates/Exchange-installation-document-DE.docx` + `-EN.docx` committed
- [ ] Engine helpers in `src/74-OpenXml.ps1`: `Open-WdTemplate`, `Set-WdContentControl`, `Merge-WdBodyAnchor`, `Test-WdTemplate`
- [ ] `New-InstallationDocument` refactor: generate chapter body XML → inject into anchor SDTs (`body_chapter_1` … `body_chapter_16`)
- [ ] New parameter `-TemplatePath <path>` for customer template override
- [ ] `Test-WdTemplate` validation runs before every generation; clear error on missing tags
- [ ] README + RELEASE-NOTES document the template customization workflow

### Modularization Checklist (see `plan-modularization.md` for full detail)
- [ ] `src/` directory with 25 modules (`00-Constants.ps1` … `99-Main.ps1`, numeric prefixes for load order)
- [ ] `tools/Merge-Source.ps1` — produces release `EXpress.ps1` from `src/*.ps1`
- [ ] `EXpress.ps1` entry: `param()` + `process{}` with `#region SOURCE-LOADER` (dev mode dot-sources `src/*.ps1`, release mode is merged by Merge-Source)
- [ ] `Build.ps1` calls Merge-Source before PS2Exe (`-SkipMerge` as escape hatch)
- [ ] Hash verification after step 1: merge output byte-identical with existing `Install-Exchange15.ps1`
- [ ] CI guard (pre-commit or GitHub Action): merge + `git diff --exit-code` on release `.ps1`
- [ ] Extract module by module (1 commit per module); after each step: `tools/Parse-Check.ps1` + `tools/Test-ScriptSanity.ps1` + Autopilot smoketest

---

## Active Bugs (v5.6.x / 5.7x)

| # | Bug | Detail |
|---|---|---|
| ~~B1~~ | ~~External relay connector conflict~~ | ~~Fixed: external placeholder changed to `192.0.2.2/32`~~ ✅ |
| ~~B2~~ | ~~HealthChecker output in install log~~ | ~~Fixed: HC suppressed with `*>&1 \| Out-Null`~~ ✅ |
| ~~B3~~ | ~~HealthChecker "not found" in report~~ | ~~Fixed: `-OutputFilePath` now uses `Join-Path ... 'HealthChecker'` as file prefix~~ ✅ |
| ~~B4~~ | ~~CPU count wrong in report~~ | ~~Fixed: sum `NumberOfCores`/`NumberOfLogicalProcessors` across all `Win32_Processor` instances~~ ✅ |
| ~~B5~~ | ~~Autodiscover SCP not derived from namespace~~ | ~~Fixed: SCP auto-derived as `autodiscover.<parent-domain>` from `-Namespace`~~ ✅ |
| ~~B6~~ | ~~Windows Update Enter = N~~ | ~~Fixed: Enter now defaults to Y~~ ✅ |
| ~~B7~~ | ~~Exchange SU always skipped~~ | ~~Fixed: `Install-ExchangeSecurityUpdate` checked `InstallWindowsUpdates` instead of `IncludeFixes`~~ ✅ |
| ~~B8~~ | ~~VC++ 2013 wrong download URL~~ | ~~Fixed: URL was delivering VC++ 2010; corrected to `CC2DF5F8-4454-44B4-802D-5EA68D086676/vcredist_x64.exe` (12.0.40664)~~ ✅ |
| ~~B9~~ | ~~VC++ detection fails for 2013~~ | ~~Fixed: `Get-VCRuntime` now has three-stage detection: VisualStudio registry → Add/Remove Programs display name → `System32\msvcr120.dll`~~ ✅ |
| ~~B10~~ | ~~BITS `0x800704DD` causes noisy retries~~ | ~~Fixed: `Get-MyPackage` detects `ERROR_NOT_LOGGED_ON` and falls back to WebDownload immediately~~ ✅ |
| ~~B11~~ | ~~WU prompt immediately skips to N~~ | ~~Fixed: added `FlushInputBuffer()` before RawUI prompt loop to clear buffered keystrokes from prior credential prompts~~ ✅ |
| ~~B12~~ | ~~`Test-AuthCertificate` null-valued expression~~ | ~~Fixed: added null-guard for `$authConfig` before `.CurrentCertificateThumbprint` access; `Get-AuthConfig` can return `$null` when Exchange PS session not fully initialized after IIS restart~~ ✅ |
| ~~B13~~ | ~~External relay connector race condition~~ | ~~Fixed: `New-AnonymousRelayConnector` now captures object returned by `New-ReceiveConnector` directly (no second `Get-ReceiveConnector` call); added 3-attempt/5s retry fallback for edge case where object is null~~ ✅ |
| ~~B14~~ | ~~Exchange SU `/norestart` not recognized~~ | ~~Fixed: removed `/norestart` from `Install-ExchangeSecurityUpdate` — Exchange SU `.exe` only accepts `/passive` and `/silent`~~ ✅ |
| ~~B15~~ | ~~Exchange SU installer self-reboot breaks Autopilot~~ | ~~Fixed: `RunOnce` + state persisted before installer launch; per-KB `ExchangeSUInstalled_<KB>` flag prevents reinstall loop on phase-5 re-entry~~ ✅ |
| ~~B16~~ | ~~`New-InstallationReport` crashes script via global trap~~ | ~~Fixed: wrong UTF-8 encoding on UTF-16 LE transcript; no try/catch on log read; no size cap; call site unprotected — all four defects fixed~~ ✅ |
| ~~B17~~ | ~~`New-InstallationReport` FormatException in Exchange section~~ | ~~Fixed: `$exContent` HERE-STRING with `-f` threw `String.Format` index-out-of-range when HTML rows contained `{n}` patterns (CSS, policy values); converted to string concatenation~~ ✅ |

---

## Planned Features (v5.6.x / v5.7)

| # | Feature | Detail |
|---|---|---|
| ~~P1~~ | ~~Receive connectors IPs in report~~ | ~~`New-InstallationReport`: show `RemoteIPRanges` per connector in section~~ ✅ |
| ~~P2~~ | ~~Pagefile status in report~~ | ~~Evaluate pagefile config vs. RAM (Exchange recommendation: RAM + 10 MB, min 32 GB)~~ ✅ |
| ~~P3~~ | ~~"-not tested-" in menu~~ | ~~Modes: DAG Join, Copy Server Config, PFX Certificate — mark as untested in menu UI~~ ✅ |
| ~~P4~~ | ~~Input validation~~ | ~~`Read-MenuInput` for namespace (valid FQDN), IP/CIDR subnets — re-prompt on invalid~~ ✅ |
| ~~P5~~ | ~~deploy-example.psd1~~ | ~~Add missing params: `SkipInstallReport`, `SkipSetupAssist`~~ ✅ |
| ~~P6~~ | ~~Dynamic SU detection via HC.ps1~~ | ~~`Get-LatestSUBuildFromHC` parses `GetExchangeBuildDictionary` in HC.ps1; `Get-InstalledExchangeBuild` reads running service version; `Install-ExchangeSecurityUpdate` warns when installed < HC latest. Exchange 2019 CU14/CU15 + 2016 CU23 Feb26SU require ESU enrollment (no public URL) — KB5074993/KB5074994/KB5074995~~ ✅ |
| ~~P7~~ | ~~Compliance mapping in installation report~~ | ~~Option 1 implemented: CIS / BSI control-ID column added to Security section~~ ✅ |
| ~~P8~~ | ~~Phase timing for phases 1–4~~ | ~~`$phSw = [Diagnostics.Stopwatch]::StartNew()` added at start of each phase~~ ✅ |

---

## Feature Backlog (F3, F24)

| # | Feature | Phase | Priority |
|---|---|---|---|
| F3 | Split Permissions | 0 | LOW — irreversible, niche use case |
| F24 | Installation-Document Template (hybrid) | 6 | MEDIUM — target v1.0 together with modularization. Convert `New-InstallationDocument` from fully code-driven OpenXML to a "style-shell + dynamic body" hybrid: ship `templates/Exchange-installation-document-{DE,EN}.docx` containing cover page (`[LOGO]` placeholder, SDT for Org/Server/Date/Scenario), header/footer with classification, `word/styles.xml` + `word/theme1.xml` + `word/numbering.xml`, and empty body-anchor SDTs (tags `body_chapter_1`, `body_chapter_5`, …). Script copies template, injects generated chapter XML into the anchor SDTs, writes output. New param `-TemplatePath <path>` for customer templates. Benefits: layout / branding / font / colour edits without PowerShell; corporate design trivial. Risks: template-schema brittleness → `Test-Template` validator required; two languages = two templates to maintain. Estimated effort: 1–2 days; lands naturally with modularization since engine moves to `src/74-OpenXml.ps1`. |

### Completed features (F6–F23)

F6 Extended Protection · F7 Auth Certificate Monitoring · F8 DAG Replication Health · F9 VSS Writer Validation · F10 EEMS Status Check · F11 Modern Auth Report · F12 Remove MSMQ · F13 Disable SSL Offloading · F14 OWA Download Domains · F15 AMSI Body Scanning · F16 IanaTimeZoneMappings · F17 Root AutoUpdate · F18 Disable MRS Proxy · F19 MAPI Encryption Required · F20 SetupLogReviewer on failure · F21 Disable Customer Feedback · F22 Word Installation Documentation · F23 Concept / Approval Document.
