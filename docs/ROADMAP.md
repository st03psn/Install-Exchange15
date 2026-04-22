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

## Active Bugs

_None open._ Bugs B1–B17 (v5.6.x / v5.7x) and Planned features P1–P8 are all resolved; see `RELEASE-NOTES.md` for the per-version fix trail.

---

## Feature Backlog (F3, F24)

| # | Feature | Phase | Priority |
|---|---|---|---|
| F3 | Split Permissions | 0 | LOW — irreversible, niche use case |
| F24 | Installation-Document Template (hybrid) | 6 | MEDIUM — target v1.0 together with modularization. Convert `New-InstallationDocument` from fully code-driven OpenXML to a "style-shell + dynamic body" hybrid: ship `templates/Exchange-installation-document-{DE,EN}.docx` containing cover page (`[LOGO]` placeholder, SDT for Org/Server/Date/Scenario), header/footer with classification, `word/styles.xml` + `word/theme1.xml` + `word/numbering.xml`, and empty body-anchor SDTs (tags `body_chapter_1`, `body_chapter_5`, …). Script copies template, injects generated chapter XML into the anchor SDTs, writes output. New param `-TemplatePath <path>` for customer templates. Benefits: layout / branding / font / colour edits without PowerShell; corporate design trivial. Risks: template-schema brittleness → `Test-Template` validator required; two languages = two templates to maintain. Estimated effort: 1–2 days; lands naturally with modularization since engine moves to `src/74-OpenXml.ps1`. |

### Completed features (F6–F23)

F6 Extended Protection · F7 Auth Certificate Monitoring · F8 DAG Replication Health · F9 VSS Writer Validation · F10 EEMS Status Check · F11 Modern Auth Report · F12 Remove MSMQ · F13 Disable SSL Offloading · F14 OWA Download Domains · F15 AMSI Body Scanning · F16 IanaTimeZoneMappings · F17 Root AutoUpdate · F18 Disable MRS Proxy · F19 MAPI Encryption Required · F20 SetupLogReviewer on failure · F21 Disable Customer Feedback · F22 Word Installation Documentation · F23 Concept / Approval Document.
