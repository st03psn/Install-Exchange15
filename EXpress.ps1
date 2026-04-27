<#
    .SYNOPSIS
    EXpress — unattended Exchange Server 2016/2019/SE installation, hardening,
    post-configuration, documentation, and day-2 standalone modes.

    Script file: EXpress.ps1
    Version:     1.3.0
    Maintainer:  st03psn

    Original author: Michel de Rooij (michel@eightwone.com).
    EXpress stands on the shoulders of giants — Michel's Install-Exchange15.ps1
    laid the groundwork for community-driven Exchange deployment automation.
    Original copyright and license notices are preserved.

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    .DESCRIPTION
    EXpress runs one of several modes, selected by parameter set:

      • Full install (default) — Exchange Mailbox role, prerequisites, AD prep
        (optional), Exchange setup, security hardening, post-config (virtual
        directory URLs, certificate import, DAG join, connectors), Exchange
        Security Update, Word installation document, HTML installation report.
      • Edge (-InstallEdge) — Edge Transport role.
      • NoSetup — prerequisites + OS hardening, skip Exchange setup.
      • Recover (-Recover) — Exchange RecoverServer mode.
      • Recipient Management Tools (-InstallRecipientManagement) — admin
        workstation deployment with *RecipientManagement PSSnapin shortcut.
      • Exchange Management Tools only (-InstallManagementTools) — server OS.
      • Standalone post-install optimization (-StandaloneOptimize) — re-runs
        VDir URLs, relay connectors, log cleanup, RBAC report, HealthChecker
        on an already-installed server.
      • Standalone documentation (-StandaloneDocument) — generates only the
        Word installation document for an existing server/org.
      • MEAC AD Split-Permissions prep (-MEACPrepareADOnly) — Domain Admin
        runs this on a non-Exchange box to pre-create the MEAC automation
        account before Exchange install.

    Execution styles:
      • Autopilot (-Autopilot) — fully automated across all six install
        phases; reboots and resumes via RunOnce using DPAPI-encrypted
        credentials from the state file. Conditional reboots (Phase 2→3 and
        Phase 5→6) skip when nothing actually requires a reboot.
      • Copilot (interactive, default) — console menu selects mode and
        toggles; the operator is prompted only where needed.

    Install phases (0–6):
      0 Preflight checks + HTML preflight report
      1 Windows features, optional Windows Updates, export source-server config
      2 .NET 4.8/4.8.1, hotfixes, VC++ 2012/2013, URL Rewrite
      3 UCMA, PrepareAD / PrepareSchema, optional AD-replication wait
      4 Exchange setup (SetupAssist on failure), transport → Manual
      5 Security hardening, performance tuning, certificate import,
        Exchange SU, virtual-directory URLs
      6 Services, IIS, DAG join, connectors, HealthChecker, HTML+PDF report,
        Word installation document, cleanup

    State is persisted in `<computername>_EXpress_State.xml` under InstallPath
    (Export-Clixml); resumable across reboots. Credentials are DPAPI-
    encrypted (user+machine bound). Transcript is always written to disk;
    console output respects $State['LogVerbose'] / $State['LogDebug'].

    .LINK
    https://github.com/st03psn/EXpress

    .LINK
    http://eightwone.com

    .NOTES
    Requires Windows PowerShell 5.1. PowerShell 6/7 is not supported —
    Exchange PSSnapin requires Windows PowerShell 5.1 exclusively.
    Must run as Administrator on a domain-joined host (except Edge Transport).

    Supported environments:
      Exchange 2016 CU23         — Windows Server 2016
      Exchange 2019 CU10–CU14    — Windows Server 2019 / 2022
      Exchange 2019 CU15+        — Windows Server 2025
      Exchange SE RTM            — Windows Server 2022 / 2025

    Autopilot mode requires an account with elevated administrator
    privileges. When EXpress prepares AD, the account additionally needs
    Schema Admin + Enterprise Admin membership.

    .REVISIONS
    Full per-version history: see RELEASE-NOTES.md in the repository root.
    EXpress is a fork of Install-Exchange15.ps1 by Michel de Rooij (v1.0–v4.23).
    Michel's versioning ends at 4.23. EXpress starts its own numbering from 1.0.

    ── EXpress (st03psn, 2026—) — newest first ──────────────────────────────────

    1.3.0   License key activation: -LicenseKey param, ConfigFile LicenseKey key, Copilot
            prompt (5-min auto-skip); Set-ExchangeLicense activates Standard/Enterprise in
            Phase 6. Bugfixes: Add-ADPermission retry with backoff (P2); Default Frontend
            ProtocolLoggingLevel Verbose; RC4 registry path guards fixed; MAPI VDir skips
            InternalAuthenticationMethods on Exchange SE RTM; auth cert NotAfter null guard;
            EP HTML report uses live IIS values (no -ADPropertiesOnly); New-WdTable -ColWidths
            parameter; explicit column widths for RBAC, retention tags, Defender paths, DR
            scenarios Word tables.
    1.2.2   ConfigFile implies Autopilot (no explicit key needed); AnonymousRelay=$false now
            suppresses connector creation (gate was subnets-only); ConfigFile missing Namespace
            aborts with error instead of silent skip; LogRetentionDays defaults to 30 in
            ConfigFile mode (was hidden inside function); deploy-example relay section
            restructured with AnonymousRelay as self-documented master switch.
    1.2.1   Bugfixes: parser error in DNS template rows (if-in-array, PS 5.1); ExistingOrg
            probe re-run after Phase 3 (catch sets $false); Test-Feature explicit config
            blocked by Condition (precedence inverted); Enable-IanaTimeZoneMappings
            ignored explicit config; HTML report — certificates empty (local fallback),
            EP duplicate row removed, HSTS N/A when no cert, PowerShell VDir URL missing;
            URL Rewrite OK step now logged.
    1.2     Debug-mode overhaul: streaming Debug log under .\Debug\ (superset of install
            log), per-phase state+system snapshots, magenta startup banner + window-title
            indicator, halt after Phase 4 in Copilot+Debug for VM snapshot. ConfigFile
            runs are now fully unattended: implicit AutoApproveWindowsUpdates=$true,
            CertificatePassword config key skips PFX prompt. ExistingOrg auto-detection
            (ADSI + Initialize-Exchange) flips defaults to $false for 8 org-wide settings
            (MaxMessageSize150MB, MessageExpiration7d, SafetyNet2d, HtmlNDR, ModernAuth,
            OWASessionTimeout6h, DisableTelemetry, MapiHttp). Language switched from
            $German to 2-letter ISO code (-Language 'DE'/'EN', future IT/FR/...).
            AnonymousRelay=$true without subnets seeds RFC 5737 placeholders (matches
            Copilot blank-answer). Menu input validation: org name, retention days, PFX
            path, MDB name. Sprint 1: 101 empty catch blocks filled. Sprint 3:
            73-DocHelpers.ps1 extraction (sections 4 + 9 of New-InstallationDocument).
            CI: PSScriptAnalyzer + JUnit Pester artefact. IMAP4/POP3 namespace config.
            PSWindowsUpdate air-gap (4-tier: installed/SourcesPath/TEMP/PSGallery);
            Get-EXpressDownloads.ps1 pre-stages PSWindowsUpdate to %TEMP%\EXpress-sources.
            Bugfixes: DriverDate cast, ByteQuantifiedSize null guard in Word doc,
            Word-doc-failed log severity ERROR (was WARN), end-of-script non-fatal error
            summary, deploy-example cleanup (org-wide warning, credentials moved to top,
            duplicates removed).
    1.1.8   Word doc + HTML report enhancements: hardware type, time zone, uptime, NIC
            drivers, readable Exchange version, EP column in VDir table, VC++ table,
            TLS cipher suites, HSTS, Download Domains; phantom cert filter at source;
            Start-EXpress.cmd dynamic path; Processing log output demoted to Verbose.
    1.1.7   Bugfix: HTML report — phantom certs filtered (DateTime.MinValue), cert expiry
            uses TotalDays, Root CA display shows '(not set)' when absent, NetBIOS
            count null-safe; Register-ExecutedCommand for IANA timezone moved inside
            Enable-IanaTimeZoneMappings so log only shows actual execution.
    1.1.6   Bugfix: EnableDownloadDomains org flag now set (CVE-2021-1730 was incomplete);
            PowerShell VDir sets ExternalUrl only (InternalUrl stays http); NetBIOS report
            now checks registry for pending-reboot state; OWA EP integer normalization;
            certificate expiry used .Days (days-component) instead of TotalDays.
    1.1.5   Docs: menu screenshots + Word doc mockup nav fix.
    1.1.4   AutoApproveWindowsUpdates toggle (default off): Security/Critical no longer
            auto-approved in Autopilot without explicit opt-in.
    1.1.3   Windows Updates: [A]=all removed; each Security/Critical update confirmed individually.
    1.1.2   NuGet auto-install, RunOnce path fix (dot-source module resolution), Exchange
            source default path, module parse errors, PS 5.1 (if...) menu crashes.
    1.1     src/ renamed to modules/. Install-target matrix: Ex2019 CU10-CU14 rejected; Ex2016
            CU23 restricted to WS2016. F26: Access Namespace mail config. Menu back/edit step.
            tools/Get-EXpressDownloads.ps1. CI merge-guard workflow.
    1.0     EXpress rename + modularization: Install-Exchange15.ps1 renamed to EXpress.ps1;
            split into 21 modules/*.ps1; dist/EXpress.ps1 merged build. Centralized downloads
            to sources/. Install-target matrix tightened to latest CU per Exchange line.
    0.8     Advanced Configuration menu + templates (v5.95, v5.96): ~55 toggles across 6
            pages, Test-Feature condition gate, config-file parity. Installation-Document
            Template support with {{token}} replacement.
    0.7     Language reform + MEAC hybrid (v5.90-v5.94.1): default output English, -German
            switch. Plain-text credentials in config file. Hybrid-aware MEAC + AD Split
            Permissions. Word doc audit-readiness: 9 new sections (change mgmt, RBAC, ports,
            compliance mapping, GDPR, backup, monitoring, acceptance tests).
    0.6     Security hardening + MEAC (v5.86-v5.88.3): Defender realtime/Tamper Protection,
            LLMNR/mDNS disable, Disable-UnnecessaryServices. MEAC Auth-Cert auto-renewal task.
            Word doc enrichment: TLS semantics, IMAP/POP3, connector detail, DNS template,
            Admin Audit Log, Anti-Spam, Crimson channels. Bugfixes: Phase 5-6 spurious reboot,
            nested-array reshape, Auth Cert validity, $state/$State shadow, (if...) crashes.
    0.5     Org-wide documentation + conditional reboots (v5.84, v5.85): all Exchange servers
            documented via CIM/WSMan remote query. Phase 2-3 and 5-6 reboots skipped when
            nothing pending. Test-RebootPending helper. VC++ 2013 URL updated.
    0.4     Word Installation Document (v5.82, v5.83): pure-PowerShell OpenXML engine, 15
            chapters. Three-tier logging (INFO/VERBOSE/DEBUG), unified file naming, dev tools.
    0.3     Installation reports + post-config (v5.4-v5.6): HTML Installation Report, PDF
            export. Anti-spam, log cleanup, reconnect session, relay improvements, reports
            subfolder. Bugfix series (v0.3.1-v0.3.4): disable services, link fixes, edge
            guards, SU reboot timing, FormatException in HTML report.
    0.2     Hardening + connector framework (v5.2, v5.3): HSTS, EOMT, VDir URLs,
            Wait-ADReplication, relay connectors, RBAC report. Add-BackgroundJob,
            New-LDAPSearch, registry idempotency, BSTR zeroing.
    0.1     Foundation (Rounds 1-7, v5.0, v5.1): WMI-to-CIM migration, Write-ToTranscript,
            security baseline, $WS2025_PREFULL fix. Pre-flight HTML report, source-server
            config export/import, HealthChecker, DAG, PFX cert. Interactive menu, Autopilot,
            Windows Updates, Exchange SU, ConfigFile, Build.ps1.

    ── Predecessor: Install-Exchange15.ps1 by Michel de Rooij (v1.0–v4.23) ─────
    EXpress is built on Michel's work. His version history (oldest first):

    1.0     Initial community release
    1.01    Added logic to prepare AD when organization present
            Fixed checks and logic to prepare AD
            Added testing for domain mixed/native mode
            Added testing for forest functional level
    1.02    Fixed small typo in post-prepare AD function
    1.03    Replaced installing most OS features in favor of /InstallWindowsComponents
            Removed installation of Office Filtering Pack
    1.1     When used for AD preparation, RSAT-ADDS-Tools won't be uninstalled
            Pending reboot detection. In AutoPilot, script will reboot and restart phase.
            Installs Server-Media-Foundation feature (UCMA 4.0 requirement)
            Validates provided credentials for AutoPilot
            Check OS version as string (should accomodate non-US OS)
    1.5     Added support for WS2008R2 (i.e. added prereqs NET45, WMF3), IEESC toggling,
            KB974405, KB2619234, KB2758857 (supersedes KB2533623). Inserted phase for
            WS2008R2 to install hotfixes (+reboot); this phase is skipped for WS2012.
            Added InstallPath to AutoPilot set (or default won't be set).
    1.51    Rewrote Test-Credentials due to missing .NET 3.5 Out of the Box in WS2008R2.
            Testing for proper loading of servermanager module in WS2008R2.
    1.52    Fix .NET / PrepareAD order for WS2008R2, relocated RebootPending check
    1.53    Fix phase of Forest/Domain Level check
    1.54    Added Parameter InstallBoth to install CAS and Mailbox, workaround as PoSHv2
            can discriminate overlapping ParameterSets (resulting in AmbigiousParameterSet)
    1.55    Feature installation bug fix on WS2012
    1.56    Changed logic of final cleanup
    1.6     Code cleanup (merged KB/QFE/package functions)
            Fixed Verbose setting not being restored when script continues after reboot
            Renamed InstallBoth to InstallMultiRole
            Added 'Yes to All' option to extract function to prevent overwrite popup
            Added detection of setup file version
            Added switch IncludeFixes, which will install recommended hotfixes
            (2008R2:KB2803754,KB2862063 2012:KB2803755,KB2862064) and KB2880833 for CU2 & CU3.
    1.61    Fixed XML not found issue when specifying different InstallPath (Cory Wood)
    1.7     Added Exchange 2013 SP1 & WS2012R2 support
            Added installing .NET Framework 4.51 (2008 R2 & 2012 - 2012R2 has 4.51)
            Added DisableRetStructPinning for Mailbox roles
            Added KB2938053 (SP1 Transport Agent Fix)
            Added switch InstallFilterPack to install Office Filter Pack (OneNote & Publisher support)
            Fixed Exchange failed setup exit code anomaly
    1.71    Uncommented RunOnce line - AutoPilot should work again
            Using strings for OS version comparisons (should fix issue w/localized OS)
            Fixed issue installing .NET 4.51 on WS2012 ('all in one' kb2858728 contains/reports
            WS2008R2/kb958488 versus WS2012/kb2881468
            Fixed inconsistency with .NET detection in WS2012
    1.72    Added CU5 support
            Added KB2971467 (CU5 Disable Shared Cache Service Managed Availability probes)
    1.73    Added CU6 support
            Added KB2997355 (Exchange Online mailboxes cannot be managed by using EAC)
            Added .NET Framework 4.52
            Removed DisableRetStructPinning (not required for .NET 4.52 or later)
    1.8     Added CU7 support
    1.9     Added CU8 support
            Fixed CU6/CU7 detection
            Added (temporary) clearing of Execution Policy GPO value
            Added Forest Level check to throw warning when it can't read value
            Added KB2985459 for WS2012
            Using different service to detect installed version
            Installs WMF4/NET452 for supported Exchange versions
            Added UseWMF3 switch to use WMF3 on WS2008R2
    2.0     Renamed script to Install-Exchange15
            Added CU9 support
            Added Exchange Server 2016 Preview support
            Fixed registry checks for GPO error messages
            Added ClearSCP switch to clear Autodiscover SCP record post-setup
            Added Import-ExchangeModule() for post-configuration using EMS
            Bug fix .NET installation
            Modified AD checks to support multi-forest deployments
            Added access checks for Installation, MDB and Log locations
            Added checks for Exchange organization/Organization parameter
    2.03    Bug & typo fix
    2.1     Replaced ClearSCP with SCP param
            Added Lock switch to lock computer during installation
            Configures High Performance Power plan
            Added installing feature RSAT-Clustering-CmdInterface
            Added pagefile configuration when it's set to 'system managed'
    2.11    Added Exchange 2016 RTM support
            Removed Exchange 2016 Preview support
    2.12    Fixed pre-CU7 .NET installation logic
    2.2     Added (temporary) blocking unsupported .NET Framework 4.6.1 (KB3133990)
            Added recommended updates KB2884597 & KB2894875 for WS2012
            Changes to output so all output/verbose/warning/error get logged
            Added check to Organization for invalid characters
            Fixed specifying an Organization name containing spaces
    2.3     Added support up to Exchange 2013 CU12 / Exchange 2016 CU1
            Switched version detection to ExSetup, now follows Build
    2.31    Fixed output error messages
    2.4     Added support up to Exchange 2013 CU13 / Exchange 2016 CU2
            Added support for .NET 4.6.1 (Exchange 2013 CU13+ / Exchange 2016 CU2+)
            Added NONET461 switch, to use .NET 4.5.2, and block .NET 4.6.1
            Added installation of .NET 4.6.1 OS-dependent required hotfixes:
            * KB2919442 and KB2919355 (~700MB!) for WS2012R2 (prerequisites).
            * KB3146716 for WS2008/WS2008R2, KB3146714 for WS2012, and KB3146715 for WS2012R2.
            Added recommended Keep-Alive and RPC timeout settings
            Added DisableSSL3 to disable SSL3 (KB187498)
    2.41    Bug fix - Setup version of Exchange 2013 CU13 is .000, not .003
    2.42    Bug fix - Installation of KB2919442 only detectable after reboot; adjusted logic
            Added /f (forceAppsClose) for .MSU installations
    2.5     Added recommended hotfixes:
            * KB3146717 (=offline version of 3146718)
            * KB2985459 (WS2012)
            * KB3041832 (WS2012R2)
            * KB3004383 (WS2008R2)
            Added logging of AD Site
            Added computername to filename of state file and log
            Changed credential prompting, will use current account
            Changed Power Plan setting to use InstanceID instead of textual match
            Fixed KeepAlive timeout setting
            Added checks for running as Enterpise & Schema admin
            Fixed NoSetup bug (would abort)
            Added check to see if Exchange server object already exists
            Added Recover switch for RecoverServer mode
    2.51    Script will abort when ExSetup has non-0 exitcode
            Script will ignore package exit codes -2145124329 (SUS_E_NOT_APPLICABLE)
    2.52    Script will abort when AD site can not be determined
            Fixed SCP parameter handling, use '-' to remove the SCP
    2.53    Fixed NoSetup logic skipping NET 4.6.1 installation
            Added .NET framework optimization post-config (7318.DrainNGenQueue)
    2.54    Fixed failing TargetPath check
    2.6     Added support for Exchange 2013 CU14 and Exchange 2016 CU3
            Fixed 7318.DrainNGenQueue routine
            Some minor cosmetics
    2.7     Added support for Windows Server 2016 (Exchange Server 2016 CU3+ only)
    2.8     Added DisableRC4 to disable RC4 (kb2868725)
            Fixed DisableSSL3, removed disabling SSL3 as client
            Disables NIC Power Management during post config
    2.9     Added support for Exchange 2016 CU4
            Added support for Exchange 2013 CU15
            Added KB3206632 to Exchange 2016 @ WS2016 requirements
    2.91    Added support for Exchange 2016 CU5
            Added support for Exchange 2013 CU16
    2.92    Cosmetics and code cleanup when installing on WS2016
            Output cosmetics when disabling RC4
    2.93    Added blocking .NET Framework 4.7
    2.95    Added support for Exchange 2016 CU6
            Added support for Exchange 2013 CU17
    2.96    Added support for Exchange 2016 CU7
            Added support for Exchange 2013 CU18
            Added FFL 2008R2 checks for Exchange 2016 CU7
            Added blocking of .NET Framework 4.7.1
            Consolidated .NET Framework blocking routines
            Modified version comparison routine
    2.97    Added support for Exchange 2016 CU8
            Added support for Exchange 2013 CU19
            Added NONET471 switch
    2.98    Added support for Exchange 2016 CU9
            Added support for Exchange 2013 CU20
            Added blocking of .NET Framework 4.7.2 (Preview)
            Added upgrade mode detection
            Added TargetPath usage for Recover mode
    2.99    Added Windows Defender exclusions (Ex2016 on WS2016)
    2.991   Fixed .NET blockade removal
            Fixed upgrade detection
            Minor bugs and cosmetics fixes
    2.99.2  Fixed Recover Mode Phase
            Fixed InstallMDBDBPath location check
            Added support for Exchange 2016 CU10
            Added support for Exchange 2013 CU21
            Added Visual C++ 2013 Redistributable prereq (Ex2016CU10+/Ex2013CU21+)
            Fixed Exchange setup result detection
            Changed code to determine AD Configuration container
            Changed script to abort on non-static IP presence
            Removed InstallFilterPack switch (obsolete)
            Code cleanup and cosmetics
    2.99.3  Fixed TargetPath-Recover parameter mutual exclusion
    2.99.4  Fixed Recover mode not adding /InstallWindowsComponents
            Added SkipRolesCheck switch
            Added Exchange 2019 Public Preview support on Windows Server 2016
    2.99.5  Added setting desktop background during setup
            Some code cleanup
    2.99.6  Added Exchange 2019 Preview on Windows Server 2019 support (desktop & core)
    2.99.7  Updated location where hotfix are being published
    2.99.8  Updated to Support Edge (Simon Poirier)
    2.99.81 Fixed phase sequencing with reboot pending
    2.99.82 Added reapplying KB2565063 (MS11-025) to IncludeFixes
            Changed downloading VC++ Package to filename indicating version
            Added post-setup Healthcheck / IIS Warmup
    2.99.9  Added support for Exchange 2016 CU11
            Updated SourcePath parameter usage (ISO)
            Added .NET Framework 4.7.2 support
            Added Windows Defender presence check
    3.0.0   Added Exchange 2019 support
            Rewritten VC++ detection
    3.0.1   Integrated Exchange 2019 RTM Cipher correction
    3.0.2   Replaced filename constructs with Join-Path
            Fixed typo in installing KB4054530
    3.0.3   Fixed typos in Join-Path constructs
    3.0.4   Fixed bug in Install-MyPackage
    3.1.0   Added support for Exchange 2019 CU1
            Added support for Exchange 2016 CU12
            Added support for Exchange 2013 CU22
            Fixed Hotfix KB3041832 url
            Fixed NoSetup Mode/EmptyRoles problem
            Added skip Health Monitor checks for InstallEdge
            Fixed potential Exchange version misreporting
    3.1.1   Fixed detection of Defender
    3.2.0   Added support for Exchange 2019 CU2
            Added support for Exchange 2016 CU13
            Added support for Exchange 2013 CU23
            Added support for NET Framework 4.8
            Added NoNET48 switch
            Added disabling of Server Manager during installation
            Removed support for Windows Server 2008R2
            Removed support for Windows Server 2012
            Removed Switch UseWMF3
    3.2.1   Updated Pagefile config for Exchange 2019 (25% mem.size)
    3.2.2   Added support for Exchange 2019 CU3
            Added support for Exchange 2016 CU14
    3.2.3   Fixed typo for Ex2019CU3 detection
    3.2.4   Added support for Exchange 2019 CU4+CU5
            Added support for Exchange 2016 CU15+CU16
    3.2.5   Fixed typo in enumeration of Exchange build to report
            Fixed specified vs used MDBLogPath (would add unspecified <DBNAME>\Log)
    3.2.6   Added support for Exchange 2019 CU6
            Added support for Exchange 2016 CU17
            Added VC++ Runtime 2012 for Exchange 2019
    3.3     Added support for Exchange 2019 CU7
            Added support for Exchange 2016 CU18
    3.4     Added support for Exchange 2019 CU8
            Added support for Exchange 2016 CU19
            Script allows non-static IP config with service Windows Azure Guest Agent, Network Agent or Telemetry Service present
    3.5     Added support for Exchange 2019 CU8
            Added support for Exchange 2016 CU19
            Added support for KB5003435 for 2019CU8+9,2016CU19+20 and 2013CU23
            Added support for KB5000871 for 2019RTM-CU7, 2016CU8-CU18 and 2013CU21+22
            Added support for Interim Update installation & detection
            Updated .NET 4.8 download link
            Updated Visual C++ 2012 download link (latest release)
            Updated Visual C++ 2013 download link (latest release)
            Corrected not installing KB3206632 on WS2019
            Corrected disabling of Server Manager during setup
            Fixed setting High Performance Plan for recent Windows builds
            Textual corrections
    3.6     Added support for Exchange 2019 CU11
            Added support for Exchange 2016 CU22
            Added support for Exchange 2019 CU10
            Added support for Exchange 2019 CU9
            Added support for Exchange 2016 CU21
            Added support for Exchange 2016 CU20
            Added IIS URL Rewrite prereq for Ex2019CU11+ & Ex2016 CU22+
            Added support for KB2999226 on for WS2012R2
            Added DiagnosticData switch to set initial DataCollectionEnabled mode
    3.61    Added mention of Exchange 2019
    3.62    Added support for Exchange 2019 CU12
            Added support for Exchange 2016 CU23
    3.7     Added support for Windows Server 2022
            Fixed logic for installing the IIS Rewrite module for Ex2016CU22+/Ex2019CU11+
            Fixed logic when to use the new /IAcceptExchangeServerLicenseTerms_DiagnosticData* switch
    3.71    Updated recommended Defender AV inclusions/exclusions
    3.8     Added support for Exchange 2019 CU13
    3.9     Added support for Exchange 2019 CU14
            Added support for .NET Framework 4.8.1
            Added NONET481 switch to use .NET 4.8 instead of 4.8.1 for Exchange 2019 CU14+
            Added DoNotEnableEP and DoNotEnableEP_FEEWS switches for Exchange 2019 CU14+
            Added deploying AUG2023 SUs for Ex2019CU13/Ex2019CU12/Ex2016CU23 when IncludeFixes specified
            Changed example to show usage of iso as source
            Added descriptive message when specifying invalid SourcePath
            Fixed detection source path when iso already mounted without drive letter assignment
    4.0     Added support for Exchange 2019 CU15
            Added support for Windows Server 2025 (Exchange 2019 CU15+)
            Removed Exchange 2013 support
            Removed Exchange 2016 CU1-22 support
            Removed Exchange 2019 RTM-CU9
            Removed Windows Server 2012 R2 support
            Added removal of obsolete MSMQ feature when installed
            Added EnableECC switch to configure Elliptic Curve Crypto support
            Added NoCBC switch to prevent configuring AES256-CBC-encrypted content support
            Added EnableAMSI switch to configure AMSI body scanning for ECP, EWS, OWA and PowerShell
            Added EnableTLS12 switch to configure TLS12
            Added EnableTLS13 switch to configure TLS13 on WS2022/WS2025 with EX2019CU15+
            Removed InstallMailbox, InstallCAS, InstallMultiRole switches
            Removed NoNet461, NoNet471, NoNet472 and NoNet48 switches
            Removed UseWMF3 switch
            Added Ex2013 detection as it cannot coexist with Ex2019CU15+
            Enabled loading Exchange module in postconf needed for possible override cmdlets
            Removed setup phase shown on wallpaper
            Set minimal required PS version to 5.1
            Code cleanup
            Functions now use approved verbs
    4.01    Removed obsolete TLS13 setup detection
    4.10    Added support for Exchange Server SE
    4.11    Fixed feature installation for WS2022/WS2025 Core
    4.12    Fixed feature installation (Web-W-Auth, should be Web-Windows-Auth)
            Using ADSI for Ex2013 detection
    4.13    Fixed race issue when installing from ISO and restarting installation
            Tested with SW_DVD9_Exchange_Server_Subscription_64bit_MultiLang_Std_Ent_.iso_MLF_X24-08113.iso
    4.20    Clearing/setting SCP now background job during install to configure it asynchronous & ASAP
    4.21    Added disabling MSExchangeAutodiscoverAppPool during setup to prevent responding to requests during setup and postconfig
    4.22    Corrected download VC++2013 runtime URL due to shortcut being unavailabe
    4.23    Fixed Edge installation (no need checking for Ex2013 in AD)

    .PARAMETER Organization
    Exchange organization name. When omitted, the PrepareAD step is skipped.

    .PARAMETER InstallEdge
    Install the Edge Transport role (Exchange 2016/2019/SE). Implies parameter set "E".

    .PARAMETER EdgeDNSSuffix
    DNS suffix to apply on the Edge host (parameter set "E" only).

    .PARAMETER Recover
    Run Exchange setup in RecoverServer mode.

    .PARAMETER MDBName
    Name of the initially created mailbox database.

    .PARAMETER MDBDBPath
    Database path for the initially created database. Requires -MDBName.

    .PARAMETER MDBLogPath
    Log path for the initially created database. Requires -MDBName.

    .PARAMETER InstallPath
    Working directory for prereq downloads, state file, transcript, and reports.
    Default: the directory containing EXpress.ps1. May be a UNC path to share prereq cache between hosts.

    .PARAMETER SourcePath
    Location of the Exchange installation files (setup.exe) or the Exchange ISO.
    ISO files are mounted on demand and dismounted when setup completes.

    .PARAMETER TargetPath
    Target directory for the Exchange binaries (Exchange ProgramFiles path).

    .PARAMETER NoSetup
    Prepare + prereqs only; do not run Exchange setup. -SourcePath is still
    required to determine Exchange version and applicable prerequisites.

    .PARAMETER Autopilot
    Fully automated installation across all phases. Reboots and resumes via
    RunOnce using DPAPI-encrypted credentials from the state file. Conditional
    Phase 2→3 and Phase 5→6 reboots skip when nothing pending requires a reboot.

    .PARAMETER Credentials
    Install-admin credentials used for Autopilot auto-logon and any elevated
    sub-process that needs explicit credentials. Format: DOMAIN\User or
    user@domain. Prompted when not supplied.

    .PARAMETER IncludeFixes
    Download and install additional recommended Exchange hotfixes for the
    detected OS/Exchange combination.

    .PARAMETER NoNet481
    Use .NET Framework 4.8 instead of 4.8.1 for Exchange 2019 CU14+ on WS2016/
    WS2019. WS2022/WS2025 always use 4.8.1 regardless of this switch.

    .PARAMETER DoNotEnableEP
    Do not enable Extended Protection on Exchange 2019 CU14+.

    .PARAMETER DoNotEnableEP_FEEWS
    Do not enable Extended Protection on the Front-End EWS virtual directory on
    Exchange 2019 CU14+.

    .PARAMETER DisableSSL3
    Disable SSL 3.0 (Schannel) after setup.

    .PARAMETER DisableRC4
    Disable RC4 ciphers (Schannel) after setup.

    .PARAMETER EnableECC
    Configure Elliptic Curve Cryptography support after setup.

    .PARAMETER NoCBC
    Do not configure AES-256-CBC-encrypted content support after setup.

    .PARAMETER EnableAMSI
    Enable AMSI body scanning for ECP / EWS / OWA / PowerShell virtual directories.

    .PARAMETER EnableTLS12
    Explicitly enable TLS 1.2 (Schannel) post-setup.

    .PARAMETER EnableTLS13
    Explicitly enable TLS 1.3 on WS2022/WS2025 for Exchange 2019 CU15+ / SE
    (default: enabled when supported).

    .PARAMETER SCP
    Reconfigure the Autodiscover Service Connection Point post-setup
    (e.g. https://autodiscover.contoso.com/autodiscover/autodiscover.xml).
    Pass '-' to remove the SCP record.

    .PARAMETER DiagnosticData
    Sets the initial Data Collection mode when deploying Exchange 2019 CU11+ or
    Exchange 2016 via the /IAcceptExchangeServerLicenseTerms_DiagnosticData* switch.

    .PARAMETER Lock
    Lock the console while the script runs.

    .PARAMETER SkipRolesCheck
    Skip the Schema Admin + Enterprise Admin membership check.

    .PARAMETER Phase
    Internal use — Autopilot uses this to resume the correct phase after reboot.

    .PARAMETER PreflightOnly
    Run preflight checks and generate the HTML preflight report only. No
    Exchange install, no system changes beyond the report.

    .PARAMETER CopyServerConfig
    Source Exchange server whose configuration (virtual directories, transport
    config, receive connectors) should be exported via Remote PowerShell and
    re-applied post-setup. Useful for swing migrations.

    .PARAMETER CertificatePath
    PFX certificate to import and enable for IIS + SMTP post-setup. The PFX
    password is prompted interactively and stored DPAPI-encrypted in state
    for the duration of the run.

    .PARAMETER DAGName
    Name of an existing Database Availability Group the new server should join
    post-setup.

    .PARAMETER SkipHealthCheck
    Skip the automatic download + execution of CSS-Exchange HealthChecker at
    the end of the installation.

    .PARAMETER NoCheckpoint
    Skip System Restore checkpoints before each phase. No effect on Windows
    Server where Checkpoint-Computer is unavailable.

    .PARAMETER InstallRecipientManagement
    Recipient Management Tools installation mode (3-phase flow). Installs
    setup.exe /roles:ManagementTools on an admin workstation, runs
    Add-PermissionForEMT.ps1, and creates a desktop shortcut loading the
    *RecipientManagement PSSnapin.

    .PARAMETER RecipientMgmtCleanup
    In Recipient Management mode, clean up legacy AD permissions after a
    successful upgrade install.

    .PARAMETER InstallManagementTools
    Exchange Management Tools installation mode (Server OS). Installs only the
    prerequisites and setup.exe /roles:ManagementTools.

    .PARAMETER ConfigFile
    Path to a PowerShell data file (.psd1) whose hashtable carries any of the
    script parameters. Makes long command lines manageable for repeat
    deployments and enables fully unattended runs.

    Fully unattended install-admin credentials may live in the config file via
    the keys AdminUser / AdminPassword (plain text). EXpress converts them to
    a PSCredential on load; a loud SECURITY WARNING is logged every run.
    Delete or scrub the config file immediately after install completes —
    the state file's DPAPI encryption is machine-bound, but the config file
    is not.

    .PARAMETER InstallWindowsUpdates
    During Phase 1 / post-setup, check for pending Windows Updates and
    applicable Exchange Security Updates (SUs), download and install them.
    Reboots integrate into the Autopilot phase flow.

    .PARAMETER SkipWindowsUpdates
    Explicitly skip the Windows Update / Exchange SU check even when the menu
    or config file would otherwise enable it.

    .PARAMETER SkipSetupAssist
    Skip the automatic download + execution of CSS-Exchange SetupAssist.ps1
    when Exchange setup fails in Phase 4.

    .PARAMETER Namespace
    External namespace (e.g. outlook.contoso.com) for configuring all Exchange
    virtual-directory internal and external URLs in Phase 6. When omitted,
    virtual-directory URLs are left at their defaults.

    .PARAMETER DownloadDomain
    Separate FQDN for OWA attachment downloads (e.g. download.contoso.com) —
    sets InternalDownloadHostName / ExternalDownloadHostName on the OWA vdir
    to mitigate CVE-2021-1730 (attachment cookie theft). Must differ from
    -Namespace. Requires DNS + certificate coverage for the host. Requires
    -Namespace.

    .PARAMETER RunEOMT
    Run the CSS-Exchange Emergency Mitigation Tool (EOMT) in Phase 5. Use
    when deploying a server that may have been exposed to publicly known
    vulnerabilities before patching.

    .PARAMETER WaitForADSync
    After PrepareAD (Phase 3), poll repadmin /showrepl /errorsonly until all
    AD replication errors clear or a 6-minute timeout elapses. Useful in
    multi-site AD deployments.

    .PARAMETER LogRetentionDays
    Register a Windows Scheduled Task (Exchange Log Cleanup, daily 02:00)
    that removes IIS log files and Exchange transport / tracking logs older
    than N days (1–365). 0 = log cleanup disabled. Task lives in the \Exchange\ task folder.

    .PARAMETER RelaySubnets
    IP ranges (e.g. '192.168.1.0/24','10.0.0.5') allowed to relay anonymously
    to accepted domains only (internal relay). Creates the
    "Anonymous Internal Relay" receive connector without
    Ms-Exch-SMTP-Accept-Any-Recipient. On success, AnonymousUsers is removed
    from the Default Frontend receive connector.

    .PARAMETER ExternalRelaySubnets
    IP ranges allowed to relay anonymously to any recipient including external
    addresses. Creates "Anonymous External Relay" and grants
    Ms-Exch-SMTP-Accept-Any-Recipient to the ANONYMOUS LOGON account
    (SID S-1-5-7 — language-independent). Use with extreme care — restrict
    to trusted senders (scanner/printer IPs).

    .PARAMETER LicenseKey
    Exchange Server product key (format: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX).
    Activates Standard or Enterprise edition at the end of Phase 6, before
    the installation report is generated. Omit to run as Trial (180-day
    evaluation). In Copilot mode the key can also be entered interactively
    with a 5-minute auto-skip prompt. Can also be set via ConfigFile:
    LicenseKey = 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX'.

    .PARAMETER StandaloneOptimize
    Standalone mode: runs post-install optimizations (VDir URLs, anonymous
    relay connectors, certificate import, DAG join, log cleanup task, RBAC
    report, HealthChecker, DB-path checks, MEAC task) on an already-installed
    Exchange server without the full install flow. Combine with -Namespace,
    -CertificatePath, -DAGName, -RelaySubnets, -LogRetentionDays,
    -SkipHealthCheck as needed.

    .PARAMETER SkipInstallReport
    Suppress the HTML installation report at Phase 6 completion. By default
    EXpress writes a comprehensive HTML report (and a PDF when Microsoft Edge
    is available) to InstallPath for customer handover and audit purposes.

    .PARAMETER NoWordDoc
    Suppress the Word installation document (.docx) at Phase 6 completion.

    .PARAMETER StandaloneDocument
    Standalone mode: generate only the Word installation document for an
    existing Exchange server / organisation. Loads state from -InstallPath;
    requires an active Exchange Management Shell or an installed Exchange
    server. Combine with -German, -CustomerDocument, -DocumentScope,
    -IncludeServers as needed.

    .PARAMETER CustomerDocument
    Mask passwords and internal IP addresses in the generated Word document
    for secure customer handover.

    .PARAMETER German
    Generate the Word installation document in German. English is the default
    for all script output. Legacy config-file key 'Language=DE' still maps
    to -German for back-compat; the old -Language command-line parameter has
    been removed.

    .PARAMETER DocumentScope
    Controls the data scope for the Word document.
      All   (default) — org-wide settings + all Exchange servers + local details
      Org            — org-wide settings only (no per-server sections)
      Local          — per-server sections only, no org-wide chapter

    .PARAMETER IncludeServers
    Restrict per-server documentation to the specified server names (string
    array). Applies when -DocumentScope is All or Local. The local server is
    always included. Example: -IncludeServers EX01,EX02

    .PARAMETER MEACAutomationCredential
    AD Split-Permissions passthrough. When a Domain Admin pre-created the MEAC
    automation account via -MEACPrepareADOnly, pass the resulting credential
    here; EXpress forwards it to MEAC as -AutomationAccountCredential. Stored
    DPAPI-encrypted in state (user+machine bound) so it survives the Autopilot
    reboot chain. Not needed in standard (non-Split) deployments — MEAC
    self-provisions SystemMailbox{b963af59-...}.

    .PARAMETER MEACIgnoreHybridConfig
    Hybrid passthrough to MEAC. Default behaviour: when a Hybrid Configuration
    is detected, MEAC registers the task in hybrid-safe mode — daily checks
    still run, but renewals are blocked (renewing without an HCW rerun would
    break Exchange Online federation). Set this switch to authorise MEAC to
    renew anyway; YOU must rerun the Hybrid Configuration Wizard afterwards.
    Pair with -MEACNotificationEmail to be alerted before the Auth Cert
    expires.

    .PARAMETER MEACIgnoreUnreachableServers
    Passthrough to MEAC. Permit task execution when some Exchange servers in
    the org are offline; MEAC validates only reachable nodes. Useful in
    multi-server orgs during planned maintenance.

    .PARAMETER MEACNotificationEmail
    Passthrough to MEAC -SendEmailNotificationTo. Address that receives MEAC
    notifications when a renewal occurs (or is blocked by hybrid detection).
    Format: valid user@domain.tld.

    .PARAMETER MEACPrepareADOnly
    Standalone mode for AD Split-Permissions environments. Run on a
    non-Exchange box as a Domain Admin with user-create permissions. EXpress
    downloads MonitorExchangeAuthCertificate.ps1 and invokes it with
    -PrepareADForAutomationOnly -ADAccountDomain <domain>, then exits. Does
    NOT touch Exchange, does NOT run prereqs. Hand the resulting credential
    to the Exchange admin, who passes it to the regular install via
    -MEACAutomationCredential. Requires -MEACADAccountDomain.

    .PARAMETER MEACADAccountDomain
    AD domain (e.g. contoso.local) under which to create the MEAC automation
    account during -MEACPrepareADOnly. Mandatory in that mode.

    .EXAMPLE
    # Start interactively — opens the installation menu (mode, toggles, inputs)
    .\EXpress.ps1

    .EXAMPLE
    # Load all parameters from a config file (skips the interactive menu)
    .\EXpress.ps1 -ConfigFile .\deploy-mbx01.psd1

    .EXAMPLE
    # Fully unattended mailbox install with Autopilot (automatic reboots through all phases)
    .\EXpress.ps1 -SourcePath D:\Exchange -Organization Contoso -Autopilot

    .EXAMPLE
    # Full install with custom DB paths, Autodiscover SCP, and certificate
    $Cred = Get-Credential
    .\EXpress.ps1 -SourcePath C:\Install\ExchangeServerSE-x64.iso -Organization Contoso `
        -MDBName MDB01 -MDBDBPath D:\MailboxData\MDB01\DB -MDBLogPath D:\MailboxData\MDB01\Log `
        -SCP https://autodiscover.contoso.com/autodiscover/autodiscover.xml `
        -CertificatePath C:\Certs\mail.pfx -Autopilot -Credentials $Cred

    .EXAMPLE
    # Swing migration: copy config from source server, import PFX, join DAG
    .\EXpress.ps1 -SourcePath D:\Exchange -Autopilot `
        -CopyServerConfig EX01 -CertificatePath D:\Certs\mail.pfx -DAGName DAG01

    .EXAMPLE
    # Generate pre-flight HTML report only (no installation)
    .\EXpress.ps1 -SourcePath D:\Exchange -PreflightOnly

    .EXAMPLE
    # Install prerequisites only, skip Exchange setup
    .\EXpress.ps1 -NoSetup -SourcePath D:\Exchange

    .EXAMPLE
    # Recover a failed server
    .\EXpress.ps1 -Recover -SourcePath D:\Exchange -Autopilot

    .EXAMPLE
    # Edge Transport role
    .\EXpress.ps1 -InstallEdge -SourcePath D:\Exchange -Autopilot

    .EXAMPLE
    # Install Recipient Management Tools on an admin workstation
    .\EXpress.ps1 -InstallRecipientManagement -SourcePath D:\Exchange -Autopilot

    .EXAMPLE
    # Install Exchange Management Tools only (Server OS)
    .\EXpress.ps1 -InstallManagementTools -SourcePath D:\Exchange

    .EXAMPLE
    # Run all post-install optimizations on an existing Exchange server (no setup required)
    .\EXpress.ps1 -StandaloneOptimize -Namespace mail.contoso.com `
        -CertificatePath C:\Certs\mail.pfx -LogRetentionDays 30 `
        -RelaySubnets '10.0.1.0/24' -ExternalRelaySubnets '10.0.2.5'

    .EXAMPLE
    # Generate the default English Word document for the full organisation (org + all servers) — ad-hoc on any Exchange server
    .\EXpress.ps1 -StandaloneDocument

    .EXAMPLE
    # Generate a German Word document (same scope) using the -German shorthand
    .\EXpress.ps1 -StandaloneDocument -German

    .EXAMPLE
    # Generate an English Word document masked for customer handover, scoped to two specific servers
    .\EXpress.ps1 -StandaloneDocument -CustomerDocument `
        -DocumentScope All -IncludeServers EX01,EX02

    .EXAMPLE
    # Document only the org-wide configuration (no per-server hardware queries), German
    .\EXpress.ps1 -StandaloneDocument -DocumentScope Org -German

    .EXAMPLE
    # Full mailbox install; suppress Word doc
    .\EXpress.ps1 -SourcePath D:\Exchange -Organization Contoso -Autopilot -NoWordDoc

#>
[cmdletbinding(DefaultParameterSetName = 'Autopilot')]
param(
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [ValidatePattern('^$|^[a-zA-Z0-9\-][a-zA-Z0-9\-\ ]{1,62}[a-zA-Z0-9\-]$')]
    [string]$Organization,
    [parameter( Mandatory = $true, ParameterSetName = 'E')]
    [switch]$InstallEdge,
    [parameter( Mandatory = $true, ParameterSetName = 'E')]
    [string]$EdgeDNSSuffix,
    [parameter( Mandatory = $true, ParameterSetName = 'Recover')]
    [switch]$Recover,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [string]$MDBName,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [string]$MDBDBPath,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [string]$MDBLogPath,
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Autopilot')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [parameter( Mandatory = $false, ParameterSetName = 'R')]
    [parameter( Mandatory = $false, ParameterSetName = 'T')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    [parameter( Mandatory = $false, ParameterSetName = 'W')]
    [string]$InstallPath = 'C:\Install',
    [parameter( Mandatory = $true, ParameterSetName = 'E')]
    [parameter( Mandatory = $true, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [parameter( Mandatory = $true, ParameterSetName = 'R')]
    [parameter( Mandatory = $true, ParameterSetName = 'T')]
    [string]$SourcePath,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [string]$TargetPath,
    [parameter( Mandatory = $true, ParameterSetName = 'NoSetup')]
    [switch]$NoSetup = $false,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [parameter( Mandatory = $false, ParameterSetName = 'R')]
    [parameter( Mandatory = $false, ParameterSetName = 'T')]
    [switch]$Autopilot,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [parameter( Mandatory = $false, ParameterSetName = 'R')]
    [parameter( Mandatory = $false, ParameterSetName = 'T')]
    [System.Management.Automation.PsCredential]$Credentials,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$IncludeFixes,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$NoNet481,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$DoNotEnableEP,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$DoNotEnableEP_FEEWS,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$DisableSSL3,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$DisableRC4,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$EnableECC,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$NoCBC,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$EnableAMSI,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$EnableTLS12,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$EnableTLS13,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [ValidateScript({ ($_ -eq '') -or ($_ -eq '-') -or (([System.URI]$_).AbsoluteUri -ne $null) })]
    [String]$SCP = '',
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$DiagnosticData,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$Lock,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$SkipRolesCheck,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Autopilot')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [ValidateRange(0, 6)]
    [int]$Phase,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$PreflightOnly,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [string]$CopyServerConfig,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    [ValidateScript({ if (-not $_ -or (Test-Path $_ -PathType Leaf)) { $true } else { throw ('PFX file not found: {0}' -f $_) } })]
    [string]$CertificatePath,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    [string]$DAGName,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    [Switch]$SkipHealthCheck,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$NoCheckpoint,
    [parameter( Mandatory = $true, ParameterSetName = 'R')]
    [switch]$InstallRecipientManagement,
    [parameter( Mandatory = $false, ParameterSetName = 'R')]
    [switch]$RecipientMgmtCleanup,
    [parameter( Mandatory = $true, ParameterSetName = 'T')]
    [switch]$InstallManagementTools,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [parameter( Mandatory = $false, ParameterSetName = 'R')]
    [parameter( Mandatory = $false, ParameterSetName = 'T')]
    [parameter( Mandatory = $false, ParameterSetName = 'Autopilot')]
    [ValidateScript({ if (-not $_ -or (Test-Path $_ -PathType Leaf)) { $true } else { throw ('ConfigFile not found: {0}' -f $_) } })]
    [string]$ConfigFile,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [parameter( Mandatory = $false, ParameterSetName = 'R')]
    [parameter( Mandatory = $false, ParameterSetName = 'T')]
    [Switch]$InstallWindowsUpdates,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [parameter( Mandatory = $false, ParameterSetName = 'R')]
    [parameter( Mandatory = $false, ParameterSetName = 'T')]
    [Switch]$SkipWindowsUpdates,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$SkipSetupAssist,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    [string]$Namespace,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    # Root mail domain for Accepted Domain + Email Address Policy (e.g. contoso.com).
    # Defaults to the parent of -Namespace when omitted (e.g. mail.contoso.com → contoso.com).
    [string]$MailDomain,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    [string]$DownloadDomain,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$RunEOMT,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [Switch]$WaitForADSync,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    [ValidateRange(0, 365)]
    [int]$LogRetentionDays,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    [string[]]$RelaySubnets,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    [string[]]$ExternalRelaySubnets,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    [ValidatePattern('^$|^[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}$')]
    [string]$LicenseKey,
    [parameter( Mandatory = $true, ParameterSetName = 'O')]
    [switch]$StandaloneOptimize,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$SkipInstallReport,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    [parameter( Mandatory = $false, ParameterSetName = 'W')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [Switch]$NoWordDoc,
    [parameter( Mandatory = $true, ParameterSetName = 'W')]
    [switch]$StandaloneDocument,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    [parameter( Mandatory = $false, ParameterSetName = 'W')]
    [switch]$CustomerDocument,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    [parameter( Mandatory = $false, ParameterSetName = 'W')]
    [switch]$German,  # Deprecated, kept for backwards compat. Use -Language 'DE' instead.
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'E')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    [parameter( Mandatory = $false, ParameterSetName = 'W')]
    [ValidatePattern('^[A-Za-z]{2}$')]
    [string]$Language = 'EN',
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    [parameter( Mandatory = $false, ParameterSetName = 'W')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [ValidateSet('All', 'Org', 'Local')]
    [string]$DocumentScope = 'All',
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    [parameter( Mandatory = $false, ParameterSetName = 'W')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [string[]]$IncludeServers = @(),
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'O')]
    [parameter( Mandatory = $false, ParameterSetName = 'W')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [string]$TemplatePath = '',

    # --- MEAC passthroughs (v5.93) — applied when Register-AuthCertificateRenewal runs
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [System.Management.Automation.PsCredential]$MEACAutomationCredential,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [switch]$MEACIgnoreHybridConfig,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [switch]$MEACIgnoreUnreachableServers,
    [parameter( Mandatory = $false, ParameterSetName = 'M')]
    [parameter( Mandatory = $false, ParameterSetName = 'NoSetup')]
    [parameter( Mandatory = $false, ParameterSetName = 'Recover')]
    [AllowEmptyString()]
    [ValidatePattern('^$|^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$MEACNotificationEmail = '',

    # --- MEAC Split-Permissions prep (v5.93) — standalone mode, AD administrator
    # runs this on a non-Exchange box BEFORE Exchange is installed. Downloads MEAC,
    # invokes it with -PrepareADForAutomationOnly + -ADAccountDomain, then exits.
    # No Exchange prereqs, no setup, no reboot loop.
    [parameter( Mandatory = $true,  ParameterSetName = 'MEACPrepareAD')]
    [switch]$MEACPrepareADOnly,
    [parameter( Mandatory = $true,  ParameterSetName = 'MEACPrepareAD')]
    [ValidatePattern('^[A-Za-z0-9][A-Za-z0-9.\-]*$')]
    [string]$MEACADAccountDomain
)

process {
    # Capture the true entry-point path before dot-sourcing modules.
    # $MyInvocation.MyCommand.Path inside a dot-sourced file resolves to
    # that module file's path, not to EXpress.ps1 — modules read this
    # variable to build the Autopilot RunOnce command correctly.
    $EXpressEntryScript = $MyInvocation.MyCommand.Path
#region SOURCE-LOADER
    foreach ($m in (Get-ChildItem (Join-Path $PSScriptRoot 'modules') -Filter '*.ps1' | Sort-Object Name)) { . $m.FullName }
#endregion SOURCE-LOADER
} #Process

