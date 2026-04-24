<#
    .SYNOPSIS
    EXpress — unattended Exchange Server 2016/2019/SE installation, hardening,
    post-configuration, documentation, and day-2 standalone modes.

    Script file: EXpress.ps1
    Version:     1.1.7
    Maintainer:  st03ps

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

    EXpress development history (pre-1.0):

    0.1     Foundation (Rounds 1-7, v5.0, v5.1): WMI-to-CIM migration, Write-ToTranscript,
            security baseline, $WS2025_PREFULL fix. Pre-flight HTML report, source-server
            config export/import, HealthChecker, DAG, PFX cert. Interactive menu, Autopilot,
            Windows Updates, Exchange SU, ConfigFile, Build.ps1.
    0.2     Hardening + connector framework (v5.2, v5.3): HSTS, EOMT, VDir URLs,
            Wait-ADReplication, relay connectors, RBAC report. Add-BackgroundJob,
            New-LDAPSearch, registry idempotency, BSTR zeroing.
    0.3     Installation reports + post-config (v5.4-v5.6): HTML Installation Report, PDF
            export. Anti-spam, log cleanup, reconnect session, relay improvements, reports
            subfolder. Bugfix series (v0.3.1-v0.3.4): disable services, link fixes, edge
            guards, SU reboot timing, FormatException in HTML report.
    0.4     Word Installation Document (v5.82, v5.83): pure-PowerShell OpenXML engine, 15
            chapters. Three-tier logging (INFO/VERBOSE/DEBUG), unified file naming, dev tools.
    0.5     Org-wide documentation + conditional reboots (v5.84, v5.85): all Exchange servers
            documented via CIM/WSMan remote query. Phase 2-3 and 5-6 reboots skipped when
            nothing pending. Test-RebootPending helper. VC++ 2013 URL updated.
    0.6     Security hardening + MEAC (v5.86-v5.88.3): Defender realtime/Tamper Protection,
            LLMNR/mDNS disable, Disable-UnnecessaryServices. MEAC Auth-Cert auto-renewal task.
            Word doc enrichment: TLS semantics, IMAP/POP3, connector detail, DNS template,
            Admin Audit Log, Anti-Spam, Crimson channels. Bugfixes: Phase 5-6 spurious reboot,
            nested-array reshape, Auth Cert validity, $state/$State shadow, (if...) crashes.
    0.7     Language reform + MEAC hybrid (v5.90-v5.94.1): default output English, -German
            switch. Plain-text credentials in config file. Hybrid-aware MEAC + AD Split
            Permissions. Word doc audit-readiness: 9 new sections (change mgmt, RBAC, ports,
            compliance mapping, GDPR, backup, monitoring, acceptance tests).
    0.8     Advanced Configuration menu + templates (v5.95, v5.96): ~55 toggles across 6
            pages, Test-Feature condition gate, config-file parity. Installation-Document
            Template support with {{token}} replacement.

    1.0     EXpress rename + modularization: Install-Exchange15.ps1 renamed to EXpress.ps1;
            split into 21 modules/*.ps1; dist/EXpress.ps1 merged build. Centralized downloads
            to sources/. Install-target matrix tightened to latest CU per Exchange line.
    1.1     src/ renamed to modules/. Install-target matrix: Ex2019 CU10-CU14 rejected; Ex2016
            CU23 restricted to WS2016. F26: Access Namespace mail config. Menu back/edit step.
            tools/Get-EXpressDownloads.ps1. CI merge-guard workflow.
    1.1.2   NuGet auto-install, RunOnce path fix (dot-source module resolution), Exchange
            source default path, module parse errors, PS 5.1 (if...) menu crashes.
    1.1.3   Windows Updates: [A]=all removed; each Security/Critical update confirmed individually.
    1.1.4   AutoApproveWindowsUpdates toggle (default off): Security/Critical no longer
            auto-approved in Autopilot without explicit opt-in.
    1.1.5   Docs: menu screenshots + Word doc mockup nav fix.
    1.1.7   Bugfix: HTML report — phantom certs filtered (DateTime.MinValue), cert expiry
            uses TotalDays, Root CA display shows '(not set)' when absent, NetBIOS
            count null-safe; Register-ExecutedCommand for IANA timezone moved inside
            Enable-IanaTimeZoneMappings so log only shows actual execution.
    1.1.6   Bugfix: EnableDownloadDomains org flag now set (CVE-2021-1730 was incomplete);
            PowerShell VDir sets ExternalUrl only (InternalUrl stays http); NetBIOS report
            now checks registry for pending-reboot state; OWA EP integer normalization;
            certificate expiry used .Days (days-component) instead of TotalDays.

    Original Install-Exchange15.ps1 by Michel de Rooij (forked after v4.23):

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
    [ValidateScript({ if ((Test-Path -Path $_ -PathType Container) -or (Get-DiskImage -ImagePath $_)) { $true } else { throw ('Specified source path or image {0} not found or inaccessible' -f $_) } })]
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
    [switch]$German,
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
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$MEACNotificationEmail,

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

    $ScriptVersion = '1.1.5'

    $ERR_OK = 0
    $ERR_PROBLEMADPREPARE = 1001
    $ERR_UNEXPECTEDOS = 1002
    $ERR_UNEXPTECTEDPHASE = 1003
    $ERR_PROBLEMADDINGFEATURE = 1004
    $ERR_NOTDOMAINJOINED = 1005
    $ERR_NOFIXEDIPADDRESS = 1006
    $ERR_CANTCREATETEMPFOLDER = 1007
    $ERR_UNKNOWNROLESSPECIFIED = 1008
    $ERR_NOACCOUNTSPECIFIED = 1009
    $ERR_RUNNINGNONADMINMODE = 1010
    $ERR_AUTOPILOTNOSTATEFILE = 1011
    $ERR_ADMIXEDMODE = 1012
    $ERR_ADFORESTLEVEL = 1013
    $ERR_INVALIDCREDENTIALS = 1014
    $ERR_MDBDBLOGPATH = 1016
    $ERR_MISSINGORGANIZATIONNAME = 1017
    $ERR_ORGANIZATIONNAMEMISMATCH = 1018
    $ERR_RUNNINGNONENTERPRISEADMIN = 1019
    $ERR_RUNNINGNONSCHEMAADMIN = 1020
    $ERR_COULDNOTDETERMINEADSITE = 1021
    $ERR_PROBLEMPACKAGEDL = 1120
    $ERR_PROBLEMPACKAGESETUP = 1121
    $ERR_PROBLEMPACKAGEEXTRACT = 1122
    $ERR_BADFORESTLEVEL = 1151
    $ERR_BADDOMAINLEVEL = 1152
    $ERR_MISSINGEXCHANGESETUP = 1201
    $ERR_PROBLEMEXCHANGESETUP = 1202
    $ERR_PROBLEMEXCHANGESERVEREXISTS = 1203
    $ERR_EX19EX2013COEXIST = 1204
    $ERR_UNSUPPORTEDEX = 1205
    $ERR_PREFLIGHTFAILED = 1030
    $ERR_CONFIGEXPORTFAILED = 1031
    $ERR_CONFIGIMPORTFAILED = 1032
    $ERR_CERTIMPORTFAILED = 1033
    $ERR_DAGJOIN = 1034
    $ERR_SOURCESERVERCONNECT = 1036
    $ERR_MEACPREPAREAD       = 1037

    $COUNTDOWN_TIMER = 10
    $WU_DOWNLOAD_TIMEOUT_SEC = 3600  # seconds before a stalled Windows Update download is aborted (60 min)
    $DOMAIN_MIXEDMODE = 0
    $FOREST_LEVEL2012 = 5
    $FOREST_LEVEL2012R2 = 6

    # Minimum FFL/DFL levels
    $EX2016_MINFORESTLEVEL = 15317
    $EX2016_MINDOMAINLEVEL = 13236
    $EX2019_MINFORESTLEVEL = 17000
    $EX2019_MINDOMAINLEVEL = 13236

    # Exchange Versions
    $EX2016_MAJOR = '15.1'
    $EX2019_MAJOR = '15.2'

    # Exchange Install registry key
    $EXCHANGEINSTALLKEY = "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup"

    # Autodiscover SCP LDAP filter template ({0} = server name)
    $AUTODISCOVER_SCP_FILTER = '(&(cn={0})(objectClass=serviceConnectionPoint)(serviceClassName=ms-Exchange-AutoDiscover-Service)(|(keywords=67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68)(keywords=77378F46-2C66-4aa9-A6A6-3E7A48B19596)))'
    # Max retries for Autodiscover SCP background jobs (30 x 10s = 5 min)
    $AUTODISCOVER_SCP_MAX_RETRIES = 30

    # Supported Exchange versions (setup.exe)
    $EX2016SETUPEXE_CU23 = '15.01.2507.006'
    $EX2019SETUPEXE_CU10 = '15.02.0922.007'
    $EX2019SETUPEXE_CU11 = '15.02.0986.005'
    $EX2019SETUPEXE_CU12 = '15.02.1118.007'
    $EX2019SETUPEXE_CU13 = '15.02.1258.012'
    $EX2019SETUPEXE_CU14 = '15.02.1544.004'
    $EX2019SETUPEXE_CU15 = '15.02.1748.008'
    $EXSESETUPEXE_RTM = '15.02.2562.017'

    # Supported Operating Systems
    $WS2016_MAJOR = '10.0'
    $WS2019_PREFULL = '10.0.17709'
    $WS2022_PREFULL = '10.0.20348'
    $WS2025_PREFULL = '10.0.26100'

    # .NET Framework Versions
    $NETVERSION_48 = 528040
    $NETVERSION_481 = 533320

    # Package exit codes
    $ERR_SUS_NOT_APPLICABLE = -2145124329   # SUS_E_NOT_APPLICABLE: package not applicable or already installed

    # Power plan GUIDs
    $POWERPLAN_HIGH_PERFORMANCE = '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'

    # FFL
    $FFL_2003 = 2
    $FFL_2008 = 3
    $FFL_2008R2 = 4
    $FFL_2012 = 5
    $FFL_2012R2 = 6
    $FFL_2016 = 7
    $FFL_2025 = 10

    function Save-State( $State) {
        Write-MyVerbose "Saving state information to $StateFile"
        Export-Clixml -InputObject $State -Path $StateFile
    }

    function Restore-State() {
        $State = @{}
        if (Test-Path $StateFile) {
            try {
                $State = Import-Clixml -Path $StateFile -ErrorAction Stop
                # Validate essential state properties
                if ($State -isnot [hashtable]) {
                    Write-MyWarning 'State file is corrupt (not a hashtable), starting fresh'
                    $State = @{}
                }
                else {
                    Write-Verbose "State information loaded from $StateFile"
                }
            }
            catch {
                Write-MyWarning ('Failed to load state file, starting fresh: {0}' -f $_.Exception.Message)
                $State = @{}
            }
        }
        else {
            Write-Verbose "No state file found at $StateFile"
        }
        return $State
    }


    function Get-OSVersionText( $OSVersion) {
        # Maps Windows build numbers to human-readable product names
        $builds = @{
            '10.0.14393' = 'Windows Server 2016'
            '10.0.17763' = 'Windows Server 2019'
            '10.0.20348' = 'Windows Server 2022'
            '10.0.26100' = 'Windows Server 2025'
        }
        $text = $builds[$OSVersion]
        if (-not $text) {
            # Unknown build — fall back to closest known version
            $text = ($builds.GetEnumerator() |
                Where-Object { [System.Version]$_.Key -le [System.Version]$OSVersion } |
                Sort-Object { [System.Version]$_.Key } |
                Select-Object -Last 1).Value
            if (-not $text) { $text = 'Windows Server (unknown)' }
        }
        return '{0} (build {1})' -f $text, $OSVersion
    }

    function Get-SetupTextVersion( $FileVersion) {
        $Versions = @{
            $EX2016SETUPEXE_CU23 = 'Exchange Server 2016 Cumulative Update 23'
            $EX2019SETUPEXE_CU10 = 'Exchange Server 2019 CU10'
            $EX2019SETUPEXE_CU11 = 'Exchange Server 2019 CU11'
            $EX2019SETUPEXE_CU12 = 'Exchange Server 2019 CU12'
            $EX2019SETUPEXE_CU13 = 'Exchange Server 2019 CU13'
            $EX2019SETUPEXE_CU14 = 'Exchange Server 2019 CU14'
            $EX2019SETUPEXE_CU15 = 'Exchange Server 2019 CU15'
            $EXSESETUPEXE_RTM    = 'Exchange Server SE RTM'
        }
        # Direct lookup first (exact CU build match)
        if ($Versions.ContainsKey($FileVersion)) {
            return '{0} (build {1})' -f $Versions[$FileVersion], $FileVersion
        }
        # Fallback: highest known CU version <= FileVersion (covers SU builds)
        $res = "Unsupported version (build $FileVersion)"
        $Versions.GetEnumerator() | Sort-Object -Property { [System.Version]$_.Key } | ForEach-Object {
            if ( [System.Version]$FileVersion -ge [System.Version]$_.Key) {
                $res = '{0} (build {1})' -f $_.Value, $FileVersion
            }
        }
        return $res
    }

    function Get-DetectedFileVersion( $File) {
        # Use FileVersionInfo directly — Get-Command triggers PowerShell command discovery
        # (PATH lookup, module analysis) which adds unnecessary overhead on ISO-mounted paths.
        if ( Test-Path $File) {
            return [System.Diagnostics.FileVersionInfo]::GetVersionInfo($File).ProductVersion
        }
        return 0
    }

    function Write-ToTranscript( $Level, $Text) {
        # Three tiers (single log file):
        #   Default     : INFO / WARNING / ERROR / EXE
        #   -Verbose    : + VERBOSE
        #   -Debug      : + DEBUG + SUPPRESSED-ERROR diff from $Error
        # Encoding note: PS 5.1 `Out-File` defaults to Unicode (UTF-16LE w/ BOM); mixing that
        # with the UTF-8 header produces "strange font" output in viewers. We pin UTF-8 (no BOM)
        # via [IO.File]::AppendAllText so every line in the file has the same encoding.
        if (-not $State['TranscriptFile']) { return }
        $Location = Split-Path $State['TranscriptFile'] -Parent
        if (-not (Test-Path $Location)) { return }
        $verboseOn = [bool]$State['LogVerbose']
        $debugOn   = [bool]$State['LogDebug']
        $shouldWrite = switch ($Level) {
            'VERBOSE' { $verboseOn -or $debugOn }
            'DEBUG'   { $debugOn }
            default   { $true }
        }
        $utf8 = [System.Text.UTF8Encoding]::new($false)
        if ($shouldWrite) {
            try {
                [System.IO.File]::AppendAllText($State['TranscriptFile'], ("{0}: [{1}] {2}`r`n" -f (Get-Date -Format u), $Level, $Text), $utf8)
            } catch { }
        }
        if ($debugOn) {
            try {
                $cur = $Error.Count
                if ($cur -gt $script:lastErrorCount) {
                    $newCount = $cur - $script:lastErrorCount
                    for ($i = $newCount - 1; $i -ge 0; $i--) {
                        $e = $Error[$i]
                        if (-not $e) { continue }
                        $inv = $e.InvocationInfo
                        $ln  = if ($inv) { $inv.ScriptLineNumber } else { '?' }
                        $cmd = if ($inv) { ($inv.Line -replace '\s+', ' ').Trim() } else { '' }
                        $typ = if ($e.Exception) { $e.Exception.GetType().FullName } else { 'Error' }
                        $msg = if ($e.Exception) { $e.Exception.Message } else { [string]$e }
                        $line = '{0}: [SUPPRESSED-ERROR] ({1}) at line {2}: {3} :: {4}' -f (Get-Date -Format u), $typ, $ln, $cmd, $msg
                        [System.IO.File]::AppendAllText($State['TranscriptFile'], ($line + "`r`n"), $utf8)
                    }
                    $script:lastErrorCount = $cur
                }
            } catch { }
        }
    }

    function Write-MyOutput( $Text) {
        Write-Output $Text
        Write-ToTranscript 'INFO' $Text
    }

    function Write-MyWarning( $Text) {
        Write-Warning $Text
        Write-ToTranscript 'WARNING' $Text
    }

    function Write-MyError( $Text) {
        Write-Error $Text
        Write-ToTranscript 'ERROR' $Text
    }

    function Write-MyVerbose( $Text) {
        Write-Verbose $Text
        Write-ToTranscript 'VERBOSE' $Text
    }

    function Write-MyDebug( $Text) {
        # Console stays silent; log line appears only when -Debug tier active.
        Write-ToTranscript 'DEBUG' $Text
    }

    # Records configuration-level commands the script actually executed, so
    # chapter 14 of the Installation Document ("Executed Cmdlets") can list them
    # chronologically with exact syntax. Call sites pass the same command line
    # they are about to run — the helper does not re-execute, only records.
    function Register-ExecutedCommand {
        param(
            [Parameter(Mandatory)][string]$Command,
            [string]$Category = ''
        )
        # After Import-Clixml (Restore-State across a reboot) the list comes back as a
        # frozen Deserialized.* type with no .Add() method. Rehydrate it into a live
        # List, preserving any entries recorded before the reboot, on first post-reboot call.
        if (-not $State.ContainsKey('ExecutedCommands') -or $null -eq $State['ExecutedCommands']) {
            $State['ExecutedCommands'] = [System.Collections.Generic.List[object]]::new()
        }
        elseif ($State['ExecutedCommands'] -isnot [System.Collections.Generic.List[object]]) {
            $rehydrated = [System.Collections.Generic.List[object]]::new()
            foreach ($item in @($State['ExecutedCommands'])) { $rehydrated.Add($item) }
            $State['ExecutedCommands'] = $rehydrated
        }
        $State['ExecutedCommands'].Add([pscustomobject]@{
            Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            Phase     = [int]($State['InstallPhase'])
            Category  = $Category
            Command   = $Command
        })
        Write-ToTranscript 'CMD' $Command
    }

    # Native-exe invoker that preserves stdout+stderr for the log. In normal mode output
    # is discarded (same as the old `$null = … 2>$null` pattern). With -Debug, the merged
    # output is written to the main log tagged [EXE] so nothing is hidden.
    function Invoke-NativeCommand {
        param(
            [Parameter(Mandatory)][string]$FilePath,
            [string[]]$Arguments = @(),
            [string]$Tag = ''
        )
        $label = if ($Tag) { $Tag } else { Split-Path $FilePath -Leaf }
        $out = & $FilePath @Arguments 2>&1
        $rc  = $LASTEXITCODE
        if ($State['LogDebug']) {
            Write-ToTranscript 'EXE' ('{0} exit={1} args=[{2}]' -f $label, $rc, ($Arguments -join ' '))
            foreach ($line in $out) {
                if ($null -eq $line) { continue }
                $text = if ($line -is [System.Management.Automation.ErrorRecord]) { 'stderr: ' + $line.Exception.Message } else { [string]$line }
                if ($text) { Write-ToTranscript 'EXE' ('  {0}' -f $text) }
            }
        }
        return $rc
    }

    function Set-RegistryValue {
        param( [string]$Path, [string]$Name, $Value, [string]$PropertyType = 'DWord')
        if ( -not (Test-Path $Path -ErrorAction SilentlyContinue)) {
            New-Item -Path $Path -Force -ErrorAction SilentlyContinue | Out-Null
        }
        else {
            $existing = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
            if ($null -ne $existing -and $existing.$Name -eq $Value) {
                Write-MyVerbose ('Registry value already set: {0}\{1} = {2}' -f $Path, $Name, $Value)
                return
            }
        }
        New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $PropertyType -Force -ErrorAction SilentlyContinue | Out-Null
    }

    function Get-PSExecutionPolicy {
        $PSPolicyKey = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell' -Name ExecutionPolicy -ErrorAction SilentlyContinue
        if ( $PSPolicyKey) {
            Write-MyWarning "PowerShell Execution Policy is set to $($PSPolicyKey.ExecutionPolicy) through GPO"
        }
        else {
            Write-MyVerbose 'PowerShell Execution Policy not configured through GPO'
        }
        return $PSPolicyKey
    }

    function Invoke-WebDownload {
        # PS 5.1-compatible web download. Uses -SkipCertificateCheck on PS 6+,
        # falls back to WebClient with TLS 1.2 and cert bypass on PS 5.1.
        param([string]$Uri, [string]$OutFile)
        if ($PSVersionTable.PSVersion.Major -ge 6) {
            Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing -SkipCertificateCheck -ErrorAction Stop
        }
        else {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $prevCallback = [Net.ServicePointManager]::ServerCertificateValidationCallback
            [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
            try {
                $wc = New-Object System.Net.WebClient
                $wc.DownloadFile($Uri, $OutFile)
            }
            finally {
                [Net.ServicePointManager]::ServerCertificateValidationCallback = $prevCallback
            }
        }
    }

    function Get-MyPackage () {
        param ( [String]$Package, [String]$URL, [String]$FileName, [String]$InstallPath)
        $res = $true
        if ( !( Test-Path $(Join-Path $InstallPath $Filename))) {
            if ( $URL) {
                Write-MyOutput "Package $Package not found, downloading to $FileName"
                Write-MyVerbose "Source: $URL"
                $destPath = Join-Path $InstallPath $Filename
                $downloaded = $false
                $savedPP = $ProgressPreference
                $ProgressPreference = 'SilentlyContinue'
                for ($attempt = 1; $attempt -le 3; $attempt++) {
                    try {
                        Start-BitsTransfer -Source $URL -Destination $destPath -ErrorAction Stop
                        $downloaded = $true
                        break
                    }
                    catch {
                        Get-BitsTransfer -ErrorAction SilentlyContinue | Where-Object { $_.JobState -notin 'Transferred','Acknowledged' } | Remove-BitsTransfer -ErrorAction SilentlyContinue
                        Remove-Item -Path $destPath -ErrorAction SilentlyContinue
                        # 0x800704DD = ERROR_NOT_LOGGED_ON: BITS has no network logon session
                        # (common in Autopilot RunOnce context after reboot). Fall back to
                        # WebClient immediately — no point retrying BITS in this scenario.
                        $isBitsLogonError = $_.Exception.Message -match '0x800704DD|not logged on to the network'
                        if ($attempt -lt 3 -and -not $isBitsLogonError) {
                            Write-MyWarning ('Download attempt {0}/3 failed, retrying in {1} seconds: {2}' -f $attempt, ($attempt * 5), $_.Exception.Message)
                            Start-Sleep -Seconds ($attempt * 5)
                        }
                        else {
                            # Final attempt or BITS network-logon error: try web download as fallback
                            try {
                                if ($isBitsLogonError) {
                                    Write-MyVerbose 'BITS unavailable (no network logon session) — using web download'
                                } else {
                                    Write-MyVerbose 'BITS failed after 3 attempts, trying web download as fallback'
                                }
                                Invoke-WebDownload -Uri $URL -OutFile $destPath
                                $downloaded = $true
                                break
                            }
                            catch {
                                Write-MyWarning ('Problem downloading package from URL: {0}' -f $_.Exception.Message)
                                Remove-Item -Path $destPath -ErrorAction SilentlyContinue
                            }
                        }
                    }
                }
                $ProgressPreference = $savedPP
                if (-not $downloaded) {
                    $res = $false
                    Write-MyWarning ('Could not download {0}. For offline or proxy-restricted deployments:' -f $FileName)
                    Write-MyOutput  ('  1. Run  .\tools\Get-EXpressDownloads.ps1  on an internet-connected machine.')
                    Write-MyOutput  ('  2. Copy the sources\ folder to {0}' -f $InstallPath)
                }
            }
            else {
                Write-MyWarning "$FileName not present, skipping"
                $res = $false
            }
        }
        else {
            Write-MyVerbose "Located $Package ($InstallPath\$FileName)"
        }
        return $res
    }

    function Get-CurrentUserName {
        return [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }

    function Test-Admin {
        $currentPrincipal = New-Object System.Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )
        return $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )
    }

    function Test-RebootPending {
        # Returns $true if Windows signals a pending reboot. Used to decide whether
        # a phase boundary really needs to reboot, or if we can continue in-process.
        $reasons = @()
        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') {
            $reasons += 'CBS RebootPending'
        }
        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') {
            $reasons += 'WindowsUpdate RebootRequired'
        }
        $pfro = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
        if ($pfro -and $pfro.PendingFileRenameOperations) {
            $reasons += 'PendingFileRenameOperations'
        }
        $cn = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName' -Name 'ComputerName' -ErrorAction SilentlyContinue
        $pcn = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName' -Name 'ComputerName' -ErrorAction SilentlyContinue
        if ($cn -and $pcn -and $cn.ComputerName -ne $pcn.ComputerName) {
            $reasons += 'Pending computer rename'
        }
        try {
            $ccm = Invoke-CimMethod -Namespace 'ROOT\ccm\ClientSDK' -ClassName 'CCM_ClientUtilities' -MethodName 'DetermineIfRebootPending' -ErrorAction Stop
            if ($ccm -and ($ccm.RebootPending -or $ccm.IsHardRebootPending)) {
                $reasons += 'CCM ClientSDK'
            }
        } catch { }
        if ($reasons.Count -gt 0) {
            Write-MyVerbose ('Reboot pending: {0}' -f ($reasons -join ', '))
            return $true
        }
        return $false
    }

    function Test-ADGroupMember ([int]$RelativeId) {
        try {
            $FRNC = Get-ForestRootNC
            $ADRootSID = ([ADSI]"LDAP://$FRNC").ObjectSID[0]
            if ($null -eq $ADRootSID) {
                Write-MyWarning 'Could not retrieve forest root SID — AD may be unreachable'
                return $false
            }
            $SID = (New-Object System.Security.Principal.SecurityIdentifier ($ADRootSID, 0)).Value.toString()
            return [Security.Principal.WindowsIdentity]::GetCurrent().Groups | Where-Object { $_.Value -eq "$SID-$RelativeId" }
        }
        catch {
            Write-MyWarning ('Test-ADGroupMember failed: {0}' -f $_.Exception.Message)
            return $false
        }
    }

    function Test-SchemaAdmin     { Test-ADGroupMember 518 }
    function Test-EnterpriseAdmin { Test-ADGroupMember 519 }

    function Test-ServerCore {
        (Get-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows NT\CurrentVersion' -Name 'InstallationType' -ErrorAction SilentlyContinue).InstallationType -eq 'Server Core'
    }

    function Test-RebootPending {
        $Pending = $False
        if ( Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue) {
            $Pending = $True
        }
        if ( Test-Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending' -ErrorAction SilentlyContinue) {
            $Pending = $True
        }
        if ( Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired' -ErrorAction SilentlyContinue) {
            $Pending = $True
        }
        return $Pending
    }

    function Enable-RunOnce {
        Write-MyOutput 'Set script to run once after reboot'
        # When compiled with PS2Exe the script runs as a standalone .exe — invoke it directly.
        # Otherwise use the current PowerShell interpreter (powershell.exe or pwsh.exe).
        $isExe = $ScriptFullName -imatch '\.exe$'
        $logFlags = ''
        if ($State['LogVerbose']) { $logFlags += ' -Verbose' }
        if ($State['LogDebug'])   { $logFlags += ' -Debug' }
        if ($isExe) {
            $RunOnce = "`"$ScriptFullName`" -InstallPath `"$InstallPath`"$logFlags"
        }
        else {
            $PSExe = (Get-Process -Id $PID).Path
            $RunOnce = "`"$PSExe`" -NoProfile -ExecutionPolicy Unrestricted -Command `"& `'$ScriptFullName`' -InstallPath `'$InstallPath`'$logFlags`""
        }
        Write-MyVerbose "RunOnce: $RunOnce"
        Set-RegistryValue -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce' -Name $ScriptName -Value $RunOnce -PropertyType String
    }

    function Disable-UAC {
        Write-MyVerbose 'Disabling User Account Control'
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 0
    }

    function Enable-UAC {
        Write-MyVerbose 'Enabling User Account Control'
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 1
    }

    function Disable-IEESC {
        $AdminKey = 'HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}'
        $UserKey  = 'HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}'
        $alreadyOff = ((Get-ItemProperty -Path $AdminKey -Name IsInstalled -ErrorAction SilentlyContinue).IsInstalled -eq 0) -and
                      ((Get-ItemProperty -Path $UserKey  -Name IsInstalled -ErrorAction SilentlyContinue).IsInstalled -eq 0)
        if ($alreadyOff) { Write-MyVerbose 'IE Enhanced Security Configuration already disabled'; return }
        Write-MyOutput 'Disabling IE Enhanced Security Configuration'
        New-Item -Path (Split-Path $AdminKey -Parent) -Name (Split-Path $AdminKey -Leaf) -ErrorAction SilentlyContinue | Out-Null
        Set-ItemProperty -Path $AdminKey -Name 'IsInstalled' -Value 0 -Force | Out-Null
        New-Item -Path (Split-Path $UserKey -Parent) -Name (Split-Path $UserKey -Leaf) -ErrorAction SilentlyContinue | Out-Null
        Set-ItemProperty -Path $UserKey  -Name 'IsInstalled' -Value 0 -Force | Out-Null
        if ( Get-Process -Name explorer.exe -ErrorAction SilentlyContinue) {
            Stop-Process -Name Explorer
        }
    }

    function Enable-IEESC {
        Write-MyVerbose 'Enabling IE Enhanced Security Configuration'
        $AdminKey = 'HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}'
        $UserKey = 'HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}'
        New-Item -Path (Split-Path $AdminKey -Parent) -Name (Split-Path $AdminKey -Leaf) -ErrorAction SilentlyContinue | Out-Null
        Set-ItemProperty -Path $AdminKey -Name 'IsInstalled' -Value 1 -Force | Out-Null
        New-Item -Path (Split-Path $UserKey -Parent) -Name (Split-Path $UserKey -Leaf) -ErrorAction SilentlyContinue | Out-Null
        Set-ItemProperty -Path $UserKey -Name 'IsInstalled' -Value 1 -Force | Out-Null
        if ( Get-Process -Name explorer.exe -ErrorAction SilentlyContinue) {
            Stop-Process -Name Explorer
        }
    }

    function Get-FullDomainAccount {
        $PlainTextAccount = $State['AdminAccount']
        if ( $PlainTextAccount.indexOf('\') -gt 0) {
            $Parts = $PlainTextAccount.split('\')
            $Domain = $Parts[0]
            $UserName = $Parts[1]
            return "$Domain\$UserName"
        }
        else {
            if ( $PlainTextAccount.indexOf('@') -gt 0) {
                return $PlainTextAccount
            }
            else {
                $Domain = $env:USERDOMAIN
                $UserName = $PlainTextAccount
                return "$Domain\$UserName"
            }
        }
    }

    function Test-LocalCredential {
        [CmdletBinding()]
        param
        (
            [string]$UserName,
            [string]$ComputerName = $env:COMPUTERNAME,
            [string]$Password
        )
        if (!($UserName) -or !($Password)) {
            Write-Warning 'Test-LocalCredential: Please specify both user name and password'
        }
        else {
            Add-Type -AssemblyName System.DirectoryServices.AccountManagement
            $DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('machine', $ComputerName)
            $DS.ValidateCredentials($UserName, $Password )
        }
    }

    function Test-Credentials {
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR((ConvertTo-SecureString $State['AdminPassword']))
        $PlainTextPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        $FullPlainTextAccount = Get-FullDomainAccount
        try {
            if ( $State['InstallEdge']) {
                $Username = $FullPlainTextAccount.split("\")[-1]
                return $( Test-LocalCredential -UserName $Username -Password $PlainTextPassword)
            }
            else {
                $dc = New-Object DirectoryServices.DirectoryEntry( $Null, $FullPlainTextAccount, $PlainTextPassword)
                if ($dc.Name) {
                    return $true
                }
                else {
                    return $false
                }
            }

        }
        catch {
            return $false
        }
        return $false
    }

    function Get-ValidatedCredentials {
        # Interactive credential prompt with validation retry loop (max 3 attempts).
        # Returns $true when valid credentials are stored in State, $false if all attempts fail.
        # Only call this when [Environment]::UserInteractive is $true.
        #
        # GUI detection: Get-Credential shows a Win32 dialog only when all three hold:
        #   1. ConsoleHost (not ISE, not PS2Exe, not a remote host)
        #   2. UserInteractive (not a service / scheduled-task session)
        #   3. A real desktop session (SESSIONNAME = Console or RDP-*; empty = Session 0 / no window station)
        # When any condition is false we go straight to Read-Host to avoid the silent-$null fallback.
        $sessionName = [string]$env:SESSIONNAME
        $useGui = (-not $IsPS2Exe) -and
                  ($Host.Name -eq 'ConsoleHost') -and
                  [Environment]::UserInteractive -and
                  ($sessionName -match '^(Console|RDP)')

        $maxAttempts = 3
        for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
            try {
                $defaultUser = if ($State['AdminAccount']) { $State['AdminAccount'] } else { [System.Security.Principal.WindowsIdentity]::GetCurrent().Name }
                $Script:Credentials = $null
                if ($useGui) {
                    $rawCred = Get-Credential -UserName $defaultUser -Message ('Enter credentials for Autopilot (attempt {0}/{1})' -f $attempt, $maxAttempts)
                    # Get-Credential can return a PSObject wrapper in some terminal environments; unwrap before assigning to typed variable.
                    $Script:Credentials = if ($rawCred -is [pscredential]) { $rawCred }
                                          elseif ($rawCred -and $rawCred.PSObject.BaseObject -is [pscredential]) { $rawCred.PSObject.BaseObject }
                                          else { $null }
                }
                if (-not $Script:Credentials) {
                    Write-MyOutput ('Enter credentials for Autopilot (attempt {0}/{1})' -f $attempt, $maxAttempts)
                    $fbUser = Read-Host -Prompt ('Username [{0}]' -f $defaultUser)
                    if ([string]::IsNullOrWhiteSpace($fbUser)) { $fbUser = $defaultUser }
                    $fbPass = Read-Host -Prompt 'Password' -AsSecureString
                    if ($fbPass -and $fbPass.Length -gt 0) {
                        $Script:Credentials = New-Object System.Management.Automation.PSCredential($fbUser, $fbPass)
                    }
                }
                if (-not $Script:Credentials) {
                    Write-MyWarning 'No credentials entered'
                }
                else {
                    $State['AdminAccount'] = $Script:Credentials.UserName
                    # ConvertFrom-SecureString without -Key uses DPAPI (user+machine bound).
                    # Autopilot always resumes as the same user on the same machine, so this is safe.
                    $State['AdminPassword'] = ($Script:Credentials.Password | ConvertFrom-SecureString)
                    Write-MyOutput ('Checking credentials (attempt {0}/{1})' -f $attempt, $maxAttempts)
                    if (Test-Credentials) {
                        Write-MyOutput 'Credentials valid'
                        return $true
                    }
                    else {
                        Write-MyWarning ("Credentials for '{0}' are invalid" -f $State['AdminAccount'])
                    }
                }
            }
            catch {
                Write-MyWarning ('Credential prompt cancelled or failed: {0}' -f $_.Exception.Message)
            }
            if ($attempt -lt $maxAttempts) {
                $choice = $Host.UI.PromptForChoice('Invalid credentials', 'Retry or quit?', @('&Retry', '&Quit'), 0)
                if ($choice -ne 0) {
                    Write-MyError 'Credential entry aborted by user'
                    return $false
                }
            }
        }
        Write-MyError ('Credential validation failed after {0} attempts' -f $maxAttempts)
        return $false
    }

    function Enable-AutoLogon {
        Write-MyVerbose 'Enabling Automatic Logon'
        # SECURITY NOTE: This writes the password in plaintext to the registry.
        # Disable-AutoLogon is called after the next login to remove these values immediately.
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR((ConvertTo-SecureString $State['AdminPassword']))
        $PlainTextPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        $PlainTextAccount = $State['AdminAccount']
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name AutoAdminLogon -Value 1 -PropertyType String
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultUserName -Value $PlainTextAccount -PropertyType String
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultPassword -Value $PlainTextPassword -PropertyType String
    }

    function Disable-AutoLogon {
        Write-MyVerbose 'Disabling Automatic Logon'
        Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name AutoAdminLogon -ErrorAction SilentlyContinue | Out-Null
        Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultUserName -ErrorAction SilentlyContinue | Out-Null
        Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultPassword -ErrorAction SilentlyContinue | Out-Null
    }

    function Disable-OpenFileSecurityWarning {
        Write-MyVerbose 'Disabling File Security Warning dialog'
        Set-RegistryValue -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations' -Name 'LowRiskFileTypes' -Value '.exe;.msp;.msu;.msi' -PropertyType String
        Set-RegistryValue -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments' -Name 'SaveZoneInformation' -Value 1
        Remove-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations' -Name 'LowRiskFileTypes' -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments' -Name 'SaveZoneInformation' -ErrorAction SilentlyContinue
    }

    function Enable-OpenFileSecurityWarning {
        Write-MyVerbose 'Enabling File Security Warning dialog'
        Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations' -Name 'LowRiskFileTypes' -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments' -Name 'SaveZoneInformation' -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations' -Name 'LowRiskFileTypes' -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments' -Name 'SaveZoneInformation' -ErrorAction SilentlyContinue
    }

    function Invoke-Extract ( $FilePath, $FileName) {
        Write-MyVerbose "Extracting $FilePath\$FileName to $FilePath"
        $FullPath = Join-Path $FilePath $FileName
        if ( Test-Path $FullPath) {
            $TempNam = "$FullPath.zip"
            try {
                Copy-Item $FullPath $TempNam -Force -ErrorAction Stop
                Expand-Archive -Path $TempNam -DestinationPath $FilePath -Force -ErrorAction Stop
            }
            catch {
                Write-MyError ('Failed to extract {0}: {1}' -f $FullPath, $_.Exception.Message)
            }
            finally {
                Remove-Item $TempNam -ErrorAction SilentlyContinue
            }
        }
        else {
            Write-MyWarning "$FilePath\$FileName not found"
        }
    }

    function Invoke-Process ( $FilePath, $FileName, $ArgumentList) {
        $rval = 0
        $mspTempDir = $null
        $FullName = Join-Path $FilePath $FileName
        if ( Test-Path $FullName) {
            switch ( ([io.fileinfo]$Filename).extension.ToUpper()) {
                '.MSU' {
                    $ArgumentList += @( $FullName)
                    $ArgumentList += @( '/f')
                    $Cmd = "$env:SystemRoot\System32\WUSA.EXE"
                }
                '.MSI' {
                    $ArgumentList += @( '/i')
                    $ArgumentList += @( $FullName)
                    $Cmd = "MSIEXEC.EXE"
                }
                '.MSP' {
                    $ArgumentList += @( '/update')
                    $ArgumentList += @( $FullName)
                    $Cmd = 'MSIEXEC.EXE'
                }
                '.CAB' {
                    $mspTempDir = Join-Path $env:TEMP ('ExSU_' + [IO.Path]::GetFileNameWithoutExtension($FileName))
                    New-Item -ItemType Directory -Path $mspTempDir -Force | Out-Null
                    $expandOut = & "$env:SystemRoot\System32\expand.exe" -F:* $FullName $mspTempDir 2>&1
                    Write-MyVerbose ('expand.exe output: {0}' -f ($expandOut -join ' | '))
                    # Exchange SU CABs are often multi-level: expand any nested CABs into the same temp dir
                    $nestedCabs = Get-ChildItem -Path $mspTempDir -Filter '*.cab' -File -ErrorAction SilentlyContinue
                    foreach ($nestedCab in $nestedCabs) {
                        $nestedOut = & "$env:SystemRoot\System32\expand.exe" -F:* $nestedCab.FullName $mspTempDir 2>&1
                        Write-MyVerbose ('Nested CAB {0}: {1}' -f $nestedCab.Name, ($nestedOut -join ' | '))
                    }
                    $extractedFiles = Get-ChildItem -Path $mspTempDir -Recurse -File -ErrorAction SilentlyContinue
                    if ($extractedFiles) {
                        Write-MyVerbose ('CAB contents: {0}' -f ($extractedFiles.Name -join ', '))
                    } else {
                        Write-MyVerbose 'CAB extraction produced no files'
                    }
                    $mspFile = $extractedFiles | Where-Object { $_.Extension -eq '.msp' } | Select-Object -First 1
                    $exeFile = $extractedFiles | Where-Object { $_.Extension -eq '.exe' -and $_.Name -notlike '*.cab' } | Select-Object -First 1
                    if ($mspFile) {
                        $ArgumentList += @('/update')
                        $ArgumentList += @($mspFile.FullName)
                        $Cmd = 'MSIEXEC.EXE'
                    }
                    elseif ($exeFile) {
                        $Cmd = $exeFile.FullName
                    }
                    else {
                        # No MSP/EXE found — WU-style CABs (Exchange SE SU) carry a compressed payload
                        # that expand.exe cannot unpack as MSP. Install directly via DISM /Add-Package.
                        Write-MyVerbose ('No MSP or EXE found in {0} — falling back to DISM /Add-Package' -f $FileName)
                        Remove-Item -Path $mspTempDir -Recurse -Force -ErrorAction SilentlyContinue
                        $mspTempDir = $null
                        $Cmd = "$env:SystemRoot\System32\dism.exe"
                        $ArgumentList = @('/Online', '/Add-Package', "/PackagePath:$FullName", '/Quiet', '/NoRestart')
                    }
                }
                default {
                    $Cmd = $FullName
                }
            }
            Write-MyVerbose "Executing $Cmd $($ArgumentList -Join ' ')"
            $rval = ( Start-Process -FilePath $Cmd -ArgumentList $ArgumentList -NoNewWindow -PassThru -Wait).Exitcode
            Write-MyVerbose "Process exited with code $rval"
            if ($mspTempDir) { Remove-Item -Path $mspTempDir -Recurse -Force -ErrorAction SilentlyContinue }
        }
        else {
            Write-MyWarning "$FullName not found"
            $rval = -1
        }
        return $rval
    }
    function Get-ForestRootNC {
        try {
            return ([ADSI]'LDAP://RootDSE').rootDomainNamingContext.toString()
        }
        catch {
            Write-MyError ('Cannot read Forest Root Naming Context (LDAP://RootDSE): {0}' -f $_.Exception.Message)
            return $null
        }
    }
    function Get-RootNC {
        try {
            return ([ADSI]'').distinguishedName.toString()
        }
        catch {
            Write-MyError ('Cannot read Root Naming Context: {0}' -f $_.Exception.Message)
            return $null
        }
    }

    function Get-ForestConfigurationNC {
        try {
            return ([ADSI]'LDAP://RootDSE').configurationNamingContext.toString()
        }
        catch {
            Write-MyError ('Cannot read Forest Configuration Naming Context: {0}' -f $_.Exception.Message)
            return $null
        }
    }

    function Get-ForestFunctionalLevel {
        $CNC = Get-ForestConfigurationNC
        try {
            $rval = ( ([ADSI]"LDAP://cn=partitions,$CNC").get('msDS-Behavior-Version') )
        }
        catch {
            Write-MyError "Can't read Forest schema version, operator possibly not member of Schema Admin group"
        }
        return $rval
    }

    function Test-DomainNativeMode {
        $NC = Get-RootNC
        return( ([ADSI]"LDAP://$NC").ntMixedDomain )
    }

    function Get-ExchangeOrganization {
        $CNC = Get-ForestConfigurationNC
        try {
            $ExOrgContainer = [ADSI]"LDAP://CN=Microsoft Exchange,CN=Services,$CNC"
            $rval = ($ExOrgContainer.PSBase.Children | Where-Object { $_.objectClass -eq 'msExchOrganizationContainer' }).Name
        }
        catch {
            Write-MyVerbose "Can't find Exchange Organization object"
            $rval = $null
        }
        return $rval
    }

    function Get-ExchangeDAGNames {
        try {
            $CNC  = Get-ForestConfigurationNC
            $exOrg = Get-ExchangeOrganization
            if (-not $exOrg) { return @() }
            $root   = [ADSI]"LDAP://CN=$exOrg,CN=Microsoft Exchange,CN=Services,$CNC"
            $result = $root.PSBase.Children | Where-Object { $_.objectClass -contains 'msExchMDBAvailabilityGroup' } |
                      ForEach-Object { [string]$_.Name }
            return @($result | Where-Object { $_ })
        } catch { return @() }
    }

    function Test-ExchangeOrganization( $Organization) {
        $CNC = Get-ForestConfigurationNC
        return( [ADSI]"LDAP://CN=$Organization,CN=Microsoft Exchange,CN=Services,$CNC")
    }

    function Get-ExchangeForestLevel {
        $CNC = Get-ForestConfigurationNC
        return ( ([ADSI]"LDAP://CN=ms-Exch-Schema-Version-Pt,CN=Schema,$CNC").rangeUpper )
    }

    function Get-ExchangeDomainLevel {
        $NC = Get-RootNC
        return( ([ADSI]"LDAP://CN=Microsoft Exchange System Objects,$NC").objectVersion )
    }

    function Add-BackgroundJob {
        param([System.Management.Automation.Job]$Job)
        if (-not $Global:BackgroundJobs) { $Global:BackgroundJobs = @() }
        # Prune completed/failed/stopped jobs to prevent unbounded list growth
        $Global:BackgroundJobs = @($Global:BackgroundJobs | Where-Object { $_.State -notin @('Completed', 'Failed', 'Stopped') })
        $Global:BackgroundJobs += $Job
    }

    function New-LDAPSearch {
        param([string]$ConfigNC, [string]$Filter)
        $s = New-Object System.DirectoryServices.DirectorySearcher
        $s.SearchRoot = "LDAP://$ConfigNC"
        $s.Filter = $Filter
        return $s
    }

    function Clear-AutodiscoverServiceConnectionPoint( [string]$Name, [switch]$Wait) {
        $ConfigNC = Get-ForestConfigurationNC
        if ($Wait) {
            $ScriptBlock = {
                param($ServerName, $ConfigNC, $FilterTemplate, $MaxRetries)
                $retries = 0
                do {
                    if ($null -ne $ConfigNC) {
                        $LDAPSearch = New-Object System.DirectoryServices.DirectorySearcher
                        $LDAPSearch.SearchRoot = 'LDAP://{0}' -f $ConfigNC
                        $LDAPSearch.Filter = $FilterTemplate -f $ServerName

                        $Results = $LDAPSearch.FindAll()
                        if ($Results.Count -gt 0) {
                            $Results | ForEach-Object {
                                Write-Host ('Removing object {0}' -f $_.Path)
                                try {
                                    ([ADSI]($_.Path)).DeleteTree()
                                    Write-Host ('Successfully cleared AutodiscoverServiceConnectionPoint for {0}' -f $ServerName)
                                }
                                catch {
                                    Write-Error ('Problem clearing AutodiscoverServiceConnectionPoint for {0}: {1}' -f $ServerName, $_.Exception.Message)
                                }
                            }
                            return $true
                        }
                        else {
                            $retries++
                            if ($retries -ge $MaxRetries) {
                                Write-Error ('AutodiscoverServiceConnectionPoint for {0} not found after {1} retries, giving up.' -f $ServerName, $MaxRetries)
                                return $false
                            }
                            Write-Host ('AutodiscoverServiceConnectionPoint not found for {0}, waiting a bit ..' -f $ServerName)
                            Start-Sleep -Seconds 10
                        }
                    }
                } while ($true)
            }

            $Job = Start-Job -ScriptBlock $ScriptBlock -ArgumentList $Name, $ConfigNC, $AUTODISCOVER_SCP_FILTER, $AUTODISCOVER_SCP_MAX_RETRIES -Name ('Clear-AutodiscoverSCP-{0}' -f $Name)
            Add-BackgroundJob $Job
            Write-MyOutput ('Started background job to clear AutodiscoverServiceConnectionPoint for {0} (Job ID: {1})' -f $Name, $Job.Id)
            return $Job
        }
        else {
            $LDAPSearch = New-LDAPSearch -ConfigNC $ConfigNC -Filter ($AUTODISCOVER_SCP_FILTER -f $Name)
            $LDAPSearch.FindAll() | ForEach-Object {

                Write-MyVerbose ('Removing object {0}' -f $_.Path)
                try {
                    ([ADSI]($_.Path)).DeleteTree()
                }
                catch {
                    Write-MyError ('Problem clearing serviceBindingInformation property on {0}: {1}' -f $_.Path, $_.Exception.Message)
                }
            }
        }
    }

    function Set-AutodiscoverServiceConnectionPoint( [string]$Name, [string]$ServiceBinding, [switch]$Wait) {
        $ConfigNC = Get-ForestConfigurationNC
        if ($Wait) {
            $ScriptBlock = {
                param($ServerName, $ConfigNC, $serviceBindingValue, $FilterTemplate, $MaxRetries)
                $retries = 0
                do {
                    if ($null -ne $ConfigNC) {
                        $LDAPSearch = New-Object System.DirectoryServices.DirectorySearcher
                        $LDAPSearch.SearchRoot = 'LDAP://{0}' -f $ConfigNC
                        $LDAPSearch.Filter = $FilterTemplate -f $ServerName

                        $Results = $LDAPSearch.FindAll()
                        if ($Results.Count -gt 0) {
                            $Results | ForEach-Object {
                                Write-Host ('Setting serviceBindingInformation on {0} to {1}' -f $_.Path, $ServiceBindingValue)
                                try {
                                    $SCPObj = $_.GetDirectoryEntry()
                                    $null = $SCPObj.Put('serviceBindingInformation', $ServiceBindingValue)
                                    $SCPObj.SetInfo()
                                    Write-Host ('Successfully set AutodiscoverServiceConnectionPoint for {0}' -f $ServerName)
                                }
                                catch {
                                    Write-Error ('Problem setting AutodiscoverServiceConnectionPoint for {0}: {1}' -f $ServerName, $_.Exception.Message)
                                }
                            }
                            return $true
                        }
                        else {
                            $retries++
                            if ($retries -ge $MaxRetries) {
                                Write-Error ('AutodiscoverServiceConnectionPoint for {0} not found after {1} retries, giving up.' -f $ServerName, $MaxRetries)
                                return $false
                            }
                            Write-Verbose ('AutodiscoverServiceConnectionPoint not found for {0}, waiting a bit ..' -f $ServerName)
                            Start-Sleep -Seconds 10
                        }
                    }
                } while ($true)
            }

            $Job = Start-Job -ScriptBlock $ScriptBlock -ArgumentList $Name, $ConfigNC, $ServiceBinding, $AUTODISCOVER_SCP_FILTER, $AUTODISCOVER_SCP_MAX_RETRIES -Name ('Set-AutodiscoverSCP-{0}' -f $Name)
            Add-BackgroundJob $Job
            Write-MyVerbose ('Started background job to clear AutodiscoverServiceConnectionPoint for {0} (Job ID: {1})' -f $Name, $Job.Id)
            return $Job
        }
        else {
            $LDAPSearch = New-LDAPSearch -ConfigNC $ConfigNC -Filter ($AUTODISCOVER_SCP_FILTER -f $Name)
            $LDAPSearch.FindAll() | ForEach-Object {
                Write-MyVerbose ('Setting serviceBindingInformation on {0} to {1}' -f $_.Path, $ServiceBinding)
                try {
                    $SCPObj = $_.GetDirectoryEntry()
                    $null = $SCPObj.Put( 'serviceBindingInformation', $ServiceBinding)
                    $SCPObj.SetInfo()
                }
                catch {
                    Write-MyError ('Problem setting serviceBindingInformation property on {0}: {1}' -f $_.Path, $_.Exception.Message)
                }
            }
        }
    }

    function Test-ExistingExchangeServer( [string]$Name) {
        $CNC = Get-ForestConfigurationNC
        $LDAPSearch = New-LDAPSearch -ConfigNC $CNC -Filter "(&(cn=$Name)(objectClass=msExchExchangeServer))"
        $Results = $LDAPSearch.FindAll()
        return ($Results.Count -gt 0)
    }

    function Get-LocalFQDNHostname {
        return ([System.Net.Dns]::GetHostByName(($env:computerName))).HostName
    }

    function Get-ADSite {
        try {
            return [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()
        }
        catch {
            return $null
        }
    }

    function Get-ExchangeServerObjects {
        $CNC = Get-ForestConfigurationNC
        $LDAPSearch = New-LDAPSearch -ConfigNC $CNC -Filter "(objectCategory=msExchExchangeServer)"
        $LDAPSearch.PropertiesToLoad.Add("cn") | Out-Null
        $LDAPSearch.PropertiesToLoad.Add("msExchCurrentServerRoles") | Out-Null
        $LDAPSearch.PropertiesToLoad.Add("serialNumber") | Out-Null
        $Results = $LDAPSearch.FindAll()
        $Results | ForEach-Object {
            [pscustomobject][ordered]@{
                CN                       = $_.Properties.cn[0]
                msExchCurrentServerRoles = $_.Properties.msexchcurrentserverroles[0]
                serialNumber             = $_.Properties.serialnumber[0]
            }
        }
    }

    function Set-EdgeDNSSuffix ([string]$DNSSuffix) {
        Write-MyVerbose 'Setting Primary DNS Suffix'
        #https://technet.microsoft.com/library%28EXCHG.150%29/ms.exch.setupreadiness.FqdnMissing.aspx?f=255&MSPPError=-2147217396
        #Update primary DNS Suffix for FQDN
        Set-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\" -Name Domain -Value $DNSSuffix
        Set-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\" -Name "NV Domain" -Value $DNSSuffix

    }

    function Import-ExchangeModule {
        if ( -not ( Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
            Write-MyVerbose 'Loading Exchange PowerShell module'
            $SetupPath = (Get-ItemProperty -Path $EXCHANGEINSTALLKEY -Name MsiInstallPath -ErrorAction SilentlyContinue).MsiInstallPath
            if (-not $SetupPath) {
                Write-MyWarning "Exchange installation path not found in registry ($EXCHANGEINSTALLKEY)"
                return
            }
            if ( ($State['InstallEdge'] -eq $true -and (Test-Path $(Join-Path $SetupPath "\bin\Exchange.ps1"))) -or ($State['InstallEdge'] -eq $false -and (Test-Path $(Join-Path $SetupPath "\bin\RemoteExchange.ps1")))) {
                if ( $State['InstallEdge']) {
                    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
                    . "$SetupPath\bin\Exchange.ps1" | Out-Null
                }
                else {
                    . "$SetupPath\bin\RemoteExchange.ps1" 6>&1 | Out-Null
                    try {
                        $savedVP = $VerbosePreference
                        $VerbosePreference = 'SilentlyContinue'
                        Connect-ExchangeServer (Get-LocalFQDNHostname) -NoShellBanner 3>&1 6>&1 | Out-Null
                        $VerbosePreference = $savedVP
                    }
                    catch {
                        $VerbosePreference = $savedVP
                        Write-MyError 'Problem loading Exchange module'
                    }
                }
                # Verify essential cmdlets are available
                $requiredCmdlets = @('Get-ExchangeServer', 'Get-MailboxDatabase')
                foreach ($cmdlet in $requiredCmdlets) {
                    if (-not (Get-Command $cmdlet -ErrorAction SilentlyContinue)) {
                        Write-MyWarning ('Exchange module loaded but cmdlet {0} not available' -f $cmdlet)
                    }
                }
            }
            else {
                Write-MyWarning "Can't determine installation path to load Exchange module"
            }
        }
        else {
            Write-MyVerbose 'Exchange module already loaded'
        }
    }

    function Reconnect-ExchangeSession {
        # After W3SVC restarts (ECC/CBC/AMSI), the implicit-remoting PS session that
        # Exchange cmdlets use gets disconnected. Remove the dead session and reconnect.
        Write-MyVerbose 'Reconnecting Exchange PS session after IIS restart'
        Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' } | Remove-PSSession -ErrorAction SilentlyContinue

        # Wait up to 90 s for the Exchange PowerShell endpoint to accept connections
        $maxWait = 90
        $elapsed = 0
        $ready   = $false
        Write-MyVerbose 'Waiting for Exchange PowerShell endpoint to become available'
        do {
            Start-Sleep -Seconds 5
            $elapsed += 5
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                $wc = New-Object System.Net.WebClient
                try { $null = $wc.DownloadString('http://localhost/PowerShell/'); $ready = $true }
                finally { $wc.Dispose() }
            }
            catch [System.Net.WebException] {
                # 401 Unauthorized = IIS is up and the Exchange endpoint exists — credentials not needed to confirm readiness
                if ($_.Exception.Response -and ([int]$_.Exception.Response.StatusCode) -eq 401) { $ready = $true }
            }
            catch { }
        } while (-not $ready -and $elapsed -lt $maxWait)

        if (-not $ready) {
            Write-MyVerbose 'Exchange PowerShell endpoint did not become available within 90 s — retrying'
        }

        # After IIS restart, implicit-remoting proxy functions are removed automatically.
        # Import-ExchangeModule's guard (Get-ExchangeServer not found) will fire and reload.
        Import-ExchangeModule
        if (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue) {
            Write-MyVerbose 'Exchange PS session reconnected'
        }
        else {
            Write-MyWarning 'Failed to reconnect Exchange PS session — subsequent Exchange cmdlets may fail'
        }
    }

    function Install-EXpress_ {
        $ver = $State['MajorSetupVersion']
        Write-MyOutput "Installing Microsoft Exchange Server ($ver)"
        $PresenceKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{CD981244-E9B8-405A-9026-6AEB9DCEF1F1}'

        if (Get-ItemProperty -Path $PresenceKey -Name InstallDate -ErrorAction SilentlyContinue) {
            Write-MyOutput 'Exchange is already installed, skipping setup'
            return
        }

        if ( $State['Recover']) {
            Write-MyOutput 'Will run Setup in recover mode'
            $Params = '/mode:RecoverServer', $State['IAcceptSwitch'], '/DoNotStartTransport', '/InstallWindowsComponents'
            if ( $State['TargetPath']) {
                $Params += "/TargetDir:`"$($State['TargetPath'])`""
            }
        }
        else {
            if ( $State['Upgrade']) {
                Write-MyOutput 'Will run Setup in upgrade mode'
                $Params = '/mode:Upgrade', $State['IAcceptSwitch']
            }
            else {
                $roles = @()
                if ( $State['InstallEdge']) {
                    $roles = 'EdgeTransport'
                }
                else {
                    $roles = 'Mailbox'
                }
                $RolesParm = $roles -join ','
                if ([string]::IsNullOrEmpty( $RolesParm)) {
                    $RolesParm = 'Mailbox'
                }
                $Params = '/mode:install', "/roles:$RolesParm", $State['IAcceptSwitch'], '/DoNotStartTransport', '/InstallWindowsComponents'
                if ( $State['InstallMailbox']) {
                    if ( $State['InstallMDBName']) {
                        $Params += "/MdbName:$($State['InstallMDBName'])"
                    }
                    if ( $State['InstallMDBDBPath']) {
                        $Params += "/DBFilePath:`"$($State['InstallMDBDBPath'])\$($State['InstallMDBName']).edb`""
                    }
                    if ( $State['InstallMDBLogPath']) {
                        $Params += "/LogFolderPath:`"$($State['InstallMDBLogPath'])`""
                    }
                }
                if ( $State['TargetPath']) {
                    $Params += "/TargetDir:`"$($State['TargetPath'])`""
                }
                if ( $State['DoNotEnableEP']) {
                    $Params += "/DoNotEnableEP"
                }
                if ( $State['DoNotEnableEP_FEEWS']) {
                    $Params += "/DoNotEnableEP_FEEWS"
                }
            }
        }

        $res = Invoke-Process $State['SourcePath'] 'setup.exe' $Params
        if ( $res -ne 0 -or -not( Get-ItemProperty -Path $PresenceKey -Name InstallDate -ErrorAction SilentlyContinue)) {
            Write-MyError 'Exchange Setup exited with non-zero value or Install info missing from registry: Please consult the Exchange setup log, i.e. C:\ExchangeSetupLogs\ExchangeSetup.log'
            Invoke-SetupAssist
            exit $ERR_PROBLEMEXCHANGESETUP
        }
    }

    function Initialize-Exchange {
        # Returns $true if PrepareAD was executed, $false if already up-to-date (skip).
        if ($State['InstallEdge']) { return $false }

        $params = @()
        if ($State['MajorSetupVersion'] -ge $EX2019_MAJOR) {
            $MinFFL = $EX2019_MINFORESTLEVEL
            $MinDFL = $EX2019_MINDOMAINLEVEL
        }
        else {
            $MinFFL = $EX2016_MINFORESTLEVEL
            $MinDFL = $EX2016_MINDOMAINLEVEL
        }

        Write-MyOutput 'Checking whether Active Directory preparation is required'
        if ($null -ne (Test-ExchangeOrganization $State['OrganizationName'])) {
            Write-MyOutput "Exchange organization '$($State['OrganizationName'])' does not exist — PrepareAD required"
            $params += '/PrepareAD', "/OrganizationName:`"$($State['OrganizationName'])`""
            $State['NewExchangeOrg'] = $true   # org created by this run — Enable-AccessNamespaceMailConfig may run
            Save-State $State
        }
        else {
            $forestlvl = Get-ExchangeForestLevel
            $domainlvl = Get-ExchangeDomainLevel
            Write-MyOutput "Exchange Forest Schema: $forestlvl (min $MinFFL), Domain: $domainlvl (min $MinDFL)"
            if (($forestlvl -lt $MinFFL) -or ($domainlvl -lt $MinDFL)) {
                Write-MyOutput 'AD schema or domain level below minimum — PrepareAD required'
                $params += '/PrepareAD'
            }
            else {
                Write-MyOutput 'Active Directory is already prepared — skipping PrepareAD'
                return $false
            }
        }

        Write-MyOutput "Preparing Active Directory — Exchange organization: $($State['OrganizationName'])"
        $params += $State['IAcceptSwitch']
        $exitCode = Invoke-Process $State['SourcePath'] 'setup.exe' $params
        if ($exitCode -ne 0) {
            Write-MyError "Exchange setup /PrepareAD failed with exit code $exitCode. Please consult the Exchange setup log, i.e. C:\ExchangeSetupLogs\ExchangeSetup.log"
            exit $ERR_PROBLEMADPREPARE
        }
        if (($null -eq (Test-ExchangeOrganization $State['OrganizationName'])) -or
            ((Get-ExchangeForestLevel) -lt $MinFFL) -or
            ((Get-ExchangeDomainLevel) -lt $MinDFL)) {
            Write-MyError 'Problem updating schema, domain or Exchange organization. Please consult the Exchange setup log, i.e. C:\ExchangeSetupLogs\ExchangeSetup.log'
            exit $ERR_PROBLEMADPREPARE
        }
        return $true
    }

    function Install-WindowsFeatures( $MajorOSVersion) {
        Write-MyOutput 'Configuring Windows Features'

        if ( $State['InstallEdge']) {
            $Feats = [array]'ADLDS'
        }
        else {
            if ( [System.Version]$WS2019_PREFULL -ge [System.Version]$MajorOSVersion) {

                # WS2019, WS2022, WS2025
                $Feats = 'Server-Media-Foundation', 'NET-Framework-45-Core', 'NET-Framework-45-ASPNET',
                'NET-WCF-HTTP-Activation45', 'NET-WCF-Pipe-Activation45', 'NET-WCF-TCP-Activation45',
                'NET-WCF-TCP-PortSharing45', 'RPC-over-HTTP-proxy', 'RSAT-Clustering',
                'RSAT-Clustering-CmdInterface', 'RSAT-Clustering-PowerShell', 'WAS-Process-Model',
                'Web-Asp-Net45', 'Web-Basic-Auth', 'Web-Client-Auth', 'Web-Digest-Auth',
                'Web-Dir-Browsing', 'Web-Dyn-Compression', 'Web-Http-Errors', 'Web-Http-Logging',
                'Web-Http-Redirect', 'Web-Http-Tracing', 'Web-ISAPI-Ext', 'Web-ISAPI-Filter',
                'Web-Metabase', 'Web-Mgmt-Service', 'Web-Net-Ext45', 'Web-Request-Monitor',
                'Web-Server', 'Web-Stat-Compression', 'Web-Static-Content', 'Web-Windows-Auth',
                'Web-WMI', 'RSAT-ADDS'

                if ( !( Test-ServerCore)) {
                    $Feats += 'RSAT-Clustering-Mgmt', 'Web-Mgmt-Console', 'Windows-Identity-Foundation'
                }
            }
            else {
                # WS2016
                $Feats = 'NET-Framework-45-Core', 'NET-Framework-45-ASPNET', 'NET-WCF-HTTP-Activation45', 'NET-WCF-Pipe-Activation45', 'NET-WCF-TCP-Activation45', 'NET-WCF-TCP-PortSharing45', 'Server-Media-Foundation', 'RPC-over-HTTP-proxy', 'RSAT-Clustering', 'RSAT-Clustering-CmdInterface', 'RSAT-Clustering-Mgmt', 'RSAT-Clustering-PowerShell', 'WAS-Process-Model', 'Web-Asp-Net45', 'Web-Basic-Auth', 'Web-Client-Auth', 'Web-Digest-Auth', 'Web-Dir-Browsing', 'Web-Dyn-Compression', 'Web-Http-Errors', 'Web-Http-Logging', 'Web-Http-Redirect', 'Web-Http-Tracing', 'Web-ISAPI-Ext', 'Web-ISAPI-Filter', 'Web-Lgcy-Mgmt-Console', 'Web-Metabase', 'Web-Mgmt-Console', 'Web-Mgmt-Service', 'Web-Net-Ext45', 'Web-Request-Monitor', 'Web-Server', 'Web-Stat-Compression', 'Web-Static-Content', 'Web-Windows-Auth', 'Web-WMI', 'Windows-Identity-Foundation', 'RSAT-ADDS'
            }
        }
        $Feats += 'Bits'

        # Only query and install features that are not yet installed.
        # Get-WindowsFeature on all features at once is much faster than per-feature calls,
        # and skipping Install-WindowsFeature entirely avoids the slow "collecting data" phase
        # when all features are already present.
        Write-MyOutput ('Checking {0} required Windows features ...' -f $Feats.Count)
        $allFeatureState = Get-WindowsFeature -Name $Feats -ErrorAction SilentlyContinue
        $missing = @($allFeatureState | Where-Object { -not $_.Installed } | ForEach-Object { $_.Name })

        if ($missing.Count -eq 0) {
            Write-MyOutput 'All required Windows features already installed — skipping feature installation'
        }
        else {
            Write-MyOutput ('Installing {0} missing Windows feature(s): {1}' -f $missing.Count, ($missing -join ', '))
            Install-WindowsFeature $missing | Out-Null
        }

        foreach ( $Feat in $Feats) {
            if ( !( (Get-WindowsFeature -Name $Feat).Installed)) {
                Write-MyError "Feature $Feat appears not to be installed"
                exit $ERR_PROBLEMADDINGFEATURE
            }
        }

        'NET-WCF-MSMQ-Activation45', 'MSMQ' | ForEach-Object {
            if ( (Get-WindowsFeature -Name $_).Installed) {
                Write-MyOutput ('Removing obsolete feature {0}' -f $_)
                Remove-WindowsFeature -Name $_
            }
        }
    }

    function Test-MyPackage( $PackageID) {
        # Some packages are released using different GUIDs, specify more than 1 using '|'
        $PackageSet = $PackageID.split('|')
        $PresenceKey = $null
        foreach ( $ID in $PackageSet) {
            Write-MyVerbose "Checking if package $ID is installed .."
            $PresenceKey = (Get-CimInstance Win32_QuickFixEngineering | Where-Object { $_.HotfixID -eq $ID }).HotfixID
            if ( !( $PresenceKey)) {
                $PresenceKey = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$ID" -Name 'DisplayName' -ErrorAction SilentlyContinue).DisplayName
                if (!( $PresenceKey)) {
                    # Alternative (seen KB2803754, 2802063 register here)
                    $PresenceKey = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$ID" -Name 'DisplayName' -ErrorAction SilentlyContinue).DisplayName
                    if ( !( $PresenceKey)) {
                        # Alternative (eg Office2010FilterPack SP1)
                        $PresenceKey = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\$ID" -Name 'DisplayName' -ErrorAction SilentlyContinue).DisplayName
                        if ( !( $PresenceKey)) {
                            # Check for installed Exchange IUs
                            switch ( $State["MajorSetupVersion"]) {
                                $EX2016_MAJOR {
                                    $IUPath = 'Exchange 2016'
                                }
                                default {
                                    if ([System.Version]$State['SetupVersion'] -ge [System.Version]$EXSESETUPEXE_RTM) {
                                        $IUPath = 'Exchange SE'
                                    }
                                    else {
                                        $IUPath = 'Exchange 2019'
                                    }
                                }
                            }
                            $PresenceKey = (Get-ItemProperty -Path ('HKLM:\Software\Microsoft\Updates\{0}\{1}' -f $IUPath, $ID) -Name 'PackageName' -ErrorAction SilentlyContinue).PackageName
                        }
                    }
                }
            }
        }
        return $PresenceKey
    }

    function Install-MyPackage {
        param ( [String]$PackageID, [string]$Package, [String]$FileName, [String]$OnlineURL, [array]$Arguments, [switch]$NoDownload, [switch]$ContinueOnError)

        if ( $PackageID) {
            Write-MyOutput "Processing $Package ($PackageID)"
            $PresenceKey = Test-MyPackage $PackageID
        }
        else {
            # Just install, don't detect
            Write-MyOutput "Processing $Package"
            $PresenceKey = $false
        }
        # All downloads land in <InstallPath>\sources\; falls back to InstallPath if
        # SourcesPath wasn't initialized yet (safety guard — Install-MyPackage may be
        # called before the dedicated sources folder is set up).
        $RunFrom = if ($State['SourcesPath']) { $State['SourcesPath'] } else { $State['InstallPath'] }
        if ( !( $PresenceKey )) {

            if ( $FileName.contains('|')) {
                # Filename contains filename (dl) and package name (after extraction)
                $PackageFile = ($FileName.Split('|'))[1]
                $FileName = ($FileName.Split('|'))[0]
                if ( !( Get-MyPackage $Package '' $FileName $RunFrom)) {
                    # Download & Extract
                    if ( !( Get-MyPackage $Package $OnlineURL $PackageFile $RunFrom)) {
                        if ($ContinueOnError) { Write-MyWarning "Could not download $Package — skipping"; return } else { Write-MyError "Problem downloading/accessing $Package"; exit $ERR_PROBLEMPACKAGEDL }
                    }
                    Write-MyOutput "Extracting Hotfix Package $Package"
                    Invoke-Extract $RunFrom $PackageFile

                    if ( !( Get-MyPackage $Package $OnlineURL $PackageFile $RunFrom)) {
                        if ($ContinueOnError) { Write-MyWarning "Could not download $Package — skipping"; return } else { Write-MyError "Problem downloading/accessing $Package"; exit $ERR_PROBLEMPACKAGEEXTRACT }
                    }
                }
            }
            else {
                if ( $NoDownload) {
                    $RunFrom = Split-Path -Path $OnlineURL -Parent
                    Write-MyVerbose "Will run $FileName straight from $RunFrom"
                }
                if ( !( Get-MyPackage $Package $OnlineURL $FileName $RunFrom)) {
                    if ($ContinueOnError) { Write-MyWarning "Could not download $Package — skipping"; return } else { Write-MyError "Problem downloading/accessing $Package"; exit $ERR_PROBLEMPACKAGEDL }
                }
            }

            Write-MyOutput "Installing $Package from $RunFrom"
            $rval = Invoke-Process $RunFrom $FileName $Arguments

            if ( $PackageID) {
                $PresenceKey = Test-MyPackage $PackageID
            }
            else {
                # Don't check post-installation
                $PresenceKey = $true
            }
            if ( ( @(3010, $ERR_SUS_NOT_APPLICABLE) -contains $rval) -or $PresenceKey) {
                switch ( $rval) {
                    3010 {
                        Write-MyVerbose "Installation $Package successful, reboot required"
                    }
                    $ERR_SUS_NOT_APPLICABLE {
                        Write-MyVerbose "$Package not applicable or blocked - ignoring"
                    }
                    default {
                        Write-MyVerbose "Installation $Package successful"
                    }
                }
            }
            else {
                if ($ContinueOnError) { Write-MyWarning "Could not install $Package — skipping"; return } else { Write-MyError "Problem installing $Package - For fixes, check $($ENV:WINDIR)\WindowsUpdate.log; For .NET Framework issues, check 'Microsoft .NET Framework 4 Setup' HTML document in $($ENV:TEMP)"; exit $ERR_PROBLEMPACKAGESETUP }
            }
        }
        else {
            Write-MyVerbose "$Package already installed"
        }
    }


    function Get-FFLText( $FFL = 0) {
        $FFLlevels = @{
            0           = 'Unknown or unsupported'
            $FFL_2003   = '2003'
            $FFL_2008   = '2008'
            $FFL_2008R2 = '2008R2'
            $FFL_2012   = '2012'
            $FFL_2012R2 = '2012R2'
            $FFL_2016   = '2016'
            $FFL_2025   = '2025'
        }
        return ($FFLlevels.GetEnumerator() | Where-Object { $FFL -ge $_.Name } | Sort-Object Name -Descending | Select-Object -First 1).Value
    }

    function Get-NetVersionText( $NetVersion = 0) {
        $NETversions = @{
            0               = 'Unknown or unsupported'
            $NETVERSION_48  = '4.8'
            $NETVERSION_481 = '4.8.1'
        }
        return ($NetVersions.GetEnumerator() | Where-Object { $NetVersion -ge $_.Name } | Sort-Object Name -Descending | Select-Object -First 1).Value
    }

    function Get-NETVersion {
        $NetVersion = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue).Release
        return [int]$NetVersion
    }

    function Set-NETFrameworkInstallBlock {
        param ( [String]$Version, [String]$KB, [string]$Key)
        $RegKey = 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\WU'
        $RegName = ('BlockNetFramework{0}' -f $Key)
        Write-MyOutput ('Set installation blockade for .NET Framework {0} ({1})' -f $Version, $KB)
        Set-RegistryValue -Path $RegKey -Name $RegName -Value 1
        if (-not (Get-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue)) {
            Write-MyError "Unable to set registry entry $RegKey\$RegName"
        }
    }

    function Remove-NETFrameworkInstallBlock {
        param ( [String]$Version, [String]$KB, [string]$Key)
        $RegKey = 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\WU'
        $RegName = ('BlockNetFramework{0}' -f $Key)
        if ( Get-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue) {
            Write-MyOutput ('Remove installation blockade for .NET Framework {0} ({1})' -f $Version, $KB)
            Remove-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue | Out-Null
        }
        if ( Get-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue) {
            Write-MyError "Unable to remove registry entry $RegKey\$RegName"
        }
    }

    function Test-Preflight {
        # Informational checks only on first run (phase 0/1); on resume these were already validated
        if ($State['InstallPhase'] -le 1) {
            Write-MyOutput 'Performing preflight checks'

            $Computer = Get-LocalFQDNHostname
            if ( $Computer) {
                Write-MyOutput "Computer name is $Computer"
            }

            Write-MyOutput 'Checking temporary installation folder'
            New-Item -Path $State['InstallPath'] -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
            if ( !( Test-Path $State['InstallPath'])) {
                Write-MyError "Can't create temporary folder $($State['InstallPath'])"
                exit $ERR_CANTCREATETEMPFOLDER
            }

            # Downloads cache: all prerequisite packages (.NET, VC++, UCMA, URL Rewrite, hotfixes,
            # Exchange SUs) and CSS-Exchange scripts (HealthChecker, EOMT, SetupAssist,
            # ExchangeExtendedProtection, MEAC, etc.) land here. Pre-staging is automatic —
            # when a file is already present the download is skipped (air-gapped / proxy scenarios).
            $State['SourcesPath'] = Join-Path $State['InstallPath'] 'sources'
            New-Item -Path $State['SourcesPath'] -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
            if ( !( Test-Path $State['SourcesPath'])) {
                Write-MyError "Can't create downloads folder $($State['SourcesPath'])"
                exit $ERR_CANTCREATETEMPFOLDER
            }
            Write-MyVerbose ('Downloads cache: {0}' -f $State['SourcesPath'])

            if ( [System.Version]$MajorOSVersion -ge [System.Version]$WS2016_MAJOR ) {
                Write-MyOutput "Operating System is $($MajorOSVersion).$($MinorOSVersion)"
            }
            else {
                Write-MyError 'Supported operating systems: Windows Server 2016 (Exchange 2016 CU23), Windows Server 2019/2022/2025 (Exchange 2019 CU15+ or Exchange Server SE)'
                exit $ERR_UNEXPECTEDOS
            }
            Write-MyOutput ('Server core mode: {0}' -f (Test-ServerCore))

            $NetVersion = Get-NETVersion
            $NetVersionText = Get-NetVersionText $NetVersion
            Write-MyOutput ".NET Framework is $NetVersion ($NetVersionText)"

            # Warn about parameters that are ignored for the selected install mode
            if ($State['InstallEdge']) {
                if ($State['DAGName'])          { Write-MyWarning 'DAGName is ignored for Edge Transport installations' }
                if ($State['Namespace'])        { Write-MyWarning 'Namespace is ignored for Edge Transport installations' }
                if ($State['CopyServerConfig']) { Write-MyWarning 'CopyServerConfig is ignored for Edge Transport installations' }
            }
        }

        if (! ( Test-Admin)) {
            Write-MyWarning 'Script not running in elevated mode, attempting auto-elevation ..'
            try {
                $scriptPath = $MyInvocation.ScriptName
                if (-not $scriptPath) { $scriptPath = $PSCommandPath }
                $argList = "-NoProfile -ExecutionPolicy Unrestricted -File `"$scriptPath`""
                # Re-pass bound parameters
                foreach ($param in $PSBoundParameters.GetEnumerator()) {
                    if ($param.Value -is [switch]) {
                        if ($param.Value) { $argList += " -$($param.Key)" }
                    }
                    elseif ($param.Value -is [System.Management.Automation.PSCredential]) {
                        # Credentials cannot be passed via command line, skip
                        Write-MyWarning 'Credentials parameter cannot be passed during auto-elevation, you will be prompted'
                    }
                    else {
                        $argList += " -$($param.Key) `"$($param.Value)`""
                    }
                }
                Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList $argList -Verb RunAs
                exit $ERR_OK
            }
            catch {
                Write-MyError ('Auto-elevation failed: {0}' -f $_.Exception.Message)
                exit $ERR_RUNNINGNONADMINMODE
            }
        }
        else {
            Write-MyOutput 'Script running in elevated mode'
        }

        # Credential validation only needed while Exchange setup is running (phases 0-4)
        if ($State['InstallPhase'] -le 4 -and $State['Autopilot']) {
            $credentialsFromCommandLine = $PSBoundParameters.ContainsKey('Credentials')
            if ( -not( $State['AdminAccount'] -and $State['AdminPassword'])) {
                # No credentials in state yet — prompt interactively if possible, else fail
                if ([Environment]::UserInteractive -and -not $credentialsFromCommandLine) {
                    if (-not (Get-ValidatedCredentials)) {
                        exit $ERR_NOACCOUNTSPECIFIED
                    }
                }
                else {
                    Write-MyError 'Autopilot specified but no credentials provided'
                    exit $ERR_NOACCOUNTSPECIFIED
                }
            }
            else {
                # Credentials already in state (command line, config file, or Autopilot resume)
                Write-MyOutput 'Checking provided credentials'
                if (Test-Credentials) {
                    Write-MyOutput 'Credentials valid'
                }
                elseif ([Environment]::UserInteractive -and -not $credentialsFromCommandLine) {
                    # Stored credentials invalid (e.g. password changed since last phase) — retry interactively
                    Write-MyWarning 'Stored credentials are no longer valid, prompting for new credentials'
                    if (-not (Get-ValidatedCredentials)) {
                        exit $ERR_INVALIDCREDENTIALS
                    }
                }
                else {
                    Write-MyError "Provided credentials don't seem to be valid"
                    exit $ERR_INVALIDCREDENTIALS
                }
            }

        }

        # Checks below are only relevant before/during setup (phases 0-4); skip after Exchange is installed
        if ($State['InstallPhase'] -le 4) {

        if ( $State["SkipRolesCheck"] -or $State['InstallEdge']) {
            Write-MyOutput 'SkipRolesCheck: Skipping validation of Schema & Enterprise Administrators membership'
        }
        else {
            if (! ( Test-SchemaAdmin)) {
                Write-MyError 'Current user is not member of Schema Administrators'
                exit $ERR_RUNNINGNONSCHEMAADMIN
            }
            else {
                Write-MyOutput 'User is member of Schema Administrators'
            }

            if (! ( Test-EnterpriseAdmin)) {
                Write-MyError 'User is not member of Enterprise Administrators'
                exit $ERR_RUNNINGNONENTERPRISEADMIN
            }
            else {
                Write-MyOutput 'User is member of Enterprise Administrators'
            }
        }
        if (!$State['InstallEdge']) {
            $ADSite = Get-ADSite
            if ( $ADSite) {
                Write-MyOutput "Computer is located in AD site $ADSite"
            }
            else {
                Write-MyError 'Could not determine Active Directory site'
                exit $ERR_COULDNOTDETERMINEADSITE
            }

            $ExOrg = Get-ExchangeOrganization
            if ( $ExOrg) {
                if ( $State['OrganizationName']) {
                    if ( $State['OrganizationName'] -ne $ExOrg) {
                        Write-MyError "OrganizationName mismatches with discovered Exchange Organization name ($ExOrg, expected $($State['OrganizationName']))"
                        exit $ERR_ORGANIZATIONNAMEMISMATCH
                    }
                }
                Write-MyOutput "Exchange Organization is: $ExOrg"
            }
            else {
                if ( $State['OrganizationName']) {
                    Write-MyOutput "Exchange Organization will be: $($State['OrganizationName'])"
                }
                else {
                    Write-MyError 'OrganizationName not specified and no Exchange Organization discovered'
                    exit $ERR_MISSINGORGANIZATIONNAME
                }
            }
        }
        Write-MyOutput 'Checking if we can access Exchange setup ..'

        if (! (Test-Path (Join-Path $State['SourcePath'] "setup.exe"))) {
            Write-MyError "Can't find Exchange setup at $($State['SourcePath'])"
            exit $ERR_MISSINGEXCHANGESETUP
        }
        else {
            Write-MyOutput "Exchange setup located at $(Join-Path $($State['SourcePath']) "setup.exe")"
        }

        # Unblock files to prevent .NET assembly sandboxing errors (Zone.Identifier from downloaded files).
        # Skip when source is a mounted ISO: UDF/ISO9660 does not support Alternate Data Streams, and
        # the ISO itself was already unblocked before mounting (see above). Querying ADS on UDF throws
        # a terminating Win32Exception ("The parameter is incorrect") that -ErrorAction cannot suppress.
        if (-not $State['SourceImage']) {
            $blockedFiles = Get-ChildItem -Path $State['SourcePath'] -Recurse -File | Where-Object {
                try { $null -ne (Get-Item -Path $_.FullName -Stream 'Zone.Identifier' -ErrorAction SilentlyContinue) }
                catch { $false }
            }
            if ($blockedFiles) {
                Write-MyWarning ('{0} blocked file(s) detected in source path, unblocking ..' -f $blockedFiles.Count)
                $blockedFiles | Unblock-File
                Write-MyOutput 'Source files unblocked successfully'
            }
        }

        $State['ExSetupVersion'] = Get-DetectedFileVersion "$($State['SourcePath'])\Setup\ServerRoles\Common\ExSetup.exe"
        $SetupVersion = $State['ExSetupVersion']
        $State['SetupVersionText'] = Get-SetupTextVersion $SetupVersion
        Write-MyOutput ('ExSetup version: {0}' -f $State['SetupVersionText'])
        if ( $SetupVersion) {
            $Num = $SetupVersion.split('.') | ForEach-Object { [string]([int]$_)
            }
            $MajorSetupVersion = [decimal]($num[0] + '.' + $num[1])
            $MinorSetupVersion = [decimal]($num[2] + '.' + $num[3])
        }
        else {
            $MajorSetupVersion = 0
            $MinorSetupVersion = 0
        }
        $State['MajorSetupVersion'] = $MajorSetupVersion
        $State['MinorSetupVersion'] = $MinorSetupVersion

        # Target install supports only the latest CU of each Exchange line:
        # Ex2016 CU23 (final), Ex2019 CU15+, Exchange SE RTM+.
        # Older Ex2019 CUs (CU10–CU14) are out of Microsoft SU support and rejected here.
        # Note: Export-SourceServerConfig queries remote source servers independently and still
        # accepts older CUs as migration sources — this gate only governs the local install target.
        if ( ($MajorSetupVersion -eq $EX2019_MAJOR -and [System.Version]$SetupVersion -lt [System.Version]$EX2019SETUPEXE_CU15) -or
            ($MajorSetupVersion -eq $EX2016_MAJOR -and [System.Version]$SetupVersion -lt [System.Version]$EX2016SETUPEXE_CU23) ) {
            Write-MyError 'Unsupported Exchange target version. Supported install targets: Exchange 2016 CU23 (final), Exchange 2019 CU15+, or Exchange Server SE. Older Exchange 2019 CUs (CU10–CU14) are out of Microsoft SU support — please install CU15 or Exchange Server SE.'
            exit $ERR_UNSUPPORTEDEX
        }

        if ( -not $State['InstallEdge'] -and [System.Version]$SetupVersion -ge [System.Version]$EX2019SETUPEXE_CU15) {
            $Ex2013Exists = Get-ExchangeServerObjects | Where-Object { $_.serialNumber -and $_.serialNumber[0] -like 'Version 15.0*' }
            if ( $Ex2013Exists) {
                Write-MyError ('Exchange 2013 detected: {0}. Exchange 2019 CU15 or later cannot co-exist with Exchange 2013' -f ($Ex2013Exists | Select-Object Name) -join ',')
                exit $ERR_EX19EX2013COEXIST
            }
        }

        # Exchange SE coexistence: SE RTM/CU1 supports EX2016 CU23 and EX2019 CU14+, but SE CU2+ does not
        if ( [System.Version]$SetupVersion -ge [System.Version]$EXSESETUPEXE_RTM) {
            $Ex2016Exists = Get-ExchangeServerObjects | Where-Object { $_.serialNumber[0] -like 'Version 15.1*' }
            if ( $Ex2016Exists) {
                Write-MyWarning ('Exchange 2016 server(s) detected: {0}. Exchange SE RTM/CU1 supports coexistence with Exchange 2016 CU23, but SE CU2+ will not. Plan decommissioning.' -f (($Ex2016Exists | Select-Object -ExpandProperty Name) -join ', '))
            }
        }

        # OS gate for Exchange Server SE — supported on WS2019, WS2022, WS2025
        if ( [System.Version]$SetupVersion -ge [System.Version]$EXSESETUPEXE_RTM -and [System.Version]$FullOSVersion -lt $WS2019_PREFULL) {
            Write-MyError 'Exchange Server SE requires Windows Server 2019, Windows Server 2022 or Windows Server 2025'
            exit $ERR_UNEXPECTEDOS
        }

        # OS gate for Exchange 2016 CU23 — target only on WS2016 per Microsoft supportability matrix
        # (WS2012/R2 are past extended support; WS2019+ are not supported by Exchange 2016 setup).
        if ( $MajorSetupVersion -eq $EX2016_MAJOR ) {
            if ( [System.Version]$FullOSVersion -lt [System.Version]$WS2016_MAJOR ) {
                Write-MyError 'Exchange 2016 CU23 requires Windows Server 2016'
                exit $ERR_UNEXPECTEDOS
            }
            if ( [System.Version]$FullOSVersion -ge $WS2019_PREFULL ) {
                Write-MyError 'Exchange 2016 CU23 is only supported on Windows Server 2016. For newer Windows Server releases install Exchange 2019 CU15+ or Exchange Server SE.'
                exit $ERR_UNEXPECTEDOS
            }
        }

        # OS gate for Exchange 2019 CU15+ — supported on WS2019, WS2022, WS2025
        if ( $MajorSetupVersion -eq $EX2019_MAJOR -and [System.Version]$FullOSVersion -lt $WS2019_PREFULL ) {
            Write-MyError 'Exchange 2019 CU15+ requires Windows Server 2019, Windows Server 2022 or Windows Server 2025'
            exit $ERR_UNEXPECTEDOS
        }

        if ( $State['NoSetup'] -or $State['Recover'] -or $State['Upgrade']) {
            Write-MyOutput 'Not checking roles (NoSetup, Recover or Upgrade mode)'
        }
        else {
            Write-MyOutput 'Checking roles to install'
            if ( !( $State['InstallMailbox']) -and !($State['InstallEdge']) ) {
                Write-MyError 'No roles specified to install'
                exit $ERR_UNKNOWNROLESSPECIFIED
            }
        }

        # Ex2019 CU15+ and Exchange SE always support DiagnosticData switch.
        # Ex2016 CU23 uses the legacy non-DiagnosticData license-accept switch.
        if ( $State["MajorSetupVersion"] -eq $EX2019_MAJOR ) {
            if ( $State['DiagnosticData']) {
                $State['IAcceptSwitch'] = '/IAcceptExchangeServerLicenseTerms_DiagnosticDataON'
                Write-MyOutput 'Will deploy Exchange with Data Collection enabled'
            }
            else {
                $State['IAcceptSwitch'] = '/IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF'
            }
        }
        else {
            $State['IAcceptSwitch'] = '/IAcceptExchangeServerLicenseTerms'
        }

        if ( !($State['InstallEdge'])) {
            if ( ( Test-ExistingExchangeServer $env:computerName) -and ($State["InstallPhase"] -eq 1)) {
                if ( $State['Recover']) {
                    Write-MyOutput 'Recovery mode specified, Exchange server object found'
                }
                else {
                    if ( Test-Path $EXCHANGEINSTALLKEY) {
                        Write-MyOutput 'Existing Exchange server object found in Active Directory, and installation seems present - switching to Upgrade mode'
                        $State['Upgrade'] = $true
                    }
                    else {
                        Write-MyError 'Existing Exchange server object found in Active Directory, but installation missing - please use Recover switch to recover a server'
                        exit $ERR_PROBLEMEXCHANGESERVEREXISTS
                    }
                }
            }

            Write-MyOutput 'Checking domain membership status ..'
            if (! ( Get-CimInstance Win32_ComputerSystem).PartOfDomain) {
                Write-MyError 'System is not domain-joined'
                exit $ERR_NOTDOMAINJOINED
            }
        }
        Write-MyOutput 'Checking NIC configuration ..'
        if (! (Get-CimInstance Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True and DHCPEnabled=False')) {
            $AzureHosted = Get-Service | Where-Object { $_.Name -ieq 'Windows Azure Guest Agent' -or $_.Name -ieq 'Windows Azure Network Agent' -or $_.Name -ieq 'Windows Azure Telemetry Service' }
            if ( $AzureHosted) {
                Write-MyError "System doesn't have a static IP addresses configured"
                exit $ERR_NOFIXEDIPADDRESS
            }
            else {
                Write-MyOutput 'Ignoring absence of static IP address assignment(s) as Azure service(s) are present.'
            }
        }
        else {
            Write-MyVerbose 'Static IP address(es) assigned.'
        }
        if ( $State['TargetPath']) {
            $Location = Split-Path $State['TargetPath'] -Qualifier
            Write-MyOutput 'Checking installation path ..'
            if ( !(Test-Path $Location)) {
                Write-MyError "MDB log location unavailable: ($Location)"
                exit $ERR_MDBDBLOGPATH
            }
        }
        if ( $State['InstallMDBLogPath']) {
            $Location = Split-Path $State['InstallMDBLogPath'] -Qualifier
            Write-MyOutput 'Checking MDB log path ..'
            if ( !(Test-Path $Location)) {
                Write-MyError "MDB log location unavailable: ($Location)"
                exit $ERR_MDBDBLOGPATH
            }
        }
        if ( $State['InstallMDBDBPath']) {
            $Location = Split-Path $State['InstallMDBDBPath'] -Qualifier
            Write-MyOutput 'Checking MDB database path ..'
            if ( !(Test-Path $Location)) {
                Write-MyError "MDB database location unavailable: ($Location)"
                exit $ERR_MDBDBLOGPATH
            }
        }
        if ( !($State['InstallEdge'])) {
            Write-MyOutput 'Checking Exchange Forest Schema Version'
            if ( $State['MajorSetupVersion'] -ge $EX2019_MAJOR) {
                $minFFL = $EX2019_MINFORESTLEVEL
                $minDFL = $EX2019_MINDOMAINLEVEL
            }
            else {
                $minFFL = $EX2016_MINFORESTLEVEL
                $minDFL = $EX2016_MINDOMAINLEVEL
            }
            $EFL = Get-ExchangeForestLevel
            if ( $EFL) {
                Write-MyOutput "Exchange Forest Schema Version is $EFL"
            }
            else {
                Write-MyOutput 'Active Directory is not prepared'
            }
            if ( $State['InstallPhase'] -ge 4) {
                if ( $null -eq $EFL -or $EFL -lt $minFFL) {
                    if ( $null -eq $EFL) {
                        Write-MyWarning 'Active Directory is not prepared. PrepareAD may have failed in a previous phase.'
                    }
                    else {
                        Write-MyWarning "Exchange Forest Schema version is $EFL (required: $minFFL)"
                    }
                    Write-MyWarning 'Rolling back to phase 3 to retry AD preparation ..'
                    $State['InstallPhase'] = 3
                    $State['LastSuccessfulPhase'] = 2
                }
            }

            Write-MyOutput 'Checking Exchange Domain Version'
            $EDV = Get-ExchangeDomainLevel
            if ( $EDV) {
                Write-MyOutput "Exchange Domain Version is $EDV"
            }
            if ( $State['InstallPhase'] -ge 4) {
                if ( $null -eq $EDV -or $EDV -lt $minDFL) {
                    if ( $null -eq $EDV) {
                        Write-MyWarning 'Exchange Domain is not prepared. PrepareAD may have failed in a previous phase.'
                    }
                    else {
                        Write-MyWarning "Exchange Domain version is $EDV (required: $minDFL)"
                    }
                    if ( $State['InstallPhase'] -ne 3) {
                        Write-MyWarning 'Rolling back to phase 3 to retry AD preparation ..'
                        $State['InstallPhase'] = 3
                        $State['LastSuccessfulPhase'] = 2
                    }
                }
                if ( $EDV -lt $minDFL) {
                    Write-MyError "Minimum required Exchange Domain version is $minDFL (current: $EDV), aborting"
                    exit $ERR_BADDOMAINLEVEL
                }
            }

            Write-MyOutput 'Checking domain mode'
            if ( Test-DomainNativeMode -eq $DOMAIN_MIXEDMODE) {
                Write-MyError 'Domain is in mixed mode, native mode is required'
                exit $ERR_ADMIXEDMODE
            }
            else {
                Write-MyOutput 'Domain is in native mode'
            }

            Write-MyOutput 'Checking Forest Functional Level'
            $FFL = Get-ForestFunctionalLevel
            if ( $MajorSetupVersion -eq $EX2019_MAJOR) {
                if ( $FFL -lt $FOREST_LEVEL2012R2) {
                    Write-MyError ('Exchange Server 2019/SE requires Forest Functionality Level 2012R2 ({0}).' -f $FFL)
                    exit $ERR_ADFORESTLEVEL
                }
                else {
                    Write-MyOutput ('Forest Functional Level is {0} ({1})' -f $FFL, (Get-FFLText $FFL))
                }
            }
            else {
                if ( $FFL -lt $FOREST_LEVEL2012) {
                    Write-MyError ('Exchange Server 2016 or later requires Forest Functionality Level 2012 ({0}).' -f $FFL)
                    exit $ERR_ADFORESTLEVEL
                }
                else {
                    Write-MyOutput ('Forest Functional Level is OK ({0})' -f $FFL)
                }
            }
        }
        if ( Get-PSExecutionPolicy) {
            # Referring to http://support.microsoft.com/kb/2810617/en
            Write-MyWarning 'PowerShell Execution Policy is configured through GPO and may prohibit Exchange Setup. Clearing entry.'
            Remove-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell -Name ExecutionPolicy -Force
        }

        } # end if ($State['InstallPhase'] -le 4)
    }

    function New-PreflightReport {
        Write-MyOutput 'Generating Pre-Flight Validation Report'
        $results = @()

        # OS Version
        $results += [PSCustomObject]@{ Check = 'Operating System'; Result = $FullOSVersion; Status = if ([System.Version]$MajorOSVersion -ge [System.Version]$WS2016_MAJOR) { 'OK' } else { 'FAIL' } }

        # Admin check
        $results += [PSCustomObject]@{ Check = 'Running as Administrator'; Result = (Test-Admin); Status = if (Test-Admin) { 'OK' } else { 'FAIL' } }

        # Domain membership
        $isDomainJoined = (Get-CimInstance Win32_ComputerSystem).PartOfDomain
        $results += [PSCustomObject]@{ Check = 'Domain Membership'; Result = $isDomainJoined; Status = if ($isDomainJoined -or $State['InstallEdge']) { 'OK' } else { 'FAIL' } }

        # Computer name
        $computerName = try { Get-LocalFQDNHostname } catch { $env:COMPUTERNAME }
        $results += [PSCustomObject]@{ Check = 'Computer Name (FQDN)'; Result = $computerName; Status = 'INFO' }

        # Static IP
        $staticIP = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True and DHCPEnabled=False'
        $results += [PSCustomObject]@{ Check = 'Static IP Address'; Result = if ($staticIP) { ($staticIP.IPAddress -join ', ') } else { 'DHCP only' }; Status = if ($staticIP) { 'OK' } else { 'WARN' } }

        # .NET Framework
        $netVer = Get-NETVersion
        $results += [PSCustomObject]@{ Check = '.NET Framework'; Result = ('{0} ({1})' -f $netVer, (Get-NetVersionText $netVer)); Status = if ($netVer -ge $NETVERSION_48) { 'OK' } else { 'WARN' } }

        # Reboot pending
        $rebootPending = Test-RebootPending
        $results += [PSCustomObject]@{ Check = 'Reboot Pending'; Result = $rebootPending; Status = if ($rebootPending) { 'WARN' } else { 'OK' } }

        # Exchange setup
        if ($State['SourcePath'] -and (Test-Path (Join-Path $State['SourcePath'] 'setup.exe'))) {
            $exVer = Get-DetectedFileVersion (Join-Path $State['SourcePath'] 'Setup\ServerRoles\Common\ExSetup.exe')
            $results += [PSCustomObject]@{ Check = 'Exchange Setup Version'; Result = $exVer; Status = 'OK' }
        }
        else {
            $results += [PSCustomObject]@{ Check = 'Exchange Setup'; Result = $State['SourcePath']; Status = 'FAIL' }
        }

        # AD checks (non-Edge only)
        if (-not $State['InstallEdge']) {
            $adSite = try { Get-ADSite } catch { $null }
            $results += [PSCustomObject]@{ Check = 'AD Site'; Result = if ($adSite) { $adSite.ToString() } else { 'Not detected' }; Status = if ($adSite) { 'OK' } else { 'FAIL' } }

            if (-not $State['SkipRolesCheck']) {
                $isSchemaAdmin = try { Test-SchemaAdmin } catch { $false }
                $isEnterpriseAdmin = try { Test-EnterpriseAdmin } catch { $false }
                $results += [PSCustomObject]@{ Check = 'Schema Admin'; Result = [bool]$isSchemaAdmin; Status = if ($isSchemaAdmin) { 'OK' } else { 'FAIL' } }
                $results += [PSCustomObject]@{ Check = 'Enterprise Admin'; Result = [bool]$isEnterpriseAdmin; Status = if ($isEnterpriseAdmin) { 'OK' } else { 'FAIL' } }
            }

            $ffl = try { Get-ForestFunctionalLevel } catch { 0 }
            $results += [PSCustomObject]@{ Check = 'Forest Functional Level'; Result = ('{0} ({1})' -f $ffl, (Get-FFLText $ffl)); Status = if ($ffl -ge $FOREST_LEVEL2012R2) { 'OK' } else { 'WARN' } }

            $exOrg = try { Get-ExchangeOrganization } catch { $null }
            $results += [PSCustomObject]@{ Check = 'Exchange Organization'; Result = if ($exOrg) { $exOrg } else { $State['OrganizationName'] }; Status = 'INFO' }
        }

        # Disk allocation unit sizes
        Get-Volume | Where-Object { $_.DriveLetter -and $_.FileSystem -eq 'NTFS' } | ForEach-Object {
            $auOk = ($_.AllocationUnitSize -eq 65536 -or -not $_.AllocationUnitSize)
            $results += [PSCustomObject]@{ Check = ('Drive {0}: Allocation Unit' -f $_.DriveLetter); Result = ('{0} bytes' -f $_.AllocationUnitSize); Status = if ($auOk) { 'OK' } else { 'WARN' } }
        }

        # Server Core
        $isCore = Test-ServerCore
        $results += [PSCustomObject]@{ Check = 'Server Core'; Result = $isCore; Status = 'INFO' }

        # Source server connectivity (if CopyServerConfig specified)
        if ($State['CopyServerConfig']) {
            $sourceReachable = Test-Connection -ComputerName $State['CopyServerConfig'] -Count 1 -Quiet -ErrorAction SilentlyContinue
            $results += [PSCustomObject]@{ Check = ('Source Server {0} Reachable' -f $State['CopyServerConfig']); Result = $sourceReachable; Status = if ($sourceReachable) { 'OK' } else { 'FAIL' } }
        }

        # Generate HTML report
        $reportPath = Join-Path $State['ReportsPath'] ('{0}_EXpress_Preflight_{1}.html' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd-HHmmss'))
        $failCount = ($results | Where-Object { $_.Status -eq 'FAIL' }).Count
        $warnCount = ($results | Where-Object { $_.Status -eq 'WARN' }).Count
        $statusColor = if ($failCount -gt 0) { '#dc3545' } elseif ($warnCount -gt 0) { '#ffc107' } else { '#28a745' }

        $htmlRows = $results | ForEach-Object {
            $color = switch ($_.Status) { 'OK' { '#d4edda' } 'FAIL' { '#f8d7da' } 'WARN' { '#fff3cd' } default { '#d1ecf1' } }
            '<tr style="background-color:{0}"><td>{1}</td><td>{2}</td><td><strong>{3}</strong></td></tr>' -f $color, $_.Check, $_.Result, $_.Status
        }

        $html = @"
<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Exchange Pre-Flight Report</title>
<style>body{font-family:Segoe UI,sans-serif;margin:20px}table{border-collapse:collapse;width:100%}
th,td{padding:8px 12px;border:1px solid #ddd;text-align:left}th{background:#343a40;color:#fff}
h1{color:#333}.summary{padding:10px;color:#fff;border-radius:4px;margin-bottom:20px}</style></head>
<body><h1>Exchange Server Pre-Flight Validation Report</h1>
<div class="summary" style="background-color:$statusColor">
<strong>Computer:</strong> $env:COMPUTERNAME | <strong>Date:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') |
<strong>Failures:</strong> $failCount | <strong>Warnings:</strong> $warnCount</div>
<table><tr><th>Check</th><th>Result</th><th>Status</th></tr>
$($htmlRows -join "`n")
</table>
<h2 style="margin-top:30px;color:#333">Exchange Database Sizing Best Practices</h2>
<table>
<tr><th>Scenario</th><th>Recommended Max DB Size</th><th>Notes</th></tr>
<tr style="background-color:#d4edda"><td>DAG (≥2 copies)</td><td>2 TB</td><td>Each database copy on a separate volume</td></tr>
<tr style="background-color:#fff3cd"><td>Standalone (no DAG)</td><td>200 GB</td><td>Limited recovery options without DAG</td></tr>
<tr style="background-color:#f8d7da"><td>Lagged DAG copy</td><td>200 GB</td><td>Replay lag reduces effective copy count</td></tr>
</table>
<ul style="margin-top:12px;font-family:Segoe UI,sans-serif">
<li>Separate database (.edb) and transaction log volumes — different spindles or SSDs</li>
<li>Use 64 KB NTFS allocation unit size on all Exchange volumes</li>
<li>Reserve ≥20% free space on database volumes at all times</li>
<li>One mailbox database per volume is strongly recommended</li>
</ul>
</body></html>
"@
        $html | Out-File $reportPath -Encoding utf8
        Write-MyOutput ('Pre-Flight Report saved to {0}' -f $reportPath)
        return $failCount
    }

    function Export-SourceServerConfig {
        param([string]$SourceServer)
        Write-MyOutput ('Exporting configuration from source server {0}' -f $SourceServer)
        $configPath = Join-Path $State['InstallPath'] ('{0}_EXpress_Config.xml' -f $SourceServer)

        try {
            $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ('http://{0}/PowerShell/' -f $SourceServer) -Authentication Kerberos -ErrorAction Stop
            Write-MyVerbose ('Connected to {0} via Remote PowerShell' -f $SourceServer)
        }
        catch {
            Write-MyError ('Failed to connect to source server {0}: {1}' -f $SourceServer, $_.Exception.Message)
            exit $ERR_SOURCESERVERCONNECT
        }

        $config = @{}
        try {
            Write-MyVerbose 'Exporting Receive Connectors'
            $config['ReceiveConnectors'] = Invoke-Command -Session $session -ScriptBlock {
                Get-ReceiveConnector -Server $using:SourceServer | Select-Object Name, Bindings, RemoteIPRanges, PermissionGroups, AuthMechanism, Enabled, TransportRole, Fqdn, Banner, MaxMessageSize, MaxRecipientsPerMessage
            }

            Write-MyVerbose 'Exporting Send Connectors'
            $config['SendConnectors'] = Invoke-Command -Session $session -ScriptBlock {
                Get-SendConnector | Select-Object Name, AddressSpaces, SmartHosts, SourceTransportServers, Enabled, DNSRoutingEnabled, MaxMessageSize, Fqdn
            }

            Write-MyVerbose 'Exporting Transport Service configuration'
            $config['TransportService'] = Invoke-Command -Session $session -ScriptBlock {
                Get-TransportService -Identity $using:SourceServer | Select-Object MaxConcurrentMailboxDeliveries, MaxConcurrentMailboxSubmissions, MaxConnectionRatePerMinute, MaxOutboundConnections, MaxPerDomainOutboundConnections, MessageExpirationTimeout, ReceiveProtocolLogPath, SendProtocolLogPath, ConnectivityLogPath, MessageTrackingLogPath
            }

            Write-MyVerbose 'Exporting Virtual Directory URLs'
            $config['OwaVDir'] = Invoke-Command -Session $session -ScriptBlock { Get-OwaVirtualDirectory -Server $using:SourceServer | Select-Object InternalUrl, ExternalUrl }
            $config['EcpVDir'] = Invoke-Command -Session $session -ScriptBlock { Get-EcpVirtualDirectory -Server $using:SourceServer | Select-Object InternalUrl, ExternalUrl }
            $config['EwsVDir'] = Invoke-Command -Session $session -ScriptBlock { Get-WebServicesVirtualDirectory -Server $using:SourceServer | Select-Object InternalUrl, ExternalUrl }
            $config['EasVDir'] = Invoke-Command -Session $session -ScriptBlock { Get-ActiveSyncVirtualDirectory -Server $using:SourceServer | Select-Object InternalUrl, ExternalUrl }
            $config['OabVDir'] = Invoke-Command -Session $session -ScriptBlock { Get-OabVirtualDirectory -Server $using:SourceServer | Select-Object InternalUrl, ExternalUrl }
            $config['MapiVDir'] = Invoke-Command -Session $session -ScriptBlock { Get-MapiVirtualDirectory -Server $using:SourceServer | Select-Object InternalUrl, ExternalUrl }
            $config['AutodiscoverVDir'] = Invoke-Command -Session $session -ScriptBlock { Get-ClientAccessServer -Identity $using:SourceServer -ErrorAction SilentlyContinue | Select-Object AutoDiscoverServiceInternalUri }
            $config['OutlookAnywhere'] = Invoke-Command -Session $session -ScriptBlock { Get-OutlookAnywhere -Server $using:SourceServer | Select-Object InternalHostname, ExternalHostname, InternalClientsRequireSsl, ExternalClientsRequireSsl, InternalClientAuthenticationMethod, ExternalClientAuthenticationMethod }

            Write-MyVerbose 'Exporting Mailbox Database info (informational)'
            $config['MailboxDatabases'] = Invoke-Command -Session $session -ScriptBlock {
                Get-MailboxDatabase -Server $using:SourceServer -Status | Select-Object Name, EdbFilePath, LogFolderPath, ProhibitSendQuota, ProhibitSendReceiveQuota, IssueWarningQuota, CircularLoggingEnabled, DeletedItemRetention, MailboxRetention
            }

            Write-MyVerbose 'Exporting Certificate info (informational)'
            $config['Certificates'] = Invoke-Command -Session $session -ScriptBlock {
                Get-ExchangeCertificate -Server $using:SourceServer | Select-Object Thumbprint, Subject, Services, NotAfter, Status
            }

            Write-MyVerbose 'Exporting Throttling Policies (informational)'
            $config['ThrottlingPolicies'] = Invoke-Command -Session $session -ScriptBlock {
                Get-ThrottlingPolicy | Where-Object { $_.IsDefault -eq $false } | Select-Object Name, EwsMaxConcurrency, EwsMaxSubscriptions, RcaMaxConcurrency, OwaMaxConcurrency
            }
        }
        catch {
            Write-MyError ('Error during config export: {0}' -f $_.Exception.Message)
            Remove-PSSession $session -ErrorAction SilentlyContinue
            exit $ERR_CONFIGEXPORTFAILED
        }
        finally {
            Remove-PSSession $session -ErrorAction SilentlyContinue
        }

        try {
            Export-Clixml -InputObject $config -Path $configPath -ErrorAction Stop
        }
        catch {
            Write-MyError ('Failed to save config export file: {0}' -f $_.Exception.Message)
            exit $ERR_CONFIGEXPORTFAILED
        }
        $State['ServerConfigExportPath'] = $configPath
        Write-MyOutput ('Server configuration exported to {0}' -f $configPath)
    }

    function Import-ServerConfig {
        $configPath = $State['ServerConfigExportPath']
        if (-not $configPath -or -not (Test-Path $configPath)) {
            Write-MyWarning 'No server configuration export found, skipping import'
            return
        }

        Write-MyOutput ('Importing server configuration from {0}' -f $configPath)
        $config = Import-Clixml -Path $configPath

        $localServer = $env:COMPUTERNAME

        # Import Virtual Directory URLs
        $vdirMappings = @(
            @{ Name = 'OWA'; Key = 'OwaVDir'; Cmd = 'Set-OwaVirtualDirectory' }
            @{ Name = 'ECP'; Key = 'EcpVDir'; Cmd = 'Set-EcpVirtualDirectory' }
            @{ Name = 'EWS'; Key = 'EwsVDir'; Cmd = 'Set-WebServicesVirtualDirectory' }
            @{ Name = 'ActiveSync'; Key = 'EasVDir'; Cmd = 'Set-ActiveSyncVirtualDirectory' }
            @{ Name = 'OAB'; Key = 'OabVDir'; Cmd = 'Set-OabVirtualDirectory' }
            @{ Name = 'MAPI'; Key = 'MapiVDir'; Cmd = 'Set-MapiVirtualDirectory' }
        )

        foreach ($vdir in $vdirMappings) {
            if ($config[$vdir.Key]) {
                try {
                    $srcVDir = $config[$vdir.Key]
                    $identity = '{0}\{1} (Default Web Site)' -f $localServer, $vdir.Name.ToLower()
                    # Verify the virtual directory exists before attempting to set it
                    $getCmd = $vdir.Cmd -replace '^Set-', 'Get-'
                    $existing = & $getCmd -Identity $identity -ErrorAction SilentlyContinue
                    if ($null -eq $existing) {
                        Write-MyWarning ('{0} virtual directory not found at {1}, skipping' -f $vdir.Name, $identity)
                        continue
                    }
                    $params = @{ Identity = $identity }
                    if ($srcVDir.InternalUrl) { $params['InternalUrl'] = $srcVDir.InternalUrl.ToString() }
                    if ($srcVDir.ExternalUrl) { $params['ExternalUrl'] = $srcVDir.ExternalUrl.ToString() }
                    & $vdir.Cmd @params -ErrorAction Stop
                    Write-MyVerbose ('Configured {0} virtual directory URLs' -f $vdir.Name)
                }
                catch {
                    Write-MyWarning ('Failed to configure {0} virtual directory: {1}' -f $vdir.Name, $_.Exception.Message)
                }
            }
        }

        # Import Outlook Anywhere settings
        if ($config['OutlookAnywhere']) {
            try {
                $oa = $config['OutlookAnywhere']
                $params = @{ Identity = ('{0}\Rpc (Default Web Site)' -f $localServer) }
                if ($oa.InternalHostname) { $params['InternalHostname'] = $oa.InternalHostname.ToString() }
                if ($oa.ExternalHostname) { $params['ExternalHostname'] = $oa.ExternalHostname.ToString() }
                if ($oa.InternalClientsRequireSsl -ne $null) { $params['InternalClientsRequireSsl'] = $oa.InternalClientsRequireSsl }
                if ($oa.ExternalClientsRequireSsl -ne $null) { $params['ExternalClientsRequireSsl'] = $oa.ExternalClientsRequireSsl }
                Set-OutlookAnywhere @params -ErrorAction Stop
                Write-MyVerbose 'Configured Outlook Anywhere settings'
            }
            catch {
                Write-MyWarning ('Failed to configure Outlook Anywhere: {0}' -f $_.Exception.Message)
            }
        }

        # Import Autodiscover SCP
        if ($config['AutodiscoverVDir'] -and $config['AutodiscoverVDir'].AutoDiscoverServiceInternalUri) {
            try {
                Set-ClientAccessServer -Identity $localServer -AutoDiscoverServiceInternalUri $config['AutodiscoverVDir'].AutoDiscoverServiceInternalUri.ToString() -ErrorAction Stop
                Write-MyVerbose 'Configured Autodiscover Service Internal URI'
            }
            catch {
                Write-MyWarning ('Failed to configure Autodiscover URI: {0}' -f $_.Exception.Message)
            }
        }

        # Import Transport Service settings
        if ($config['TransportService']) {
            try {
                $ts = $config['TransportService']
                $tsParams = @{
                    Identity                        = $localServer
                    MaxConcurrentMailboxDeliveries  = $ts.MaxConcurrentMailboxDeliveries
                    MaxConcurrentMailboxSubmissions = $ts.MaxConcurrentMailboxSubmissions
                    MaxOutboundConnections          = $ts.MaxOutboundConnections
                    MaxPerDomainOutboundConnections = $ts.MaxPerDomainOutboundConnections
                    ErrorAction                     = 'Stop'
                }
                if ($ts.MessageExpirationTimeout) { $tsParams['MessageExpirationTimeout'] = $ts.MessageExpirationTimeout }
                Set-TransportService @tsParams
                Write-MyVerbose 'Configured Transport Service settings'
            }
            catch {
                Write-MyWarning ('Failed to configure Transport Service: {0}' -f $_.Exception.Message)
            }
        }

        # Import Receive Connectors
        if ($config['ReceiveConnectors']) {
            foreach ($rc in $config['ReceiveConnectors']) {
                try {
                    # Skip default connectors (they are created by setup)
                    $existing = Get-ReceiveConnector -Server $localServer -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $rc.Name }
                    if ($existing) {
                        Write-MyVerbose ('Receive Connector {0} already exists, updating' -f $rc.Name)
                        Set-ReceiveConnector -Identity $existing.Identity -MaxMessageSize $rc.MaxMessageSize -ErrorAction Stop
                    }
                    else {
                        Write-MyVerbose ('Creating Receive Connector {0}' -f $rc.Name)
                        New-ReceiveConnector -Name $rc.Name -Server $localServer -Bindings $rc.Bindings -RemoteIPRanges $rc.RemoteIPRanges -PermissionGroups $rc.PermissionGroups -AuthMechanism $rc.AuthMechanism -TransportRole $rc.TransportRole -ErrorAction Stop
                    }
                }
                catch {
                    Write-MyWarning ('Failed to configure Receive Connector {0}: {1}' -f $rc.Name, $_.Exception.Message)
                }
            }
        }

        # Log informational items
        if ($config['MailboxDatabases']) {
            Write-MyOutput 'Source server Mailbox Database configuration (for reference):'
            $config['MailboxDatabases'] | ForEach-Object {
                Write-MyOutput ('  DB: {0} | Path: {1} | ProhibitSend: {2} | CircularLog: {3}' -f $_.Name, $_.EdbFilePath, $_.ProhibitSendQuota, $_.CircularLoggingEnabled)
            }
        }

        if ($config['Certificates']) {
            Write-MyOutput 'Source server certificates (for reference):'
            $config['Certificates'] | ForEach-Object {
                Write-MyOutput ('  Cert: {0} | Services: {1} | Expires: {2}' -f $_.Subject, $_.Services, $_.NotAfter)
            }
        }

        Write-MyOutput 'Server configuration import completed'
    }

    function Test-DBLogPathSeparation {
        if (-not $State['MDBDBPath'] -or -not $State['MDBLogPath']) {
            Write-MyVerbose 'MDBDBPath or MDBLogPath not set, skipping DB/Log separation check'
            return
        }
        $dbRoot  = [System.IO.Path]::GetPathRoot($State['MDBDBPath']).TrimEnd('\')
        $logRoot = [System.IO.Path]::GetPathRoot($State['MDBLogPath']).TrimEnd('\')

        Write-MyOutput ('Checking DB/Log path separation — DB root: {0}  Log root: {1}' -f $dbRoot, $logRoot)

        if ($dbRoot -and $logRoot -and ($dbRoot -eq $logRoot)) {
            Write-MyWarning ('Database and transaction logs share the same volume ({0}). Microsoft recommends separate volumes for performance and recoverability.' -f $dbRoot)
        }
        else {
            Write-MyOutput 'Database and transaction logs are on separate volumes (best practice confirmed).'
        }

        if ($State['DAGName']) {
            Write-MyOutput 'DAG environment: Microsoft recommends max 2 TB per mailbox database (200 GB for lagged copies).'
        }
        else {
            Write-MyOutput 'Standalone (no DAG): Microsoft recommends keeping mailbox databases under 200 GB for optimal recoverability.'
        }
    }

    function Wait-ADReplication {
        if (-not $State['WaitForADSync']) { return }
        Write-MyOutput 'Checking AD replication health after PrepareAD (-WaitForADSync)'
        $maxAttempts = 18   # 18 x 20 s = 6 min
        $healthy     = $false
        for ($i = 1; $i -le $maxAttempts; $i++) {
            try {
                # repadmin /showrepl /errorsonly always outputs DC header lines (site\name,
                # DSA Options, object GUID, etc.) even when there are no errors. A single-DC
                # environment with no replication partners produces only these header lines.
                # Match only lines that indicate actual replication failures:
                #   "N consecutive failure(s)" — failure counter line
                #   "Last attempt @ <date> FAILED" — failure detail line
                $replErrors = & repadmin /showrepl /errorsonly 2>&1 |
                    Where-Object { $_ -match 'consecutive failure|Last attempt .* FAILED' }
                if (-not $replErrors) {
                    Write-MyOutput ('AD replication healthy (attempt {0}/{1})' -f $i, $maxAttempts)
                    $healthy = $true
                    break
                }
                Write-MyVerbose ('Replication errors ({0}/{1}): {2}' -f $i, $maxAttempts, ($replErrors -join ' | '))
                Write-MyOutput ('Waiting for AD replication... ({0}/{1})' -f $i, $maxAttempts)
            }
            catch {
                Write-MyWarning ('repadmin check failed: {0}' -f $_.Exception.Message)
                break
            }
            if ($i -lt $maxAttempts) { Start-Sleep -Seconds 20 }
        }
        if (-not $healthy) {
            Write-MyWarning 'AD replication errors still present after WaitForADSync timeout — review before continuing.'
        }
    }

    function Register-ExchangeLogCleanup {
        $days = if ($State['LogRetentionDays'] -and [int]$State['LogRetentionDays'] -gt 0) { [int]$State['LogRetentionDays'] } else { 30 }
        Write-MyOutput 'Registering Exchange log cleanup scheduled task'

        # Use folder from menu/config if already provided; otherwise prompt interactively
        $defaultScriptFolder = 'C:\#service'
        $scriptFolder = if ($State['LogCleanupFolder']) { $State['LogCleanupFolder'] } else { $defaultScriptFolder }
        if ($State['LogCleanupFolder']) {
            Write-MyVerbose ('Log cleanup folder from configuration: {0}' -f $scriptFolder)
        } elseif ([Environment]::UserInteractive) {
            Write-MyOutput ('Enter folder for log cleanup script [{0}] (ENTER = default, S = skip, auto-accept in 2 min):' -f $defaultScriptFolder)
            $inputBuffer = ''
            try {
                try { $host.UI.RawUI.FlushInputBuffer() } catch { }
                $totalSecs = 120
                $deadline = [DateTime]::Now.AddSeconds($totalSecs)
                while ([DateTime]::Now -lt $deadline) {
                    $secsLeft = [int]($deadline - [DateTime]::Now).TotalSeconds
                    Write-Progress -Id 2 -Activity 'Log cleanup folder' `
                        -Status ('Auto-accept in {0}s  |  ENTER = accept  |  S = skip' -f $secsLeft) `
                        -PercentComplete (($totalSecs - $secsLeft) * 100 / $totalSecs)
                    if ($host.UI.RawUI.KeyAvailable) {
                        $key = $host.UI.RawUI.ReadKey('IncludeKeyDown,NoEcho')
                        if ($key.VirtualKeyCode -eq 13) {           # Enter
                            Write-Host ''
                            break
                        }
                        elseif ($key.VirtualKeyCode -eq 27) {       # Escape — use default
                            $inputBuffer = ''
                            Write-Host ''
                            break
                        }
                        elseif ($key.VirtualKeyCode -eq 8) {        # Backspace
                            if ($inputBuffer.Length -gt 0) {
                                $inputBuffer = $inputBuffer.Substring(0, $inputBuffer.Length - 1)
                                Write-Host "`b `b" -NoNewline
                            }
                        }
                        elseif ($key.Character -ge ' ') {
                            $inputBuffer += $key.Character
                            Write-Host $key.Character -NoNewline
                        }
                    }
                    Start-Sleep -Milliseconds 100
                }
                Write-Progress -Id 2 -Activity 'Log cleanup folder' -Completed
                if ($inputBuffer.Trim().ToUpper() -eq 'S') {
                    Write-MyVerbose 'Log cleanup task registration skipped by user'
                    return
                }
                if ($inputBuffer.Trim() -ne '') { $scriptFolder = $inputBuffer.Trim() }
            }
            catch {
                # Console does not support RawUI — accept default silently (non-interactive environment)
                Write-MyVerbose ('Log cleanup folder auto-accepted (no interactive console): {0}' -f $scriptFolder)
            }
        }

        if (-not (Test-Path $scriptFolder)) {
            New-Item -Path $scriptFolder -ItemType Directory -Force | Out-Null
            Write-MyVerbose ('Created script folder: {0}' -f $scriptFolder)
        }

        $scriptPath = Join-Path $scriptFolder 'Invoke-ExchangeLogCleanup.ps1'
        $logFolder  = Join-Path $scriptFolder 'logs'

        $cleanupScript = @"
# Exchange Log Cleanup Script — generated by EXpress.ps1
# Runs daily via Scheduled Task; retention: $days days for Exchange/IIS logs, 30 days for own logs

param([int]`$DaysToKeep = $days)

`$ScriptDir  = Split-Path -Path `$MyInvocation.MyCommand.Path
`$LogFolder  = Join-Path `$ScriptDir 'logs'
`$LogFile    = Join-Path `$LogFolder ('LogCleanup_{0}.log' -f (Get-Date -Format 'yyyyMM'))
`$cutoff     = (Get-Date).AddDays(-`$DaysToKeep)

if (-not (Test-Path `$LogFolder)) { New-Item -Path `$LogFolder -ItemType Directory | Out-Null }

function Write-Log {
    param([string]`$Message, [string]`$Level = 'Info')
    `$line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), `$Level, `$Message
    Add-Content -Path `$LogFile -Value `$line
}

Write-Log 'Exchange log cleanup started'
Write-Log ('Removing files older than {0} days' -f `$DaysToKeep)

# IIS logs — try dynamic path from metabase, fall back to default
`$iisRoot = `$null
try {
    Import-Module WebAdministration -ErrorAction Stop
    `$iisRoot = ((Get-WebConfigurationProperty -Filter 'system.applicationHost/sites/siteDefaults' -Name logFile).directory) -replace '%SystemDrive%', `$env:SystemDrive
} catch { }
if (-not `$iisRoot) { `$iisRoot = Join-Path `$env:SystemDrive 'inetpub\logs\LogFiles' }
if (Test-Path `$iisRoot) {
    `$files = @(Get-ChildItem -Path `$iisRoot -Recurse -File -Filter '*.log' | Where-Object { `$_.LastWriteTime -lt `$cutoff })
    `$files | Remove-Item -Force -ErrorAction SilentlyContinue
    Write-Log ('IIS: removed {0} log file(s) from {1}' -f `$files.Count, `$iisRoot)
}

# Exchange logs — entire Logging\ and TransportRoles\Logs\ trees (covers EWS, OWA, HttpProxy, RpcClientAccess, transport, tracking, monitoring, etc.)
`$exSetup = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup' -ErrorAction SilentlyContinue).MsiInstallPath
if (`$exSetup) {
    foreach (`$path in @((Join-Path `$exSetup 'Logging'), (Join-Path `$exSetup 'TransportRoles\Logs'))) {
        if (Test-Path `$path) {
            `$files = @(Get-ChildItem -Path `$path -Recurse -File -Filter '*.log' | Where-Object { `$_.LastWriteTime -lt `$cutoff })
            `$files | Remove-Item -Force -ErrorAction SilentlyContinue
            if (`$files.Count -gt 0) { Write-Log ('Exchange: removed {0} file(s) from {1}' -f `$files.Count, `$path) }
        }
    }
}

# HTTPERR logs
`$httpErrPath = Join-Path `$env:SystemRoot 'System32\LogFiles\HTTPERR'
if (Test-Path `$httpErrPath) {
    `$files = @(Get-ChildItem -Path `$httpErrPath -File -Filter '*.log' | Where-Object { `$_.LastWriteTime -lt `$cutoff })
    `$files | Remove-Item -Force -ErrorAction SilentlyContinue
    if (`$files.Count -gt 0) { Write-Log ('HTTPERR: removed {0} file(s) from {1}' -f `$files.Count, `$httpErrPath) }
}

# Self-cleanup: purge own log files older than 30 days
`$ownCutoff = (Get-Date).AddDays(-30)
Get-ChildItem -Path `$LogFolder -File -Filter '*.log' |
    Where-Object { `$_.LastWriteTime -lt `$ownCutoff } |
    Remove-Item -Force -ErrorAction SilentlyContinue

Write-Log 'Exchange log cleanup finished'
"@
        try {
            $cleanupScript | Out-File -FilePath $scriptPath -Encoding utf8 -Force
            Write-MyOutput ('Log cleanup script saved to: {0}' -f $scriptPath)

            $action    = New-ScheduledTaskAction -Execute 'powershell.exe' `
                             -Argument ('-NonInteractive -NoProfile -ExecutionPolicy Bypass -File "{0}"' -f $scriptPath)
            $trigger   = New-ScheduledTaskTrigger -Daily -At '02:00'
            $settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2)
            $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
            $taskName  = 'Exchange Log Cleanup'
            $taskPath  = '\Exchange\'
            Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath -ErrorAction SilentlyContinue |
                Unregister-ScheduledTask -Confirm:$false
            Register-ExecutedCommand -Category 'ScheduledTask' -Command ("Register-ScheduledTask -TaskName '$taskName' -TaskPath '$taskPath' -Action (New-ScheduledTaskAction …Clean-ExchangeLogs.ps1 -RetentionDays $days) -Trigger (Daily 02:00) -Principal SYSTEM -RunLevel Highest")
            Register-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Action $action `
                -Trigger $trigger -Settings $settings -Principal $principal -ErrorAction Stop | Out-Null
            Write-MyOutput ('Scheduled task "{0}" registered — runs daily at 02:00, retention {1} days' -f $taskName, $days)
        }
        catch {
            Write-MyWarning ('Failed to register log cleanup task: {0}' -f $_.Exception.Message)
        }
    }
    function Get-MEACAutomationCredentialFromState {
        # Rehydrates -MEACAutomationCredential across the Autopilot reboot chain.
        # Only used for AD Split-Permissions deployments where a Domain Admin has
        # pre-created the SystemMailbox{b963af59-…} account via -MEACPrepareADOnly
        # and the credential needs to survive reboots between Phase 0 and Phase 6.
        # Standard (non-Split) deployments never populate this — MEAC self-provisions.
        if (-not $State['MEACAutomationUser'] -or -not $State['MEACAutomationPW']) { return $null }
        try {
            $sec = ConvertTo-SecureString $State['MEACAutomationPW'] -ErrorAction Stop
            return New-Object System.Management.Automation.PSCredential($State['MEACAutomationUser'], $sec)
        }
        catch {
            Write-MyWarning ('MEAC: could not decrypt stored automation credential (DPAPI mismatch?): {0}' -f $_.Exception.Message)
            return $null
        }
    }

    function Register-AuthCertificateRenewal {
        # MEAC — CSS-Exchange MonitorExchangeAuthCertificate.ps1. Creates a daily
        # scheduled task that auto-renews the Exchange Auth Certificate 60 days
        # before expiry. Without it, Auth Cert expiry causes a full outage
        # (OAuth, Hybrid, EWS). Skip on Edge and Management-only installs.
        #
        # v5.93: MEAC self-provisions SystemMailbox{b963af59-…}, the Auth Certificate
        # Management role group, and the batch-logon right. EXpress does not layer a
        # credential system on top — but MEAC still needs a password supplied at
        # registration time so Task Scheduler can run the daily task AS that user
        # (Task Scheduler refuses to register a task-under-user-identity without a
        # credential). EXpress generates a strong random password inline and passes
        # it via MEAC -Password; it is never persisted.
        #
        # v5.93: hybrid-aware + documented passthroughs. Detects hybrid via
        # Get-HybridConfiguration; in hybrid mode MEAC refuses to renew by default
        # (to avoid breaking HCW federation silently) — operator opts into renewal
        # via -MEACIgnoreHybridConfig with the implicit promise to rerun HCW.
        # Split-Permissions path: DA pre-creates the account (separate run with
        # -MEACPrepareADOnly); Exchange admin passes the resulting credential here
        # as -MEACAutomationCredential, forwarded to MEAC -AutomationAccountCredential.
        if ($State['InstallEdge'] -or $State['InstallManagementTools'] -or $State['InstallRecipientManagement']) {
            Write-MyVerbose 'Auth Certificate renewal task: not applicable to this install mode'
            return
        }
        # Idempotency: on re-run after a failure, skip MEAC entirely when both
        # the scheduled task and the auto-provisioned automation account already
        # exist. Re-invoking MEAC in this state is redundant and can surface
        # spurious errors from the registration path.
        $existingTask = Get-ScheduledTask -TaskName 'Daily Auth Certificate Check' -ErrorAction SilentlyContinue
        $existingUser = $null
        if (Get-Command Get-User -ErrorAction SilentlyContinue) {
            $existingUser = Get-User -Identity 'SystemMailbox{b963af59-3975-4f92-9d58-ad0b1fe3a1a3}' -ErrorAction SilentlyContinue
        }
        if ($existingTask -and $existingUser) {
            Write-MyOutput 'Auth Certificate renewal already registered (scheduled task + automation account present) - skipping MEAC'
            return
        }
        Write-MyOutput 'Registering Auth Certificate renewal (MEAC / CSS-Exchange MonitorExchangeAuthCertificate.ps1)'
        $meacPath = Join-Path $State['SourcesPath'] 'MonitorExchangeAuthCertificate.ps1'
        $meacUrl  = 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/MonitorExchangeAuthCertificate.ps1'
        if (-not (Test-Path $meacPath)) {
            try {
                Invoke-WebDownload -Uri $meacUrl -OutFile $meacPath
                Write-MyVerbose ('MEAC downloaded, SHA256: {0}' -f (Get-FileHash $meacPath -Algorithm SHA256).Hash)
            }
            catch {
                Write-MyWarning ('Could not download MonitorExchangeAuthCertificate.ps1: {0}' -f $_.Exception.Message)
                return
            }
        }

        # --- Hybrid detection (transparent) ---------------------------------------
        # Get-HybridConfiguration returns an object only when HCW has been run.
        # Blank / non-existent object → not hybrid. Errors are treated as not-hybrid
        # rather than blocking, because MEAC itself re-detects authoritatively.
        $hybridDetected = $false
        if (Get-Command Get-HybridConfiguration -ErrorAction SilentlyContinue) {
            try {
                $hc = Get-HybridConfiguration -ErrorAction Stop
                $hybridDetected = [bool]($hc -and ($hc.Domains -or $hc.Features -or $hc.ClientAccessServers -or $hc.TransportServers))
            } catch { $hybridDetected = $false }
        }

        # --- Build parameter set for MEAC -----------------------------------------
        $meacParams = @{
            ConfigureScriptToRunViaScheduledTask = $true
            Confirm                              = $false
        }

        # Split-Permissions: DA pre-created the automation account; pass it through.
        # Prefer state (survives Autopilot reboot chain) over current-run CLI parameter.
        $autoCred = if ($State['MEACAutomationUser']) { Get-MEACAutomationCredentialFromState } else { $MEACAutomationCredential }
        if ($autoCred) {
            $meacParams.AutomationAccountCredential = $autoCred
            Write-MyOutput ('MEAC: using pre-created automation account {0} (AD Split-Permissions passthrough)' -f $autoCred.UserName)
        }
        else {
            # Standard (non-Split) deployment: MEAC self-provisions the
            # SystemMailbox{b963af59-…} account, but Task Scheduler still needs a
            # password to register a task that runs as that user. MEAC accepts
            # either -Password <SecureString> or -AutomationAccountCredential; we
            # take the simpler path and generate a strong random password inline.
            #
            # The password is transient: once MEAC sets it on the account AND
            # registers the scheduled task, Windows stores the credential in the
            # task's (DPAPI-protected) credential store. EXpress never needs it
            # again — if re-registration becomes necessary (e.g. password policy
            # rotation), re-running this function generates a fresh password and
            # MEAC resets both the account and the task atomically.
            #
            # Therefore: not persisted to state, not logged, not returned.
            $charset = [char[]](
                (65..90) +             # A-Z
                (97..122) +            # a-z
                (50..57) +             # 2-9  (skip 0/1 for clarity)
                @(33, 35, 36, 37, 38, 42, 43, 45, 61, 63, 64)   # ! # $ % & * + - = ? @
            )
            $pwBytes = [byte[]]::new(32)
            [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($pwBytes)
            $pwChars = foreach ($b in $pwBytes) { $charset[$b % $charset.Length] }
            $pwSecure = ConvertTo-SecureString -String (-join $pwChars) -AsPlainText -Force
            $meacParams.Password = $pwSecure
            Remove-Variable pwBytes, pwChars -ErrorAction SilentlyContinue
            Write-MyVerbose 'MEAC: generated transient 32-char password for SystemMailbox{b963af59-…} automation account (not persisted; Task Scheduler stores it DPAPI-protected)'
        }

        if ($MEACIgnoreHybridConfig) {
            $meacParams.IgnoreHybridConfig = $true
            Write-MyWarning 'MEAC: -IgnoreHybridConfig enabled. MEAC will renew the Auth Certificate when due; you MUST rerun the Hybrid Configuration Wizard afterwards or OAuth/federation with Exchange Online will break.'
        }
        if ($MEACIgnoreUnreachableServers) {
            $meacParams.IgnoreUnreachableServers = $true
            Write-MyVerbose 'MEAC: -IgnoreUnreachableServers enabled'
        }
        if ($MEACNotificationEmail) {
            $meacParams.SendEmailNotificationTo = $MEACNotificationEmail
            Write-MyOutput ('MEAC: renewal notifications will be sent to {0}' -f $MEACNotificationEmail)
        }

        # Hybrid advisory — transparent to the operator.
        if ($hybridDetected -and -not $MEACIgnoreHybridConfig) {
            Write-MyOutput ''
            Write-MyOutput 'MEAC: Hybrid configuration detected.'
            Write-MyOutput '      Task registered in hybrid-safe mode — MEAC will REFUSE to renew the'
            Write-MyOutput '      Auth Certificate until -MEACIgnoreHybridConfig is passed (which also'
            Write-MyOutput '      obliges you to rerun HCW afterwards to re-federate with Exchange Online).'
            Write-MyOutput '      Without the flag, daily checks still run; wire up -MEACNotificationEmail'
            Write-MyOutput '      to receive an alert 60 days before expiry.'
            Write-MyOutput ''
        }

        # --- Run MEAC -------------------------------------------------------------
        try {
            Push-Location $State['InstallPath']
            $meacArgStr = ($meacParams.GetEnumerator() | ForEach-Object {
                $v = if ($_.Value -is [System.Security.SecureString]) { '<SecureString>' }
                     elseif ($_.Value -is [System.Management.Automation.PSCredential]) { '<PSCredential>' }
                     elseif ($_.Value -is [bool] -or $_.Value -is [switch]) { '' }
                     else { "'$($_.Value)'" }
                if ($v) { "-$($_.Key) $v" } else { "-$($_.Key)" }
            }) -join ' '
            Register-ExecutedCommand -Category 'ScheduledTask' -Command (".\MonitorExchangeAuthCertificate.ps1 $meacArgStr")
            & $meacPath @meacParams *>&1 | ForEach-Object { Write-MyVerbose ('MEAC: {0}' -f $_) }
        }
        catch {
            Write-MyWarning ('MEAC registration failed: {0}' -f $_.Exception.Message)
            return
        }
        finally {
            Pop-Location
        }
        $meacTask = Get-ScheduledTask -TaskName 'Daily Auth Certificate Check' -ErrorAction SilentlyContinue
        if ($meacTask) {
            Write-MyOutput 'MEAC scheduled task registered — auth cert will auto-renew 60 days before expiry'
        }
        else {
            Write-MyWarning 'MEAC: task "Daily Auth Certificate Check" not found after registration — check MEAC log in Exchange Logging\AuthCertificateMonitoring\ for details'
        }
    }

    function Add-ServerToSendConnectors {
        if ($State['InstallEdge']) {
            Write-MyVerbose 'Edge role — skipping Send Connector update'
            return
        }
        try {
            $sendConnectors = Get-SendConnector -ErrorAction Stop | Where-Object {
                $srvList = @($_.SourceTransportServers | ForEach-Object { $_.Name })
                $srvList -notcontains $env:COMPUTERNAME
            }
            if (-not $sendConnectors -or $sendConnectors.Count -eq 0) {
                Write-MyVerbose 'All Send Connectors already include this server'
                return
            }
            Write-MyOutput ('{0} Send Connector(s) do not include this server:' -f $sendConnectors.Count)
            foreach ($sc in $sendConnectors) {
                Write-MyOutput ('  - {0}' -f $sc.Name)
            }
            $answer = 'Y'
            if ([Environment]::UserInteractive) {
                Write-MyOutput 'Add this server as source transport server? [Y/N/S=skip] (default: Y):'
                try {
                    $key = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                    $answer = $key.Character.ToString().ToUpper()
                }
                catch { $answer = 'Y' }
                Write-MyOutput $answer
            }
            if ($answer -notin @('N', 'S')) {
                foreach ($sc in $sendConnectors) {
                    $sources = [System.Collections.Generic.List[object]]($sc.SourceTransportServers)
                    $sources.Add($env:COMPUTERNAME) | Out-Null
                    Register-ExecutedCommand -Category 'SendConnector' -Command ("Set-SendConnector -Identity '$($sc.Identity)' -SourceTransportServers $($sources -join ',')")
                    Set-SendConnector -Identity $sc.Identity -SourceTransportServers $sources -ErrorAction Stop
                    Write-MyOutput ('Added {0} to Send Connector: {1}' -f $env:COMPUTERNAME, $sc.Name)
                }
            }
            else {
                Write-MyVerbose 'Send Connector update skipped by user'
            }
        }
        catch {
            Write-MyWarning ('Failed to update Send Connectors: {0}' -f $_.Exception.Message)
        }
    }

    function Install-AntispamAgents {
        if ($State['InstallEdge']) {
            Write-MyVerbose 'Edge role has antispam agents built-in, skipping'
            return
        }
        $exSetup = (Get-ItemProperty -Path $EXCHANGEINSTALLKEY -ErrorAction SilentlyContinue).MsiInstallPath
        if (-not $exSetup) {
            Write-MyWarning 'Exchange install path not found — cannot install antispam agents'
            return
        }
        $installScript = Join-Path $exSetup 'Scripts\Install-AntispamAgents.ps1'
        if (-not (Test-Path $installScript)) {
            Write-MyWarning ('Install-AntispamAgents.ps1 not found at: {0}' -f $installScript)
            return
        }

        # Check if antispam agents are already installed by looking for a known agent
        $existingAgents = Get-TransportAgent -ErrorAction SilentlyContinue |
                          Where-Object { $_.Identity -like '*Filter*' -or $_.Identity -like '*Antispam*' }
        if (-not $existingAgents) {
            Write-MyOutput 'Installing Exchange antispam agents'
            # Capture everything that bypasses the pipeline (Write-Host / $host.UI.WriteWarningLine /
            # [Console]::WriteLine) by redirecting Console.Out + Console.Error to StringWriters.
            # Combined with *>&1 this catches both stream-based and host-UI-based output.
            $capOut    = [System.IO.StringWriter]::new()
            $capErr    = [System.IO.StringWriter]::new()
            $origOut   = [System.Console]::Out
            $origErr   = [System.Console]::Error
            $records   = $null
            # Suppress warnings in the current session scope: the Exchange PSSnapin autoload
            # ("Please exit Windows PowerShell", "restart MSExchangeTransport") fires in the
            # caller's session — not in the child script's streams — so *>&1 alone cannot
            # capture it. $WarningPreference = SilentlyContinue mutes this session-level noise.
            $savedWP = $WarningPreference
            $WarningPreference = 'SilentlyContinue'
            try {
                [System.Console]::SetOut($capOut)
                [System.Console]::SetError($capErr)
                $records = & $installScript *>&1
            }
            catch {
                [System.Console]::SetOut($origOut)
                [System.Console]::SetError($origErr)
                Write-MyWarning ('Failed to install antispam agents: {0}' -f $_.Exception.Message)
                return
            }
            finally {
                [System.Console]::SetOut($origOut)
                [System.Console]::SetError($origErr)
                $WarningPreference = $savedWP
            }

            # Route pipeline records: warnings/verbose to debug log, Exchange agent objects
            # collected into a single summary list, everything else to stdout.
            # Suppresses the well-known "Please restart MSExchangeTransport" warning from
            # the Exchange-shipped Install-AntispamAgents.ps1 (we restart the service below).
            $agentSummary = [System.Collections.Generic.List[string]]::new()
            foreach ($r in $records) {
                if ($null -eq $r) { continue }
                if ($r -is [System.Management.Automation.WarningRecord]) {
                    $msg = $r.Message
                    # We restart MSExchangeTransport below, so suppress the "restart required"
                    # and PSSnapin-autoload "exit Windows PowerShell" warnings shipped with
                    # Install-AntispamAgents.ps1.
                    if ($msg -match '(restart is required|restart the Microsoft Exchange Transport|exit Windows PowerShell to complete)') { continue }
                    Write-MyDebug ('[antispam] WARN: {0}' -f $msg)
                }
                elseif ($r -is [System.Management.Automation.ErrorRecord]) {
                    Write-MyWarning ('[antispam] {0}' -f $r.Exception.Message)
                }
                elseif ($r -is [System.Management.Automation.VerboseRecord] -or
                        $r -is [System.Management.Automation.DebugRecord]) {
                    Write-MyDebug ('[antispam] {0}' -f $r.Message)
                }
                elseif ($r -is [System.Management.Automation.InformationRecord]) {
                    Write-MyDebug ('[antispam] {0}' -f $r.MessageData)
                }
                else {
                    # TransportAgent objects — each would otherwise render as its own table.
                    # Collect a compact one-liner per agent instead.
                    $idName = if ($r.PSObject.Properties['Identity']) { [string]$r.Identity } else { '' }
                    if ($idName) {
                        $prio    = if ($r.PSObject.Properties['Priority']) { $r.Priority } else { '-' }
                        $enabled = if ($r.PSObject.Properties['Enabled'])  { $r.Enabled  } else { '-' }
                        $agentSummary.Add(('  {0,-32} Priority={1,-3} Enabled={2}' -f $idName, $prio, $enabled))
                    }
                    else {
                        $line = ($r | Out-String).TrimEnd()
                        if ($line) { Write-MyDebug ('[antispam] {0}' -f $line) }
                    }
                }
            }
            if ($agentSummary.Count -gt 0) {
                Write-MyOutput ('[antispam] Installed {0} transport agents:' -f $agentSummary.Count)
                foreach ($s in $agentSummary) { Write-MyOutput $s }
            }

            # Route host-UI output captured via Console redirection — demote to debug.
            foreach ($captured in @($capOut.ToString(), $capErr.ToString())) {
                if (-not $captured) { continue }
                foreach ($line in ($captured -split "`r?`n")) {
                    if (-not $line) { continue }
                    if ($line -match 'restart the Microsoft Exchange Transport') { continue }
                    Write-MyDebug ('[antispam] {0}' -f $line.TrimEnd())
                }
            }

            $installDidRun = $true
        }
        else {
            Write-MyVerbose ('Antispam agents already installed ({0} found), skipping install script' -f $existingAgents.Count)
            $installDidRun = $false
        }

        # Configure agents: only Recipient Filter Agent enabled, all other antispam
        # agents shipped by Install-AntispamAgents.ps1 disabled.
        #
        # Enumerate via Get-TransportAgent instead of hard-coded names — the exact
        # Identity casing varies across CUs/locales ("Sender Id Agent" vs
        # "Sender ID Agent"); a hard-coded lookup left Sender ID / Protocol Analysis
        # enabled in the field. Match by regex against all antispam agent names
        # and pass the actual Identity back to Enable/Disable-TransportAgent.
        $antispamRegex  = '(Content Filter|Sender\s*Id|Sender Filter|Recipient Filter|Protocol Analysis)'
        $recipientRegex = 'Recipient Filter'

        $configChanged = $false
        $savedWP       = $WarningPreference
        $WarningPreference = 'SilentlyContinue'   # mutes session-level "restart required" /
                                                  # "exit Windows PowerShell" host-UI warnings
                                                  # emitted by Enable-/Disable-TransportAgent
        try {
            $allAgents = @(Get-TransportAgent -ErrorAction SilentlyContinue)
            foreach ($agent in $allAgents) {
                $id = [string]$agent.Identity
                if ($id -notmatch $antispamRegex) { continue }
                $wantEnabled = ($id -match $recipientRegex)
                if ($agent.Enabled -eq $wantEnabled) {
                    Write-MyVerbose ('Already {0}: {1}' -f ($(if ($wantEnabled) {'enabled'} else {'disabled'})), $id)
                    continue
                }
                if ($wantEnabled) {
                    Register-ExecutedCommand -Category 'Antispam' -Command ("Enable-TransportAgent -Identity '$id'")
                    Enable-TransportAgent  -Identity $id -Confirm:$false -WarningAction SilentlyContinue -ErrorAction SilentlyContinue *>&1 | Out-Null
                    Write-MyOutput ('Enabled: {0}' -f $id)
                }
                else {
                    Register-ExecutedCommand -Category 'Antispam' -Command ("Disable-TransportAgent -Identity '$id'")
                    Disable-TransportAgent -Identity $id -Confirm:$false -WarningAction SilentlyContinue -ErrorAction SilentlyContinue *>&1 | Out-Null
                    Write-MyOutput ('Disabled: {0}' -f $id)
                }
                $configChanged = $true
            }
        }
        finally {
            $WarningPreference = $savedWP
        }

        # Enable the recipient lookup against the GAL. Without this the Recipient
        # Filter Agent only applies block-list / tarpit logic — the main value
        # (rejecting mail for non-existent recipients on Authoritative domains)
        # stays off. Accepted Domains with DomainType=Authoritative are covered
        # implicitly; Internal/External Relay domains are skipped by design.
        try {
            $rfc = Get-RecipientFilterConfig -ErrorAction Stop
            if (-not $rfc.RecipientValidationEnabled) {
                Register-ExecutedCommand -Category 'Antispam' -Command 'Set-RecipientFilterConfig -RecipientValidationEnabled $true'
                Set-RecipientFilterConfig -RecipientValidationEnabled $true -Confirm:$false -ErrorAction Stop
                Write-MyOutput 'Enabled recipient lookup (RecipientFilterConfig.RecipientValidationEnabled = True)'
                $configChanged = $true
            }
            else {
                Write-MyVerbose 'Recipient lookup already enabled'
            }
        }
        catch {
            Write-MyWarning ('Could not configure RecipientFilterConfig: {0}' -f $_.Exception.Message)
        }

        # One restart, at the end — covers both the install (if it ran) and any
        # enable/disable changes. Nothing downstream looks at the agents before
        # this point, so a single bounce is sufficient.
        if ($installDidRun -or $configChanged) {
            Write-MyOutput 'Restarting MSExchangeTransport (may take ~30s)'
            Restart-Service MSExchangeTransport -Force -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            Write-MyOutput 'MSExchangeTransport restarted'
        }
        Write-MyOutput 'Antispam agents configured: only Recipient Filter Agent is enabled'
    }

    function New-AnonymousRelayConnector {
        $hasInternal = $State['RelaySubnets']         -and $State['RelaySubnets'].Count         -gt 0
        $hasExternal = $State['ExternalRelaySubnets'] -and $State['ExternalRelaySubnets'].Count -gt 0

        if (-not $hasInternal -and -not $hasExternal) {
            Write-MyVerbose 'No RelaySubnets or ExternalRelaySubnets specified, skipping relay connector setup'
            return
        }
        if ($State['InstallEdge']) {
            Write-MyVerbose 'Edge role — skipping relay connector setup (use EdgeSync for relay)'
            return
        }

        $server  = $env:COMPUTERNAME
        $success = $true

        # --- Internal relay connector (no Ms-Exch-SMTP-Accept-Any-Recipient) ---
        # Anonymous senders can deliver to accepted domains only; cannot relay externally.
        # RFC 5737 TEST-NET placeholder — never routable, used when no subnets were specified.
        # Prevents the relay connector from matching real SMTP traffic until the admin sets proper IPs.
        $RELAY_PLACEHOLDER          = '192.0.2.1/32'
        $RELAY_PLACEHOLDER_EXTERNAL = '192.0.2.2/32'   # Different RFC 5737 address — avoids Bindings+RemoteIPRanges conflict
        $internalIsPlaceholder = ($State['RelaySubnets'].Count -eq 1 -and $State['RelaySubnets'][0] -eq $RELAY_PLACEHOLDER)
        $externalIsPlaceholder = ($State['ExternalRelaySubnets'].Count -eq 1 -and $State['ExternalRelaySubnets'][0] -in @($RELAY_PLACEHOLDER, $RELAY_PLACEHOLDER_EXTERNAL))

        if ($hasInternal) {
            $intName    = ('Anonymous Internal Relay - {0}' -f $server)
            $subnetList = $State['RelaySubnets'] -join ', '
            if ($internalIsPlaceholder) {
                Write-MyWarning 'Internal relay connector: no subnets specified — using placeholder IP (192.0.2.1/32, RFC 5737).'
                Write-MyWarning 'No real SMTP traffic will match this connector until you set RemoteIPRanges to your actual relay sources.'
            }
            Write-MyOutput ('Configuring internal relay connector "{0}" — subnets: {1}' -f $intName, $subnetList)
            try {
                $existing = Get-ReceiveConnector -Identity "$server\$intName" -ErrorAction SilentlyContinue
                if ($existing) {
                    Register-ExecutedCommand -Category 'ReceiveConnector' -Command ("Set-ReceiveConnector -Identity '$server\$intName' -RemoteIPRanges $($State['RelaySubnets'] -join ',') -AuthMechanism Tls -ProtocolLoggingLevel Verbose -Banner '220 Mail Service'")
                    Set-ReceiveConnector -Identity "$server\$intName" -RemoteIPRanges $State['RelaySubnets'] `
                        -AuthMechanism Tls -ProtocolLoggingLevel Verbose -Banner '220 Mail Service' -ErrorAction Stop
                    Write-MyVerbose 'Internal relay connector already exists — RemoteIPRanges, TLS, logging and banner updated'
                }
                else {
                    Register-ExecutedCommand -Category 'ReceiveConnector' -Command ("New-ReceiveConnector -Name '$intName' -Server '$server' -TransportRole FrontendTransport -RemoteIPRanges $($State['RelaySubnets'] -join ',') -Bindings '0.0.0.0:25' -PermissionGroups AnonymousUsers -AuthMechanism Tls -ProtocolLoggingLevel Verbose -Banner '220 Mail Service'")
                    New-ReceiveConnector -Name $intName -Server $server -TransportRole FrontendTransport `
                        -RemoteIPRanges $State['RelaySubnets'] -Bindings '0.0.0.0:25' `
                        -PermissionGroups AnonymousUsers -AuthMechanism Tls `
                        -ProtocolLoggingLevel Verbose -Banner '220 Mail Service' -ErrorAction Stop | Out-Null
                    Write-MyOutput 'Internal relay connector created (TLS offered, accepted domains only, no external relay right, hardened banner)'
                }
            }
            catch {
                Write-MyWarning ('Failed to configure internal relay connector: {0}' -f $_.Exception.Message)
                $success = $false
            }
        }

        # --- External relay connector (Ms-Exch-SMTP-Accept-Any-Recipient granted) ---
        # Anonymous senders from these IPs can relay to any recipient including external.
        if ($hasExternal) {
            $extName = ('Anonymous External Relay - {0}' -f $server)
            Write-MyWarning 'SECURITY: External relay connector allows anonymous relay to ANY recipient.'
            if ($externalIsPlaceholder) {
                Write-MyWarning 'External relay connector: no subnets specified — using placeholder IP (192.0.2.2/32, RFC 5737).'
                Write-MyWarning 'No real SMTP traffic will match this connector until you set RemoteIPRanges to your actual relay sources.'
            }
            Write-MyWarning ('         External relay subnets: {0}' -f ($State['ExternalRelaySubnets'] -join ', '))
            try {
                # Resolve ANONYMOUS LOGON by SID (S-1-5-7) — language-independent
                $anonLogon = ([System.Security.Principal.SecurityIdentifier]'S-1-5-7').Translate(
                    [System.Security.Principal.NTAccount]).Value
                Write-MyVerbose ('Resolved ANONYMOUS LOGON account: {0}' -f $anonLogon)

                $existing = Get-ReceiveConnector -Identity "$server\$extName" -ErrorAction SilentlyContinue
                $connObj = $null
                if ($existing) {
                    Register-ExecutedCommand -Category 'ReceiveConnector' -Command ("Set-ReceiveConnector -Identity '$server\$extName' -RemoteIPRanges $($State['ExternalRelaySubnets'] -join ',') -AuthMechanism Tls -ProtocolLoggingLevel Verbose -Banner '220 Mail Service'")
                    Set-ReceiveConnector -Identity "$server\$extName" -RemoteIPRanges $State['ExternalRelaySubnets'] `
                        -AuthMechanism Tls -ProtocolLoggingLevel Verbose -Banner '220 Mail Service' -ErrorAction Stop
                    Write-MyVerbose 'External relay connector already exists — RemoteIPRanges, TLS, logging and banner updated'
                    $connObj = $existing
                }
                else {
                    # Capture the returned object directly to avoid a race condition where
                    # Get-ReceiveConnector fails immediately after creation (Exchange AD not yet updated).
                    Register-ExecutedCommand -Category 'ReceiveConnector' -Command ("New-ReceiveConnector -Name '$extName' -Server '$server' -TransportRole FrontendTransport -RemoteIPRanges $($State['ExternalRelaySubnets'] -join ',') -Bindings '0.0.0.0:25' -PermissionGroups AnonymousUsers -AuthMechanism Tls -ProtocolLoggingLevel Verbose -Banner '220 Mail Service'")
                    $connObj = New-ReceiveConnector -Name $extName -Server $server -TransportRole FrontendTransport `
                        -RemoteIPRanges $State['ExternalRelaySubnets'] -Bindings '0.0.0.0:25' `
                        -PermissionGroups AnonymousUsers -AuthMechanism Tls `
                        -ProtocolLoggingLevel Verbose -Banner '220 Mail Service' -ErrorAction Stop
                }
                # Fallback: if the object is somehow null, retry Get-ReceiveConnector with backoff
                if (-not $connObj) {
                    for ($retry = 1; $retry -le 3 -and -not $connObj; $retry++) {
                        Write-MyVerbose ('Waiting for external relay connector to register in Exchange (attempt {0}/3)...' -f $retry)
                        Start-Sleep -Seconds 5
                        $connObj = Get-ReceiveConnector -Identity "$server\$extName" -ErrorAction SilentlyContinue
                    }
                }
                Register-ExecutedCommand -Category 'ReceiveConnector' -Command ("Get-ReceiveConnector '$server\$extName' | Add-ADPermission -User '$anonLogon' -ExtendedRights 'Ms-Exch-SMTP-Accept-Any-Recipient'")
                $connObj | Add-ADPermission -User $anonLogon `
                    -ExtendedRights 'Ms-Exch-SMTP-Accept-Any-Recipient' -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
                Write-MyOutput ('External relay connector created with Ms-Exch-SMTP-Accept-Any-Recipient for {0}' -f $anonLogon)
            }
            catch {
                Write-MyWarning ('Failed to configure external relay connector: {0}' -f $_.Exception.Message)
                $success = $false
            }
        }

        # --- Remove AnonymousUsers from Default Frontend connector ---
        # Only when at least one dedicated relay connector was configured successfully.
        # This prevents unauthenticated inbound from arbitrary IPs while keeping
        # relay restricted to the explicitly defined subnets above.
        # Skip Default Frontend hardening when only placeholder IPs are set — the relay connector
        # won't match real traffic yet, so removing AnonymousUsers would break inbound mail.
        $onlyPlaceholders = ($hasInternal -and $internalIsPlaceholder) -and (-not $hasExternal -or $externalIsPlaceholder)
        if ($success -and -not $onlyPlaceholders) {
            $defaultName = ('Default Frontend {0}' -f $server)
            try {
                $rc = Get-ReceiveConnector -Identity "$server\$defaultName" -ErrorAction SilentlyContinue
                if ($rc -and ($rc.PermissionGroups -match 'AnonymousUsers')) {
                    $pgList  = ($rc.PermissionGroups.ToString() -split ',\s*') | Where-Object { $_.Trim() -ne 'AnonymousUsers' }
                    Register-ExecutedCommand -Category 'ReceiveConnector' -Command ("Set-ReceiveConnector -Identity '$server\$defaultName' -PermissionGroups '$($pgList -join ',')'  # AnonymousUsers removed")
                    Set-ReceiveConnector -Identity "$server\$defaultName" -PermissionGroups ($pgList -join ',') -ErrorAction Stop
                    Write-MyOutput ('Removed AnonymousUsers from "{0}" receive connector' -f $defaultName)
                }
                else {
                    Write-MyVerbose ('AnonymousUsers already absent from "{0}"' -f $defaultName)
                }
            }
            catch {
                Write-MyWarning ('Failed to modify Default Frontend connector: {0}' -f $_.Exception.Message)
            }
        }
        elseif ($onlyPlaceholders) {
            Write-MyWarning ('Default Frontend connector NOT modified — relay connector uses placeholder IPs only. Set real RemoteIPRanges first, then remove AnonymousUsers from "{0}" manually.' -f ('Default Frontend {0}' -f $server))
        }
        else {
            Write-MyWarning 'One or more relay connectors failed — Default Frontend connector was NOT modified'
        }
    }

    function Enable-AccessNamespaceMailConfig {
        # F26 — Configure the Access Namespace as an Accepted Domain and update the
        # default Email Address Policy so that mailboxes get a primary SMTP address
        # @<AccessNamespace> (e.g. @mail.contoso.com).
        #
        # Steps:
        #  1. Add the Access Namespace as an Authoritative Accepted Domain (skip if present).
        #  2. Update the default Email Address Policy to make @<namespace> the primary
        #     SMTP template; retain the AD/internal domain as a secondary address.
        #  3. Remove pure-internal domains (.local / nonroutable) from the policy
        #     templates — they serve no purpose as email addresses.
        #  4. Apply the updated policy to all mailboxes via Update-EmailAddressPolicy.
        #
        # Safety: the function is idempotent.  Running it a second time re-checks each
        # step and skips anything already in place.
        #
        if ($State['InstallEdge']) { Write-MyVerbose 'Enable-AccessNamespaceMailConfig: skipped (Edge Transport)'; return }
        if (-not $State['Namespace']) { Write-MyVerbose 'Enable-AccessNamespaceMailConfig: no namespace set — skipping'; return }

        # MailDomain is the root domain used for email addresses (e.g. contoso.com).
        # It defaults to the parent of the access namespace (drop leftmost label).
        # e.g. Namespace=outlook.domain.de → MailDomain=domain.de
        $ns = if ($State['MailDomain']) {
            $State['MailDomain']
        } else {
            $part = ($State['Namespace'] -split '\.', 2)[1]
            if ($part -match '\.') { $part } else { $State['Namespace'] }
        }

        Write-MyOutput ('Configuring access namespace mail settings — mail domain: {0}' -f $ns)

        # ── 1. Accepted Domain ──────────────────────────────────────────────────
        try {
            $existing = Get-AcceptedDomain -ErrorAction Stop | Where-Object { $_.DomainName -eq $ns }
            if ($existing) {
                Write-MyVerbose ('Accepted domain already present: {0} ({1})' -f $ns, $existing.DomainType)
            }
            else {
                New-AcceptedDomain -Name $ns -DomainName $ns -DomainType Authoritative -ErrorAction Stop | Out-Null
                Register-ExecutedCommand -Category 'ExchangePolicy' -Command ("New-AcceptedDomain -Name '{0}' -DomainName '{0}' -DomainType Authoritative" -f $ns)
                Write-MyOutput ('Accepted domain added: {0} (Authoritative)' -f $ns)
            }
        }
        catch {
            Write-MyWarning ('Could not create accepted domain {0}: {1}' -f $ns, $_.Exception.Message)
            return
        }

        # ── 2. Create a new Email Address Policy named after the mail domain ──────
        # A new policy is created rather than modifying the built-in Default Policy.
        # Name = mail domain (e.g. "contoso.com"); primary SMTP template = %m@domain.
        try {
            $policyName = $ns
            $nsTemplate = "SMTP:%m@$ns"   # uppercase SMTP = primary; %m = mailbox alias

            $existing = Get-EmailAddressPolicy -ErrorAction Stop | Where-Object { $_.Name -ieq $policyName }
            if ($existing) {
                $alreadyPrimary = (($existing.EnabledEmailAddressTemplates | Select-Object -First 1) -ieq $nsTemplate)
                if ($alreadyPrimary) {
                    Write-MyVerbose ("Email Address Policy '{0}' already configured correctly — no change needed" -f $policyName)
                } else {
                    Set-EmailAddressPolicy -Identity $existing.Identity `
                        -EnabledEmailAddressTemplates @($nsTemplate) -ErrorAction Stop
                    Register-ExecutedCommand -Category 'ExchangePolicy' `
                        -Command ("Set-EmailAddressPolicy -Identity '{0}' -EnabledEmailAddressTemplates @('{1}')" -f $policyName, $nsTemplate)
                    Write-MyOutput ("Email Address Policy '{0}' updated — primary SMTP: %m@{1}" -f $policyName, $ns)
                    Update-EmailAddressPolicy -Identity $existing.Identity -ErrorAction Stop
                    Register-ExecutedCommand -Category 'ExchangePolicy' -Command ("Update-EmailAddressPolicy -Identity '{0}'" -f $policyName)
                    Write-MyOutput 'Email Address Policy applied.'
                }
            } else {
                New-EmailAddressPolicy -Name $policyName -IncludedRecipients AllRecipients `
                    -EnabledEmailAddressTemplates @($nsTemplate) -Priority 1 `
                    -ErrorAction Stop | Out-Null
                Register-ExecutedCommand -Category 'ExchangePolicy' `
                    -Command ("New-EmailAddressPolicy -Name '{0}' -IncludedRecipients AllRecipients -EnabledEmailAddressTemplates @('{1}') -Priority 1" -f $policyName, $nsTemplate)
                Write-MyOutput ("Email Address Policy '{0}' created — primary SMTP: %m@{1}" -f $policyName, $ns)
                Update-EmailAddressPolicy -Identity $policyName -ErrorAction Stop
                Register-ExecutedCommand -Category 'ExchangePolicy' -Command ("Update-EmailAddressPolicy -Identity '{0}'" -f $policyName)
                Write-MyOutput 'Email Address Policy applied.'
            }
        }
        catch {
            Write-MyWarning ('Email Address Policy configuration failed: {0}' -f $_.Exception.Message)
        }
    }

    function Invoke-EOMT {
        if (-not $State['RunEOMT']) {
            Write-MyVerbose 'RunEOMT not specified, skipping EOMT'
            return
        }
        Write-MyOutput 'Running CSS-Exchange Emergency Mitigation Tool (EOMT)'
        $eomtPath = Join-Path $State['SourcesPath'] 'EOMT.ps1'
        $eomtUrl  = 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/EOMT.ps1'

        if (-not (Test-Path $eomtPath)) {
            $downloaded = $false
            $savedPP = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                try {
                    Write-MyVerbose ('Downloading EOMT from {0} (attempt {1}/3)' -f $eomtUrl, $attempt)
                    Start-BitsTransfer -Source $eomtUrl -Destination $eomtPath -ErrorAction Stop
                    $downloaded = $true
                    break
                }
                catch {
                    Get-BitsTransfer -ErrorAction SilentlyContinue | Where-Object { $_.JobState -notin 'Transferred','Acknowledged' } | Remove-BitsTransfer -ErrorAction SilentlyContinue
                    Remove-Item -Path $eomtPath -ErrorAction SilentlyContinue
                    if ($attempt -eq 3) {
                        try {
                            Invoke-WebDownload -Uri $eomtUrl -OutFile $eomtPath
                            $downloaded = $true
                        }
                        catch {
                            Write-MyWarning ('Could not download EOMT after 3 attempts: {0}' -f $_.Exception.Message)
                        }
                    }
                    else {
                        Start-Sleep -Seconds ($attempt * 5)
                    }
                }
            }
            $ProgressPreference = $savedPP
            if (-not $downloaded) { return }
        }

        if (Test-Path $eomtPath) {
            try {
                Write-MyVerbose ('EOMT SHA256: {0}' -f (Get-FileHash -Path $eomtPath -Algorithm SHA256).Hash)
                & $eomtPath
                Write-MyOutput 'EOMT completed successfully'
            }
            catch {
                Write-MyWarning ('EOMT execution failed: {0}' -f $_.Exception.Message)
            }
        }
    }

    function Set-HSTSHeader {
        if ($State['InstallEdge']) {
            Write-MyVerbose 'Edge role has no OWA/ECP — skipping HSTS configuration'
            return
        }
        Write-MyOutput 'Configuring HSTS (Strict-Transport-Security) for OWA and ECP'
        try {
            Import-Module WebAdministration -ErrorAction Stop
            $site = 'IIS:\Sites\Default Web Site'
            foreach ($vDir in @('owa', 'ecp')) {
                $path = '{0}\{1}' -f $site, $vDir
                if (-not (Test-Path $path)) {
                    Write-MyVerbose ('Virtual directory /{0} not found in IIS, skipping HSTS' -f $vDir)
                    continue
                }
                $filter   = 'system.webServer/httpProtocol/customHeaders/add[@name="Strict-Transport-Security"]'
                $existing = Get-WebConfigurationProperty -PSPath $path -Filter $filter -Name '.' -ErrorAction SilentlyContinue
                if ($existing) {
                    Write-MyVerbose ('HSTS header already present on /{0}' -f $vDir)
                }
                else {
                    Add-WebConfigurationProperty -PSPath $path -Filter 'system.webServer/httpProtocol/customHeaders' -Name '.' -Value @{ name = 'Strict-Transport-Security'; value = 'max-age=31536000' }
                    Write-MyOutput ('HSTS header configured on /{0} (max-age=31536000)' -f $vDir)
                }
            }
        }
        catch {
            Write-MyWarning ('Failed to configure HSTS: {0}' -f $_.Exception.Message)
        }
    }

    function Import-ExchangeCertificateFromPFX {
        if (-not $State['CertificatePath'] -or -not $State['CertificatePassword']) {
            Write-MyVerbose 'No certificate import requested'
            return
        }

        $pfxPath = $State['CertificatePath']
        if (-not (Test-Path $pfxPath)) {
            Write-MyError ('PFX file not found: {0}' -f $pfxPath)
            return
        }

        Write-MyOutput ('Importing certificate from {0}' -f $pfxPath)
        try {
            $secPwd = ConvertTo-SecureString $State['CertificatePassword']
            Register-ExecutedCommand -Category 'Certificate' -Command ("Import-ExchangeCertificate -FileData ([IO.File]::ReadAllBytes('$pfxPath')) -Password <SecureString> -PrivateKeyExportable `$true")
            $cert = Import-ExchangeCertificate -FileData ([System.IO.File]::ReadAllBytes($pfxPath)) -Password $secPwd -PrivateKeyExportable $true -ErrorAction Stop
            Write-MyOutput ('Certificate imported: {0} (Thumbprint: {1})' -f $cert.Subject, $cert.Thumbprint)

            # Detect wildcard certificate (CN=* or SAN with *.domain)
            $isWildcard = ($cert.Subject -match 'CN=\*') -or ($cert.SubjectAlternativeNames -match '^\*\.')
            if ($isWildcard) {
                # Wildcard: enable for IIS and SMTP only (IMAP/POP use specific SANs)
                Register-ExecutedCommand -Category 'Certificate' -Command ("Enable-ExchangeCertificate -Thumbprint '$($cert.Thumbprint)' -Services IIS,SMTP -Force")
                Enable-ExchangeCertificate -Thumbprint $cert.Thumbprint -Services IIS,SMTP -Force -ErrorAction Stop
                Write-MyOutput ('Wildcard certificate enabled for IIS and SMTP services')
            }
            else {
                # Named certificate: also enable for IMAP and POP
                Register-ExecutedCommand -Category 'Certificate' -Command ("Enable-ExchangeCertificate -Thumbprint '$($cert.Thumbprint)' -Services IIS,SMTP,IMAP,POP -Force")
                Enable-ExchangeCertificate -Thumbprint $cert.Thumbprint -Services IIS,SMTP,IMAP,POP -Force -ErrorAction Stop
                Write-MyOutput ('Certificate enabled for IIS, SMTP, IMAP and POP services')
            }
        }
        catch {
            Write-MyError ('Failed to import/enable certificate: {0}' -f $_.Exception.Message)
        }
    }

    function Set-VirtualDirectoryURLs {
        if (-not $State['Namespace']) {
            Write-MyVerbose 'No Namespace specified, skipping Virtual Directory URL configuration'
            return
        }

        $ns     = $State['Namespace']
        $server = $env:COMPUTERNAME
        $errors = 0
        $changed = 0
        Write-MyOutput ('Configuring Virtual Directory URLs for namespace: {0}' -f $ns)

        # Exchange VDir cmdlets call ShouldContinue("host can't be resolved") when the namespace
        # doesn't resolve in DNS — ShouldContinue cannot be suppressed by -Confirm:$false or
        # preference variables. Add a temporary hosts entry if needed and remove it afterwards.
        $hostsFile      = "$env:SystemRoot\System32\drivers\etc\hosts"
        $tempHostsMark  = '# EXpress-temp-vdir'
        $hostsBackup    = $null
        $dlDomain       = $State['DownloadDomain']
        $nsResolves     = $false
        $dlResolves     = $false
        try { [System.Net.Dns]::GetHostEntry($ns) | Out-Null; $nsResolves = $true } catch { }
        if ($dlDomain) { try { [System.Net.Dns]::GetHostEntry($dlDomain) | Out-Null; $dlResolves = $true } catch { } }
        if (-not $nsResolves) {
            Write-MyVerbose ('Namespace {0} not resolvable — adding temporary hosts entry to suppress VDir confirmation prompt' -f $ns)
            $hostsBackup = [System.IO.File]::ReadAllBytes($hostsFile)
            "`r`n127.0.0.1`t$ns`t$tempHostsMark" | Add-Content -Path $hostsFile -Encoding ASCII -ErrorAction SilentlyContinue
        }
        if ($dlDomain -and -not $dlResolves) {
            Write-MyVerbose ('Download domain {0} not resolvable — adding temporary hosts entry' -f $dlDomain)
            if (-not $hostsBackup) { $hostsBackup = [System.IO.File]::ReadAllBytes($hostsFile) }
            "`r`n127.0.0.1`t$dlDomain`t$tempHostsMark" | Add-Content -Path $hostsFile -Encoding ASCII -ErrorAction SilentlyContinue
        }

        # Helper: compare a vdir URL property (Uri object or string) to a target string
        function Test-VdirUrl($current, $target) {
            if (-not $current) { return $false }
            return ([string]$current -eq $target)
        }

        # OWA — set URL and UPN logon format
        try {
            $vd = Get-OwaVirtualDirectory -Identity "$server\owa (Default Web Site)" -ADPropertiesOnly -ErrorAction Stop
            $urlOk    = (Test-VdirUrl $vd.InternalUrl "https://$ns/owa") -and (Test-VdirUrl $vd.ExternalUrl "https://$ns/owa")
            $formatOk = ([string]$vd.LogonFormat -eq 'PrincipalName')
            if ($urlOk -and $formatOk) {
                Write-MyVerbose 'OWA: URLs and logon format already set, skipping'
            } else {
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-OwaVirtualDirectory -Identity '$server\owa (Default Web Site)' -InternalUrl 'https://$ns/owa' -ExternalUrl 'https://$ns/owa' -LogonFormat PrincipalName -DefaultDomain ''")
                Set-OwaVirtualDirectory -Identity "$server\owa (Default Web Site)" `
                    -InternalUrl "https://$ns/owa" -ExternalUrl "https://$ns/owa" `
                    -LogonFormat PrincipalName -DefaultDomain '' `
                    -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
                Write-MyVerbose 'OWA virtual directory configured (UPN logon)'
                $changed++
            }
        }
        catch { Write-MyWarning ('OWA: {0}' -f $_.Exception.Message); $errors++ }

        # OWA Download Domains — CVE-2021-1730 mitigation (isolates attachment downloads to a separate hostname)
        if ($dlDomain) {
            try {
                $vd = Get-OwaVirtualDirectory -Identity "$server\owa (Default Web Site)" -ADPropertiesOnly -ErrorAction Stop
                $dlOk = ([string]$vd.ExternalDownloadHostName -eq $dlDomain) -and ([string]$vd.InternalDownloadHostName -eq $dlDomain)
                if ($dlOk) {
                    Write-MyVerbose ('OWA Download Domains already set to {0}, skipping' -f $dlDomain)
                } else {
                    Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-OwaVirtualDirectory -Identity '$server\owa (Default Web Site)' -ExternalDownloadHostName '$dlDomain' -InternalDownloadHostName '$dlDomain'")
                    Set-OwaVirtualDirectory -Identity "$server\owa (Default Web Site)" `
                        -ExternalDownloadHostName $dlDomain -InternalDownloadHostName $dlDomain `
                        -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
                    Write-MyVerbose ('OWA Download Domains configured: {0} (CVE-2021-1730 mitigation)' -f $dlDomain)
                    $changed++
                }
            }
            catch { Write-MyWarning ('OWA Download Domains: {0}' -f $_.Exception.Message); $errors++ }
            # EnableDownloadDomains must be set at org level for CVE-2021-1730 mitigation to take effect
            try {
                $tc = Get-OrganizationConfig -ErrorAction Stop
                if (-not $tc.EnableDownloadDomains) {
                    Register-ExecutedCommand -Category 'VirtualDirectories' -Command 'Set-OrganizationConfig -EnableDownloadDomains $true'
                    Set-OrganizationConfig -EnableDownloadDomains $true -ErrorAction Stop
                    Write-MyVerbose 'EnableDownloadDomains enabled at org level (CVE-2021-1730)'
                    $changed++
                } else {
                    Write-MyVerbose 'EnableDownloadDomains already enabled at org level'
                }
            }
            catch { Write-MyWarning ('EnableDownloadDomains: {0}' -f $_.Exception.Message); $errors++ }
        }

        # ECP
        try {
            $vd = Get-EcpVirtualDirectory -Identity "$server\ecp (Default Web Site)" -ADPropertiesOnly -ErrorAction Stop
            if ((Test-VdirUrl $vd.InternalUrl "https://$ns/ecp") -and (Test-VdirUrl $vd.ExternalUrl "https://$ns/ecp")) {
                Write-MyVerbose 'ECP: URLs already set, skipping'
            } else {
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-EcpVirtualDirectory -Identity '$server\ecp (Default Web Site)' -InternalUrl 'https://$ns/ecp' -ExternalUrl 'https://$ns/ecp'")
                Set-EcpVirtualDirectory -Identity "$server\ecp (Default Web Site)" `
                    -InternalUrl "https://$ns/ecp" -ExternalUrl "https://$ns/ecp" `
                    -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
                Write-MyVerbose 'ECP virtual directory configured'
                $changed++
            }
        }
        catch { Write-MyWarning ('ECP: {0}' -f $_.Exception.Message); $errors++ }

        # EWS
        try {
            $vd = Get-WebServicesVirtualDirectory -Identity "$server\EWS (Default Web Site)" -ADPropertiesOnly -ErrorAction Stop
            if ((Test-VdirUrl $vd.InternalUrl "https://$ns/EWS/Exchange.asmx") -and (Test-VdirUrl $vd.ExternalUrl "https://$ns/EWS/Exchange.asmx")) {
                Write-MyVerbose 'EWS: URLs already set, skipping'
            } else {
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-WebServicesVirtualDirectory -Identity '$server\EWS (Default Web Site)' -InternalUrl 'https://$ns/EWS/Exchange.asmx' -ExternalUrl 'https://$ns/EWS/Exchange.asmx'")
                Set-WebServicesVirtualDirectory -Identity "$server\EWS (Default Web Site)" `
                    -InternalUrl "https://$ns/EWS/Exchange.asmx" -ExternalUrl "https://$ns/EWS/Exchange.asmx" `
                    -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
                Write-MyVerbose 'EWS virtual directory configured'
                $changed++
            }
        }
        catch { Write-MyWarning ('EWS: {0}' -f $_.Exception.Message); $errors++ }

        # OAB
        try {
            $vd = Get-OabVirtualDirectory -Identity "$server\OAB (Default Web Site)" -ADPropertiesOnly -ErrorAction Stop
            if ((Test-VdirUrl $vd.InternalUrl "https://$ns/OAB") -and (Test-VdirUrl $vd.ExternalUrl "https://$ns/OAB")) {
                Write-MyVerbose 'OAB: URLs already set, skipping'
            } else {
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-OabVirtualDirectory -Identity '$server\OAB (Default Web Site)' -InternalUrl 'https://$ns/OAB' -ExternalUrl 'https://$ns/OAB'")
                Set-OabVirtualDirectory -Identity "$server\OAB (Default Web Site)" `
                    -InternalUrl "https://$ns/OAB" -ExternalUrl "https://$ns/OAB" `
                    -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
                Write-MyVerbose 'OAB virtual directory configured'
                $changed++
            }
        }
        catch { Write-MyWarning ('OAB: {0}' -f $_.Exception.Message); $errors++ }

        # ActiveSync
        try {
            $vd = Get-ActiveSyncVirtualDirectory -Identity "$server\Microsoft-Server-ActiveSync (Default Web Site)" -ADPropertiesOnly -ErrorAction Stop
            if ((Test-VdirUrl $vd.InternalUrl "https://$ns/Microsoft-Server-ActiveSync") -and (Test-VdirUrl $vd.ExternalUrl "https://$ns/Microsoft-Server-ActiveSync")) {
                Write-MyVerbose 'ActiveSync: URLs already set, skipping'
            } else {
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-ActiveSyncVirtualDirectory -Identity '$server\Microsoft-Server-ActiveSync (Default Web Site)' -InternalUrl 'https://$ns/Microsoft-Server-ActiveSync' -ExternalUrl 'https://$ns/Microsoft-Server-ActiveSync'")
                Set-ActiveSyncVirtualDirectory -Identity "$server\Microsoft-Server-ActiveSync (Default Web Site)" `
                    -InternalUrl "https://$ns/Microsoft-Server-ActiveSync" -ExternalUrl "https://$ns/Microsoft-Server-ActiveSync" `
                    -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
                Write-MyVerbose 'ActiveSync virtual directory configured'
                $changed++
            }
        }
        catch { Write-MyWarning ('ActiveSync: {0}' -f $_.Exception.Message); $errors++ }

        # MAPI — URL first, auth methods in a separate try (not available on all builds)
        try {
            $vd = Get-MapiVirtualDirectory -Identity "$server\mapi (Default Web Site)" -ADPropertiesOnly -ErrorAction Stop
            if ((Test-VdirUrl $vd.InternalUrl "https://$ns/mapi") -and (Test-VdirUrl $vd.ExternalUrl "https://$ns/mapi")) {
                Write-MyVerbose 'MAPI: URLs already set, skipping'
            } else {
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-MapiVirtualDirectory -Identity '$server\mapi (Default Web Site)' -InternalUrl 'https://$ns/mapi' -ExternalUrl 'https://$ns/mapi'")
                Set-MapiVirtualDirectory -Identity "$server\mapi (Default Web Site)" `
                    -InternalUrl "https://$ns/mapi" -ExternalUrl "https://$ns/mapi" `
                    -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
                Write-MyVerbose 'MAPI virtual directory URL configured'
                $changed++
            }
        }
        catch { Write-MyWarning ('MAPI URL: {0}' -f $_.Exception.Message); $errors++ }

        try {
            Set-MapiVirtualDirectory -Identity "$server\mapi (Default Web Site)" `
                -InternalAuthenticationMethods NTLM,Negotiate,OAuth `
                -ExternalAuthenticationMethods NTLM,Negotiate,OAuth `
                -ErrorAction Stop -WarningAction SilentlyContinue
            Write-MyVerbose 'MAPI authentication methods configured'
            Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-MapiVirtualDirectory -Identity '$server\mapi (Default Web Site)' -InternalAuthenticationMethods NTLM,Negotiate,OAuth -ExternalAuthenticationMethods NTLM,Negotiate,OAuth")
        }
        catch { Write-MyVerbose ('MAPI auth methods not supported on this build: {0}' -f $_.Exception.Message) }

        # PowerShell — ExternalUrl only; InternalUrl stays http (Exchange internal services use http by default)
        try {
            $vd = Get-PowerShellVirtualDirectory -Identity "$server\PowerShell (Default Web Site)" -ADPropertiesOnly -ErrorAction Stop
            if (Test-VdirUrl $vd.ExternalUrl "https://$ns/powershell") {
                Write-MyVerbose 'PowerShell: ExternalUrl already set, skipping'
            } else {
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-PowerShellVirtualDirectory -Identity '$server\PowerShell (Default Web Site)' -ExternalUrl 'https://$ns/powershell'")
                Set-PowerShellVirtualDirectory -Identity "$server\PowerShell (Default Web Site)" `
                    -ExternalUrl "https://$ns/powershell" `
                    -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
                Write-MyVerbose 'PowerShell virtual directory ExternalUrl configured'
                $changed++
            }
        }
        catch { Write-MyWarning ('PowerShell URL: {0}' -f $_.Exception.Message); $errors++ }

        # Autodiscover SCP — always use autodiscover.<parent-domain>, not the namespace hostname
        try {
            $cas = Get-ClientAccessService -Identity $server -ErrorAction Stop
            $nsParts   = $ns -split '\.'
            $scpHost   = if ($nsParts[0] -eq 'autodiscover') { $ns } else { 'autodiscover.' + ($nsParts[1..($nsParts.Length-1)] -join '.') }
            $scpTarget = "https://$scpHost/Autodiscover/Autodiscover.xml"
            if ([string]$cas.AutoDiscoverServiceInternalUri -eq $scpTarget) {
                Write-MyVerbose 'Autodiscover SCP: already set, skipping'
            } else {
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-ClientAccessService -Identity '$server' -AutoDiscoverServiceInternalUri '$scpTarget'")
                Set-ClientAccessService -Identity $server `
                    -AutoDiscoverServiceInternalUri $scpTarget `
                    -ErrorAction Stop -WarningAction SilentlyContinue
                Write-MyVerbose 'Autodiscover SCP configured'
                $changed++
            }
        }
        catch { Write-MyWarning ('Autodiscover SCP: {0}' -f $_.Exception.Message); $errors++ }

        # Restore hosts file to exact pre-modification state using the binary backup.
        if ($hostsBackup) {
            try {
                [System.IO.File]::WriteAllBytes($hostsFile, $hostsBackup)
                Write-MyVerbose 'Temporary hosts entries removed (hosts file restored from backup)'
            }
            catch { Write-MyVerbose ('Could not restore hosts file: {0}' -f $_.Exception.Message) }
        }

        if ($errors -eq 0) {
            if ($changed -gt 0) {
                Write-MyOutput ('Virtual Directory URLs configured for https://{0} (OWA logon: UPN)' -f $ns)
            } else {
                Write-MyOutput ('Virtual Directory URLs already correct for https://{0} — no changes made' -f $ns)
            }
        }
        else {
            Write-MyWarning ('{0} virtual directory(s) could not be configured — check warnings above' -f $errors)
        }
    }

    function Join-DAG {
        if (-not $State['DAGName']) {
            return
        }

        Write-MyOutput ('Joining Database Availability Group: {0}' -f $State['DAGName'])

        # Ensure Exchange module is loaded
        Import-ExchangeModule

        try {
            $dag = Get-DatabaseAvailabilityGroup -Identity $State['DAGName'] -ErrorAction Stop
            if ($null -eq $dag) {
                Write-MyError ('DAG {0} not found' -f $State['DAGName'])
                exit $ERR_DAGJOIN
            }
            if ($dag.Servers -contains $env:COMPUTERNAME) {
                Write-MyOutput ('Server {0} is already a member of DAG {1}' -f $env:COMPUTERNAME, $State['DAGName'])
                return
            }

            Register-ExecutedCommand -Category 'DAG' -Command ("Add-DatabaseAvailabilityGroupServer -Identity '$($State['DAGName'])' -MailboxServer '$env:COMPUTERNAME'")
            Add-DatabaseAvailabilityGroupServer -Identity $State['DAGName'] -MailboxServer $env:COMPUTERNAME -ErrorAction Stop
            Write-MyOutput ('Successfully joined DAG {0}' -f $State['DAGName'])
        }
        catch {
            Write-MyError ('Failed to join DAG {0}: {1}' -f $State['DAGName'], $_.Exception.Message)
            exit $ERR_DAGJOIN
        }
    }

    function Invoke-HealthChecker {
        if ($State['SkipHealthCheck']) {
            Write-MyVerbose 'SkipHealthCheck specified, skipping HealthChecker'
            return
        }

        Write-MyOutput 'Running CSS-Exchange HealthChecker'
        $hcPath = Join-Path $State['SourcesPath'] 'HealthChecker.ps1'
        $hcUrl = 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/HealthChecker.ps1'

        # Download if not present
        if (-not (Test-Path $hcPath)) {
            $downloaded = $false
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                try {
                    Write-MyVerbose ('Downloading HealthChecker from {0} (attempt {1}/3)' -f $hcUrl, $attempt)
                    Start-BitsTransfer -Source $hcUrl -Destination $hcPath -ErrorAction Stop
                    $downloaded = $true
                    break
                }
                catch {
                    if ($attempt -eq 3) {
                        try {
                            Invoke-WebDownload -Uri $hcUrl -OutFile $hcPath
                            $downloaded = $true
                        }
                        catch {
                            Write-MyWarning ('Could not download HealthChecker after 3 attempts: {0}' -f $_.Exception.Message)
                        }
                    }
                    else {
                        Start-Sleep -Seconds ($attempt * 5)
                    }
                }
            }
            if ($downloaded -and (Test-Path $hcPath)) {
                $hash = (Get-FileHash -Path $hcPath -Algorithm SHA256).Hash
                Write-MyVerbose ('HealthChecker downloaded, SHA256: {0}' -f $hash)
            }
            elseif (-not $downloaded) {
                return
            }
        }

        if (Test-Path $hcPath) {
            try {
                # HC writes ExchangeAllServersReport-*.html to the *current directory*, not -OutputFilePath.
                # Push-Location so both the XML (-OutputFilePath) and the HTML land in ReportsPath.
                Push-Location $State['ReportsPath']
                $hcBefore = [datetime]::Now
                & $hcPath -OutputFilePath $State['ReportsPath'] -SkipVersionCheck *>&1 | Out-Null
                & $hcPath -BuildHtmlServersReport -SkipVersionCheck *>&1 | Out-Null
                Pop-Location
                $hcReport = Get-ChildItem -Path $State['ReportsPath'] -ErrorAction SilentlyContinue |
                    Where-Object { $_.LastWriteTime -ge $hcBefore -and $_.Extension -match '\.html?' -and $_.Name -match '^(ExchangeAllServersReport|HealthChecker|HCExchangeServerReport)' } |
                    Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($hcReport) {
                    # Rename to SERVER_HCExchangeServerReport-<timestamp>.html
                    $hcTimestamp = $hcReport.Name -replace '^(?:ExchangeAllServersReport|HealthChecker|HCExchangeServerReport)-', ''
                    $newHcName   = '{0}_HCExchangeServerReport-{1}' -f $env:COMPUTERNAME, $hcTimestamp
                    $newHcPath   = Join-Path $State['ReportsPath'] $newHcName
                    try {
                        Rename-Item -Path $hcReport.FullName -NewName $newHcName -ErrorAction Stop
                        $State['HCReportPath'] = $newHcPath
                        Write-MyOutput ('HealthChecker report saved to {0}' -f $newHcPath)
                    }
                    catch {
                        $State['HCReportPath'] = $hcReport.FullName
                        Write-MyOutput ('HealthChecker report saved to {0}' -f $hcReport.FullName)
                    }
                } else {
                    Write-MyOutput ('HealthChecker completed — report in {0}' -f $State['ReportsPath'])
                }
                # On Domain Controllers there are no local security groups — the SAM database is
                # replaced by AD. HC's "Exchange Server Membership" check enumerates Win32_GroupUser
                # for the local "Exchange Servers" / "Exchange Trusted Subsystem" groups, which don't
                # exist on DCs, so it always reports "failed/blank" regardless of AD group membership.
                # This is a HC limitation; the server IS a member via the domain group.
                $dcRole = try { (Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue).DomainRole } catch { 3 }
                if ($dcRole -ge 4) {
                    Write-MyWarning 'NOTE: This server is a Domain Controller. HC "Exchange Server Membership" will show failed/blank — DCs have no local security groups. Exchange group membership is via AD domain groups and is correct.'
                }
            }
            catch {
                Pop-Location -ErrorAction SilentlyContinue
                Write-MyWarning ('HealthChecker execution failed: {0}' -f $_.Exception.Message)
            }
        }
    }

    function Invoke-SetupAssist {
        if ($State['SkipSetupAssist']) {
            Write-MyVerbose 'SkipSetupAssist specified, skipping SetupAssist'
            return
        }

        Write-MyOutput 'Running CSS-Exchange SetupAssist to diagnose setup failure'
        $saPath = Join-Path $State['SourcesPath'] 'SetupAssist.ps1'
        $saUrl  = 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/SetupAssist.ps1'

        if (-not (Test-Path $saPath)) {
            $downloaded = $false
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                try {
                    Write-MyVerbose ('Downloading SetupAssist from {0} (attempt {1}/3)' -f $saUrl, $attempt)
                    Start-BitsTransfer -Source $saUrl -Destination $saPath -ErrorAction Stop
                    $downloaded = $true
                    break
                }
                catch {
                    if ($attempt -eq 3) {
                        try {
                            Invoke-WebDownload -Uri $saUrl -OutFile $saPath
                            $downloaded = $true
                        }
                        catch {
                            Write-MyWarning ('Could not download SetupAssist after 3 attempts: {0}' -f $_.Exception.Message)
                        }
                    }
                    else {
                        Start-Sleep -Seconds ($attempt * 5)
                    }
                }
            }
            if ($downloaded -and (Test-Path $saPath)) {
                Write-MyVerbose ('SetupAssist downloaded, SHA256: {0}' -f (Get-FileHash -Path $saPath -Algorithm SHA256).Hash)
            }
            elseif (-not $downloaded) {
                return
            }
        }

        if (Test-Path $saPath) {
            try {
                & $saPath
            }
            catch {
                Write-MyWarning ('SetupAssist execution failed: {0}' -f $_.Exception.Message)
            }
        }

        # SetupLogReviewer — additional log analysis tool
        $slrPath = Join-Path $State['SourcesPath'] 'SetupLogReviewer.ps1'
        $slrUrl  = 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/SetupLogReviewer.ps1'

        if (-not (Test-Path $slrPath)) {
            $downloaded = $false
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                try {
                    Write-MyVerbose ('Downloading SetupLogReviewer from {0} (attempt {1}/3)' -f $slrUrl, $attempt)
                    Start-BitsTransfer -Source $slrUrl -Destination $slrPath -ErrorAction Stop
                    $downloaded = $true
                    break
                }
                catch {
                    if ($attempt -eq 3) {
                        try {
                            Invoke-WebDownload -Uri $slrUrl -OutFile $slrPath
                            $downloaded = $true
                        }
                        catch {
                            Write-MyWarning ('Could not download SetupLogReviewer after 3 attempts: {0}' -f $_.Exception.Message)
                        }
                    }
                    else {
                        Start-Sleep -Seconds ($attempt * 5)
                    }
                }
            }
            if ($downloaded -and (Test-Path $slrPath)) {
                Write-MyVerbose ('SetupLogReviewer downloaded, SHA256: {0}' -f (Get-FileHash -Path $slrPath -Algorithm SHA256).Hash)
            }
        }

        if (Test-Path $slrPath) {
            try {
                Write-MyOutput 'Running CSS-Exchange SetupLogReviewer to analyze setup logs'
                & $slrPath
            }
            catch {
                Write-MyWarning ('SetupLogReviewer execution failed: {0}' -f $_.Exception.Message)
            }
        }
    }

    function Test-AuthCertificate {
        try {
            $authConfig = Get-AuthConfig -ErrorAction Stop
            if (-not $authConfig) {
                Write-MyVerbose 'Test-AuthCertificate: Get-AuthConfig returned null — Exchange PS session may not be fully initialized'
                return
            }
            $thumbprint = $authConfig.CurrentCertificateThumbprint
            if (-not $thumbprint) {
                Write-MyWarning 'Exchange Auth Certificate: no thumbprint configured in AuthConfig'
                return
            }
            $cert = Get-ExchangeCertificate -Thumbprint $thumbprint -ErrorAction SilentlyContinue
            if (-not $cert) {
                Write-MyWarning ('Exchange Auth Certificate (thumbprint {0}) not found on this server' -f $thumbprint)
                return
            }
            $daysLeft = ($cert.NotAfter - (Get-Date)).Days
            if ($daysLeft -le 0) {
                Write-MyWarning ('Exchange Auth Certificate EXPIRED {0} day(s) ago (expires {1}, thumbprint {2}). Renew: New-ExchangeCertificate, then Set-AuthConfig -NewCertificateThumbprint / -PublishCertificate' -f [Math]::Abs($daysLeft), $cert.NotAfter.ToString('yyyy-MM-dd'), $thumbprint)
            }
            elseif ($daysLeft -le 60) {
                Write-MyWarning ('Exchange Auth Certificate expires in {0} days on {1} (thumbprint {2}). Renew soon: New-ExchangeCertificate, then Set-AuthConfig -NewCertificateThumbprint / -PublishCertificate' -f $daysLeft, $cert.NotAfter.ToString('yyyy-MM-dd'), $thumbprint)
            }
            else {
                Write-MyOutput ('Exchange Auth Certificate valid for {0} days (expires {1}, thumbprint {2})' -f $daysLeft, $cert.NotAfter.ToString('yyyy-MM-dd'), $thumbprint)
            }
        }
        catch {
            Write-MyVerbose ('Test-AuthCertificate: {0}' -f $_.Exception.Message)
        }
    }

    function Test-DAGReplicationHealth {
        # F8: Validates mailbox database copy replication after DAG join.
        if (-not $State['DAGName']) { Write-MyVerbose 'No DAG configured, skipping replication health check'; return }
        Write-MyOutput ('Checking DAG database copy replication health on {0}' -f $env:computername)
        try {
            $copies = @(Get-MailboxDatabaseCopyStatus -Server $env:computername -ErrorAction Stop)
            if ($copies.Count -eq 0) { Write-MyVerbose 'No mailbox database copies found on this server'; return }
            $warns = 0
            foreach ($copy in $copies) {
                $ok  = $copy.Status -in 'Mounted', 'Healthy'
                $msg = 'DB copy {0}: Status={1}, CopyQueue={2}, ReplayQueue={3}' -f $copy.DatabaseName, $copy.Status, $copy.CopyQueueLength, $copy.ReplayQueueLength
                if ($ok) { Write-MyVerbose $msg } else { Write-MyWarning $msg; $warns++ }
            }
            if ($warns -eq 0) {
                Write-MyOutput ('DAG replication health: {0} copy/copies OK' -f $copies.Count)
            }
            else {
                Write-MyWarning ('{0} database copy/copies not healthy — review replication status' -f $warns)
            }
        }
        catch {
            Write-MyWarning ('DAG replication health check failed: {0}' -f $_.Exception.Message)
        }
    }

    function Test-VSSWriters {
        # F9: Checks all VSS writers are in a stable state. Unstable writers can break Exchange online backup.
        Write-MyOutput 'Checking VSS writer health'
        try {
            $output = & vssadmin.exe list writers 2>&1
            $currentWriter = ''
            $warns = 0
            foreach ($line in $output) {
                if ($line -match "Writer name:\s+'(.+)'") { $currentWriter = $Matches[1] }
                elseif ($line -match 'State:\s*\[\d+\]\s+(.+)') {
                    $stateText = $Matches[1].Trim()
                    if ($stateText -notmatch '^Stable') {
                        Write-MyWarning ('VSS Writer "{0}": {1}' -f $currentWriter, $stateText)
                        $warns++
                    }
                }
            }
            if ($warns -eq 0) { Write-MyVerbose 'All VSS writers are stable' }
            else { Write-MyWarning ('{0} VSS writer(s) not stable — check Volume Shadow Copy Service' -f $warns) }
        }
        catch {
            Write-MyWarning ('VSS writer check failed: {0}' -f $_.Exception.Message)
        }
    }

    function Test-EEMSStatus {
        # F10: Exchange Emergency Mitigation Service (EEMS) — available from Exchange 2019 CU11+ and SE.
        # EEMS applies automatic security mitigations for critical CVEs before patches are available.
        $svc = Get-Service MSExchangeMitigation -ErrorAction SilentlyContinue
        if (-not $svc) { Write-MyVerbose 'EEMS service not present (Exchange 2016 or 2019 pre-CU11)'; return }
        $statusLabel = if ($svc.Status -eq 'Running') { 'Running (OK)' } else { $svc.Status.ToString() }
        Write-MyOutput ('Exchange Emergency Mitigation Service (EEMS): {0}' -f $statusLabel)
        if ($svc.Status -ne 'Running') {
            Write-MyWarning 'EEMS is not running — automatic CVE mitigations will not be applied'
        }
        try {
            $orgCfg = Get-OrganizationConfig -ErrorAction Stop
            if ($orgCfg.PSObject.Properties['MitigationsEnabled']) {
                if (-not $orgCfg.MitigationsEnabled) {
                    Write-MyWarning 'EEMS mitigations disabled org-wide (Set-OrganizationConfig -MitigationsEnabled $true to re-enable)'
                }
                else {
                    Write-MyVerbose ('EEMS mitigations enabled: {0}' -f $orgCfg.MitigationsEnabled)
                }
                $blocked = $orgCfg.MitigationsBlocked
                if ($blocked) {
                    Write-MyWarning ('EEMS blocked mitigations: {0}' -f ($blocked -join ', '))
                }
            }
        }
        catch {
            Write-MyVerbose ('EEMS org config check: {0}' -f $_.Exception.Message)
        }
    }

    function Get-ModernAuthReport {
        # F11: Verifies Modern Authentication (OAuth2) is enabled. Required for Outlook 2016+,
        # Microsoft Teams, mobile clients, and any Hybrid / Azure AD configuration.
        Write-MyOutput 'Checking Modern Authentication (OAuth2) configuration'
        try {
            $orgCfg = Get-OrganizationConfig -ErrorAction Stop
            if ($orgCfg.OAuth2ClientProfileEnabled) {
                Write-MyVerbose 'Modern Authentication (OAuth2): Enabled (OK)'
            }
            else {
                Write-MyWarning 'Modern Authentication (OAuth2) is DISABLED — required for Outlook 2016+, Teams, mobile clients, and Hybrid. Enable: Set-OrganizationConfig -OAuth2ClientProfileEnabled $true'
            }
        }
        catch {
            Write-MyVerbose ('Modern Auth report: {0}' -f $_.Exception.Message)
        }
    }

    function Get-RemoteServerData {
        <#
        .SYNOPSIS
            Collects hardware/OS/pagefile/volume/NIC data from a remote Exchange server via CIM/WSMan.
        .DESCRIPTION
            Uses CIM over WSMan (WinRM TCP 5985/5986, Kerberos). NOT WMI/DCOM.
            Returns a uniform hashtable; on failure sets Reachable = $false with Error text.
            Timeout 30 s; always disposes CimSession in finally.
            Pre-requisites on target: see tools\Enable-EXpressRemoteQuery.ps1 or docs\remote-query-setup.md.
        #>
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][string]$ComputerName,
            [int]$TimeoutSec = 30
        )

        $result = @{
            ComputerName = $ComputerName
            Reachable    = $false
            Error        = $null
            OS           = $null
            CPU          = $null
            ComputerSys  = $null
            PageFile     = $null
            Volumes      = @()
            NICs         = @()
        }

        $session = $null
        try {
            $opt = New-CimSessionOption -Protocol WSMan
            $session = New-CimSession -ComputerName $ComputerName -SessionOption $opt `
                                      -OperationTimeoutSec $TimeoutSec -ErrorAction Stop

            $result.OS          = Get-CimInstance -CimSession $session -ClassName Win32_OperatingSystem          -ErrorAction Stop
            $result.CPU         = Get-CimInstance -CimSession $session -ClassName Win32_Processor                -ErrorAction Stop
            $result.ComputerSys = Get-CimInstance -CimSession $session -ClassName Win32_ComputerSystem           -ErrorAction Stop
            $result.PageFile    = Get-CimInstance -CimSession $session -ClassName Win32_PageFileSetting          -ErrorAction SilentlyContinue
            $result.Volumes     = @(Get-CimInstance -CimSession $session -ClassName Win32_Volume -Filter 'DriveType=3' -ErrorAction SilentlyContinue)
            $result.NICs        = @(Get-CimInstance -CimSession $session -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=TRUE' -ErrorAction SilentlyContinue)
            $result.Reachable   = $true
        }
        catch {
            $result.Error = $_.Exception.Message
            Write-MyVerbose ('Get-RemoteServerData {0}: {1}' -f $ComputerName, $_.Exception.Message)
        }
        finally {
            if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
        }

        return $result
    }

    function Invoke-RemoteQueryWithPrompt {
        <#
        .SYNOPSIS
            Wraps Get-RemoteServerData with interactive retry/skip prompt on failure.
        .DESCRIPTION
            Copilot (interactive) mode: on failure, shows hint pointing to Enable-EXpressRemoteQuery.ps1
            and offers [R]etry / [S]kip with a 10-minute auto-skip timeout (Write-Progress -Id 2).
            Autopilot mode or non-interactive session: silent skip.
        #>
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][string]$ComputerName,
            [int]$TimeoutSec = 600
        )

        $data = Get-RemoteServerData -ComputerName $ComputerName
        if ($data.Reachable) { return $data }

        $nonInteractive = $State['Autopilot'] -or -not [Environment]::UserInteractive
        if ($nonInteractive) {
            Write-MyVerbose ('Remote query skipped (non-interactive) for {0}: {1}' -f $ComputerName, $data.Error)
            return $data
        }

        while (-not $data.Reachable) {
            Write-Host ''
            Write-MyWarning ('Remote query failed for {0}' -f $ComputerName)
            Write-Host ('    Error : {0}' -f $data.Error) -ForegroundColor Yellow
            Write-Host  '    Fix   : Run tools\Enable-EXpressRemoteQuery.ps1 on the target server,' -ForegroundColor Yellow
            Write-Host  '            or apply GPO per docs\remote-query-setup.md' -ForegroundColor Yellow
            Write-Host ''
            Write-Host '    [R] Retry    [S] Skip    (auto-skip in 10:00)' -ForegroundColor Cyan

            $choice = $null
            try { $host.UI.RawUI.FlushInputBuffer() } catch { }
            $deadline = [DateTime]::Now.AddSeconds($TimeoutSec)
            while ([DateTime]::Now -lt $deadline -and -not $choice) {
                $secsLeft = [int]($deadline - [DateTime]::Now).TotalSeconds
                $mm = [int]($secsLeft / 60); $ss = $secsLeft % 60
                Write-Progress -Id 2 -Activity ('Remote query: {0}' -f $ComputerName) `
                    -Status ('Auto-skip in {0:D2}:{1:D2}  |  [R] Retry  |  [S] Skip' -f $mm, $ss) `
                    -PercentComplete (($TimeoutSec - $secsLeft) * 100 / $TimeoutSec)
                if ($host.UI.RawUI.KeyAvailable) {
                    $key = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                    switch ($key.Character.ToString().ToUpper()) {
                        'R' { $choice = 'Retry' }
                        'S' { $choice = 'Skip'  }
                    }
                }
                Start-Sleep -Milliseconds 100
            }
            Write-Progress -Id 2 -Activity ('Remote query: {0}' -f $ComputerName) -Completed

            if (-not $choice) {
                Write-MyOutput ('Auto-skip: {0}' -f $ComputerName)
                return $data
            }
            if ($choice -eq 'Skip') {
                Write-MyOutput ('Skipped remote query for {0}' -f $ComputerName)
                return $data
            }

            Write-MyOutput ('Retrying remote query for {0}...' -f $ComputerName)
            $data = Get-RemoteServerData -ComputerName $ComputerName
        }

        return $data
    }

    function Get-OrganizationReportData {
        # Collects org-wide Exchange settings. No server-specific data here.
        # Safe to call from New-InstallationDocument in all scenarios.
        $org = @{}

        # Org config
        try { $org.OrgConfig = Get-OrganizationConfig -ErrorAction Stop } catch { $org.OrgConfig = $null }

        # Accepted / Remote Domains
        try { $org.AcceptedDomains = @(Get-AcceptedDomain -ErrorAction Stop) } catch { $org.AcceptedDomains = @() }
        try { $org.RemoteDomains   = @(Get-RemoteDomain   -ErrorAction Stop) } catch { $org.RemoteDomains   = @() }

        # Email Address Policies
        try { $org.EmailAddressPolicies = @(Get-EmailAddressPolicy -ErrorAction Stop) } catch { $org.EmailAddressPolicies = @() }

        # Transport
        try { $org.TransportConfig = Get-TransportConfig -ErrorAction Stop } catch { $org.TransportConfig = $null }
        try { $org.TransportRules  = @(Get-TransportRule  -ErrorAction Stop | Select-Object Name, State, Priority, Mode, Comments) } catch { $org.TransportRules = @() }

        # Journal / Retention / DLP
        try { $org.JournalRules     = @(Get-JournalRule      -ErrorAction Stop) } catch { $org.JournalRules     = @() }
        try { $org.RetentionPolicies   = @(Get-RetentionPolicy    -ErrorAction Stop) } catch { $org.RetentionPolicies   = @() }
        try { $org.RetentionPolicyTags = @(Get-RetentionPolicyTag -ErrorAction Stop) } catch { $org.RetentionPolicyTags = @() }
        try { $org.DlpPolicies      = @(Get-DlpPolicy        -ErrorAction Stop) } catch { $org.DlpPolicies      = @() }

        # Mobile / OWA policies
        try { $org.MobileDevicePolicies = @(Get-MobileDeviceMailboxPolicy -ErrorAction Stop) } catch { $org.MobileDevicePolicies = @() }
        try { $org.OwaPolicies          = @(Get-OwaMailboxPolicy          -ErrorAction Stop) } catch { $org.OwaPolicies          = @() }

        # DAGs (all)
        try {
            $org.DAGs = @(Get-DatabaseAvailabilityGroup -Status -ErrorAction Stop | ForEach-Object {
                $dag = $_
                $copies = @{}
                try {
                    Get-MailboxDatabaseCopyStatus -Server ($dag.Servers | Select-Object -First 1) -ErrorAction SilentlyContinue | ForEach-Object {
                        $copies[$_.DatabaseName] = $copies[$_.DatabaseName] + @($_)
                    }
                } catch { }
                [pscustomobject]@{
                    DAG             = $dag
                    DatabaseCopies  = $copies
                }
            })
        } catch { $org.DAGs = @() }

        # Send Connectors (org-scoped, not per-server)
        try { $org.SendConnectors = @(Get-SendConnector -ErrorAction Stop) } catch { $org.SendConnectors = @() }

        # Federation
        try { $org.FederationTrust  = @(Get-FederationTrust  -ErrorAction Stop) } catch { $org.FederationTrust  = @() }
        try { $org.FederationOrg    = Get-FederatedOrganizationIdentifier -ErrorAction SilentlyContinue } catch { $org.FederationOrg = $null }

        # Hybrid
        try { $org.HybridConfig = Get-HybridConfiguration -ErrorAction Stop } catch { $org.HybridConfig = $null }

        # OAuth / AuthConfig
        try { $org.AuthConfig = Get-AuthConfig -ErrorAction Stop } catch { $org.AuthConfig = $null }
        try { $org.IntraOrgConnectors = @(Get-IntraOrganizationConnector -ErrorAction Stop) } catch { $org.IntraOrgConnectors = @() }

        # RBAC role groups (with members). Keep members as name/recipient-type only — full DN bloats the doc.
        $org.RoleGroups = @()
        try {
            $rgList = @(Get-RoleGroup -ErrorAction Stop | Sort-Object Name)
            foreach ($rg in $rgList) {
                $mem = @()
                try {
                    $mem = @(Get-RoleGroupMember -Identity $rg.Name -ErrorAction Stop |
                             Select-Object @{n='Name';e={$_.Name}}, @{n='Type';e={$_.RecipientType}})
                } catch { }
                $org.RoleGroups += [pscustomobject]@{
                    Name        = $rg.Name
                    Description = $rg.Description
                    Members     = $mem
                    ManagedBy   = @($rg.ManagedBy | ForEach-Object { $_.ToString() })
                }
            }
        } catch { }

        # Admin Audit Log Config (org-wide; controls which cmdlets/parameters are recorded in the admin audit log)
        try { $org.AdminAuditLog = Get-AdminAuditLogConfig -ErrorAction Stop } catch { $org.AdminAuditLog = $null }

        # Anti-spam filter configuration (org-wide settings objects; only present when anti-spam agents are installed)
        try { $org.ContentFilterConfig   = Get-ContentFilterConfig   -ErrorAction Stop } catch { $org.ContentFilterConfig   = $null }
        try { $org.SenderFilterConfig    = Get-SenderFilterConfig    -ErrorAction Stop } catch { $org.SenderFilterConfig    = $null }
        try { $org.RecipientFilterConfig = Get-RecipientFilterConfig -ErrorAction Stop } catch { $org.RecipientFilterConfig = $null }
        try { $org.SenderIdConfig        = Get-SenderIdConfig        -ErrorAction Stop } catch { $org.SenderIdConfig        = $null }

        # Auth Certificate (current + next) — org-wide (replicated to all servers).
        try { $org.AuthCertCurrent = Get-AuthConfig -ErrorAction Stop |
                 Select-Object CurrentCertificateThumbprint, PreviousCertificateThumbprint, NextCertificateThumbprint,
                               ServiceName, Realm } catch { $org.AuthCertCurrent = $null }

        # Scheduled Tasks (Exchange-related: MEAC auth-cert renewal, EXpress log cleanup).
        # ServerManager auto-start disable is an OS-level hardening step documented in Chapter 8, not a scheduled
        # task worth listing here.
        $org.ScheduledTasks = @()
        try {
            $foundTasks = @{}
            # Direct name lookup for known task names (fast path)
            # CSS-Exchange MEAC task is named "Daily Auth Certificate Check" (Register-AuthCertificateRenewalTask.ps1 default)
            $knownNames = @('Daily Auth Certificate Check','MonitorExchangeAuthCertificate','Exchange Log Cleanup','EXpressLogCleanup')
            foreach ($tn in $knownNames) {
                try { Get-ScheduledTask -TaskName $tn -ErrorAction SilentlyContinue | ForEach-Object { $foundTasks[$_.TaskName] = $_ } } catch { }
            }
            # Broad pattern search — catches variants across CSS-Exchange releases
            try {
                Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
                    $_.TaskName -match 'Daily Auth Certificate|MonitorExchangeAuth|ExchangeLogClean|EXpressLog'
                } | ForEach-Object { $foundTasks[$_.TaskName] = $_ }
            } catch { }
            foreach ($task in $foundTasks.Values) {
                $info = try { Get-ScheduledTaskInfo -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction SilentlyContinue } catch { $null }
                $org.ScheduledTasks += [pscustomobject]@{
                    Name      = $task.TaskName
                    Path      = $task.TaskPath
                    State     = $task.State
                    LastRun   = if ($info) { $info.LastRunTime }   else { $null }
                    NextRun   = if ($info) { $info.NextRunTime }   else { $null }
                    LastResult= if ($info) { $info.LastTaskResult } else { $null }
                    Actions   = @($task.Actions | ForEach-Object { if ($_.Execute) { "$($_.Execute) $($_.Arguments)".Trim() } })
                }
            }
        } catch { }

        return $org
    }

    function Get-ServerReportData {
        param([Parameter(Mandatory)][string]$ServerName)

        $srv = @{ ServerName = $ServerName }

        # Exchange server object
        try { $srv.ExServer = Get-ExchangeServer -Identity $ServerName -ErrorAction Stop } catch { $srv.ExServer = $null }

        # Databases on this server
        try { $srv.Databases = @(Get-MailboxDatabase -Server $ServerName -Status -ErrorAction Stop) } catch { $srv.Databases = @() }

        # Virtual directories
        try { $srv.VDirOWA    = @(Get-OwaVirtualDirectory              -Server $ServerName -ADPropertiesOnly -ErrorAction Stop) } catch { $srv.VDirOWA    = @() }
        try { $srv.VDirECP    = @(Get-EcpVirtualDirectory              -Server $ServerName -ADPropertiesOnly -ErrorAction Stop) } catch { $srv.VDirECP    = @() }
        try { $srv.VDirEWS    = @(Get-WebServicesVirtualDirectory      -Server $ServerName -ADPropertiesOnly -ErrorAction Stop) } catch { $srv.VDirEWS    = @() }
        try { $srv.VDirAS     = @(Get-ActiveSyncVirtualDirectory       -Server $ServerName -ADPropertiesOnly -ErrorAction Stop) } catch { $srv.VDirAS     = @() }
        try { $srv.VDirOAB    = @(Get-OabVirtualDirectory              -Server $ServerName -ADPropertiesOnly -ErrorAction Stop) } catch { $srv.VDirOAB    = @() }
        try { $srv.VDirMAPI   = @(Get-MapiVirtualDirectory             -Server $ServerName -ADPropertiesOnly -ErrorAction Stop) } catch { $srv.VDirMAPI   = @() }
        try { $srv.VDirPW     = @(Get-PowerShellVirtualDirectory       -Server $ServerName -ADPropertiesOnly -ErrorAction Stop) } catch { $srv.VDirPW     = @() }
        try { $srv.AutodiscoverSCP = Get-ClientAccessService           -Identity $ServerName -ErrorAction Stop } catch { $srv.AutodiscoverSCP = $null }

        # Connectors (Receive only — Send is org-scoped)
        try { $srv.ReceiveConnectors = @(Get-ReceiveConnector -Server $ServerName -ErrorAction Stop) } catch { $srv.ReceiveConnectors = @() }

        # IMAP/POP settings (local only — remote Exchange management remoting requires a separate PS session)
        $srv.ImapSettings = $null
        $srv.PopSettings  = $null
        if ($ServerName -ieq $env:COMPUTERNAME) {
            try { $srv.ImapSettings = Get-ImapSettings -Server $ServerName -ErrorAction Stop } catch { }
            try { $srv.PopSettings  = Get-PopSettings  -Server $ServerName -ErrorAction Stop } catch { }
        }

        # Certificates
        try { $srv.Certificates = @(Get-ExchangeCertificate -Server $ServerName -ErrorAction Stop) } catch { $srv.Certificates = @() }

        # Transport agents (only present on servers with Hub Transport)
        try { $srv.TransportAgents = @(Get-TransportAgent -ErrorAction Stop) } catch { $srv.TransportAgents = @() }

        # Database copy status (per server; runs against local server where available)
        try { $srv.DatabaseCopies = @(Get-MailboxDatabaseCopyStatus -Server $ServerName -ErrorAction Stop |
                                      Select-Object Name, DatabaseName, Status, ContentIndexState, CopyQueueLength, ReplayQueueLength, ActivationPreference, MailboxServer) } catch { $srv.DatabaseCopies = @() }

        # Defender preferences — only meaningful for local server (remote would need CIM/PSSession)
        $srv.DefenderExclusions = $null
        if ($ServerName -ieq $env:COMPUTERNAME) {
            try {
                $mp = Get-MpPreference -ErrorAction Stop
                $srv.DefenderExclusions = [pscustomobject]@{
                    ExclusionPath      = @($mp.ExclusionPath)
                    ExclusionProcess   = @($mp.ExclusionProcess)
                    ExclusionExtension = @($mp.ExclusionExtension)
                    RealTimeEnabled    = -not $mp.DisableRealtimeMonitoring
                }
            } catch { }
        }

        # IIS log configuration (local only — remote IIS queries require WinRM/PSSession, out of scope)
        $srv.IISLogs = $null
        if ($ServerName -ieq $env:COMPUTERNAME) {
            try {
                Import-Module WebAdministration -ErrorAction SilentlyContinue
                $sites = @(Get-Website -ErrorAction SilentlyContinue | Where-Object { $_.Name -in 'Default Web Site','Exchange Back End' } | ForEach-Object {
                    [pscustomobject]@{
                        Name      = $_.Name
                        LogDir    = $_.LogFile.Directory
                        LogFormat = $_.LogFile.LogFormat
                        Period    = $_.LogFile.Period
                    }
                })
                $srv.IISLogs = [pscustomobject]@{
                    Sites = $sites
                    ExchangeLogPath = Join-Path (Split-Path $env:ExchangeInstallPath -Parent) 'Logging'
                }
            } catch { }
        }

        # Hardware/OS data — direct CIM for local server, CIM/WSMan + prompt for remote
        if ($ServerName -ieq $env:COMPUTERNAME) {
            $srv.RemoteData = @{
                ComputerName = $ServerName
                Reachable    = $true
                Error        = $null
                OS           = Get-CimInstance Win32_OperatingSystem          -ErrorAction SilentlyContinue
                CPU          = Get-CimInstance Win32_Processor                 -ErrorAction SilentlyContinue
                ComputerSys  = Get-CimInstance Win32_ComputerSystem            -ErrorAction SilentlyContinue
                PageFile     = Get-CimInstance Win32_PageFileSetting           -ErrorAction SilentlyContinue
                Volumes      = @(Get-CimInstance Win32_Volume -Filter 'DriveType=3'                          -ErrorAction SilentlyContinue)
                NICs         = @(Get-CimInstance Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=TRUE'  -ErrorAction SilentlyContinue)
            }
        } else {
            $srv.RemoteData = Invoke-RemoteQueryWithPrompt -ComputerName $ServerName
        }

        return $srv
    }

    function Get-InstallationReportData {
        param(
            [ValidateSet('All','Org','Local')][string]$Scope = 'All',
            [string[]]$IncludeServers = @()
        )

        $data = @{
            Org     = $null
            Servers = @()
            Local   = @{}
        }

        # Org-wide data
        if ($Scope -in 'All','Org') {
            Write-MyVerbose 'Collecting org-wide Exchange configuration'
            $data.Org = Get-OrganizationReportData
        }

        # Per-server data
        if ($Scope -in 'All','Local') {
            try {
                $allServers = @(Get-ExchangeServer -ErrorAction Stop | Sort-Object Name)
                if ($IncludeServers.Count -gt 0) {
                    $allServers = @($allServers | Where-Object { $_.Name -in $IncludeServers })
                }
                foreach ($srv in $allServers) {
                    Write-MyVerbose ('Collecting data for server {0}' -f $srv.Name)
                    $srvData = Get-ServerReportData -ServerName $srv.Name
                    $srvData.IsLocalServer = ($srv.Name -ieq $env:COMPUTERNAME)
                    $data.Servers += $srvData
                }
            } catch {
                Write-MyWarning ('Get-InstallationReportData: could not enumerate Exchange servers: {0}' -f $_.Exception.Message)
            }
        }

        # Local system data (always, for the server running EXpress)
        $data.Local.OS          = Get-CimInstance Win32_OperatingSystem        -ErrorAction SilentlyContinue
        $data.Local.CPU         = Get-CimInstance Win32_Processor               -ErrorAction SilentlyContinue
        $data.Local.ComputerSys = Get-CimInstance Win32_ComputerSystem          -ErrorAction SilentlyContinue
        $data.Local.PageFile    = Get-CimInstance Win32_PageFileSetting         -ErrorAction SilentlyContinue
        $data.Local.Volumes     = @(Get-CimInstance Win32_Volume -Filter 'DriveType=3' -ErrorAction SilentlyContinue)
        $data.Local.NICs        = @(Get-CimInstance Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=TRUE' -ErrorAction SilentlyContinue)

        return $data
    }

    function New-InstallationReport {
        Write-MyOutput 'Generating Installation Report'
        $reportPath = Join-Path $State['ReportsPath'] ('{0}_EXpress_Report_{1}.html' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd-HHmmss'))
        $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

        function Format-Badge($text, $type) {
            $colors = @{ ok='background:#107c10;color:#fff'; warn='background:#d83b01;color:#fff'; fail='background:#c50f1f;color:#fff'; info='background:#0078d4;color:#fff'; na='background:#8a8886;color:#fff' }
            '<span style="display:inline-block;padding:2px 10px;border-radius:12px;font-size:12px;font-weight:600;' + $colors[$type.ToLower()] + '">' + $text + '</span>'
        }
        function New-HtmlSection($id, $title, $content) {
            '<section id="' + $id + '" class="section"><h2 class="section-title">' + $title + '</h2><div class="section-body">' + $content + '</div></section>'
        }

        # ── 1. INSTALLATION PARAMETERS ────────────────────────────────────────
        $instRows = [System.Collections.Generic.List[string]]::new()
        $instMode = if ($State['InstallEdge']) { 'Edge Transport' } elseif ($State['InstallRecipientManagement']) { 'Recipient Management Tools' } elseif ($State['InstallManagementTools']) { 'Management Tools' } elseif ($State['StandaloneOptimize']) { 'Standalone Optimize' } elseif ($State['NoSetup']) { 'Optimization Only' } else { 'Mailbox Server' }
        $instRows.Add('<tr><td>Script Version</td><td>EXpress v{0}</td></tr>' -f $ScriptVersion)
        $instRows.Add('<tr><td>Report Generated</td><td>{0}</td></tr>' -f $reportDate)
        $instRows.Add('<tr><td>Server</td><td>{0}</td></tr>' -f $env:COMPUTERNAME)
        $instRows.Add('<tr><td>Installation Mode</td><td>{0}</td></tr>' -f $instMode)
        $instRows.Add('<tr><td>Organization</td><td>{0}</td></tr>' -f $State['OrganizationName'])
        $instRows.Add(('<tr><td>Setup Version</td><td>{0} ({1})</td></tr>' -f $State['SetupVersion'], (Get-SetupTextVersion $State['SetupVersion'])))
        $instRows.Add('<tr><td>Install Path</td><td>{0}</td></tr>' -f $State['InstallPath'])
        if ($State['Namespace'])        { $instRows.Add('<tr><td>Namespace</td><td>{0}</td></tr>' -f $State['Namespace']) }
        if ($State['DownloadDomain'])   { $instRows.Add('<tr><td>OWA Download Domain</td><td>{0}</td></tr>' -f $State['DownloadDomain']) }
        if ($State['DAGName'])          { $instRows.Add('<tr><td>DAG</td><td>{0}</td></tr>' -f $State['DAGName']) }
        if ($State['CertificatePath'])  { $instRows.Add('<tr><td>Certificate Path</td><td>{0}</td></tr>' -f $State['CertificatePath']) }
        if ($State['CopyServerConfig']) { $instRows.Add('<tr><td>Source Server Config</td><td>{0}</td></tr>' -f $State['CopyServerConfig']) }
        if ($State['LogRetentionDays']) { $instRows.Add('<tr><td>Log Retention</td><td>{0} days</td></tr>' -f $State['LogRetentionDays']) }
        $modeText = if ($State['ConfigDriven']) { 'Autopilot (fully automated)' } else { 'Copilot (interactive)' }
        $instRows.Add('<tr><td>Mode</td><td>{0}</td></tr>' -f $modeText)
        $instRows.Add('<tr><td>TLS 1.2 Enforced</td><td>{0}</td></tr>' -f $State['EnableTLS12'])
        $instRows.Add('<tr><td>TLS 1.3 Enforced</td><td>{0}</td></tr>' -f $State['EnableTLS13'])
        $instRows.Add('<tr><td>SSL 3 Disabled</td><td>{0}</td></tr>' -f $State['DisableSSL3'])
        $instRows.Add('<tr><td>RC4 Disabled</td><td>{0}</td></tr>' -f $State['DisableRC4'])
        $instRows.Add('<tr><td>Log File</td><td><code>{0}</code></td></tr>' -f $State['TranscriptFile'])

        # ── 2. SYSTEM INFORMATION ─────────────────────────────────────────────
        $sysRows = [System.Collections.Generic.List[string]]::new()
        try {
            $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
            $sysRows.Add(('<tr><td>Operating System</td><td>{0}</td><td>{1}</td></tr>' -f $os.Caption, (Format-Badge 'OK' 'ok')))
            $sysRows.Add('<tr><td>OS Build</td><td>{0}</td><td></td></tr>' -f $os.Version)
            $sysRows.Add('<tr><td>Last Boot</td><td>{0}</td><td></td></tr>' -f $os.LastBootUpTime.ToString('yyyy-MM-dd HH:mm:ss'))
            $totalRAM = [math]::Round($os.TotalVisibleMemorySize / 1MB, 0)
            $sysRows.Add('<tr><td>Total RAM</td><td>{0} GB</td><td></td></tr>' -f $totalRAM)
        } catch { $sysRows.Add('<tr><td colspan="3">OS info unavailable</td></tr>') }
        try {
            $cpuList     = @(Get-CimInstance Win32_Processor -ErrorAction Stop)
            $cpu         = $cpuList[0]
            $totalCores  = ($cpuList | Measure-Object -Property NumberOfCores -Sum).Sum
            $totalLogical= ($cpuList | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
            $sysRows.Add(('<tr><td>CPU</td><td>{0}</td><td>{1} cores / {2} logical</td></tr>' -f $cpu.Name.Trim(), $totalCores, $totalLogical))
        } catch { }
        try {
            $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
            $sysRows.Add(('<tr><td>Computer Name (FQDN)</td><td>{0}.{1}</td><td></td></tr>' -f $cs.DNSHostName, $cs.Domain))
        } catch { }
        # Volumes — exclude DVD-ROM and removable drives
        try {
            Get-Volume -ErrorAction SilentlyContinue | Where-Object {
                $_.DriveLetter -and $_.DriveType -notin 'CD-ROM','Removable' -and $_.Size -gt 0
            } | ForEach-Object {
                $freeGB    = [math]::Round($_.SizeRemaining / 1GB, 1)
                $totalGB   = [math]::Round($_.Size / 1GB, 1)
                $freePct   = [math]::Round($_.SizeRemaining / $_.Size * 100, 0)
                $auBadge   = if ($_.AllocationUnitSize -eq 65536) { Format-Badge '64 KB ✓' 'ok' } elseif ($_.AllocationUnitSize) { Format-Badge ('{0} KB !' -f ($_.AllocationUnitSize / 1KB)) 'warn' } else { '' }
                $diskBadge = if ($freePct -lt 10) { Format-Badge ('Free {0}%' -f $freePct) 'fail' } elseif ($freePct -lt 20) { Format-Badge ('Free {0}%' -f $freePct) 'warn' } else { Format-Badge ('Free {0}%' -f $freePct) 'ok' }
                $sysRows.Add(('<tr><td>Volume {0}:</td><td>{1} GB free of {2} GB &nbsp; {3}</td><td>{4} &nbsp; Alloc: {5}</td></tr>' -f $_.DriveLetter, $freeGB, $totalGB, $diskBadge, $_.FileSystem, $auBadge))
            }
        } catch { }
        # Network adapters — IP address + DNS servers
        try {
            $nicIPs  = @{}
            $nicDNS  = @{}
            Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                Where-Object { $_.InterfaceAlias -notlike '*Loopback*' } |
                ForEach-Object { $nicIPs[$_.InterfaceAlias] = ('{0}/{1}' -f $_.IPAddress, $_.PrefixLength) }
            Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                Where-Object { $_.InterfaceAlias -notlike '*Loopback*' -and $_.ServerAddresses } |
                ForEach-Object { $nicDNS[$_.InterfaceAlias] = ($_.ServerAddresses -join ', ') }
            foreach ($nic in ($nicIPs.Keys | Sort-Object)) {
                $dns = if ($nicDNS[$nic]) { $nicDNS[$nic] } else { '<em>not set</em>' }
                $sysRows.Add(('<tr><td>NIC: {0}</td><td>{1}</td><td>DNS: {2}</td></tr>' -f $nic, $nicIPs[$nic], $dns))
            }
        } catch { }
        $sysContent = '<table class="data-table"><tr><th>Property</th><th>Value</th><th>Detail / Status</th></tr>' + ($sysRows -join '') + '</table>'

        # ── 3. ACTIVE DIRECTORY ───────────────────────────────────────────────
        $adRows = [System.Collections.Generic.List[string]]::new()
        try {
            $cs2 = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
            $adRows.Add('<tr><td>Domain</td><td>{0}</td><td></td></tr>' -f $cs2.Domain)
        } catch { }
        try {
            $ffl = Get-ForestFunctionalLevel
            $fflBadge = if ($ffl -ge $FOREST_LEVEL2012R2) { Format-Badge 'OK' 'ok' } else { Format-Badge 'WARN' 'warn' }
            $adRows.Add(('<tr><td>Forest Functional Level</td><td>{0} ({1})</td><td>{2}</td></tr>' -f $ffl, (Get-FFLText $ffl), $fflBadge))
        } catch { }
        try {
            $exOrg = Get-ExchangeOrganization
            if ($exOrg) { $adRows.Add('<tr><td>Exchange Organization</td><td>{0}</td><td></td></tr>' -f $exOrg) }
            $exFL = Get-ExchangeForestLevel
            $adRows.Add('<tr><td>Exchange Forest Schema</td><td>{0}</td><td></td></tr>' -f $exFL)
            $exDL = Get-ExchangeDomainLevel
            $adRows.Add('<tr><td>Exchange Domain Level</td><td>{0}</td><td></td></tr>' -f $exDL)
        } catch { }
        $adContent = '<table class="data-table"><tr><th>Property</th><th>Value</th><th>Status</th></tr>' + ($adRows -join '') + '</table>'

        # ── 4. EXCHANGE CONFIGURATION ─────────────────────────────────────────
        $exRows    = [System.Collections.Generic.List[string]]::new()
        $vdirRows  = [System.Collections.Generic.List[string]]::new()
        $connRows  = [System.Collections.Generic.List[string]]::new()
        $dbRows    = [System.Collections.Generic.List[string]]::new()
        $certRows  = [System.Collections.Generic.List[string]]::new()
        $exVersion = Get-SetupTextVersion $State['SetupVersion']

        try {
            $exSrv = Get-ExchangeServer $env:COMPUTERNAME -ErrorAction Stop
            $exVersion = $exSrv.AdminDisplayVersion.ToString()
            $exRows.Add('<tr><td>Exchange Version</td><td>{0}</td><td></td></tr>' -f $exSrv.AdminDisplayVersion)
            $exRows.Add('<tr><td>Server Role</td><td>{0}</td><td></td></tr>' -f ($exSrv.ServerRole -join ', '))
            $exRows.Add('<tr><td>Edition</td><td>{0}</td><td></td></tr>' -f $exSrv.Edition)
            $exRows.Add('<tr><td>AD Site</td><td>{0}</td><td></td></tr>' -f $exSrv.Site)
        } catch { $exRows.Add('<tr><td colspan="3">Exchange server query unavailable</td></tr>') }
        # Autodiscover SCP (Client Access Service, not a virtual directory)
        try {
            $cas = Get-ClientAccessService -Identity $env:COMPUTERNAME -ErrorAction SilentlyContinue
            if ($cas) { $exRows.Add('<tr><td>Autodiscover SCP</td><td>{0}</td><td></td></tr>' -f $cas.AutoDiscoverServiceInternalUri) }
        } catch { }
        try {
            $orgCfg = Get-OrganizationConfig -ErrorAction Stop
            $exRows.Add('<tr><td>Organization Name</td><td>{0}</td><td></td></tr>' -f $orgCfg.Name)
            $oauthBadge = if ($orgCfg.OAuth2ClientProfileEnabled) { Format-Badge 'Enabled' 'ok' } else { Format-Badge 'Disabled' 'warn' }
            $exRows.Add(('<tr><td>Modern Auth (OAuth2)</td><td>{0}</td><td>{1}</td></tr>' -f $orgCfg.OAuth2ClientProfileEnabled, $oauthBadge))
            $mapiBadge = if ($orgCfg.MapiHttpEnabled) { Format-Badge 'Enabled' 'ok' } else { Format-Badge 'Disabled' 'warn' }
            $exRows.Add(('<tr><td>MAPI/HTTP</td><td>{0}</td><td>{1}</td></tr>' -f $orgCfg.MapiHttpEnabled, $mapiBadge))
        } catch { }
        try {
            Get-AcceptedDomain -ErrorAction Stop | ForEach-Object {
                $exRows.Add(('<tr><td>Accepted Domain</td><td>{0}</td><td>{1}</td></tr>' -f $_.DomainName, (Format-Badge $_.DomainType 'info')))
            }
        } catch { }

        # Virtual directories — Autodiscover SCP + OWA, ECP, EWS, OAB, ActiveSync, MAPI
        $vdirRows.Add('<tr><th>Service</th><th>Internal URL</th><th>External URL</th></tr>')
        try {
            $cas = Get-ClientAccessService -Identity $env:COMPUTERNAME -ErrorAction Stop
            $scpUri = if ($cas.AutoDiscoverServiceInternalUri) { $cas.AutoDiscoverServiceInternalUri.AbsoluteUri } else { '<em>not set</em>' }
            $vdirRows.Add(('<tr><td>Autodiscover SCP</td><td>{0}</td><td><em>n/a (SCP)</em></td></tr>' -f $scpUri))
        } catch { }
        @(
            @{ Name='OWA';         Cmd={ Get-OwaVirtualDirectory         -Server $env:COMPUTERNAME -ADPropertiesOnly -ErrorAction SilentlyContinue | Select-Object -First 1 } }
            @{ Name='ECP';         Cmd={ Get-EcpVirtualDirectory         -Server $env:COMPUTERNAME -ADPropertiesOnly -ErrorAction SilentlyContinue | Select-Object -First 1 } }
            @{ Name='EWS';         Cmd={ Get-WebServicesVirtualDirectory  -Server $env:COMPUTERNAME -ADPropertiesOnly -ErrorAction SilentlyContinue | Select-Object -First 1 } }
            @{ Name='OAB';         Cmd={ Get-OabVirtualDirectory         -Server $env:COMPUTERNAME -ADPropertiesOnly -ErrorAction SilentlyContinue | Select-Object -First 1 } }
            @{ Name='ActiveSync';  Cmd={ Get-ActiveSyncVirtualDirectory   -Server $env:COMPUTERNAME -ADPropertiesOnly -ErrorAction SilentlyContinue | Select-Object -First 1 } }
            @{ Name='MAPI';        Cmd={ Get-MapiVirtualDirectory         -Server $env:COMPUTERNAME -ADPropertiesOnly -ErrorAction SilentlyContinue | Select-Object -First 1 } }
        ) | ForEach-Object {
            try {
                $vd = & $_.Cmd
                if ($vd) {
                    $int = if ($vd.InternalUrl) { $vd.InternalUrl.AbsoluteUri } else { '<em>not set</em>' }
                    $ext = if ($vd.ExternalUrl) { $vd.ExternalUrl.AbsoluteUri } else { '<em>not set</em>' }
                    $vdirRows.Add(('<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>' -f $_.Name, $int, $ext))
                }
            } catch { }
        }

        # Mailbox databases — try with status, fall back without
        $dbRows.Add('<tr><th>Database</th><th>DB Path</th><th>Log Path</th><th>Status</th></tr>')
        try {
            $dbs = Get-MailboxDatabase -Server $env:COMPUTERNAME -Status -ErrorAction SilentlyContinue
            if (-not $dbs) {
                $dbs = Get-MailboxDatabase -Server $env:COMPUTERNAME -ErrorAction Stop
            }
            if ($dbs) {
                $dbs | ForEach-Object {
                    $mounted = if ($null -ne $_.Mounted) { $_.Mounted } else { $null }
                    $mountedText  = if ($null -eq $mounted) { 'Unknown' } elseif ($mounted) { 'Mounted' } else { 'Dismounted' }
                    $mountedBadge = if ($null -eq $mounted) { Format-Badge 'Unknown' 'na' } elseif ($mounted) { Format-Badge 'Mounted' 'ok' } else { Format-Badge 'Dismounted' 'warn' }
                    $dbRows.Add(('<tr><td>{0}</td><td><code>{1}</code></td><td><code>{2}</code></td><td>{3}</td></tr>' -f $_.Name, $_.EdbFilePath, $_.LogFolderPath, $mountedBadge))
                }
            } else {
                $dbRows.Add('<tr><td colspan="4"><em>No mailbox databases found on this server</em></td></tr>')
            }
        } catch { $dbRows.Add(('<tr><td colspan="4">Query failed: {0}</td></tr>' -f $_.Exception.Message)) }

        # Receive connectors
        $connRows.Add('<tr><th>Connector</th><th>Bindings</th><th>Remote IP Ranges</th><th>Auth</th><th>Permission Groups</th></tr>')
        try {
            Get-ReceiveConnector -Server $env:COMPUTERNAME -ErrorAction Stop | ForEach-Object {
                $connRows.Add(('<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td></tr>' -f $_.Name, ($_.Bindings -join '<br>'), ($_.RemoteIPRanges -join '<br>'), $_.AuthMechanism, $_.PermissionGroups))
            }
        } catch { $connRows.Add('<tr><td colspan="5">Receive connector query failed</td></tr>') }

        # Certificates
        $certRows.Add('<tr><th>Subject</th><th>Expiry</th><th>Services</th><th>Thumbprint</th></tr>')
        try {
            Get-ExchangeCertificate -Server $env:COMPUTERNAME -ErrorAction Stop | ForEach-Object {
                # Skip phantom entries: no subject, no thumbprint, or NotAfter = DateTime.MinValue
                if ([string]::IsNullOrEmpty($_.Thumbprint) -or $_.NotAfter -le [datetime]'1970-01-01') { return }
                $daysLeft = [int][Math]::Floor(($_.NotAfter - (Get-Date)).TotalDays)
                $expiryBadge = if ($daysLeft -lt 30) { Format-Badge ('Expires {0}d!' -f $daysLeft) 'fail' } elseif ($daysLeft -lt 90) { Format-Badge ('Expires {0}d' -f $daysLeft) 'warn' } else { Format-Badge ('{0} ({1}d)' -f $_.NotAfter.ToString('yyyy-MM-dd'), $daysLeft) 'ok' }
                $certRows.Add(('<tr><td>{0}</td><td>{1}</td><td>{2}</td><td><code>{3}</code></td></tr>' -f $_.Subject, $expiryBadge, $_.Services, $_.Thumbprint))
            }
        } catch { $certRows.Add('<tr><td colspan="4">Certificate query failed</td></tr>') }

        # Exchange Optimizations — org/transport level settings
        $exchOptRows = [System.Collections.Generic.List[string]]::new()
        $exchOptRows.Add('<tr><th>Setting</th><th>Current Value</th><th>Recommendation</th><th>Status</th></tr>')
        try {
            $orgCfg2 = Get-OrganizationConfig -ErrorAction SilentlyContinue
            if ($orgCfg2) {
                $toEnabled  = $orgCfg2.ActivityBasedAuthenticationTimeoutEnabled
                $toInterval = $orgCfg2.ActivityBasedAuthenticationTimeoutInterval
                $toBadge    = if ($toEnabled -and $toInterval -le [TimeSpan]'06:00:00') { Format-Badge '✓' 'ok' } else { Format-Badge 'Review' 'warn' }
                $exchOptRows.Add(('<tr><td>OWA/ECP Session Timeout</td><td>Enabled: {0} / Interval: {1}</td><td>Enabled, ≤ 6 h (security compliance)</td><td>{2}</td></tr>' -f $toEnabled, $toInterval, $toBadge))

                $ceipBadge = if (-not $orgCfg2.CustomerFeedbackEnabled) { Format-Badge 'Disabled ✓' 'ok' } else { Format-Badge 'Enabled' 'warn' }
                $exchOptRows.Add(('<tr><td>CEIP / Telemetry</td><td>{0}</td><td>Disabled (privacy / GDPR)</td><td>{1}</td></tr>' -f $orgCfg2.CustomerFeedbackEnabled, $ceipBadge))
            }
        } catch { }
        try {
            $transCfg2 = Get-TransportConfig -ErrorAction SilentlyContinue
            if ($transCfg2) {
                # MaxSendSize / MaxReceiveSize may be Unlimited ($null .Value) on fresh orgs.
                $maxSendMB = if ($transCfg2.MaxSendSize    -and $transCfg2.MaxSendSize.Value)    { [math]::Round($transCfg2.MaxSendSize.Value.ToBytes()    / 1MB, 0) } else { $null }
                $maxRecvMB = if ($transCfg2.MaxReceiveSize -and $transCfg2.MaxReceiveSize.Value) { [math]::Round($transCfg2.MaxReceiveSize.Value.ToBytes() / 1MB, 0) } else { $null }
                $maxSendDisp = if ($null -ne $maxSendMB) { ('{0} MB' -f $maxSendMB) } else { 'Unlimited / not set' }
                $maxRecvDisp = if ($null -ne $maxRecvMB) { ('{0} MB' -f $maxRecvMB) } else { 'Unlimited / not set' }
                $sizeBadge = if ($null -ne $maxSendMB -and $maxSendMB -ge 50) { Format-Badge '✓' 'ok' } else { Format-Badge 'Default 25 MB' 'warn' }
                $exchOptRows.Add(('<tr><td>Max Message Size</td><td>Send: {0} / Recv: {1}</td><td>≥ 50 MB (modern workflow files)</td><td>{2}</td></tr>' -f $maxSendDisp, $maxRecvDisp, $sizeBadge))

                $ndrBadge = if ($transCfg2.InternalDsnSendHtml -and $transCfg2.ExternalDsnSendHtml) { Format-Badge 'Enabled ✓' 'ok' } else { Format-Badge 'Plain text' 'warn' }
                $exchOptRows.Add(('<tr><td>HTML Non-Delivery Reports</td><td>Internal: {0} / External: {1}</td><td>Enabled (improves end-user NDR readability)</td><td>{2}</td></tr>' -f $transCfg2.InternalDsnSendHtml, $transCfg2.ExternalDsnSendHtml, $ndrBadge))

                $snBadge = if ($transCfg2.SafetyNetHoldTime -ge [TimeSpan]'2.00:00:00') { Format-Badge '✓' 'ok' } else { Format-Badge 'Short' 'warn' }
                $exchOptRows.Add(('<tr><td>Safety Net Hold Time</td><td>{0}</td><td>≥ 2 days (message redelivery after DB failover)</td><td>{1}</td></tr>' -f $transCfg2.SafetyNetHoldTime, $snBadge))
            }
        } catch { }
        try {
            $transSvc2 = Get-TransportService -Identity $env:COMPUTERNAME -ErrorAction SilentlyContinue
            if ($transSvc2) {
                $expBadge = if ($transSvc2.MessageExpirationTimeout -ge [TimeSpan]'7.00:00:00') { Format-Badge '✓' 'ok' } else { Format-Badge 'Default 2d' 'warn' }
                $exchOptRows.Add(('<tr><td>Message Expiration Timeout</td><td>{0}</td><td>7 days (delay NDRs during multi-day outages)</td><td>{1}</td></tr>' -f $transSvc2.MessageExpirationTimeout, $expBadge))
            }
        } catch { }

        $exContent = '<table class="data-table"><tr><th>Property</th><th>Value</th><th>Status</th></tr>' + ($exRows -join '') + '</table>' +
            '<h3 class="subsection">Virtual Directory URLs</h3><table class="data-table">' + ($vdirRows -join '') + '</table>' +
            '<h3 class="subsection">Mailbox Databases</h3><table class="data-table">' + ($dbRows -join '') + '</table>' +
            '<h3 class="subsection">Receive Connectors</h3><table class="data-table">' + ($connRows -join '') + '</table>' +
            '<h3 class="subsection">Certificates</h3><table class="data-table">' + ($certRows -join '') + '</table>' +
            '<h3 class="subsection">Exchange Optimizations</h3><table class="data-table">' + ($exchOptRows -join '') + '</table>'

        # ── 5. SECURITY SETTINGS (with Exchange best-practice + reference column) ─
        $secRows = [System.Collections.Generic.List[string]]::new()
        function Get-SecRegVal($path, $name) { try { (Get-ItemProperty -Path $path -Name $name -ErrorAction Stop).$name } catch { $null } }
        function Format-RefLink($url, $label) { '<a href="' + $url + '" target="_blank" style="font-size:0.85em;white-space:nowrap">' + $label + ' ↗</a>' }

        # TLS protocols — show current value + Exchange recommendation
        @(
            @{ Proto='1.0'; Rec='Disabled'; LegacyRisk=$true;  CisId='CIS L1 / PCI-DSS 4.2.1'; RefUrl='https://techcommunity.microsoft.com/t5/exchange-team-blog/exchange-server-tls-guidance-part-1-getting-ready-for-tls-1-2/ba-p/607649'; RefLabel='Exchange TLS Guide' }
            @{ Proto='1.1'; Rec='Disabled'; LegacyRisk=$true;  CisId='CIS L1 / PCI-DSS 4.2.1'; RefUrl='https://techcommunity.microsoft.com/t5/exchange-team-blog/exchange-server-tls-guidance-part-1-getting-ready-for-tls-1-2/ba-p/607649'; RefLabel='Exchange TLS Guide' }
            @{ Proto='1.2'; Rec='Enabled';  LegacyRisk=$false; CisId='CIS L1 / PCI-DSS 4.2.1'; RefUrl='https://techcommunity.microsoft.com/blog/exchange/exchange-server-tls-guidance-part-2-enabling-tls-1-2-and-identifying-clients-not/607761'; RefLabel='Exchange TLS Guide' }
            @{ Proto='1.3'; Rec='Enabled (Exchange SE / 2019 CU15+ on WS2022+)'; LegacyRisk=$false; CisId='Best practice'; RefUrl='https://support.microsoft.com/en-us/topic/partial-tls-1-3-support-for-exchange-server-2019-5f4058f5-b288-4859-9a85-9aac680f50fe'; RefLabel='Exchange Blog' }
        ) | ForEach-Object {
            $proto      = $_.Proto
            $srvEnabled = Get-SecRegVal "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS $proto\Server" 'Enabled'
            $cliEnabled = Get-SecRegVal "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS $proto\Client" 'Enabled'
            $isEnabled  = -not ($srvEnabled -eq 0 -or $cliEnabled -eq 0)
            $valText    = if ($null -eq $srvEnabled -and $null -eq $cliEnabled) { 'OS Default' } else { "Srv=$srvEnabled / Cli=$cliEnabled" }
            $label      = if ($isEnabled) { 'Enabled' } else { 'Disabled' }
            $badgeType  = if ($_.LegacyRisk) { if ($isEnabled) { 'warn' } else { 'ok' } } else { if ($isEnabled) { 'ok' } else { 'warn' } }
            $secRows.Add(('<tr><td>TLS {0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td></tr>' -f $proto, $valText, $_.Rec, (Format-Badge $label $badgeType), (Format-RefLink $_.RefUrl $_.RefLabel), $_.CisId))
        }
        $strongCrypto = Get-SecRegVal 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319' 'SchUseStrongCrypto'
        $strongBadge  = if ($strongCrypto -eq 1) { Format-Badge 'Enabled' 'ok' } else { Format-Badge 'Not set' 'warn' }
        $secRows.Add(('<tr><td>.NET Strong Crypto</td><td>SchUseStrongCrypto = {0}</td><td>1 (required)</td><td>{1}</td><td>{2}</td><td>CIS L1 §18.3</td></tr>' -f $strongCrypto, $strongBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/dotnet/framework/network-programming/tls' 'MS Learn')))
        try {
            $smb1 = (Get-SmbServerConfiguration -ErrorAction Stop).EnableSMB1Protocol
            $smb1Badge = if ($smb1) { Format-Badge 'Enabled (risk)' 'warn' } else { Format-Badge 'Disabled' 'ok' }
            $secRows.Add(('<tr><td>SMBv1</td><td>{0}</td><td>Disabled</td><td>{1}</td><td>{2}</td><td>CIS L1 §18.3 / BSI SYS.1</td></tr>' -f $smb1, $smb1Badge, (Format-RefLink 'https://techcommunity.microsoft.com/t5/storage-at-microsoft/stop-using-smb1/ba-p/425858' 'Microsoft Blog')))
        } catch { }
        $wdigest = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest' 'UseLogonCredential'
        $wdigestBadge = if ($wdigest -eq 0) { Format-Badge 'Disabled' 'ok' } else { Format-Badge 'Enabled (risk)' 'warn' }
        $secRows.Add(('<tr><td>WDigest Caching</td><td>UseLogonCredential = {0}</td><td>0 = Disabled</td><td>{1}</td><td>{2}</td><td>CIS L1 §18.9.48 / DISA STIG</td></tr>' -f $wdigest, $wdigestBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/windows-server/security/credentials-protection-and-management/configuring-additional-lsa-protection' 'MS Learn')))
        $lsaPPL = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' 'RunAsPPL'
        $lsaBadge = if ($lsaPPL -eq 1) { Format-Badge 'Enabled' 'ok' } else { Format-Badge 'Not set' 'warn' }
        $secRows.Add(('<tr><td>LSA Protection (RunAsPPL)</td><td>{0}</td><td>1 = Enabled (Ex2019 CU12+/SE)</td><td>{1}</td><td>{2}</td><td>CIS L2 §2.3.11 / DISA STIG</td></tr>' -f $lsaPPL, $lsaBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/windows-server/security/credentials-protection-and-management/configuring-additional-lsa-protection' 'MS Learn')))
        $lmLevel = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' 'LmCompatibilityLevel'
        $lmBadge = if ($lmLevel -ge 5) { Format-Badge "Level $lmLevel ✓" 'ok' } else { Format-Badge "Level $lmLevel" 'warn' }
        $secRows.Add(('<tr><td>LM Compatibility Level</td><td>{0}</td><td>5 = NTLMv2 only</td><td>{1}</td><td>{2}</td><td>CIS L1 §2.3.11.7 / BSI</td></tr>' -f $lmLevel, $lmBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/windows/security/threat-protection/security-policy-settings/network-security-lan-manager-authentication-level' 'MS Learn')))
        $credGuard = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard' 'EnableVirtualizationBasedSecurity'
        $cgBadge = if ($credGuard -eq 0) { Format-Badge 'Disabled' 'ok' } else { Format-Badge 'Enabled (review)' 'warn' }
        $secRows.Add(('<tr><td>Credential Guard</td><td>EnableVBS = {0}</td><td>0 = Disabled (Exchange servers)</td><td>{1}</td><td>{2}</td><td>CIS L2</td></tr>' -f $credGuard, $cgBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/exchange/plan-and-deploy/virtualization' 'Exchange Virtualization')))
        $http2 = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\HTTP\Parameters' 'EnableHttp2Tls'
        $http2Badge = if ($http2 -eq 0) { Format-Badge 'Disabled' 'ok' } else { Format-Badge 'Enabled' 'warn' }
        $secRows.Add(('<tr><td>HTTP/2 over TLS</td><td>EnableHttp2Tls = {0}</td><td>0 = Disabled (MAPI/RPC compat)</td><td>{1}</td><td>{2}</td><td>MS Exchange</td></tr>' -f $http2, $http2Badge, (Format-RefLink 'https://techcommunity.microsoft.com/blog/exchange/released-2022-h1-cumulative-updates-for-exchange-server/3285026' 'Exchange Blog')))
        # Serialized Data Signing — registry value name: EnableSerializationDataSigning
        $serialSign = Get-SecRegVal 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Diagnostics' 'EnableSerializationDataSigning'
        $serialBadge = if ($serialSign -eq 1) { Format-Badge 'Enabled' 'ok' } else { Format-Badge 'Not set' 'warn' }
        $secRows.Add(('<tr><td>Serialized Data Signing</td><td>{0}</td><td>1 = Enabled</td><td>{1}</td><td>{2}</td><td>MS Exchange</td></tr>' -f $serialSign, $serialBadge, (Format-RefLink 'https://techcommunity.microsoft.com/blog/exchange/released-2022-h1-cumulative-updates-for-exchange-server/3285026' 'Exchange Blog')))
        $uacVal = Get-SecRegVal 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' 'EnableLUA'
        $uacBadge = if ($uacVal -eq 1 -or $null -eq $uacVal) { Format-Badge 'Enabled' 'ok' } else { Format-Badge 'Disabled!' 'fail' }
        $secRows.Add(('<tr><td>UAC (EnableLUA)</td><td>{0}</td><td>1 = Enabled (re-enabled after setup)</td><td>{1}</td><td>{2}</td><td>CIS L1 §17.2 / BSI SYS.2.1</td></tr>' -f $uacVal, $uacBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/windows/security/application-security/application-control/user-account-control/' 'MS Learn')))

        # IPv4 over IPv6 preference
        $ipv4Comp = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters' 'DisabledComponents'
        $ipv4ValText = if ($null -eq $ipv4Comp) { 'Not set (OS default)' } else { '0x{0:X}' -f [int]$ipv4Comp }
        $ipv4Badge = if ($ipv4Comp -eq 0x20) { Format-Badge '0x20 ✓' 'ok' } else { Format-Badge 'Not configured' 'warn' }
        $secRows.Add(('<tr><td>IPv4 over IPv6 preference</td><td>{0}</td><td>0x20 = prefer IPv4 (keep IPv6 loopback)</td><td>{1}</td><td>{2}</td><td>MS Exchange</td></tr>' -f $ipv4ValText, $ipv4Badge, (Format-RefLink 'https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/configure-ipv6-in-windows' 'Exchange Blog')))

        # NetBIOS over TCP/IP
        # SetTcpipNetbios(2) may return 1 (pending reboot), leaving the live CIM value stale.
        # Cross-check the registry (NetbiosOptions=2) so the report reflects the pending state.
        try {
            $nbNics = @(Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True' -ErrorAction Stop)
            $nbDisabled = 0
            foreach ($nbNic in $nbNics) {
                if ($nbNic.TcpipNetbiosOptions -eq 2) { $nbDisabled++; continue }
                $nbReg = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\NetBT\Parameters\Interfaces\Tcpip_$($nbNic.SettingID)" -Name 'NetbiosOptions' -ErrorAction SilentlyContinue).NetbiosOptions
                if ($nbReg -eq 2) { $nbDisabled++ }
            }
            $nbBadge = if ($nbNics.Count -gt 0 -and $nbDisabled -eq $nbNics.Count) { Format-Badge 'Disabled ✓' 'ok' } else { Format-Badge ('{0}/{1} disabled' -f $nbDisabled, $nbNics.Count) 'warn' }
            $secRows.Add(('<tr><td>NetBIOS over TCP/IP</td><td>{0} of {1} NICs disabled</td><td>Disabled on all NICs (reduces LLMNR/NBT-NS attack surface)</td><td>{2}</td><td>{3}</td><td>CIS §18 / BSI</td></tr>' -f $nbDisabled, $nbNics.Count, $nbBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/disable-netbios-tcp-ip-using-dhcp' 'MS Learn')))
        } catch { }

        # Root CA auto-update
        $rootAU = Get-SecRegVal 'HKLM:\SOFTWARE\Policies\Microsoft\SystemCertificates\AuthRoot' 'DisableRootAutoUpdate'
        $rootAUBadge = if ($rootAU -ne 1) { Format-Badge 'Enabled ✓' 'ok' } else { Format-Badge 'Disabled by policy!' 'warn' }
        $rootAUDisplay = if ($null -eq $rootAU) { '(not set — default enabled)' } else { 'DisableRootAutoUpdate = {0}' -f $rootAU }
        $secRows.Add(('<tr><td>Root CA Auto-Update</td><td>{0}</td><td>Must not be disabled (Exchange Online / M365 connectivity)</td><td>{1}</td><td>{2}</td><td>MS Exchange</td></tr>' -f $rootAUDisplay, $rootAUBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/security/trusted-root/release-notes' 'MS Trusted Root')))

        # Extended Protection (OWA VDir — evaluate Frontend only; Back End is internal and EP=None by design)
        if (-not $State['InstallEdge']) {
            try {
                $owaVdirs = @(Get-OwaVirtualDirectory -Server $env:COMPUTERNAME -ADPropertiesOnly -ErrorAction SilentlyContinue)
                $owaFe    = $owaVdirs | Where-Object { $_.Name -notlike '*Back End*' -and $_.WebSiteName -notlike '*Back End*' } | Select-Object -First 1
                if (-not $owaFe) { $owaFe = $owaVdirs | Select-Object -First 1 }
                if ($owaFe) {
                    $siteLabel = if ($owaFe.WebSiteName) { $owaFe.WebSiteName } else { $owaFe.Name }
                    $epVal = [string]$owaFe.ExtendedProtectionTokenChecking
                    if ([string]::IsNullOrEmpty($epVal)) { $epVal = 'None' }
                    # Normalize integer forms returned when deserializing AD properties
                    if ($epVal -eq '2') { $epVal = 'Require' } elseif ($epVal -eq '1') { $epVal = 'Allow' } elseif ($epVal -eq '0') { $epVal = 'None' }
                    $epBadge = if ($epVal -in 'Require','Allow') { Format-Badge "$epVal ✓" 'ok' } else { Format-Badge "$epVal (risk)" 'warn' }
                    $secRows.Add(('<tr><td>Extended Protection (OWA)</td><td>{0} — {1}</td><td>Require or Allow</td><td>{2}</td><td>{3}</td><td>MS Exchange</td></tr>' -f $siteLabel, $epVal, $epBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/exchange/plan-and-deploy/post-installation-tasks/security-best-practices/exchange-extended-protection' 'MS Learn')))
                }
            } catch { }
        }

        # SSL Offloading (Outlook Anywhere) — must be off for Extended Protection
        if (-not $State['InstallEdge']) {
            try {
                $oaVdir = Get-OutlookAnywhere -Server $env:COMPUTERNAME -ErrorAction SilentlyContinue
                if ($oaVdir) {
                    $oaBadge = if (-not $oaVdir.SSLOffloading) { Format-Badge 'Disabled ✓' 'ok' } else { Format-Badge 'Enabled (blocks EP)' 'warn' }
                    $secRows.Add(('<tr><td>SSL Offloading (Outlook Anywhere)</td><td>{0}</td><td>False (required for Extended Protection channel binding)</td><td>{1}</td><td>{2}</td><td>MS Exchange</td></tr>' -f $oaVdir.SSLOffloading, $oaBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/exchange/plan-and-deploy/post-installation-tasks/security-best-practices/exchange-extended-protection' 'MS Learn')))
                }
            } catch { }
        }

        # MRS Proxy (EWS)
        if ($State['InstallMailbox'] -and -not $State['InstallEdge']) {
            try {
                $ewsVdir = Get-WebServicesVirtualDirectory -Server $env:COMPUTERNAME -ADPropertiesOnly -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($ewsVdir) {
                    $mrsBadge = if (-not $ewsVdir.MRSProxyEnabled) { Format-Badge 'Disabled ✓' 'ok' } else { Format-Badge 'Enabled (review)' 'warn' }
                    $secRows.Add(('<tr><td>MRS Proxy (EWS)</td><td>{0}</td><td>False (enable only for cross-forest migrations)</td><td>{1}</td><td>{2}</td><td>MS Exchange</td></tr>' -f $ewsVdir.MRSProxyEnabled, $mrsBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/exchange/architecture/mailbox-servers/mrs-proxy-endpoint' 'MS Learn')))
                }
            } catch { }
        }

        # MAPI Encryption Required
        if ($State['InstallMailbox'] -and -not $State['InstallEdge']) {
            try {
                $mbxSrv = Get-MailboxServer -Identity $env:COMPUTERNAME -ErrorAction SilentlyContinue
                if ($mbxSrv) {
                    $mapiEncBadge = if ($mbxSrv.MAPIEncryptionRequired) { Format-Badge 'Required ✓' 'ok' } else { Format-Badge 'Not required' 'warn' }
                    $secRows.Add(('<tr><td>MAPI Encryption Required</td><td>{0}</td><td>True (forces encrypted Outlook MAPI connections)</td><td>{1}</td><td>{2}</td><td>MS Exchange</td></tr>' -f $mbxSrv.MAPIEncryptionRequired, $mapiEncBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/exchange/clients/mapi-over-http/configure-mapi-over-http' 'MS Learn')))
                }
            } catch { }
        }

        # SMTP Banner hardening
        if (-not $State['InstallEdge']) {
            try {
                $feBannerConns = @(Get-ReceiveConnector -Server $env:COMPUTERNAME -ErrorAction SilentlyContinue | Where-Object { $_.TransportRole -eq 'FrontendTransport' })
                if ($feBannerConns.Count -gt 0) {
                    $bannersHardened = ($feBannerConns | Where-Object { $_.Banner -and $_.Banner -notlike '*Microsoft ESMTP*' }).Count
                    $bannerBadge = if ($bannersHardened -eq $feBannerConns.Count) { Format-Badge 'Hardened ✓' 'ok' } else { Format-Badge ('{0}/{1} hardened' -f $bannersHardened, $feBannerConns.Count) 'warn' }
                    $secRows.Add(('<tr><td>SMTP Banner</td><td>{0}/{1} Frontend connectors hardened</td><td>Generic banner (hides Exchange version from attackers)</td><td>{2}</td><td>{3}</td><td>CIS / DISA STIG</td></tr>' -f $bannersHardened, $feBannerConns.Count, $bannerBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/exchange/mail-flow/connectors/receive-connectors' 'MS Learn')))
                }
            } catch { }
        }

        $secContent = '<table class="data-table"><tr><th>Setting</th><th>Current Value</th><th>Exchange Recommendation</th><th>Status</th><th>Reference</th><th>CIS / BSI</th></tr>' + ($secRows -join '') + '</table>'

        # ── 6. PERFORMANCE SETTINGS (with best-practice column) ───────────────
        $perfRows = [System.Collections.Generic.List[string]]::new()
        try {
            $plan = Get-CimInstance -Namespace 'root\cimv2\power' -ClassName Win32_PowerPlan -Filter 'IsActive=True' -ErrorAction Stop
            $isHighPerf = $plan.InstanceID -like "*$POWERPLAN_HIGH_PERFORMANCE*"
            $planBadge  = if ($isHighPerf) { Format-Badge 'High Performance ✓' 'ok' } else { Format-Badge 'Not High Performance' 'warn' }
            $perfRows.Add(('<tr><td>Power Plan</td><td>{0}</td><td>High Performance</td><td>{1}</td></tr>' -f $plan.ElementName, $planBadge))
        } catch { }
        try {
            $pf = Get-CimInstance Win32_PageFileSetting -ErrorAction Stop | Select-Object -First 1
            if ($pf) {
                $ramMB        = [math]::Round((Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue).TotalPhysicalMemory / 1MB, 0)
                $recMB        = if ($State['MajorSetupVersion'] -ge $EX2019_MAJOR) { [int]($ramMB * 0.25) } else { [math]::Min($ramMB + 10, 32768 + 10) }
                $pfOk         = $pf.InitialSize -ge $recMB -and $pf.MaximumSize -ge $recMB
                $pfBadge      = if ($pfOk) { Format-Badge '✓' 'ok' } else { Format-Badge 'Below recommendation' 'warn' }
                $perfRows.Add(('<tr><td>Pagefile</td><td>{0} — Init: {1} MB / Max: {2} MB</td><td>≥ {3} MB</td><td>{4}</td></tr>' -f $pf.Name, $pf.InitialSize, $pf.MaximumSize, $recMB, $pfBadge))
            }
        } catch { }
        $keepAlive = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters' 'KeepAliveTime'
        if ($keepAlive) {
            $kaBadge = if ($keepAlive -le 900000) { Format-Badge '✓' 'ok' } else { Format-Badge 'High' 'warn' }
            $perfRows.Add(('<tr><td>TCP KeepAliveTime</td><td>{0} ms ({1} min)</td><td>900000 ms (15 min)</td><td>{2}</td></tr>' -f $keepAlive, [math]::Round($keepAlive / 60000, 0), $kaBadge))
        }
        try {
            Get-NetAdapterRss -ErrorAction SilentlyContinue | Where-Object { $_.Enabled } | ForEach-Object {
                $perfRows.Add(('<tr><td>RSS: {0}</td><td>Enabled — Queues: {1}</td><td>Enabled; Queues = physical cores</td><td>{2}</td></tr>' -f $_.Name, $_.NumberOfReceiveQueues, (Format-Badge 'Enabled ✓' 'ok')))
            }
        } catch { }
        $maxConcAPI = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters' 'MaxConcurrentApi'
        if ($maxConcAPI) {
            $mcaBadge = if ($maxConcAPI -ge 10) { Format-Badge '✓' 'ok' } else { Format-Badge 'Low' 'warn' }
            $perfRows.Add(('<tr><td>Netlogon MaxConcurrentApi</td><td>{0}</td><td>&ge; logical cores (min 10)</td><td>{1}</td></tr>' -f $maxConcAPI, $mcaBadge))
        }
        $ctsPct = Get-SecRegVal 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Search\SystemParameters' 'CtsProcessorAffinityPercentage'
        if ($null -ne $ctsPct) {
            $ctsBadge = if ($ctsPct -eq 0) { Format-Badge '0% ✓' 'ok' } else { Format-Badge "$ctsPct%" 'warn' }
            $perfRows.Add(('<tr><td>CTS Processor Affinity %</td><td>{0}</td><td>0 (Exchange Search best practice)</td><td>{1}</td></tr>' -f $ctsPct, $ctsBadge))
        }
        $tcpOffload = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters' 'DisableTaskOffload'
        if ($null -ne $tcpOffload) {
            $toBadge = if ($tcpOffload -eq 1) { Format-Badge 'Disabled ✓' 'ok' } else { Format-Badge 'Enabled' 'warn' }
            $perfRows.Add(('<tr><td>TCP Task Offload</td><td>DisableTaskOffload = {0}</td><td>1 = Disabled</td><td>{1}</td></tr>' -f $tcpOffload, $toBadge))
        }
        $wSearch = (Get-Service -Name WSearch -ErrorAction SilentlyContinue)
        if ($wSearch) {
            $wsBadge = if ($wSearch.StartType -eq 'Disabled') { Format-Badge 'Disabled ✓' 'ok' } else { Format-Badge $wSearch.StartType 'warn' }
            $perfRows.Add(('<tr><td>Windows Search Service</td><td>{0}</td><td>Disabled (Exchange uses own indexing)</td><td>{1}</td></tr>' -f $wSearch.StartType, $wsBadge))
        }

        # RPC Minimum Connection Timeout
        $rpcTO = Get-SecRegVal 'HKLM:\SOFTWARE\Microsoft\Rpc' 'MinimumConnectionTimeout'
        if ($null -ne $rpcTO) {
            $rpcBadge = if ($rpcTO -ge 120) { Format-Badge "✓ ${rpcTO}s" 'ok' } else { Format-Badge "${rpcTO}s (low)" 'warn' }
            $perfRows.Add(('<tr><td>RPC Min Connection Timeout</td><td>{0} s</td><td>120 s (prevents premature RPC timeouts under load)</td><td>{1}</td></tr>' -f $rpcTO, $rpcBadge))
        }

        # TCP Chimney Offload
        $tcpChim = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters' 'EnableTCPChimney'
        if ($null -ne $tcpChim) {
            $chimBadge = if ($tcpChim -eq 0) { Format-Badge 'Disabled ✓' 'ok' } else { Format-Badge 'Enabled' 'warn' }
            $perfRows.Add(('<tr><td>TCP Chimney Offload</td><td>EnableTCPChimney = {0}</td><td>0 = Disabled (Microsoft recommendation for Exchange)</td><td>{1}</td></tr>' -f $tcpChim, $chimBadge))
        }

        # NodeRunner Max Memory
        $nodeRunMem = Get-SecRegVal 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Search\SystemParameters' 'NodeRunnerMaxMemory'
        if ($null -ne $nodeRunMem) {
            $nrBadge = if ($nodeRunMem -eq 0) { Format-Badge 'Unlimited ✓' 'ok' } else { Format-Badge "${nodeRunMem} MB (capped)" 'warn' }
            $perfRows.Add(('<tr><td>NodeRunner Max Memory</td><td>{0}</td><td>0 = unlimited (Exchange Search best practice)</td><td>{1}</td></tr>' -f $nodeRunMem, $nrBadge))
        }

        $perfContent = '<table class="data-table"><tr><th>Setting</th><th>Current Value</th><th>Exchange Recommendation</th><th>Status</th></tr>' + ($perfRows -join '') + '</table>'

        # ── 7. RBAC ROLE GROUP MEMBERSHIP ─────────────────────────────────────
        $rbacGroups = @(
            'Organization Management', 'Server Management', 'Recipient Management',
            'Help Desk', 'Hygiene Management', 'Compliance Management',
            'Records Management', 'Discovery Management',
            'Public Folder Management', 'View-Only Organization Management'
        )
        $rbacRows = [System.Collections.Generic.List[string]]::new()
        $rbacRows.Add('<tr><th>Role Group</th><th>Members</th></tr>')
        foreach ($group in $rbacGroups) {
            try {
                $members = @(Get-RoleGroupMember -Identity $group -ErrorAction Stop)
                $memberHtml = if ($members.Count -gt 0) {
                    ($members | ForEach-Object { '<code>{0}</code> <span style="color:#888;font-size:11px">({1})</span>' -f [string]$_.Name, [string]$_.RecipientType }) -join '<br>'
                } else { '<em style="color:#888">no members</em>' }
                $rbacRows.Add(('<tr><td>{0}</td><td>{1}</td></tr>' -f $group, $memberHtml))
            }
            catch {
                $rbacRows.Add(('<tr><td>{0}</td><td style="color:#c50f1f"><em>Query failed: {1}</em></td></tr>' -f $group, $_.Exception.Message))
            }
        }
        $rbacContent = '<table class="data-table">' + ($rbacRows -join '') + '</table>'

        # ── 9. INSTALLATION LOG ───────────────────────────────────────────────
        # B16 fixes:
        # 1. ReadAllText was called with UTF-8 encoding; PS 5.1 transcripts are UTF-16 LE —
        #    removed explicit encoding so .NET auto-detects the BOM (handles both UTF-8/UTF-16).
        # 2. Wrapped in try/catch — an IOExceptionon a large/locked transcript file propagated
        #    to the global trap { break } and killed the entire script.
        # 3. Capped to the last 2000 lines — a transcript accumulated over multiple reboots
        #    (30h+ install) can be several MB; embedding the full file makes the HTML report
        #    unusably large and strains memory during the regex-replace operations.
        $logContent = if ($State['TranscriptFile'] -and (Test-Path $State['TranscriptFile'])) {
            try {
                $logLines   = [System.IO.File]::ReadAllLines($State['TranscriptFile'])
                $totalLines = $logLines.Count
                $maxLines   = 2000
                $truncated  = $totalLines -gt $maxLines
                $logLines   = if ($truncated) { $logLines[($totalLines - $maxLines)..($totalLines - 1)] } else { $logLines }
                $logText    = $logLines -join "`n"
                $logEscaped = $logText -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
                $truncNote  = if ($truncated) { '<div style="color:#ff9800;font-size:11px;margin-bottom:6px">&#9888; Log truncated — showing last {0} of {1} lines. Full log: <code>{2}</code></div>' -f $maxLines, $totalLines, $State['TranscriptFile'] } else { '' }
                $truncNote + '<pre style="font-family:Consolas,monospace;font-size:12px;line-height:1.5;background:#1e1e1e;color:#d4d4d4;padding:16px;border-radius:4px;overflow:auto;max-height:600px;white-space:pre-wrap;word-break:break-all">' + $logEscaped + '</pre>'
            } catch {
                '<p style="color:#8a8886"><em>Log file could not be read ({0}): {1}</em></p>' -f $State['TranscriptFile'], $_.Exception.Message
            }
        } else {
            '<p style="color:#8a8886"><em>Log file not found: {0}</em></p>' -f $State['TranscriptFile']
        }

        # ── BUILD HTML ────────────────────────────────────────────────────────
        $css = @'
:root{--primary:#1a2332;--accent:#0078d4;--ok:#107c10;--warn:#d83b01;--fail:#c50f1f;--bg:#f3f2f1;--card:#fff;--border:#e1dfdd}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:var(--bg);color:#252423;font-size:14px}
header{background:var(--primary);color:#fff;padding:24px 40px;display:flex;align-items:center;gap:16px}
header h1{font-size:22px;font-weight:300;letter-spacing:.5px}
.logo{width:44px;height:44px;background:var(--accent);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:18px;font-weight:700;flex-shrink:0}
.summary-bar{background:var(--accent);color:#fff;padding:12px 40px;display:flex;gap:40px;font-size:13px;flex-wrap:wrap}
.summary-bar span{opacity:.8}.summary-bar strong{opacity:1;font-weight:600}
.container{display:flex}
.toc{width:210px;min-width:210px;background:var(--primary);padding:20px 0;position:sticky;top:0;height:100vh;overflow-y:auto;flex-shrink:0}
.toc-title{color:#888;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:1px;padding:16px 20px 6px}
.toc a{display:block;padding:9px 20px;color:#c8c8c8;text-decoration:none;font-size:13px;border-left:3px solid transparent;transition:all .15s}
.toc a:hover{color:#fff;background:rgba(255,255,255,.08);border-left-color:var(--accent)}
main{flex:1;padding:32px 36px;max-width:1200px;overflow-x:auto}
.section{background:var(--card);border-radius:8px;margin-bottom:24px;box-shadow:0 1px 4px rgba(0,0,0,.08);overflow:hidden}
.section-title{background:var(--primary);color:#fff;padding:12px 20px;font-size:15px;font-weight:400;letter-spacing:.3px}
.section-body{padding:20px}
.subsection{margin:22px 0 10px;font-size:13px;font-weight:700;color:var(--primary);border-bottom:2px solid var(--border);padding-bottom:6px;text-transform:uppercase;letter-spacing:.4px}
.data-table{width:100%;border-collapse:collapse;font-size:13px}
.data-table th{background:#f3f2f1;font-weight:600;text-align:left;padding:8px 12px;border-bottom:2px solid var(--border);color:var(--primary)}
.data-table td{padding:7px 12px;border-bottom:1px solid #f0efee;vertical-align:middle}
.data-table tr:last-child td{border-bottom:none}
.data-table tr:hover td{background:#faf9f8}
code{font-family:Consolas,monospace;font-size:12px;background:#f0f0f0;padding:1px 5px;border-radius:3px;word-break:break-all}
footer{background:var(--primary);color:#888;padding:16px 40px;font-size:12px;text-align:center;margin-top:8px}
@media print{
  @page{size:A4 portrait;margin:1.8cm 1.5cm 2cm 1.5cm}
  @page:first{margin-top:1cm}
  *{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}
  .toc,.summary-bar{display:none!important}
  body{background:#fff;font-size:10pt}
  header{padding:12px 20px;break-after:avoid}
  header h1{font-size:14pt}
  .container{display:block}
  main{padding:0;max-width:100%}
  .section{background:#fff;box-shadow:none;border:1pt solid #ccc;margin-bottom:14pt;page-break-inside:avoid;break-inside:avoid}
  .section-title{padding:7pt 12pt;font-size:11pt}
  .section-body{padding:10pt 12pt}
  .subsection{font-size:9pt;margin:14pt 0 6pt}
  .data-table{font-size:8.5pt;width:100%;table-layout:fixed}
  .data-table th,.data-table td{padding:4pt 6pt;word-break:break-word}
  .data-table tr{page-break-inside:avoid;break-inside:avoid}
  code{font-size:7.5pt;word-break:break-all}
  #installlog pre{max-height:none;font-size:6.5pt;page-break-inside:auto}
  #healthchecker iframe{display:none}
  #healthchecker::after{content:"HealthChecker report is embedded in the HTML version of this document.";font-style:italic;color:#666;font-size:10pt}
  a{color:inherit;text-decoration:none}
  footer{font-size:8pt;padding:8pt 20pt;margin-top:6pt}
  /* Print document header on continuation pages via border-top trick */
  .section:nth-child(n+2){border-top:2pt solid #1a2332}
}
'@

        # ── 9. HEALTHCHECKER (embedded via srcdoc — no external file dependency) ──
        $hcContent = if ($State['SkipHealthCheck']) {
            '<p style="color:#666;font-size:13px;">HealthChecker was skipped (<code>-SkipHealthCheck</code> specified).</p>'
        } elseif ($State['HCReportPath'] -and (Test-Path $State['HCReportPath'])) {
            try {
                $hcRaw     = [System.IO.File]::ReadAllText($State['HCReportPath'], [System.Text.Encoding]::UTF8)
                # Escape for HTML attribute value: & → &amp;  " → &quot;
                $hcSrcdoc  = $hcRaw -replace '&', '&amp;' -replace '"', '&quot;'
                '<iframe srcdoc="' + $hcSrcdoc + '" style="width:100%;height:860px;border:1px solid #e1dfdd;border-radius:4px;" title="HealthChecker Report" sandbox="allow-same-origin allow-scripts allow-popups"></iframe>'
            }
            catch {
                '<p style="color:#d83b01;font-size:13px;">HealthChecker report could not be embedded: {0}</p>' -f $_.Exception.Message
            }
        } else {
            '<p style="color:#d83b01;font-size:13px;">HealthChecker report not found — HC may have failed to run. Check the installation log for details.</p>'
        }

        # ── 0. MANAGEMENT SUMMARY ─────────────────────────────────────────────
        $mgmtRows = [System.Collections.Generic.List[string]]::new()

        # Count security badge types from built HTML string
        $secOK   = ([regex]::Matches($secContent,   'background:#107c10')).Count
        $secWarn = ([regex]::Matches($secContent,   'background:#d83b01')).Count
        $secFail = ([regex]::Matches($secContent,   'background:#c50f1f')).Count
        $perfOK  = ([regex]::Matches($perfContent,  'background:#107c10')).Count
        $perfWarn= ([regex]::Matches($perfContent,  'background:#d83b01')).Count

        $secStatusBadge  = if ($secFail -gt 0)  { Format-Badge "$secFail critical, $secWarn warnings" 'fail' }
                           elseif ($secWarn -gt 0) { Format-Badge "$secWarn warnings" 'warn' }
                           else { Format-Badge "All $secOK items OK" 'ok' }
        $perfStatusBadge = if ($perfWarn -gt 0) { Format-Badge "$perfWarn items to review" 'warn' } else { Format-Badge "All $perfOK items OK" 'ok' }

        $certStatus = try {
            $certs = @(Get-ExchangeCertificate -Server $env:COMPUTERNAME -ErrorAction Stop | Where-Object { $_.Thumbprint -and $_.NotAfter -gt [datetime]'1970-01-01' })
            $expiring = @($certs | Where-Object { [int][Math]::Floor(($_.NotAfter - (Get-Date)).TotalDays) -le 90 })
            if ($expiring.Count -gt 0) { Format-Badge ('{0} expiring within 90 days' -f $expiring.Count) 'warn' }
            else { Format-Badge ('{0} certificate(s) — all valid' -f $certs.Count) 'ok' }
        } catch { Format-Badge 'Could not query' 'na' }

        $vdirStatus = if ($State['Namespace']) { Format-Badge ('Configured ({0})' -f $State['Namespace']) 'ok' } else { Format-Badge 'Not configured' 'warn' }
        $suStatus   = if ($State['IncludeFixes']) { Format-Badge 'Enabled' 'ok' } else { Format-Badge 'Skipped' 'warn' }
        $hcStatus   = if ($State['SkipHealthCheck']) { Format-Badge 'Skipped' 'na' } elseif ($State['HCReportPath']) { Format-Badge 'Completed — see section below' 'ok' } else { Format-Badge 'Failed / not found' 'warn' }
        $modeStatus = if ($State['ConfigDriven']) { Format-Badge 'Autopilot (fully automated)' 'info' } else { Format-Badge 'Copilot (interactive)' 'info' }

        $mgmtRows.Add(('<tr><td style="width:220px"><strong>Exchange Version</strong></td><td>{0}</td><td>{1}</td></tr>' -f $exVersion, (Format-Badge 'Installed' 'ok')))
        $serverFqdn = try { '{0}.{1}' -f (Get-CimInstance Win32_ComputerSystem -EA SilentlyContinue).DNSHostName, (Get-CimInstance Win32_ComputerSystem -EA SilentlyContinue).Domain } catch { '' }
        $mgmtRows.Add(('<tr><td><strong>Server</strong></td><td>{0}</td><td>{1}</td></tr>' -f ('{0} ({1})' -f $env:COMPUTERNAME, $serverFqdn), (Format-Badge 'OK' 'ok')))
        $mgmtRows.Add('<tr><td><strong>Organization</strong></td><td>{0}</td><td></td></tr>' -f $State['OrganizationName'])
        $mgmtRows.Add(('<tr><td><strong>Installation Mode</strong></td><td>{0}</td><td>{1}</td></tr>' -f $instMode, $modeStatus))
        $mgmtRows.Add(('<tr><td><strong>Virtual Directory URLs</strong></td><td>{0}</td><td>{1}</td></tr>' -f $State['Namespace'], $vdirStatus))
        $mgmtRows.Add(('<tr><td><strong>Security Hardening</strong></td><td>{0} OK / {1} warnings / {2} critical</td><td>{3}</td></tr>' -f $secOK, $secWarn, $secFail, $secStatusBadge))
        $mgmtRows.Add(('<tr><td><strong>Performance Settings</strong></td><td>{0} OK / {1} to review</td><td>{2}</td></tr>' -f $perfOK, $perfWarn, $perfStatusBadge))
        $mgmtRows.Add('<tr><td><strong>Certificates</strong></td><td></td><td>{0}</td></tr>' -f $certStatus)
        $mgmtRows.Add('<tr><td><strong>Security Updates</strong></td><td></td><td>{0}</td></tr>' -f $suStatus)
        $mgmtRows.Add('<tr><td><strong>HealthChecker</strong></td><td></td><td>{0}</td></tr>' -f $hcStatus)
        $mgmtRows.Add('<tr><td><strong>Report Generated</strong></td><td>{0}</td><td></td></tr>' -f $reportDate)

        # Action items — surface any WARN/FAIL as bullet list
        $actionItems = [System.Collections.Generic.List[string]]::new()
        if ($secFail -gt 0)  { $actionItems.Add('<li><strong>Security:</strong> {0} critical finding(s) require immediate attention — see Security Settings section.</li>' -f $secFail) }
        if ($secWarn -gt 0)  { $actionItems.Add('<li><strong>Security:</strong> {0} warning(s) — review Security Settings section.</li>' -f $secWarn) }
        if ($perfWarn -gt 0) { $actionItems.Add('<li><strong>Performance:</strong> {0} setting(s) below recommendation — review Performance &amp; Tuning section.</li>' -f $perfWarn) }
        if (-not $State['Namespace']) { $actionItems.Add('<li><strong>Virtual Directories:</strong> No access namespace configured. OWA/ECP/EWS URLs may still point to server hostname.</li>') }
        if (-not $State['IncludeFixes']) { $actionItems.Add('<li><strong>Security Updates:</strong> Exchange Security Update installation was skipped. Apply the latest SU manually.</li>') }

        $actionHtml = if ($actionItems.Count -gt 0) {
            '<h3 class="subsection">Action Items</h3><ul style="margin:8px 0 0 20px;line-height:1.8;font-size:13px">' + ($actionItems -join '') + '</ul>'
        } else {
            '<p style="color:#107c10;font-size:13px;margin-top:12px">&#10003; No critical action items identified.</p>'
        }

        $mgmtContent = '<table class="data-table"><tr><th>Item</th><th>Detail</th><th>Status</th></tr>' + ($mgmtRows -join '') + '</table>' + $actionHtml

        $sections = @(
            (New-HtmlSection 'summary'      'Management Summary'        $mgmtContent)
            (New-HtmlSection 'params'       'Installation Parameters'   ('<table class="data-table">' + ($instRows -join '') + '</table>'))
            (New-HtmlSection 'system'       'System Information'        $sysContent)
            (New-HtmlSection 'ad'           'Active Directory'          $adContent)
            (New-HtmlSection 'exchange'     'Exchange Configuration'    $exContent)
            (New-HtmlSection 'security'     'Security Settings'         $secContent)
            (New-HtmlSection 'performance'  'Performance &amp; Tuning'  $perfContent)
            (New-HtmlSection 'rbac'         'RBAC Role Group Membership' $rbacContent)
            (New-HtmlSection 'installlog'   'Installation Log'          $logContent)
            (New-HtmlSection 'healthchecker' 'HealthChecker'            $hcContent)
        )

        $toc = @(
            '<div class="toc-title">Contents</div>'
            '<a href="#summary">Management Summary</a>'
            '<a href="#params">Installation Parameters</a>'
            '<a href="#system">System Information</a>'
            '<a href="#ad">Active Directory</a>'
            '<a href="#exchange">Exchange Configuration</a>'
            '<a href="#security">Security Settings</a>'
            '<a href="#performance">Performance &amp; Tuning</a>'
            '<a href="#rbac">RBAC Role Groups</a>'
            '<a href="#installlog">Installation Log</a>'
            '<a href="#healthchecker">HealthChecker</a>'
        )

        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Exchange Installation Report — $env:COMPUTERNAME</title>
<style>$css</style>
</head>
<body>
<header>
  <div class="logo">Ex</div>
  <div>
    <h1>Exchange Server Installation Report</h1>
    <div style="font-size:12px;opacity:.65;margin-top:4px">Generated by EXpress v$ScriptVersion</div>
  </div>
</header>
<div class="summary-bar">
  <div><span>Server: </span><strong>$env:COMPUTERNAME</strong></div>
  <div><span>Exchange: </span><strong>$exVersion</strong></div>
  <div><span>Organization: </span><strong>$($State['OrganizationName'])</strong></div>
  <div><span>Report Date: </span><strong>$reportDate</strong></div>
</div>
<div class="container">
  <nav class="toc">$($toc -join '')</nav>
  <main>$($sections -join '')</main>
</div>
<footer>Exchange Server Installation Report &bull; $env:COMPUTERNAME &bull; $reportDate &bull; EXpress v$ScriptVersion</footer>
</body>
</html>
"@

        try {
            $html | Out-File -FilePath $reportPath -Encoding utf8 -ErrorAction Stop
            Write-MyOutput ('Installation Report saved to {0}' -f $reportPath)
        }
        catch {
            Write-MyWarning ('Could not write Installation Report: {0}' -f $_.Exception.Message)
            return
        }

        # Optional PDF via Microsoft Edge headless
        $edgeExe = @(
            "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
            "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
        ) | Where-Object { Test-Path $_ } | Select-Object -First 1

        if ($edgeExe) {
            $pdfPath    = $reportPath -replace '\.html$', '.pdf'
            $edgeStdErr = Join-Path $env:TEMP 'edge_headless_stderr.txt'
            Write-MyVerbose ('Generating PDF via Edge headless: {0}' -f $pdfPath)
            try {
                $fileUri  = 'file:///{0}' -f ($reportPath -replace '\\', '/')
                $edgeArgs = '--headless', '--disable-gpu', '--run-all-compositor-stages-before-draw',
                            '--log-level=3',        # suppress DevTools/renderer noise
                            '--disable-extensions', # no extension interference
                            '--print-to-pdf-no-header',  # remove browser URL/date chrome
                            '--virtual-time-budget=8000', # allow srcdoc iframe time to render
                            "--print-to-pdf=`"$pdfPath`"", "`"$fileUri`""
                $proc = Start-Process -FilePath $edgeExe -ArgumentList $edgeArgs -NoNewWindow -Wait -PassThru `
                            -RedirectStandardError $edgeStdErr -ErrorAction Stop
                Remove-Item $edgeStdErr -ErrorAction SilentlyContinue
                if ($proc.ExitCode -eq 0 -and (Test-Path $pdfPath)) {
                    Write-MyOutput ('Installation Report PDF saved to {0}' -f $pdfPath)
                }
                else {
                    Write-MyVerbose ('Edge PDF export exit code: {0}' -f $proc.ExitCode)
                }
            }
            catch {
                Write-MyVerbose ('Edge PDF generation failed: {0}' -f $_.Exception.Message)
            }
        }
        else {
            Write-MyVerbose 'Microsoft Edge not found — skipping PDF export. Open the HTML report in a browser and use File > Print > Save as PDF.'
        }
    }

    # ── OpenXML Engine (shared with tools/Build-ConceptTemplate.ps1) ─────────────
    # Pure PowerShell, no Office/COM required. PS2Exe-safe.

    function Invoke-XmlEscape { param([string]$Text) [Security.SecurityElement]::Escape([string]$Text) }

    function New-WdHeading {
        param([string]$Text, [int]$Level = 1)
        '<w:p><w:pPr><w:pStyle w:val="Heading{0}"/></w:pPr><w:r><w:t xml:space="preserve">{1}</w:t></w:r></w:p>' -f $Level, (Invoke-XmlEscape $Text)
    }
    function New-WdParagraph {
        param([string]$Text)
        if (-not $Text) { return '<w:p/>' }
        '<w:p><w:r><w:t xml:space="preserve">{0}</w:t></w:r></w:p>' -f (Invoke-XmlEscape $Text)
    }
    function New-WdPageBreak { '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' }
    function New-WdCentered {
        # Centered paragraph with configurable size (half-points) and optional bold.
        param([string]$Text, [int]$SizeHalfPt = 22, [bool]$Bold = $false, [string]$Color = '1F3864')
        $boldTag = if ($Bold) { '<w:b/>' } else { '' }
        $sb = '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:after="120"/></w:pPr><w:r><w:rPr>{0}<w:color w:val="{1}"/><w:sz w:val="{2}"/></w:rPr><w:t xml:space="preserve">{3}</w:t></w:r></w:p>' -f $boldTag, $Color, $SizeHalfPt, (Invoke-XmlEscape $Text)
        return $sb
    }
    function New-WdSpacer {
        # Vertical spacer paragraph (empty paragraph with configurable top spacing in twentieths of a point).
        param([int]$SpaceBefore = 240)
        '<w:p><w:pPr><w:spacing w:before="{0}" w:after="0"/></w:pPr></w:p>' -f $SpaceBefore
    }
    function New-WdToc {
        # Dynamic Table of Contents field. Word shows "Right-click → Update Field" or F9 after opening.
        # Levels 1-3 covers Heading1/2/3; \h = hyperlinks, \z = hide tab in web view, \u = use outline levels.
        param([string]$Title = 'Inhaltsverzeichnis')
        $titlePara = '<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t xml:space="preserve">{0}</w:t></w:r></w:p>' -f (Invoke-XmlEscape $Title)
        $tocField  = '<w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \o &quot;1-3&quot; \h \z \u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:i/><w:color w:val="808080"/></w:rPr><w:t xml:space="preserve">(Rechtsklick → Felder aktualisieren bzw. F9, um das Inhaltsverzeichnis zu aktualisieren)</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
        return $titlePara + $tocField
    }
    function New-WdBullet {
        param([string]$Text, [int]$Level = 0)
        '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="{0}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">{1}</w:t></w:r></w:p>' -f $Level, (Invoke-XmlEscape $Text)
    }
    function New-WdCode {
        param([string]$Text)
        '<w:p><w:pPr><w:pStyle w:val="Code"/><w:spacing w:after="120"/></w:pPr><w:r><w:t xml:space="preserve">{0}</w:t></w:r></w:p>' -f (Invoke-XmlEscape $Text)
    }
    function New-WdTable {
        # -Compact: shrinks runs to 8pt (half-point size = 16) for wide tables.
        # Word auto-layout distributes columns proportionally across the page width; with
        # long content in 6+ columns each cell wraps aggressively at 11pt default. 8pt
        # gives ~40% more horizontal characters per line and lifts most wrap to a single
        # break instead of cascading wraps on every column.
        param([string[]]$Headers, [object[]]$Rows, [switch]$Compact)
        $sb = [System.Text.StringBuilder]::new()
        $null = $sb.Append('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>')
        $colCount = if ($Headers) { $Headers.Count } else { 0 }
        # Font-size half-points: 22 = 11pt (default), 16 = 8pt (compact).
        $szHalfPt  = if ($Compact) { 16 } else { 22 }
        $cellRPr   = if ($Compact) { '<w:rPr><w:sz w:val="{0}"/></w:rPr>' -f $szHalfPt } else { '' }
        $headerRPr = if ($Compact) { '<w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="{0}"/></w:rPr>' -f $szHalfPt } else { '<w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>' }
        if ($Headers) {
            $null = $sb.Append('<w:tr><w:trPr><w:tblHeader/></w:trPr>')
            foreach ($h in $Headers) {
                $null = $sb.Append('<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="2F5496"/></w:tcPr>')
                $null = $sb.Append(('<w:p><w:r>{0}<w:t xml:space="preserve">{1}</w:t></w:r></w:p></w:tc>' -f $headerRPr, (Invoke-XmlEscape $h)))
            }
            $null = $sb.Append('</w:tr>')
        }
        # PS 5.1 flattens `@( @(a,b), @(c,d) )` literals to `@(a,b,c,d)` before the array is
        # bound to this parameter. Detect that case by scanning for any array-typed element; if
        # all elements are scalars and the total count is a multiple of $colCount, reshape into
        # rows of $colCount cells. Callers who pass `List[object[]].ToArray()` or use `,@(...)`
        # per row are unaffected because their rows remain array-typed.
        if ($Rows -and $colCount -gt 1 -and $Rows.Count -gt 0) {
            $anyArrayRow = $false
            foreach ($r in $Rows) {
                if ($null -ne $r -and -not ($r -is [string]) -and ($r -is [System.Collections.IEnumerable])) { $anyArrayRow = $true; break }
            }
            if (-not $anyArrayRow -and ($Rows.Count % $colCount -eq 0)) {
                $reshaped = New-Object 'System.Collections.Generic.List[object[]]'
                for ($i = 0; $i -lt $Rows.Count; $i += $colCount) {
                    $buf = New-Object 'object[]' $colCount
                    for ($j = 0; $j -lt $colCount; $j++) { $buf[$j] = $Rows[$i + $j] }
                    $reshaped.Add($buf)
                }
                $Rows = $reshaped.ToArray()
            }
        }
        foreach ($row in $Rows) {
            # Callers that forget the `,@(...)` prefix on literal jagged arrays cause PS 5.1
            # to flatten the outer @(...), so each $row arrives as a scalar string instead of
            # a row array. Normalize both cases here to avoid emitting ragged tables, which
            # some Word versions flag as invalid and refuse to render past that point.
            $cells = @($row)
            $null = $sb.Append('<w:tr>')
            foreach ($cell in $cells) {
                $cellStr = [string]$cell
                if ($cellStr -match "`n") {
                    # Multi-line cell: split at newlines, render each line as a separate run
                    # with <w:br/> between them. Font shrunk to 9pt (18 half-pt) so long paths
                    # fit on one line each.
                    $mlSz  = if ($Compact) { $szHalfPt } else { 18 }
                    $mlRPr = '<w:rPr><w:sz w:val="{0}"/></w:rPr>' -f $mlSz
                    $brRun = '<w:r>{0}<w:br/></w:r>' -f $mlRPr
                    $lines = $cellStr -split "`n"
                    $runs  = ($lines | ForEach-Object { '<w:r>{0}<w:t xml:space="preserve">{1}</w:t></w:r>' -f $mlRPr, (Invoke-XmlEscape $_) }) -join $brRun
                    $null  = $sb.Append(('<w:tc><w:p>{0}</w:p></w:tc>' -f $runs))
                }
                else {
                    $null = $sb.Append(('<w:tc><w:p><w:r>{0}<w:t xml:space="preserve">{1}</w:t></w:r></w:p></w:tc>' -f $cellRPr, (Invoke-XmlEscape $cellStr)))
                }
            }
            # Pad short rows to header width so cell counts match the first/header row.
            for ($pad = $cells.Count; $pad -lt $colCount; $pad++) {
                $null = $sb.Append(('<w:tc><w:p><w:r>{0}<w:t xml:space="preserve"></w:t></w:r></w:p></w:tc>' -f $cellRPr))
            }
            $null = $sb.Append('</w:tr>')
        }
        $null = $sb.Append('</w:tbl>')
        $sb.ToString()
    }
    function New-WdDocumentXml {
        param([string[]]$BodyParts)
        $body = $BodyParts -join "`n"
        @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
$body
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId3"/>
      <w:footerReference w:type="default" r:id="rId4"/>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1800" w:header="709" w:footer="709" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"@
    }
    function New-WdFile {
        param([string]$OutputPath, [string[]]$BodyParts, [string]$DocTitle = '', [string]$HeaderLabel = '', [string]$LogoPath = '')
        Add-Type -AssemblyName System.IO.Compression
        $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
        $hasLogo = $LogoPath -and (Test-Path $LogoPath -PathType Leaf)
        $fs  = [System.IO.File]::Open($OutputPath, [System.IO.FileMode]::Create)
        $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        function Add-ZipEntry([string]$name, [string]$content) {
            $entry  = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
            $stream = $entry.Open()
            $bytes  = $utf8NoBom.GetBytes($content)
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Dispose()
        }
        function Add-ZipBinaryEntry([string]$name, [byte[]]$bytes) {
            $entry  = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
            $stream = $entry.Open()
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Dispose()
        }
        $d    = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ssZ')
        $te   = Invoke-XmlEscape $DocTitle
        $heSrc = if ($HeaderLabel) { $HeaderLabel } else { $DocTitle }
        $he   = Invoke-XmlEscape $heSrc
        $pngCT = if ($hasLogo) { "`n  <Default Extension=`"png`" ContentType=`"image/png`"/>" } else { '' }
        Add-ZipEntry '[Content_Types].xml' @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>$pngCT
  <Override PartName="/word/document.xml"  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/header1.xml"   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml"   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/docProps/core.xml"  ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>
"@
        Add-ZipEntry '_rels/.rels' @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>
'@
        Add-ZipEntry 'docProps/core.xml' @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>$te</dc:title>
  <dc:creator>EXpress v$ScriptVersion</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">$d</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">$d</dcterms:modified>
</cp:coreProperties>
"@
        $logoRel = if ($hasLogo) { "`n  <Relationship Id=`"rId5`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`" Target=`"media/logo.png`"/>" } else { '' }
        Add-ZipEntry 'word/_rels/document.xml.rels' @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"    Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"    Target="header1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"    Target="footer1.xml"/>$logoRel
</Relationships>
"@
        Add-ZipEntry 'word/styles.xml' @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
      <w:sz w:val="22"/><w:szCs w:val="22"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:pageBreakBefore/><w:keepNext/><w:keepLines/><w:spacing w:before="480" w:after="80"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:b/><w:color w:val="2F5496"/><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:jc w:val="center"/><w:spacing w:before="240" w:after="120"/><w:contextualSpacing/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:b/><w:color w:val="1F3864"/><w:sz w:val="72"/><w:szCs w:val="72"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle">
    <w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:after="120"/><w:contextualSpacing/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:i/><w:color w:val="2E74B5"/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="TOCHeading">
    <w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/><w:next w:val="Normal"/>
    <w:pPr><w:pageBreakBefore/><w:outlineLvl w:val="9"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:b/><w:color w:val="2F5496"/><w:sz w:val="40"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="TOC1">
    <w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:spacing w:before="120" w:after="0"/></w:pPr>
    <w:rPr><w:b/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="TOC2">
    <w:name w:val="toc 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:spacing w:after="0"/><w:ind w:left="220"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="TOC3">
    <w:name w:val="toc 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:spacing w:after="0"/><w:ind w:left="440"/></w:pPr>
  </w:style>
  <w:style w:type="character" w:styleId="PlaceholderText">
    <w:name w:val="Placeholder Text"/>
    <w:rPr><w:color w:val="808080"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="40"/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:b/><w:color w:val="2E74B5"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="40"/><w:outlineLvl w:val="2"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:b/><w:color w:val="1F3864"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading4">
    <w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="20"/><w:outlineLvl w:val="3"/></w:pPr>
    <w:rPr><w:i/><w:color w:val="2E74B5"/><w:sz w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Code">
    <w:name w:val="Code"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="0" w:after="0"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Courier New"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:ind w:left="720"/></w:pPr>
  </w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal">
    <w:name w:val="Normal Table"/>
    <w:tblPr><w:tblCellMar>
      <w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
      <w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
    </w:tblCellMar></w:tblPr>
  </w:style>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/>
    <w:tblPr><w:tblBorders>
      <w:top    w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:left   w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:right  w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    </w:tblBorders></w:tblPr>
  </w:style>
</w:styles>
'@
        Add-ZipEntry 'word/numbering.xml' @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#x2022;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/><w:numFmt w:val="bullet"/>
      <w:lvlText w:val="o"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:hint="default"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>
'@
        Add-ZipEntry 'word/header1.xml' @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:jc w:val="right"/>
      <w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="2F5496"/></w:pBdr>
      <w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr>
    </w:pPr>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr>
      <w:t>$he</w:t>
    </w:r>
  </w:p>
</w:hdr>
"@
        Add-ZipEntry 'word/footer1.xml' @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pBdr><w:top w:val="single" w:sz="6" w:space="1" w:color="2F5496"/></w:pBdr>
      <w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs>
      <w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr>
    </w:pPr>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:t>INTERN</w:t></w:r>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:tab/></w:r>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve"> / </w:t></w:r>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:ftr>
'@
        if ($hasLogo) { Add-ZipBinaryEntry 'word/media/logo.png' ([System.IO.File]::ReadAllBytes($LogoPath)) }
        Add-ZipEntry 'word/document.xml' (New-WdDocumentXml $BodyParts)
        $zip.Dispose()
        $fs.Dispose()
    }
    function Test-WdTemplate {
        # Validates that a DOCX template contains all required {{token}} placeholders.
        # Searches across all XML parts (document, header, footer, etc.).
        # Returns @{ Valid=[bool]; Missing=[string[]] }.
        param([string]$Path, [string[]]$RequiredTags = @('document_body'))
        Add-Type -AssemblyName System.IO.Compression
        $missing = [System.Collections.Generic.List[string]]::new()
        try {
            $fs  = [System.IO.File]::OpenRead($Path)
            $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Read)
            $allXml = ''
            foreach ($entry in $zip.Entries) {
                if ($entry.FullName -match '\.xml$') {
                    $sr      = [System.IO.StreamReader]::new($entry.Open())
                    $allXml += $sr.ReadToEnd()
                    $sr.Dispose()
                }
            }
            $zip.Dispose()
            $fs.Dispose()
        } catch {
            return @{ Valid = $false; Missing = @(('Cannot open template: ' + $_)) }
        }
        foreach ($tag in $RequiredTags) {
            if ($allXml -notlike ('*{{' + $tag + '}}*')) { $null = $missing.Add($tag) }
        }
        return @{ Valid = ($missing.Count -eq 0); Missing = $missing.ToArray() }
    }
    function Write-WdFromTemplate {
        # Copies a DOCX template to $OutputPath and replaces all {{token}} placeholders
        # in every XML part with the corresponding values from the $Tokens hashtable.
        # Special token 'document_body': replaces the entire anchor paragraph
        #   <w:p><w:r><w:t>{{document_body}}</w:t></w:r></w:p>
        # with the supplied chapter XML string (multiple <w:p> elements).
        # All other tokens are XML-escaped before substitution.
        param([string]$TemplatePath, [string]$OutputPath, [hashtable]$Tokens)
        Add-Type -AssemblyName System.IO.Compression
        $enc = [System.Text.UTF8Encoding]::new($false)
        [System.IO.File]::Copy($TemplatePath, $OutputPath, $true)
        $fs  = [System.IO.File]::Open($OutputPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
        $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Update)
        $xmlEntries = @($zip.Entries | Where-Object { $_.FullName -match '\.xml$' })
        foreach ($entry in $xmlEntries) {
            $sr      = [System.IO.StreamReader]::new($entry.Open())
            $content = $sr.ReadToEnd()
            $sr.Dispose()
            $modified = $false
            foreach ($kv in $Tokens.GetEnumerator()) {
                $marker = '{{' + $kv.Key + '}}'
                if ($content.Contains($marker)) {
                    if ($kv.Key -eq 'document_body') {
                        $anchor = '<w:p><w:r><w:t>' + $marker + '</w:t></w:r></w:p>'
                        $content = $content.Replace($anchor, $kv.Value)
                    } else {
                        $content = $content.Replace($marker, (Invoke-XmlEscape $kv.Value))
                    }
                    $modified = $true
                }
            }
            if ($modified) {
                $entryName = $entry.FullName
                $entry.Delete()
                $newEntry = $zip.CreateEntry($entryName)
                $sw = [System.IO.StreamWriter]::new($newEntry.Open(), $enc)
                $sw.Write($content)
                $sw.Dispose()
            }
        }
        $zip.Dispose()
        $fs.Dispose()
    }

    # ── New-InstallationDocument (F22) ────────────────────────────────────────────
    function New-InstallationDocument {
        # Default to EN unless the state explicitly requests DE. Previous logic
        # ($State['Language'] -ne 'EN') flipped to DE whenever the key was missing
        # (e.g. Phase-5 re-entry against a pre-5.93 state file without the key).
        $DE   = ($State['Language'] -eq 'DE')
        $cust = [bool]$State['CustomerDocument']
        $lang = if ($DE) { 'DE' } else { 'EN' }
        $scope         = if ($State['DocumentScope']) { $State['DocumentScope'] } else { 'All' }
        $includeFilter = if ($State['IncludeServers']) { @($State['IncludeServers'] -split ',') } else { @() }
        $isAdHoc       = [bool]$State['StandaloneDocument'] -and -not $State['InstallPhase']
        $docStem  = if ($DE) { 'ExchangeServer-Dokumentation' } else { 'ExchangeServer-Documentation' }
        $docPath  = Join-Path $State['ReportsPath'] ('{0}_EXpress_{1}_{2}_{3}.docx' -f $env:COMPUTERNAME, $docStem, $lang, (Get-Date -Format 'yyyyMMdd-HHmmss'))
        $docTitle = if ($DE) { 'Exchange Server Installationsdokumentation' } else { 'Exchange Server Installation Documentation' }
        Write-MyOutput ('Generating Word Installation Document ({0}): {1}' -f $lang, $docPath)

        Write-MyVerbose 'Collecting installation report data'
        $rd = Get-InstallationReportData -Scope $scope -IncludeServers $includeFilter

        function Mask-Ip([string]$text) {
            if (-not $cust) { return $text }
            $text -replace '\b(10|172\.(1[6-9]|2[0-9]|3[01])|192\.168)\.\d{1,3}\.\d{1,3}\b', 'x.x.x.x'
        }
        function Mask-Val([string]$text) { if ($cust -and $text) { '[redacted]' } else { $text } }
        function SafeVal([object]$v, [string]$fallback = '') { if ($null -eq $v -or "$v" -eq '') { $fallback } else { "$v" } }
        # L / Lc: language helper. PS 5.1 cannot use (if ...) as a command argument; these helpers keep call sites compact.
        function L([string]$d, [string]$e) { if ($DE) { $d } else { $e } }
        function Lc([bool]$c, [string]$a, [string]$b) { if ($c) { $a } else { $b } }
        function Get-SecReg($path, $name) { try { (Get-ItemProperty -Path $path -Name $name -ErrorAction Stop).$name } catch { $null } }
        # Format-RegBool: translate registry 0/1 (or $false/$true) to localised enabled/disabled text.
        function Format-RegBool($v) {
            if ($null -eq $v -or "$v" -eq '') { return (L '(nicht gesetzt)' '(not set)') }
            # Use [bool] instead of [int]: Exchange cmdlet properties can return SwitchParameter,
            # which [int] cannot cast in PS 5.1 but [bool] handles via implicit conversion.
            if ([bool]$v) { return (L 'aktiviert' 'enabled') }
            return (L 'deaktiviert' 'disabled')
        }
        function Format-RemoteSysRows($remData) {
            $rows = [System.Collections.Generic.List[object[]]]::new()
            if (-not $remData -or -not $remData.Reachable) {
                $errMsg = if ($remData -and $remData.Error) { $remData.Error } else { (L 'WinRM nicht erreichbar' 'WinRM not reachable') }
                $rows.Add(@((L 'Systemdetails' 'System details'), (L ('Nicht abrufbar: {0} — Abhilfe: tools\Enable-EXpressRemoteQuery.ps1' -f $errMsg) ('Not available: {0} — Fix: tools\Enable-EXpressRemoteQuery.ps1' -f $errMsg))))
                return ,$rows
            }
            if ($remData.OS) {
                $rows.Add(@((L 'Betriebssystem' 'Operating system'), $remData.OS.Caption))
                $rows.Add(@((L 'OS-Build' 'OS build'), $remData.OS.Version))
                $rows.Add(@((L 'Letzter Neustart' 'Last boot'), $remData.OS.LastBootUpTime.ToString('yyyy-MM-dd HH:mm:ss')))
                $rows.Add(@((L 'RAM gesamt' 'Total RAM'), ('{0} GB' -f [math]::Round($remData.OS.TotalVisibleMemorySize / 1MB, 0))))
            }
            if ($remData.CPU) {
                $cpuList = @($remData.CPU)
                $totalCores   = ($cpuList | Measure-Object NumberOfCores -Sum).Sum
                $totalLogical = ($cpuList | Measure-Object NumberOfLogicalProcessors -Sum).Sum
                $rows.Add(@('CPU', ('{0} — {1} {2} / {3} {4}' -f $cpuList[0].Name.Trim(), $totalCores, (L 'Kerne' 'cores'), $totalLogical, (L 'logisch' 'logical'))))
            }
            if ($remData.ComputerSys) {
                $rows.Add(@((L 'Computername (FQDN)' 'Computer name (FQDN)'), ('{0}.{1}' -f $remData.ComputerSys.DNSHostName, $remData.ComputerSys.Domain)))
            }
            foreach ($vol in $remData.Volumes) {
                if ($vol.DriveLetter -and $vol.Capacity -gt 0) {
                    $freeGB = [math]::Round($vol.FreeSpace / 1GB, 1)
                    $totGB  = [math]::Round($vol.Capacity / 1GB, 1)
                    $pct    = [math]::Round($vol.FreeSpace / $vol.Capacity * 100, 0)
                    $au     = if ($vol.BlockSize) { '{0} KB' -f ($vol.BlockSize / 1KB) } else { '?' }
                    $rows.Add(@(('Volume {0}:' -f $vol.DriveLetter), ('{0} GB {1} / {2} GB ({3}% free) — AU: {4}' -f $freeGB, $vol.FileSystem, $totGB, $pct, $au)))
                }
            }
            if ($remData.PageFile) {
                $pf    = $remData.PageFile
                $ramMB = if ($remData.OS) { [math]::Round($remData.OS.TotalVisibleMemorySize / 1KB, 0) } else { 0 }
                $recMB = $ramMB + 10
                $rows.Add(@((L 'Auslagerungsdatei' 'Page file'), ('{0} — Init: {1} MB / Max: {2} MB — {3}: {4} MB' -f $pf.Name, $pf.InitialSize, $pf.MaximumSize, (L 'Empfehlung RAM+10MB' 'Recommended RAM+10MB'), $recMB)))
            }
            foreach ($nic in $remData.NICs) {
                $ips = if ($nic.IPAddress) { (Mask-Ip ($nic.IPAddress -join ', ')) } else { (L '(keine IP)' '(no IP)') }
                $dns = if ($nic.DNSServerSearchOrder) { (Mask-Ip ($nic.DNSServerSearchOrder -join ', ')) } else { (L '(nicht gesetzt)' '(not set)') }
                $rows.Add(@(('NIC: {0}' -f $nic.Description), ('{0} — DNS: {1}' -f $ips, $dns)))
            }
            return ,$rows
        }

        $parts = [System.Collections.Generic.List[string]]::new()

        # ── Template check (F24) ─────────────────────────────────────────────────
        # When -TemplatePath is supplied and valid, the cover page is driven by the
        # template DOCX; $parts contains only the chapter body XML.
        $tplPath = $State['TemplatePath']
        $useTpl  = $tplPath -and (Test-Path $tplPath -PathType Leaf)
        if ($useTpl) {
            $tplCheck = Test-WdTemplate -Path $tplPath -RequiredTags @('document_body')
            if (-not $tplCheck.Valid) {
                Write-MyWarning ('Template missing required tokens: ' + ($tplCheck.Missing -join ', ') + ' — falling back to built-in cover page.')
                $useTpl = $false
            } else {
                Write-MyVerbose ('Using custom template: ' + $tplPath)
            }
        }

        $instMode = if ($isAdHoc) { (L 'Ad-hoc-Inventar' 'Ad-hoc Inventory') } elseif ($State['InstallEdge']) { 'Edge Transport' } elseif ($State['InstallRecipientManagement']) { 'Recipient Management Tools' } elseif ($State['InstallManagementTools']) { 'Management Tools' } elseif ($State['StandaloneOptimize']) { 'Standalone Optimize' } elseif ($State['NoSetup']) { 'Optimization Only' } else { 'Mailbox Server' }
        $scenario = if ($isAdHoc) { (L 'Ad-hoc-Inventar (vorhandene Umgebung)' 'Ad-hoc inventory (existing environment)') } elseif ($rd.Servers.Count -le 1) { (L 'Neue Exchange-Umgebung' 'New Exchange environment') } else { (L 'Server-Ergänzung zu bestehender Umgebung' 'Server added to existing environment') }
        $classification = (Lc $cust 'CUSTOMER' 'INTERN')

        # Cover page variables — needed both for built-in cover page and template tokens.
        $company  = SafeVal $State['CompanyName'] ''
        $author   = SafeVal $State['Author']      ''
        $coverSub = (L 'Installation, Hybridbereitstellung, Mailflow' 'Installation, Hybrid deployment, Mail flow')
        # Logo probe: sources\logo.png (user-placed) → beside the script → assets\logo.png (repo default)
        $_logoRoot = if ($PSScriptRoot) { $PSScriptRoot } else { $State['InstallPath'] }
        $logoFile = @(
            (Join-Path $State['SourcesPath'] 'logo.png'),
            (Join-Path $_logoRoot 'logo.png'),
            (Join-Path $_logoRoot 'assets\logo.png')
        ) | Where-Object { Test-Path $_ -PathType Leaf } | Select-Object -First 1
        if (-not $logoFile) { $logoFile = Join-Path $State['SourcesPath'] 'logo.png' }   # fallback path (will fail Test-Path gracefully)

        if (-not $useTpl) {
        # ── Deckblatt (Cover Page) ───────────────────────────────────────────────
        # Layout nach Referenzvorlage: Produkt (groß) / Titel (XXL) / Untertitel / Version+Datum+Autor.
        # Company/Author sind State-gesteuert ($State['CompanyName'], $State['Author']) ohne Default-Branding.
        $null = $parts.Add((New-WdSpacer 1440))
        if (Test-Path $logoFile -PathType Leaf) {
            # Logo centered, 6 cm wide (2160000 EMU), proportional height for 400×80 source: 432000 EMU
            $null = $parts.Add('<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="240"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="2160000" cy="432000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="logo"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="logo"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2160000" cy="432000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')
        }
        $null = $parts.Add((New-WdCentered -Text 'Microsoft Exchange Server SE' -SizeHalfPt 40 -Bold $true -Color '1F3864'))
        $null = $parts.Add(('<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t xml:space="preserve">{0}</w:t></w:r></w:p>' -f (Invoke-XmlEscape (L 'Installation & Konfiguration' 'Installation & Configuration'))))
        $null = $parts.Add(('<w:p><w:pPr><w:pStyle w:val="Subtitle"/></w:pPr><w:r><w:t xml:space="preserve">{0}</w:t></w:r></w:p>' -f (Invoke-XmlEscape $coverSub)))
        $null = $parts.Add((New-WdSpacer 1200))
        $orgLine = SafeVal $State['OrganizationName'] ''
        if ($orgLine) { $null = $parts.Add((New-WdCentered -Text $orgLine -SizeHalfPt 28 -Bold $true -Color '1F3864')) }
        $null = $parts.Add((New-WdCentered -Text $env:COMPUTERNAME -SizeHalfPt 24 -Color '404040'))
        $null = $parts.Add((New-WdCentered -Text $scenario          -SizeHalfPt 22 -Color '404040'))
        $null = $parts.Add((New-WdCentered -Text (('{0}: {1}' -f (L 'Installationsmodus' 'Installation mode'), $instMode)) -SizeHalfPt 22 -Color '404040'))
        $null = $parts.Add((New-WdSpacer 1440))
        $null = $parts.Add((New-WdCentered -Text (('{0}: {1}' -f (L 'Versionsnummer' 'Version'), ('{0} / EXpress v{1}' -f (Get-Date -Format 'yyyy-MM-dd'), $ScriptVersion))) -SizeHalfPt 22 -Color '404040'))
        $null = $parts.Add((New-WdCentered -Text (('{0}: {1}' -f (L 'Datum' 'Date'), (Get-Date -Format 'dd.MM.yyyy')))                       -SizeHalfPt 22 -Color '404040'))
        if ($author)  { $null = $parts.Add((New-WdCentered -Text (('{0}: {1}' -f (L 'Autor' 'Author'), $author))                              -SizeHalfPt 22 -Color '404040')) }
        if ($company) { $null = $parts.Add((New-WdCentered -Text $company                                                                      -SizeHalfPt 22 -Color '404040')) }
        $null = $parts.Add((New-WdSpacer 600))
        $null = $parts.Add((New-WdCentered -Text $classification -SizeHalfPt 22 -Bold $true -Color 'C00000'))
        $null = $parts.Add((New-WdPageBreak))
        } # end if (-not $useTpl)

        # ── Hinweise zu diesem Dokument ─────────────────────────────────────────
        # Struktur nach Referenzvorlage: Anpassungsvorbehalt, Genderhinweis, Warenzeichen,
        # Screenshots/Mockups, Copyright. Firmenname aus $State['CompanyName'] — kein Default.
        $companyRef = if ($company) { $company } else { (L 'der Hersteller dieses Dokuments' 'the publisher of this document') }
        $null = $parts.Add((New-WdHeading (L 'Hinweise zu diesem Dokument' 'Notes on this document') 1))
        $null = $parts.Add((New-WdParagraph (L ('{0} behält sich vor, den beschriebenen Funktionsumfang jederzeit an neue Anforderungen und Erkenntnisse anzupassen. Dadurch kann es gegebenenfalls zu Abweichungen zwischen diesem Dokument und der ausgelieferten Software kommen.' -f $companyRef) ('{0} reserves the right to adapt the functional scope described herein to new requirements and insights at any time. This may result in deviations between this document and the delivered software.' -f $companyRef))))
        $null = $parts.Add((New-WdParagraph (L 'Genderhinweis: Aus Gründen der besseren Lesbarkeit wird auf eine geschlechtsneutrale Differenzierung verzichtet. Entsprechende Begriffe gelten im Sinne der Gleichbehandlung grundsätzlich für alle Geschlechter. Die verkürzte Sprachform beinhaltet keine Wertung.' 'Gender note: For better readability, gender-neutral differentiation is omitted. Corresponding terms apply to all genders in the sense of equal treatment. The abbreviated language form does not imply a value judgement.')))
        $null = $parts.Add((New-WdParagraph (L 'Die hier genannten Produkte und Namen sind eingetragene Warenzeichen und/oder geschützte Marken und damit Eigentum der jeweiligen Rechteinhaber, u. a. der Microsoft Corporation (Microsoft, Exchange Server, Windows Server, Active Directory, Microsoft 365, Intune), Intel Corporation und weiterer.' 'The products and names mentioned here are registered trademarks and/or protected brands and therefore the property of the respective rights holders, including Microsoft Corporation (Microsoft, Exchange Server, Windows Server, Active Directory, Microsoft 365, Intune), Intel Corporation and others.')))
        $null = $parts.Add((New-WdParagraph (L 'Bitte beachten Sie: Teilweise zeigen dargestellte Ausgaben und Tabellen eine beispielhafte Konfiguration, um die beschriebenen Prozesse und Funktionalitäten zu dokumentieren. In Abstimmung mit dem Auftraggeber werden in der Vorbereitungsphase offene Fragen für die konkrete Umsetzung besprochen.' 'Please note: Some of the outputs and tables shown depict an exemplary configuration to document the described processes and functionality. Open questions regarding the concrete implementation are discussed with the contracting party during the preparation phase.')))
        $copyrightHolder = if ($company) { $company } else { (L '(Hersteller)' '(publisher)') }
        $null = $parts.Add((New-WdParagraph (L ('© Copyright {0}. Alle Rechte vorbehalten. Die Weitergabe und Vervielfältigung dieser Publikation oder von Teilen daraus sind, zu welchem Zweck und in welcher Form auch immer, ohne ausdrückliche schriftliche Genehmigung nicht gestattet. In dieser Publikation enthaltene Informationen können ohne vorherige Ankündigung geändert werden.' -f $copyrightHolder) ('© Copyright {0}. All rights reserved. Reproduction or distribution of this publication or parts thereof, for any purpose and in any form, is not permitted without express written approval. Information contained in this publication may be changed without prior notice.' -f $copyrightHolder))))
        $null = $parts.Add((New-WdParagraph (L 'Dieses Dokument wurde automatisch durch EXpress (Install-Exchange15.ps1) generiert und spiegelt die Konfiguration der Exchange-Umgebung zum Erstellungszeitpunkt wider. Spätere Änderungen sind nicht berücksichtigt. EXpress wird "wie besehen" ohne Gewährleistung bereitgestellt; die Verantwortung für die Einhaltung organisatorischer, rechtlicher sowie regulatorischer Vorgaben (z. B. DSGVO, GoBD, BAIT/VAIT, ISO 27001) liegt beim Betreiber.' 'This document was generated automatically by EXpress (Install-Exchange15.ps1) and reflects the Exchange environment configuration at the time of generation. Subsequent changes are not reflected. EXpress is provided "as is" without warranty; responsibility for compliance with organisational, legal and regulatory requirements (e.g. GDPR, SOX, ISO 27001) lies with the operator.')))
        $null = $parts.Add((New-WdHeading (L 'Versionshistorie' 'Revision History') 2))
        $revAuthor = if ($author) { $author } else { ('EXpress v{0}' -f $ScriptVersion) }
        $null = $parts.Add((New-WdTable -Headers @((L 'Version' 'Version'), (L 'Datum' 'Date'), (L 'Autor' 'Author'), (L 'Änderung' 'Change')) -Rows @(
            @('1.0', (Get-Date -Format 'dd.MM.yyyy'), $revAuthor, (L 'Automatische Erstgenerierung' 'Automatic initial generation'))
        )))

        # ── Dynamisches Inhaltsverzeichnis ───────────────────────────────────────
        $null = $parts.Add((New-WdToc (L 'Inhaltsverzeichnis' 'Table of Contents')))

        # ── 1. Dokumenteigenschaften ─────────────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '1. Dokumenteigenschaften' '1. Document Properties') 1))
        $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows @(
            @((L 'Dokument' 'Document'), $docTitle)
            @('EXpress Version', "v$ScriptVersion")
            @((L 'Erstellt auf Server' 'Generated on server'), $env:COMPUTERNAME)
            @((L 'Exchange-Organisation' 'Exchange Organisation'), (SafeVal $State['OrganizationName'] (L '(nicht gesetzt)' '(not set)')))
            @((L 'Szenario' 'Scenario'), $scenario)
            @((L 'Installationsmodus' 'Installation mode'), $instMode)
            @((L 'Installiert durch' 'Installed by'), (SafeVal $State['InstallingUser'] (L '(unbekannt)' '(unknown)')))
            @((L 'Erstellt am' 'Generated on'), (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
            @((L 'Klassifizierung' 'Classification'), $classification)
        )))

        # ── 1.1 Freigabe und Change-Management ───────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '1.1 Freigabe und Change-Management' '1.1 Sign-off and Change Management') 2))
        $null = $parts.Add((New-WdParagraph (L 'Die folgende Tabelle dient als formaler Freigabenachweis dieser Installation. Bitte nach Abschluss der Installation und Durchführung der Abnahmetests ausfüllen (siehe auch Kapitel 16 Abnahmetest).' 'The table below serves as formal sign-off evidence for this installation. Please complete after finishing the installation and acceptance tests (see also chapter 16 Acceptance Testing).')))
        $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows @(
            ,@((L 'Change-Request-Nr.' 'Change request no.'), '')
            ,@((L 'Genehmigt von' 'Approved by'), '')
            ,@((L 'Genehmigungsdatum' 'Approval date'), '')
            ,@((L 'Abnahme durch' 'Accepted by'), '')
            ,@((L 'Abnahmedatum' 'Acceptance date'), '')
            ,@((L 'Bemerkungen' 'Remarks'), '')
        )))

        # ── 2. Installationsparameter (nur bei tatsächlichem Setup-Lauf) ─────────
        if (-not $isAdHoc) {
            $null = $parts.Add((New-WdHeading (L '2. Installationsparameter' '2. Installation Parameters') 1))
            $null = $parts.Add((New-WdParagraph (L 'Die folgende Tabelle dokumentiert die bei der Installation verwendeten Parameter. Sie dient als Nachweis der gewählten Konfiguration und als Referenz für spätere Änderungen oder eine Neuinstallation. Im Autopilot-Modus wurden alle Parameter aus einer Konfigurationsdatei geladen; im Copilot-Modus wurden sie während des Installationslaufs interaktiv abgefragt.' 'The table below documents the parameters used during installation. It serves as evidence of the chosen configuration and as a reference for later changes or a reinstallation. In Autopilot mode, all parameters were loaded from a configuration file; in Copilot mode they were interactively collected during the installation run.')))
            $modeText = if ($State['ConfigDriven']) { (L 'Autopilot (vollautomatisch)' 'Autopilot (fully automated)') } else { (L 'Copilot (interaktiv)' 'Copilot (interactive)') }
            $paramRows = [System.Collections.Generic.List[object[]]]::new()
            $paramRows.Add(@((L 'Setup-Version' 'Setup version'), (SafeVal (& { try { (Get-SetupTextVersion $State['SetupVersion']) } catch { $State['SetupVersion'] } }))))
            $paramRows.Add(@((L 'Installationspfad' 'Install path'), (SafeVal $State['InstallPath'])))
            if ($State['Namespace'])        { $paramRows.Add(@('Namespace', (SafeVal $State['Namespace']))) }
            if ($State['DownloadDomain'])   { $paramRows.Add(@('OWA Download Domain', (SafeVal $State['DownloadDomain']))) }
            if ($State['DAGName'])          { $paramRows.Add(@('DAG', (SafeVal $State['DAGName']))) }
            if ($State['CertificatePath'])  { $paramRows.Add(@((L 'Zertifikatspfad' 'Certificate path'), (Mask-Val (SafeVal $State['CertificatePath'])))) }
            if ($State['LogRetentionDays']) { $paramRows.Add(@((L 'Log-Aufbewahrung' 'Log retention'), ('{0} {1}' -f $State['LogRetentionDays'], (L 'Tage' 'days')))) }
            if ($State['RelaySubnets'])     { $paramRows.Add(@((L 'Relay-Subnetze' 'Relay subnets'), (Mask-Ip (($State['RelaySubnets'] -join ', '))))) }
            $paramRows.Add(@((L 'Modus' 'Mode'), $modeText))
            $paramRows.Add(@('TLS 1.2', (Format-RegBool $State['EnableTLS12'])))
            $paramRows.Add(@('TLS 1.3', (Format-RegBool $State['EnableTLS13'])))
            # PS 5.1: (if ...) cannot be used inline as an array element — assign first (Known Pitfall)
            $tls10text = if ($null -eq $State['DisableSSL3']) { (L '(nicht gesetzt)' '(not set)') } elseif ($State['DisableSSL3']) { (L 'deaktiviert' 'disabled') } else { (L 'aktiv' 'active') }
            $paramRows.Add(@('TLS 1.0 / TLS 1.1', $tls10text))
            $paramRows.Add(@((L 'Logdatei' 'Log file'), (SafeVal $State['TranscriptFile'])))
            $null = $parts.Add((New-WdTable -Headers @((L 'Parameter' 'Parameter'), (L 'Wert' 'Value')) -Rows $paramRows.ToArray()))
        }

        # ── 3. IST-Aufnahme Active Directory ─────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '3. Active Directory — Voraussetzungen und Status' '3. Active Directory — Prerequisites and Status') 1))
        $null = $parts.Add((New-WdParagraph (L 'Exchange Server SE ist vollständig von Active Directory abhängig: Alle Konfigurationsdaten werden im AD gespeichert, die Authentifizierung erfolgt über Kerberos/NTLM gegen AD-Domänencontroller, und der Transport-Dienst nutzt AD-Standortinformationen für die Nachrichtenweiterleitung. Im Rahmen der Preflight-Prüfung wurden die AD-Voraussetzungen verifiziert. Die folgende Tabelle zeigt den ermittelten AD-Status zum Zeitpunkt der Installation.' 'Exchange Server SE is fully dependent on Active Directory: all configuration data is stored in AD, authentication is handled via Kerberos/NTLM against AD domain controllers, and the transport service uses AD site information for message routing. The AD prerequisites were verified during the preflight check. The table below shows the AD status at the time of installation.')))
        $adRows = [System.Collections.Generic.List[object[]]]::new()
        try { $localCS = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue; if ($localCS) { $adRows.Add(@((L 'Domäne' 'Domain'), $localCS.Domain)) } } catch { }
        try { $ffl = Get-ForestFunctionalLevel; $adRows.Add(@((L 'Forest Functional Level' 'Forest functional level'), ('{0} ({1})' -f $ffl, (Get-FFLText $ffl)))) } catch { }
        try {
            $exOrg = Get-ExchangeOrganization
            if ($exOrg) { $adRows.Add(@((L 'Exchange-Organisation' 'Exchange organisation'), $exOrg)) }
            $adRows.Add(@((L 'Exchange Forest Schema (rangeUpper)' 'Exchange forest schema (rangeUpper)'), (SafeVal (Get-ExchangeForestLevel))))
            $adRows.Add(@((L 'Exchange Domain Level' 'Exchange domain level'), (SafeVal (Get-ExchangeDomainLevel))))
        } catch { }
        try {
            $fsmoRoles = @{}
            $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            $fsmoRoles[(L 'Schema Master' 'Schema Master')] = $forest.SchemaRoleOwner.Name
            $fsmoRoles[(L 'Domain Naming Master' 'Domain Naming Master')] = $forest.NamingRoleOwner.Name
            $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $fsmoRoles[(L 'PDC Emulator' 'PDC Emulator')] = $domain.PdcRoleOwner.Name
            $fsmoRoles[(L 'RID Master' 'RID Master')] = $domain.RidRoleOwner.Name
            $fsmoRoles[(L 'Infrastructure Master' 'Infrastructure Master')] = $domain.InfrastructureRoleOwner.Name
            foreach ($role in $fsmoRoles.Keys) { $adRows.Add(@($role, (Mask-Ip $fsmoRoles[$role]))) }
        } catch { }
        if ($adRows.Count -eq 0) { $adRows.Add(@((L '(AD-Daten nicht abrufbar)' '(AD data not available)'), '')) }
        $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $adRows.ToArray()))
        $null = $parts.Add((New-WdParagraph (L 'Hinweis: Ein Forest Functional Level von mindestens Windows Server 2012 R2 (Level 6) ist für Exchange SE erforderlich. Schema- und Domänenerweiterungen (PrepareSchema / PrepareAD / PrepareDomain) wurden von EXpress automatisch durchgeführt.' 'Note: A Forest Functional Level of at least Windows Server 2012 R2 (Level 6) is required for Exchange SE. Schema and domain extensions (PrepareSchema / PrepareAD / PrepareDomain) were performed automatically by EXpress.')))

        # ── 4. Organisation — übergreifende Konfiguration ────────────────────────
        if ($scope -in 'All','Org') {
            $null = $parts.Add((New-WdHeading (L '4. Organisation — übergreifende Konfiguration' '4. Organisation — Global Configuration') 1))
            $null = $parts.Add((New-WdParagraph (L 'Die Exchange-Organisation umfasst alle Exchange-Server in der AD-Gesamtstruktur. Die folgenden Abschnitte dokumentieren die organisationsweiten Einstellungen, die auf alle Server und Postfächer in der Organisation wirken.' 'The Exchange organisation encompasses all Exchange servers in the AD forest. The following sections document the organisation-wide settings that apply to all servers and mailboxes in the organisation.')))
            $orgD = $rd.Org

            # 4.1 Org-Übersicht
            $null = $parts.Add((New-WdHeading (L '4.1 Org-Übersicht' '4.1 Organisation Overview') 2))
            $orgRows = [System.Collections.Generic.List[object[]]]::new()
            if ($orgD -and $orgD.OrgConfig) {
                $oc = $orgD.OrgConfig
                $orgRows.Add(@((L 'Name' 'Name'), (SafeVal $oc.Name)))
                $orgRows.Add(@((L 'Version' 'Version'), (SafeVal $oc.AdminDisplayVersion)))
                $orgRows.Add(@((L 'MAPI/HTTP' 'MAPI/HTTP'), (SafeVal $oc.MapiHttpEnabled)))
                $orgRows.Add(@((L 'Modern Auth (OAuth2)' 'Modern Auth (OAuth2)'), (SafeVal $oc.OAuth2ClientProfileEnabled)))
                $orgRows.Add(@((L 'CEIP deaktiviert' 'CEIP disabled'), (SafeVal (-not $oc.CustomerFeedbackEnabled))))
                if ($null -ne $oc.DefaultPublicFolderMailbox) { $orgRows.Add(@((L 'Standard-PF-Postfach' 'Default PF mailbox'), (SafeVal $oc.DefaultPublicFolderMailbox))) }
            }
            $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $orgRows.ToArray()))

            # 4.2 Accepted Domains
            $null = $parts.Add((New-WdHeading (L '4.2 Accepted Domains' '4.2 Accepted Domains') 2))
            $adDomRows = [System.Collections.Generic.List[object[]]]::new()
            foreach ($dom in $orgD.AcceptedDomains) { $adDomRows.Add(@($dom.DomainName, $dom.DomainType, (Lc $dom.Default (L 'Standard' 'Default') ''))) }
            $null = $parts.Add((New-WdTable -Headers @((L 'Domäne' 'Domain'), (L 'Typ' 'Type'), (L 'Standard' 'Default')) -Rows $adDomRows.ToArray()))

            # 4.3 Remote Domains
            $null = $parts.Add((New-WdHeading (L '4.3 Remote Domains' '4.3 Remote Domains') 2))
            $rdRows = [System.Collections.Generic.List[object[]]]::new()
            foreach ($rd2 in $orgD.RemoteDomains) { $rdRows.Add(@($rd2.DomainName, (SafeVal $rd2.ContentType), (Lc $rd2.AutoReplyEnabled (L 'Auto-Reply aktiv' 'Auto-reply active') ''))) }
            $null = $parts.Add((New-WdTable -Headers @((L 'Domäne' 'Domain'), (L 'Content-Typ' 'Content type'), (L 'Hinweis' 'Note')) -Rows $rdRows.ToArray()))

            # 4.4 E-Mail-Adressrichtlinien
            $null = $parts.Add((New-WdHeading (L '4.4 E-Mail-Adressrichtlinien' '4.4 Email Address Policies') 2))
            $eapRows = [System.Collections.Generic.List[object[]]]::new()
            foreach ($pol in $orgD.EmailAddressPolicies) { $eapRows.Add(@($pol.Name, (SafeVal $pol.RecipientFilter), (SafeVal ($pol.EnabledEmailAddressTemplates -join ', ')))) }
            $null = $parts.Add((New-WdTable -Headers @((L 'Name' 'Name'), (L 'Empfängerfilter' 'Recipient filter'), (L 'Adressvorlagen' 'Address templates')) -Rows $eapRows.ToArray()))

            # 4.5 Transport Rules
            $null = $parts.Add((New-WdHeading (L '4.5 Transportregeln' '4.5 Transport Rules') 2))
            $trRows = [System.Collections.Generic.List[object[]]]::new()
            foreach ($tr in $orgD.TransportRules) { $trRows.Add(@($tr.Name, $tr.State, $tr.Priority, (SafeVal $tr.Comments))) }
            if ($trRows.Count -eq 0) { $trRows.Add(@((L '(keine Regeln konfiguriert)' '(no rules configured)'), '', '', '')) }
            $null = $parts.Add((New-WdTable -Headers @((L 'Name' 'Name'), (L 'Status' 'State'), (L 'Priorität' 'Priority'), (L 'Kommentar' 'Comment')) -Rows $trRows.ToArray()))

            # 4.6 Transport-Konfiguration (Org)
            $null = $parts.Add((New-WdHeading (L '4.6 Transport-Konfiguration' '4.6 Transport Configuration') 2))
            $tcRows = [System.Collections.Generic.List[object[]]]::new()
            if ($orgD.TransportConfig) {
                $tc2 = $orgD.TransportConfig
                # MaxSendSize / MaxReceiveSize may be Unlimited ($null .Value) on a fresh org
                # or when the Exchange snap-in is unavailable. Format explicitly with a null-guard.
                $fmtSize = {
                    param($sz)
                    if ($null -eq $sz) { return (L 'nicht gesetzt' 'not set') }
                    if ($null -eq $sz.Value) { return (L 'unbegrenzt' 'Unlimited') }
                    '{0} MB' -f [math]::Round($sz.Value.ToBytes() / 1MB, 0)
                }
                $tcRows.Add(@((L 'Max. Sendegröße' 'Max send size'),    (& $fmtSize $tc2.MaxSendSize)))
                $tcRows.Add(@((L 'Max. Empfangsgröße' 'Max receive size'), (& $fmtSize $tc2.MaxReceiveSize)))
                $tcRows.Add(@('Safety Net Hold Time', (SafeVal $tc2.SafetyNetHoldTime)))
                $tcRows.Add(@((L 'HTML-NDRs (intern/extern)' 'HTML NDRs (internal/external)'), ('{0} / {1}' -f $tc2.InternalDsnSendHtml, $tc2.ExternalDsnSendHtml)))
            }
            $null = $parts.Add((New-WdTable -Headers @((L 'Einstellung' 'Setting'), (L 'Wert' 'Value')) -Rows $tcRows.ToArray()))

            # 4.7 Journal / DLP / Retention
            $null = $parts.Add((New-WdHeading (L '4.7 Journal-, DLP- und Aufbewahrungsrichtlinien' '4.7 Journal, DLP and Retention Policies') 2))
            $null = $parts.Add((New-WdParagraph (L 'Journaling erfasst eine Kopie aller oder ausgewählter E-Mails an eine Compliance-Postfachadresse — häufig gesetzlich vorgeschrieben (GoBD, MiFID II, SOX). Aufbewahrungsrichtlinien (Retention Policies) steuern die automatische Verschiebung oder Löschung von E-Mails nach definierten Zeiträumen (Messaging Records Management, MRM). DLP-Richtlinien (Data Loss Prevention) erkennen sensible Inhalte (z. B. Kreditkartennummern, Ausweisdaten) in E-Mails und können diese blockieren, umleiten oder markieren. In rein on-premises-Umgebungen ohne Exchange Online ist DLP nur mit eigenem Regelwerk verfügbar; die vordefinierten Microsoft 365-Vorlagen sind auf EXO beschränkt.' 'Journaling captures a copy of all or selected emails to a compliance mailbox address — often legally required (GoBD, MiFID II, SOX). Retention policies control automatic moving or deletion of emails after defined periods (Messaging Records Management, MRM). DLP policies (Data Loss Prevention) detect sensitive content (e.g. credit card numbers, ID data) in emails and can block, redirect or tag them. In purely on-premises environments without Exchange Online, DLP is only available with a custom rule set; the predefined Microsoft 365 templates are restricted to EXO.')))
            if ($orgD.JournalRules.Count -gt 0) {
                $jRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($jr in $orgD.JournalRules) { $jRows.Add(@($jr.Name, (SafeVal $jr.JournalEmailAddress), $jr.Scope, (Lc $jr.Enabled (L 'Aktiv' 'Enabled') (L 'Inaktiv' 'Disabled')))) }
                $null = $parts.Add((New-WdTable -Headers @((L 'Journal-Regel' 'Journal rule'), (L 'Empfänger' 'Recipient'), 'Scope', (L 'Status' 'Status')) -Rows $jRows.ToArray()))
            }
            if ($orgD.RetentionPolicies.Count -gt 0) {
                $rpRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($rp in $orgD.RetentionPolicies) { $rpRows.Add(@($rp.Name, (SafeVal ($rp.RetentionPolicyTagLinks -join ', ')))) }
                $null = $parts.Add((New-WdTable -Headers @((L 'Aufbewahrungsrichtlinie' 'Retention policy'), (L 'Verknüpfte Tags' 'Linked tags')) -Rows $rpRows.ToArray()))
            }
            if ($orgD.RetentionPolicyTags -and $orgD.RetentionPolicyTags.Count -gt 0) {
                $null = $parts.Add((New-WdParagraph (L 'Konfigurierte Aufbewahrungs-Tags (Retention Tags) — definieren je Postfachordner oder benutzergewählt, nach welcher Frist welche Aktion (Verschieben ins Archiv, Löschen mit/ohne Wiederherstellung, MarkAsPastRetentionLimit) ausgeführt wird:' 'Configured retention tags — define per mailbox folder or user-selectable which action (move to archive, delete with/without recovery, MarkAsPastRetentionLimit) is executed after which retention period:')))
                $rtRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($rt in ($orgD.RetentionPolicyTags | Sort-Object Type, Name)) {
                    $age = if ($null -ne $rt.AgeLimitForRetention) { ('{0} {1}' -f $rt.AgeLimitForRetention.Days, (L 'Tage' 'days')) } else { (L '(unbegrenzt)' '(unlimited)') }
                    $rtRows.Add(@(
                        $rt.Name,
                        (SafeVal $rt.Type),
                        $age,
                        (SafeVal $rt.RetentionAction),
                        (Lc $rt.RetentionEnabled (L 'Aktiv' 'Enabled') (L 'Inaktiv' 'Disabled'))
                    ))
                }
                $null = $parts.Add((New-WdTable -Headers @((L 'Tag-Name' 'Tag name'), (L 'Typ' 'Type'), (L 'Aufbewahrung' 'Retention'), (L 'Aktion' 'Action'), (L 'Status' 'Status')) -Rows $rtRows.ToArray() -Compact))
            }
            if ($orgD.DlpPolicies.Count -gt 0) {
                $dlpRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($dp in $orgD.DlpPolicies) { $dlpRows.Add(@($dp.Name, $dp.Mode, (Lc $dp.Activated (L 'Aktiv' 'Active') (L 'Inaktiv' 'Inactive')))) }
                $null = $parts.Add((New-WdTable -Headers @('DLP', 'Mode', (L 'Status' 'Status')) -Rows $dlpRows.ToArray()))
            }
            if ($orgD.JournalRules.Count -eq 0 -and $orgD.RetentionPolicies.Count -eq 0 -and $orgD.DlpPolicies.Count -eq 0) {
                $null = $parts.Add((New-WdParagraph (L '(Keine Journal-, DLP- oder Aufbewahrungsregeln konfiguriert)' '(No journal, DLP or retention policies configured)')))
            }

            # 4.8 Mobile / OWA Policies
            $null = $parts.Add((New-WdHeading (L '4.8 Mobile- und OWA-Richtlinien' '4.8 Mobile and OWA Policies') 2))
            $null = $parts.Add((New-WdParagraph (L 'Mobile Device Mailbox Policies steuern, welche Anforderungen mobile Geräte (ActiveSync, Exchange Active Sync/EAS) für die Verbindung mit Exchange erfüllen müssen: PIN-Schutz, Geräteverschlüsselung, Passwort-Komplexität, Fernlöschung (Remote Wipe). In Hybrid-Umgebungen übernehmen Intune-MDM-Richtlinien zunehmend diese Funktion; Exchange ActiveSync bleibt für on-premises-verwaltete Geräte relevant. OWA-Richtlinien kontrollieren den Funktionsumfang in Outlook Web App: Dateianhänge, S/MIME, OneNote-Integration, Skype for Business, SharePoint-Zugriff. In Hybrid-Szenarien ist die OWA-Policy-Zuweisung zwischen on-premises und EXO-Postfächern zu synchronisieren.' 'Mobile Device Mailbox Policies control which requirements mobile devices (ActiveSync, Exchange Active Sync/EAS) must meet to connect to Exchange: PIN protection, device encryption, password complexity, remote wipe. In hybrid environments, Intune MDM policies are increasingly taking over this function; Exchange ActiveSync remains relevant for on-premises-managed devices. OWA policies control the feature scope in Outlook Web App: file attachments, S/MIME, OneNote integration, Skype for Business, SharePoint access. In hybrid scenarios, OWA policy assignment between on-premises and EXO mailboxes needs to be synchronised.')))
            if ($orgD.MobileDevicePolicies.Count -gt 0) {
                $mobRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($mp in $orgD.MobileDevicePolicies) { $mobRows.Add(@($mp.Name, (Lc $mp.IsDefault (L 'Standard' 'Default') ''), (SafeVal $mp.DevicePasswordEnabled), (SafeVal $mp.DeviceEncryptionEnabled))) }
                $null = $parts.Add((New-WdTable -Headers @((L 'Richtlinie' 'Policy'), (L 'Standard' 'Default'), (L 'PIN erforderlich' 'PIN required'), (L 'Verschlüsselung' 'Encryption')) -Rows $mobRows.ToArray()))
            }
            if ($orgD.OwaPolicies.Count -gt 0) {
                $owaPolRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($op in $orgD.OwaPolicies) { $owaPolRows.Add(@($op.Name, (Lc $op.IsDefault (L 'Standard' 'Default') ''), (SafeVal $op.LogonFormat))) }
                $null = $parts.Add((New-WdTable -Headers @((L 'OWA-Richtlinie' 'OWA policy'), (L 'Standard' 'Default'), (L 'Anmeldung' 'Logon format')) -Rows $owaPolRows.ToArray()))
            }

            # 4.9 DAGs (alle)
            $null = $parts.Add((New-WdHeading (L '4.9 Database Availability Groups' '4.9 Database Availability Groups') 2))
            if ($orgD.DAGs -and $orgD.DAGs.Count -gt 0) {
                foreach ($dagEntry in $orgD.DAGs) {
                    $dag2 = $dagEntry.DAG
                    $null = $parts.Add((New-WdHeading $dag2.Name 3))
                    $dagInfoRows = @(
                        @((L 'Mitglieder' 'Members'), ($dag2.Servers -join ', '))
                        @('FSW', (Mask-Ip (SafeVal $dag2.WitnessServer)))
                        @('Alternate FSW', (Mask-Ip (SafeVal $dag2.AlternateWitnessServer)))
                        @('DAC Mode', (SafeVal $dag2.DatacenterActivationMode))
                        @((L 'Replikationsnetz' 'Replication networks'), (SafeVal ($dag2.ReplicationDagNetwork -join ', ')))
                    )
                    $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $dagInfoRows))
                    $copyRows = [System.Collections.Generic.List[object[]]]::new()
                    try {
                        Get-MailboxDatabaseCopyStatus -Server ($dag2.Servers | Select-Object -First 1) -ErrorAction SilentlyContinue | ForEach-Object {
                            $copyRows.Add(@($_.Name, $_.Status, $_.CopyQueueLength, $_.ReplayQueueLength, (SafeVal $_.ContentIndexState)))
                        }
                    } catch { }
                    if ($copyRows.Count -gt 0) {
                        $null = $parts.Add((New-WdTable -Headers @((L 'DB-Kopie' 'DB copy'), (L 'Status' 'Status'), 'Copy-Q', 'Replay-Q', (L 'Suchindex' 'Content index')) -Rows $copyRows.ToArray()))
                    }
                }
            } else {
                $null = $parts.Add((New-WdParagraph (L '(Keine DAG konfiguriert — Standalone-Umgebung)' '(No DAG configured — standalone environment)')))
            }

            # 4.10 Send Connectors
            $null = $parts.Add((New-WdHeading (L '4.10 Send Connectors' '4.10 Send Connectors') 2))
            $scRows = [System.Collections.Generic.List[object[]]]::new()
            foreach ($sc in $orgD.SendConnectors) {
                $enabledSc  = if ($sc.Enabled) { (L 'aktiviert' 'enabled') } else { (L 'deaktiviert' 'disabled') }
                $reqTlsSc   = Lc ([bool]$sc.RequireTLS) (L 'ja' 'yes') (L 'nein' 'no')
                $maxMsgSc   = if ($sc.MaxMessageSize) { $sc.MaxMessageSize.ToString() } else { '—' }
                $scRows.Add(@($sc.Name, ($sc.AddressSpaces -join ', '), (Mask-Ip (SafeVal ($sc.SmartHosts -join ', '))), (Mask-Ip ($sc.SourceTransportServers -join ', ')), (SafeVal $sc.Fqdn '—'), $reqTlsSc, $maxMsgSc, $enabledSc))
            }
            if ($scRows.Count -eq 0) { $scRows.Add(@((L '(keine konfiguriert)' '(none configured)'), '', '', '', '', '', '', '')) }
            $null = $parts.Add((New-WdTable -Headers @((L 'Name' 'Name'), (L 'Adressraum' 'Address space'), 'Smarthost', (L 'Quell-Server' 'Source servers'), 'FQDN', 'TLS', (L 'Max. Größe' 'Max size'), (L 'Status' 'Status')) -Rows $scRows.ToArray()))

            # 4.11 Federation / Hybrid / OAuth
            $null = $parts.Add((New-WdHeading (L '4.11 Federation, Hybrid und OAuth' '4.11 Federation, Hybrid and OAuth') 2))
            $null = $parts.Add((New-WdParagraph (L 'Federation und Hybrid-Konfiguration verbinden die on-premises Exchange-Organisation mit Exchange Online (Microsoft 365) bzw. anderen Exchange-Organisationen. Eine Hybrid-Konfiguration ist Voraussetzung für eine schrittweise Migration in die Cloud, für Cross-Premises-Postfachbewegungen (New-MoveRequest), für geteilte Kalenderfreigaben (Free/Busy), Nachrichtenverfolgung und für die gemeinsame Nutzung der gleichen SMTP-Domäne zwischen on-premises und Cloud. OAuth ermöglicht serverseitige Authentifizierung zwischen Exchange Server und anderen Workloads (EXO, SharePoint, Skype for Business).' 'Federation and hybrid configuration connect the on-premises Exchange organisation with Exchange Online (Microsoft 365) or other Exchange organisations. A hybrid configuration is a prerequisite for a staged cloud migration, for cross-premises mailbox moves (New-MoveRequest), for shared calendar/free-busy, message tracing, and for sharing a single SMTP namespace between on-premises and the cloud. OAuth enables server-to-server authentication between Exchange Server and other workloads (EXO, SharePoint, Skype for Business).')))
            if ($orgD.FederationTrust -and $orgD.FederationTrust.Count -gt 0) {
                $fedRows = $orgD.FederationTrust | ForEach-Object { @($_.Name, (SafeVal $_.ApplicationUri), (SafeVal $_.TokenIssuerUri)) }
                $null = $parts.Add((New-WdTable -Headers @((L 'Federation Trust' 'Federation trust'), 'Application URI', 'Token Issuer') -Rows $fedRows))
            }
            if ($orgD.HybridConfig) {
                $hyb2 = $orgD.HybridConfig
                $hybRows2 = @(
                    @((L 'Hybrid-Features' 'Hybrid features'), (SafeVal ($hyb2.Features -join ', ')))
                    @((L 'On-Premises SMTP-Domänen' 'On-premises SMTP domains'), (SafeVal ($hyb2.OnPremisesSMTPDomains -join ', ')))
                    @((L 'Edge-Transport-Server' 'Edge Transport servers'), (SafeVal ($hyb2.EdgeTransportServers -join ', ')))
                    @((L 'Client Access Server' 'Client Access servers'), (SafeVal ($hyb2.ClientAccessServers -join ', ')))
                    @((L 'Empfangs-Connector' 'Receive connector'), (SafeVal ($hyb2.ReceivingTransportServers -join ', ')))
                    @((L 'Sende-Connector' 'Send connector'), (SafeVal ($hyb2.SendingTransportServers -join ', ')))
                    @((L 'Externe SMTP-Domänen' 'External SMTP domains'), (SafeVal ($hyb2.ExternalIPAddresses -join ', ')))
                    @((L 'TLS-Zertifikatsname' 'TLS certificate name'), (SafeVal $hyb2.TlsCertificateName))
                )
                $null = $parts.Add((New-WdTable -Headers @((L 'Hybrid-Eigenschaft' 'Hybrid property'), (L 'Wert' 'Value')) -Rows $hybRows2))
                $null = $parts.Add((New-WdParagraph (L 'Hinweis: Hybrid Configuration Wizard (HCW) prüft und aktualisiert diese Einstellungen automatisch. Änderungen sollten stets über den HCW oder Set-HybridConfiguration erfolgen, nicht über manuelle ADSIEdit- oder Registry-Eingriffe.' 'Note: Hybrid Configuration Wizard (HCW) validates and updates these settings automatically. Changes should always be made via HCW or Set-HybridConfiguration, never via manual ADSIEdit or registry edits.')))
            }
            if ($orgD.IntraOrgConnectors -and $orgD.IntraOrgConnectors.Count -gt 0) {
                $iocRows = $orgD.IntraOrgConnectors | ForEach-Object { @($_.Name, (SafeVal $_.TargetAddressDomains), (SafeVal $_.DiscoveryEndpoint), (Lc $_.Enabled (L 'Aktiv' 'Active') (L 'Inaktiv' 'Inactive'))) }
                $null = $parts.Add((New-WdTable -Headers @('IntraOrg Connector', (L 'Zieldomänen' 'Target domains'), 'Discovery', (L 'Status' 'Status')) -Rows $iocRows))
            }
            if (-not $orgD.FederationTrust -and -not $orgD.HybridConfig -and -not ($orgD.IntraOrgConnectors | Where-Object { $_ })) {
                $null = $parts.Add((New-WdParagraph (L '(Keine Federation/Hybrid-Konfiguration vorhanden — reine on-premises Umgebung)' '(No federation/hybrid configuration present — on-premises only environment)')))
            }

            # 4.12 AuthConfig + Auth-Zertifikat
            $null = $parts.Add((New-WdHeading (L '4.12 Auth-Zertifikat und OAuth-Konfiguration' '4.12 Auth Certificate and OAuth Configuration') 2))
            $null = $parts.Add((New-WdParagraph (L 'Das Auth-Zertifikat ist das zentrale Sicherheitsobjekt für die server-interne OAuth-Kommunikation (OAuth 2.0). Es signiert die Token, die Exchange-Dienste untereinander und gegenüber Exchange Online austauschen. Die Lebensdauer beträgt standardmäßig 5 Jahre; läuft das Zertifikat ab, schlägt OAuth fehl (Hybrid-Szenarien, Exchange Online Federation, OWA/ECP-Rückfragen auf andere Server). MEAC (MonitorExchangeAuthCertificate.ps1) übernimmt die automatische Erneuerung 60 Tage vor Ablauf durch einen geplanten Task (siehe Kapitel 7).' 'The Auth Certificate is the central security artifact for server-internal OAuth communication (OAuth 2.0). It signs the tokens that Exchange services exchange among themselves and with Exchange Online. Default lifetime is 5 years; once it expires OAuth fails (hybrid scenarios, Exchange Online federation, OWA/ECP cross-server calls). MEAC (MonitorExchangeAuthCertificate.ps1) handles automatic renewal 60 days before expiry via a scheduled task (see chapter 7).')))
            if ($orgD.AuthConfig) {
                $ac = $orgD.AuthConfig
                $fmtTp = {
                    param($thumb)
                    if (-not $thumb) { return (L '(nicht gesetzt)' '(not set)') }
                    if ($cust)       { return ('{0}...' -f $thumb.Substring(0, [Math]::Min(8, $thumb.Length))) }
                    [string]$thumb
                }
                $tp     = & $fmtTp $ac.CurrentCertificateThumbprint
                $tpNext = & $fmtTp $ac.NextCertificateThumbprint
                $tpPrev = & $fmtTp $ac.PreviousCertificateThumbprint
                # Auth cert validity: AuthConfig does not expose NotAfter directly — look up the cert
                # by thumbprint from the local server's Exchange cert store.
                $validUntil = (L '(unbekannt)' '(unknown)')
                $daysLeft   = $null
                if ($ac.CurrentCertificateThumbprint) {
                    try {
                        $authCert = Get-ExchangeCertificate -Thumbprint $ac.CurrentCertificateThumbprint -Server $env:COMPUTERNAME -ErrorAction Stop
                        if ($authCert -and $authCert.NotAfter) {
                            $validUntil = $authCert.NotAfter.ToString('yyyy-MM-dd')
                            $daysLeft = [int]([Math]::Floor(($authCert.NotAfter - (Get-Date)).TotalDays))
                        }
                    } catch {
                        try {
                            $certStore = Get-ChildItem -Path 'Cert:\LocalMachine\My' -ErrorAction Stop | Where-Object { $_.Thumbprint -eq $ac.CurrentCertificateThumbprint } | Select-Object -First 1
                            if ($certStore) {
                                $validUntil = $certStore.NotAfter.ToString('yyyy-MM-dd')
                                $daysLeft = [int]([Math]::Floor(($certStore.NotAfter - (Get-Date)).TotalDays))
                            }
                        } catch { }
                    }
                }
                $validUntilCell = if ($null -ne $daysLeft) { ('{0} ({1} Tage verbleibend / {1} days remaining)' -f $validUntil, $daysLeft) } else { $validUntil }
                $authRows = [System.Collections.Generic.List[object[]]]::new()
                $authRows.Add(@((L 'Aktuelles Auth-Zertifikat (Fingerabdruck)' 'Current Auth cert thumbprint'), $tp))
                $authRows.Add(@((L 'Gültig bis' 'Valid until'), $validUntilCell))
                $authRows.Add(@((L 'Nächstes Auth-Zertifikat' 'Next Auth certificate'), $tpNext))
                $authRows.Add(@((L 'Vorheriges Auth-Zertifikat' 'Previous Auth certificate'), $tpPrev))
                $authRows.Add(@((L 'Realm' 'Realm'), (SafeVal $ac.Realm (L '(leer — Default)' '(empty — default)'))))
                $authRows.Add(@((L 'Service Name' 'Service name'), (SafeVal $ac.ServiceName (L '(nicht gesetzt)' '(not set)'))))
                $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $authRows.ToArray()))
            } else {
                $null = $parts.Add((New-WdParagraph (L '(AuthConfig nicht abrufbar)' '(AuthConfig not available)')))
            }
            $null = $parts.Add((New-WdParagraph (L 'Wichtig: Eine manuelle Rotation des Auth-Zertifikats wird ausschließlich im Notfall empfohlen. Reguläre Rotation erfolgt über den MEAC-Task oder per Set-AuthConfig -PublishCertificate nach vorheriger Erzeugung eines "Next"-Zertifikats. Nach einer Rotation ist IISRESET auf allen Exchange-Servern erforderlich.' 'Important: Manual rotation of the Auth Certificate is only recommended as an emergency procedure. Regular rotation is handled by the MEAC task or via Set-AuthConfig -PublishCertificate after creating a "Next" certificate. After any rotation an IISRESET is required on all Exchange servers.')))

            # 4.13 Namensräume-Übersicht
            $null = $parts.Add((New-WdHeading (L '4.13 Namensräume — konsolidierte Übersicht' '4.13 Namespaces — Consolidated Overview') 2))
            $null = $parts.Add((New-WdParagraph (L 'Diese Tabelle aggregiert die Internal- und External-URLs aller Client-zugewandten Dienste über alle Exchange-Server hinweg. Identische URLs über alle Server sind Voraussetzung für Load Balancing ohne Session Affinity (ab Exchange 2016). Abweichende URLs innerhalb eines Dienstes deuten auf inkonsistente Namespace-Konfiguration hin und sollten korrigiert werden.' 'This table aggregates internal and external URLs for all client-facing services across all Exchange servers. Identical URLs across all servers are a prerequisite for load balancing without session affinity (since Exchange 2016). Diverging URLs within one service indicate inconsistent namespace configuration and should be corrected.')))
            $nsRows = [System.Collections.Generic.List[object[]]]::new()
            $vdirServices = @(
                @{ Name='OWA'        ; Prop='VDirOWA'  }
                @{ Name='ECP'        ; Prop='VDirECP'  }
                @{ Name='EWS'        ; Prop='VDirEWS'  }
                @{ Name='OAB'        ; Prop='VDirOAB'  }
                @{ Name='ActiveSync' ; Prop='VDirAS'   }
                @{ Name='MAPI'       ; Prop='VDirMAPI' }
                @{ Name='PowerShell' ; Prop='VDirPW'   }
            )
            foreach ($svc in $vdirServices) {
                $intUrls = @(); $extUrls = @()
                foreach ($srv2 in $rd.Servers) {
                    $vd = $srv2.($svc.Prop) | Select-Object -First 1
                    if ($vd) {
                        if ($vd.InternalUrl) { $intUrls += $vd.InternalUrl.AbsoluteUri }
                        if ($vd.ExternalUrl) { $extUrls += $vd.ExternalUrl.AbsoluteUri }
                    }
                }
                $intU = if ($intUrls) { ($intUrls | Select-Object -Unique) -join ', ' } else { (L '(nicht gesetzt)' '(not set)') }
                $extU = if ($extUrls) { ($extUrls | Select-Object -Unique) -join ', ' } else { (L '(nicht gesetzt)' '(not set)') }
                $consistency = if (($intUrls | Select-Object -Unique).Count -le 1 -and ($extUrls | Select-Object -Unique).Count -le 1) { (L 'konsistent' 'consistent') } else { (L 'ABWEICHUNG' 'DIVERGENT') }
                $nsRows.Add(@($svc.Name, (Mask-Ip $intU), (Mask-Ip $extU), $consistency))
            }
            $autodiscoverUrls = @($rd.Servers | ForEach-Object { if ($_.AutodiscoverSCP -and $_.AutodiscoverSCP.AutoDiscoverServiceInternalUri) { $_.AutodiscoverSCP.AutoDiscoverServiceInternalUri.ToString() } } | Where-Object { $_ })
            if ($autodiscoverUrls) {
                $adIn = ($autodiscoverUrls | Select-Object -Unique) -join ', '
                $adC  = if (($autodiscoverUrls | Select-Object -Unique).Count -le 1) { (L 'konsistent' 'consistent') } else { (L 'ABWEICHUNG' 'DIVERGENT') }
                $nsRows.Add(@('Autodiscover SCP', (Mask-Ip $adIn), '—', $adC))
            }
            $null = $parts.Add((New-WdTable -Headers @((L 'Dienst' 'Service'), (L 'Interne URL' 'Internal URL'), (L 'Externe URL' 'External URL'), (L 'Konsistenz' 'Consistency')) -Rows $nsRows.ToArray()))

            # 4.14 Datenbank-Kopien-Status (DAG-übergreifend)
            $anyCopies = @($rd.Servers | ForEach-Object { $_.DatabaseCopies } | Where-Object { $_ })
            if ($anyCopies.Count -gt 0) {
                $null = $parts.Add((New-WdHeading (L '4.14 Datenbank-Kopien-Status' '4.14 Database Copy Status') 2))
                $null = $parts.Add((New-WdParagraph (L 'Der Status aller Datenbankkopien wird serverübergreifend erfasst. CopyQueueLength bezeichnet die Anzahl der noch nicht replizierten Log-Dateien auf die Kopie, ReplayQueueLength die Anzahl der noch nicht eingespielten Logs. Im Normalbetrieb sollten beide Werte einstellig bleiben. ContentIndexState = "Healthy" ist erforderlich für die Postfachsuche. Eine dauerhaft hohe Queue deutet auf Netzwerk- oder I/O-Probleme hin.' 'The status of all database copies is collected across all servers. CopyQueueLength is the number of log files not yet replicated to the copy, ReplayQueueLength the number of logs not yet replayed. In normal operation both values should stay single-digit. ContentIndexState = "Healthy" is required for mailbox search. A persistently high queue indicates network or I/O problems.')))
                $dcRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($srv2 in $rd.Servers) {
                    foreach ($dc in $srv2.DatabaseCopies) {
                        $dcRows.Add(@($dc.DatabaseName, $dc.MailboxServer, $dc.Status, $dc.CopyQueueLength, $dc.ReplayQueueLength, (SafeVal $dc.ContentIndexState), (SafeVal $dc.ActivationPreference)))
                    }
                }
                $null = $parts.Add((New-WdTable -Headers @((L 'Datenbank' 'Database'), (L 'Server' 'Server'), (L 'Status' 'Status'), 'Copy-Q', 'Replay-Q', (L 'Suchindex' 'Content index'), (L 'AktPref' 'ActPref')) -Rows $dcRows.ToArray()))
            }

            # 4.15 RBAC — Rollengruppen
            if ($orgD.RoleGroups -and $orgD.RoleGroups.Count -gt 0) {
                $null = $parts.Add((New-WdHeading (L '4.15 RBAC — Rollengruppen' '4.15 RBAC — Role Groups') 2))
                $null = $parts.Add((New-WdParagraph (L 'Role-Based Access Control (RBAC) steuert, welche Exchange-Cmdlets und -Parameter ein Benutzer ausführen darf. Built-in-Rollengruppen wie "Organization Management", "Recipient Management" oder "View-Only Organization Management" werden von Exchange bereitgestellt. Benutzerdefinierte Rollengruppen erlauben feingranulare Delegation (z. B. Helpdesk ohne Zugriff auf Transport oder Hybrid). Diese Tabelle zeigt alle Rollengruppen mit ihren Mitgliedern — eine Dokumentation ist wichtig für Audits und Zugriffskontrollen.' 'Role-Based Access Control (RBAC) governs which Exchange cmdlets and parameters a user may run. Built-in role groups such as "Organization Management", "Recipient Management" or "View-Only Organization Management" are provided by Exchange. Custom role groups allow fine-grained delegation (e.g. helpdesk without access to transport or hybrid). This table lists all role groups with their members — documentation matters for audits and access reviews.')))
                $rgRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($rg in $orgD.RoleGroups) {
                    $memStr = if ($rg.Members -and $rg.Members.Count -gt 0) {
                        ($rg.Members | ForEach-Object { if ($cust) { ('{0} ({1})' -f (Mask-Val $_.Name), $_.Type) } else { ('{0} ({1})' -f $_.Name, $_.Type) } }) -join '; '
                    } else { (L '(keine Mitglieder)' '(no members)') }
                    $rgRows.Add(@($rg.Name, (SafeVal $rg.Description), $memStr))
                }
                $null = $parts.Add((New-WdTable -Headers @((L 'Rollengruppe' 'Role group'), (L 'Beschreibung' 'Description'), (L 'Mitglieder' 'Members')) -Rows $rgRows.ToArray()))
                $null = $parts.Add((New-WdParagraph (L 'Hinweis: Eine detaillierte RBAC-Aufstellung mit verwalteten Rollen liefert der Befehl Get-RoleGroup | Format-List und Get-ManagementRoleAssignment. EXpress legt optional einen separaten RBAC-Report (.txt) im Reports-Verzeichnis ab.' 'Note: A detailed RBAC listing with managed roles is available via Get-RoleGroup | Format-List and Get-ManagementRoleAssignment. EXpress optionally writes a separate RBAC report (.txt) to the reports directory.')))
            }

            # 4.16 Audit-Konfiguration
            $null = $parts.Add((New-WdHeading (L '4.16 Audit-Konfiguration' '4.16 Audit Configuration') 2))
            $null = $parts.Add((New-WdParagraph (L 'Das Admin-Auditprotokoll zeichnet alle Exchange-Verwaltungscmdlets auf, die von Administratoren ausgeführt werden (wer hat wann was geändert). Es ist Grundlage für Compliance-Anforderungen wie ISO 27001, BSI-Grundschutz und DSGVO-Rechenschaftspflicht. Das Protokoll wird in einem dedizierten verborgenen Postfach in der Exchange-Organisation gespeichert und kann per Search-AdminAuditLog abgefragt werden. Die Aufbewahrungsfrist (AdminAuditLogAgeLimit) bestimmt, wie lange Einträge erhalten bleiben (Standard: 90 Tage).' 'The admin audit log records all Exchange management cmdlets executed by administrators (who changed what and when). It is the basis for compliance requirements such as ISO 27001, BSI baseline protection and GDPR accountability. The log is stored in a dedicated hidden mailbox in the Exchange organisation and can be queried via Search-AdminAuditLog. The retention period (AdminAuditLogAgeLimit) determines how long entries are kept (default: 90 days).')))
            if ($orgD.AdminAuditLog) {
                $aal = $orgD.AdminAuditLog
                $aalRows = [System.Collections.Generic.List[object[]]]::new()
                $aalRows.Add(@((L 'Admin-Auditprotokoll aktiviert' 'Admin audit log enabled'),  (Format-RegBool $aal.AdminAuditLogEnabled)))
                $aalRows.Add(@((L 'Aufbewahrungsfrist' 'Log age limit'),                         (SafeVal $aal.AdminAuditLogAgeLimit (L '(Standard: 90 Tage)' '(default: 90 days)'))))
                $aalRows.Add(@((L 'Log-Postfach' 'Log mailbox'),                                 (SafeVal $aal.AdminAuditLogMailbox   (L '(Standard — automatisch)' '(default — automatic)'))))
                $aalCmdlets    = if ($aal.AdminAuditLogCmdlets)    { $aal.AdminAuditLogCmdlets -join ', '    } else { $null }
                $aalExclusions = if ($aal.AdminAuditLogExclusions) { $aal.AdminAuditLogExclusions -join ', ' } else { $null }
                $aalRows.Add(@((L 'Aufgezeichnete Cmdlets' 'Logged cmdlets'),  (SafeVal $aalCmdlets    (L '(alle — Standard)' '(all — default)'))))
                $aalRows.Add(@((L 'Ausgeschlossene Cmdlets' 'Excluded cmdlets'), (SafeVal $aalExclusions (L '(keine)' '(none)'))))
                $aalRows.Add(@((L 'Test-Cmdlet-Protokollierung' 'Test cmdlet logging'),          (Format-RegBool $aal.TestCmdletLoggingEnabled)))
                $aalRows.Add(@((L 'Log-Level' 'Log level'),                                      (SafeVal $aal.LogLevel)))
                $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $aalRows.ToArray()))
            } else {
                $null = $parts.Add((New-WdParagraph (L '(Admin-Auditprotokoll-Konfiguration nicht abrufbar)' '(Admin audit log configuration not available)')))
            }

            # 4.17 Service-Accounts und Berechtigungen (Exchange RBAC)
            $null = $parts.Add((New-WdHeading (L '4.17 Service-Accounts und Berechtigungen' '4.17 Service Accounts and Permissions') 2))
            $null = $parts.Add((New-WdParagraph (L 'Exchange Server verwendet rollenbasierte Zugriffssteuerung (RBAC). Die folgende Tabelle dokumentiert die Mitglieder der wichtigsten Exchange-Rollengruppen. Privilegierte Konten sollten auf das Minimum beschränkt sein (Principle of Least Privilege). Dienstkonten für externe Integrationen (Backup, Monitoring, Archivierung) sollten dedizierte AD-Konten mit minimalen Exchange-Berechtigungen nutzen.' 'Exchange Server uses Role-Based Access Control (RBAC). The table below documents the members of the most important Exchange role groups. Privileged accounts should be limited to the minimum necessary (Principle of Least Privilege). Service accounts for external integrations (backup, monitoring, archiving) should use dedicated AD accounts with minimum Exchange permissions.')))
            $rbacRoles2 = @('Organization Management','Server Management','Recipient Management','Hygiene Management','Compliance Management','View-Only Organization Management')
            $rbacRows2 = [System.Collections.Generic.List[object[]]]::new()
            foreach ($rg2 in $rbacRoles2) {
                try {
                    $rgMembers = @(Get-RoleGroupMember $rg2 -ErrorAction SilentlyContinue)
                    $rgMemberList = if ($rgMembers -and $rgMembers.Count -gt 0) { ($rgMembers | ForEach-Object { if ($_.Name) { $_.Name } else { $_.DisplayName } }) -join "`n" } else { (L '(leer)' '(empty)') }
                    $rbacRows2.Add(@($rg2, $rgMemberList))
                } catch {
                    $rbacRows2.Add(@($rg2, (L '(nicht abfragbar)' '(not available)')))
                }
            }
            if ($rbacRows2.Count -eq 0) { $rbacRows2.Add(@((L '(RBAC-Daten nicht abrufbar)' '(RBAC data not available)'), '')) }
            $null = $parts.Add((New-WdTable -Headers @((L 'Rollengruppe' 'Role group'), (L 'Mitglieder' 'Members')) -Rows $rbacRows2.ToArray()))

            # Exchange Online / Microsoft 365 was formerly 4.17 inside "Organisation".
            # Moved to its own top-level section 15 (before Operative Runbooks) — belongs
            # to customer-ops context, not org-config telemetry. See below, just before
            # "Operative Runbooks".
        }

        # ── 5. Server in der Organisation ────────────────────────────────────────
        if ($scope -in 'All','Local') {
            $null = $parts.Add((New-WdHeading (L '5. Server in der Organisation' '5. Servers in the Organisation') 1))
            $null = $parts.Add((New-WdParagraph (L 'Die folgenden Abschnitte dokumentieren jeden Exchange-Server in der Organisation. Der neu installierte Server ist mit dem Hinweis "← Neu installiert" gekennzeichnet. Systemdetails (Hardware, Volumes, NICs) werden über WinRM/CIM abgefragt — bei nicht erreichbaren Servern erscheint ein entsprechender Hinweis.' 'The following sections document each Exchange server in the organisation. The newly installed server is marked with "← Newly installed". System details (hardware, volumes, NICs) are retrieved via WinRM/CIM — for unreachable servers a corresponding note is shown.')))
            if ($rd.Servers.Count -eq 0) {
                $null = $parts.Add((New-WdParagraph (L '(Keine Exchange-Server abfragbar)' '(No Exchange servers available)')))
            }
            $srvCounter = 0
            foreach ($srvD in $rd.Servers) {
                $srvCounter++
                $srvName   = $srvD.ServerName
                $isLocal   = $srvD.IsLocalServer
                $exSrv2    = $srvD.ExServer
                $srvLabel  = if ($isLocal) { ('{0} ← {1}' -f $srvName, (L 'Neu installiert / lokaler Server' 'Newly installed / local server')) } else { $srvName }
                $null = $parts.Add((New-WdHeading ('5.{0} {1}' -f $srvCounter, $srvLabel) 2))

                # 5.x.1 Identität
                $null = $parts.Add((New-WdHeading (L 'Identität' 'Identity') 3))
                $idRows = [System.Collections.Generic.List[object[]]]::new()
                if ($exSrv2) {
                    $idRows.Add(@((L 'Exchange-Version' 'Exchange version'), $exSrv2.AdminDisplayVersion.ToString()))
                    $idRows.Add(@('FQDN', (SafeVal $exSrv2.Fqdn)))
                    $idRows.Add(@((L 'Serverrolle' 'Server role'), ($exSrv2.ServerRole -join ', ')))
                    $idRows.Add(@((L 'Edition' 'Edition'), $exSrv2.Edition.ToString()))
                    $idRows.Add(@((L 'AD-Standort' 'AD site'), $exSrv2.Site.ToString()))
                    $idRows.Add(@((L 'Installiert am' 'Installed on'), (SafeVal $exSrv2.WhenCreated)))
                }
                if ($srvD.AutodiscoverSCP) { $idRows.Add(@('Autodiscover SCP', (SafeVal $srvD.AutodiscoverSCP.AutoDiscoverServiceInternalUri))) }
                $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $idRows.ToArray()))

                # 5.x.2 Systemdetails (CIM)
                $null = $parts.Add((New-WdHeading (L 'Systemdetails' 'System Details') 3))
                $sysDetailRows = Format-RemoteSysRows $srvD.RemoteData
                $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $sysDetailRows.ToArray()))

                # 5.x.3 Datenbanken
                $null = $parts.Add((New-WdHeading (L 'Postfachdatenbanken' 'Mailbox Databases') 3))
                $dbRows2 = [System.Collections.Generic.List[object[]]]::new()
                foreach ($db2 in $srvD.Databases) {
                    $mounted2 = if ($null -ne $db2.Mounted) { if ($db2.Mounted) { (L 'Eingehängt' 'Mounted') } else { (L 'Ausgehängt' 'Dismounted') } } else { (L 'Unbekannt' 'Unknown') }
                    $dbRows2.Add(@($db2.Name, (SafeVal $db2.EdbFilePath), (SafeVal $db2.LogFolderPath), $mounted2))
                }
                if ($dbRows2.Count -eq 0) { $dbRows2.Add(@((L '(keine Datenbank auf diesem Server)' '(no database on this server)'), '', '', '')) }
                $null = $parts.Add((New-WdTable -Headers @((L 'Datenbank' 'Database'), (L 'DB-Pfad' 'DB path'), (L 'Log-Pfad' 'Log path'), (L 'Status' 'Status')) -Rows $dbRows2.ToArray()))

                # 5.x.4 Virtuelle Verzeichnisse
                $null = $parts.Add((New-WdHeading (L 'Virtuelle Verzeichnisse' 'Virtual Directories') 3))
                $vd2Rows = [System.Collections.Generic.List[object[]]]::new()
                $vdirSources = @(
                    @{ Name='OWA';        Data=$srvD.VDirOWA  }
                    @{ Name='ECP';        Data=$srvD.VDirECP  }
                    @{ Name='EWS';        Data=$srvD.VDirEWS  }
                    @{ Name='OAB';        Data=$srvD.VDirOAB  }
                    @{ Name='ActiveSync'; Data=$srvD.VDirAS   }
                    @{ Name='MAPI';       Data=$srvD.VDirMAPI }
                )
                foreach ($vde in $vdirSources) {
                    $vd3 = $vde.Data | Select-Object -First 1
                    if ($vd3) {
                        $int2 = if ($vd3.InternalUrl) { $vd3.InternalUrl.AbsoluteUri } else { (L '(nicht gesetzt)' '(not set)') }
                        $ext2 = if ($vd3.ExternalUrl) { $vd3.ExternalUrl.AbsoluteUri } else { (L '(nicht gesetzt)' '(not set)') }
                        $vd2Rows.Add(@($vde.Name, (Mask-Ip $int2), (Mask-Ip $ext2)))
                    }
                }
                $null = $parts.Add((New-WdTable -Headers @((L 'Dienst' 'Service'), (L 'Intern' 'Internal'), (L 'Extern' 'External')) -Rows $vd2Rows.ToArray()))

                # 5.x.5 Receive Connectors — split into two tables (network / security)
                # A single 8-column table wraps every cell in portrait Word; splitting into
                # 4 + 5 columns keeps each row legible. Name repeats as the join key.
                $null = $parts.Add((New-WdHeading (L 'Receive Connectors' 'Receive Connectors') 3))
                $rcNetRows = [System.Collections.Generic.List[object[]]]::new()
                $rcSecRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($rc in $srvD.ReceiveConnectors) {
                    $reqTlsRc = Lc ([bool]$rc.RequireTLS) (L 'ja' 'yes') (L 'nein' 'no')
                    $maxMsgRc = if ($rc.MaxMessageSize) { $rc.MaxMessageSize.ToString() } else { '—' }
                    $rcNetRows.Add(@($rc.Name, (Mask-Ip ($rc.Bindings -join ', ')), (Mask-Ip ($rc.RemoteIPRanges -join ', ')), (SafeVal $rc.Fqdn '—')))
                    $rcSecRows.Add(@($rc.Name, $rc.AuthMechanism, $rc.PermissionGroups, $reqTlsRc, $maxMsgRc))
                }
                if ($rcNetRows.Count -eq 0) {
                    $rcNetRows.Add(@((L '(keine)' '(none)'), '', '', ''))
                    $rcSecRows.Add(@((L '(keine)' '(none)'), '', '', '', ''))
                }
                $null = $parts.Add((New-WdParagraph (L 'Netzwerk:' 'Network:')))
                $null = $parts.Add((New-WdTable -Compact -Headers @((L 'Connector' 'Connector'), 'Bindings', (L 'Remote-IPs' 'Remote IPs'), 'FQDN') -Rows $rcNetRows.ToArray()))
                $null = $parts.Add((New-WdParagraph (L 'Sicherheit und Limits:' 'Security and limits:')))
                $null = $parts.Add((New-WdTable -Compact -Headers @((L 'Connector' 'Connector'), 'Auth', (L 'Berechtigungen' 'Permissions'), 'TLS', (L 'Max. Größe' 'Max size')) -Rows $rcSecRows.ToArray()))

                # 5.x.6 IMAP/POP3-Konfiguration (nur lokaler Server)
                if ($srvD.IsLocalServer -and ($srvD.ImapSettings -or $srvD.PopSettings)) {
                    $null = $parts.Add((New-WdHeading (L 'IMAP/POP3-Konfiguration' 'IMAP/POP3 Configuration') 3))
                    $protoSrvRows = [System.Collections.Generic.List[object[]]]::new()
                    if ($srvD.ImapSettings) {
                        $im = $srvD.ImapSettings
                        $protoSrvRows.Add(@((L 'IMAP4 — Externer Namespace' 'IMAP4 — External namespace'),      (SafeVal (($im.ExternalConnectionSettings | ForEach-Object { $_.ToString() }) -join '; ') (L '(nicht gesetzt — bitte manuell ergänzen)' '(not set — please fill in manually)'))))
                        $protoSrvRows.Add(@((L 'IMAP4 — Interner Namespace' 'IMAP4 — Internal namespace'),      (SafeVal (($im.InternalConnectionSettings | ForEach-Object { $_.ToString() }) -join '; ') (L '(nicht gesetzt)' '(not set)'))))
                        $protoSrvRows.Add(@((L 'IMAP4 — X.509-Zertifikatname' 'IMAP4 — X.509 certificate name'), (SafeVal $im.X509CertificateName (L '(nicht gesetzt)' '(not set)'))))
                        $protoSrvRows.Add(@((L 'IMAP4 — Anmeldetyp' 'IMAP4 — Login type'),                       (SafeVal $im.LoginType)))
                    }
                    if ($srvD.PopSettings) {
                        $pop = $srvD.PopSettings
                        $protoSrvRows.Add(@((L 'POP3 — Externer Namespace' 'POP3 — External namespace'),         (SafeVal (($pop.ExternalConnectionSettings | ForEach-Object { $_.ToString() }) -join '; ') (L '(nicht gesetzt — bitte manuell ergänzen)' '(not set — please fill in manually)'))))
                        $protoSrvRows.Add(@((L 'POP3 — Interner Namespace' 'POP3 — Internal namespace'),         (SafeVal (($pop.InternalConnectionSettings | ForEach-Object { $_.ToString() }) -join '; ') (L '(nicht gesetzt)' '(not set)'))))
                        $protoSrvRows.Add(@((L 'POP3 — X.509-Zertifikatname' 'POP3 — X.509 certificate name'),  (SafeVal $pop.X509CertificateName (L '(nicht gesetzt)' '(not set)'))))
                    }
                    $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $protoSrvRows.ToArray()))
                }

                # 5.x.7 Zertifikate
                $null = $parts.Add((New-WdHeading (L 'Zertifikate' 'Certificates') 3))
                $certRows2 = [System.Collections.Generic.List[object[]]]::new()
                foreach ($cert2 in $srvD.Certificates) {
                    $expiry2   = if ($cert2.NotAfter) { $cert2.NotAfter.ToString('yyyy-MM-dd') } else { '?' }
                    $daysLeft2 = if ($cert2.NotAfter) { [int][Math]::Floor(($cert2.NotAfter - (Get-Date)).TotalDays) } else { 0 }
                    $tp2 = if ($cust) { ('{0}...' -f $cert2.Thumbprint.Substring(0, [Math]::Min(8, $cert2.Thumbprint.Length))) } else { $cert2.Thumbprint }
                    $certRows2.Add(@($cert2.Subject, $expiry2, ('{0}d' -f $daysLeft2), $cert2.Services, $tp2))
                }
                if ($certRows2.Count -eq 0) { $certRows2.Add(@((L '(keine)' '(none)'), '', '', '', '')) }
                $null = $parts.Add((New-WdTable -Compact -Headers @('Subject', (L 'Ablauf' 'Expiry'), (L 'Verbleibend' 'Remaining'), (L 'Dienste' 'Services'), (L 'Fingerabdruck' 'Thumbprint')) -Rows $certRows2.ToArray()))

                # 5.x.7 Transport Agents
                if ($srvD.TransportAgents -and $srvD.TransportAgents.Count -gt 0) {
                    $null = $parts.Add((New-WdHeading (L 'Transport Agents' 'Transport Agents') 3))
                    $taRows = [System.Collections.Generic.List[object[]]]::new()
                    foreach ($ta in $srvD.TransportAgents) {
                        $taState = if ($ta.Enabled) { (L 'Aktiv' 'Enabled') } else { (L 'Inaktiv' 'Disabled') }
                        # TransportAgent.Name can be empty after implicit remoting deserialization; fall back to Identity.
                        $taName = if ($ta.Name) { [string]$ta.Name } elseif ($ta.Identity) { [string]$ta.Identity } else { '(unbenannt)' }
                        $taRows.Add(@($taName, $taState, $ta.Priority))
                    }
                    $null = $parts.Add((New-WdTable -Headers @('Agent', (L 'Status' 'Status'), (L 'Priorität' 'Priority')) -Rows $taRows.ToArray()))
                }
            }
        }

        # ── 6. Netzwerk & DNS (lokal) ─────────────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '6. Netzwerk und DNS (lokaler Server)' '6. Network and DNS (local server)') 1))
        $null = $parts.Add((New-WdParagraph (L 'Die folgende Tabelle zeigt die Netzwerkkonfiguration des lokalen Exchange-Servers. Für Exchange Server ist eine korrekte DNS-Auflösung (Forward und Reverse) eine grundlegende Betriebsvoraussetzung. Als DNS-Server müssen ausschließlich Active-Directory-integrierte DNS-Server der eigenen Domäne eingetragen sein — kein öffentlicher DNS (z. B. 8.8.8.8), da Exchange für Autodiscover, SCP-Lookups und interne Namensauflösung auf AD-DNS angewiesen ist.' 'The table below shows the network configuration of the local Exchange server. Correct DNS resolution (forward and reverse) is a fundamental operational requirement for Exchange Server. Only Active Directory-integrated DNS servers of the own domain must be configured — no public DNS (e.g. 8.8.8.8), as Exchange relies on AD DNS for Autodiscover, SCP lookups and internal name resolution.')))
        $netRows = [System.Collections.Generic.List[object[]]]::new()
        try {
            $nicIPs = @{}; $nicDNS = @{}
            Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.InterfaceAlias -notlike '*Loopback*' } | ForEach-Object { $nicIPs[$_.InterfaceAlias] = ('{0}/{1}' -f $_.IPAddress, $_.PrefixLength) }
            Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.InterfaceAlias -notlike '*Loopback*' -and $_.ServerAddresses } | ForEach-Object { $nicDNS[$_.InterfaceAlias] = ($_.ServerAddresses -join ', ') }
            foreach ($nic in ($nicIPs.Keys | Sort-Object)) {
                $ip2  = Mask-Ip $nicIPs[$nic]
                $dns2 = if ($nicDNS[$nic]) { Mask-Ip $nicDNS[$nic] } else { (L '(nicht gesetzt)' '(not set)') }
                $netRows.Add(@(('NIC: {0}' -f $nic), ('{0} — DNS: {1}' -f $ip2, $dns2)))
            }
        } catch { }
        if ($netRows.Count -eq 0) { $netRows.Add(@((L '(keine NIC-Daten abrufbar)' '(no NIC data available)'), '')) }
        $null = $parts.Add((New-WdTable -Headers @((L 'NIC / Eigenschaft' 'NIC / Property'), (L 'Wert' 'Value')) -Rows $netRows.ToArray()))

        # 6.1 DNS-Einträge (relevant für Exchange-Dienste)
        $null = $parts.Add((New-WdHeading (L '6.1 DNS-Einträge (Exchange-Dienste)' '6.1 DNS Records (Exchange services)') 2))
        $null = $parts.Add((New-WdParagraph (L 'Für einen Exchange-Server müssen grundsätzlich die folgenden öffentlichen DNS-Einträge je SMTP-Domäne auf dem für diese Domäne zuständigen DNS eingetragen sein: autodiscover.<domain> (A oder CNAME auf den externen Namespace), MX (zeigt auf den eingehenden Mailflow — direkt auf den Exchange-Namespace, einen Smarthost oder einen eingehenden Cloud-Filter), sowie die Authentifizierungseinträge SPF (TXT), DKIM (TXT via Selektor) und DMARC (_dmarc.<domain> TXT). Bei Hybrid-Szenarien kommt ein CNAME auf onmicrosoft.com hinzu, außerdem ggf. _autodiscover._tcp SRV.' 'For an Exchange server the following public DNS records must exist per SMTP domain on the DNS authoritative for that domain: autodiscover.<domain> (A or CNAME pointing to the external namespace), MX (controls incoming mail flow — directly to the Exchange namespace, to a smart host, or to an inbound cloud filter), and the authentication records SPF (TXT), DKIM (TXT via selector) and DMARC (_dmarc.<domain> TXT). In hybrid scenarios a CNAME to onmicrosoft.com is added, plus optionally _autodiscover._tcp SRV.')))
        $null = $parts.Add((New-WdParagraph (L 'Hinweis: In Split-DNS-Szenarien (AD-Domäne entspricht einer gerouteten SMTP-Domäne) existieren diese Einträge zusätzlich auf dem internen AD-DNS; für rein interne AD-Domänen (z. B. .local/.lan) sind MX/SPF/DKIM/DMARC nicht relevant. Eine automatische Auflösung aller Einträge aus dem Server heraus ist nicht aussagekräftig, da die Antworten je nach DNS-View (intern/extern) abweichen und sich externe Einträge typischerweise erst nach Umzug der primären Maildomäne bzw. mit weiteren akzeptierten Domänen ergänzen.' 'Note: In split-DNS scenarios (AD domain identical to a routed SMTP domain) these records also exist on the internal AD DNS; for purely internal AD domains (e.g. .local/.lan) MX/SPF/DKIM/DMARC are not relevant. Automatic resolution of all records from the server itself is not conclusive, since answers differ depending on the DNS view (internal/external), and external records are typically added only after cut-over of the primary mail domain or when additional accepted domains are configured.')))

        # Autodiscover SCP (internal clients) — always sensible to document for a fresh server
        $scpRows = [System.Collections.Generic.List[object[]]]::new()
        try {
            $casList = Get-ClientAccessService -ErrorAction SilentlyContinue
            foreach ($cas in $casList) {
                $scpRows.Add(@($cas.Name, (SafeVal ([string]$cas.AutoDiscoverServiceInternalUri))))
            }
        } catch { }
        if ($scpRows.Count -gt 0) {
            $null = $parts.Add((New-WdParagraph (L 'Autodiscover Service Connection Point (SCP) — für domänenmitgliedschaftsfähige Clients im internen Netzwerk maßgeblich. Wird im AD (CN=Configuration) gespeichert und von Outlook bevorzugt vor DNS-basiertem Autodiscover verwendet.' 'Autodiscover Service Connection Point (SCP) — decisive for domain-joined clients on the internal network. Stored in AD (CN=Configuration) and preferred by Outlook over DNS-based autodiscover.')))
            $null = $parts.Add((New-WdTable -Headers @((L 'Client Access Server' 'Client Access server'), 'AutoDiscoverServiceInternalUri') -Rows $scpRows.ToArray()))
        }

        # DNS record template — pre-filled with accepted domain names, answers left blank for manual completion.
        # Automatic DNS resolution from the server is unreliable (internal DNS view differs from external; records
        # may not exist yet at installation time). External records are verified after go-live via mxtoolbox.com etc.
        $dnsTemplateRows = [System.Collections.Generic.List[object[]]]::new()
        $authDomainNames = @()
        if ($rd.Org -and $rd.Org.AcceptedDomains) {
            $authDomainNames = @($rd.Org.AcceptedDomains | Where-Object { $_.DomainType -eq 'Authoritative' } | Select-Object -ExpandProperty DomainName | Select-Object -First 5)
        }
        if (-not $authDomainNames -or $authDomainNames.Count -eq 0) { $authDomainNames = @('<domain>') }
        foreach ($d in $authDomainNames) {
            $d = [string]$d
            $dnsTemplateRows.Add(@('A / CNAME', "autodiscover.$d",             (L '(bitte manuell ergänzen)' '(please fill in manually)')))
            $dnsTemplateRows.Add(@('MX',        $d,                            (L '(bitte manuell ergänzen)' '(please fill in manually)')))
            $dnsTemplateRows.Add(@('TXT (SPF)',  $d,                            (L '(bitte manuell ergänzen)' '(please fill in manually)')))
            $dnsTemplateRows.Add(@('TXT (DKIM)', "selector1._domainkey.$d",    (L '(bitte manuell ergänzen)' '(please fill in manually)')))
            $dnsTemplateRows.Add(@('TXT (DMARC)',"_dmarc.$d",                 (L '(bitte manuell ergänzen)' '(please fill in manually)')))
        }
        $null = $parts.Add((New-WdParagraph (L 'Externe DNS-Einträge sind nach Go-Live über den autoritativen öffentlichen DNS zu prüfen (z. B. mxtoolbox.com, dig, nslookup). Die folgende Tabelle zeigt die typischerweise erforderlichen Einträge — bitte nach Einrichtung manuell ergänzen.' 'External DNS records must be verified after go-live via the authoritative public DNS (e.g. mxtoolbox.com, dig, nslookup). The table below lists the typically required records — please fill in after setup.')))
        $null = $parts.Add((New-WdTable -Headers @('Type', (L 'Name' 'Name'), (L 'Wert / Antwort' 'Value / Answer')) -Rows $dnsTemplateRows.ToArray()))

        # 6.2 Erforderliche Ports und Firewall-Regeln
        $null = $parts.Add((New-WdHeading (L '6.2 Erforderliche Ports und Firewall-Regeln' '6.2 Required Ports and Firewall Rules') 2))
        $null = $parts.Add((New-WdParagraph (L 'Die folgende Tabelle listet die für den Exchange Server-Betrieb erforderlichen TCP-Ports auf. Externe Ports müssen durch eine Firewall oder einen Reverse-Proxy abgesichert werden — Exchange Server sollte niemals direkt aus dem Internet erreichbar sein.' 'The table below lists the TCP ports required for Exchange Server operation. External ports must be secured by a firewall or reverse proxy — Exchange Server should never be directly reachable from the internet.')))
        $null = $parts.Add((New-WdTable -Headers @('Port', 'Protokoll', (L 'Dienst / Verwendung' 'Service / Purpose'), (L 'Sichtbarkeit' 'Visibility')) -Rows @(
            ,@('25',    'TCP', (L 'SMTP eingehend (extern + intern)' 'SMTP inbound (external + internal)'),                                               (L 'extern + intern' 'external + internal'))
            ,@('587',   'TCP', (L 'SMTP Submission / AUTH (Client-Einlieferung)' 'SMTP Submission / AUTH (client submission)'),                             (L 'intern / auth. Clients' 'internal / auth. clients'))
            ,@('443',   'TCP', (L 'HTTPS: OWA, ECP, EWS, Autodiscover, MAPI/HTTP, ActiveSync, OAB' 'HTTPS: OWA, ECP, EWS, Autodiscover, MAPI/HTTP, ActiveSync, OAB'), (L 'extern + intern' 'external + internal'))
            ,@('80',    'TCP', (L 'HTTP — Redirect auf HTTPS (am Reverse-Proxy)' 'HTTP — redirect to HTTPS (at reverse proxy)'),                           (L 'extern (Redirect)' 'external (redirect)'))
            ,@('993',   'TCP', (L 'IMAP4S (wenn aktiviert)' 'IMAP4S (if enabled)'),                                                                        (L 'intern / optional' 'internal / optional'))
            ,@('995',   'TCP', (L 'POP3S (wenn aktiviert)' 'POP3S (if enabled)'),                                                                          (L 'intern / optional' 'internal / optional'))
            ,@('135',   'TCP', (L 'RPC Endpoint Mapper (MAPI/RPC Legacy)' 'RPC Endpoint Mapper (MAPI/RPC legacy)'),                                        (L 'intern' 'internal'))
            ,@('445',   'TCP', (L 'SMB — DAG-Cluster, File Share Witness' 'SMB — DAG cluster, File Share Witness'),                                        (L 'intern (DAG)' 'internal (DAG)'))
            ,@('3268',  'TCP', (L 'Global Catalog LDAP' 'Global Catalog LDAP'),                                                                            (L 'intern (AD)' 'internal (AD)'))
            ,@('3269',  'TCP', (L 'Global Catalog LDAPS' 'Global Catalog LDAPS'),                                                                          (L 'intern (AD)' 'internal (AD)'))
            ,@('5985',  'TCP', (L 'WinRM HTTP (EMS, EXpress)' 'WinRM HTTP (EMS, EXpress)'),                                                               (L 'intern' 'internal'))
            ,@('5986',  'TCP', (L 'WinRM HTTPS (EMS, EXpress)' 'WinRM HTTPS (EMS, EXpress)'),                                                             (L 'intern' 'internal'))
            ,@('64327', 'TCP', (L 'DAG-Replikation (Mailbox Replication Service)' 'DAG Replication (Mailbox Replication Service)'),                        (L 'intern (DAG)' 'internal (DAG)'))
        ) -Compact))
        $null = $parts.Add((New-WdParagraph (L 'Hinweis: IMAP4 und POP3 sind auf Exchange Server standardmäßig deaktiviert und sollten nur bei explizitem Bedarf aktiviert werden. Port 80 (HTTP) sollte am Reverse-Proxy ausschließlich auf HTTPS (443) umgeleitet werden — Exchange-Dienste dürfen nicht unverschlüsselt exponiert sein.' 'Note: IMAP4 and POP3 are disabled by default on Exchange Server and should only be enabled when explicitly required. Port 80 (HTTP) should be redirected to HTTPS (443) at the reverse proxy — Exchange services must not be exposed unencrypted.')))

        # ── 7. Exchange-Installation (lokal, nur wenn kein Ad-hoc) ────────────────
        if (-not $isAdHoc) {
            $null = $parts.Add((New-WdHeading (L '7. Exchange-Installation (lokal)' '7. Exchange Installation (local)') 1))
            $null = $parts.Add((New-WdParagraph (L 'Die Exchange Server-Installation wurde mit EXpress vollautomatisch (Autopilot) bzw. interaktiv (Copilot) durchgeführt. EXpress übernimmt alle Installationsphasen 0–6 inklusive Windows-Features, .NET, VC++, URL Rewrite, UCMA, Active-Directory-Vorbereitung (PrepareSchema/PrepareAD), Exchange-Setup, Sicherheitshärtung und Post-Konfiguration. Die folgende Tabelle dokumentiert die installierte Exchange-Instanz auf diesem Server.' 'The Exchange Server installation was performed fully automated (Autopilot) or interactively (Copilot) using EXpress. EXpress handles all installation phases 0–6 including Windows features, .NET, VC++, URL Rewrite, UCMA, Active Directory preparation (PrepareSchema/PrepareAD), Exchange setup, security hardening and post-configuration. The table below documents the installed Exchange instance on this server.')))
            $exInstRows2 = [System.Collections.Generic.List[object[]]]::new()
            try {
                $exSrvLocal = Get-ExchangeServer $env:COMPUTERNAME -ErrorAction Stop
                $exInstRows2.Add(@((L 'Exchange-Version' 'Exchange version'), $exSrvLocal.AdminDisplayVersion.ToString()))
                $exInstRows2.Add(@((L 'Serverrolle' 'Server role'), ($exSrvLocal.ServerRole -join ', ')))
                $exInstRows2.Add(@((L 'Edition' 'Edition'), $exSrvLocal.Edition.ToString()))
                $exInstRows2.Add(@((L 'AD-Standort' 'AD site'), $exSrvLocal.Site.ToString()))
            } catch { }
            $exInstRows2.Add(@((L 'Installationspfad' 'Install path'), (SafeVal $State['InstallPath'])))
            $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $exInstRows2.ToArray()))

            # 7.1 Geplante Tasks (MEAC + Log Cleanup) — operative Exchange-Aufgaben, keine OS-Härtungen
            if ($rd.Org -and $rd.Org.ScheduledTasks -and $rd.Org.ScheduledTasks.Count -gt 0) {
                $null = $parts.Add((New-WdHeading (L '7.1 Geplante Tasks' '7.1 Scheduled Tasks') 2))
                $null = $parts.Add((New-WdParagraph (L 'EXpress registriert zwei operative geplante Aufgaben für den Exchange-Betrieb: MEAC (MonitorExchangeAuthCertificate.ps1) überwacht täglich das Exchange Auth-Zertifikat und erneuert es automatisch 60 Tage vor Ablauf — damit werden OAuth-/Hybrid-Ausfälle zuverlässig verhindert. EXpress Log-Cleanup bereinigt Exchange-Log-Verzeichnisse (Transport-Logs, IIS-Logs, HttpProxy-Logs, ETL/ETW) entsprechend der konfigurierten Aufbewahrungsfrist und verhindert ein Volllaufen des Log-Volumes.' 'EXpress registers two operational scheduled tasks for Exchange operations: MEAC (MonitorExchangeAuthCertificate.ps1) monitors the Exchange Auth certificate daily and automatically renews it 60 days before expiry — reliably preventing OAuth/Hybrid outages. EXpress Log-Cleanup purges Exchange log directories (transport logs, IIS logs, HttpProxy logs, ETL/ETW) according to the configured retention period, preventing the log volume from filling up.')))
                $stRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($t in $rd.Org.ScheduledTasks) {
                    $last = if ($t.LastRun)  { $t.LastRun.ToString('yyyy-MM-dd HH:mm')  } else { '—' }
                    $next = if ($t.NextRun)  { $t.NextRun.ToString('yyyy-MM-dd HH:mm')  } else { '—' }
                    $res  = if ($null -ne $t.LastResult) { ('0x{0:X}' -f $t.LastResult) } else { '—' }
                    $purpose =
                        if     ($t.Name -match 'Daily Auth Certificate|MonitorExchangeAuthCertificate|Monitor Exchange Auth') { (L 'Auto-Erneuerung Exchange Auth-Zertifikat (OAuth/Hybrid) — MEAC/CSS-Exchange' 'Auto-renewal of Exchange Auth certificate (OAuth/Hybrid) — MEAC/CSS-Exchange') }
                        elseif ($t.Name -match 'Log.?Cleanup|EXpressLogCleanup')                                            { (L 'Bereinigung Exchange-Log-Verzeichnisse' 'Cleanup of Exchange log directories') }
                        else                                                                           { '' }
                    $stRows.Add(@($t.Name, (SafeVal $t.Path), (SafeVal $t.State), $last, $next, $res, $purpose))
                }
                $null = $parts.Add((New-WdTable -Headers @((L 'Aufgabe' 'Task'), (L 'Pfad' 'Path'), (L 'Status' 'State'), (L 'Letzter Lauf' 'Last run'), (L 'Nächster Lauf' 'Next run'), (L 'Ergebnis' 'Result'), (L 'Zweck' 'Purpose')) -Rows $stRows.ToArray()))
            }

            # 7.2 Sicherheitsupdate-Stand
            $null = $parts.Add((New-WdHeading (L '7.2 Sicherheitsupdate-Stand' '7.2 Security Update Status') 2))
            $null = $parts.Add((New-WdParagraph (L 'Für Auditierbarkeit und Compliance ist der Patch-Stand des Exchange-Servers zu dokumentieren. Exchange Security Updates (SU) beheben kritische Sicherheitslücken (CVE) und müssen innerhalb der internen Patch-Window-Frist eingespielt werden. Neue SUs erscheinen monatlich (Patch Tuesday) oder außerplanmäßig bei kritischen Lücken. Der aktuelle Patch-Stand lässt sich über HealthChecker und Get-ExchangeDiagnosticInfo überprüfen.' 'For auditability and compliance, the patch status of the Exchange server must be documented. Exchange Security Updates (SU) fix critical vulnerabilities (CVE) and must be applied within the internal patch window. New SUs are released monthly (Patch Tuesday) or out-of-band for critical issues. The current patch status can be verified via HealthChecker and Get-ExchangeDiagnosticInfo.')))
            $suRows = [System.Collections.Generic.List[object[]]]::new()
            try {
                $exSrvSU = Get-ExchangeServer $env:COMPUTERNAME -ErrorAction Stop
                $suRows.Add(@((L 'Exchange-Version (Build)' 'Exchange version (build)'), $exSrvSU.AdminDisplayVersion.ToString()))
            } catch { }
            try {
                $osVer = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue)
                if ($osVer) {
                    $suRows.Add(@((L 'Windows-Version' 'Windows version'), ('{0} (Build {1})' -f $osVer.Caption, $osVer.BuildNumber)))
                    $suRows.Add(@((L 'Letzter Systemstart' 'Last system boot'), $osVer.LastBootUpTime.ToString('yyyy-MM-dd HH:mm:ss')))
                }
            } catch { }
            if ($State['ExchangeSUVersion']) { $suRows.Add(@((L 'Exchange SU (dieser Lauf)' 'Exchange SU (this run)'), (SafeVal $State['ExchangeSUVersion']))) }
            $suRows.Add(@((L 'Empfehlung' 'Recommendation'), (L 'HealthChecker nach jedem SU ausführen — https://aka.ms/ExchangeHealthChecker' 'Run HealthChecker after every SU — https://aka.ms/ExchangeHealthChecker')))
            $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $suRows.ToArray()))
        }

        # ── 8. Optimierungen und Härtungen (lokal) ────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '8. Optimierungen und Härtungen (lokaler Server)' '8. Optimisations and Hardening (local server)') 1))
        $null = $parts.Add((New-WdParagraph (L 'Im Rahmen der Installation wurden auf diesem Server gezielte Sicherheitshärtungen und Leistungsoptimierungen angewendet. Die Maßnahmen orientieren sich an den Empfehlungen des Microsoft Exchange-Teams, dem CIS Benchmark sowie Best Practices für Exchange Server in Unternehmensumgebungen. Die folgende Tabelle dokumentiert den aktuellen Konfigurationsstatus der wichtigsten Härtungsmaßnahmen.' 'As part of the installation, targeted security hardening measures and performance optimisations were applied to this server. The measures are based on the recommendations of the Microsoft Exchange team, the CIS Benchmark, and best practices for Exchange Server in enterprise environments. The following table documents the current configuration status of the most important hardening measures.')))

        $null = $parts.Add((New-WdHeading (L '8.1 TLS und Kryptografie' '8.1 TLS and Cryptography') 2))
        $null = $parts.Add((New-WdParagraph (L 'Exchange Server kommuniziert intern (MAPI, EWS, Autodiscover) und extern (SMTP, OWA, ActiveSync) ausschließlich über TLS-verschlüsselte Verbindungen. TLS 1.0 und 1.1 gelten als unsicher (POODLE, BEAST) und wurden deaktiviert. TLS 1.2 ist das Mindestprotokoll; TLS 1.3 wird auf Windows Server 2022/2025 zusätzlich aktiviert. Die .NET Strong Crypto-Einstellung stellt sicher, dass auch alle .NET-Anwendungen auf diesem Server ausschließlich sichere Cipher Suites verwenden.' 'Exchange Server communicates internally (MAPI, EWS, Autodiscover) and externally (SMTP, OWA, ActiveSync) exclusively over TLS-encrypted connections. TLS 1.0 and 1.1 are considered insecure (POODLE, BEAST) and have been disabled. TLS 1.2 is the minimum protocol; TLS 1.3 is additionally enabled on Windows Server 2022/2025. The .NET Strong Crypto setting ensures that all .NET applications on this server also use only secure cipher suites.')))
        $tlsRows = [System.Collections.Generic.List[object[]]]::new()
        # Helper: derive a semantic protocol state from Enabled + DisabledByDefault registry values.
        # Raw "Enabled=0" / "Disabled=1" values are ambiguous at a glance ("Disabled=0 means active?"),
        # so translate into plain text: Enabled / Disabled / OS-Default.
        function Get-TlsProtocolState([string]$proto, [bool]$shouldBeEnabled) {
            $base = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\{0}\Server' -f $proto
            $en  = Get-SecReg $base 'Enabled'
            $dbd = Get-SecReg $base 'DisabledByDefault'
            $effEnabled = $null
            if ($null -ne $en)        { $effEnabled = ([int]$en -ne 0) }
            elseif ($null -ne $dbd)   { $effEnabled = ([int]$dbd -eq 0) }
            if ($null -eq $effEnabled) {
                # Not present in registry → OS default. WS2022+/SE: TLS 1.0/1.1 disabled by default.
                $osDefEnabled = ($proto -in 'TLS 1.2','TLS 1.3')
                return if ($osDefEnabled) { (L 'aktiviert (OS-Standard)' 'enabled (OS default)') } else { (L 'deaktiviert (OS-Standard)' 'disabled (OS default)') }
            }
            $stateText = if ($effEnabled) { (L 'aktiviert' 'enabled') } else { (L 'deaktiviert' 'disabled') }
            $warn      = ''
            if ($shouldBeEnabled -and -not $effEnabled)    { $warn = (L ' — ACHTUNG: sollte aktiviert sein'   ' — WARNING: should be enabled') }
            if (-not $shouldBeEnabled -and $effEnabled)    { $warn = (L ' — ACHTUNG: sollte deaktiviert sein' ' — WARNING: should be disabled') }
            if ($warn) { '{0}{1}' -f $stateText, $warn } else { $stateText }
        }
        $tlsRows.Add(@('TLS 1.0 Server', (Get-TlsProtocolState 'TLS 1.0' $false)))
        $tlsRows.Add(@('TLS 1.1 Server', (Get-TlsProtocolState 'TLS 1.1' $false)))
        $tlsRows.Add(@('TLS 1.2 Server', (Get-TlsProtocolState 'TLS 1.2' $true)))
        $tlsRows.Add(@('TLS 1.3 Server', (Get-TlsProtocolState 'TLS 1.3' $true)))
        $tlsRows.Add(@('.NET Strong Crypto (v4)', (Format-RegBool (Get-SecReg 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319' 'SchUseStrongCrypto'))))
        $tlsRows.Add(@('.NET Strong Crypto (v2)', (Format-RegBool (Get-SecReg 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727' 'SchUseStrongCrypto'))))
        $null = $parts.Add((New-WdTable -Headers @((L 'Maßnahme' 'Measure'), (L 'Registrierungswert / Status' 'Registry value / status')) -Rows $tlsRows.ToArray()))

        $null = $parts.Add((New-WdHeading (L '8.2 Authentifizierung und Credential-Schutz' '8.2 Authentication and Credential Protection') 2))
        $null = $parts.Add((New-WdParagraph (L 'WDigest-Authentifizierung speichert Anmeldeinformationen im Klartextformat im LSASS-Speicher und ist für Pass-the-Hash- und Credential-Dumping-Angriffe (Mimikatz) anfällig. Sie wurde deaktiviert. LSA-Schutz (RunAsPPL) verhindert das Injizieren von unsigniertem Code in den LSASS-Prozess — ein zentraler Schutz gegen moderne Angriffswerkzeuge. Der LM-Kompatibilitätslevel bestimmt, welche Authentifizierungsprotokolle zugelassen werden; Level 5 (nur NTLMv2/Kerberos) entspricht dem aktuellen Sicherheitsstandard. Credential Guard (VBS) isoliert Credential-Hashes in einer virtualisierten Umgebung und ist auf Exchange-Servern zu deaktivieren, da Exchange interne Dienst-Konten mit NTLM-Authentifizierung nutzt.' 'WDigest authentication stores credentials in cleartext in LSASS memory and is vulnerable to pass-the-hash and credential dumping attacks (Mimikatz). It has been disabled. LSA protection (RunAsPPL) prevents injection of unsigned code into the LSASS process — a central protection against modern attack tools. The LM compatibility level determines which authentication protocols are permitted; level 5 (NTLMv2/Kerberos only) meets the current security standard. Credential Guard (VBS) isolates credential hashes in a virtualised environment and must be disabled on Exchange servers, as Exchange uses internal service accounts with NTLM authentication.')))
        $authRows = [System.Collections.Generic.List[object[]]]::new()
        $authRows.Add(@('WDigest UseLogonCredential', (Format-RegBool (Get-SecReg 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest' 'UseLogonCredential'))))
        $authRows.Add(@('LSA RunAsPPL',               (Format-RegBool (Get-SecReg 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' 'RunAsPPL'))))
        $lmLevel = Get-SecReg 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' 'LmCompatibilityLevel'
        $lmText  = if ($null -eq $lmLevel) { (L 'nicht gesetzt (Standard: 3)' 'not set (default: 3)') } else { 'Level {0}' -f $lmLevel }
        $authRows.Add(@('LM Compatibility Level', $lmText))
        $authRows.Add(@('Credential Guard (VBS)',  (Format-RegBool (Get-SecReg 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard' 'EnableVirtualizationBasedSecurity'))))
        $null = $parts.Add((New-WdTable -Headers @((L 'Maßnahme' 'Measure'), (L 'Registrierungswert / Status' 'Registry value / status')) -Rows $authRows.ToArray()))

        $null = $parts.Add((New-WdHeading (L '8.3 Netzwerkprotokolle' '8.3 Network Protocols') 2))
        $null = $parts.Add((New-WdParagraph (L 'SMBv1 ist ein veraltetes Dateifreigabeprotokoll ohne Verschlüsselung, das für WannaCry, NotPetya und ähnliche Ransomware-Angriffe genutzt wurde. Es wurde vollständig deaktiviert. HTTP/2 für Exchange-Webdienste wird deaktiviert, da es mit bestimmten Load-Balancer-Konfigurationen und dem Exchange Extended Protection-Mechanismus interferiert. SSL-Offloading (Beendigung der TLS-Verbindung am Load Balancer) ist deaktiviert, da Extended Protection eine End-to-End-TLS-Bindung erfordert.' 'SMBv1 is an outdated file-sharing protocol without encryption that was exploited by WannaCry, NotPetya and similar ransomware attacks. It has been completely disabled. HTTP/2 for Exchange web services is disabled as it interferes with certain load balancer configurations and the Exchange Extended Protection mechanism. SSL offloading (terminating the TLS connection at the load balancer) is disabled because Extended Protection requires end-to-end TLS binding.')))
        $protoRows = [System.Collections.Generic.List[object[]]]::new()
        try { $smb1 = (Get-SmbServerConfiguration -ErrorAction Stop).EnableSMB1Protocol; $protoRows.Add(@('SMBv1', (Format-RegBool $smb1))) } catch { }
        $protoRows.Add(@('HTTP/2 Cleartext (Exchange FE)', (Format-RegBool (Get-SecReg 'HKLM:\SYSTEM\CurrentControlSet\Services\HTTP\Parameters' 'EnableHttp2Cleartext'))))
        $null = $parts.Add((New-WdTable -Headers @((L 'Maßnahme' 'Measure'), (L 'Registrierungswert / Status' 'Registry value / status')) -Rows $protoRows.ToArray()))

        $null = $parts.Add((New-WdHeading (L '8.4 Exchange-spezifische Härtung' '8.4 Exchange-specific Hardening') 2))
        $null = $parts.Add((New-WdParagraph (L 'Extended Protection (EPA) verhindert Man-in-the-Middle-Angriffe auf HTTP-Verbindungen, indem die TLS-Channel-Binding-Information in die Authentifizierung einbezogen wird. Serialized Data Signing (SDS) schützt vor Deserialisierungsangriffen auf Exchange-interne Kommunikation. AMSI-Body-Scanning prüft HTTP-Anfragen (OWA, ECP, EWS, PowerShell) auf bekannte Angriffsmuster durch die Windows Defender AMSI-Schnittstelle. Die MAPI-Verschlüsselung stellt sicher, dass Outlook-MAPI-Verbindungen ausschließlich verschlüsselt erfolgen. Strict Mode für Powershell-Remoting und die Deaktivierung der PowerShell Autodiscover-App-Pools senken die Angriffsfläche der Exchange-Management-Schnittstellen weiter.' 'Extended Protection (EPA) prevents man-in-the-middle attacks on HTTP connections by incorporating TLS channel binding information into authentication. Serialized Data Signing (SDS) protects against deserialization attacks on Exchange internal communication. AMSI body scanning checks HTTP requests (OWA, ECP, EWS, PowerShell) for known attack patterns via the Windows Defender AMSI interface. MAPI encryption ensures that Outlook MAPI connections are exclusively encrypted. Strict mode for PowerShell remoting and disabling the PowerShell Autodiscover app pools further reduce the attack surface of the Exchange management interfaces.')))
        $exHardRows = [System.Collections.Generic.List[object[]]]::new()
        # Pull authoritative values from Exchange where available; fall back to registry-only hints otherwise.
        $epaState = '(unknown)'
        try {
            $epAuthDirs = @(Get-ExchangeServer $env:COMPUTERNAME -ErrorAction Stop | Out-Null)  # ensure EMS available
            $vdAuth = @()
            try { Get-OwaVirtualDirectory -Server $env:COMPUTERNAME -ErrorAction Stop       | ForEach-Object { $vdAuth += ('OWA={0}'        -f $_.ExtendedProtectionTokenChecking) } } catch { }
            try { Get-EcpVirtualDirectory -Server $env:COMPUTERNAME -ErrorAction Stop       | ForEach-Object { $vdAuth += ('ECP={0}'        -f $_.ExtendedProtectionTokenChecking) } } catch { }
            try { Get-WebServicesVirtualDirectory -Server $env:COMPUTERNAME -ErrorAction Stop | ForEach-Object { $vdAuth += ('EWS={0}'       -f $_.ExtendedProtectionTokenChecking) } } catch { }
            try { Get-OabVirtualDirectory -Server $env:COMPUTERNAME -ErrorAction Stop       | ForEach-Object { $vdAuth += ('OAB={0}'        -f $_.ExtendedProtectionTokenChecking) } } catch { }
            try { Get-ActiveSyncVirtualDirectory -Server $env:COMPUTERNAME -ErrorAction Stop | ForEach-Object { $vdAuth += ('EAS={0}'       -f $_.ExtendedProtectionTokenChecking) } } catch { }
            try { Get-MapiVirtualDirectory -Server $env:COMPUTERNAME -ErrorAction Stop      | ForEach-Object { $vdAuth += ('MAPI={0}'       -f $_.ExtendedProtectionTokenChecking) } } catch { }
            try { Get-AutodiscoverVirtualDirectory -Server $env:COMPUTERNAME -ErrorAction Stop | ForEach-Object { $vdAuth += ('Autodiscover={0}' -f $_.ExtendedProtectionTokenChecking) } } catch { }
            if ($vdAuth.Count -gt 0) { $epaState = ($vdAuth -join ', ') }
        } catch { }
        $exHardRows.Add(@('Extended Protection (EPA)', $epaState, (L 'Channel-Binding-Schutz gegen MITM auf IIS-VDirs' 'Channel-binding protection against MITM on IIS VDirs')))
        # Registry value name is EnableSerializationDataSigning (Microsoft's actual spelling), not EnableSerializedDataSigning.
        $exHardRows.Add(@('Serialized Data Signing', (Format-RegBool (Get-SecReg 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Diagnostics' 'EnableSerializationDataSigning')), (L 'Schutz gegen Deserialisierungs-Angriffe (ab März 2024 verpflichtend)' 'Protection against deserialization attacks (required since March 2024)')))
        $amsiVal  = Get-SecReg 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Diagnostics' 'DisableAMSIScanning'
        $amsiText = if ($null -eq $amsiVal) { (L 'aktiviert (Standard)' 'enabled (default)') } elseif (([int]"$amsiVal") -eq 0) { (L 'aktiviert' 'enabled') } else { (L 'deaktiviert' 'disabled') }
        $exHardRows.Add(@('AMSI Body Scanning', $amsiText, (L 'HTTP-Request-Scan über Windows Defender AMSI' 'HTTP request scan via Windows Defender AMSI')))
        $mapiEnc = try { (Get-RpcClientAccess -Server $env:COMPUTERNAME -ErrorAction Stop | Select-Object -First 1).EncryptionRequired.ToString() } catch { '(unknown)' }
        $exHardRows.Add(@('MAPI Encryption Required', (SafeVal $mapiEnc), (L 'Outlook-/MAPI-Verbindungen nur verschlüsselt' 'Outlook/MAPI connections encrypted only')))
        # Throttling / rate limiting for Exchange Web Services (mitigates abuse / DoS on EWS endpoint)
        $throt = try {
            $tp = Get-ThrottlingPolicy -ErrorAction Stop | Where-Object { $_.IsDefault } | Select-Object -First 1
            if ($tp -and $null -ne $tp.EwsMaxConcurrency) { $tp.EwsMaxConcurrency.ToString() }
            else { (L '(nicht gesetzt — Standard: 27)' '(not set — default: 27)') }
        } catch { (L '(nicht abrufbar)' '(not available)') }
        $exHardRows.Add(@('EWS Max Concurrency (default policy)', $throt, (L 'Throttling-Policy gegen EWS-Überlastung' 'Throttling policy against EWS overload')))
        # Authentication flags on OWA/ECP
        $owaBasic = try { (Get-OwaVirtualDirectory -Server $env:COMPUTERNAME -ErrorAction Stop | Select-Object -First 1).BasicAuthentication.ToString() } catch { '(unknown)' }
        $exHardRows.Add(@('OWA Basic Authentication', (SafeVal $owaBasic), (L 'Basic-Auth auf OWA ist gegen Credential-Stuffing anfällig' 'Basic auth on OWA is vulnerable to credential stuffing')))
        # PowerShell Autodiscover app pool (F19: disabled by EXpress; mitigates ProxyLogon-style vectors)
        $psPool = try { (Get-Website | Where-Object { $_.Name -eq 'Default Web Site' } | Out-Null); (Get-WebAppPoolState -Name 'MSExchangePowerShellAppPool' -ErrorAction Stop).Value } catch { '(unknown)' }
        $exHardRows.Add(@('MSExchangePowerShellAppPool', (SafeVal $psPool), (L 'Remote-PowerShell-Pool — Started/Stopped' 'Remote PowerShell pool — Started/Stopped')))
        $autodiscPool = try { (Get-WebAppPoolState -Name 'MSExchangeAutodiscoverAppPool' -ErrorAction Stop).Value } catch { (L '(nicht abrufbar)' '(not available)') }
        $exHardRows.Add(@('MSExchangeAutodiscoverAppPool', (SafeVal $autodiscPool), (L 'Autodiscover PowerShell-AppPool — aktueller Status' 'Autodiscover PowerShell app pool — current state')))
        # ECC certificate support (cipher modernization)
        $exHardRows.Add(@('ECC Certificate Support', (Format-RegBool (Get-SecReg 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Diagnostics' 'EnableEccCertificateSupport')), (L 'Moderne ECC-Zertifikate in Exchange zugelassen' 'Modern ECC certificates permitted in Exchange')))
        # Setup-Override files (SettingOverride framework) — CVE-bezogene Kill-Switches
        try {
            $overrides = @(Get-ExchangeSettingOverride -ErrorAction Stop) 2>$null
            if ($overrides) {
                $ovList = ($overrides | ForEach-Object { '{0}:{1}={2}' -f $_.ComponentName, $_.SectionName, ($_.Parameters -join ',') }) -join '; '
                $exHardRows.Add(@('Exchange SettingOverrides', (SafeVal $ovList), (L 'Aktive Konfigurations-Overrides (CVE-Mitigationen, Features)' 'Active configuration overrides (CVE mitigations, features)')))
            }
        } catch { }
        $null = $parts.Add((New-WdTable -Headers @((L 'Härtungsmaßnahme' 'Hardening measure'), (L 'Status / Wert' 'Status / value'), (L 'Zweck' 'Purpose')) -Rows $exHardRows.ToArray()))

        # 8.5 Windows Defender Exclusions
        $localSrvData = @($rd.Servers | Where-Object { $_.IsLocalServer }) | Select-Object -First 1
        if ($localSrvData -and $localSrvData.DefenderExclusions) {
            $null = $parts.Add((New-WdHeading (L '8.5 Windows Defender — Ausnahmen' '8.5 Windows Defender — Exclusions') 2))
            $null = $parts.Add((New-WdParagraph (L 'Microsoft dokumentiert umfangreiche Pfad-, Prozess- und Dateityp-Ausnahmen für Exchange Server, ohne die Antivirus-Software Datenbank-Dateien, Transport-Warteschlangen oder Logs blockiert und Leistung wie Stabilität schwer beeinträchtigt. EXpress trägt diese Ausnahmen automatisch in Windows Defender ein. Bei Drittanbieter-Antivirus müssen dieselben Pfade manuell in das entsprechende Produkt übernommen werden. Weitere Informationen: Microsoft Docs "Exchange antivirus software".' 'Microsoft documents extensive path, process and filetype exclusions for Exchange Server without which antivirus software would block database files, transport queues or logs and severely impact performance and stability. EXpress automatically registers these exclusions with Windows Defender. For third-party antivirus, the same paths must be manually configured in the corresponding product. Further information: Microsoft Docs "Exchange antivirus software".')))
            $exr = $localSrvData.DefenderExclusions
            $defRows = [System.Collections.Generic.List[object[]]]::new()
            $defRows.Add(@((L 'Echtzeit-Überwachung' 'Real-time monitoring'), (Lc $exr.RealTimeEnabled (L 'aktiv' 'enabled') (L 'inaktiv' 'disabled'))))
            $defRows.Add(@((L 'Pfad-Ausnahmen' 'Path exclusions'), (SafeVal (($exr.ExclusionPath | Sort-Object) -join "`n") (L '(keine)' '(none)'))))
            $defRows.Add(@((L 'Prozess-Ausnahmen' 'Process exclusions'), (SafeVal (($exr.ExclusionProcess | Sort-Object) -join "`n") (L '(keine)' '(none)'))))
            $defRows.Add(@((L 'Dateityp-Ausnahmen' 'Extension exclusions'), (SafeVal (($exr.ExclusionExtension | Sort-Object) -join "`n") (L '(keine)' '(none)'))))
            $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $defRows.ToArray()))
        }

        # 8.6 IIS- und Exchange-Logs
        $null = $parts.Add((New-WdHeading (L '8.6 Protokollierung — IIS und Exchange' '8.6 Logging — IIS and Exchange') 2))
        $null = $parts.Add((New-WdParagraph (L 'Exchange Server schreibt umfangreiche Betriebsprotokolle in den Logging-Pfad unter dem Exchange-Installationsverzeichnis (Transport, Managed Availability, HttpProxy, CAS). IIS protokolliert Zugriffe auf OWA, ECP, EWS, ActiveSync, MAPI, OAB. Ohne automatische Bereinigung füllen diese Logs innerhalb weniger Wochen das Log-Volume vollständig auf. EXpress registriert hierfür einen geplanten Task (siehe 7.1), der Logs älter als der konfigurierten Aufbewahrungsfrist (Standard: 30 Tage) automatisch entfernt. Die tatsächlichen IIS-Log-Pfade zeigt die folgende Tabelle.' 'Exchange Server writes extensive operational logs to the logging path below the Exchange installation directory (Transport, Managed Availability, HttpProxy, CAS). IIS logs access to OWA, ECP, EWS, ActiveSync, MAPI, OAB. Without automatic cleanup these logs fill the log volume completely within a few weeks. EXpress registers a scheduled task for this purpose (see 7.1) which automatically removes logs older than the configured retention (default: 30 days). Actual IIS log paths are shown in the table below.')))
        $null = $parts.Add((New-WdParagraph (L 'Hinweis zu Forensik und Compliance: Die regelmäßige lokale Bereinigung dient ausschließlich dazu, ein Vollaufen des Log-Volumes (und damit den Ausfall von Transport, IIS und Managed Availability) zu verhindern — sie ist kein Ersatz für eine revisionssichere Langzeit-Aufbewahrung. Für forensische Auswertung sicherheitsrelevanter Vorfälle (Authentifizierungs-Anomalien, EWS-/MAPI-Zugriffsmuster, Transport-Spuren bei Datenabfluss) und zur Erfüllung gesetzlicher Aufbewahrungspflichten (BSI APP.5.2, DSGVO Rechenschaftspflicht, GoBD) sind IIS-, HttpProxy-, MessageTracking-, Transport- und Windows-Security-Eventlogs idealerweise per Log-Forwarder (z. B. NXLog, WEF/WEC, Filebeat, Azure Monitor Agent) an ein zentrales SIEM (z. B. Splunk, Elastic Security, Microsoft Sentinel, Wazuh, IBM QRadar) auszuleiten. Die Aufbewahrungsdauer im SIEM sollte sich an der internen Sicherheitsleitlinie und branchenspezifischen Vorgaben orientieren (typisch 12 Monate Hot-Storage, 7 Jahre Archiv). Erst diese Kombination — kurze Aufbewahrung am Server, lange Aufbewahrung im SIEM — erfüllt sowohl operative Stabilitätsanforderungen als auch forensische und Compliance-Anforderungen.' 'Note on forensics and compliance: Periodic local cleanup is intended solely to prevent the log volume from filling up (which would take down Transport, IIS and Managed Availability) — it is not a substitute for tamper-evident long-term retention. For forensic investigation of security-relevant incidents (authentication anomalies, EWS/MAPI access patterns, transport traces during data exfiltration) and to meet legal retention obligations (BSI APP.5.2, GDPR accountability, GoBD), IIS, HttpProxy, MessageTracking, Transport and Windows Security event logs should ideally be forwarded via a log shipper (e.g. NXLog, WEF/WEC, Filebeat, Azure Monitor Agent) to a central SIEM (e.g. Splunk, Elastic Security, Microsoft Sentinel, Wazuh, IBM QRadar). Retention in the SIEM should follow the organisation''s security policy and industry-specific requirements (typically 12 months hot storage, 7 years archive). Only this combination — short retention on the server, long retention in the SIEM — satisfies both operational stability and forensic/compliance requirements.')))
        $logRows = [System.Collections.Generic.List[object[]]]::new()
        $logRows.Add(@((L 'Exchange Logging-Pfad' 'Exchange logging path'), (SafeVal (Join-Path (Split-Path $env:ExchangeInstallPath -Parent) 'Logging'))))
        $logRows.Add(@((L 'ETL/Diagnostic-Pfad' 'ETL/Diagnostic path'), (SafeVal (Join-Path (Split-Path $env:ExchangeInstallPath -Parent) 'Bin\Search\Ceres\HostController\Data\Events'))))
        $retDays = if ($State['LogRetentionDays']) { $State['LogRetentionDays'] } else { 30 }
        $logRows.Add(@((L 'Aufbewahrung (EXpress Log Cleanup)' 'Retention (EXpress log cleanup)'), ('{0} {1}' -f $retDays, (L 'Tage' 'days'))))
        if ($localSrvData -and $localSrvData.IISLogs) {
            foreach ($site in $localSrvData.IISLogs.Sites) {
                $logRows.Add(@(('IIS: {0}' -f $site.Name), ('{0} — Format: {1} — Period: {2}' -f (SafeVal $site.LogDir), (SafeVal $site.LogFormat), (SafeVal $site.Period))))
            }
        }
        $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $logRows.ToArray()))

        # 8.7 Kerberos Load Balancing
        $null = $parts.Add((New-WdHeading (L '8.7 Kerberos Load Balancing' '8.7 Kerberos Load Balancing') 2))
        $null = $parts.Add((New-WdParagraph (L 'In Umgebungen mit Hardware- oder Software-Load-Balancern (NLB, F5, Kemp, HAProxy u. a.) ist Kerberos-Authentifizierung für Exchange-Dienste ohne Session Affinity möglich, sofern ein dedizierter Kerberos-Service-Account (KSA) konfiguriert wird. Ohne KSA fällt Kerberos auf NTLM zurück, wenn der Client an einen anderen Server weitergeleitet wird als den, den er ursprünglich kontaktiert hat — NTLM erzeugt höhere Latenz und kann in großen Umgebungen zu NtLM-Stau führen. Der KSA erhält einen Service Principal Name (SPN) für jeden HTTPS-Dienst (OWA, EWS, Autodiscover, MAPI, ECP, ActiveSync, OAB) und wird in AD als Konto mit gesetztem "Kerberos-Einschränkungen zulassen" hinterlegt. Ab Exchange 2016 mit CAS-Array-Entfall ist Kerberos-LB eine optionale, aber empfehlenswerte Konfiguration für Umgebungen mit mehreren Exchange-Servern hinter einem LB.' 'In environments with hardware or software load balancers (NLB, F5, Kemp, HAProxy etc.), Kerberos authentication for Exchange services without session affinity is possible provided a dedicated Kerberos Service Account (KSA) is configured. Without a KSA, Kerberos falls back to NTLM when a client is redirected to a different server than the one it originally contacted — NTLM causes higher latency and can lead to NTLM saturation in large environments. The KSA is assigned a Service Principal Name (SPN) for each HTTPS service (OWA, EWS, Autodiscover, MAPI, ECP, ActiveSync, OAB) and registered in AD as an account with "Constrain Kerberos delegation" set. Since Exchange 2016 with the removal of the CAS array, Kerberos LB is an optional but recommended configuration for environments with multiple Exchange servers behind a load balancer.')))
        $krbRows = [System.Collections.Generic.List[object[]]]::new()
        try {
            $cas = @(Get-ClientAccessService -ErrorAction Stop)
            foreach ($c in $cas) {
                $ksa = try { $c.AlternateServiceAccountCredential | Select-Object -First 1 } catch { $null }
                $ksaName = if ($ksa -and $ksa.Credential) { $ksa.Credential.UserName } elseif ($c.AlternateServiceAccountCredential) { (SafeVal ($c.AlternateServiceAccountCredential -join ', ')) } else { (L '(kein KSA konfiguriert)' '(no KSA configured)') }
                $krbRows.Add(@($c.Name, $ksaName, (SafeVal $c.AutoDiscoverServiceInternalUri)))
            }
        } catch { }
        if ($krbRows.Count -eq 0) { $krbRows.Add(@((L '(Get-ClientAccessService nicht verfügbar)' '(Get-ClientAccessService not available)'), '', '')) }
        $null = $parts.Add((New-WdTable -Headers @((L 'CAS-Server' 'CAS server'), (L 'Kerberos-Service-Account' 'Kerberos service account'), 'Autodiscover URI') -Rows $krbRows.ToArray()))
        $null = $parts.Add((New-WdParagraph (L 'Konfigurationsreferenz: Set-ClientAccessService -AlternateServiceAccountCredential. Weitere Details und SPN-Registrierung: Microsoft Docs "Configure Kerberos authentication for load-balanced Exchange servers".' 'Configuration reference: Set-ClientAccessService -AlternateServiceAccountCredential. Further details and SPN registration: Microsoft Docs "Configure Kerberos authentication for load-balanced Exchange servers".')))

        # 8.8 Compliance-Mapping CIS / BSI IT-Grundschutz
        $null = $parts.Add((New-WdHeading (L '8.8 Compliance-Mapping (CIS / BSI IT-Grundschutz)' '8.8 Compliance Mapping (CIS / BSI)') 2))
        $null = $parts.Add((New-WdParagraph (L 'Die folgende Tabelle ordnet die von EXpress angewendeten Härtungsmaßnahmen den relevanten Kontrollen aus dem CIS Benchmark for Microsoft Windows Server und dem BSI IT-Grundschutz-Kompendium zu. Sie dient als Nachweis für Audits und interne Compliance-Prüfungen.' 'The table below maps the hardening measures applied by EXpress to the relevant controls from the CIS Benchmark for Microsoft Windows Server and the BSI IT-Grundschutz Compendium. It serves as evidence for audits and internal compliance reviews.')))
        $null = $parts.Add((New-WdParagraph (L 'Wichtiger Hinweis zur Protokoll-Auswertung: Mehrere der nachfolgenden Kontrollen — insbesondere Admin Audit Log, Mailbox Audit Log, Windows Security Eventlog und IIS-Zugriffsprotokolle — entfalten ihren vollen Compliance- und forensischen Nutzen erst, wenn die erzeugten Ereignisse zentral zusammengeführt, korreliert und revisionssicher aufbewahrt werden. EXpress aktiviert und konfiguriert die Protokollquellen auf dem Server, sieht jedoch ausdrücklich keine SIEM-Anbindung vor — diese ist organisationsweit zu planen und liegt außerhalb des Scopes einer Server-Installation. Für die Erfüllung von BSI APP.5.2 A13 (Protokollierung), BSI OPS.1.1.5 (Protokollierung), CIS Control 8 (Audit Log Management) sowie der DSGVO-Rechenschaftspflicht (Art. 5 Abs. 2) ist die Anbindung an ein SIEM (Security Information and Event Management) dringend empfohlen. Ein SIEM ermöglicht: (1) zentrale Korrelation über mehrere Exchange-Server, Domain Controller und Edge-Komponenten hinweg; (2) Alarmierung bei Anomalien (Brute-Force-Versuche, ungewöhnliche EWS-/PowerShell-Zugriffe, Mass-Mail-Abfluss); (3) revisionssichere Langzeit-Aufbewahrung über die lokale Bereinigungsfrist hinaus; (4) Nachweisführung gegenüber Auditoren ohne Eingriff am Produktivsystem. Empfohlene Quellen für die Auslieferung: Windows Security/System/Application-Eventlog, IIS-W3C-Logs, Exchange MessageTracking, HttpProxy, Managed Availability, sowie das Admin- und Mailbox-Audit-Log via Search-AdminAuditLog / Search-MailboxAuditLog oder New-MailboxAuditLogSearch.' 'Important note on log evaluation: Several of the controls below — in particular Admin Audit Log, Mailbox Audit Log, Windows Security event log and IIS access logs — only deliver their full compliance and forensic value when the generated events are centrally aggregated, correlated and retained tamper-evidently. EXpress enables and configures the log sources on the server, but explicitly does not provide SIEM integration — this must be planned organisation-wide and is out of scope for a server installation. To meet BSI APP.5.2 A13 (logging), BSI OPS.1.1.5 (logging), CIS Control 8 (Audit Log Management) and the GDPR accountability obligation (Art. 5(2)), integration with a SIEM (Security Information and Event Management) is strongly recommended. A SIEM enables: (1) central correlation across multiple Exchange servers, domain controllers and edge components; (2) alerting on anomalies (brute-force attempts, unusual EWS/PowerShell access, mass mail exfiltration); (3) tamper-evident long-term retention beyond the local cleanup period; (4) audit evidence without touching the production system. Recommended sources for forwarding: Windows Security/System/Application event log, IIS W3C logs, Exchange MessageTracking, HttpProxy, Managed Availability, plus the Admin and Mailbox Audit Log via Search-AdminAuditLog / Search-MailboxAuditLog or New-MailboxAuditLogSearch.')))
        $null = $parts.Add((New-WdTable -Headers @((L 'Maßnahme' 'Measure'), (L 'CIS-Kontrolle' 'CIS Control'), (L 'BSI-Grundschutz' 'BSI Control'), (L 'Status' 'Status')) -Rows @(
            ,@((L 'TLS 1.0 / 1.1 deaktiviert' 'TLS 1.0 / 1.1 disabled'),                          'CIS WS2022 18.4.x',   'BSI SYS.1.2 A5',  (L 'Umgesetzt' 'Implemented'))
            ,@((L 'TLS 1.2 erzwungen + .NET Strong Crypto' 'TLS 1.2 enforced + .NET Strong Crypto'), 'CIS WS2022 18.4.x',   'BSI SYS.1.2 A5',  (L 'Umgesetzt' 'Implemented'))
            ,@('RC4 / 3DES / NULL Ciphers deaktiviert',                                              'CIS WS2022 2.3.11.x', 'BSI SYS.1.2 A6',  (L 'Umgesetzt' 'Implemented'))
            ,@((L 'SMBv1 deaktiviert' 'SMBv1 disabled'),                                             'CIS WS2022 18.3.4',   'BSI NET.3.4 A2',  (L 'Umgesetzt' 'Implemented'))
            ,@('NTLMv2 (LmCompatibilityLevel = 5)',                                                   'CIS WS2022 2.3.11.8', 'BSI SYS.1.2 A7',  (L 'Umgesetzt' 'Implemented'))
            ,@((L 'WDigest deaktiviert' 'WDigest disabled'),                                         'CIS WS2022 18.3.7',   'BSI SYS.1.6 A3',  (L 'Umgesetzt' 'Implemented'))
            ,@((L 'LSA-Schutz aktiviert' 'LSA Protection enabled'),                                  'CIS WS2022 18.4.5',   'BSI SYS.1.6 A5',  (L 'Umgesetzt' 'Implemented'))
            ,@('Extended Protection for Authentication (EPA)',                                         'CIS WS2022 18.4.x',   'BSI APP.5.2 A10', (L 'Umgesetzt' 'Implemented'))
            ,@('Serialized Data Signing',                                                              'MS Exchange SE Baseline', 'BSI APP.5.2 A10', (L 'Umgesetzt' 'Implemented'))
            ,@((L 'Defender Ausnahmen (Exchange-VSS, Transport, IIS)' 'Defender exclusions (Exchange VSS, Transport, IIS)'), 'MS Exchange Best Practice', 'BSI APP.5.2 A4', (L 'Umgesetzt' 'Implemented'))
            ,@('LLMNR / mDNS deaktiviert',                                                            'CIS WS2022 18.5.4.2', 'BSI NET.3.1 A10', (L 'Umgesetzt' 'Implemented'))
            ,@((L 'Dienste minimiert (Browser/Fax/Xcopy u. a.)' 'Services minimised (Browser/Fax/Xcopy etc.)'), 'CIS WS2022 5.x', 'BSI SYS.1.2 A3', (L 'Umgesetzt' 'Implemented'))
            ,@((L 'Admin Audit Log aktiviert' 'Admin Audit Log enabled'),                             'CIS EX2019 1.1',      'BSI APP.5.2 A13', (L 'Umgesetzt' 'Implemented'))
            ,@((L 'SIEM-Anbindung / zentrale Log-Auswertung' 'SIEM integration / central log evaluation'), 'CIS Control 8',     'BSI OPS.1.1.5 / APP.5.2 A13', (L 'Out of Scope — organisationsweit zu planen' 'Out of scope — to be planned organisation-wide'))
            ,@((L 'Log-Bereinigung am Server (Volume-Schutz)' 'Local log cleanup (volume protection)'), 'MS Best Practice',     'BSI APP.5.2 A4',  (L 'Umgesetzt — geplante Aufgabe (siehe 7.1)' 'Implemented — scheduled task (see 7.1)'))
            ,@((L 'Defender Echtzeit deaktiviert (Exchange-Konflikt mit AWL)' 'Defender realtime disabled (Exchange AWL conflict)'), 'CIS WS2022 n/a', 'BSI SYS.1.2 A4', (L 'Ausnahme — Exchange-AWL-Konflikt; AV-Ausnahmen gesetzt' 'Exception — Exchange AWL conflict; AV exclusions applied'))
        ) -Compact))

        # 8.9 Datenschutz und DSGVO-Relevanz
        $null = $parts.Add((New-WdHeading (L '8.9 Datenschutz und DSGVO-Relevanz' '8.9 Data Protection and GDPR Relevance') 2))
        $null = $parts.Add((New-WdParagraph (L 'Exchange Server verarbeitet personenbezogene Daten (E-Mail-Inhalte, Adressdaten, Kalendereinträge, Postfachberechtigungen) und ist daher für Organisationen in der EU als Datenverarbeitungssystem im Sinne der DSGVO (Art. 4 Nr. 2) einzustufen. Die folgende Checkliste fasst die datenschutzrelevanten Aspekte zusammen.' 'Exchange Server processes personal data (email content, address data, calendar entries, mailbox permissions) and must therefore be classified as a data processing system under the GDPR (Art. 4 No. 2) for organisations in the EU. The checklist below summarises the data protection-relevant aspects.')))
        $null = $parts.Add((New-WdTable -Headers @((L 'Datenschutzaspekt' 'Data protection aspect'), (L 'Status / Hinweis' 'Status / Note')) -Rows @(
            ,@((L 'Transportverschlüsselung (TLS 1.2+)' 'Transport encryption (TLS 1.2+)'),                          (L 'Umgesetzt — TLS 1.2 auf allen Verbindungspunkten erzwungen' 'Implemented — TLS 1.2 enforced on all connection points'))
            ,@((L 'Ruheverschlüsselung (Encryption at rest)' 'Encryption at rest'),                                   (L 'BitLocker (OS-Ebene) empfohlen; Exchange-native DB-Verschlüsselung nicht verfügbar' 'BitLocker (OS level) recommended; Exchange-native DB encryption not available'))
            ,@((L 'Admin-Auditprotokoll' 'Admin Audit Log'),                                                           (L 'Umgesetzt — administrative Cmdlet-Ausführungen werden protokolliert' 'Implemented — administrative cmdlet executions are logged'))
            ,@((L 'Postfach-Zugriffsprotokoll (Mailbox Audit Logging)' 'Mailbox Audit Logging'),                      (L 'Ab Exchange 2019: standardmäßig aktiviert (Default Audit Logging)' 'From Exchange 2019: enabled by default (Default Audit Logging)'))
            ,@((L 'Aufbewahrungsrichtlinien / Löschfristen' 'Retention policies / deletion periods'),                  (L 'Über Compliance-Tags und Retention Policies im Compliance Center konfigurieren' 'Configure via Compliance Tags and Retention Policies in the Compliance Center'))
            ,@((L 'Verarbeitungsverzeichnis (Art. 30 DSGVO)' 'Records of processing activities (Art. 30 GDPR)'),      (L 'Exchange Server ist im Verarbeitungsverzeichnis zu führen' 'Exchange Server must be included in the records of processing activities'))
            ,@((L 'DSFA / DPIA (Art. 35 DSGVO)' 'DPIA (Art. 35 GDPR)'),                                              (L 'Bei umfangreicher Verarbeitung sensibler Daten ggf. erforderlich' 'May be required for extensive processing of sensitive data'))
            ,@((L 'Auftragsverarbeitung (AV-Vertrag / DPA)' 'Data Processing Agreement (DPA)'),                       (L 'Mit M365/EOP/AIP-Diensten ist ein AV-Vertrag (Microsoft DPA) abzuschließen' 'A DPA (Microsoft DPA) must be concluded for M365/EOP/AIP services'))
        )))

        # ── 9. Anti-Spam / Agents (lokal) ─────────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '9. Transport-Agents und Anti-Spam (lokaler Server)' '9. Transport Agents and Anti-Spam (local server)') 1))
        $null = $parts.Add((New-WdParagraph (L 'Exchange Server enthält integrierte Anti-Spam-Agents, die auf Mailbox-Servern standardmäßig nicht aktiviert sind. EXpress aktiviert die Anti-Spam-Agents und konfiguriert sie so, dass ausschließlich der Recipient Filter Agent aktiv bleibt — dieser prüft, ob Empfänger im Active Directory existieren, und lehnt E-Mails an nicht vorhandene Empfänger bereits auf SMTP-Ebene ab (Directory Harvest Attack Protection). Content Filter, Sender Filter und andere Agents werden deaktiviert, da diese Aufgaben in Unternehmensumgebungen typischerweise durch dedizierte Gateway-Lösungen (z. B. Hornetsecurity, Proofpoint, Mimeacst) oder Exchange Online Protection (EOP) übernommen werden.' 'Exchange Server includes built-in anti-spam agents that are not enabled by default on Mailbox servers. EXpress enables the anti-spam agents and configures them so that only the Recipient Filter Agent remains active — this checks whether recipients exist in Active Directory and rejects emails to non-existent recipients at the SMTP level (Directory Harvest Attack Protection). Content Filter, Sender Filter and other agents are disabled, as these tasks are typically handled by dedicated gateway solutions (e.g. Hornetsecurity, Proofpoint, Mimeacst) or Exchange Online Protection (EOP) in enterprise environments.')))
        $agentRows2 = [System.Collections.Generic.List[object[]]]::new()
        try {
            # Collect agents from all transport scopes (HubTransport is the default; on Mailbox
            # servers the FrontendTransport and MailboxSubmission/Delivery scopes each expose a
            # separate agent list). Deduplicate by Identity to keep the table compact.
            $seenAg = @{}
            $scopes = @('TransportService','FrontendTransport','MailboxSubmission','MailboxDelivery')
            $collected = @()
            foreach ($sc in $scopes) {
                try { $collected += @(Get-TransportAgent -TransportService $sc -ErrorAction SilentlyContinue) } catch { }
            }
            if (-not $collected -or $collected.Count -eq 0) {
                $collected = @(Get-TransportAgent -ErrorAction SilentlyContinue)
            }
            # Lookup used by section 9.1 to cross-reference org-wide *FilterConfig.Enabled
            # with the actual TransportAgent.Enabled state. Without this cross-reference the
            # doc shows "Enabled=True" for Content/Sender/Sender-ID even after the installer
            # has disabled the corresponding agents, because *FilterConfig.Enabled is just the
            # org-level feature switch, not the effective pipeline state.
            $script:__agentByKind = @{}
            foreach ($ag in $collected) {
                if (-not $ag) { continue }
                $agName = if ($ag.Name) { [string]$ag.Name } elseif ($ag.Identity) { [string]$ag.Identity } else { '(unbenannt)' }
                $kind = switch -Regex ($agName) {
                    'Content Filter'         { 'Content'; break }
                    'Sender Filter'          { 'Sender'; break }
                    'Recipient Filter'       { 'Recipient'; break }
                    'Sender ?Id|Sender Id'   { 'SenderId'; break }
                    'Connection Filter(ing)?'{ 'Connection'; break }
                    'Protocol Analysis'      { 'ProtocolAnalysis'; break }
                    default                  { $null }
                }
                if ($kind -and -not $script:__agentByKind.ContainsKey($kind)) {
                    $script:__agentByKind[$kind] = $ag
                }
                if ($seenAg.ContainsKey($agName)) { continue }
                $seenAg[$agName] = $true
                $agentState2 = if ($ag.Enabled) { (L 'Aktiv' 'Enabled') } else { (L 'Inaktiv' 'Disabled') }
                $agentRows2.Add(@($agName, $agentState2, $ag.Priority))
            }
        } catch { }

        # Helper — renders the effective pipeline state for a filter's underlying TransportAgent.
        # Distinguishes three cases so a reader can tell the difference between "org switch on, agent off"
        # (EXpress default: org config says Enabled, agent is disabled → filter inert) and the other two.
        function script:Get-EffectiveAgentState {
            param([string]$Kind)
            $ag = $script:__agentByKind[$Kind]
            if (-not $ag) { return (L 'Nicht installiert' 'Not installed') }
            if ($ag.Enabled) { return (L 'Aktiv — Agent läuft im Transport-Pipeline' 'Enabled — agent runs in transport pipeline') }
            return (L 'Inaktiv — Transport-Agent ist deaktiviert, Filter greift nicht (Org-Schalter ist nur ein Feature-Flag)' 'Inactive — transport agent is disabled, filter does not fire (org switch is only a feature flag)')
        }
        if ($agentRows2.Count -eq 0) { $agentRows2.Add(@((L '(keine konfiguriert)' '(none configured)'), '', '')) }
        $null = $parts.Add((New-WdTable -Headers @('Agent', (L 'Status' 'Status'), (L 'Priorität' 'Priority')) -Rows $agentRows2.ToArray()))

        # 9.1 Anti-Spam-Filter-Konfiguration (org-weite Filtereinstellungen)
        $hasAnyFilter = $orgD.ContentFilterConfig -or $orgD.SenderFilterConfig -or $orgD.RecipientFilterConfig -or $orgD.SenderIdConfig
        if ($hasAnyFilter) {
            $null = $parts.Add((New-WdHeading (L '9.1 Anti-Spam-Filter-Konfiguration' '9.1 Anti-Spam Filter Configuration') 2))
            $null = $parts.Add((New-WdParagraph (L 'Die folgenden Tabellen zeigen die organisationsweite Konfiguration der installierten Anti-Spam-Filter-Agents. In reinen on-premises-Umgebungen ohne vorgelagerten Cloud-Filter (EOP/Hornetsecurity/Proofpoint) sind diese Einstellungen aktiv wirksam. In Hybrid-Umgebungen oder mit vorgelagerten Gateways werden Content- und Sender-Filter typischerweise deaktiviert (Recipient Filter bleibt für Directory Harvest Attack Protection aktiv).' 'The following tables show the organisation-wide configuration of the installed anti-spam filter agents. In pure on-premises environments without an upstream cloud filter (EOP/Hornetsecurity/Proofpoint), these settings are actively effective. In hybrid environments or with upstream gateways, Content and Sender Filters are typically disabled (Recipient Filter remains active for Directory Harvest Attack Protection).')))
            $null = $parts.Add((New-WdParagraph (L 'Hinweis zur Unterscheidung: "Effektiver Status (Transport-Agent)" zeigt, ob der Agent tatsächlich in der Transport-Pipeline läuft (Get-TransportAgent). "Org-Konfig Enabled" ist nur der organisationsweite Feature-Schalter (Get-*FilterConfig) und sagt nichts darüber aus, ob der Filter wirklich greift. EXpress deaktiviert standardmäßig alle Transport-Agents außer dem Recipient Filter — "Org-Konfig Enabled = True" bei deaktiviertem Transport-Agent bedeutet daher: Filter greift nicht.' 'Note on interpretation: "Effective status (transport agent)" shows whether the agent actually runs in the transport pipeline (Get-TransportAgent). "Org config Enabled" is only the organisation-wide feature switch (Get-*FilterConfig) and says nothing about whether the filter actually fires. EXpress disables all transport agents by default except Recipient Filter — "Org config Enabled = True" with a disabled transport agent therefore means: filter does not fire.')))
            if ($orgD.ContentFilterConfig) {
                $cf = $orgD.ContentFilterConfig
                $cfRows = [System.Collections.Generic.List[object[]]]::new()
                $cfRows.Add(@((L 'Effektiver Status (Transport-Agent)' 'Effective status (transport agent)'), (Get-EffectiveAgentState 'Content')))
                $cfRows.Add(@((L 'Org-Konfig Enabled (Feature-Flag)' 'Org config Enabled (feature flag)'),   (Format-RegBool $cf.Enabled)))
                $cfRows.Add(@((L 'Aktion (SCL ≥ 6)' 'Action (SCL ≥ 6)'),  (SafeVal $cf.SCLRejectEnabled (L '(nicht gesetzt)' '(not set)'))))
                $cfRows.Add(@((L 'SCL-Ablehneschwellenwert' 'SCL reject threshold'), (SafeVal $cf.SCLRejectThreshold)))
                $cfRows.Add(@((L 'SCL-Löschschwellenwert' 'SCL delete threshold'),  (SafeVal $cf.SCLDeleteThreshold)))
                $cfRows.Add(@((L 'SCL-Quarantäneschwellenwert' 'SCL quarantine threshold'), (SafeVal $cf.SCLQuarantineThreshold)))
                $cfRows.Add(@((L 'Quarantäne-Postfach' 'Quarantine mailbox'),       (SafeVal $cf.QuarantineMailbox (L '(nicht gesetzt)' '(not set)'))))
                $null = $parts.Add((New-WdHeading (L 'Content Filter' 'Content Filter') 3))
                $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $cfRows.ToArray()))
            }
            if ($orgD.SenderFilterConfig) {
                $sf = $orgD.SenderFilterConfig
                $sfRows = [System.Collections.Generic.List[object[]]]::new()
                $sfRows.Add(@((L 'Effektiver Status (Transport-Agent)' 'Effective status (transport agent)'), (Get-EffectiveAgentState 'Sender')))
                $sfRows.Add(@((L 'Org-Konfig Enabled (Feature-Flag)' 'Org config Enabled (feature flag)'),   (Format-RegBool $sf.Enabled)))
                $sfRows.Add(@((L 'Leere Absender blockieren' 'Block blank senders'), (Format-RegBool $sf.BlankSenderBlockingEnabled)))
                $sfBlockedSenders = if ($sf.BlockedSenders) { $sf.BlockedSenders -join '; ' } else { $null }
                $sfBlockedDomains = if ($sf.BlockedDomains) { $sf.BlockedDomains -join '; ' } else { $null }
                $sfRows.Add(@((L 'Blockliste (Absender)' 'Block list (senders)'), (SafeVal $sfBlockedSenders (L '(leer)' '(empty)'))))
                $sfRows.Add(@((L 'Blockliste (Domänen)' 'Block list (domains)'),  (SafeVal $sfBlockedDomains (L '(leer)' '(empty)'))))
                $null = $parts.Add((New-WdHeading (L 'Sender Filter' 'Sender Filter') 3))
                $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $sfRows.ToArray()))
            }
            if ($orgD.RecipientFilterConfig) {
                $rf = $orgD.RecipientFilterConfig
                $rfRows = [System.Collections.Generic.List[object[]]]::new()
                $rfRows.Add(@((L 'Effektiver Status (Transport-Agent)' 'Effective status (transport agent)'), (Get-EffectiveAgentState 'Recipient')))
                $rfRows.Add(@((L 'Org-Konfig Enabled (Feature-Flag)' 'Org config Enabled (feature flag)'),   (Format-RegBool $rf.Enabled)))
                $rfBlockedRecipients = if ($rf.BlockedRecipients) { $rf.BlockedRecipients -join '; ' } else { $null }
                $rfRows.Add(@((L 'Blockliste (Empfänger)' 'Block list (recipients)'), (SafeVal $rfBlockedRecipients (L '(leer)' '(empty)'))))
                $rfRows.Add(@((L 'Empfänger-Lookup aktiviert' 'Recipient lookup enabled'), (Format-RegBool $rf.RecipientValidationEnabled)))
                $null = $parts.Add((New-WdHeading (L 'Recipient Filter' 'Recipient Filter') 3))
                $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $rfRows.ToArray()))
            }
            if ($orgD.SenderIdConfig) {
                $si = $orgD.SenderIdConfig
                $siRows = [System.Collections.Generic.List[object[]]]::new()
                $siRows.Add(@((L 'Effektiver Status (Transport-Agent)' 'Effective status (transport agent)'), (Get-EffectiveAgentState 'SenderId')))
                $siRows.Add(@((L 'Org-Konfig Enabled (Feature-Flag)' 'Org config Enabled (feature flag)'),   (Format-RegBool $si.Enabled)))
                $siRows.Add(@((L 'Aktion (Spoofed)' 'Action (spoofed)'),             (SafeVal $si.SpoofedDomainAction)))
                $siRows.Add(@((L 'Aktion (Temporary Error)' 'Action (temp error)'),  (SafeVal $si.TempErrorAction)))
                $null = $parts.Add((New-WdHeading 'Sender ID' 3))
                $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $siRows.ToArray()))
            }
        }

        # ── 10. Backup- & DR-Readiness (lokal) ────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '10. Backup- und DR-Readiness' '10. Backup and DR Readiness') 1))
        $null = $parts.Add((New-WdParagraph (L 'Exchange Server unterstützt datenbankebene Sicherungen über die Volume Shadow Copy Service (VSS)-Schnittstelle. Eine ordnungsgemäß funktionierende VSS-Integration ist Voraussetzung für konsistente Exchange-Backups durch Backup-Software (Veeam, Windows Server Backup, Commvault u. a.). Nach einem Backup werden die Transaktionsprotokolle automatisch abgeschnitten (Log Truncation) — vorausgesetzt, Circular Logging ist deaktiviert. Für die Disaster-Recovery-Fähigkeit sind funktionierende VSS Writer, korrekte Exchange-Defender-Ausnahmen und ein regelmäßig getestetes Restore-Verfahren entscheidend.' 'Exchange Server supports database-level backups via the Volume Shadow Copy Service (VSS) interface. Correctly functioning VSS integration is a prerequisite for consistent Exchange backups by backup software (Veeam, Windows Server Backup, Commvault, etc.). After a backup, transaction logs are automatically truncated — provided Circular Logging is disabled. For disaster recovery capability, functioning VSS writers, correct Exchange Defender exclusions and a regularly tested restore procedure are essential.')))
        $null = $parts.Add((New-WdHeading (L '10.1 VSS Writer Status' '10.1 VSS Writer Status') 2))
        $null = $parts.Add((New-WdParagraph (L 'Alle Exchange-relevanten VSS Writer müssen im Zustand "Stabil" sein. Fehlerhafte Writer führen zu inkonsistenten oder fehlschlagenden Backups. Bei dauerhaft fehlerhaften Writern ist ein Neustart des betroffenen Dienstes (Microsoft Exchange Writer → MSExchangeIS) oder ein Server-Neustart erforderlich.' 'All Exchange-relevant VSS writers must be in a "Stable" state. Faulty writers lead to inconsistent or failed backups. For persistently faulty writers, a restart of the affected service (Microsoft Exchange Writer → MSExchangeIS) or a server restart is required.')))
        $vssRows = [System.Collections.Generic.List[object[]]]::new()
        try {
            $vssOut = (vssadmin list writers 2>&1) -join "`n"
            $curWriter = ''
            foreach ($line in ($vssOut -split "`n")) {
                if ($line -match "Writer name:\s+'(.+)'") { $curWriter = $Matches[1] }
                elseif ($line -match 'State:\s*\[\d+\]\s+(.+)') { $vssRows.Add(@($curWriter, $Matches[1].Trim())) }
            }
        } catch { $vssRows.Add(@((L 'VSS-Abfrage fehlgeschlagen' 'VSS query failed'), '')) }
        if ($vssRows.Count -eq 0) { $vssRows.Add(@((L '(keine VSS Writer gefunden)' '(no VSS writers found)'), '')) }
        $null = $parts.Add((New-WdTable -Headers @((L 'VSS Writer' 'VSS Writer'), (L 'Zustand' 'State')) -Rows $vssRows.ToArray()))
        $null = $parts.Add((New-WdHeading (L '10.2 Empfehlungen Backup-Strategie' '10.2 Backup Strategy Recommendations') 2))
        $null = $parts.Add((New-WdParagraph (L 'Für Exchange Server werden folgende Backup-Praktiken empfohlen:' 'The following backup practices are recommended for Exchange Server:')))
        $null = $parts.Add((New-WdBullet (L 'Tägliche VSS-Vollsicherung der Exchange-Datenbanken über eine Exchange-aware Backup-Lösung (kein File-Level-Backup laufender EDB-Dateien)' 'Daily VSS full backup of Exchange databases via an Exchange-aware backup solution (no file-level backup of running EDB files)')))
        $null = $parts.Add((New-WdBullet (L 'Transaktionsprotokolle werden nach erfolgreichem Backup automatisch abgeschnitten — Circular Logging sollte deaktiviert bleiben' 'Transaction logs are automatically truncated after a successful backup — Circular Logging should remain disabled')))
        $null = $parts.Add((New-WdBullet (L 'Restore-Test mindestens einmal jährlich in einer Testumgebung (Recovery Database, RDB) durchführen' 'Perform restore test at least once annually in a test environment (Recovery Database, RDB)')))
        $null = $parts.Add((New-WdBullet (L 'Backup der Active-Directory-Domänencontroller separat sicherstellen (Exchange ist AD-abhängig)' 'Ensure separate backup of Active Directory domain controllers (Exchange is AD-dependent)')))
        $null = $parts.Add((New-WdHeading (L '10.3 Disaster-Recovery-Szenarien' '10.3 Disaster Recovery Scenarios') 2))
        $null = $parts.Add((New-WdParagraph (L 'Die folgende Tabelle gibt einen Überblick über typische DR-Szenarien und die empfohlene Vorgehensweise.' 'The table below provides an overview of typical DR scenarios and the recommended approach.')))
        $drRows = @(
            ,@((L 'Datenbankausfall (keine DAG)' 'Database failure (no DAG)'), (L 'Restore aus Backup in Recovery Database (RDB), Mailbox-Merge in Produktionsdatenbank' 'Restore from backup into Recovery Database (RDB), mailbox merge into production database'))
            ,@((L 'Datenbankausfall (DAG vorhanden)' 'Database failure (DAG present)'), (L 'Automatischer/manueller Failover auf Datenbankkopie; fehlerhafte Kopie per Update-MailboxDatabaseCopy reseed' 'Automatic/manual failover to database copy; reseed faulty copy via Update-MailboxDatabaseCopy'))
            ,@((L 'Server-Totalausfall' 'Complete server failure'), (L 'setup.exe /m:RecoverServer auf ersetztem Server; danach Datenbanken mounten bzw. DAG-Kopien reseed' 'setup.exe /m:RecoverServer on replacement server; then mount databases or reseed DAG copies'))
            ,@((L 'Verlust des File Share Witness (FSW)' 'Loss of File Share Witness (FSW)'), (L 'DAG kann noch lesen; Alternate FSW übernimmt automatisch (wenn konfiguriert). Manuell: Set-DatabaseAvailabilityGroup -AlternateWitnessServer' 'DAG can still read; Alternate FSW takes over automatically (if configured). Manually: Set-DatabaseAvailabilityGroup -AlternateWitnessServer'))
            ,@((L 'Active-Directory-Ausfall' 'Active Directory failure'), (L 'Exchange kann ohne AD nicht starten (Ausnahme: Edge Transport). AD-Wiederherstellung hat Vorrang.' 'Exchange cannot start without AD (exception: Edge Transport). AD recovery takes priority.'))
        )
        $null = $parts.Add((New-WdTable -Headers @((L 'Szenario' 'Scenario'), (L 'Vorgehensweise' 'Procedure')) -Rows $drRows))

        # 10.4 Backup-Nachweis und Testzyklus
        $null = $parts.Add((New-WdHeading (L '10.4 Backup-Nachweis und Testzyklus' '10.4 Backup Evidence and Test Cycle') 2))
        $null = $parts.Add((New-WdParagraph (L 'Für Auditierbarkeit muss dokumentiert sein, dass Backups regelmäßig durchgeführt und getestet werden. Bitte nach Abschluss der ersten Produktionsbackups und nach jedem Restore-Test ausfüllen.' 'For auditability it must be documented that backups are performed and tested regularly. Please complete after the first production backups and after each restore test.')))
        $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert / Datum' 'Value / Date')) -Rows @(
            ,@((L 'Backup-Lösung (Produkt)' 'Backup solution (product)'), '')
            ,@((L 'Erstes erfolgreiches Backup' 'First successful backup'), '')
            ,@((L 'Backup-Frequenz' 'Backup frequency'), '')
            ,@((L 'Aufbewahrungsdauer Backups' 'Backup retention period'), '')
            ,@((L 'Letzter Restore-Test (Datum)' 'Last restore test (date)'), '')
            ,@((L 'Restore-Test durchgeführt von' 'Restore test performed by'), '')
            ,@((L 'Restore-Ergebnis' 'Restore result'), '')
            ,@((L 'Nächster Restore-Test geplant' 'Next restore test planned'), '')
        )))

        # ── 11. HealthChecker ──────────────────────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '11. HealthChecker' '11. HealthChecker') 1))
        $null = $parts.Add((New-WdParagraph (L 'Der Microsoft CSS Exchange HealthChecker ist ein offizielles Diagnoseskript des Microsoft Exchange-Teams (https://aka.ms/ExchangeHealthChecker). Er prüft den Exchange-Server auf bekannte Konfigurationsprobleme, fehlende Sicherheitsupdates, falsche Registry-Einstellungen, TLS-Konfiguration, Zertifikatsprobleme, OS-Konfiguration und Performance-Indikatoren. Der HealthChecker wird am Ende jeder EXpress-Installation automatisch ausgeführt. Das Ergebnis sollte nach der Installation gesichtet und offene Findings abgearbeitet werden.' 'The Microsoft CSS Exchange HealthChecker is an official diagnostic script from the Microsoft Exchange team (https://aka.ms/ExchangeHealthChecker). It checks the Exchange server for known configuration issues, missing security updates, incorrect registry settings, TLS configuration, certificate issues, OS configuration and performance indicators. HealthChecker is automatically executed at the end of every EXpress installation. The result should be reviewed after installation and any open findings addressed.')))
        $hcPath = SafeVal $State['HCReportPath']
        if ($hcPath) {
            $null = $parts.Add((New-WdParagraph ((L 'HealthChecker HTML-Report (generiert während der Installation): ' 'HealthChecker HTML report (generated during installation): ') + $hcPath)))
        } else {
            $null = $parts.Add((New-WdParagraph (L 'HealthChecker wurde nicht ausgeführt oder der Report-Pfad ist nicht verfügbar. Bitte manuell ausführen: https://aka.ms/ExchangeHealthChecker' 'HealthChecker was not run or the report path is not available. Please run manually: https://aka.ms/ExchangeHealthChecker')))
        }

        # ── 12. Monitoring-Readiness ───────────────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '12. Monitoring-Readiness' '12. Monitoring Readiness') 1))
        $null = $parts.Add((New-WdParagraph (L 'Exchange Server enthält mit Managed Availability ein eingebautes Selbstheilungssystem, das Komponenten überwacht und bei Fehler automatisch Recover-Aktionen auslöst (Dienst-Neustart, IIS-Reset, Server-Failover). Managed Availability ersetzt jedoch kein aktives externes Monitoring. Für den produktiven Betrieb wird ein dediziertes Monitoring-System empfohlen, das Exchange-spezifische Metriken, Event-Log-Einträge und Service-Zustände überwacht.' 'Exchange Server includes Managed Availability, a built-in self-healing system that monitors components and automatically triggers recovery actions on failure (service restart, IIS reset, server failover). However, Managed Availability does not replace active external monitoring. A dedicated monitoring system is recommended for production operation, monitoring Exchange-specific metrics, event log entries and service states.')))
        $monRows = [System.Collections.Generic.List[object[]]]::new()
        try { $svc2 = Get-Service MSExchangeMitigation -ErrorAction SilentlyContinue; if ($svc2) { $monRows.Add(@('EEMS (MSExchangeMitigation)', $svc2.Status.ToString())) } } catch { }
        try {
            $evtLogs2 = @('Application','System','MSExchange Management') | ForEach-Object {
                try { '{0}: MaxSize={1}MB' -f $_, [math]::Round((Get-WinEvent -ListLog $_ -ErrorAction Stop).MaximumSizeInBytes / 1MB, 0) } catch { }
            } | Where-Object { $_ }
            if ($evtLogs2) { $monRows.Add(@((L 'Event-Log-Größen' 'Event log sizes'), ($evtLogs2 -join '; '))) }
        } catch { }
        if ($monRows.Count -eq 0) { $monRows.Add(@((L '(keine Daten abrufbar)' '(no data available)'), '')) }
        $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert / Status' 'Value / status')) -Rows $monRows.ToArray()))
        $null = $parts.Add((New-WdParagraph (L 'Empfehlungen für das Monitoring nach Go-Live:' 'Recommendations for monitoring after go-live:')))
        $null = $parts.Add((New-WdBullet (L 'Perfmon-Baseline innerhalb von 4 Wochen nach Go-Live aufzeichnen (MSExchangeIS, RPC-Latenz, Disk-Queue, CPU)' 'Record Perfmon baseline within 4 weeks of go-live (MSExchangeIS, RPC latency, disk queue, CPU)')))
        $null = $parts.Add((New-WdBullet (L 'Event-IDs überwachen: 1009 (MSExchangeIS), 2142/2144 (RPC-Latenz), 4999 (Watson), 1022 (Datenbankfehler)' 'Monitor event IDs: 1009 (MSExchangeIS), 2142/2144 (RPC latency), 4999 (Watson), 1022 (database errors)')))
        $null = $parts.Add((New-WdBullet (L 'Exchange-Zertifikatsablauf überwachen — Auth-Zertifikat (2 Jahre) und IIS/SMTP-Zertifikat (kundenabhängig). MEAC-Scheduled-Task übernimmt Auth-Cert-Erneuerung automatisch.' 'Monitor Exchange certificate expiry — Auth certificate (2 years) and IIS/SMTP certificate (customer-dependent). MEAC scheduled task handles Auth Cert renewal automatically.')))
        $null = $parts.Add((New-WdBullet (L 'Datenbankkopienstatus (DAG): Get-MailboxDatabaseCopyStatus täglich prüfen oder per Monitoring automatisieren' 'Database copy status (DAG): Check Get-MailboxDatabaseCopyStatus daily or automate via monitoring')))

        # 12.1 Exchange Crimson Event Log Channels
        $null = $parts.Add((New-WdHeading (L '12.1 Exchange Crimson Event Log Kanäle' '12.1 Exchange Crimson Event Log Channels') 2))
        $null = $parts.Add((New-WdParagraph (L 'Exchange schreibt strukturierte Ereignisdaten in dedizierte Windows-Ereigniskanäle ("Crimson Channels") unterhalb von Microsoft-Exchange-*. Diese Kanäle sind feingranularer als das Application-Protokoll und ermöglichen gezieltes Monitoring einzelner Exchange-Subsysteme. Die folgende Tabelle zeigt alle aktivierten oder bereits beschriebenen Exchange-Ereigniskanäle auf diesem Server.' 'Exchange writes structured event data to dedicated Windows event channels ("Crimson channels") under Microsoft-Exchange-*. These channels are more granular than the Application log and allow targeted monitoring of individual Exchange subsystems. The table below shows all enabled or already written Exchange event channels on this server.')))
        $crimsonRows = [System.Collections.Generic.List[object[]]]::new()
        try {
            $exchLogs = @(Get-WinEvent -ListLog 'Microsoft-Exchange*' -ErrorAction SilentlyContinue |
                Where-Object { ($_.IsEnabled -or $_.RecordCount -gt 0) -and $_.LogName -match '/Operational$|/Admin$' } |
                Sort-Object LogName)
            foreach ($log in $exchLogs) {
                $sizeMB   = if ($log.MaximumSizeInBytes -gt 0) { '{0} MB' -f [math]::Round($log.MaximumSizeInBytes / 1MB, 0) } else { '—' }
                $records  = if ($log.RecordCount -gt 0) { $log.RecordCount.ToString() } else { '0' }
                # NOTE: $logState not $state/$State — PowerShell is case-insensitive; $state would shadow the outer $State hashtable.
                $logState = if ($log.IsEnabled) { (L 'aktiv' 'enabled') } else { (L 'inaktiv' 'disabled') }
                $crimsonRows.Add(@($log.LogName, $logState, $sizeMB, $records))
            }
        } catch { }
        if ($crimsonRows.Count -eq 0) { $crimsonRows.Add(@((L '(keine Kanäle gefunden oder WinEvent nicht verfügbar)' '(no channels found or WinEvent not available)'), '', '', '')) }
        $null = $parts.Add((New-WdTable -Headers @((L 'Kanal' 'Channel'), (L 'Status' 'State'), (L 'Max. Größe' 'Max size'), (L 'Einträge' 'Records')) -Rows $crimsonRows.ToArray()))
        $null = $parts.Add((New-WdParagraph (L 'Wichtige Kanäle für das Exchange-Monitoring: Microsoft-Exchange-HighAvailability/Operational (DAG-Failover), Microsoft-Exchange-ManagedAvailability/Monitoring (Selbstheilung), Microsoft-Exchange-Store Driver/Operational (Mailbox-Speicher), Microsoft-Exchange-Transport/Operational (Mailflow). Für historische Fehlersuche: Get-WinEvent -LogName "Microsoft-Exchange-*" -MaxEvents 1000.' 'Key channels for Exchange monitoring: Microsoft-Exchange-HighAvailability/Operational (DAG failover), Microsoft-Exchange-ManagedAvailability/Monitoring (self-healing), Microsoft-Exchange-Store Driver/Operational (mailbox store), Microsoft-Exchange-Transport/Operational (mail flow). For historical troubleshooting: Get-WinEvent -LogName "Microsoft-Exchange-*" -MaxEvents 1000.')))

        # 12.2 Monitoring-Checkliste Go-Live
        $null = $parts.Add((New-WdHeading (L '12.2 Monitoring-Checkliste Go-Live' '12.2 Monitoring Checklist Go-Live') 2))
        $null = $parts.Add((New-WdParagraph (L 'Die folgende Checkliste dokumentiert den Aufbau des produktiven Monitorings nach Go-Live. Bitte nach Einrichtung jedes Monitoring-Elements ausfüllen.' 'The checklist below documents the setup of production monitoring after go-live. Please complete after each monitoring element is configured.')))
        $null = $parts.Add((New-WdTable -Headers @((L 'Monitoring-Element' 'Monitoring element'), (L 'Tool / System' 'Tool / system'), (L 'Eingerichtet (Datum)' 'Configured (date)'), (L 'Verantwortlich' 'Owner')) -Rows @(
            ,@((L 'Exchange-Dienst-Überwachung (MSExchange*)' 'Exchange service monitoring (MSExchange*)'), '', '', '')
            ,@((L 'Zertifikatsablauf-Überwachung (IIS/SMTP)' 'Certificate expiry monitoring (IIS/SMTP)'), '', '', '')
            ,@((L 'Postfachvolumen / Datenbankgröße' 'Mailbox volume / database size'), '', '', '')
            ,@((L 'Datenbankkopien-Status (DAG)' 'Database copy status (DAG)'), '', '', '')
            ,@((L 'Mailflow-Test (eingehend + ausgehend)' 'Mail flow test (inbound + outbound)'), '', '', '')
            ,@((L 'Log-Volume-Auslastung' 'Log volume utilisation'), '', '', '')
            ,@((L 'Event-ID-Alerting (1009, 4999, 1022)' 'Event ID alerting (1009, 4999, 1022)'), '', '', '')
            ,@((L 'Perfmon-Baseline aufgezeichnet' 'Perfmon baseline recorded'), '', '', '')
            ,@('HealthChecker (nach jedem SU/CU)', '', '', '')
        )))

        # ── 13. Public Folders ─────────────────────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '13. Öffentliche Ordner' '13. Public Folders') 1))
        $null = $parts.Add((New-WdParagraph (L 'Öffentliche Ordner (Public Folders) sind eine Legacy-Kollaborationsfunktion in Exchange, die seit Exchange 2013 auf Postfach-Infrastruktur (Public Folder Mailboxes) umgestellt wurde ("Modern Public Folders"). Sie ermöglichen gemeinsamen Zugriff auf E-Mail, Kalender, Kontakte und Dateien in einer Ordnerhierarchie, die allen Benutzern oder ausgewählten Gruppen zugänglich ist. In modernen Umgebungen werden Public Folders zunehmend durch Shared Mailboxes (gemeinsamer Posteingang, geteilter Kalender) und Microsoft Teams/SharePoint (Dokumentenablage, Teamzusammenarbeit) abgelöst. Microsoft hat mehrfach die Abkündigung von Public Folders angekündigt und empfiehlt für neue Implementierungen ausschließlich die modernen Alternativen.' 'Public Folders are a legacy collaboration feature in Exchange that has been migrated to mailbox infrastructure (Public Folder Mailboxes) since Exchange 2013 ("Modern Public Folders"). They allow shared access to email, calendars, contacts and files in a folder hierarchy accessible to all users or selected groups. In modern environments, Public Folders are increasingly replaced by Shared Mailboxes (shared inbox, shared calendar) and Microsoft Teams/SharePoint (document storage, team collaboration). Microsoft has announced the deprecation of Public Folders multiple times and recommends only the modern alternatives for new implementations.')))
        $null = $parts.Add((New-WdParagraph (L 'Hinweis zur Migration: Öffentliche Ordner können nach Exchange Online migriert werden (Migration zu EXO Modern Public Folders). Alternativ können Inhalte in Shared Mailboxes oder SharePoint-Dokumentbibliotheken überführt werden. Für die Migration zu EXO ist das Skript-Paket unter https://aka.ms/publicfoldermigration verfügbar.' 'Migration note: Public Folders can be migrated to Exchange Online (migration to EXO Modern Public Folders). Alternatively, contents can be transferred to Shared Mailboxes or SharePoint document libraries. The script package for migration to EXO is available at https://aka.ms/publicfoldermigration.')))
        try {
            $pfMailboxes = @(Get-Mailbox -PublicFolder -ErrorAction SilentlyContinue)
            if ($pfMailboxes -and $pfMailboxes.Count -gt 0) {
                $null = $parts.Add((New-WdParagraph (L 'Folgende Public-Folder-Postfächer sind in der Organisation konfiguriert:' 'The following Public Folder mailboxes are configured in the organisation:')))
                $pfRows = $pfMailboxes | ForEach-Object { @($_.Name, (SafeVal $_.PrimarySmtpAddress), (SafeVal $_.Database), (SafeVal $_.IsRootPublicFolderMailbox)) }
                $null = $parts.Add((New-WdTable -Headers @((L 'Name' 'Name'), 'SMTP', (L 'Datenbank' 'Database'), (L 'Root-PF-Postfach' 'Root PF mailbox')) -Rows $pfRows))
                try {
                    $pfStats = Get-PublicFolderStatistics -ErrorAction SilentlyContinue | Measure-Object -Property ItemCount, TotalItemSize -Sum
                    if ($pfStats) {
                        $pfCountRow = [System.Collections.Generic.List[object[]]]::new()
                        $pfCountRow.Add(@((L 'Anzahl Öffentliche Ordner (gesamt)' 'Total public folder count'), (SafeVal ($pfStats | Where-Object { $_.Property -eq 'ItemCount' } | Select-Object -ExpandProperty Sum))))
                        $null = $parts.Add((New-WdTable -Headers @((L 'Statistik' 'Statistic'), (L 'Wert' 'Value')) -Rows $pfCountRow.ToArray()))
                    }
                } catch { }
            } else {
                $null = $parts.Add((New-WdParagraph (L 'Öffentliche Ordner sind in dieser Organisation nicht konfiguriert. Es sind keine Public-Folder-Postfächer vorhanden.' 'Public Folders are not configured in this organisation. No Public Folder mailboxes exist.')))
            }
        } catch {
            $null = $parts.Add((New-WdParagraph (L 'Abfrage nicht möglich (Edge/Management-Tools-Modus oder keine Exchange-Session).' 'Query not possible (Edge/Management Tools mode or no Exchange session).')))
        }

        # ── 14. Ausgeführte Konfigurationsbefehle (nur bei tatsächlichem Setup-Lauf) ──
        # Chronological list of the config-level cmdlets the script actually ran
        # during this installation (recorded via Register-ExecutedCommand). Covers
        # Virtual Directory URLs, antispam config, relay connectors, certificate
        # import/enable, DAG join, send-connector source updates, and scheduled
        # tasks. Low-level hardening (registry/Schannel/services) is described in
        # the preceding chapters and is not repeated here to keep the list readable.
        if (-not $isAdHoc) {
            $null = $parts.Add((New-WdHeading (L '14. Ausgeführte Konfigurationsbefehle' '14. Executed configuration commands') 1))
            $execCmds = @()
            if ($State.ContainsKey('ExecutedCommands') -and $State['ExecutedCommands']) {
                $execCmds = @($State['ExecutedCommands'])
            }
            if ($execCmds.Count -eq 0) {
                $null = $parts.Add((New-WdParagraph (L 'Während dieses Laufs wurden keine Konfigurationsbefehle aufgezeichnet (z. B. reiner Tools-Modus oder Lauf ohne Namespace/Zertifikat/DAG).' 'No configuration commands were recorded during this run (e.g. tools-only mode or run without namespace/certificate/DAG).')))
            }
            else {
                $null = $parts.Add((New-WdParagraph (L 'Die folgenden Befehle wurden in chronologischer Reihenfolge mit der angegebenen Syntax ausgeführt. Passwörter und Secure-Strings sind durch Platzhalter ersetzt.' 'The following commands were executed in chronological order with the shown syntax. Passwords and secure strings are replaced by placeholders.')))
                $byCat = $execCmds | Group-Object -Property Category | Sort-Object Name
                $catIdx = 0
                foreach ($g in $byCat) {
                    $catIdx++
                    $catLabel = if ($g.Name) { $g.Name } else { (L 'Sonstige' 'Other') }
                    $null = $parts.Add((New-WdHeading ('14.{0} {1}' -f $catIdx, $catLabel) 2))
                    foreach ($e in $g.Group) {
                        foreach ($cmd in ($e.Command -split '; ')) {
                            $null = $parts.Add((New-WdCode $cmd.Trim()))
                        }
                    }
                }
            }
            $null = $parts.Add((New-WdParagraph (L 'Die vollständige Installationsausgabe (inkl. Statusmeldungen, Versionsprüfungen und Fehlern) steht in der EXpress-Logdatei (siehe Kapitel 1 "Dokumenteigenschaften" → "Logdatei").' 'The complete installation output (including status messages, version checks, and errors) is available in the EXpress log file (see chapter 1 "Document Properties" → "Log file").' )))
        }

        # ── 15. Exchange Online und Microsoft 365 (promoted from former §4.17) ─────
        # Placed here, directly before the runbooks, so hybrid/EXO considerations are
        # read together with day-2 operations rather than buried inside §4 org-config.
        $null = $parts.Add((New-WdHeading (L '15. Exchange Online und Microsoft 365' '15. Exchange Online and Microsoft 365') 1))
        $null = $parts.Add((New-WdParagraph (L 'Exchange Online (EXO) ist die cloud-gehostete E-Mail-Plattform in Microsoft 365. In Hybrid-Szenarien koexistieren Exchange Server on-premises und Exchange Online — Postfächer können auf beiden Plattformen liegen, E-Mails werden plattformübergreifend weitergeleitet (Shared Namespace), und Benutzer erfahren keine funktionalen Unterschiede. Der Hybrid Configuration Wizard (HCW) richtet die notwendigen Konnektoren, Zertifikate und OAuth-Vertrauensbeziehungen ein.' 'Exchange Online (EXO) is the cloud-hosted email platform in Microsoft 365. In hybrid scenarios, Exchange Server on-premises and Exchange Online coexist — mailboxes can reside on either platform, emails are routed across platforms (Shared Namespace), and users experience no functional differences. The Hybrid Configuration Wizard (HCW) sets up the necessary connectors, certificates and OAuth trust relationships.')))
        $null = $parts.Add((New-WdParagraph (L 'Folgende Aspekte sind in Hybrid-Umgebungen besonders zu beachten:' 'The following aspects are particularly important in hybrid environments:')))
        $null = $parts.Add((New-WdBullet (L 'Mailflow-Routing: In Centralised Mail Transport (CMT) läuft alle E-Mail über den on-premises-Server — ideal für Compliance/Archivierung. In dezentralem Routing sendet EXO direkt. CMT verursacht höhere Latenz und Abhängigkeit vom on-premises-System.' 'Mail flow routing: In Centralised Mail Transport (CMT) all email passes through the on-premises server — ideal for compliance/archiving. In decentralised routing EXO sends directly. CMT causes higher latency and dependency on the on-premises system.')))
        $null = $parts.Add((New-WdBullet (L 'Free/Busy-Integration: Verfügbarkeitsanzeige zwischen on-premises- und EXO-Postfächern erfordert funktionierende OAuth/Federation-Vertrauensbeziehung (Get-FederationTrust, Get-IntraOrganizationConnector). Bei Fehler sehen Benutzer "Keine Informationen" für cloud-Kalender.' 'Free/Busy integration: Availability display between on-premises and EXO mailboxes requires a functioning OAuth/Federation trust (Get-FederationTrust, Get-IntraOrganizationConnector). On failure, users see "No information" for cloud calendars.')))
        $null = $parts.Add((New-WdBullet (L 'Postfach-Migration (Move Request): Postfächer werden über New-MoveRequest zwischen on-premises und EXO bewegt. MRSProxy-Endpunkt muss auf dem on-premises-CAS extern erreichbar sein (TCP 443, mrsProxy.svc).' 'Mailbox migration (Move Request): Mailboxes are moved between on-premises and EXO via New-MoveRequest. MRSProxy endpoint must be externally reachable on the on-premises CAS (TCP 443, mrsProxy.svc).')))
        $null = $parts.Add((New-WdBullet (L 'Exchange Online Protection (EOP) / Defender for Office 365: In Hybrid-Szenarien ist EOP für eingehende E-Mails aus dem Internet der primäre Schutz. On-premises Anti-Spam-Filter (Content Filter, Sender Filter) werden typischerweise deaktiviert, da EOP/MDO die Filterung bereits vollständig übernimmt.' 'Exchange Online Protection (EOP) / Defender for Office 365: In hybrid scenarios, EOP is the primary protection for inbound email from the internet. On-premises anti-spam filters (Content Filter, Sender Filter) are typically disabled as EOP/MDO already performs complete filtering.')))
        $null = $parts.Add((New-WdBullet (L 'Namespace-Planung: Alle HTTPS-Dienste (OWA, EWS, Autodiscover, MAPI) sollten über einen einzigen externen FQDN erreichbar sein, der auf den on-premises-Exchange oder einen vorgelagerten Reverse-Proxy zeigt. EXO-Benutzer nutzen denselben Autodiscover-FQDN; der SCP-Record im AD ist für interne Clients maßgebend.' 'Namespace planning: All HTTPS services (OWA, EWS, Autodiscover, MAPI) should be reachable via a single external FQDN pointing to the on-premises Exchange or a reverse proxy. EXO users use the same Autodiscover FQDN; the SCP record in AD is authoritative for internal clients.')))
        $null = $parts.Add((New-WdBullet (L 'Lizenzierung: Exchange Online-Postfächer benötigen eine M365-Lizenz mit Exchange Online-Plan (F1, E1, E3, E5). On-premises-Postfächer benötigen Exchange Server-CALs (Standard/Enterprise). In Hybrid-Szenarien dürfen keine EXO-Lizenzen für on-premises-Postfächer zugewiesen werden.' 'Licensing: Exchange Online mailboxes require an M365 licence with an Exchange Online plan (F1, E1, E3, E5). On-premises mailboxes require Exchange Server CALs (Standard/Enterprise). In hybrid scenarios, EXO licences must not be assigned to on-premises mailboxes.')))
        if ($scope -in 'All','Org','Local' -and $orgD -and $orgD.HybridConfig) {
            $hyb3 = $orgD.HybridConfig
            $eo365Rows = [System.Collections.Generic.List[object[]]]::new()
            $eo365Rows.Add(@((L 'Hybrid-Konfiguration' 'Hybrid configuration'), (L 'Aktiv — Hybrid Configuration Wizard wurde ausgeführt' 'Active — Hybrid Configuration Wizard has been run')))
            if ($hyb3.OnPremisesSMTPDomains) { $eo365Rows.Add(@((L 'Freigegebene SMTP-Domänen' 'Shared SMTP domains'), ($hyb3.OnPremisesSMTPDomains -join ', '))) }
            if ($hyb3.Features) { $eo365Rows.Add(@((L 'HCW-Features' 'HCW features'), ($hyb3.Features -join ', '))) }
            $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $eo365Rows.ToArray()))
        } else {
            $null = $parts.Add((New-WdParagraph (L 'Hybrid Configuration Wizard wurde (noch) nicht ausgeführt — diese Exchange-Umgebung ist rein on-premises. Für eine spätere Migration zu Exchange Online ist der HCW der empfohlene Einstiegspunkt: https://aka.ms/HybridWizard' 'Hybrid Configuration Wizard has not (yet) been run — this Exchange environment is purely on-premises. For a later migration to Exchange Online, HCW is the recommended entry point: https://aka.ms/HybridWizard')))
        }

        # ── 16. Abnahmetest / Funktionsnachweis ───────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '16. Abnahmetest und Funktionsnachweis' '16. Acceptance Testing and Functional Verification') 1))
        $null = $parts.Add((New-WdParagraph (L 'Nach Abschluss der Installation sind die folgenden Funktions- und Akzeptanztests durchzuführen und zu dokumentieren. Die Testergebnisse dienen als Nachweis für die formale Abnahme des Systems (vgl. Kapitel 1.1 Freigabe und Change-Management). Bitte Ergebnis und Datum nach jedem Test eintragen.' 'After completing the installation, the following functional and acceptance tests must be performed and documented. The test results serve as evidence for the formal acceptance of the system (cf. chapter 1.1 Sign-off and Change Management). Please enter result and date after each test.')))
        # Build OWA / ECP / EWS / Autodiscover URLs from namespace if available
        $nsBase = if ($State['Namespace']) { 'https://' + $State['Namespace'] } else { 'https://<Namespace>' }
        $null = $parts.Add((New-WdTable -Headers @((L 'Testfall' 'Test case'), (L 'Prüfpunkt' 'Check'), (L 'Ergebnis' 'Result'), (L 'Datum / Tester' 'Date / Tester')) -Rows @(
            ,@('OWA',         ('{0}/owa — Login mit Testpostfach / Login with test mailbox' -f $nsBase),                              '', '')
            ,@('ECP',         ('{0}/ecp — Admin-Login, Postfach erstellen / Admin login, create mailbox' -f $nsBase),                '', '')
            ,@('EWS',         ('{0}/ews/exchange.asmx — HTTP 200 / 401' -f $nsBase),                                                '', '')
            ,@('Autodiscover', ('{0}/autodiscover/autodiscover.xml — HTTP 200 / 401' -f $nsBase),                                   '', '')
            ,@('SMTP eingehend', (L 'Testmail an internes Postfach senden (extern → Exchange)' 'Send test mail to internal mailbox (external → Exchange)'),   '', '')
            ,@('SMTP ausgehend', (L 'Testmail vom Exchange nach extern senden' 'Send test mail from Exchange to external'),           '', '')
            ,@('MAPI/HTTP',     (L 'Outlook-Client verbinden (Autodiscover, kein TCP 135 erforderlich)' 'Connect Outlook client (Autodiscover, no TCP 135 required)'), '', '')
            ,@('ActiveSync',    (L 'Mobiles Gerät verbinden (EAS, HTTPS 443)' 'Connect mobile device (EAS, HTTPS 443)'),             '', '')
            ,@('Zertifikat',    (L 'TLS-Zertifikat gültig, kein Browser-Warning' 'TLS certificate valid, no browser warning'),       '', '')
            ,@('DAG',           (L 'DAG-Datenbankkopien-Status: alle Healthy / Mounted' 'DAG database copy status: all Healthy / Mounted'), '', '')
            ,@('Backup',        (L 'Erstes VSS-Backup erfolgreich, Logs abgeschnitten' 'First VSS backup successful, logs truncated'), '', '')
            ,@('HealthChecker',  (L 'Keine kritischen Findings (Reds)' 'No critical findings (Reds)'),                               '', '')
        )))

        # ── 17. Operative Runbooks ─────────────────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '17. Operative Runbooks' '17. Operational Runbooks') 1))
        $null = $parts.Add((New-WdParagraph (L 'Dieses Kapitel enthält vorgefertigte Befehlssequenzen für die häufigsten operativen Aufgaben auf Exchange Server. Die Befehle sind in der Exchange Management Shell (EMS) auszuführen, sofern nicht anders angegeben. Platzhalter (<Server>, <DB> etc.) sind vor der Ausführung durch die tatsächlichen Werte zu ersetzen.' 'This chapter contains pre-built command sequences for the most common operational tasks on Exchange Server. Commands are to be executed in the Exchange Management Shell (EMS) unless otherwise stated. Placeholders (<Server>, <DB>, etc.) must be replaced with actual values before execution.')))
        $null = $parts.Add((New-WdHeading (L '17.1 DAG-Wartungsmodus' '17.1 DAG Maintenance Mode') 2))
        $null = $parts.Add((New-WdParagraph (L 'Vor Wartungsarbeiten (Patches, Hardwarearbeiten) an einem DAG-Mitglied muss der Server in den Wartungsmodus versetzt werden. Dies löst einen kontrollierten Failover aller aktiven Datenbanken auf andere DAG-Mitglieder aus und verhindert, dass während der Wartung neue Datenbanken aktiviert werden.' 'Before maintenance work (patches, hardware work) on a DAG member, the server must be placed in maintenance mode. This triggers a controlled failover of all active databases to other DAG members and prevents new databases from being activated during maintenance.')))
        $null = $parts.Add((New-WdCode 'Set-ServerComponentState <Server> -Component ServerWideOffline -State Inactive -Requester Maintenance'))
        $null = $parts.Add((New-WdCode 'Suspend-MailboxDatabaseCopy <DB>\<Server> -SuspendComment "Wartung"'))
        $null = $parts.Add((New-WdCode '# Wartungsarbeiten durchführen'))
        $null = $parts.Add((New-WdCode 'Resume-MailboxDatabaseCopy <DB>\<Server>'))
        $null = $parts.Add((New-WdCode 'Set-ServerComponentState <Server> -Component ServerWideOffline -State Active -Requester Maintenance'))
        $null = $parts.Add((New-WdHeading (L '17.2 Cumulative Update / Security Update installieren' '17.2 Install Cumulative Update / Security Update') 2))
        $null = $parts.Add((New-WdParagraph (L 'Exchange-Updates (CU und SU) müssen als lokaler Administrator oder als SYSTEM-Konto ausgeführt werden. Empfohlen wird die Ausführung über einen geplanten Task als SYSTEM (PSExec oder Task Scheduler). Vor dem Update: DAG-Wartungsmodus aktivieren, Backup erstellen, Health-Checker-Baseline sichern. Nach dem Update: Health-Checker erneut ausführen.' 'Exchange updates (CU and SU) must be executed as local administrator or SYSTEM account. Execution via a scheduled task as SYSTEM (PSExec or Task Scheduler) is recommended. Before the update: enable DAG maintenance mode, create backup, save HealthChecker baseline. After the update: run HealthChecker again.')))
        $null = $parts.Add((New-WdCode '# Als SYSTEM (PSExec): psexec -s setup.exe ...'))
        $null = $parts.Add((New-WdCode 'setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF /PrepareAllDomains'))
        $null = $parts.Add((New-WdCode 'setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF /Mode:Upgrade'))
        $null = $parts.Add((New-WdHeading (L '17.3 Zertifikatstausch' '17.3 Certificate Replacement') 2))
        $null = $parts.Add((New-WdParagraph (L 'Exchange-Zertifikate (IIS, SMTP) laufen typischerweise nach 1–3 Jahren ab. Der Tausch muss auf allen Exchange-Servern der Organisation durchgeführt werden. Das Auth-Zertifikat (OAuth) wird durch den MEAC-Scheduled-Task automatisch 60 Tage vor Ablauf erneuert und erfordert keinen manuellen Eingriff.' 'Exchange certificates (IIS, SMTP) typically expire after 1–3 years. The replacement must be performed on all Exchange servers in the organisation. The Auth certificate (OAuth) is automatically renewed 60 days before expiry by the MEAC scheduled task and does not require manual intervention.')))
        $null = $parts.Add((New-WdCode 'Import-ExchangeCertificate -FileName <pfx> -Password (ConvertTo-SecureString <pwd> -AsPlainText -Force) -Server <srv>'))
        $null = $parts.Add((New-WdCode 'Enable-ExchangeCertificate -Thumbprint <tp> -Services IIS,SMTP -Server <srv> -Confirm:$false'))
        $null = $parts.Add((New-WdHeading (L '17.4 Aktive Datenbank verschieben (Failover)' '17.4 Move Active Database (Failover)') 2))
        $null = $parts.Add((New-WdParagraph (L 'Manueller Failover einer aktiven Datenbankkopie auf ein anderes DAG-Mitglied — z. B. vor Wartungsarbeiten oder zur Lastverteilung.' 'Manual failover of an active database copy to another DAG member — e.g. before maintenance or for load balancing.')))
        $null = $parts.Add((New-WdCode 'Move-ActiveMailboxDatabase <DB> -ActivateOnServer <TargetServer> -Confirm:$false'))
        $null = $parts.Add((New-WdCode 'Get-MailboxDatabaseCopyStatus <DB>\* | Select Name, Status, CopyQueueLength, ReplayQueueLength'))
        $null = $parts.Add((New-WdHeading (L '17.5 Datenbankkopie neu erstellen (Reseed)' '17.5 Reseed Database Copy') 2))
        $null = $parts.Add((New-WdParagraph (L 'Wenn eine passive Datenbankkopie in einem DAG stark in Verzug geraten ist oder beschädigt wurde, kann sie neu erstellt (reseeded) werden. Der Reseed kopiert die aktive Datenbank vollständig auf das Ziel-DAG-Mitglied.' 'If a passive database copy in a DAG has fallen significantly behind or been corrupted, it can be reseeded. The reseed fully copies the active database to the target DAG member.')))
        $null = $parts.Add((New-WdCode 'Update-MailboxDatabaseCopy <DB>\<Server> -DeleteExistingFiles'))
        $null = $parts.Add((New-WdCode 'Get-MailboxDatabaseCopyStatus <DB>\<Server>  # Status verfolgen / monitor status'))
        $null = $parts.Add((New-WdHeading (L '17.6 Server wiederherstellen (RecoverServer)' '17.6 Recover Server') 2))
        $null = $parts.Add((New-WdParagraph (L 'Bei einem vollständigen Serverausfall ohne DAG-Redundanz kann Exchange auf einem neuen Server mit denselben Eigenschaften (Name, IP) wiederhergestellt werden. Voraussetzung: AD-Computerkonto noch vorhanden, Exchange-Datenbanken aus Backup verfügbar.' 'In case of a complete server failure without DAG redundancy, Exchange can be restored on a new server with the same properties (name, IP). Prerequisite: AD computer account still exists, Exchange databases available from backup.')))
        $null = $parts.Add((New-WdCode 'setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF /m:RecoverServer'))

        # ── 18. Offene Punkte ──────────────────────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '18. Offene Punkte' '18. Open Items') 1))
        # Comma operator prefix prevents PS 5.1 from flattening the jagged array when
        # binding to [object[][]]; without it Rows becomes a flat 15-element array.
        $null = $parts.Add((New-WdTable -Headers @('Nr.', (L 'Offener Punkt' 'Open item'), (L 'Verantwortlich' 'Owner'), (L 'Fällig am' 'Due date'), (L 'Status' 'Status')) -Rows @(
            ,@('1', '', '', '', '')
            ,@('2', '', '', '', '')
            ,@('3', '', '', '', '')
        )))

        # Write document
        $headerLabel = if ($DE) { 'EXCHANGE SERVER INSTALLATIONSDOKUMENTATION' } else { 'EXCHANGE SERVER INSTALLATION DOCUMENTATION' }
        if ($useTpl) {
            # F24: inject chapter body into customer template and fill cover page tokens.
            $tplTokens = @{
                document_body  = ($parts -join '')
                Organization   = (SafeVal $State['OrganizationName'] '')
                ServerName     = $env:COMPUTERNAME
                Scenario       = $scenario
                InstallMode    = $instMode
                Version        = ((Get-Date -Format 'yyyy-MM-dd') + ' / EXpress v' + $ScriptVersion)
                DateLong       = (Get-Date -Format 'dd.MM.yyyy')
                Author         = $author
                Company        = $company
                Classification = $classification
                HeaderLabel    = $headerLabel
                DocTitle       = $docTitle
                CoverSub       = $coverSub
            }
            Write-WdFromTemplate -TemplatePath $tplPath -OutputPath $docPath -Tokens $tplTokens
        } else {
            New-WdFile -OutputPath $docPath -BodyParts $parts.ToArray() -DocTitle $docTitle -HeaderLabel $headerLabel -LogoPath $logoFile
        }
        $State['WordDocPath'] = $docPath
        Write-MyOutput ('Word Installation Document: {0}' -f $docPath)
    }

    function Get-RBACReport {
        Write-MyOutput 'Generating RBAC role group membership report'
        $reportPath = Join-Path $State['ReportsPath'] ('{0}_EXpress_RBAC_{1}.txt' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd-HHmmss'))

        $roleGroups = @(
            'Organization Management',
            'Server Management',
            'Recipient Management',
            'Help Desk',
            'Hygiene Management',
            'Compliance Management',
            'Records Management',
            'Discovery Management',
            'Public Folder Management',
            'View-Only Organization Management'
        )

        $lines = [System.Collections.Generic.List[string]]::new()
        $lines.Add('Exchange RBAC Role Group Membership Report')
        $lines.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        $lines.Add("Server: $env:COMPUTERNAME")
        $lines.Add('-' * 60)

        foreach ($group in $roleGroups) {
            try {
                $members = @(Get-RoleGroupMember -Identity $group -ErrorAction Stop)
                $lines.Add('')
                $lines.Add("[$group]")
                if ($members.Count -gt 0) {
                    foreach ($member in $members) {
                        try {
                            $memberName = [string]$member.Name
                            $memberType = [string]$member.RecipientType
                            $lines.Add("  $memberName ($memberType)")
                        }
                        catch {
                            $lines.Add('  (could not display member)')
                        }
                    }
                }
                else {
                    $lines.Add('  (no members)')
                }
            }
            catch {
                $errMsg = if ($null -ne $_.Exception -and $_.Exception.Message) { $_.Exception.Message } else { $_.ToString() }
                $lines.Add('')
                $lines.Add("[$group] - could not retrieve: $errMsg")
            }
        }

        try {
            $lines | Set-Content -Path $reportPath -Encoding UTF8 -ErrorAction Stop
            Write-MyOutput "RBAC report saved to $reportPath"
        }
        catch {
            Write-MyWarning "Could not save RBAC report: $($_.Exception.Message)"
            $lines | ForEach-Object { Write-MyOutput $_ }
        }
    }

    function Get-OptimizationCatalog {
        # ─── Exchange Optimization Catalog ────────────────────────────────────────
        # To add a new optimization: append a hashtable to the array below.
        # Required fields:
        #   Key         – Single letter (A–Z) used as menu toggle key
        #   Name        – Unique identifier (used internally)
        #   Label       – Short display name shown in the menu (max 26 chars)
        #   Hint        – One-liner shown alongside the toggle (max 28 chars)
        #   Description – Full explanation shown in the description panel
        #   Default     – $true = selected by default, $false = opt-in
        #   Action      – ScriptBlock executed when the optimization is applied
        # Optional:
        #   Condition   – ScriptBlock returning $true if this entry is applicable.
        #                 If omitted the entry is always shown.
        # ──────────────────────────────────────────────────────────────────────────
        return @(
            @{
                Key         = 'A'
                Name        = 'ModernAuth'
                Label       = 'Modern Authentication'
                Hint        = 'Outlook 2016+, Teams, mobile'
                Description = 'Enables OAuth2 / Modern Authentication org-wide (Set-OrganizationConfig -OAuth2ClientProfileEnabled $true). Required for Outlook 2016+, Microsoft Teams, all mobile clients and any Hybrid / Azure AD configuration. Safe to enable on all Exchange 2016 / 2019 / SE installations. Without this, Outlook falls back to Basic Auth which Microsoft is deprecating.'
                Default     = $true
                Action      = {
                    Write-MyOutput 'Enabling Modern Authentication (OAuth2)'
                    Set-OrganizationConfig -OAuth2ClientProfileEnabled $true -ErrorAction Stop -WarningAction SilentlyContinue
                }
            }
            @{
                Key         = 'B'
                Name        = 'SessionTimeout'
                Label       = 'OWA Session Timeout (6h)'
                Hint        = 'Auto-logout after inactivity'
                Description = 'Sets activity-based OWA/ECP session timeout to 6 hours (Set-OrganizationConfig -ActivityBasedAuthenticationTimeoutEnabled $true -ActivityBasedAuthenticationTimeoutInterval 06:00:00). After 6 hours of inactivity the browser session is automatically logged out. Recommended for open-plan or shared workstation environments and for compliance requirements that mandate session expiry.'
                Default     = $true
                Action      = {
                    Write-MyOutput 'Configuring OWA/ECP session timeout (6 hours inactivity)'
                    Set-OrganizationConfig -ActivityBasedAuthenticationTimeoutEnabled $true -ActivityBasedAuthenticationTimeoutInterval '06:00:00' -ErrorAction Stop -WarningAction SilentlyContinue
                }
            }
            @{
                Key         = 'C'
                Name        = 'DisableTelemetry'
                Label       = 'Disable Telemetry (CEIP)'
                Hint        = 'Privacy / DSGVO: no Watson'
                Description = 'Disables the Microsoft Customer Experience Improvement Program (CEIP) and Watson crash telemetry (Set-OrganizationConfig -CustomerFeedbackEnabled $false). Prevents Exchange from sending diagnostic and usage data to Microsoft. Recommended for environments with strict data-privacy requirements (GDPR / DSGVO) or where external telemetry is blocked by policy.'
                Default     = $true
                Action      = {
                    Write-MyOutput 'Disabling CEIP / telemetry'
                    Set-OrganizationConfig -CustomerFeedbackEnabled $false -ErrorAction Stop -WarningAction SilentlyContinue
                }
            }
            @{
                Key         = 'D'
                Name        = 'MapiHttp'
                Label       = 'MAPI over HTTP (explicit)'
                Hint        = 'Replaces legacy RPC/HTTP'
                Description = 'Explicitly enables MAPI over HTTP (Set-OrganizationConfig -MapiHttpEnabled $true). MAPI/HTTP replaces the older Outlook Anywhere (RPC/HTTP), offering faster failover, better behaviour across NAT and load balancers, and improved Outlook startup performance. Enabled by default since Exchange 2016, but explicit activation avoids edge cases after upgrades or migrations.'
                Default     = $true
                Action      = {
                    Write-MyOutput 'Enabling MAPI over HTTP'
                    Set-OrganizationConfig -MapiHttpEnabled $true -ErrorAction Stop -WarningAction SilentlyContinue
                }
            }
            @{
                Key         = 'E'
                Name        = 'MaxMessageSize'
                Label       = 'Max Message Size (150MB)'
                Hint        = 'Org-wide send/receive limit'
                Description = 'Raises the organisation-wide maximum message size to 150MB for both send and receive, and limits recipients per message to 500 (Set-TransportConfig -MaxSendSize/-MaxReceiveSize/-MaxRecipientEnvelopeLimit). The Exchange default of 25MB is often too restrictive for modern file-sharing workflows. Frontend Receive Connectors are updated consistently. Adjust to match your storage capacity and bandwidth.'
                Default     = $true
                Action      = {
                    Write-MyOutput 'Setting org-wide max message size to 150MB'
                    Set-TransportConfig -MaxSendSize 150MB -MaxReceiveSize 150MB -MaxRecipientEnvelopeLimit 500 -ErrorAction Stop -WarningAction SilentlyContinue
                    Get-ReceiveConnector | Where-Object { $_.TransportRole -eq 'FrontendTransport' } | ForEach-Object {
                        Set-ReceiveConnector -Identity $_.Identity -MaxMessageSize 150MB -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                    }
                }
            }
            @{
                Key         = 'F'
                Name        = 'MessageExpiration'
                Label       = 'Message expiration 7 days'
                Hint        = 'Delay NDRs on outage (default: 2d)'
                Description = 'Extends message expiration timeout from the default 2 days to 7 days (Set-TransportService -MessageExpirationTimeout 7.00:00:00). During an outage or connectivity loss, messages remain in the queue for up to 7 days before an NDR is generated. Recommended for environments where short outages should not immediately result in delivery failure notifications. Skipped when CopyServerConfig is active (value is imported from source server).'
                Default     = $true
                Condition   = { -not $State['CopyServerConfig'] }
                Action      = {
                    $current = (Get-TransportService -Identity $env:COMPUTERNAME).MessageExpirationTimeout
                    if ($current -ne [TimeSpan]'7.00:00:00') {
                        Write-MyOutput 'Setting message expiration timeout to 7 days'
                        Set-TransportService -Identity $env:COMPUTERNAME -MessageExpirationTimeout 7.00:00:00 -ErrorAction Stop -WarningAction SilentlyContinue
                    }
                    else {
                        Write-MyVerbose 'Message expiration timeout already set to 7 days'
                    }
                }
            }
            @{
                Key         = 'G'
                Name        = 'ConnectorBanner'
                Label       = 'Harden SMTP Banner'
                Hint        = 'Remove Exchange version info'
                Description = 'Replaces the default SMTP greeting banner on all Frontend Receive Connectors with a generic "220 Mail Service" message (Set-ReceiveConnector -Banner). The default banner discloses the exact Exchange version, which helps attackers identify applicable CVEs. This is a low-effort hardening step recommended by security benchmarks (CIS, DISA STIG).'
                Default     = $true
                Action      = {
                    Write-MyOutput 'Hardening SMTP banner on Frontend Receive Connectors'
                    Get-ReceiveConnector | Where-Object { $_.TransportRole -eq 'FrontendTransport' } | ForEach-Object {
                        Set-ReceiveConnector -Identity $_.Identity -Banner '220 Mail Service' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                    }
                }
            }
            @{
                Key         = 'H'
                Name        = 'HtmlNDR'
                Label       = 'HTML Non-Delivery Reports'
                Hint        = 'Readable bounce messages'
                Description = 'Configures Exchange to generate HTML-formatted Non-Delivery Reports for both internal and external messages (Set-TransportConfig -InternalDsnSendHtml $true -ExternalDsnSendHtml $true). Plain-text NDRs are difficult for end users to interpret. HTML NDRs include formatted error descriptions and suggested next steps, reducing helpdesk escalations.'
                Default     = $true
                Action      = {
                    Write-MyOutput 'Enabling HTML-formatted Non-Delivery Reports'
                    Set-TransportConfig -InternalDsnSendHtml $true -ExternalDsnSendHtml $true -ErrorAction Stop -WarningAction SilentlyContinue
                }
            }
            @{
                Key         = 'I'
                Name        = 'ShadowRedundancy'
                Label       = 'Shadow Redundancy (DAG)'
                Hint        = 'Prefer remote shadow copy'
                Description = 'Configures Shadow Message Redundancy to prefer a remote DAG member as the shadow server (Set-TransportConfig -ShadowMessagePreferenceSetting PreferRemote). In a DAG, this ensures the redundant copy of each in-flight message is held on a different physical server than the primary, improving resilience against single-server failure during transport. Only effective in a DAG deployment.'
                Default     = $false
                Condition   = { $State['DAGName'] }
                Action      = {
                    Write-MyOutput 'Configuring Shadow Redundancy to prefer remote DAG member'
                    Set-TransportConfig -ShadowMessagePreferenceSetting PreferRemote -ErrorAction Stop -WarningAction SilentlyContinue
                }
            }
            @{
                Key         = 'J'
                Name        = 'SafetyNet'
                Label       = 'Safety Net Hold Time (2d)'
                Hint        = 'Explicit redelivery hold time'
                Description = 'Explicitly sets the Safety Net message hold time to 2 days (Set-TransportConfig -SafetyNetHoldTime 2.00:00:00). Safety Net retains a redundant copy of successfully delivered messages, enabling redelivery after a database failure or mailbox switchover. The 2-day default is appropriate for most environments; adjust to match your backup and recovery SLA.'
                Default     = $true
                Action      = {
                    Write-MyOutput 'Setting Safety Net hold time to 2 days'
                    Set-TransportConfig -SafetyNetHoldTime '2.00:00:00' -ErrorAction Stop -WarningAction SilentlyContinue
                }
            }
        )
    }

    function Invoke-SingleOptimization {
        param($Opt)
        try {
            & $Opt.Action
        }
        catch {
            Write-MyWarning ('Optimization [{0}] {1} failed: {2}' -f $Opt.Key, $Opt.Label, $_.Exception.Message)
        }
    }

    function Invoke-ExchangeOptimizations {
        $catalog    = Get-OptimizationCatalog
        $applicable = @($catalog | Where-Object { -not $_.ContainsKey('Condition') -or (& $_.Condition) })

        if ($applicable.Count -eq 0) {
            Write-MyVerbose 'No applicable Exchange org/transport optimizations for this configuration'
            return
        }

        # Selection state: Key -> bool
        $sel = @{}
        foreach ($opt in $applicable) { $sel[$opt.Key] = $opt.Default }

        # ── Autopilot / non-interactive: apply defaults without menu ──────────
        if ($State['Autopilot'] -or -not [Environment]::UserInteractive) {
            $defaults = @($applicable | Where-Object { $sel[$_.Key] })
            Write-MyOutput ('Applying Exchange optimizations — {0} of {1} selected (defaults)' -f $defaults.Count, $applicable.Count)
            foreach ($opt in $defaults) { Invoke-SingleOptimization $opt }
            return
        }

        # ── Interactive menu ──────────────────────────────────────────────────
        $byKey    = @{}
        foreach ($opt in $applicable) { $byKey[$opt.Key] = $opt }
        $keys     = @($applicable | ForEach-Object { $_.Key })
        $half     = [Math]::Ceiling($keys.Count / 2)
        $lastKey  = ''
        $statusMsg = ''

        $useRawKey = $false
        try { $null = $host.UI.RawUI.KeyAvailable; $useRawKey = $true } catch { }

        function Draw-OptimizationMenu {
            param([string]$Status = '', [string]$LastKey = '')
            Clear-Host
            Write-Host ('=' * 62) -ForegroundColor Cyan
            Write-Host ('  EXpress v{0} — Exchange Optimizations' -f $script:ScriptVersion) -ForegroundColor Cyan
            Write-Host ('=' * 62) -ForegroundColor Cyan
            Write-Host ''
            Write-Host '  Toggle optimizations to apply in Phase 5:' -ForegroundColor Yellow
            Write-Host ''

            # Two-column toggle list
            for ($r = 0; $r -lt $half; $r++) {
                $lk = $keys[$r]
                $rk = if (($r + $half) -lt $keys.Count) { $keys[$r + $half] } else { $null }
                $lo = $byKey[$lk]

                $lv   = if ($sel[$lk]) { 'X' } else { ' ' }
                $lStr = '  [{0}] [{1}] {2,-26}' -f $lk, $lv, $lo.Label

                $lColor = if ($lk -eq $LastKey) { [System.ConsoleColor]::Yellow } else { [System.ConsoleColor]::White }
                Write-Host $lStr -ForegroundColor $lColor -NoNewline

                if ($rk) {
                    $ro    = $byKey[$rk]
                    $rv    = if ($sel[$rk]) { 'X' } else { ' ' }
                    $rStr  = '   [{0}] [{1}] {2,-26}' -f $rk, $rv, $ro.Label
                    $rColor = if ($rk -eq $LastKey) { [System.ConsoleColor]::Yellow } else { [System.ConsoleColor]::White }
                    Write-Host $rStr -ForegroundColor $rColor
                } else {
                    Write-Host ''
                }
            }

            Write-Host ''

            # Description panel — shows full description of last-toggled option
            Write-Host ('  ' + [string]::new([char]0x2500, 58)) -ForegroundColor DarkGray
            if ($LastKey -and $byKey.ContainsKey($LastKey)) {
                $opt  = $byKey[$LastKey]
                $optState = if ($sel[$LastKey]) { 'ENABLED' } else { 'DISABLED' }  # NOTE: not $state — shadows outer $State hashtable
                Write-Host ('  [{0}] {1}  ({2})' -f $LastKey, $opt.Label, $optState) -ForegroundColor Yellow
                Write-Host ''
                # Word-wrap description at 58 chars
                $words  = ($opt.Description -replace '\s+', ' ').Trim() -split ' '
                $line   = '  '
                foreach ($w in $words) {
                    if (($line + $w).Length -gt 60) {
                        Write-Host $line
                        $line = '  ' + $w + ' '
                    }
                    else { $line += $w + ' ' }
                }
                if ($line.Trim()) { Write-Host $line }
            }
            else {
                Write-Host '  Press a letter key to see a detailed description.' -ForegroundColor DarkGray
            }
            Write-Host ('  ' + [string]::new([char]0x2500, 58)) -ForegroundColor DarkGray
            Write-Host ''

            if ($Status) { Write-Host "  $Status" -ForegroundColor Yellow; Write-Host '' }
        }

        while ($true) {
            Draw-OptimizationMenu -Status $statusMsg -LastKey $lastKey
            $statusMsg = ''

            if ($useRawKey) {
                Write-Host ('  Press {0} to toggle  |  ENTER = apply  |  S = skip all: ' -f ($keys -join '/')) -NoNewline -ForegroundColor Cyan
                try {
                    $keyInfo = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                    $vk  = $keyInfo.VirtualKeyCode
                    $raw = $keyInfo.Character.ToString().ToUpper()
                    Write-Host $raw
                    if ($vk -eq 13)    { break }        # Enter = apply
                    if ($vk -eq 27 -or $raw -eq 'S') { Write-MyOutput 'Exchange optimizations skipped'; return }
                }
                catch {
                    $useRawKey = $false
                    $raw = (Read-Host '').Trim().ToUpper()
                    if ($raw -eq '')   { break }
                    if ($raw -eq 'S')  { Write-MyOutput 'Exchange optimizations skipped'; return }
                }
            }
            else {
                $raw = (Read-Host ('  Toggle [{0}]  |  ENTER = apply  |  S = skip all' -f ($keys -join '/'))).Trim().ToUpper()
                if ($raw -eq '')  { break }
                if ($raw -eq 'S') { Write-MyOutput 'Exchange optimizations skipped'; return }
            }

            if ($raw.Length -eq 1 -and $byKey.ContainsKey($raw)) {
                $sel[$raw] = -not $sel[$raw]
                $lastKey   = $raw
            }
            elseif ($raw.Length -gt 0) {
                $statusMsg = "Unknown key '$raw' — use the listed letters, ENTER or S"
            }
        }

        # Apply selected optimizations
        $applied = 0
        foreach ($opt in $applicable | Where-Object { $sel[$_.Key] }) {
            Invoke-SingleOptimization $opt
            $applied++
        }
        Write-MyOutput ('{0} of {1} Exchange optimization(s) applied' -f $applied, $applicable.Count)
    }

    function Install-PendingWindowsUpdates {
        # Installs pending Windows security and critical updates.
        # Interactive mode: prompts per update (Y/N/A=all/S=skip rest).
        # Autopilot mode:   installs all without prompting.
        # Download + install runs in a background job with $WU_DOWNLOAD_TIMEOUT_SEC timeout;
        # on timeout the update step is skipped and Exchange installation continues.
        # Uses PSWindowsUpdate module when available; falls back to Windows Update Agent COM API.
        # Sets $State['RebootRequired'] = $true when a reboot is needed after updates.

        if (-not $State['InstallWindowsUpdates']) {
            Write-MyVerbose 'InstallWindowsUpdates not set, skipping Windows Update check'
            return
        }

        # Interactive prompts whenever a real console is available.
        # Autopilot does NOT suppress the prompt — if someone is at the keyboard they can still
        # review each update. In a truly headless run [Environment]::UserInteractive is $false.
        $isInteractive = [Environment]::UserInteractive

        Write-MyOutput 'Checking for pending Windows Updates (Security + Critical)'

        # --- Detect PSWindowsUpdate module ---
        $useModule = $false
        if (Get-Module -ListAvailable -Name PSWindowsUpdate -ErrorAction SilentlyContinue) {
            $useModule = $true
        }
        else {
            Write-MyVerbose 'PSWindowsUpdate module not found, attempting to install from PSGallery'
            try {
                # Ensure NuGet provider present unattended — without this Install-Module
                # prompts interactively even in non-interactive/Autopilot sessions.
                # Install-PackageProvider may fail to reach the provider index URI but
                # Install-Module -ForceBootstrap handles NuGet bootstrap itself without prompting.
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null
                Install-Module -Name PSWindowsUpdate -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                $useModule = $true
                Write-MyOutput 'PSWindowsUpdate module installed'
            }
            catch {
                Write-MyWarning ('Could not install PSWindowsUpdate: {0}. Falling back to WUA COM API' -f $_.Exception.Message)
            }
        }

        # --- Phase 1: Search (fast, runs in main thread) ---
        $candidates = @()   # [PSCustomObject]@{ Title; KB; Severity }

        if ($useModule) {
            try {
                Import-Module PSWindowsUpdate -ErrorAction Stop
                $wuList = Get-WindowsUpdate -Category 'Security Updates','Critical Updates' -NotTitle 'Preview' -ErrorAction Stop
                $candidates = @($wuList | ForEach-Object {
                    [PSCustomObject]@{ Title = $_.Title; KB = $_.KB; Severity = $_.MsrcSeverity }
                })
            }
            catch {
                Write-MyWarning ('PSWindowsUpdate search error: {0}' -f $_.Exception.Message)
            }
        }
        else {
            try {
                $wuaSession  = New-Object -ComObject Microsoft.Update.Session
                $wuaSearcher = $wuaSession.CreateUpdateSearcher()
                $wuaResult   = $wuaSearcher.Search('IsInstalled=0 and IsHidden=0 and BrowseOnly=0')
                $candidates  = @(foreach ($u in $wuaResult.Updates) {
                    if ($u.MsrcSeverity -in @('Critical','Important') -or $u.AutoSelectOnWebSites) {
                        [PSCustomObject]@{ Title = $u.Title; KB = ($u.KBArticleIDs | Select-Object -First 1); Severity = $u.MsrcSeverity }
                    }
                })
            }
            catch {
                Write-MyWarning ('WUA COM API search error: {0}' -f $_.Exception.Message)
            }
        }

        if ($candidates.Count -eq 0) {
            Write-MyOutput 'No pending Windows security/critical updates found'
            return
        }

        Write-MyOutput ('{0} update(s) found' -f $candidates.Count)

        # --- Phase 2: Per-update prompt ---
        # Autopilot auto-approves only when AutoApproveWindowsUpdates is explicitly set in
        # the Advanced Configuration — security updates are a deliberate opt-in decision.
        $approvedKBs     = @()
        $autoApproveAll  = (-not $isInteractive) -and $State['AutoApproveWindowsUpdates']

        if ((-not $isInteractive) -and (-not $State['AutoApproveWindowsUpdates'])) {
            Write-MyWarning ('Found {0} pending Windows update(s) — skipping in Autopilot because AutoApproveWindowsUpdates is not set. Enable it in Advanced Configuration to install automatically.' -f $candidates.Count)
            $candidates | ForEach-Object { Write-MyVerbose ('  Pending: {0} ({1})' -f $_.Title, $_.Severity) }
            return
        }

        for ($idx = 0; $idx -lt $candidates.Count; $idx++) {
            $u = $candidates[$idx]
            $label = '[{0}/{1}] {2} — {3}' -f ($idx + 1), $candidates.Count, $u.Title, $(if ($u.Severity) { $u.Severity } else { 'Unknown' })

            if ($autoApproveAll) {
                Write-MyOutput ('Auto-approved: {0}' -f $label)
                if ($u.KB) { $approvedKBs += $u.KB }
                continue
            }

            # Timed prompt: auto-skip (N) after 120 seconds with no keypress.
            # Uses RawUI.ReadKey so no Enter is required; falls back to Read-Host
            # (blocking, no timeout) when console is unavailable (redirected stdin).
            $WU_PROMPT_TIMEOUT_SEC = 120
            Write-Host ('{0}' -f $label) -ForegroundColor Cyan
            $ans = ''
            if ($host.UI.RawUI -and $host.UI.RawUI.KeyAvailable -ne $null) {
                # Flush any buffered keystrokes (e.g. from credential prompts or prior Read-Host
                # calls) so a stale keystroke doesn't immediately resolve the prompt as 'N'.
                try { $host.UI.RawUI.FlushInputBuffer() } catch { }
                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                Write-Host ('  Install? [Y/N/S=skip remaining] (auto-No in {0}s) ' -f $WU_PROMPT_TIMEOUT_SEC) -NoNewline -ForegroundColor DarkCyan
                while ($sw.Elapsed.TotalSeconds -lt $WU_PROMPT_TIMEOUT_SEC) {
                    $secsLeft = [int]($WU_PROMPT_TIMEOUT_SEC - $sw.Elapsed.TotalSeconds)
                    Write-Progress -Id 2 -Activity 'Windows Update' `
                        -Status ('Auto-No in {0}s  |  Y = install  |  N = skip  |  S = skip remaining' -f $secsLeft) `
                        -PercentComplete ($sw.Elapsed.TotalSeconds * 100 / $WU_PROMPT_TIMEOUT_SEC)
                    if ($host.UI.RawUI.KeyAvailable) {
                        $key = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                        $ans = $key.Character.ToString().ToUpper()
                        Write-Host $ans
                        break
                    }
                    Start-Sleep -Milliseconds 200
                }
                Write-Progress -Id 2 -Activity 'Windows Update' -Completed
                if ($ans -eq '') {
                    Write-Host 'N (timeout)'
                    $ans = 'N'
                }
            }
            else {
                $ans = (Read-Host '  Install? [Y=yes / N=no / S=skip remaining] (default: Y)').Trim().ToUpper()
                if ($ans -eq '') { $ans = 'Y' }
            }
            switch ($ans) {
                'S' { Write-MyOutput 'Skipping all remaining updates'; $idx = $candidates.Count; continue }
                'N' { Write-MyOutput ('Skipping: {0}' -f $u.Title) }
                default { if ($u.KB) { $approvedKBs += $u.KB }; Write-MyOutput ('Approved: {0}' -f $u.Title) }
            }
        }

        if ($approvedKBs.Count -eq 0) {
            Write-MyOutput 'No updates approved for installation — skipping Windows Update step'
            return
        }

        # --- Phase 3: Download + Install in background job with timeout ---
        Write-MyOutput ('Installing {0} approved update(s) (timeout: {1}s) ...' -f $approvedKBs.Count, $WU_DOWNLOAD_TIMEOUT_SEC)

        if ($useModule) {
            $wuJob = Start-Job -ScriptBlock {
                param([string[]]$kbs)
                Import-Module PSWindowsUpdate -ErrorAction Stop
                $result = Install-WindowsUpdate -KBArticleID $kbs -AcceptAll -IgnoreReboot -ErrorAction Stop
                $result | Select-Object Title, KB, Result, RebootRequired
            } -ArgumentList (,$approvedKBs)
        }
        else {
            $wuJob = Start-Job -ScriptBlock {
                param([string[]]$kbs)
                $session    = New-Object -ComObject Microsoft.Update.Session
                $searcher   = $session.CreateUpdateSearcher()
                # KBArticleID is not a valid WUA search criterion — search all pending and filter in-memory
                $allPending = $searcher.Search('IsInstalled=0 and IsHidden=0 and BrowseOnly=0')
                $toInstall  = New-Object -ComObject Microsoft.Update.UpdateColl
                foreach ($u in $allPending.Updates) {
                    foreach ($kb in @($u.KBArticleIDs)) {
                        if ($kbs -contains $kb) { $null = $toInstall.Add($u); break }
                    }
                }
                if ($toInstall.Count -eq 0) { return @{ Installed=0; RebootRequired=$false; ResultCode='' } }
                $dl = $session.CreateUpdateDownloader()
                $dl.Updates = $toInstall
                $dl.Download() | Out-Null
                $inst = $session.CreateUpdateInstaller()
                $inst.Updates = $toInstall
                $instResult = $inst.Install()
                @{ Installed = $toInstall.Count; RebootRequired = $instResult.RebootRequired; ResultCode = $instResult.ResultCode }
            } -ArgumentList (,$approvedKBs)
        }

        # --- Polling loop: show progress + allow keyboard cancellation ---
        $pollInterval = 5   # seconds between status checks
        $elapsed      = 0
        $cancelled    = $false
        Write-Host '  Press X to cancel Windows Update installation at any time.' -ForegroundColor DarkGray

        while ($wuJob.State -eq 'Running') {
            Start-Sleep -Seconds $pollInterval
            $elapsed += $pollInterval

            $remaining  = $WU_DOWNLOAD_TIMEOUT_SEC - $elapsed
            $pct        = [Math]::Min(99, [int]($elapsed * 100 / $WU_DOWNLOAD_TIMEOUT_SEC))
            $statusText = 'Installing {0} update(s) — {1}s elapsed  |  auto-abort in {2}s  |  X = cancel' -f $approvedKBs.Count, $elapsed, $remaining
            Write-Progress -Activity 'Windows Updates' -Status $statusText -PercentComplete $pct

            # Non-blocking key check for cancellation
            if ([Console]::KeyAvailable) {
                $key = [Console]::ReadKey($true)
                if ($key.Key -in @([ConsoleKey]::X, [ConsoleKey]::Q)) {
                    Write-Progress -Activity 'Windows Updates' -Completed
                    Write-MyWarning 'Windows Update installation cancelled by user — continuing Exchange installation without updates'
                    Stop-Job  $wuJob -ErrorAction SilentlyContinue
                    Remove-Job $wuJob -Force -ErrorAction SilentlyContinue
                    $cancelled = $true
                    break
                }
            }

            if ($elapsed -ge $WU_DOWNLOAD_TIMEOUT_SEC) {
                Write-Progress -Activity 'Windows Updates' -Completed
                Stop-Job  $wuJob -ErrorAction SilentlyContinue
                Remove-Job $wuJob -Force -ErrorAction SilentlyContinue
                Write-MyWarning ('Windows Update timed out after {0}s — continuing Exchange installation without updates' -f $WU_DOWNLOAD_TIMEOUT_SEC)
                $cancelled = $true
                break
            }
        }
        Write-Progress -Activity 'Windows Updates' -Completed
        if ($cancelled) { return }

        $jobOut    = Receive-Job $wuJob -ErrorVariable wuErrors
        Remove-Job $wuJob -Force -ErrorAction SilentlyContinue

        if ($wuErrors) {
            Write-MyWarning ('Windows Update error: {0}' -f $wuErrors[0].Exception.Message)
        }

        $rebootNeeded = $false
        if ($useModule) {
            $installed    = @($jobOut | Where-Object { $_.Result -eq 'Installed' -and $_.KB -and ($approvedKBs -contains $_.KB) }).Count
            $rebootNeeded = ($jobOut | Where-Object { $_.RebootRequired }) -as [bool]
            Write-MyOutput ('{0} update(s) installed' -f $installed)
        }
        else {
            $rebootNeeded = $jobOut.RebootRequired
            Write-MyOutput ('{0} update(s) installed, WUA result code: {1}' -f $jobOut.Installed, $jobOut.ResultCode)
        }

        if ($rebootNeeded) {
            Write-MyWarning 'Windows Updates require a reboot'
            $State['RebootRequired'] = $true
        }
        else {
            Write-MyOutput 'Windows Updates installed, no reboot required'
        }
    }

    # Known Exchange Security Updates (SU): hashtable of SetupVersion -> SU info
    # Format: @{ '<ExSetup build>' = @{ KB='KBxxxxxxx'; URL='<msp url>'; TargetVersion='<build after SU>' } }
    # Maps RTM setup.exe version -> latest known Security Update.
    # Keys are ExSetup.exe ProductVersion strings (from Get-DetectedFileVersion on setup.exe).
    # FileName must be the .exe installer name; URL must be a direct download link.
    # Update this table whenever Microsoft releases a new Exchange Security Update.
    $ExchangeSUMap = @{
        # Exchange SE RTM (15.02.2562.017) -> Feb 2026 SU (KB5074992)
        # No URL: the WU-catalog CAB is not installable via DISM/expand.exe.
        # Place ExchangeSubscriptionEdition-KB5074992-x64-en.exe (from Microsoft Download Center)
        # in <InstallPath>\sources\ before running, or apply via Windows Update / WSUS.
        '15.02.2562.017' = @{
            KB            = 'KB5074992'
            FileName      = 'ExchangeSubscriptionEdition-KB5074992-x64-en.exe'
            URL           = $null
            TargetVersion = '15.02.2562.037'
        }
        # Exchange 2019 CU15 (15.02.1748.008) -> Jan 2025 SU (KB5049233 SU3 V2)
        '15.02.1748.008' = @{
            KB            = 'KB5049233'
            FileName      = 'Exchange2019-KB5049233-x64-en.exe'
            URL           = 'https://download.microsoft.com/download/8/0/b/80b356e4-f7b1-4e11-9586-d3132a7a2fc3/Exchange2019-KB5049233-x64-en.exe'
            TargetVersion = '15.02.1748.016'
        }
        # Exchange 2019 CU14 (15.02.1544.004) -> Jan 2025 SU (KB5049233 SU3 V2)
        '15.02.1544.004' = @{
            KB            = 'KB5049233'
            FileName      = 'Exchange2019-KB5049233-x64-en.exe'
            URL           = 'https://download.microsoft.com/download/8/0/b/80b356e4-f7b1-4e11-9586-d3132a7a2fc3/Exchange2019-KB5049233-x64-en.exe'
            TargetVersion = '15.02.1544.014'
        }
        # Exchange 2019 CU13 (15.02.1258.012) -> Jan 2025 SU (KB5049233 SU7 V2)
        '15.02.1258.012' = @{
            KB            = 'KB5049233'
            FileName      = 'Exchange2019-KB5049233-x64-en.exe'
            URL           = 'https://download.microsoft.com/download/4/e/5/4e5cbbcc-5894-457d-88c4-c0b2ff7f208f/Exchange2019-KB5049233-x64-en.exe'
            TargetVersion = '15.02.1258.032'
        }
        # Exchange 2016 CU23 (15.01.2507.006) -> Jan 2025 SU (KB5049233 SU14 V2)
        '15.01.2507.006' = @{
            KB            = 'KB5049233'
            FileName      = 'Exchange2016-KB5049233-x64-en.exe'
            URL           = 'https://download.microsoft.com/download/0/9/9/0998c26c-8eb6-403a-b97a-ae44c4db5e20/Exchange2016-KB5049233-x64-en.exe'
            TargetVersion = '15.01.2507.043'
        }
    }

    function Get-LatestExchangeSecurityUpdate {
        # Returns SU info hashtable for the currently installed Exchange setup version, or $null if up to date / not applicable.
        $currentBuild = $State['SetupVersion']
        if (-not $currentBuild) { return $null }
        if ($ExchangeSUMap.ContainsKey($currentBuild)) {
            return $ExchangeSUMap[$currentBuild]
        }
        return $null
    }

    function Get-InstalledExchangeBuild {
        # Returns the installed Exchange build from the MSExchangeServiceHost service binary.
        try {
            $svcPath = (Get-CimInstance -Query 'SELECT * FROM win32_service WHERE name="MSExchangeServiceHost"' -ErrorAction Stop).PathName
            if ($svcPath) { return Get-DetectedFileVersion $svcPath.Trim('"') }
        }
        catch { }
        return $null
    }

    function Get-LatestSUBuildFromHC {
        # Parses HealthChecker.ps1's GetExchangeBuildDictionary to find the latest known SU
        # build for the installed Exchange CU. Returns a version string ('15.02.1748.043') or $null.
        $hcPath = Join-Path $State['SourcesPath'] 'HealthChecker.ps1'
        if (-not (Test-Path $hcPath)) { return $null }

        # Map setup.exe version to HC CU key
        $cuLookup = @{
            $EXSESETUPEXE_RTM    = 'RTM'
            $EX2019SETUPEXE_CU15 = 'CU15'
            $EX2019SETUPEXE_CU14 = 'CU14'
            $EX2019SETUPEXE_CU13 = 'CU13'
            $EX2016SETUPEXE_CU23 = 'CU23'
        }
        $cu = $cuLookup[$State['ExSetupVersion']]
        if (-not $cu) { return $null }

        try { $hcContent = Get-Content $hcPath -Raw -ErrorAction Stop }
        catch { return $null }

        # Find the CU block in GetExchangeBuildDictionary:
        #   "RTM"|"CUxx" = (NewCUAndSUObject "base.build" @{ "FebxxSU" = "x.x.x.x" ... })
        $cuPattern = '"' + [regex]::Escape($cu) + '"\s*=\s*\(NewCUAndSUObject\s+"[\d.]+"\s+@\{([^}]+)\}\)'
        $cuMatch   = [regex]::Match($hcContent, $cuPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
        if (-not $cuMatch.Success) { return $null }

        # Extract all SU version strings and pick the highest
        $builds = [regex]::Matches($cuMatch.Groups[1].Value, '"[\w]+"\s*=\s*"(\d+\.\d+\.\d+\.\d+)"') |
                  ForEach-Object { [System.Version]$_.Groups[1].Value } |
                  Sort-Object -Descending
        if (-not $builds -or $builds.Count -eq 0) { return $null }

        # Normalise from HC format (15.2.1748.43) to script format (15.02.1748.043)
        $b = $builds[0]
        return '{0}.{1:D2}.{2}.{3:D3}' -f $b.Major, $b.Minor, $b.Build, $b.Revision
    }

    function Install-ExchangeSecurityUpdate {
        # Downloads and installs an Exchange Security Update (.exe, .cab, or .msp).
        # P6: also does a dynamic gap-check against HealthChecker.ps1's build dictionary.
        if (-not $State['IncludeFixes']) {
            Write-MyVerbose 'IncludeFixes not set, skipping Exchange SU check'
            return
        }

        # Get the currently installed Exchange build; skip redundant reinstalls
        $installedBuild = Get-InstalledExchangeBuild
        if ($installedBuild) { Write-MyVerbose ('Installed Exchange build: {0}' -f $installedBuild) }

        $su = Get-LatestExchangeSecurityUpdate
        # B15: skip if we already installed this exact KB in a previous phase-5 run.
        # Exchange SU installers may trigger their own system reboot before the script's
        # phase-end logic runs (Enable-RunOnce / LastSuccessfulPhase update). On the next
        # run, the build version reported by Get-InstalledExchangeBuild may still show the
        # pre-SU value (service binary cache / timing), causing an endless install loop.
        # Persisting a per-KB flag in State prevents the reinstall.
        if ($su) {
            $suFlag = 'ExchangeSUInstalled_{0}' -f $su.KB
            if ($State[$suFlag]) {
                Write-MyVerbose ('Exchange SU {0} already installed in a previous run — skipping' -f $su.KB)
                return
            }
        }

        if (-not $su) {
            Write-MyOutput 'No known Exchange Security Update applicable for this build'
        }
        else {
            $targetVer    = try { [System.Version]$su.TargetVersion } catch { $null }
            $installedVer = if ($installedBuild) { try { [System.Version]$installedBuild } catch { $null } } else { $null }

            if ($installedVer -and $targetVer -and $installedVer -ge $targetVer) {
                Write-MyOutput ('Exchange build {0} already at or above SU target {1} ({2}), skipping install' -f $installedBuild, $su.TargetVersion, $su.KB)
            }
            else {
                Write-MyOutput ('Exchange Security Update {0} available for build {1} -> {2}' -f $su.KB, $State['SetupVersion'], $su.TargetVersion)
                $suPath = Join-Path $State['SourcesPath'] $su.FileName
                if (-not (Test-Path $suPath)) {
                    if ($su.URL) {
                        Write-MyOutput ('Downloading {0}' -f $su.KB)
                        $null = Get-MyPackage -Package $su.KB -URL $su.URL -FileName $su.FileName -InstallPath $State['SourcesPath']
                    }
                    if (-not (Test-Path $suPath)) {
                        Write-MyWarning ('Exchange SU {0}: installer not available for automatic download.' -f $su.KB)
                        Write-MyOutput  ('  Download:  https://support.microsoft.com/help/{0}' -f ($su.KB -replace '^KB', ''))
                        Write-MyOutput  ('  Place EXE: {0}' -f $suPath)

                        # Interactive countdown — user has 5 min to place the file, then ENTER to install.
                        # Autopilot / non-interactive: skip silently (no file available, no reboot loop).
                        if ([Environment]::UserInteractive -and -not $State['ConfigDriven']) {
                            Write-MyOutput 'Place the installer, then press ENTER — or skip now with ENTER / auto-skip after 5 min:'
                            $suTotalSecs = 300
                            $suDeadline  = [DateTime]::Now.AddSeconds($suTotalSecs)
                            try {
                                try { $host.UI.RawUI.FlushInputBuffer() } catch { }
                                while ([DateTime]::Now -lt $suDeadline) {
                                    $secsLeft = [int]($suDeadline - [DateTime]::Now).TotalSeconds
                                    Write-Progress -Id 2 -Activity ('Exchange SU {0}' -f $su.KB) `
                                        -Status ('Place {0} in {1} then ENTER  |  auto-skip in {2}s' -f $su.FileName, $State['SourcesPath'], $secsLeft) `
                                        -PercentComplete ([int](($suTotalSecs - $secsLeft) * 100 / $suTotalSecs))
                                    if ($host.UI.RawUI.KeyAvailable) {
                                        $key = $host.UI.RawUI.ReadKey('IncludeKeyDown,NoEcho')
                                        Write-Host ''
                                        if ($key.VirtualKeyCode -in 13, 27) { break }
                                    }
                                    Start-Sleep -Milliseconds 100
                                }
                                Write-Progress -Id 2 -Activity ('Exchange SU {0}' -f $su.KB) -Completed
                            }
                            catch { }
                        }
                    }
                }
                if (Test-Path $suPath) {
                    Write-MyOutput ('Installing Exchange SU {0}' -f $su.KB)
                    # B15: In Autopilot mode, pre-set RunOnce + save state before launching the
                    # installer. Exchange SU installers (.exe) may call ExitWindowsEx internally
                    # and reboot the machine before this script's phase-end logic runs, leaving
                    # LastSuccessfulPhase = 4 and no RunOnce set — so the script would not
                    # auto-resume. Pre-setting RunOnce here ensures the script always restarts.
                    if ($State['Autopilot']) {
                        Disable-UAC
                        Enable-AutoLogon
                        Enable-RunOnce
                        Save-State $State
                    }
                    # Exchange SU installers only accept /passive or /silent — /norestart is not supported.
                    # Exit code 3010 = success + reboot required; handled below.
                    $rc = Invoke-Process -FilePath $State['SourcesPath'] -FileName $su.FileName -ArgumentList '/passive'
                    if ($rc -eq 0 -or $rc -eq 3010) {
                        Write-MyOutput ('Exchange SU {0} installed successfully' -f $su.KB)
                        # Persist a per-KB installed flag immediately so phase-5 re-entry after
                        # the reboot skips the SU (build version check alone is unreliable when
                        # the service binary cache has not yet been flushed after the SU reboot).
                        $State['ExchangeSUInstalled_{0}' -f $su.KB] = $true
                        Save-State $State
                        if ($rc -eq 3010) {
                            Write-MyWarning 'Exchange SU requires a reboot'
                            $State['RebootRequired'] = $true
                        }
                    }
                    else {
                        Write-MyWarning ('Exchange SU {0} install failed (exit code {1}). Try applying via Windows Update or see https://support.microsoft.com/help/{2}' -f $su.KB, $rc, ($su.KB -replace '^KB', ''))
                    }
                }
            }
        }

        # P6 — Dynamic gap-check: download HC.ps1 if not present and compare installed
        # build against HC's GetExchangeBuildDictionary (single attempt, non-blocking).
        $hcPath = Join-Path $State['SourcesPath'] 'HealthChecker.ps1'
        if (-not (Test-Path $hcPath)) {
            try {
                Write-MyVerbose 'Downloading HealthChecker.ps1 for Exchange SU version check'
                Invoke-WebDownload -Uri 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/HealthChecker.ps1' -OutFile $hcPath
            }
            catch { Write-MyVerbose ('Could not download HealthChecker.ps1 for SU check: {0}' -f $_.Exception.Message) }
        }

        $hcLatest = Get-LatestSUBuildFromHC
        if ($hcLatest) {
            $hcLatestVer  = try { [System.Version]$hcLatest } catch { $null }
            # Re-query installed build after potential SU install above
            $currentBuild = Get-InstalledExchangeBuild
            $currentVer   = if ($currentBuild) { try { [System.Version]$currentBuild } catch { $null } } else { $null }
            if ($currentVer -and $hcLatestVer) {
                if ($currentVer -lt $hcLatestVer) {
                    Write-MyWarning ('Exchange build {0} is behind latest known SU {1} (per HealthChecker). Newer SU may require ESU enrollment — see https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates for the latest update.' -f $currentBuild, $hcLatest)
                }
                else {
                    Write-MyOutput ('Exchange build {0} is current per HealthChecker (latest known: {1})' -f $currentBuild, $hcLatest)
                }
            }
        }
    }

    function Test-IsClientOS {
        # Returns $true when running on a client SKU (Windows 10/11), $false on Server
        $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
        # ProductType: 1 = Workstation/Client, 2 = Domain Controller, 3 = Server
        return ($osInfo.ProductType -eq 1)
    }

    function Install-RecipientManagementPrereqs {
        # Phase 1 of Recipient Management install: OS detection and prerequisite installation
        if (Test-IsClientOS) {
            Write-MyOutput 'Client OS detected, installing RSAT Active Directory tools via Add-WindowsCapability'
            try {
                $cap = Get-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools*' -ErrorAction Stop
                if ($cap.State -ne 'Installed') {
                    Add-WindowsCapability -Online -Name $cap.Name -ErrorAction Stop | Out-Null
                    Write-MyOutput 'RSAT ADDS tools installed'
                }
                else {
                    Write-MyOutput 'RSAT ADDS tools already installed'
                }
            }
            catch {
                Write-MyError ('Failed to install RSAT ADDS tools: {0}' -f $_.Exception.Message)
                exit $ERR_PROBLEMADDINGFEATURE
            }
        }
        else {
            Write-MyOutput 'Server OS detected, installing RSAT-ADDS via Install-WindowsFeature'
            try {
                if (-not (Get-WindowsFeature -Name 'RSAT-ADDS').Installed) {
                    Install-WindowsFeature -Name 'RSAT-ADDS' -ErrorAction Stop | Out-Null
                    Write-MyOutput 'RSAT-ADDS installed'
                }
                else {
                    Write-MyOutput 'RSAT-ADDS already installed'
                }
            }
            catch {
                Write-MyError ('Failed to install RSAT-ADDS feature: {0}' -f $_.Exception.Message)
                exit $ERR_PROBLEMADDINGFEATURE
            }
        }
    }

    function Install-RecipientManagement {
        # Phase 2 of Recipient Management install: run setup.exe /roles:ManagementTools + EMT permission script
        Write-MyVerbose 'Validating Exchange organization is reachable'
        if (-not (Test-ExchangeOrganization)) {
            Write-MyWarning 'Exchange organization not detected in Active Directory - installation may fail if AD was not prepared'
        }

        $setupExe = Join-Path $State['SourcePath'] 'setup.exe'
        if (-not (Test-Path $setupExe)) {
            Write-MyError ('Exchange setup.exe not found at {0}' -f $setupExe)
            exit $ERR_UNEXPTECTEDPHASE
        }

        Write-MyOutput 'Running Exchange setup.exe /roles:ManagementTools /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF'
        $rc = Invoke-Process -FilePath $State['SourcePath'] -FileName 'setup.exe' -ArgumentList '/mode:install /roles:ManagementTools /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF'
        if ($rc -ne 0) {
            Write-MyError ('Exchange setup returned exit code {0}' -f $rc)
            exit $ERR_UNEXPTECTEDPHASE
        }
        Write-MyOutput 'Exchange Management Tools setup completed'

        # Run CSS-Exchange Add-PermissionForEMT.ps1 if available (pre-stage in sources\).
        # This script was removed from CSS-Exchange releases; only runs if the file is pre-staged.
        $emtScript = Join-Path $State['SourcesPath'] 'Add-PermissionForEMT.ps1'
        $emtUrl = $null   # no longer available from CSS-Exchange releases
        if (Test-Path $emtScript) {
            try {
                Write-MyOutput 'Running Add-PermissionForEMT.ps1'
                & $emtScript
            }
            catch {
                Write-MyWarning ('Add-PermissionForEMT.ps1 execution failed: {0}' -f $_.Exception.Message)
            }
        }
    }

    function New-RecipientManagementShortcut {
        # Phase 3 of Recipient Management install: create desktop shortcut loading the RecipientManagement snapin
        try {
            $desktop = [Environment]::GetFolderPath('CommonDesktopDirectory')
            $shortcutPath = Join-Path $desktop 'Exchange Recipient Management.lnk'
            $shell = New-Object -ComObject WScript.Shell
            $shortcut = $shell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = (Get-Command powershell.exe).Source
            $shortcut.Arguments = '-NoExit -Command "Add-PSSnapin *RecipientManagement; Write-Host ''Recipient Management snap-in loaded'' -ForegroundColor Green"'
            $shortcut.IconLocation = '%SystemRoot%\System32\dsa.msc, 0'
            $shortcut.Description = 'Exchange Recipient Management PowerShell'
            $shortcut.Save()
            Write-MyOutput ('Desktop shortcut created: {0}' -f $shortcutPath)
        }
        catch {
            Write-MyWarning ('Could not create desktop shortcut: {0}' -f $_.Exception.Message)
        }
    }

    function Invoke-RecipientManagementADCleanup {
        # Optional AD cleanup after Recipient Management upgrade install
        Write-MyOutput 'RecipientMgmtCleanup requested - reviewing legacy Exchange permissions'
        Write-MyWarning 'AD cleanup is a manual safety gate. Review the following and run required Set-ADPermission commands manually if desired.'
        Write-MyOutput 'Reference: https://learn.microsoft.com/en-us/exchange/plan-and-deploy/post-installation-tasks/post-installation-tasks'
    }

    function Install-ManagementToolsPrereqs {
        # Phase 1 of Management Tools install: Windows prerequisites
        Write-MyOutput 'Installing Windows prerequisites for Exchange Management Tools'
        if (Test-IsClientOS) {
            Write-MyError 'Exchange Management Tools setup requires a Windows Server OS. Use -InstallRecipientManagement for client OS installs.'
            exit $ERR_UNEXPECTEDOS
        }
        $features = @('RSAT-ADDS', 'NET-Framework-45-Features')
        foreach ($f in $features) {
            if (-not (Get-WindowsFeature -Name $f -ErrorAction SilentlyContinue).Installed) {
                try {
                    Install-WindowsFeature -Name $f -ErrorAction Stop | Out-Null
                    Write-MyOutput ('Installed Windows feature: {0}' -f $f)
                }
                catch {
                    Write-MyWarning ('Could not install {0}: {1}' -f $f, $_.Exception.Message)
                }
            }
        }
    }

    function Install-ManagementToolsRuntimePrereqs {
        # Phase 2 of Management Tools install: runtime prerequisites (VC++, URL Rewrite)
        Write-MyOutput 'Installing runtime prerequisites for Exchange Management Tools'
        # Management Tools only needs the baseline runtimes, not the full Exchange server stack.
        # Reuse existing VC++ helper functions where applicable (Install-MyPackage with the same IDs).
        Write-MyVerbose 'VC++ and URL Rewrite prerequisites are pulled in by setup.exe /roles:ManagementTools on demand'
    }

    function Install-ManagementToolsOnly {
        # Phase 3 of Management Tools install: run setup /roles:ManagementTools
        $setupExe = Join-Path $State['SourcePath'] 'setup.exe'
        if (-not (Test-Path $setupExe)) {
            Write-MyError ('Exchange setup.exe not found at {0}' -f $setupExe)
            exit $ERR_UNEXPTECTEDPHASE
        }
        Write-MyOutput 'Running Exchange setup.exe /roles:ManagementTools /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF'
        $rc = Invoke-Process -FilePath $State['SourcePath'] -FileName 'setup.exe' -ArgumentList '/mode:install /roles:ManagementTools /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF'
        if ($rc -ne 0) {
            Write-MyError ('Exchange setup returned exit code {0}' -f $rc)
            exit $ERR_UNEXPTECTEDPHASE
        }
        Write-MyOutput 'Exchange Management Tools installed successfully'
    }

    function Cleanup {
        Write-MyOutput "Cleaning up .."

        if ( (Get-WindowsFeature -Name 'Bits').Installed) {
            Write-MyOutput "Removing BITS feature"
            Remove-WindowsFeature Bits
        }
        Write-MyVerbose "Removing state file $Statefile"
        Remove-Item $Statefile
    }

    function Write-PhaseProgress {
        # Lightweight wrapper: Write-Progress for phase-level and step-level feedback.
        # Id 0 = overall install progress (Phase X of 6)
        # Id 1 = current-phase step progress (used in Phase 5 only)
        # PS2Exe does not render Write-Progress visually — fall back to Write-MyOutput for
        # meaningful milestones so progress is still visible in the console window.
        param(
            [int]$Id = 0,
            [string]$Activity,
            [string]$Status,
            [int]$PercentComplete = -1,
            [switch]$Completed
        )
        if ($Completed) {
            Write-Progress -Id $Id -Activity $Activity -Completed
        }
        elseif ($PercentComplete -ge 0) {
            Write-Progress -Id $Id -Activity $Activity -Status $Status -PercentComplete $PercentComplete
        }
        else {
            Write-Progress -Id $Id -Activity $Activity -Status $Status
        }

        # PS2Exe fallback: emit status as plain output so progress is not lost
        if ($IsPS2Exe -and -not $Completed -and $Status) {
            if ($Id -eq 0) {
                # Phase-level: only log when status changes (major transitions)
                Write-MyOutput ('[{0}] {1}' -f $Activity, $Status)
            }
            elseif ($Id -eq 1) {
                # Step-level (Phase 5): log every step
                Write-MyOutput ('  -> {0}' -f $Status)
            }
        }
    }

    function LockScreen {
        Write-MyVerbose 'Locking system'
        rundll32.exe user32.dll, LockWorkStation
    }

    function Clear-DesktopBackground {
        # Remove the desktop wallpaper during install — reduces visual distraction and
        # avoids Windows trying to render/cache wallpaper images while setup runs.
        # No restore needed: the server reboots multiple times during installation.
        # Uses registry + RUNDLL32 to avoid slow Add-Type/C# compilation on each phase start.
        Write-MyVerbose 'Clearing desktop background'
        Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name Wallpaper -Value '' -ErrorAction SilentlyContinue
        Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name WallpaperStyle -Value '0' -ErrorAction SilentlyContinue
        $p = Start-Process -FilePath 'RUNDLL32.EXE' -ArgumentList 'user32.dll, UpdatePerUserSystemParameters' -NoNewWindow -Wait -PassThru -ErrorAction SilentlyContinue
        if ($p -and $p.ExitCode -ne 0) {
            Write-MyWarning "RUNDLL32 UpdatePerUserSystemParameters exited with code $($p.ExitCode)"
        }
    }

    function Enable-HighPerformancePowerPlan {
        Write-MyVerbose 'Configuring Power Plan'
        $CurrentPlan = Get-CimInstance -Namespace root/cimv2/power -ClassName Win32_PowerPlan | Where-Object { $_.IsActive }
        if ($CurrentPlan.InstanceID -match $POWERPLAN_HIGH_PERFORMANCE) {
            Write-MyVerbose 'High Performance power plan already active'
        }
        else {
            $p = Start-Process -FilePath 'powercfg.exe' -ArgumentList ('/setactive', $POWERPLAN_HIGH_PERFORMANCE) -NoNewWindow -PassThru -Wait
            if ($p.ExitCode -ne 0) {
                Write-MyWarning "powercfg /setactive exited with code $($p.ExitCode)"
            }
            $CurrentPlan = Get-CimInstance -Namespace root/cimv2/power -ClassName Win32_PowerPlan | Where-Object { $_.IsActive }
            Write-MyOutput "Power Plan active: $($CurrentPlan.ElementName)"
        }
    }

    function Disable-NICPowerManagement {
        # http://support.microsoft.com/kb/2740020
        Write-MyVerbose 'Disabling Power Management on Network Adapters'
        # Find physical adapters that are OK and are not disabled
        $NICs = Get-CimInstance -ClassName Win32_NetworkAdapter | Where-Object { $_.AdapterTypeId -eq 0 -and $_.PhysicalAdapter -and $_.ConfigManagerErrorCode -eq 0 -and $_.ConfigManagerErrorCode -ne 22 }
        foreach ( $NIC in $NICs) {
            $PNPDeviceID = ($NIC.PNPDeviceID).ToUpper()
            $NICPowerMgt = Get-CimInstance -ClassName MSPower_DeviceEnable -Namespace root/wmi | Where-Object { $_.instancename -match [regex]::escape( $PNPDeviceID) }
            if ($NICPowerMgt.Enable) {
                Set-CimInstance -InputObject $NICPowerMgt -Property @{ Enable = $false }
                $NICPowerMgt = Get-CimInstance -ClassName MSPower_DeviceEnable -Namespace root/wmi | Where-Object { $_.instancename -match [regex]::escape( $PNPDeviceID) }
                if ($NICPowerMgt.Enable) {
                    Write-MyError "Problem disabling power management on $($NIC.Name) ($PNPDeviceID)"
                }
                else {
                    Write-MyOutput "Disabled power management on $($NIC.Name) ($PNPDeviceID)"
                }
            }
            else {
                Write-MyVerbose "Power management already disabled on $($NIC.Name) ($PNPDeviceID)"
            }
        }
    }

    function Set-Pagefile {
        Write-MyVerbose 'Checking Pagefile Configuration'
        $CS = Get-CimInstance -ClassName Win32_ComputerSystem
        if ($CS.AutomaticManagedPagefile) {
            Write-MyVerbose 'System configured to use Automatic Managed Pagefile, reconfiguring'
            try {
                $InstalledMem = $CS.TotalPhysicalMemory
                if ( $State["MajorSetupVersion"] -ge $EX2019_MAJOR) {
                    # 25% of RAM
                    $DesiredSize = [int]($InstalledMem / 4 / 1MB)
                    Write-MyVerbose ('Configuring PageFile to 25% of Total Memory: {0}MB' -f $DesiredSize)
                }
                else {
                    # RAM + 10 MB, with maximum of 32GB + 10MB
                    $DesiredSize = (($InstalledMem + 10MB), (32GB + 10MB) | Measure-Object -Minimum).Minimum / 1MB
                    Write-MyVerbose ('Configuring PageFile Total Memory+10MB with maximum of 32GB+10MB: {0}MB' -f $DesiredSize)
                }
                Set-CimInstance -InputObject $CS -Property @{ AutomaticManagedPagefile = $false }
                $CPF = Get-CimInstance -ClassName Win32_PageFileSetting
                Set-CimInstance -InputObject $CPF -Property @{ InitialSize = [int]$DesiredSize; MaximumSize = [int]$DesiredSize }
                Register-ExecutedCommand -Category 'Hardening' -Command 'Set-CimInstance -ClassName Win32_ComputerSystem -Property @{AutomaticManagedPagefile=$false}'
                Register-ExecutedCommand -Category 'Hardening' -Command ("Set-CimInstance -ClassName Win32_PageFileSetting -Property @{{InitialSize={0};MaximumSize={0}}}  # {0} MB" -f [int]$DesiredSize)
            }
            catch {
                Write-MyError "Problem reconfiguring pagefile: $($_.Exception.Message)"
            }
            $CPF = Get-CimInstance -ClassName Win32_PageFileSetting
            Write-MyOutput "Pagefile set to manual, initial/maximum size: $($CPF.InitialSize)MB / $($CPF.MaximumSize)MB"
        }
        else {
            Write-MyVerbose 'Manually configured page file, skipping configuration'
        }
    }

    function Set-TCPSettings {
        $currentRPC = (Get-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows NT\RPC' -Name 'MinimumConnectionTimeout' -ErrorAction SilentlyContinue).MinimumConnectionTimeout
        if ($currentRPC -eq 120) {
            Write-MyVerbose 'RPC Timeout already set to 120 seconds'
        }
        else {
            Write-MyOutput 'Setting RPC Timeout to 120 seconds'
            Set-RegistryValue -Path 'HKLM:\Software\Policies\Microsoft\Windows NT\RPC' -Name 'MinimumConnectionTimeout' -Value 120
        }
        $currentKA = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters' -Name 'KeepAliveTime' -ErrorAction SilentlyContinue).KeepAliveTime
        if ($currentKA -eq 900000) {
            Write-MyVerbose 'Keep-Alive Timeout already set to 15 minutes'
        }
        else {
            Write-MyOutput 'Setting Keep-Alive Timeout to 15 minutes'
            Set-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters' -Name 'KeepAliveTime' -Value 900000
        }
    }

    function Disable-SMBv1 {
        Write-MyOutput 'Disabling SMBv1 protocol (security best practice)'
        try {
            $feature = Get-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -ErrorAction SilentlyContinue
            if ($feature -and $feature.State -eq 'Enabled') {
                Disable-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -NoRestart -ErrorAction Stop | Out-Null
                Write-MyVerbose 'SMBv1 Windows feature disabled'
            }
            else {
                Write-MyVerbose 'SMBv1 Windows feature already disabled or not present'
            }
        }
        catch {
            Write-MyWarning ('Problem disabling SMBv1 feature: {0}' -f $_.Exception.Message)
        }
        try {
            Set-SmbServerConfiguration -EnableSMB1Protocol $false -Force -ErrorAction Stop
            Write-MyVerbose 'SMBv1 server protocol disabled'
        }
        catch {
            Write-MyWarning ('Problem disabling SMBv1 server config: {0}' -f $_.Exception.Message)
        }
    }

    function Disable-WindowsSearchService {
        Write-MyOutput 'Disabling Windows Search service (Exchange uses own content indexing)'
        $svc = Get-Service WSearch -ErrorAction SilentlyContinue
        if ($svc) {
            if ($svc.Status -eq 'Running') {
                Stop-Service WSearch -Force -ErrorAction SilentlyContinue
            }
            Set-Service WSearch -StartupType Disabled -ErrorAction SilentlyContinue
            Write-MyVerbose 'Windows Search service disabled'
        }
        else {
            Write-MyVerbose 'Windows Search service not found'
        }
    }

    function Disable-UnnecessaryServices {
        Write-MyOutput 'Disabling unnecessary Windows services (security hardening)'
        $services = @(
            @{ Name = 'Spooler';  Desc = 'Print Spooler (PrintNightmare attack surface, CVE-2021-34527)' }
            @{ Name = 'Fax';      Desc = 'Fax service (not required on Exchange)' }
            @{ Name = 'seclogon'; Desc = 'Secondary Logon (pass-the-hash / privilege escalation vector)' }
            @{ Name = 'SCardSvr'; Desc = 'Smart Card (not required on Exchange)' }
        )
        foreach ($svc in $services) {
            $s = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
            if ($s) {
                if ($s.Status -eq 'Running') {
                    Stop-Service -Name $svc.Name -Force -ErrorAction SilentlyContinue
                    Register-ExecutedCommand -Category 'Hardening' -Command ('Stop-Service -Name {0} -Force' -f $svc.Name)
                }
                Set-Service -Name $svc.Name -StartupType Disabled -ErrorAction SilentlyContinue
                Register-ExecutedCommand -Category 'Hardening' -Command ('Set-Service -Name {0} -StartupType Disabled  # {1}' -f $svc.Name, $svc.Desc)
                Write-MyVerbose ('Disabled: {0}' -f $svc.Desc)
            }
            else {
                Write-MyVerbose ('Service not found, skipping: {0}' -f $svc.Name)
            }
        }
    }

    function Disable-ShutdownEventTracker {
        # Redundant with Event IDs 1074/6006/6008; dialog blocks unattended Autopilot reboots
        Write-MyOutput 'Disabling Shutdown Event Tracker (redundant with event log; blocks unattended reboots)'
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Reliability' -Name 'ShutdownReasonOn' -Value 0
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Reliability' -Name 'ShutdownReasonUI' -Value 0
    }

    function Disable-WDigestCredentialCaching {
        # Prevents cleartext credential storage in LSASS memory (Mimikatz mitigation)
        Write-MyOutput 'Disabling WDigest credential caching (security hardening)'
        Set-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest' -Name 'UseLogonCredential' -Value 0
    }

    function Disable-HTTP2 {
        # HTTP/2 causes known issues with Exchange MAPI/RPC connections
        Write-MyOutput 'Disabling HTTP/2 protocol (Exchange compatibility)'
        Set-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\HTTP\Parameters' -Name 'EnableHttp2Tls' -Value 0
        Set-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\HTTP\Parameters' -Name 'EnableHttp2Cleartext' -Value 0
    }

    function Disable-TCPOffload {
        # Microsoft recommends disabling TCP offload features on Exchange servers.
        # chimney=disabled was removed from netsh in WS2019 — only apply on WS2016 (build < 17763).
        Write-MyOutput 'Disabling TCP Task Offload and autotuning settings'
        try {
            $osBuild = [int](Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name CurrentBuildNumber -ErrorAction SilentlyContinue).CurrentBuildNumber
            if ($osBuild -gt 0 -and $osBuild -lt 17763) {
                Invoke-NativeCommand -FilePath 'netsh.exe' -Arguments @('int','tcp','set','global','chimney=disabled') -Tag 'netsh chimney' | Out-Null
                if ($LASTEXITCODE -ne 0) { Write-MyWarning ('netsh chimney=disabled exited with code {0}' -f $LASTEXITCODE) }
            }
            Invoke-NativeCommand -FilePath 'netsh.exe' -Arguments @('int','tcp','set','global','autotuninglevel=restricted') -Tag 'netsh autotuninglevel' | Out-Null
            if ($LASTEXITCODE -ne 0) { Write-MyWarning ('netsh autotuninglevel=restricted exited with code {0}' -f $LASTEXITCODE) }
            Set-NetOffloadGlobalSetting -TaskOffload Disabled -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            Write-MyVerbose 'TCP offload settings configured'
        }
        catch {
            Write-MyWarning ('Problem configuring TCP offload: {0}' -f $_.Exception.Message)
        }
    }

    function Test-DiskAllocationUnitSize {
        # Exchange best practice: database and log volumes should use 64KB allocation units
        Write-MyOutput 'Checking disk allocation unit sizes (64KB recommended for Exchange volumes)'
        Get-Volume | Where-Object { $_.DriveLetter -and $_.FileSystem -eq 'NTFS' } | ForEach-Object {
            $letter = $_.DriveLetter
            $auSize = $_.AllocationUnitSize
            if ($auSize -and $auSize -ne 65536) {
                Write-MyWarning ('Drive {0}: uses {1} byte allocation units (64KB/65536 recommended for Exchange database/log volumes)' -f $letter, $auSize)
            }
            else {
                Write-MyVerbose ('Drive {0}: allocation unit size OK ({1})' -f $letter, $auSize)
            }
        }
    }

    function Disable-UnnecessaryScheduledTasks {
        Write-MyOutput 'Disabling unnecessary scheduled tasks (performance optimization)'
        $tasksToDisable = @(
            '\Microsoft\Windows\Defrag\ScheduledDefrag'
        )
        foreach ($taskName in $tasksToDisable) {
            try {
                $task = Get-ScheduledTask -TaskName (Split-Path $taskName -Leaf) -TaskPath ((Split-Path $taskName -Parent) + '\') -ErrorAction SilentlyContinue
                if ($task -and $task.State -ne 'Disabled') {
                    $task | Disable-ScheduledTask | Out-Null
                    Write-MyVerbose ('Disabled scheduled task: {0}' -f $taskName)
                }
                else {
                    Write-MyVerbose ('Scheduled task already disabled or not found: {0}' -f $taskName)
                }
            }
            catch {
                Write-MyWarning ('Problem disabling scheduled task {0}: {1}' -f $taskName, $_.Exception.Message)
            }
        }
    }

    function Disable-ServerManagerAtLogon {
        # Disable Server Manager at logon for ALL users (machine-wide).
        # Three layers are used for complete coverage:
        #   1. Machine-wide Group Policy key — overrides per-user HKCU settings
        #   2. Default user hive — applies to new user profiles created after this point
        #   3. Scheduled task — belt-and-suspenders, prevents task-triggered launch
        # Idempotent: silent if all three layers are already configured.
        $policyPath    = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Server\ServerManager'
        $alreadyPolicy = (Get-ItemProperty -Path $policyPath -Name 'DoNotOpenAtLogon' -ErrorAction SilentlyContinue).DoNotOpenAtLogon -eq 1
        $smTask        = Get-ScheduledTask -TaskName 'ServerManager' -TaskPath '\Microsoft\Windows\Server Manager\' -ErrorAction SilentlyContinue
        $alreadyTask   = -not $smTask -or $smTask.State -eq 'Disabled'
        if ($alreadyPolicy -and $alreadyTask) {
            Write-MyVerbose 'Server Manager at logon already disabled — skipping'
            return
        }
        Write-MyOutput 'Disabling Server Manager at logon for all users'

        # Layer 1: Machine-wide policy (overrides HKCU for all users)
        $policyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Server\ServerManager'
        if (-not (Test-Path $policyPath -ErrorAction SilentlyContinue)) {
            New-Item -Path $policyPath -Force -ErrorAction SilentlyContinue | Out-Null
        }
        Set-RegistryValue -Path $policyPath -Name 'DoNotOpenAtLogon' -Value 1 -PropertyType DWord

        # Layer 2: Default user profile hive (new users created after this point)
        $defaultHive    = 'C:\Users\Default\NTUSER.DAT'
        $defaultHiveKey = 'HKU\ExchangeInstallDefault'
        if (Test-Path $defaultHive) {
            Invoke-NativeCommand -FilePath 'reg.exe' -Arguments @('load', $defaultHiveKey, $defaultHive) -Tag 'reg load default hive' | Out-Null
            if (Test-Path "Registry::$defaultHiveKey\Software\Microsoft\ServerManager") {
                Set-ItemProperty -Path "Registry::$defaultHiveKey\Software\Microsoft\ServerManager" -Name 'DoNotOpenServerManagerAtLogon' -Value 1 -Type DWord -ErrorAction SilentlyContinue
            }
            else {
                New-Item -Path "Registry::$defaultHiveKey\Software\Microsoft\ServerManager" -Force -ErrorAction SilentlyContinue | Out-Null
                New-ItemProperty -Path "Registry::$defaultHiveKey\Software\Microsoft\ServerManager" -Name 'DoNotOpenServerManagerAtLogon' -Value 1 -PropertyType DWord -Force -ErrorAction SilentlyContinue | Out-Null
            }
            Invoke-NativeCommand -FilePath 'reg.exe' -Arguments @('unload', $defaultHiveKey) -Tag 'reg unload default hive' | Out-Null
        }

        # Layer 3: Disable the ServerManager scheduled task (machine-wide)
        $smTask = Get-ScheduledTask -TaskName 'ServerManager' -TaskPath '\Microsoft\Windows\Server Manager\' -ErrorAction SilentlyContinue
        if ($smTask -and $smTask.State -ne 'Disabled') {
            $smTask | Disable-ScheduledTask | Out-Null
            Write-MyVerbose 'Disabled scheduled task: \Microsoft\Windows\Server Manager\ServerManager'
        }
    }

    function Set-CRLCheckTimeout {
        # Prevents Exchange startup delays when CRL endpoints are unreachable
        Write-MyOutput 'Configuring Certificate Revocation List check timeout (15 seconds)'
        $regPath = 'HKLM:\SOFTWARE\Microsoft\Cryptography\OID\EncodingType 0\CertDllCreateCertificateChainEngine\Config'
        if (-not (Test-Path $regPath -ErrorAction SilentlyContinue)) {
            New-Item -Path $regPath -Force -ErrorAction SilentlyContinue | Out-Null
        }
        Set-RegistryValue -Path $regPath -Name 'ChainUrlRetrievalTimeoutMilliseconds' -Value 15000
    }

    function Disable-CredentialGuard {
        # HealthChecker flags Credential Guard as causing performance issues on Exchange servers.
        # On Windows Server 2025 it is enabled by default and must be explicitly disabled.
        Write-MyOutput 'Disabling Credential Guard (Exchange performance best practice)'
        Set-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA' -Name 'LsaCfgFlags' -Value 0
        Set-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard' -Name 'EnableVirtualizationBasedSecurity' -Value 0
    }

    function Set-LmCompatibilityLevel {
        # HealthChecker recommends level 5: send NTLMv2 only, refuse LM and NTLM
        Write-MyOutput 'Setting LAN Manager compatibility level to 5 (NTLMv2 only)'
        Set-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' -Name 'LmCompatibilityLevel' -Value 5
    }

    function Enable-RSSOnAllNICs {
        # HealthChecker warns if RSS is disabled or queue count does not match physical core count
        Write-MyOutput 'Enabling Receive Side Scaling (RSS) on all supported NICs'
        $physicalCores = (Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue |
            Measure-Object -Property NumberOfCores -Sum).Sum
        if (-not $physicalCores -or $physicalCores -lt 1) { $physicalCores = 1 }
        Write-MyVerbose ('Physical core count: {0} — setting RSS queue count to match' -f $physicalCores)
        Register-ExecutedCommand -Category 'Hardening' -Command 'Enable-NetAdapterRss -Name *'
        Register-ExecutedCommand -Category 'Hardening' -Command ("Set-NetAdapterRss -Name * -NumberOfReceiveQueues $physicalCores  # = physical core count")
        try {
            Get-NetAdapterRss -ErrorAction SilentlyContinue | ForEach-Object {
                if (-not $_.Enabled) {
                    Write-MyVerbose ('Enabling RSS on adapter: {0}' -f $_.Name)
                    Enable-NetAdapterRss -Name $_.Name -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                }
                if ($_.NumberOfReceiveQueues -ne $physicalCores) {
                    Set-NetAdapterRss -Name $_.Name -NumberOfReceiveQueues $physicalCores -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                    Write-MyVerbose ('Set RSS queues to {0} on adapter: {1}' -f $physicalCores, $_.Name)
                }
                else {
                    Write-MyVerbose ('RSS queues already at {0} on adapter: {1}' -f $physicalCores, $_.Name)
                }
            }
        }
        catch {
            Write-MyWarning ('Problem configuring RSS: {0}' -f $_.Exception.Message)
        }
    }

    function Set-IPv4OverIPv6Preference {
        # Microsoft recommendation for Exchange: prefer IPv4 over IPv6 (DisabledComponents = 0x20).
        # Disables IPv6 on all non-loopback interfaces but keeps the IPv6 loopback intact,
        # which Exchange internal components rely on. Full IPv6 disable (0xFF) is not recommended.
        $regPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters'
        $current = (Get-ItemProperty -Path $regPath -Name DisabledComponents -ErrorAction SilentlyContinue).DisabledComponents
        if ($current -eq 0x20) {
            Write-MyVerbose 'IPv4 over IPv6 preference already set (DisabledComponents = 0x20) (OK)'
        } else {
            Set-RegistryValue -Path $regPath -Name 'DisabledComponents' -Value 0x20 -PropertyType DWord
            Write-MyOutput 'IPv4 over IPv6 preference set (DisabledComponents = 0x20) — effective after next reboot'
            # Do not flag RebootRequired: the value is re-read at boot and the install's
            # end-of-Phase-6 reboot (or the next natural reboot) activates it. Forcing a
            # mid-install reboot here would trigger the Phase 5→6 skip-logic unnecessarily.
        }
    }

    function Disable-NetBIOSOnAllNICs {
        # Disables NetBIOS over TCP/IP on all NICs. Exchange does not require NetBIOS;
        # disabling it reduces attack surface (LLMNR/NBT-NS poisoning, credential capture).
        # SetTcpipNetbios(2) = Disable NetBIOS over TCP/IP
        Write-MyOutput 'Disabling NetBIOS over TCP/IP on all NICs'
        try {
            $nics = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True' -ErrorAction Stop
            $changed = 0
            foreach ($nic in $nics) {
                $result = ($nic | Invoke-CimMethod -MethodName SetTcpipNetbios -Arguments @{ TcpipNetbiosOptions = [uint32]2 } -ErrorAction SilentlyContinue).ReturnValue
                if ($result -eq 0) {
                    Write-MyVerbose ('NetBIOS disabled on: {0}' -f $nic.Description)
                    $changed++
                } elseif ($result -eq 1) {
                    Write-MyVerbose ('NetBIOS disable on {0}: takes effect after next reboot' -f $nic.Description)
                    $changed++
                    # Do not flag RebootRequired: the setting activates on the next boot
                    # anyway (end-of-Phase-6 reboot covers it). Forcing a Phase 5→6 reboot
                    # for a NIC flag is unnecessary.
                } else {
                    Write-MyWarning ('NetBIOS disable on {0} returned code {1}' -f $nic.Description, $result)
                }
            }
            Write-MyVerbose ('NetBIOS disabled on {0} NIC(s)' -f $changed)
        } catch {
            Write-MyWarning ('Failed to disable NetBIOS: {0}' -f $_.Exception.Message)
        }
    }

    function Disable-LLMNR {
        # CIS L1 18.5.4.2: Disable Link-Local Multicast Name Resolution.
        # LLMNR broadcasts unresolved names to the local subnet; Responder-class tools
        # answer with spoofed records and capture NTLM hashes. Exchange relies on DNS.
        Write-MyOutput 'Disabling LLMNR (Link-Local Multicast Name Resolution)'
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient' -Name 'EnableMulticast' -Value 0
    }

    function Disable-MDNS {
        # WS2022+ enables mDNS by default (port 5353 UDP). Same poisoning vector as LLMNR.
        # Registry value EnableMDNS under Dnscache\Parameters disables it globally.
        Write-MyOutput 'Disabling mDNS (Multicast DNS)'
        Set-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\Dnscache\Parameters' -Name 'EnableMDNS' -Value 0
    }

    function Enable-LSAProtection {
        # Enables LSA Protection (RunAsPPL) to prevent credential theft from LSASS memory.
        # Exchange 2019 CU12+ and Exchange SE are compatible with LSA Protection.
        # Earlier Exchange versions (2016, pre-CU12 2019) may conflict with legacy auth providers.
        # The setting takes effect after the next reboot.
        $regPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'
        $current = (Get-ItemProperty -Path $regPath -Name RunAsPPL -ErrorAction SilentlyContinue).RunAsPPL
        if ($current -eq 1) {
            Write-MyVerbose 'LSA Protection (RunAsPPL) already enabled'
            return
        }
        Write-MyOutput 'Enabling LSA Protection (RunAsPPL) — effective after next reboot'
        Set-RegistryValue -Path $regPath -Name 'RunAsPPL' -Value 1 -PropertyType DWord
        # Audit mode first (2) is not used here as Exchange servers are domain-joined production systems
        # and Exchange 2019 CU12+/SE are fully compatible with RunAsPPL = 1.
    }

    function Set-MaxConcurrentAPI {
        # Netlogon MaxConcurrentApi limits simultaneous Kerberos/NTLM authentication requests
        # against domain controllers. Exchange generates heavy auth load; the default (10) can
        # cause 0xC000005E (No logon servers) errors under load on busy servers.
        # Microsoft recommendation for Exchange: raise to match logical processor count (min 10).
        # Edge Transport is not domain-joined — Netlogon optimization does not apply.
        if ($State['InstallEdge']) { Write-MyVerbose 'Set-MaxConcurrentAPI: skipped (Edge Transport)'; return }
        Write-MyOutput 'Setting Netlogon MaxConcurrentApi for Kerberos authentication optimization'
        $logicalProcs = (Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue).NumberOfLogicalProcessors
        if (-not $logicalProcs -or $logicalProcs -lt 10) { $logicalProcs = 10 }
        $regPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters'
        Register-ExecutedCommand -Category 'Hardening' -Command ("Set-ItemProperty '$regPath' MaxConcurrentApi $logicalProcs  # = logical processor count, min 10")
        Set-RegistryValue -Path $regPath -Name 'MaxConcurrentApi' -Value $logicalProcs -PropertyType DWord
        Write-MyVerbose ('MaxConcurrentApi set to {0}' -f $logicalProcs)
    }

    function Set-CtsProcessorAffinityPercentage {
        # HealthChecker flags any non-zero value as harmful to Exchange Search performance
        # Edge Transport uses a different search stack — this registry path does not exist there.
        if ($State['InstallEdge']) { Write-MyVerbose 'Set-CtsProcessorAffinityPercentage: skipped (Edge Transport)'; return }
        Write-MyOutput 'Setting CtsProcessorAffinityPercentage to 0 (Exchange Search best practice)'
        $regPath = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Search\SystemParameters'
        if (-not (Test-Path $regPath -ErrorAction SilentlyContinue)) {
            New-Item -Path $regPath -Force -ErrorAction SilentlyContinue | Out-Null
        }
        Set-RegistryValue -Path $regPath -Name 'CtsProcessorAffinityPercentage' -Value 0
    }

    function Enable-SerializedDataSigning {
        # HealthChecker validates this security feature (mitigates PowerShell serialization attacks)
        Write-MyOutput 'Enabling Serialized Data Signing (security hardening)'
        $regPath = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Diagnostics'
        if (-not (Test-Path $regPath -ErrorAction SilentlyContinue)) {
            New-Item -Path $regPath -Force -ErrorAction SilentlyContinue | Out-Null
        }
        Set-RegistryValue -Path $regPath -Name 'EnableSerializationDataSigning' -Value 1
    }

    function Set-NodeRunnerMemoryLimit {
        # HealthChecker flags any non-zero memoryLimitMegabytes as a Search performance limiter
        # Edge Transport does not run Exchange Search / NodeRunner.
        if ($State['InstallEdge']) { Write-MyVerbose 'Set-NodeRunnerMemoryLimit: skipped (Edge Transport)'; return }
        Write-MyOutput 'Setting NodeRunner memory limit to 0 (unlimited, Exchange Search best practice)'
        $exchangeInstallPath = (Get-ItemProperty -Path $EXCHANGEINSTALLKEY -Name MsiInstallPath -ErrorAction SilentlyContinue).MsiInstallPath
        if ($exchangeInstallPath) {
            $configFile = Join-Path $exchangeInstallPath 'Bin\Search\Ceres\Runtime\1.0\noderunner.exe.config'
            if (Test-Path $configFile) {
                try {
                    $xml = [XML](Get-Content $configFile)
                    $node = $xml.SelectSingleNode('//nodeRunnerSettings')
                    if ($node -and $node.memoryLimitMegabytes -ne '0') {
                        $node.memoryLimitMegabytes = '0'
                        $xml.Save($configFile)
                        Write-MyVerbose 'NodeRunner memoryLimitMegabytes set to 0'
                    }
                    else {
                        Write-MyVerbose 'NodeRunner memoryLimitMegabytes already 0 or node not found'
                    }
                }
                catch {
                    Write-MyWarning ('Problem configuring NodeRunner: {0}' -f $_.Exception.Message)
                }
            }
            else {
                Write-MyVerbose 'NodeRunner config file not found (may not be installed yet)'
            }
        }
    }

    function Enable-MAPIFrontEndServerGC {
        # HealthChecker recommends Server GC for MAPI Front End App Pool on systems with 20+ GB RAM
        Write-MyOutput 'Checking MAPI Front End App Pool GC mode'
        $installedMem = (Get-CimInstance -ClassName Win32_ComputerSystem).TotalPhysicalMemory
        if ($installedMem -ge 20GB) {
            $exchangeInstallPath = (Get-ItemProperty -Path $EXCHANGEINSTALLKEY -Name MsiInstallPath -ErrorAction SilentlyContinue).MsiInstallPath
            if ($exchangeInstallPath) {
                $configFile = Join-Path $exchangeInstallPath 'bin\MSExchangeMapiFrontEndAppPool_CLRConfig.config'
                if (Test-Path $configFile) {
                    try {
                        $xml = [XML](Get-Content $configFile)
                        $gcNode = $xml.SelectSingleNode('//gcServer')
                        if ($gcNode -and $gcNode.enabled -ne 'true') {
                            $gcNode.enabled = 'true'
                            $xml.Save($configFile)
                            Write-MyOutput 'Enabled Server GC for MAPI Front End App Pool (20+ GB RAM detected)'
                        }
                        else {
                            Write-MyVerbose 'Server GC already enabled or node not found'
                        }
                    }
                    catch {
                        Write-MyWarning ('Problem configuring MAPI FE GC: {0}' -f $_.Exception.Message)
                    }
                }
                else {
                    Write-MyVerbose 'MAPI FE config file not found (may not be installed yet)'
                }
            }
        }
        else {
            Write-MyVerbose 'Less than 20 GB RAM, skipping Server GC configuration'
        }
    }

    function Disable-SSL3 {
        # SSL3 disabling/Poodle, https://support.microsoft.com/en-us/kb/187498
        Write-MyVerbose 'Disabling SSL3 protocol for services'
        Set-RegistryValue -Path 'HKLM:\System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server' -Name 'Enabled' -Value 0
    }

    function Disable-RC4 {
        # https://support.microsoft.com/en-us/kb/2868725
        # Note: Can't use regular New-Item as registry path contains '/' (always interpreted as path splitter)
        Write-MyVerbose 'Disabling RC4 protocol for services'
        $RC4Keys = @('RC4 128/128', 'RC4 40/128', 'RC4 56/128')
        $RegKey = 'SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers'
        $RegName = "Enabled"
        foreach ( $RC4Key in $RC4Keys) {
            if ( -not( Get-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue)) {
                if ( -not (Test-Path $RegKey -ErrorAction SilentlyContinue)) {
                    $RegHandle = (Get-Item 'HKLM:\').OpenSubKey( $RegKey, $true)
                    $RegHandle.CreateSubKey( $RC4Key) | Out-Null
                    $RegHandle.Close()
                }
            }
            Write-MyVerbose "Setting registry $RegKey\$RegName\RC4Key to 0"
            New-ItemProperty -Path (Join-Path (Join-Path 'HKLM:\' $RegKey) $RC4Key) -Name $RegName -Value 0 -Force -ErrorAction SilentlyContinue | Out-Null
            Register-ExecutedCommand -Category 'Hardening' -Command ("New-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\{0}' -Name Enabled -Value 0 -Force" -f $RC4Key)
        }
    }

    function Enable-ECC {
        # https://learn.microsoft.com/en-us/exchange/architecture/client-access/certificates?view=exchserver-2019#elliptic-curve-cryptography-certificates-support-in-exchange-server
        Write-MyOutput 'Enabling Elliptic Curve Cryptography (ECC) certificate support'

        $RegKey = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Diagnostics'
        $RegName = 'EnableEccCertificateSupport'

        if ( -not( Get-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue)) {
            Write-MyVerbose ('Setting {0}\{1} to 1' -f $RegKey, $RegName)
            New-ItemProperty -Path $RegKey -Name $RegName -Value 1 -Type String -Force -ErrorAction SilentlyContinue | Out-Null
        }

        # If overrides were configured, disable these (obsolete and not fully supporting ECC)
        $Override = Get-SettingOverride | Where-Object { ($_.SectionName -eq "ECCCertificateSupport") -and ($_.Parameters -eq "Enabled=true") }
        if ( $Override) {
            Write-MyVerbose ('Override for ECC found, removing (obsolete)')
            $Override | Remove-SettingOverride
            Get-ExchangeDiagnosticInfo -Process Microsoft.Exchange.Directory.TopologyService -Component VariantConfiguration -Argument Refresh | Out-Null
            $script:p5NeedsIisRestart = $true
        }
        else {
            Write-MyVerbose ('No override configuration for ECC found')
        }
    }

    function Enable-CBC {
        # https://support.microsoft.com/en-us/topic/enable-support-for-aes256-cbc-encrypted-content-in-exchange-server-august-2023-su-add63652-ee17-4428-8928-ddc45339f99e
        Write-MyOutput 'Enabling AES256-CBC encryption mode support'

        $Override = Get-SettingOverride | Where-Object { ($_.SectionName -eq "EnableEncryptionAlgorithmCBC") -and ($_.Parameters -eq "Enabled=True") }
        if ( $Override) {
            Write-MyVerbose ('Configuration for CBC already configured')
        }
        else {
            New-SettingOverride -Name "EnableEncryptionAlgorithmCBC" -Parameters @("Enabled=True") -Component Encryption -Reason "Enable CBC encryption" -Section EnableEncryptionAlgorithmCBC | Out-Null
            Get-ExchangeDiagnosticInfo -Process Microsoft.Exchange.Directory.TopologyService -Component VariantConfiguration -Argument Refresh | Out-Null
            $script:p5NeedsIisRestart = $true
        }
    }

    function Enable-AMSI {
        param(
            [string[]]$ConfigParam = @("EnabledEcp=True", "EnabledEws=True", "EnabledOwa=True", "EnabledPowerShell=True")
        )
        # https://learn.microsoft.com/en-us/exchange/antispam-and-antimalware/amsi-integration-with-exchange?view=exchserver-2019#enable-exchange-server-amsi-body-scanning
        # Edge Transport is not domain-joined and has no org connection; New-SettingOverride would fail.
        if ($State['InstallEdge']) { Write-MyVerbose 'Enable-AMSI: skipped (Edge Transport — no org connection)'; return }
        Write-MyOutput 'Enabling AMSI body scanning for OWA, ECP, EWS and PowerShell'

        $amsiOverride = Get-SettingOverride | Where-Object { $_.SectionName -eq 'AmsiRequestBodyScanning' }
        if ($amsiOverride) {
            Write-MyVerbose 'AMSI body scanning override already configured'
        }
        else {
            New-SettingOverride -Name "EnableAMSIBodyScan" -Component Cafe -Section AmsiRequestBodyScanning -Parameters $ConfigParam -Reason "Enabling AMSI body Scan" | Out-Null
            Get-ExchangeDiagnosticInfo -Process Microsoft.Exchange.Directory.TopologyService -Component VariantConfiguration -Argument Refresh | Out-Null
            $script:p5NeedsIisRestart = $true
        }
    }

    function Enable-IanaTimeZoneMappings {
        # Exchange 2019 CU14+ ships IanaTimeZoneMappings.xml in the bin folder.
        # HealthChecker flags its absence as a calendar timezone issue.
        $setupKey = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup' -ErrorAction SilentlyContinue
        if (-not $setupKey) { Write-MyVerbose 'Enable-IanaTimeZoneMappings: Exchange setup registry key not found'; return }
        $exBin = Join-Path $setupKey.MsiInstallPath 'Bin'
        $mappingFile = Join-Path $exBin 'IanaTimeZoneMappings.xml'
        if (Test-Path $mappingFile) {
            Write-MyVerbose ('IANA timezone mappings file present: {0}' -f $mappingFile)
        }
        else {
            Write-MyWarning ('IANA timezone mappings file not found ({0}). Calendar timezone issues may occur. Update Exchange to a newer CU to resolve.' -f $mappingFile)
        }

        # Exchange 2019 CU14+ / SE: enable IANA timezone IDs for calendar items if supported
        try {
            $orgConfig = Get-OrganizationConfig -ErrorAction Stop
            if ($orgConfig.PSObject.Properties['UseIanaTimeZoneId']) {
                if (-not $orgConfig.UseIanaTimeZoneId) {
                    Register-ExecutedCommand -Category 'ExchangeTuning' -Command 'Set-OrganizationConfig -UseIanaTimeZoneId $true'
                    Set-OrganizationConfig -UseIanaTimeZoneId $true -ErrorAction Stop
                    Write-MyOutput 'IANA timezone IDs enabled for calendar items (UseIanaTimeZoneId)'
                }
                else {
                    Write-MyVerbose 'IANA timezone IDs already enabled (UseIanaTimeZoneId)'
                }
            }
            else {
                Write-MyVerbose 'UseIanaTimeZoneId not available on this Exchange version — skipping'
            }
        }
        catch {
            Write-MyVerbose ('Enable-IanaTimeZoneMappings: {0}' -f $_.Exception.Message)
        }
    }

    function Disable-SSLOffloading {
        # F13: SSL offloading at a reverse proxy prevents Extended Protection channel-binding from working.
        # Always set to $false — Exchange should terminate TLS itself, not receive plaintext from a proxy.
        if ($State['InstallEdge']) { return }
        Write-MyOutput 'Configuring Outlook Anywhere SSL offloading (required for Extended Protection)'
        try {
            $oa = Get-OutlookAnywhere -Server $env:computername -ErrorAction SilentlyContinue
            if ($oa) {
                if ($oa.SSLOffloading) {
                    Set-OutlookAnywhere -Identity $oa.Identity -SSLOffloading $false -Confirm:$false -ErrorAction Stop
                    Register-ExecutedCommand -Category 'ExchangeTuning' -Command ("Set-OutlookAnywhere -Identity '{0}' -SSLOffloading `$false" -f $oa.Identity)
                    Write-MyVerbose 'Outlook Anywhere SSL offloading disabled'
                }
                else {
                    Write-MyVerbose 'Outlook Anywhere SSL offloading already disabled (OK)'
                }
            }
            else {
                Write-MyVerbose 'No Outlook Anywhere virtual directory found on this server'
            }
        }
        catch {
            Write-MyWarning ('Could not configure Outlook Anywhere SSL offloading: {0}' -f $_.Exception.Message)
        }
    }

    function Enable-ExtendedProtection {
        # F6: Windows Extended Protection (channel binding) mitigates NTLM relay / pass-the-hash attacks on IIS.
        # Prerequisite: SSL offloading must be disabled (F13), TLS 1.2 must be enforced.
        # Exchange 2019 CU14+ / SE: EP is enabled by setup — this function validates the configuration.
        # Exchange 2016 / 2019 pre-CU14: downloads and runs ExchangeExtendedProtectionManagement.ps1 from CSS-Exchange.
        if ($State['DoNotEnableEP']) { Write-MyVerbose 'DoNotEnableEP set — skipping Extended Protection'; return }
        if ($State['InstallEdge'])   { Write-MyVerbose 'Edge Transport — Extended Protection not applicable'; return }

        $exSetupVer    = [System.Version]$State['ExSetupVersion']
        $isCU14OrNewer = $exSetupVer -ge [System.Version]$EX2019SETUPEXE_CU14

        if ($isCU14OrNewer) {
            Write-MyOutput 'Exchange 2019 CU14+ / SE — Extended Protection enabled by setup; validating OWA'
            try {
                $owa = Get-OwaVirtualDirectory -Server $env:computername -ErrorAction SilentlyContinue
                if ($owa) {
                    $ep = $owa.ExtendedProtectionTokenChecking
                    if ($ep -eq 'None') {
                        Write-MyWarning ('OWA Extended Protection is None (expected Require/Allow for Exchange {0}). Review ExtendedProtectionTokenChecking on all virtual directories.' -f $State['ExSetupVersion'])
                    }
                    else {
                        Write-MyVerbose ('OWA ExtendedProtectionTokenChecking: {0} (OK)' -f $ep)
                    }
                }
            }
            catch { Write-MyVerbose ('Extended Protection validation: {0}' -f $_.Exception.Message) }
            return
        }

        # Exchange 2016 / 2019 pre-CU14: configure via CSS-Exchange ExchangeExtendedProtectionManagement.ps1
        Write-MyOutput 'Enabling Extended Protection via CSS-Exchange ExchangeExtendedProtectionManagement.ps1'
        $epPath = Join-Path $State['SourcesPath'] 'ExchangeExtendedProtectionManagement.ps1'
        $epUrl  = 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/ExchangeExtendedProtectionManagement.ps1'
        # Note: previously named ExchangeExtendedProtection.ps1 — renamed in CSS-Exchange 2024 releases

        if (-not (Test-Path $epPath)) {
            try {
                Invoke-WebDownload -Uri $epUrl -OutFile $epPath
                Write-MyVerbose ('ExchangeExtendedProtectionManagement.ps1 downloaded, SHA256: {0}' -f (Get-FileHash $epPath -Algorithm SHA256).Hash)
            }
            catch {
                Write-MyWarning ('Could not download ExchangeExtendedProtectionManagement.ps1: {0}' -f $_.Exception.Message)
                return
            }
        }

        try {
            $epArgs    = @('-ExchangeServerNames', $env:computername)
            $epSkipEWS = if ($State['DoNotEnableEP_FEEWS']) { ' -SkipEWS' } else { '' }
            if ($epSkipEWS) { $epArgs += '-SkipEWS' }
            $epCmd = '& ExchangeExtendedProtectionManagement.ps1 -ExchangeServerNames {0}{1}' -f $env:computername, $epSkipEWS
            Register-ExecutedCommand -Category 'ExchangeTuning' -Command $epCmd
            & $epPath @epArgs *>&1 | ForEach-Object { Write-ToTranscript ([string]$_) }
        }
        catch {
            Write-MyWarning ('ExchangeExtendedProtectionManagement.ps1 failed: {0}' -f $_.Exception.Message)
        }
    }

    function Enable-RootCertificateAutoUpdate {
        # F17: Prevents Exchange Online connectivity failures caused by stale/missing root CA certificates.
        # Group Policy or hardening baselines sometimes disable Windows automatic root certificate updates,
        # which breaks connectivity to Exchange Online, Microsoft 365, and any modern PKI-dependent service.
        Write-MyOutput 'Verifying automatic root certificate update (AuthRoot policy)'
        $regPath = 'HKLM:\SOFTWARE\Policies\Microsoft\SystemCertificates\AuthRoot'
        try {
            $val = (Get-ItemProperty -Path $regPath -Name DisableRootAutoUpdate -ErrorAction SilentlyContinue).DisableRootAutoUpdate
            if ($val -eq 1) {
                Set-RegistryValue -Path $regPath -Name 'DisableRootAutoUpdate' -Value 0 -PropertyType DWord
                Write-MyOutput 'Root certificate auto-update re-enabled (was disabled by policy — required for Exchange Online / M365 connectivity)'
            }
            else {
                Write-MyVerbose 'Root certificate auto-update: not disabled by policy (OK)'
            }
        }
        catch {
            Write-MyVerbose ('Root certificate auto-update check: {0}' -f $_.Exception.Message)
        }
    }

    function Disable-MRSProxy {
        # F18: MRS Proxy enables cross-forest / cross-org mailbox moves. Disable when not in use —
        # HealthChecker flags an enabled MRS Proxy endpoint as unnecessary attack surface.
        # Re-enable with: Set-WebServicesVirtualDirectory -MRSProxyEnabled $true -Confirm:$false
        if (-not $State['InstallMailbox']) { return }
        Write-MyOutput 'Disabling MRS Proxy on EWS virtual directory (enable manually for cross-forest migrations)'
        try {
            Get-WebServicesVirtualDirectory -Server $env:computername -ErrorAction Stop |
                Set-WebServicesVirtualDirectory -MRSProxyEnabled $false -Confirm:$false -ErrorAction Stop
            Write-MyVerbose 'MRS Proxy disabled (Set-WebServicesVirtualDirectory -MRSProxyEnabled $false)'
        }
        catch {
            Write-MyWarning ('Could not disable MRS Proxy: {0}' -f $_.Exception.Message)
        }
    }

    function Set-MAPIEncryptionRequired {
        # F19: Requires MAPI encryption on all MAPI-over-RPC Outlook connections.
        # Prevents signing-only or cleartext MAPI sessions. ExchangeDsc / HealthChecker recommendation.
        if (-not $State['InstallMailbox']) { return }
        Write-MyOutput 'Setting MAPI encryption as required on mailbox server'
        try {
            Set-MailboxServer -Identity $env:computername -MAPIEncryptionRequired $true -Confirm:$false -ErrorAction Stop
            Write-MyVerbose 'MAPI encryption required (Set-MailboxServer -MAPIEncryptionRequired $true)'
        }
        catch {
            Write-MyWarning ('Could not set MAPI encryption required: {0}' -f $_.Exception.Message)
        }
    }

    function Set-SchannelProtocol {
        param( [string]$Protocol, [bool]$Enable )
        $base = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols"
        $enabled       = if ($Enable) { 1 } else { 0 }
        $disabledByDef = if ($Enable) { 0 } else { 1 }
        New-Item -Path $base -Name $Protocol -ErrorAction SilentlyContinue | Out-Null
        foreach ( $role in 'Client', 'Server') {
            New-Item -Path "$base\$Protocol" -Name $role -ErrorAction SilentlyContinue | Out-Null
            Set-ItemProperty -Path "$base\$Protocol\$role" -Name 'DisabledByDefault' -Value $disabledByDef -Type DWord
            Set-ItemProperty -Path "$base\$Protocol\$role" -Name 'Enabled'           -Value $enabled       -Type DWord
            Register-ExecutedCommand -Category 'Hardening' -Command ("Set-ItemProperty '{0}\{1}\{2}' -Name Enabled -Value {3}; Set-ItemProperty '{0}\{1}\{2}' -Name DisabledByDefault -Value {4}" -f $base, $Protocol, $role, $enabled, $disabledByDef)
        }
    }

    function Set-NetFrameworkStrongCrypto {
        # HealthChecker requires all 4 paths (v4.0 + v2.0, 64-bit + 32-bit)
        foreach ( $path in 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319',
                            'HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319',
                            'HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727',
                            'HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v2.0.50727') {
            if (Test-Path $path) {
                Set-ItemProperty -Path $path -Name 'SystemDefaultTlsVersions' -Value 1 -Type DWord
                Set-ItemProperty -Path $path -Name 'SchUseStrongCrypto'        -Value 1 -Type DWord
                Register-ExecutedCommand -Category 'Hardening' -Command ("Set-ItemProperty '{0}' -Name SystemDefaultTlsVersions -Value 1; Set-ItemProperty '{0}' -Name SchUseStrongCrypto -Value 1" -f $path)
            }
        }
    }

    function Set-TLSSettings {

        param(
            [switch]$TLS12,
            [switch]$TLS13
        )

        if ( $TLS12) {
            Write-MyVerbose 'Enabling TLS 1.2 and configuring .NET Framework strong crypto'
            Set-NetFrameworkStrongCrypto
            Set-SchannelProtocol -Protocol 'TLS 1.2' -Enable $true
        }
        else {
            Write-MyVerbose 'Disabling TLS 1.2'
            Set-SchannelProtocol -Protocol 'TLS 1.2' -Enable $false
        }

        if ( [System.Version]$FullOSVersion -ge [System.Version]$WS2022_PREFULL -and [System.Version]$State['ExSetupVersion'] -ge [System.Version]$EX2019SETUPEXE_CU15) {
            if ( $TLS13) {
                Write-MyVerbose 'Enabling TLS 1.3 and configuring .NET Framework strong crypto'
                Set-NetFrameworkStrongCrypto
                Set-SchannelProtocol -Protocol 'TLS 1.3' -Enable $true
                # Configure the TLS 1.3 cipher suites (cmdlet requires WS2022+)
                if (Get-Command Enable-TlsCipherSuite -ErrorAction SilentlyContinue) {
                    Enable-TlsCipherSuite -Name TLS_AES_256_GCM_SHA384 -Position 0
                    Enable-TlsCipherSuite -Name TLS_AES_128_GCM_SHA256 -Position 1
                }
                else {
                    Write-MyWarning 'Enable-TlsCipherSuite cmdlet not available on this OS, skipping TLS 1.3 cipher suite configuration'
                }
            }
            else {
                Write-MyVerbose 'Disabling TLS 1.3'
                Set-SchannelProtocol -Protocol 'TLS 1.3' -Enable $false
                Disable-TlsCipherSuite -Name TLS_AES_256_GCM_SHA384 -ErrorAction SilentlyContinue
                Disable-TlsCipherSuite -Name TLS_AES_128_GCM_SHA256 -ErrorAction SilentlyContinue
            }
        }
        else {
            Write-MyWarning 'TLS13 configuration not supported for this OS or Exchange version'
        }

    }

    function Enable-WindowsDefenderExclusions {

        if ( Get-Command -Name Add-MpPreference -ErrorAction SilentlyContinue) {
            $SystemRoot = $Env:SystemRoot
            $SystemDrive = $Env:SystemDrive

            Write-MyOutput 'Configuring Windows Defender folder exclusions'
            if ( $State['TargetPath']) {
                $InstallFolder = $State['TargetPath']
            }
            else {
                # TargetPath not specified, using default location
                $InstallFolder = 'C:\Program Files\Microsoft\Exchange Server\V15'
            }

            $Locations = @(
                "$SystemRoot|Cluster",
                "$InstallFolder|ClientAccess\OAB,FIP-FS,GroupMetrics,Logging,Mailbox",
                "$InstallFolder\TransportRoles\Data|IpFilter,Queue,SenderReputation,Temp",
                "$InstallFolder\TransportRoles|Logs,Pickup,Replay",
                "$InstallFolder\UnifiedMessaging|Grammars,Prompts,Temp,VoiceMail",
                "$InstallFolder|Working\OleConverter",
                "$SystemDrive\InetPub\Temp|IIS Temporary Compressed Files",
                "$SystemDrive|Temp\OICE_*"
            )

            foreach ( $Location in $Locations) {
                $Parts = $Location -split '\|'
                $Items = $Parts[1] -split ','
                foreach ( $Item in $Items) {
                    $ExcludeLocation = Join-Path -Path $Parts[0] -ChildPath $Item
                    Write-MyVerbose "WindowsDefender: Excluding location $ExcludeLocation"
                    try {
                        Add-MpPreference -ExclusionPath $ExcludeLocation -ErrorAction SilentlyContinue
                    }
                    catch {
                        Write-MyWarning $_.Exception.Message
                    }
                }
            }

            Write-MyOutput 'Configuring Windows Defender exclusions: NodeRunner process'
            $Processes = @(
                "$InstallFolder\Bin|ComplianceAuditService.exe,Microsoft.Exchange.Directory.TopologyService.exe,Microsoft.Exchange.EdgeSyncSvc.exe,Microsoft.Exchange.Notifications.Broker.exe,Microsoft.Exchange.ProtectedServiceHost.exe,Microsoft.Exchange.RPCClientAccess.Service.exe,Microsoft.Exchange.Search.Service.exe,Microsoft.Exchange.Store.Service.exe,Microsoft.Exchange.Store.Worker.exe,MSExchangeCompliance.exe,MSExchangeDagMgmt.exe,MSExchangeDelivery.exe,MSExchangeFrontendTransport.exe,MSExchangeMailboxAssistants.exe,MSExchangeMailboxReplication.exe,MSExchangeRepl.exe,MSExchangeSubmission.exe,MSExchangeThrottling.exe,OleConverter.exe,UmService.exe,UmWorkerProcess.exe,wsbexchange.exe,EdgeTransport.exe,Microsoft.Exchange.AntispamUpdateSvc.exe,Microsoft.Exchange.Diagnostics.Service.exe,Microsoft.Exchange.Servicehost.exe,MSExchangeHMHost.exe,MSExchangeHMWorker.exe,MSExchangeTransport.exe,MSExchangeTransportLogSearch.exe",
                "$InstallFolder\FIP-FS\Bin|fms.exe,ScanEngineTest.exe,ScanningProcess.exe,UpdateService.exe",
                "$InstallFolder|Bin\Search\Ceres|HostController\HostControllerService.exe,Runtime\1.0\Noderunner.exe,ParserServer\ParserServer.exe",
                "$InstallFolder|FrontEnd\PopImap|Microsoft.Exchange.Imap4.exe,Microsoft.Exchange.Pop3.exe",
                "$InstallFolder|ClientAccess\PopImap\Microsoft.Exchange.Imap4service.exe,Microsoft.Exchange.Pop3service.exe",
                "$InstallFolder|FrontEnd\CallRouter|Microsoft.Exchange.UM.CallRouter.exe",
                "$InstallFolder|TransportRoles\agents\Hygiene\Microsoft.Exchange.ContentFilter.Wrapper.exe"
            )

            foreach ( $Process in $Processes) {
                $Parts = $Process -split '\|'
                $Items = $Parts[1] -split ','
                foreach ( $Item in $Items) {
                    $ExcludeProcess = Join-Path -Path $Parts[0] -ChildPath $Item
                    Write-MyVerbose "WindowsDefender: Excluding process $ExcludeProcess"
                    try {
                        Add-MpPreference -ExclusionProcess $ExcludeProcess -ErrorAction SilentlyContinue
                    }
                    catch {
                        Write-MyWarning $_.Exception.Message
                    }
                }
            }

            $Extensions = 'dsc', 'txt', 'cfg', 'grxml', 'lzx', 'config', 'chk', 'edb', 'jfm', 'jrs', 'log', 'que'
            foreach ( $Extension in $Extensions) {
                $ExcludeExtension = '.{0}' -f $Extension
                Write-MyVerbose "WindowsDefender: Excluding extension $ExcludeExtension"
                try {
                    Add-MpPreference -ExclusionExtension $ExcludeExtension -ErrorAction SilentlyContinue
                }
                catch {
                    Write-MyWarning $_.Exception.Message
                }
            }
            Register-ExecutedCommand -Category 'Hardening' -Command ("Add-MpPreference -ExclusionPath '{0}\Mailbox','{0}\Logging','{0}\FIP-FS',...  # see chapter 8.5 for complete path list" -f $InstallFolder)
            Register-ExecutedCommand -Category 'Hardening' -Command ("Add-MpPreference -ExclusionProcess '{0}\Bin\MSExchangeDelivery.exe','{0}\Bin\MSExchangeTransport.exe',...  # see chapter 8.5 for complete process list" -f $InstallFolder)
            Register-ExecutedCommand -Category 'Hardening' -Command 'Add-MpPreference -ExclusionExtension .edb,.jrs,.jfm,.chk,.log,.que,.cfg,.grxml,.lzx,.config,.dsc,.txt'
        }
        else {
            Write-MyVerbose 'Windows Defender not installed'
        }
    }

    function Disable-DefenderTamperProtection {
        # Tamper Protection blocks Set-MpPreference from taking effect. It cannot be disabled
        # via PowerShell/registry once MDE/Intune enforces it — those must be set via the
        # Security Center / Intune policy. On unmanaged devices we can flip the registry flag
        # as best-effort. Re-enabled in Enable-DefenderTamperProtection.
        if (-not (Get-Command -Name Get-MpComputerStatus -ErrorAction SilentlyContinue)) { return }
        try {
            $status = Get-MpComputerStatus -ErrorAction Stop
            if (-not $status.IsTamperProtected) {
                Write-MyVerbose 'Defender Tamper Protection already off — nothing to do'
                return
            }
            Write-MyOutput 'Attempting to disable Defender Tamper Protection (best-effort, registry)'
            $tpPath = 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Features'
            # Capture current value so we can restore it, even if not present
            $prev   = (Get-ItemProperty -Path $tpPath -Name 'TamperProtection' -ErrorAction SilentlyContinue).TamperProtection
            if ($null -eq $prev) { $State['DefenderTPPrev'] = '__absent__' } else { $State['DefenderTPPrev'] = [int]$prev }
            Set-RegistryValue -Path $tpPath -Name 'TamperProtection' -Value 0
            Start-Sleep -Seconds 2
            $post = Get-MpComputerStatus -ErrorAction SilentlyContinue
            if ($post -and $post.IsTamperProtected) {
                Write-MyWarning 'Tamper Protection still active — likely enforced by Intune/MDE. Realtime disable may be ignored.'
                Write-MyWarning '  Disable Tamper Protection manually in Windows Security / Intune before install, or accept that setup runs with AV active.'
            }
            else {
                Write-MyVerbose 'Tamper Protection flag cleared successfully'
            }
            $State['DefenderTPDisabledByEXpress'] = $true
            Save-State $State
        }
        catch {
            Write-MyWarning ('Could not inspect/disable Tamper Protection: {0}' -f $_.Exception.Message)
        }
    }

    function Enable-DefenderTamperProtection {
        # Restore the Tamper Protection registry value we captured before flipping it.
        if (-not $State['DefenderTPDisabledByEXpress']) { return }
        try {
            $tpPath = 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Features'
            $prev   = $State['DefenderTPPrev']
            if ($prev -eq '__absent__') {
                Remove-ItemProperty -Path $tpPath -Name 'TamperProtection' -ErrorAction SilentlyContinue
                Write-MyOutput 'Tamper Protection registry value removed (original state)'
            }
            elseif ($null -ne $prev) {
                Set-RegistryValue -Path $tpPath -Name 'TamperProtection' -Value ([int]$prev)
                Write-MyOutput ('Tamper Protection registry value restored to {0}' -f $prev)
            }
            $State.Remove('DefenderTPDisabledByEXpress') | Out-Null
            $State.Remove('DefenderTPPrev') | Out-Null
            Save-State $State
        }
        catch {
            Write-MyWarning ('Could not restore Tamper Protection: {0}' -f $_.Exception.Message)
        }
    }

    function Disable-DefenderRealtimeMonitoring {
        # Temporarily disable Defender real-time scanning during Exchange install/hardening.
        # Setup and SU runs generate massive file I/O (ECP/OWA .config unpacking, assembly
        # ngen, transport agents) that Defender scans inline, causing setup to stall or fail
        # with random file-lock errors. Re-enabled at the start of Phase 6.
        # Accepted risk: GPO/Intune may re-enable during the window. Flag is idempotent.
        if (-not (Get-Command -Name Set-MpPreference -ErrorAction SilentlyContinue)) {
            Write-MyVerbose 'Windows Defender not installed — skipping realtime disable'
            return
        }
        # Tamper Protection must be cleared first, otherwise Set-MpPreference is silently ignored.
        Disable-DefenderTamperProtection
        try {
            $pref = Get-MpPreference -ErrorAction Stop
            if ($pref.DisableRealtimeMonitoring) {
                Write-MyVerbose 'Defender realtime monitoring already disabled — leaving as-is'
                return
            }
            Write-MyOutput 'Disabling Windows Defender realtime monitoring during Exchange install'
            Set-MpPreference -DisableRealtimeMonitoring $true -ErrorAction Stop
            Start-Sleep -Seconds 1
            $post = Get-MpPreference -ErrorAction SilentlyContinue
            if ($post -and -not $post.DisableRealtimeMonitoring) {
                Write-MyWarning 'Realtime monitoring did not stay disabled — Tamper Protection or policy override active. Continuing with AV on.'
                return
            }
            $State['DefenderRealtimeDisabledByEXpress'] = $true
            Save-State $State
        }
        catch {
            Write-MyWarning ('Could not disable Defender realtime monitoring: {0}' -f $_.Exception.Message)
        }
    }

    function Enable-DefenderRealtimeMonitoring {
        # Re-enable Defender realtime scanning. With -Force the function always attempts
        # to turn realtime on, regardless of whether EXpress was the one to disable it —
        # this is used right before the Word report generates so the report reflects an
        # active protection state after installation. Without -Force, it only reverses
        # an EXpress-initiated disable (flag set in Disable-DefenderRealtimeMonitoring).
        param([switch]$Force)
        if (-not (Get-Command -Name Set-MpPreference -ErrorAction SilentlyContinue)) { return }
        $shouldAct = $Force -or $State['DefenderRealtimeDisabledByEXpress']
        if ($shouldAct) {
            try {
                $pref = Get-MpPreference -ErrorAction Stop
                if (-not $pref.DisableRealtimeMonitoring) {
                    Write-MyVerbose 'Defender realtime monitoring already enabled'
                }
                else {
                    if ($Force) { Write-MyOutput 'Ensuring Windows Defender realtime monitoring is enabled (pre-report)' }
                    else        { Write-MyOutput 'Re-enabling Windows Defender realtime monitoring' }
                    Set-MpPreference -DisableRealtimeMonitoring $false -ErrorAction Stop
                    Start-Sleep -Seconds 1
                    $post = Get-MpPreference -ErrorAction SilentlyContinue
                    if ($post -and $post.DisableRealtimeMonitoring) {
                        Write-MyWarning 'Realtime monitoring still disabled after set — Tamper Protection or policy override active.'
                    }
                }
                if ($State['DefenderRealtimeDisabledByEXpress']) {
                    $State.Remove('DefenderRealtimeDisabledByEXpress') | Out-Null
                    Save-State $State
                }
            }
            catch {
                Write-MyWarning ('Could not re-enable Defender realtime monitoring: {0}' -f $_.Exception.Message)
            }
        }
        else {
            Write-MyVerbose 'Defender realtime monitoring was not disabled by EXpress — skipping re-enable'
        }
        # Restore Tamper Protection regardless — it may have been flipped without realtime change.
        Enable-DefenderTamperProtection
    }

    # Return location of mounted drive if ISO specified
    function Resolve-SourcePath {
        param (
            [String]$SourceImage
        )
        $disk = Get-DiskImage -ImagePath $SourceImage -ErrorAction SilentlyContinue
        if ( $disk) {
            if ( $disk.Attached) {
                $vol = $disk | Get-Volume -ErrorAction SilentlyContinue
                if ( $vol) {
                    $Drive = $vol.DriveLetter
                }
                else {
                    Write-Verbose ('{0} already attached but no drive letter - will mount again' -f $SourceImage)
                    $Drive = (Mount-DiskImage -ImagePath $SourceImage -PassThru | Get-Volume).DriveLetter
                }
            }
            else {
                $Drive = (Mount-DiskImage -ImagePath $SourceImage -PassThru | Get-Volume).DriveLetter
            }
            $SourcePath = '{0}:\' -f $Drive
            Write-Verbose ('Mounted {0} on drive {1}' -f $SourceImage, $SourcePath)
            return $SourcePath
        }
        else {
            return $null
        }
    }

    function Get-VCRuntime {
        param (
            [String]$version,
            [String]$MinBuild = ''
        )
        Write-MyVerbose ('Looking for presence of Visual C++ v{0} Runtime' -f $version)
        $presence = $false
        $build = $null

        # Primary check: VisualStudio registry paths (used by VC++ 2015+ / VS 14.x bundles,
        # and some variants of 2012/2013 installers).
        $RegPaths = @(
            'HKLM:\Software\WOW6432Node\Microsoft\VisualStudio\{0}\VC\Runtimes\x64',
            'HKLM:\Software\Microsoft\VisualStudio\{0}\VC\Runtimes\x64',
            'HKLM:\Software\WOW6432Node\Microsoft\VisualStudio\{0}\VC\VCRedist\x64',
            'HKLM:\Software\Microsoft\VisualStudio\{0}\VC\VCRedist\x64')
        foreach ( $RegPath in $RegPaths) {
            $Key = (Get-ItemProperty -Path ($RegPath -f $version) -Name Installed -ErrorAction SilentlyContinue).Installed
            if ( $Key -eq 1) {
                $build = (Get-ItemProperty -Path ($RegPath -f $version) -Name Version -ErrorAction SilentlyContinue).Version
                $presence = $true
                break
            }
        }

        # Fallback 1: scan Add/Remove Programs for matching display name.
        # VC++ 2013 (12.0) and older standalone redistributables do not write to
        # the VisualStudio\{ver}\VC\Runtimes path — they only register here.
        if (-not $presence) {
            $yearMap = @{ '10.0' = '2010'; '11.0' = '2012'; '12.0' = '2013'; '14.0' = '2015' }
            $yearStr  = $yearMap[$version]
            if ($yearStr) {
                foreach ($hive in @('HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
                                    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')) {
                    # Match display name without requiring "(x64)" — the format varies by installer
                    # version (e.g. "...Redistributable (x64)..." vs "...x64 Minimum Runtime...").
                    $entry = Get-ChildItem $hive -ErrorAction SilentlyContinue |
                             Get-ItemProperty -ErrorAction SilentlyContinue |
                             Where-Object { $_.DisplayName -like "Microsoft Visual C++ $yearStr*" } |
                             Sort-Object DisplayVersion -Descending |
                             Select-Object -First 1
                    if ($entry) {
                        $build    = $entry.DisplayVersion
                        $presence = $true
                        Write-MyVerbose ('Found Visual C++ v{0} in Add/Remove Programs: {1}' -f $version, $entry.DisplayName)
                        break
                    }
                }
            }
        }

        # Fallback 2: check the runtime DLL in System32 — the same check Exchange Setup uses.
        # msvcr110.dll = VC++ 2012 (11.0), msvcr120.dll = VC++ 2013 (12.0)
        if (-not $presence) {
            $dllMap = @{ '11.0' = 'msvcr110.dll'; '12.0' = 'msvcr120.dll' }
            $dll    = $dllMap[$version]
            if ($dll) {
                $dllPath = Join-Path $env:SystemRoot "System32\$dll"
                if (Test-Path $dllPath) {
                    $build    = (Get-Item $dllPath -ErrorAction SilentlyContinue).VersionInfo.ProductVersion
                    $presence = $true
                    Write-MyVerbose ('Found Visual C++ v{0} via {1}, version {2}' -f $version, $dll, $build)
                }
            }
        }

        if ($presence) {
            Write-MyVerbose ('Found Visual C++ Runtime v{0}, build {1}' -f $version, $build)
            if ($MinBuild -and $build -and ([System.Version]$build -lt [System.Version]$MinBuild)) {
                Write-MyVerbose ('Visual C++ v{0} build {1} is older than required minimum {2} — will update' -f $version, $build, $MinBuild)
                return $false
            }
        }
        else {
            Write-MyVerbose ('Could not find Visual C++ v{0} Runtime installed' -f $version)
        }
        return $presence
    }

    function Start-DisableMSExchangeAutodiscoverAppPoolJob {

        $ScriptBlock = {
            # IIS:\ PSDrive is not available in Start-Job child processes without an explicit import.
            Import-Module WebAdministration -ErrorAction SilentlyContinue

            $maxWaitSec = 600   # give up after 10 minutes
            $elapsed    = 0
            $interval   = 10
            do {
                # Use Test-Path instead of Get-WebAppPoolState: the latter internally calls
                # Get-WebItemState which throws PathNotFound and is NOT suppressed by -ErrorAction SilentlyContinue.
                if (Test-Path 'IIS:\AppPools\MSExchangeAutodiscoverAppPool') {
                    Write-Verbose 'Stopping and blocking startup of MSExchangeAutodiscoverAppPool'
                    if ( (Get-WebAppPoolState -Name 'MSExchangeAutodiscoverAppPool').Value -ine 'Stopped') {
                        try {
                            Stop-WebAppPool -Name 'MSExchangeAutodiscoverAppPool' -ErrorAction Stop
                        }
                        catch {
                            Write-Error ('Failed to stop app pool: {0}' -f $_.Exception.Message)
                        }
                    }
                    try {
                        Set-ItemProperty "IIS:\AppPools\MSExchangeAutodiscoverAppPool" -Name "autoStart" -Value $false -ErrorAction Stop
                        Set-ItemProperty "IIS:\AppPools\MSExchangeAutodiscoverAppPool" -Name "startMode" -Value "OnDemand" -ErrorAction Stop
                    }
                    catch {
                        Write-Error ('Failed to update app pool properties: {0}' -f $_.Exception.Message)
                    }
                    return $true
                }
                Start-Sleep -Seconds $interval
                $elapsed += $interval
            } while ($elapsed -lt $maxWaitSec)
            Write-Verbose 'MSExchangeAutodiscoverAppPool did not appear within 10 minutes — giving up'
            return $false
        }

        $Job = Start-Job -ScriptBlock $ScriptBlock -Name ('DisableMSExchangeAutodiscoverAppPoolJob-{0}' -f $env:COMPUTERNAME)
        Add-BackgroundJob $Job

        Write-MyOutput ('Started background job to disable MSExchangeAutodiscoverAppPool (Job ID: {0})' -f $Job.Id)
        return $Job
    }

    function Enable-MSExchangeAutodiscoverAppPool {
        # Use Test-Path instead of Get-WebAppPoolState: the latter internally calls
        # Get-WebItemState which throws a provider PathNotFound error that is NOT
        # suppressed by -ErrorAction SilentlyContinue.
        if (-not (Test-Path 'IIS:\AppPools\MSExchangeAutodiscoverAppPool' -ErrorAction SilentlyContinue)) {
            Write-MyVerbose 'MSExchangeAutodiscoverAppPool not found'
            return $false
        }

        Write-MyOutput 'Starting and enabling startup of MSExchangeAutodiscoverAppPool'
        try {
            Start-WebAppPool -Name 'MSExchangeAutodiscoverAppPool' -ErrorAction Stop
        }
        catch {
            Write-MyWarning ('Failed to start app pool: {0}' -f $_.Exception.Message)
        }
        try {
            Set-ItemProperty 'IIS:\AppPools\MSExchangeAutodiscoverAppPool' -Name 'autoStart' -Value $true  -ErrorAction Stop
            Set-ItemProperty 'IIS:\AppPools\MSExchangeAutodiscoverAppPool' -Name 'startMode' -Value 'OnDemand' -ErrorAction Stop
        }
        catch {
            Write-MyWarning ('Failed to update app pool properties: {0}' -f $_.Exception.Message)
        }
        return $true
    }

    function Stop-BackgroundJobs {
        if ($Global:BackgroundJobs -and $Global:BackgroundJobs.Count -gt 0) {
            Write-MyVerbose "Cleaning up $($Global:BackgroundJobs.Count) background job(s)..."
            foreach ($Job in $Global:BackgroundJobs) {
                if ($Job.State -eq 'Running') {
                    # Wait up to 30 seconds for job to finish gracefully
                    $null = $Job | Wait-Job -Timeout 30 -ErrorAction SilentlyContinue
                    if ($Job.State -eq 'Running') {
                        Write-MyWarning ('Background job {0} (ID {1}) did not finish within 30 seconds, forcing stop' -f $Job.Name, $Job.Id)
                        Stop-Job -Job $Job -ErrorAction SilentlyContinue
                    }
                }
                $JobOutput = Receive-Job -Job $Job
                Write-MyVerbose ('Cleanup background job: {0} (ID {1}), Output {2}' -f $Job.Name, $Job.Id, $JobOutput)
                Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue
            }
            $Global:BackgroundJobs = @()
            Write-MyVerbose "Background job cleanup completed."
        }
    }

    function Get-AdvancedFeatureCatalog {
        # Advanced Configuration catalog. Each entry:
        #   Name        — unique key persisted in $State['AdvancedFeatures'] and config file
        #   Label       — short display text in the Advanced menu (max ~30 chars)
        #   Description — one-line explanation shown in description panel
        #   Default     — $true/$false; matches current behaviour unless noted
        #   Category    — 'TLS', 'Hardening', 'Performance', 'ExchangePolicy', 'PostConfig', 'InstallFlow'
        #   Condition   — (optional) scriptblock; entry hidden when it returns $false
        return [ordered]@{
            # ─── Security / TLS ──────────────────────────────────────────────
            DisableSSL3         = @{ Category='TLS'; Label='Disable SSL 3.0';                Default=$true;  Description='Disable legacy SSL 3.0 (POODLE, CVE-2014-3566).' }
            DisableRC4          = @{ Category='TLS'; Label='Disable RC4 cipher';             Default=$true;  Description='Disable deprecated RC4 stream cipher.' }
            EnableECC           = @{ Category='TLS'; Label='Prefer ECC key exchange';        Default=$true;  Description='Enable ECC cipher suites and prefer over RSA.' }
            NoCBC               = @{ Category='TLS'; Label='Disable CBC ciphers';            Default=$false; Description='Disables CBC cipher suites. Not recommended — breaks compatibility with several clients.' }
            EnableAMSI          = @{ Category='TLS'; Label='Enable AMSI';                    Default=$true;  Description='Antimalware Scan Interface for Exchange transport and OWA.' }
            EnableTLS12         = @{ Category='TLS'; Label='Enforce TLS 1.2';                Default=$true;  Description='Enforce TLS 1.2; disables TLS 1.0/1.1 on SChannel and .NET StrongCrypto.' }
            EnableTLS13         = @{ Category='TLS'; Label='Enable TLS 1.3';                 Default=$true;  Description='Enable TLS 1.3 (Windows Server 2022+).'; Condition={ [System.Version]$script:FullOSVersion -ge [System.Version]$script:WS2022_PREFULL } }
            DoNotEnableEP       = @{ Category='TLS'; Label='Opt-out: Extended Protection';   Default=$false; Description='Skip Extended Protection configuration. Required for Hybrid + Modern Hybrid Topology where EP is incompatible.' }

            # ─── Security / Hardening ────────────────────────────────────────
            SMBv1Disable        = @{ Category='Hardening'; Label='Disable SMBv1';             Default=$true;  Description='Remove SMBv1 (WannaCry mitigation, MS17-010).' }
            NetBIOSDisable      = @{ Category='Hardening'; Label='Disable NetBIOS/TCP';       Default=$true;  Description='Disable NetBIOS over TCP/IP on all NICs (reduces attack surface).' }
            LLMNRDisable        = @{ Category='Hardening'; Label='Disable LLMNR';             Default=$true;  Description='Disable Link-Local Multicast Name Resolution (CIS L1 §18.5.4.2).' }
            MDNSDisable         = @{ Category='Hardening'; Label='Disable mDNS';              Default=$true;  Description='Disable Multicast DNS responder (WS2022+).' }
            WDigestDisable      = @{ Category='Hardening'; Label='Disable WDigest caching';   Default=$true;  Description='Prevent plaintext credentials in LSASS memory.' }
            LSAProtection       = @{ Category='Hardening'; Label='Enable LSA Protection';     Default=$true;  Description='RunAsPPL for LSASS to prevent credential dumping.' }
            LmCompat5           = @{ Category='Hardening'; Label='LmCompatibilityLevel=5';    Default=$true;  Description='Enforce NTLMv2, refuse LM/NTLMv1.' }
            SerializedDataSig   = @{ Category='Hardening'; Label='SerializedDataSigning';     Default=$true;  Description='Exchange SerializedDataSigning (MS-mandatory post CVE-2023-21529).' }
            ShutdownTrackerOff  = @{ Category='Hardening'; Label='Disable Shutdown Tracker';  Default=$true;  Description='Suppress the Shutdown Event Tracker reason dialog on server shutdowns.' }
            HSTS                = @{ Category='Hardening'; Label='HSTS on OWA/ECP';           Default=$true;  Description='HTTP Strict-Transport-Security header on OWA/ECP virtual directories.' }
            MAPIEncryption      = @{ Category='Hardening'; Label='Required MAPI encryption';  Default=$true;  Description='Set-RpcClientAccess -EncryptionRequired $true.' }
            HTTP2Disable        = @{ Category='Hardening'; Label='Disable HTTP/2';            Default=$true;  Description='Workaround for Exchange compatibility issues with HTTP/2.' }
            CredentialGuardOff  = @{ Category='Hardening'; Label='Disable Credential Guard';  Default=$true;  Description='Exchange is incompatible with Credential Guard; disable if enabled.' }
            UnnecessaryServices = @{ Category='Hardening'; Label='Disable unneeded services'; Default=$true;  Description='Disable Print Spooler, Xbox, Geolocation and other unneeded services on Exchange servers.' }
            WindowsSearchOff    = @{ Category='Hardening'; Label='Disable Windows Search';    Default=$true;  Description='Disable Windows Search service (not used by Exchange; uses CPU/IO).' }
            CRLTimeout          = @{ Category='Hardening'; Label='CRL Check Timeout';         Default=$true;  Description='Tune CRL retrieval timeout to avoid slow startup when OCSP/CRL endpoints are unreachable.' }
            RootCAAutoUpdate    = @{ Category='Hardening'; Label='Root CA Auto-Update';       Default=$true;  Description='Keep Automatic Root Certificates Update enabled (required for Modern Auth / O365 Hybrid).' }
            SMTPBannerHarden    = @{ Category='Hardening'; Label='Harden SMTP banner';        Default=$true;  Description='Replace Exchange version banner on Frontend Receive Connectors with "220 Mail Service".' }

            # ─── Performance / Tuning ────────────────────────────────────────
            MaxConcurrentAPI    = @{ Category='Performance'; Label='MaxConcurrentAPI';        Default=$true;  Description='MS KB 2688798 — raise MaxConcurrentApi to prevent NTLM auth bottlenecks.' }
            DiskAllocHint       = @{ Category='Performance'; Label='Disk allocation hint';    Default=$true;  Description='Emit warning when DB/log volumes are not formatted with 64K NTFS cluster size.' }
            CtsProcAffinity     = @{ Category='Performance'; Label='Content conv. affinity';  Default=$true;  Description='Limit Content Conversion processor affinity to stabilise CPU load.' }
            NodeRunnerMemLimit  = @{ Category='Performance'; Label='NodeRunner RAM cap';      Default=$true;  Description='Cap Exchange Search NodeRunner memory to prevent runaway allocations.' }
            MapiFeGC            = @{ Category='Performance'; Label='MAPI FrontEnd Server GC'; Default=$true;  Description='Enable Server GC mode for MAPI FrontEnd AppPool.' }
            NICPowerMgmtOff     = @{ Category='Performance'; Label='NIC Power Management';    Default=$true;  Description='Disable "Allow computer to turn off this device" on all NICs.' }
            RSSEnable           = @{ Category='Performance'; Label='Receive Side Scaling';    Default=$true;  Description='Enable RSS on all NICs for multi-core packet processing.' }
            TCPTuning           = @{ Category='Performance'; Label='TCP tuning';              Default=$true;  Description='Autotuning, Chimney offload and related TCP stack tweaks for Exchange workloads.' }
            TCPOffloadOff       = @{ Category='Performance'; Label='Disable TCP offload';     Default=$true;  Description='Disable TCP checksum/segmentation offload (avoids driver bugs on Exchange).' }
            IPv4OverIPv6Off     = @{ Category='Performance'; Label='Prefer IPv4 over IPv6';    Default=$true;  Description='Prefer IPv4 over IPv6 (DisabledComponents=0x20) — avoids Exchange DNS-lookup delays on IPv6-only hosts.' }

            # ─── Exchange Org Policy (current Optimization Catalog A–J) ──────
            ModernAuth          = @{ Category='ExchangePolicy'; Label='Modern Auth (OAuth2)';    Default=$true;  Description='Org-wide OAuth2 / Modern Authentication. Required for Outlook 2016+, Teams, mobile.' }
            OWASessionTimeout6h = @{ Category='ExchangePolicy'; Label='OWA Session Timeout 6h';  Default=$true;  Description='Activity-based OWA/ECP session timeout at 6h inactivity.' }
            DisableTelemetry    = @{ Category='ExchangePolicy'; Label='Disable CEIP telemetry';  Default=$true;  Description='Set-OrganizationConfig -CustomerFeedbackEnabled $false (GDPR/DSGVO).' }
            MapiHttp            = @{ Category='ExchangePolicy'; Label='MAPI over HTTP';          Default=$true;  Description='Explicit MapiHttpEnabled — replaces legacy RPC/HTTP.' }
            MaxMessageSize150MB = @{ Category='ExchangePolicy'; Label='Max message size 150MB';  Default=$true;  Description='Raise org-wide + Frontend receive connector max message size to 150 MB.' }
            MessageExpiration7d = @{ Category='ExchangePolicy'; Label='Expiration 7 days';       Default=$true;  Description='Extend transport message expiration to 7 days. Condition: not CopyServerConfig.'; Condition={ -not $script:State['CopyServerConfig'] } }
            HtmlNDR             = @{ Category='ExchangePolicy'; Label='HTML NDR formatting';     Default=$true;  Description='Set-TransportConfig -InternalDsnSendHtml / -ExternalDsnSendHtml.' }
            ShadowRedundancy    = @{ Category='ExchangePolicy'; Label='Shadow Redundancy';       Default=$false; Description='Prefer remote DAG member for shadow copies. DAG-only.'; Condition={ [bool]$script:State['DAGName'] } }
            SafetyNet2d         = @{ Category='ExchangePolicy'; Label='Safety Net 2d hold';      Default=$true;  Description='Safety Net hold time set to 2 days.' }

            # ─── Post-Config / Integration ───────────────────────────────────
            MECA                = @{ Category='PostConfig'; Label='MECA Auth Cert Renewal';  Default=$true;  Description='Register CSS-Exchange MonitorExchangeAuthCertificate scheduled task for automatic renewal.' }
            AntispamAgents      = @{ Category='PostConfig'; Label='Install Antispam Agents'; Default=$true;  Description='Install built-in antispam agents (Mailbox role only; no effect on Edge).' }
            SSLOffloading       = @{ Category='PostConfig'; Label='SSL Offloading tuning';   Default=$true;  Description='IIS/OWA SSL offload settings for load-balanced deployments.' }
            MRSProxy            = @{ Category='PostConfig'; Label='Enable MRS Proxy';        Default=$true;  Description='Enable MRS Proxy on EWS for cross-forest/cross-org mailbox moves.' }
            IANATimezone        = @{ Category='PostConfig'; Label='IANA timezone mapping';   Default=$true;  Description='Configure IANA ↔ Windows timezone mapping (iCal interop).' }
            AnonymousRelay      = @{ Category='PostConfig'; Label='Anonymous relay connector'; Default=$true; Description='Create anonymous internal/external relay connector if RelaySubnets is configured.'; Condition={ [bool]$script:State['RelaySubnets'] -or [bool]$script:State['ExternalRelaySubnets'] } }
            AccessNamespaceMail = @{ Category='PostConfig'; Label='Access Namespace mail config'; Default=$true; Description='Add Access Namespace as Authoritative Accepted Domain and set it as primary SMTP in the default Email Address Policy. Removes .local/nonroutable templates. Only available when EXpress created the Exchange org.'; Condition={ [bool]$script:State['Namespace'] -and [bool]$script:State['NewExchangeOrg'] } }
            SkipHealthCheck     = @{ Category='PostConfig'; Label='Opt-out: HealthChecker';  Default=$false; Description='Skip CSS-Exchange HealthChecker run at end of Phase 6.' }
            RBACReport          = @{ Category='PostConfig'; Label='RBAC Report';             Default=$true;  Description='Generate RBAC (role assignments / role groups) HTML report.' }
            RunEOMT             = @{ Category='PostConfig'; Label='Run EOMT';                Default=$false; Description='Run CSS-Exchange Emergency Mitigation Tool (legacy CUs; no-op on current CUs).' }

            # ─── Install-Flow / Debug ────────────────────────────────────────
            AutoApproveWindowsUpdates = @{ Category='InstallFlow'; Label='Auto-approve Windows Updates'; Default=$false; Description='Autopilot: approve all pending Security/Critical Windows Updates without prompting. Off by default — deliberate opt-in required.' }
            DiagnosticData      = @{ Category='InstallFlow'; Label='Send diagnostic data';   Default=$false; Description='/IAcceptExchangeServerLicenseTerms_DiagnosticDataON — share setup telemetry with Microsoft.' }
            Lock                = @{ Category='InstallFlow'; Label='Lock screen during run'; Default=$false; Description='Lock the console while the installation is in progress (Autopilot only).' }
            SkipRolesCheck      = @{ Category='InstallFlow'; Label='Skip AD roles check';    Default=$false; Description='Skip Schema/Enterprise/Domain Admin membership check (use with caution).' }
            NoCheckpoint        = @{ Category='InstallFlow'; Label='Skip System Restore';    Default=$false; Description='Skip pre-install System Restore checkpoints.' }
            NoNet481            = @{ Category='InstallFlow'; Label='Skip .NET 4.8.1';        Default=$false; Description='Skip .NET 4.8.1 install (debug only — may break Exchange setup).' }
            WaitForADSync       = @{ Category='InstallFlow'; Label='Wait for AD replication'; Default=$false; Description='After PrepareAD, wait up to 6 min for error-free AD replication before continuing.' }
        }
    }

    function Show-AdvancedMenu {
        # Interactive Advanced Configuration menu. 2 categories per page (~3 pages total).
        # Navigation uses Enter / Backspace / 0 / Esc — never letter keys — so all
        # A-Z letters are available for toggling items without conflicts.
        # Returns a hashtable @{Name=bool} of all toggle states, or $null on cancel.
        # $InitialValues: pre-seed toggle state (e.g. from a previous C-press in the main menu).
        param([hashtable]$InitialValues = $null)

        $catalog = Get-AdvancedFeatureCatalog

        $categoryDefs = @(
            @{ Key='TLS';            Title='Security / TLS' }
            @{ Key='Hardening';      Title='Security / Hardening' }
            @{ Key='Performance';    Title='Performance / Tuning' }
            @{ Key='ExchangePolicy'; Title='Exchange Org Policy' }
            @{ Key='PostConfig';     Title='Post-Config / Integration' }
            @{ Key='InstallFlow';    Title='Install-Flow / Debug' }
        )

        # Initialize selection state; filter entries whose Condition is $false.
        $sel     = @{}
        $visible = @{}
        foreach ($cat in $categoryDefs) { $visible[$cat.Key] = @() }
        $existing = if ($InitialValues -is [hashtable])                { $InitialValues }
                    elseif ($State['AdvancedFeatures'] -is [hashtable]) { $State['AdvancedFeatures'] }
                    else                                                 { @{} }
        foreach ($name in $catalog.Keys) {
            $entry = $catalog[$name]
            if ($entry.ContainsKey('Condition')) {
                try { if (-not (& $entry.Condition)) { continue } } catch { continue }
            }
            $sel[$name] = if ($existing.ContainsKey($name)) { [bool]$existing[$name] } else { [bool]$entry.Default }
            $visible[$entry.Category] += ,$name
        }

        # Build pages: 2 non-empty categories per page.
        $activeCats = @($categoryDefs | Where-Object { $visible[$_.Key].Count -gt 0 })
        if ($activeCats.Count -eq 0) { Write-MyVerbose 'No advanced features applicable'; return $sel }
        $pages = @()
        for ($i = 0; $i -lt $activeCats.Count; $i += 2) {
            $pg = @($activeCats[$i])
            if ($i + 1 -lt $activeCats.Count) { $pg += $activeCats[$i + 1] }
            $pages += ,@{ Cats = $pg }
        }

        $useRawKey = $false
        try { $null = $host.UI.RawUI.KeyAvailable; $useRawKey = $true } catch { }

        $pageIdx   = 0
        $lastName  = ''
        $statusMsg = ''

        while ($true) {
            # Flatten all visible names on this page in category order (for letter assignment).
            $pageNames = @()
            foreach ($cat in $pages[$pageIdx].Cats) { $pageNames += $visible[$cat.Key] }
            $count = $pageNames.Count

            Clear-Host
            Write-Host ('=' * 70) -ForegroundColor Cyan
            Write-Host ('  EXpress v{0} — Advanced Configuration  (page {1}/{2})' -f $script:ScriptVersion, ($pageIdx + 1), $pages.Count) -ForegroundColor Cyan
            Write-Host ('=' * 70) -ForegroundColor Cyan

            # Render each category section with its own 2-column block.
            $offset = 0
            foreach ($cat in $pages[$pageIdx].Cats) {
                $catNames = @($visible[$cat.Key])
                if ($catNames.Count -eq 0) { continue }
                $sep = [string]::new([char]0x2500, [Math]::Max(0, 52 - $cat.Title.Length))
                Write-Host ''
                Write-Host ('  -- {0} {1}' -f $cat.Title, $sep) -ForegroundColor Yellow
                Write-Host ''
                $half = [int][Math]::Ceiling($catNames.Count / 2)
                for ($r = 0; $r -lt $half; $r++) {
                    $li      = $r
                    $ri      = $r + $half
                    $lName   = $catNames[$li]
                    $lLetter = [char]([int][char]'A' + $offset + $li)
                    $lEntry  = $catalog[$lName]
                    $lMark   = if ($sel[$lName]) { 'X' } else { ' ' }
                    $lColor  = if ($lName -eq $lastName) { [System.ConsoleColor]::Yellow } else { [System.ConsoleColor]::White }
                    Write-Host ('  [{0}] [{1}] {2,-28}' -f $lLetter, $lMark, $lEntry.Label) -ForegroundColor $lColor -NoNewline
                    if ($ri -lt $catNames.Count) {
                        $rName   = $catNames[$ri]
                        $rLetter = [char]([int][char]'A' + $offset + $ri)
                        $rEntry  = $catalog[$rName]
                        $rMark   = if ($sel[$rName]) { 'X' } else { ' ' }
                        $rColor  = if ($rName -eq $lastName) { [System.ConsoleColor]::Yellow } else { [System.ConsoleColor]::White }
                        Write-Host ('   [{0}] [{1}] {2,-28}' -f $rLetter, $rMark, $rEntry.Label) -ForegroundColor $rColor
                    } else {
                        Write-Host ''
                    }
                }
                $offset += $catNames.Count
            }

            # Description panel
            Write-Host ''
            Write-Host ('  ' + [string]::new([char]0x2500, 66)) -ForegroundColor DarkGray
            if ($lastName -and $catalog.Contains($lastName)) {
                $opt       = $catalog[$lastName]
                $dispState = if ($sel[$lastName]) { 'ENABLED' } else { 'DISABLED' }
                Write-Host ('  {0}  ({1})' -f $opt.Label, $dispState) -ForegroundColor Yellow
                Write-Host ''
                $words = ($opt.Description -replace '\s+', ' ').Trim() -split ' '
                $line  = '  '
                foreach ($w in $words) {
                    if (($line + $w).Length -gt 68) { Write-Host $line; $line = '  ' + $w + ' ' }
                    else { $line += $w + ' ' }
                }
                if ($line.Trim()) { Write-Host $line }
            } else {
                Write-Host '  Press a letter to toggle it and see its description.' -ForegroundColor DarkGray
            }
            Write-Host ('  ' + [string]::new([char]0x2500, 66)) -ForegroundColor DarkGray
            Write-Host ''

            if ($statusMsg) { Write-Host ('  ' + $statusMsg) -ForegroundColor Yellow; Write-Host ''; $statusMsg = '' }

            $lastLetter = [char]([byte][char]'A' + $count - 1)
            $navFwd  = if ($pageIdx -lt $pages.Count - 1) { 'Enter=Next' } else { 'Enter=Apply' }
            $navBack = if ($pageIdx -gt 0) { '  Back=Prev' } else { '' }
            Write-Host ('  [A-{0}]=Toggle  |  {1}{2}  |  0=Skip-all  |  Esc=Cancel: ' -f $lastLetter, $navFwd, $navBack) -NoNewline -ForegroundColor Cyan

            # Read one key. Nav is encoded as a sentinel string so no letter key is
            # ever reserved for navigation (fixes A/N/P/S conflicts in the old design).
            $action = 'UNKNOWN'
            if ($useRawKey) {
                try {
                    $keyInfo = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                    $vk = $keyInfo.VirtualKeyCode
                    $ch = $keyInfo.Character.ToString().ToUpper()
                    if     ($vk -eq 27)           { Write-Host ''; $action = 'CANCEL' }
                    elseif ($vk -eq 13)           { Write-Host ''; $action = 'NEXT'   }
                    elseif ($vk -in @(8, 37))     { Write-Host ''; $action = 'PREV'   }   # Backspace or Left arrow
                    elseif ($ch -eq '0')          { Write-Host '0'; $action = 'SKIP'  }
                    elseif ($ch -match '^[A-Z]$') { Write-Host $ch; $action = $ch     }
                    else                          { Write-Host '' }
                } catch {
                    $useRawKey = $false
                }
            }
            if (-not $useRawKey -and $action -eq 'UNKNOWN') {
                $typed = (Read-Host '').Trim().ToUpper()
                if     ($typed -eq '')             { $action = 'NEXT'   }
                elseif ($typed -in @('-','<','B')) { $action = 'PREV'   }
                elseif ($typed -eq '0')            { $action = 'SKIP'   }
                elseif ($typed -eq 'Q')            { $action = 'CANCEL' }
                elseif ($typed -match '^[A-Z]$')   { $action = $typed   }
            }

            switch ($action) {
                'NEXT' {
                    if ($pageIdx -lt $pages.Count - 1) { $pageIdx++; $lastName = '' }
                    else { return $sel }
                }
                'PREV' {
                    if ($pageIdx -gt 0) { $pageIdx--; $lastName = '' }
                    else { $statusMsg = 'Already on first page' }
                }
                'SKIP' {
                    foreach ($n in $sel.Keys.Clone()) { $sel[$n] = [bool]$catalog[$n].Default }
                    Write-MyOutput 'Advanced configuration skipped — using defaults'
                    return $sel
                }
                'CANCEL' {
                    Write-MyOutput 'Advanced configuration cancelled — continuing with defaults'
                    return $null
                }
                default {
                    if ($action -match '^[A-Z]$') {
                        $idx = [byte][char]$action - [byte][char]'A'
                        if ($idx -ge 0 -and $idx -lt $count) {
                            $targetName = $pageNames[$idx]
                            $sel[$targetName] = -not $sel[$targetName]
                            $lastName = $targetName
                        } else {
                            $statusMsg = "No item on key '$action' — valid range A-$lastLetter"
                        }
                    }
                }
            }
        }
    }

    function Invoke-AdvancedConfigurationPrompt {
        # Offers the Advanced Configuration menu with a 60-second auto-skip (default = skip).
        # Autopilot / non-interactive: returns immediately without prompting.
        # Returns $true if the menu was shown and settings saved, $false if skipped.
        if ($State['Autopilot'] -or -not [Environment]::UserInteractive) { return $false }
        if ($State.ContainsKey('SuppressAdvancedPrompt') -and $State['SuppressAdvancedPrompt']) { return $false }

        $timeoutSec = 60
        Write-Host ''
        Write-Host ('  Configure advanced options? [y/N] (auto-skip in {0}s) ' -f $timeoutSec) -NoNewline -ForegroundColor Cyan

        $deadline = (Get-Date).AddSeconds($timeoutSec)
        $answer   = ''
        while ((Get-Date) -lt $deadline) {
            if ($host.UI.RawUI.KeyAvailable) {
                $k = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                if ($k.VirtualKeyCode -eq 13 -or $k.VirtualKeyCode -eq 27) { break }
                $answer = $k.Character.ToString().ToUpper()
                Write-Host $answer
                break
            }
            Start-Sleep -Milliseconds 200
            $remaining = [int]([Math]::Ceiling(($deadline - (Get-Date)).TotalSeconds))
            Write-Progress -Id 2 -Activity 'Advanced configuration prompt' -Status ('Auto-skip in {0}s — press Y to configure, N/Enter to skip' -f $remaining) -SecondsRemaining $remaining
        }
        Write-Progress -Id 2 -Activity 'Advanced configuration prompt' -Completed

        if ($answer -ne 'Y') {
            Write-MyOutput 'Advanced configuration skipped — continuing with defaults'
            return $false
        }
        Write-Host ''

        $result = Show-AdvancedMenu
        if ($null -eq $result) {
            Write-MyOutput 'Advanced configuration cancelled — continuing with defaults'
            return $false
        }

        $State['AdvancedFeatures'] = $result
        Save-State $State
        $changed = @($result.Keys | Where-Object { $result[$_] -ne [bool](Get-AdvancedFeatureCatalog)[$_].Default }).Count
        Write-MyOutput ('Advanced configuration applied — {0} setting(s) differ from defaults' -f $changed)
        return $true
    }

    function Test-Feature {
        # Returns $true when an Advanced feature is enabled.
        # Precedence: $State['AdvancedFeatures'][Name] > catalog default.
        # Condition scriptblock (if any) is evaluated first — returns $false when not met,
        # regardless of the stored or default value. Prevents config-file bypass of
        # runtime-gated features (e.g. ShadowRedundancy without DAG, EnableTLS13 on WS2019).
        # Unknown names return $false (fail closed) and log a verbose warning.
        param([Parameter(Mandatory)][string]$Name)

        $catalog = Get-AdvancedFeatureCatalog
        if (-not $catalog.Contains($Name)) {
            Write-MyVerbose ("Test-Feature: unknown feature name '{0}' — returning `$false" -f $Name)
            return $false
        }

        $entry = $catalog[$Name]
        if ($entry.ContainsKey('Condition')) {
            try   { if (-not (& $entry.Condition)) { return $false } }
            catch { return $false }
        }

        $features = $State['AdvancedFeatures']
        if ($features -is [hashtable] -and $features.ContainsKey($Name)) {
            return [bool]$features[$Name]
        }
        return [bool]$entry.Default
    }

    function Show-InstallationMenu {
        # Interactive console menu. Returns a hashtable of all chosen settings, or $null if user cancelled.
        # Uses Read-Host for all input so it works reliably over RDP, Hyper-V console and Windows Terminal.

        $modes = @{
            1 = 'Exchange Server (Mailbox)'
            2 = 'Exchange Server (Edge Transport)  [not tested]'
            3 = 'Recipient Management Tools         [not tested]'
            4 = 'Exchange Management Tools only     [not tested]'
            5 = 'Recovery Mode                      [not tested]'
            6 = 'Standalone Optimize                [not tested]'
            7 = 'Generate Installation Document     [not tested]'
        }

        # Toggle definitions: Key=letter, Name=parameter name, Default=initial state
        # Main menu exposes installation-flow toggles only; ~55 hardening/tuning options
        # live in the Advanced Configuration menu (see Get-AdvancedFeatureCatalog).

        # Name = parameter/cfg key; Label = display text shown in menu
        $toggleDefs = [ordered]@{
            'A' = @{ Name='Autopilot';             Label='Autopilot (auto-reboot)';        Default=$true  }
            'B' = @{ Name='IncludeFixes';          Label='Install Exchange SU';            Default=$true  }
            'N' = @{ Name='PreflightOnly';         Label='Preflight only (no install)';    Default=$false }
            'R' = @{ Name='InstallWindowsUpdates'; Label='Install Windows Updates';        Default=$true  }
            'U' = @{ Name='GenerateDoc';           Label='Generate Installation Document'; Default=$false }
            'V' = @{ Name='German';                Label='Language:  DE (default EN)';     Default=$false }
        }

        # Toggles disabled per mode (letters that cannot be toggled in that mode)
        $disabledToggles = @{
            1 = @()
            2 = @('U','V')                                    # Edge: no installation doc
            3 = @('B','N','R','U','V')                        # Recipient Mgmt: only Autopilot
            4 = @('B','U','V')                                # Mgmt Tools: no setup, no doc
            5 = @()
            6 = @('B','N','R')                                # Standalone Optimize: no setup, no WU, no preflight
            7 = @('A','B','N','R','U')                        # Document-only: only language matters
        }

        # Initialize toggle states from defaults
        $toggleState = @{}
        foreach ($k in $toggleDefs.Keys) { $toggleState[$k] = $toggleDefs[$k].Default }

        $selectedMode = 0

        # Returns extra letters that should be disabled based on current toggle state
        function Get-DynamicDisabled {
            param([hashtable]$TS)
            $extra = @()
            if ($TS['N'])      { $extra += @('B','R') }   # PreflightOnly: SU/WU irrelevant
            if (-not $TS['U']) { $extra += 'V' }          # V (language) only meaningful when doc is generated
            return $extra
        }

        function Write-MenuLine {
            param([string]$Line, [System.ConsoleColor]$Color = [System.ConsoleColor]::White)
            Write-Host $Line -ForegroundColor $Color
        }

        function Draw-Menu {
            param([int]$Mode, [hashtable]$ToggState, [string]$StatusMsg = '', [array]$ExtraDisabled = @(), [int]$AdvCount = 0)
            Clear-Host
            Write-MenuLine ('=' * 60) Cyan
            Write-MenuLine ('  EXpress v{0}  —  Copilot' -f $ScriptVersion) Cyan
            Write-MenuLine ('=' * 60) Cyan
            Write-Host ''
            Write-MenuLine '  Installation Mode:' Yellow
            for ($i = 1; $i -le 7; $i++) {
                $marker = if ($Mode -eq $i) { '>' } else { ' ' }
                $color  = if ($Mode -eq $i) { [System.ConsoleColor]::Green } else { [System.ConsoleColor]::Gray }
                Write-Host ('    [{0}] {1}  {2}' -f $i, $marker, $modes[$i]) -ForegroundColor $color
            }
            Write-Host ''
            Write-MenuLine '  Switches (press letter to toggle, C=Advanced, then ENTER to start):' Yellow

            $disabled = @(if ($Mode -gt 0) { $disabledToggles[$Mode] } else { @() }) + $ExtraDisabled
            $letters  = @($toggleDefs.Keys)
            # Render two columns
            for ($r = 0; $r -lt [Math]::Ceiling($letters.Count / 2); $r++) {
                $left  = $letters[$r]
                $right = $letters[$r + [Math]::Ceiling($letters.Count / 2)]
                $leftDis  = $disabled -contains $left
                $rightDis = $right -and ($disabled -contains $right)
                $leftVal  = if ($ToggState[$left])  { 'X' } else { ' ' }
                $rightVal = if ($right -and $ToggState[$right]) { 'X' } else { ' ' }
                $leftStr  = '  [{0}] {1,-28} [{2}]' -f $left,  $toggleDefs[$left].Label,  $leftVal
                $rightStr = if ($right) { '   [{0}] {1,-28} [{2}]' -f $right, $toggleDefs[$right].Label, $rightVal } else { '' }
                $lColor = if ($leftDis)  { [System.ConsoleColor]::DarkGray } else { [System.ConsoleColor]::White }
                $rColor = if ($rightDis) { [System.ConsoleColor]::DarkGray } else { [System.ConsoleColor]::White }
                Write-Host $leftStr  -ForegroundColor $lColor -NoNewline
                Write-Host $rightStr -ForegroundColor $rColor
            }
            # Advanced Configuration shortcut
            $advStatus = if ($AdvCount -gt 0) { "($AdvCount customized)" } else { '(defaults)' }
            Write-Host ('  [C] Advanced Configuration...          {0}' -f $advStatus) -ForegroundColor Cyan

            Write-Host ''
            if ($StatusMsg) { Write-Host "  $StatusMsg" -ForegroundColor Yellow }
        }

        Write-MyVerbose 'Menu: Show-InstallationMenu started'

        # Advanced Configuration state — populated when user presses C.
        # Starts empty (@{}) so Test-Feature falls back to catalog defaults.
        $advancedFeatures = @{}

        # --- Step 1: Mode selection ---
        while ($selectedMode -lt 1 -or $selectedMode -gt 7) {
            Draw-Menu -Mode $selectedMode -ToggState $toggleState -AdvCount $advancedFeatures.Count
            $raw = Read-Host '  Mode [1-7]'
            if ($raw -match '^[1-7]$') {
                $selectedMode = [int]$raw
                Write-MyVerbose ('Menu: Mode {0} selected ({1})' -f $selectedMode, $modes[$selectedMode])
                # Apply mode-specific toggle defaults
                switch ($selectedMode) {
                    2 { $toggleState['G'] = $false; $toggleState['I'] = $false }
                    3 { foreach ($k in $disabledToggles[3]) { $toggleState[$k] = $false } }
                    6 { foreach ($k in $disabledToggles[6]) { $toggleState[$k] = $false } }
                    7 { foreach ($k in $disabledToggles[7]) { $toggleState[$k] = $false } }
                }
            }
        }

        # --- Step 2: Toggle switches ---
        # Try RawUI.ReadKey (no Enter needed); fall back to Read-Host if console is not interactive
        # (e.g. stdin redirected, PS2Exe without console, or restricted host).
        $useRawKey = $false
        try {
            $null = $host.UI.RawUI.KeyAvailable  # throws if RawUI is not available
            $useRawKey = $true
        } catch { }

        $statusMsg = ''
        while ($true) {
            $dynDisabled = Get-DynamicDisabled $toggleState
            Draw-Menu -Mode $selectedMode -ToggState $toggleState -StatusMsg $statusMsg -ExtraDisabled $dynDisabled -AdvCount $advancedFeatures.Count
            $statusMsg = ''

            if ($useRawKey) {
                Write-Host '  Press letter to toggle, C=Advanced, ENTER to start: ' -NoNewline -ForegroundColor Cyan
                try {
                    $keyInfo = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                    $vk  = $keyInfo.VirtualKeyCode
                    $raw = $keyInfo.Character.ToString().ToUpper()
                    Write-Host $raw  # echo the pressed key
                    if ($vk -eq 13) { break }                          # Enter
                    if ($vk -eq 27) { return $null }                   # Escape = cancel
                } catch {
                    # RawUI failed mid-session — fall back
                    $useRawKey = $false
                    $raw = (Read-Host '').Trim().ToUpper()
                    if ($raw -eq '') { break }
                }
            }
            else {
                $raw = (Read-Host '  Toggle letter / C=Advanced / ENTER to start').Trim().ToUpper()
                if ($raw -eq '') { break }
            }

            if ($raw -eq 'C') {
                # Open Advanced Configuration menu; preserve state across multiple C-presses.
                $advResult = Show-AdvancedMenu -InitialValues $advancedFeatures
                if ($null -ne $advResult) {
                    $advancedFeatures = $advResult
                    $changed = ($advancedFeatures.Keys | Where-Object { $advancedFeatures[$_] -ne [bool](Get-AdvancedFeatureCatalog)[$_].Default }).Count
                    Write-MyVerbose ('Menu: Advanced Configuration applied — {0} setting(s) differ from defaults' -f $changed)
                }
            }
            elseif ($raw.Length -eq 1 -and $toggleDefs.Contains($raw)) {
                $dynNow = Get-DynamicDisabled $toggleState
                if (($disabledToggles[$selectedMode] -contains $raw) -or ($dynNow -contains $raw)) {
                    $statusMsg = "[$raw] is not available in this configuration"
                }
                else {
                    $toggleState[$raw] = -not $toggleState[$raw]
                    $toggState = if ($toggleState[$raw]) { 'ON' } else { 'OFF' }
                    Write-MyVerbose ('Menu: Toggle [{0}] {1} -> {2}' -f $raw, $toggleDefs[$raw].Label, $toggState)
                    # Reset any toggles that became disabled by this change
                    $dynAfter = Get-DynamicDisabled $toggleState
                    foreach ($x in $dynAfter) {
                        if ($toggleState[$x]) {
                            $toggleState[$x] = $false
                            Write-MyVerbose ('Menu: Toggle [{0}] auto-cleared (now disabled)' -f $x)
                        }
                    }
                }
            }
            elseif ($raw.Length -gt 0) {
                $statusMsg = "Unknown key '$raw' — press a listed letter, C=Advanced, or ENTER to start"
            }
        }

        # --- Step 3: String inputs (context-dependent) ---
        Clear-Host
        Write-MenuLine ('=' * 60) Cyan
        Write-MenuLine ("  EXpress v{0} - Mode: {1}" -f $ScriptVersion, $modes[$selectedMode]) Cyan
        Write-MenuLine ('=' * 60) Cyan
        Write-Host ''
        Write-MenuLine '  Enter values (leave blank for default, shown in [brackets]):' Yellow
        Write-Host ''

        function Read-MenuInput {
            param(
                [string]$Prompt,
                [string]$Default = '',
                [bool]$Required = $false,
                [scriptblock]$Validate = $null,
                [string]$ValidateMessage = 'Invalid input — please try again'
            )
            while ($true) {
                if ($Default) {
                    Write-Host -NoNewline ("  {0} " -f $Prompt)
                    Write-Host -NoNewline ("[{0}]: " -f $Default) -ForegroundColor Green
                    $val = Read-Host
                } else {
                    $val = Read-Host ("  {0}" -f $Prompt)
                }
                if ($val -eq '') { $val = $Default }
                if ($Required -and -not $val) {
                    Write-Host '  (required — cannot be empty)' -ForegroundColor Yellow
                }
                elseif ($val -and $Validate -and -not (& $Validate $val)) {
                    Write-Host "  $ValidateMessage" -ForegroundColor Yellow
                }
                else { return $val }
            }
        }

        $validateFQDN = { param($v) $v -match '^[a-zA-Z0-9]([a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?(\.[a-zA-Z0-9]([a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?)+$' }
        $validateCIDRList = {
            param($v)
            ($v -split '\s*,\s*') | Where-Object { $_ } | ForEach-Object {
                $_ -match '^\d{1,3}(\.\d{1,3}){3}(/([0-9]|[12]\d|3[0-2]))?$'
            } | Where-Object { -not $_ } | Measure-Object | Select-Object -ExpandProperty Count | ForEach-Object { $_ -eq 0 }
        }

        $cfg = @{}
        $cfg['Mode']       = $selectedMode
        $cfg['InstallPath'] = if ($ScriptFullName) { Split-Path $ScriptFullName -Parent } else { $PWD.Path }
        if ($selectedMode -notin @(6, 7)) {
            $defaultIso = Join-Path $cfg['InstallPath'] 'sources\ExchangeServerSE-x64.iso'
            $srcTry = 0
            while ($true) {
                $srcPath = Read-MenuInput -Prompt 'Exchange source (folder or .iso)' -Default $defaultIso -Required $true
                if (Test-Path $srcPath) { $cfg['SourcePath'] = $srcPath; break }
                $srcTry++
                Write-Host ("  Path not found: {0}" -f $srcPath) -ForegroundColor Yellow
                if ($srcTry -ge 3) {
                    Write-Host '  3 failed attempts — returning to main menu.' -ForegroundColor Red
                    return $null
                }
                Write-Host ("  Attempt {0}/3 — verify the path and try again." -f $srcTry) -ForegroundColor Yellow
            }
        }

        if ($selectedMode -eq 1) {
            # Detect existing Exchange organisation from AD (requires domain connectivity)
            $detectedOrg = ''
            try {
                $configNC  = ([ADSI]'LDAP://RootDSE').configurationNamingContext
                $searcher  = New-Object System.DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC")
                $searcher.Filter = '(objectClass=msExchOrganizationContainer)'
                $searcher.PropertiesToLoad.Add('name') | Out-Null
                $result = $searcher.FindOne()
                if ($result) { $detectedOrg = $result.Properties['name'][0] }
            } catch { }

            if ($detectedOrg) {
                Write-Host ("  Existing Exchange organisation detected: {0}" -f $detectedOrg) -ForegroundColor Green
                $cfg['Organization'] = Read-MenuInput -Prompt 'Organization name      (ENTER = keep existing)' -Default $detectedOrg
            } else {
                Write-Host '  No existing Exchange organisation found in AD.' -ForegroundColor Yellow
                # Require an org name — cannot install into a new org without a name
                $orgInput = ''
                while (-not $orgInput) {
                    $orgInput = (Read-Host '  Organization name      (required for new org)').Trim()
                    if (-not $orgInput) {
                        Write-Host '  Organisation name is required when no existing organisation is found. Enter Q to quit.' -ForegroundColor Yellow
                        if ($orgInput -imatch '^[Qq]$') { return $null }
                    }
                }
                $cfg['Organization'] = $orgInput
            }

            # Detect current Autodiscover SCP URL from AD
            $currentSCP = ''
            try {
                $configNC2 = ([ADSI]'LDAP://RootDSE').configurationNamingContext
                $scpSearch = New-Object System.DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC2")
                $scpSearch.Filter = "(&(cn=$($env:COMPUTERNAME))(objectClass=serviceConnectionPoint)(serviceClassName=ms-Exchange-AutoDiscover-Service))"
                $scpSearch.PropertiesToLoad.Add('serviceBindingInformation') | Out-Null
                $scpResult = $scpSearch.FindOne()
                if ($scpResult) { $currentSCP = $scpResult.Properties['serviceBindingInformation'][0] }
            } catch { }

            $cfg['MDBName']          = Read-MenuInput -Prompt 'Mailbox DB name        (blank = default name)'
            $cfg['MDBDBPath']        = Read-MenuInput -Prompt 'Mailbox DB path        (blank = Exchange default)'
            $cfg['MDBLogPath']       = Read-MenuInput -Prompt 'Mailbox log path       (blank = Exchange default)'
            if ($currentSCP) {
                Write-Host ("  Current Autodiscover SCP: {0}" -f $currentSCP) -ForegroundColor DarkGray
                $cfg['SCP']          = Read-MenuInput -Prompt 'Autodiscover SCP URL   (ENTER = keep current, - = remove)' -Default $currentSCP
            } else {
                $cfg['SCP']          = Read-MenuInput -Prompt 'Autodiscover SCP URL   (blank = let Setup set, - = remove)'
            }
            $cfg['TargetPath']       = Read-MenuInput -Prompt 'Exchange install path  (blank = C:\Program Files\Microsoft\Exchange Server\V15)'
            $knownDAGs = Get-ExchangeDAGNames
            if ($knownDAGs.Count -gt 0) {
                Write-Host ("  DAGs found in AD: {0}" -f ($knownDAGs -join ', ')) -ForegroundColor DarkGray
                $cfg['DAGName'] = Read-MenuInput -Prompt ('DAG name               ({0}, blank = no DAG join)' -f ($knownDAGs -join ' / ')) -Default ($knownDAGs[0])
            } else {
                $cfg['DAGName'] = Read-MenuInput -Prompt 'DAG name               (blank = no DAG join)'
            }
            $cfg['CopyServerConfig'] = Read-MenuInput -Prompt 'Copy config from server (FQDN, blank = none) [not tested]' -Validate $validateFQDN -ValidateMessage 'Not a valid FQDN (e.g. ex01.contoso.com)'
            $cfg['CertificatePath']  = Read-MenuInput -Prompt 'PFX certificate path   (blank = none)        [not tested]'
            $cfg['Namespace']        = Read-MenuInput -Prompt 'Access namespace       (e.g. mail.contoso.com, blank = skip URL config)' -Validate $validateFQDN -ValidateMessage 'Not a valid FQDN (e.g. mail.contoso.com)'
            if ($cfg['Namespace']) {
                # Default mail domain = parent of access namespace (drop leftmost label)
                $defaultMailDomain = ($cfg['Namespace'] -split '\.', 2)[1]
                if ($defaultMailDomain -notmatch '\.') { $defaultMailDomain = $cfg['Namespace'] }
                $cfg['MailDomain']     = Read-MenuInput -Prompt 'Mail domain             (e.g. contoso.com — for Accepted Domain + email addresses)' -Default $defaultMailDomain -Validate $validateFQDN -ValidateMessage 'Not a valid domain (e.g. contoso.com)'
                $cfg['DownloadDomain'] = Read-MenuInput -Prompt 'OWA download domain     (e.g. download.contoso.com, blank = skip CVE-2021-1730)' -Validate $validateFQDN -ValidateMessage 'Not a valid FQDN (e.g. download.contoso.com)'
            }
            if ((Read-MenuInput -Prompt 'Enable log cleanup task? [Y/N]' -Default 'Y') -imatch '^[Yy]$') {
                $retDays = Read-MenuInput -Prompt 'Log retention days' -Default '30' -Required $true
                $cfg['LogRetentionDays'] = [int]$retDays
                $cfg['LogCleanupFolder'] = Read-MenuInput -Prompt 'Log cleanup script folder' -Default 'C:\#service'
            } else {
                $cfg['LogRetentionDays'] = 0
                $cfg['LogCleanupFolder'] = ''
            }
            if ((Read-MenuInput -Prompt 'Create relay connectors? [Y/N]' -Default 'N') -imatch '^[Yy]$') {
                $relay = Read-MenuInput -Prompt 'Internal relay subnets  (comma-separated CIDRs, blank = placeholder)' -Validate $validateCIDRList -ValidateMessage 'Invalid format — use e.g. 192.168.1.0/24,10.0.0.5'
                $cfg['RelaySubnets'] = if ($relay) { $relay -split '\s*,\s*' | Where-Object { $_ } } else { @('192.0.2.1/32') }
                $extRelay = Read-MenuInput -Prompt 'External relay subnets  (comma-separated CIDRs, blank = placeholder)' -Validate $validateCIDRList -ValidateMessage 'Invalid format — use e.g. 192.168.2.0/24,10.0.1.5'
                $cfg['ExternalRelaySubnets'] = if ($extRelay) { $extRelay -split '\s*,\s*' | Where-Object { $_ } } else { @('192.0.2.2/32') }
            } else {
                $cfg['RelaySubnets'] = @()
                $cfg['ExternalRelaySubnets'] = @()
            }
        }
        elseif ($selectedMode -eq 2) {
            $cfg['EdgeDNSSuffix'] = Read-MenuInput -Prompt 'Edge DNS suffix (e.g. edge.contoso.com)' -Required $true -Validate $validateFQDN -ValidateMessage 'Not a valid FQDN (e.g. edge.contoso.com)'
            $cfg['TargetPath']    = Read-MenuInput -Prompt 'Exchange install path  (blank = Exchange default)'
        }
        elseif ($selectedMode -eq 3) {
            $cfg['RecipientMgmtCleanup'] = (Read-MenuInput -Prompt 'Run AD cleanup after install? [Y/N]' -Default 'N') -imatch '^[Yy]'
        }
        elseif ($selectedMode -eq 7) {
            # Mode 7 always generates the doc; language is picked via toggle V (see toggleDefs above).
            $custInput = Read-MenuInput -Prompt 'Redact sensitive values for customer? [Y/N]' -Default 'N'
            $cfg['CustomerDocument'] = ($custInput -imatch '^[Yy]$')
        }
        elseif ($selectedMode -eq 6) {
            $cfg['Namespace']        = Read-MenuInput -Prompt 'Access namespace       (e.g. mail.contoso.com, blank = skip URL config)' -Validate $validateFQDN -ValidateMessage 'Not a valid FQDN (e.g. mail.contoso.com)'
            if ($cfg['Namespace']) {
                $defaultMailDomain2 = ($cfg['Namespace'] -split '\.', 2)[1]
                if ($defaultMailDomain2 -notmatch '\.') { $defaultMailDomain2 = $cfg['Namespace'] }
                $cfg['MailDomain']     = Read-MenuInput -Prompt 'Mail domain             (e.g. contoso.com — for Accepted Domain + email addresses)' -Default $defaultMailDomain2 -Validate $validateFQDN -ValidateMessage 'Not a valid domain (e.g. contoso.com)'
                $cfg['DownloadDomain'] = Read-MenuInput -Prompt 'OWA download domain     (e.g. download.contoso.com, blank = skip CVE-2021-1730)' -Validate $validateFQDN -ValidateMessage 'Not a valid FQDN (e.g. download.contoso.com)'
            }
            $cfg['CertificatePath']  = Read-MenuInput -Prompt 'PFX certificate path   (blank = none)        [not tested]'
            $knownDAGs2 = Get-ExchangeDAGNames
            if ($knownDAGs2.Count -gt 0) {
                Write-Host ("  DAGs found in AD: {0}" -f ($knownDAGs2 -join ', ')) -ForegroundColor DarkGray
                $cfg['DAGName'] = Read-MenuInput -Prompt ('DAG name               ({0}, blank = no DAG join)' -f ($knownDAGs2 -join ' / ')) -Default ($knownDAGs2[0])
            } else {
                $cfg['DAGName'] = Read-MenuInput -Prompt 'DAG name               (blank = no DAG join)'
            }
            if ((Read-MenuInput -Prompt 'Enable log cleanup task? [Y/N]' -Default 'Y') -imatch '^[Yy]$') {
                $retDays = Read-MenuInput -Prompt 'Log retention days' -Default '30' -Required $true
                $cfg['LogRetentionDays'] = [int]$retDays
                $cfg['LogCleanupFolder'] = Read-MenuInput -Prompt 'Log cleanup script folder' -Default 'C:\#service'
            } else {
                $cfg['LogRetentionDays'] = 0
                $cfg['LogCleanupFolder'] = ''
            }
            if ((Read-MenuInput -Prompt 'Create relay connectors? [Y/N]' -Default 'N') -imatch '^[Yy]$') {
                $relay = Read-MenuInput -Prompt 'Internal relay subnets  (comma-separated CIDRs, blank = placeholder)' -Validate $validateCIDRList -ValidateMessage 'Invalid format — use e.g. 192.168.1.0/24,10.0.0.5'
                $cfg['RelaySubnets'] = if ($relay) { $relay -split '\s*,\s*' | Where-Object { $_ } } else { @('192.0.2.1/32') }
                $extRelay = Read-MenuInput -Prompt 'External relay subnets  (comma-separated CIDRs, blank = placeholder)' -Validate $validateCIDRList -ValidateMessage 'Invalid format — use e.g. 192.168.2.0/24,10.0.1.5'
                $cfg['ExternalRelaySubnets'] = if ($extRelay) { $extRelay -split '\s*,\s*' | Where-Object { $_ } } else { @('192.0.2.2/32') }
            } else {
                $cfg['RelaySubnets'] = @()
                $cfg['ExternalRelaySubnets'] = @()
            }
        }

        # Copy toggle values into cfg
        foreach ($k in $toggleDefs.Keys) {
            $cfg[$toggleDefs[$k].Name] = $toggleState[$k]
        }
        # Advanced Configuration — persisted so state-assignment can pick it up via Test-Feature.
        $cfg['AdvancedFeatures'] = $advancedFeatures

        # --- Step 4: Summary + confirmation ---
        # Build ordered list of editable fields per mode for the E=Edit path.
        # Each entry: Key (cfg hashtable key), Label (display), Prompt (Read-Host text),
        #             Validate (scriptblock or $null), ValidateMsg, Required.
        $editFields = [System.Collections.Generic.List[hashtable]]::new()
        if ($selectedMode -in @(1, 6)) {
            if ($selectedMode -eq 1) {
                $editFields.Add(@{ Key='SourcePath';    Label='Exchange source';      Prompt='Exchange source (folder or .iso)';                               Required=$true;  Validate={ param($v) Test-Path $v }; ValidateMsg='Path not found — enter a valid folder or .iso file path' })
                $editFields.Add(@{ Key='Organization';  Label='Organization name';    Prompt='Organization name';                                              Required=$false; Validate=$null;         ValidateMsg='' })
                $editFields.Add(@{ Key='MDBName';       Label='Mailbox DB name';      Prompt='Mailbox DB name        (blank = default name)';                  Required=$false; Validate=$null;         ValidateMsg='' })
                $editFields.Add(@{ Key='MDBDBPath';     Label='Mailbox DB path';      Prompt='Mailbox DB path        (blank = Exchange default)';              Required=$false; Validate=$null;         ValidateMsg='' })
                $editFields.Add(@{ Key='MDBLogPath';    Label='Mailbox log path';     Prompt='Mailbox log path       (blank = Exchange default)';              Required=$false; Validate=$null;         ValidateMsg='' })
                $editFields.Add(@{ Key='SCP';           Label='Autodiscover SCP URL'; Prompt='Autodiscover SCP URL   (blank = let Setup set, - = remove)';    Required=$false; Validate=$null;         ValidateMsg='' })
                $editFields.Add(@{ Key='TargetPath';    Label='Exchange install path';Prompt='Exchange install path  (blank = C:\Program Files\Microsoft\Exchange Server\V15)'; Required=$false; Validate=$null; ValidateMsg='' })
                $editFields.Add(@{ Key='DAGName';       Label='DAG name';             Prompt='DAG name               (blank = no DAG join)';                  Required=$false; Validate=$validateFQDN; ValidateMsg='Not a valid FQDN (e.g. dag01.contoso.com)' })
                $editFields.Add(@{ Key='CertificatePath'; Label='PFX certificate';   Prompt='PFX certificate path   (blank = none)';                          Required=$false; Validate=$null;         ValidateMsg='' })
            }
            $editFields.Add(@{ Key='Namespace';      Label='Access Namespace';        Prompt='Access namespace       (e.g. mail.contoso.com, blank = skip URL config)'; Required=$false; Validate=$validateFQDN; ValidateMsg='Not a valid FQDN (e.g. mail.contoso.com)' })
            $editFields.Add(@{ Key='MailDomain';     Label='Mail domain';             Prompt='Mail domain             (e.g. contoso.com — for Accepted Domain + email addresses)'; Required=$false; Validate=$validateFQDN; ValidateMsg='Not a valid domain (e.g. contoso.com)' })
            $editFields.Add(@{ Key='DownloadDomain'; Label='OWA download domain';     Prompt='OWA download domain    (e.g. download.contoso.com, blank = skip CVE-2021-1730)'; Required=$false; Validate=$validateFQDN; ValidateMsg='Not a valid FQDN (e.g. download.contoso.com)' })
        }

        while ($true) {
            Clear-Host
            Write-MenuLine ('=' * 60) Cyan
            Write-MenuLine '  Summary' Cyan
            Write-MenuLine ('=' * 60) Cyan
            Write-Host ''
            Write-Host ('  Mode    : {0}' -f $modes[$selectedMode]) -ForegroundColor Green
            if ($cfg['SourcePath'])    { Write-Host ('  Source  : {0}' -f $cfg['SourcePath']) }
            Write-Host                   ('  Install : {0}' -f $cfg['InstallPath'])
            if ($cfg['Organization'])  { Write-Host ('  Org     : {0}' -f $cfg['Organization']) }
            if ($cfg['MDBName'])       { Write-Host ('  MDB     : {0}' -f $cfg['MDBName']) }
            if ($cfg['MDBDBPath'])     { Write-Host ('  DB Path : {0}' -f $cfg['MDBDBPath']) }
            if ($cfg['MDBLogPath'])    { Write-Host ('  Log Path: {0}' -f $cfg['MDBLogPath']) }
            if ($cfg['SCP'])           { Write-Host ('  SCP     : {0}' -f $cfg['SCP']) }
            if ($cfg['TargetPath'])    { Write-Host ('  Target  : {0}' -f $cfg['TargetPath']) }
            if ($cfg['DAGName'])       { Write-Host ('  DAG     : {0}' -f $cfg['DAGName']) }
            if ($cfg['Namespace'])     { Write-Host ('  Namespace: {0}' -f $cfg['Namespace']) -ForegroundColor Cyan }
            if ($cfg['MailDomain'])    { Write-Host ('  MailDomain: {0}' -f $cfg['MailDomain']) -ForegroundColor Cyan }
            if ($cfg['DownloadDomain']){ Write-Host ('  DL Domain: {0}' -f $cfg['DownloadDomain']) }
            if ($cfg['CertificatePath']){ Write-Host ('  Cert    : {0}' -f $cfg['CertificatePath']) }
            if ($cfg['EdgeDNSSuffix']) { Write-Host ('  Edge DNS: {0}' -f $cfg['EdgeDNSSuffix']) }
            # Active switches
            $finalDisabled  = @($disabledToggles[$selectedMode]) + (Get-DynamicDisabled $toggleState)
            $activeToggles = ($toggleDefs.Keys | Where-Object { $toggleState[$_] -and ($finalDisabled -notcontains $_) }) -join ', '
            if ($activeToggles) { Write-Host ('  Switches: {0}' -f $activeToggles) }
            Write-Host ''

            $editHint = if ($editFields.Count -gt 0) { ' / E=edit a field' } else { '' }
            $confirm = Read-Host ("  Start? [Y=yes{0} / N=back to menu / Q=quit]" -f $editHint)

            if ($confirm -imatch '^[Yy]') { return $cfg }
            if ($confirm -imatch '^[Qq]') { return $null }

            if ($confirm -imatch '^[Ee]' -and $editFields.Count -gt 0) {
                # Show numbered list of editable fields with current values
                Write-Host ''
                Write-Host '  Edit a field — current values:' -ForegroundColor Cyan
                for ($fi = 0; $fi -lt $editFields.Count; $fi++) {
                    $fld  = $editFields[$fi]
                    $fval = if ($cfg[$fld.Key]) { $cfg[$fld.Key] } else { '(empty)' }
                    Write-Host ('  {0,2}.  {1,-24} : {2}' -f ($fi + 1), $fld.Label, $fval)
                }
                Write-Host ''
                $pick = (Read-Host '  Field number (ENTER = cancel)').Trim()
                if ($pick -match '^\d+$') {
                    $idx = [int]$pick - 1
                    if ($idx -ge 0 -and $idx -lt $editFields.Count) {
                        $fld      = $editFields[$idx]
                        $curVal   = if ($cfg[$fld.Key]) { $cfg[$fld.Key] } else { '' }
                        $valMsg   = if ($fld.ValidateMsg) { $fld.ValidateMsg } else { 'Invalid input' }
                        $newVal   = Read-MenuInput -Prompt $fld.Prompt -Default $curVal -Required $fld.Required -Validate $fld.Validate -ValidateMessage $valMsg
                        $cfg[$fld.Key] = $newVal
                        # Clear DownloadDomain if Namespace was cleared
                        if ($fld.Key -eq 'Namespace' -and -not $newVal) { $cfg['DownloadDomain'] = '' }
                    }
                }
                continue
            }

            Write-MyVerbose 'Menu: Back to mode selection'
            # N or anything else = restart from mode selection
            $selectedMode = 0
            while ($selectedMode -lt 1 -or $selectedMode -gt 7) {
                Draw-Menu -Mode $selectedMode -ToggState $toggleState
                $raw = Read-Host '  Mode [1-7]'
                if ($raw -match '^[1-7]$') { $selectedMode = [int]$raw }
            }
        }
    }
    ########################################
    # MAIN
    ########################################

    #Requires -Version 5.1

    # Pin verbose/debug preferences early — before any cmdlet that would emit stream 4/5 output
    # (Get-CimInstance below is a stream-4 spammer when -Verbose was passed on the command line,
    # which also happens on Autopilot RunOnce resume when the original launch used -Verbose).
    # Our custom Write-MyVerbose / Write-MyDebug write to the log via $State['LogVerbose'/'LogDebug']
    # flags, set a few lines further down — decoupled from $VerbosePreference.
    $VerbosePreference = 'SilentlyContinue'
    $DebugPreference   = 'SilentlyContinue'

    # $EXpressEntryScript is set in EXpress.ps1 before the dot-source loop so it always
    # points to EXpress.ps1 itself. $MyInvocation.MyCommand.Path inside a dot-sourced file
    # resolves to that module file's path — using it here would break the Autopilot RunOnce key.
    # PS2Exe: both are empty; fall back to the process image path.
    $ScriptFullName = if ($EXpressEntryScript) {
        $EXpressEntryScript
    } elseif ($MyInvocation.MyCommand.Path) {
        $MyInvocation.MyCommand.Path
    } else {
        [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    }
    # Detect PS2Exe compiled run: MyCommand.Path is empty; Write-Progress is not rendered visually
    $IsPS2Exe = -not $MyInvocation.MyCommand.Path
    $ScriptName = $ScriptFullName.Split("\")[-1]
    if (-not $PSBoundParameters.ContainsKey('InstallPath')) {
        $InstallPath = Split-Path $ScriptFullName -Parent
    }
    $ParameterString = $PSBoundParameters.getEnumerator() -join " "
    $OSVersionParts = (Get-CimInstance -ClassName Win32_OperatingSystem).Version.Split('.')
    $MajorOSVersion = '{0}.{1}' -f $OSVersionParts[0], $OSVersionParts[1]
    $MinorOSVersion = $OSVersionParts[2]
    $FullOSVersion  = '{0}.{1}' -f $MajorOSVersion, $MinorOSVersion

    $State = @{}
    $StateFile = "$InstallPath\$($env:computerName)_EXpress_State.xml"
    $State = Restore-State
    # Ensure reports folder exists on Autopilot resume (state restored from XML)
    if ($State['ReportsPath'] -and -not (Test-Path $State['ReportsPath'])) {
        New-Item -Path $State['ReportsPath'] -ItemType Directory -Force | Out-Null
    }

    $BackgroundJobs = @()

    Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
        Stop-BackgroundJobs
    } | Out-Null
    trap {
        Write-MyWarning 'Script termination detected, cleaning up background jobs...'
        Stop-BackgroundJobs
        break
    }

    # --- Logging bootstrap (must run BEFORE any Write-MyOutput/Write-MyVerbose so pre-menu
    #     messages land in the single log file).
    $script:lastErrorCount = $Error.Count
    # Capture fresh-vs-resume BEFORE we populate $State; otherwise the later `if ($script:isFreshStart)`
    # check would mis-classify a fresh start as a resume once LogVerbose/LogDebug/TranscriptFile
    # have been seeded.
    $script:isFreshStart = ($State.Count -eq 0)
    $boundVerbose = $PSBoundParameters.ContainsKey('Verbose') -and [bool]$PSBoundParameters['Verbose']
    $boundDebug   = $PSBoundParameters.ContainsKey('Debug')   -and [bool]$PSBoundParameters['Debug']
    # $VerbosePreference / $DebugPreference have already been pinned to SilentlyContinue at the
    # top of process{} (before the first Get-CimInstance). Our Write-MyVerbose / Write-MyDebug
    # wrappers append to the log via the State flags below — independent of the preference vars.
    $State['LogVerbose'] = $boundVerbose -or $boundDebug -or [bool]$State['LogVerbose']
    $State['LogDebug']   = $boundDebug   -or [bool]$State['LogDebug']

    # Seed TranscriptFile early so pre-menu / Test-Preflight lines land in the same file.
    # On Autopilot resume the restored state already carries TranscriptFile; the init block
    # below reuses it instead of creating a second timestamped file.
    $earlyReports = if ($State['ReportsPath']) { $State['ReportsPath'] } else { Join-Path $InstallPath 'reports' }
    if (-not (Test-Path $earlyReports)) { New-Item -Path $earlyReports -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null }
    if (-not $State['TranscriptFile']) {
        $State['TranscriptFile'] = Join-Path $earlyReports ('{0}_EXpress_Install_{1}.log' -f $env:computerName, (Get-Date -Format 'yyyyMMdd-HHmmss'))
    }
    if ($State['LogVerbose'] -or $State['LogDebug']) {
        $tier = if ($State['LogDebug']) { 'DEBUG' } else { 'VERBOSE' }
        try {
            $hdr = @(
                '',
                '==================================================================',
                ('  {0} session start: {1}' -f $tier, (Get-Date -Format u)),
                ('  PID {0}  User {1}\{2}  PS {3}  Host {4}' -f $PID, $env:USERDOMAIN, $env:USERNAME, $PSVersionTable.PSVersion, $Host.Name),
                ('  Invocation: {0}' -f ($MyInvocation.Line -replace '\s+', ' ').Trim()),
                '=================================================================='
            ) -join [Environment]::NewLine
            [System.IO.File]::AppendAllText($State['TranscriptFile'], ($hdr + "`r`n"), [System.Text.UTF8Encoding]::new($false))
        } catch {
            Write-Warning ('Could not write log header: {0}' -f $_.Exception.Message)
        }
    }

    # Now everything is in place: start normal logging (console + file).
    if ($State.Count -gt 0 -and -not $ParameterString -and -not $script:isFreshStart) {
        $ParameterString = '[resuming from phase {0}]' -f $State['InstallPhase']
    }
    Write-MyOutput  "Script $ScriptFullName v$ScriptVersion called using $ParameterString"
    # ParameterSetName is the internal binding set name — always "Autopilot" as default even
    # in Copilot launches. Keep for forensics but at DEBUG tier so it doesn't confuse normal logs.
    Write-MyDebug   "ParameterSet used for binding: $($PsCmdlet.ParameterSetName)"
    Write-MyOutput  ('Running on OS build {0}' -f $FullOSVersion)
    if ($State['LogVerbose'] -or $State['LogDebug']) {
        $tierLabel = if ($State['LogDebug']) { 'DEBUG' } else { 'VERBOSE' }
        Write-MyOutput ('Log tier: {0} - file: {1}' -f $tierLabel, $State['TranscriptFile'])
    }

    # --- v5.93: MEAC Split-Permissions standalone prep mode ------------------------
    # Runs before any Exchange-specific checks so it works on a non-Exchange AD-admin
    # box. Downloads MEAC, invokes -PrepareADForAutomationOnly, exits. No state, no
    # phase loop, no reboot.
    if ($PsCmdlet.ParameterSetName -eq 'MEACPrepareAD') {
        Write-MyOutput ('MEAC Split-Permissions AD preparation mode (domain: {0})' -f $MEACADAccountDomain)
        if (-not (Test-Admin)) {
            Write-MyError 'Must run elevated (Administrator) to create AD accounts.'
            exit $ERR_RUNNINGNONADMINMODE
        }
        $meacPrepPath = Join-Path $env:TEMP 'MonitorExchangeAuthCertificate.ps1'
        $meacPrepUrl  = 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/MonitorExchangeAuthCertificate.ps1'
        try {
            Invoke-WebDownload -Uri $meacPrepUrl -OutFile $meacPrepPath
            Write-MyVerbose ('MEAC downloaded, SHA256: {0}' -f (Get-FileHash $meacPrepPath -Algorithm SHA256).Hash)
        }
        catch {
            Write-MyError ('Could not download MonitorExchangeAuthCertificate.ps1: {0}' -f $_.Exception.Message)
            exit $ERR_MEACPREPAREAD
        }
        try {
            & $meacPrepPath -PrepareADForAutomationOnly -ADAccountDomain $MEACADAccountDomain -Confirm:$false *>&1 |
                ForEach-Object { Write-MyOutput ('MEAC: {0}' -f $_) }
        }
        catch {
            Write-MyError ('MEAC PrepareAD failed: {0}' -f $_.Exception.Message)
            exit $ERR_MEACPREPAREAD
        }
        Write-MyOutput ''
        Write-MyOutput 'MEAC automation account prepared in AD.'
        Write-MyOutput 'Hand the SystemMailbox{b963af59-3975-4f92-9d58-ad0b1fe3a1a3} credential to the'
        Write-MyOutput 'Exchange administrator, who passes it via -MEACAutomationCredential during the'
        Write-MyOutput 'normal install.'
        exit $ERR_OK
    }

    if ($script:isFreshStart) {
        # No state, initialize settings from parameters.
        # When started interactively with no meaningful parameters (default Autopilot set, no bound params
        # other than the defaults), show the interactive installation menu.
        $isInteractiveStart = [Environment]::UserInteractive -and
                              ($PsCmdlet.ParameterSetName -eq 'Autopilot') -and
                              ($PSBoundParameters.Keys | Where-Object { $_ -notin @('InstallPath','Verbose','Debug') }).Count -eq 0

        if ($isInteractiveStart) {
            # Auto-detect config.psd1 in the same folder as the script / compiled .exe
            if (-not $ConfigFile) {
                $autoConfigPath = Join-Path (Split-Path $ScriptFullName -Parent) 'config.psd1'
                if (Test-Path $autoConfigPath -PathType Leaf) {
                    Write-Host ("Found 'config.psd1' in script folder ({0})." -f (Split-Path $ScriptFullName -Parent)) -ForegroundColor Cyan
                    $useAuto = (Read-Host 'Use this configuration file? [Y=yes / N=show menu]').Trim().ToUpper()
                    if ($useAuto -eq 'Y') {
                        $ConfigFile = $autoConfigPath
                        Write-MyOutput ("Auto-detected configuration loaded: {0}" -f $ConfigFile)
                    }
                }
            }
        }

        if ($isInteractiveStart -and -not $ConfigFile) {
            $menuResult = Show-InstallationMenu
            if (-not $menuResult) {
                Write-Output 'Installation cancelled.'
                exit $ERR_OK
            }
            # Map menu result back to parameter-equivalent variables so the standard state init below can run
            $mode            = $menuResult['Mode']
            $SourcePath      = $menuResult['SourcePath']
            $InstallPath     = if ($menuResult['InstallPath']) { $menuResult['InstallPath'] } else { Split-Path $ScriptFullName -Parent }
            $Organization    = $menuResult['Organization']
            $MDBName         = $menuResult['MDBName']
            $MDBDBPath       = $menuResult['MDBDBPath']
            $MDBLogPath      = $menuResult['MDBLogPath']
            $SCP             = if ($menuResult['SCP']) { $menuResult['SCP'] } else { '' }
            $TargetPath      = $menuResult['TargetPath']
            $DAGName         = $menuResult['DAGName']
            $CopyServerConfig    = $menuResult['CopyServerConfig']
            $CertificatePath     = $menuResult['CertificatePath']
            $EdgeDNSSuffix       = $menuResult['EdgeDNSSuffix']
            $Autopilot           = [switch]($menuResult['Autopilot'])
            $IncludeFixes        = [switch]($menuResult['IncludeFixes'])
            $PreflightOnly       = [switch]($menuResult['PreflightOnly'])
            $InstallWindowsUpdates   = [switch]($menuResult['InstallWindowsUpdates'])
            # Non-menu toggles (C/D/E/… removed in v5.95) — values come from Advanced
            # Configuration menu (Invoke-AdvancedConfigurationPrompt) or $ConfigFile.
            # Param variables retain their on-cmdline / catalog defaults here.
            $InstallEdge         = [switch]($mode -eq 2)
            $Recover             = [switch]($mode -eq 5)
            $StandaloneOptimize  = [switch]($mode -eq 6)
            $StandaloneDocument  = [switch]($mode -eq 7)
            # Toggle U "Generate Installation Document" → invert to NoWordDoc.
            # Mode 7 always generates the doc regardless of the toggle.
            $NoWordDoc           = [switch](-not ([bool]$menuResult['GenerateDoc'] -or $mode -eq 7))
            $German              = [switch]([bool]$menuResult['German'])
            $CustomerDocument    = [switch]([bool]$menuResult['CustomerDocument'])
            $NoSetup             = [switch]($false)
            $InstallRecipientManagement = [switch]($mode -eq 3)
            $InstallManagementTools     = [switch]($mode -eq 4)
            $RecipientMgmtCleanup = [switch]($menuResult['RecipientMgmtCleanup'])
            $Namespace           = $menuResult['Namespace']
            $DownloadDomain      = $menuResult['DownloadDomain']
            $LogRetentionDays    = if ($menuResult['LogRetentionDays']) { [int]$menuResult['LogRetentionDays'] } else { 0 }
            $RelaySubnets        = $menuResult['RelaySubnets']
            $ExternalRelaySubnets = $menuResult['ExternalRelaySubnets']
            # Reload state file path with potentially updated InstallPath
            $StateFile = "$InstallPath\$($env:computerName)_EXpress_State.xml"

            # Log confirmed menu selection (here in the caller so Write-MyOutput / Write-MyVerbose
            # don't pollute Show-InstallationMenu's return pipeline with extra string values).
            $modeLabel = @{1='Exchange Server (Mailbox)';2='Edge Transport';3='Recipient Mgmt';4='Mgmt Tools';5='Recovery';6='Standalone Optimize';7='Installation Document'}
            Write-MyOutput '================ Menu selection confirmed ================'
            Write-MyOutput ('Menu: Mode     : {0}' -f $modeLabel[$menuResult['Mode']])
            Write-MyOutput ('Menu: Source   : {0}' -f $SourcePath)
            Write-MyOutput ('Menu: Install  : {0}' -f $InstallPath)
            if ($Organization)    { Write-MyOutput ('Menu: Org      : {0}' -f $Organization) }
            if ($Namespace)       { Write-MyOutput ('Menu: Namespace: {0}' -f $Namespace) }
            if ($DownloadDomain)  { Write-MyOutput ('Menu: DL Domain: {0}' -f $DownloadDomain) }
            if ($DAGName)         { Write-MyOutput ('Menu: DAG      : {0}' -f $DAGName) }
            if ($CertificatePath) { Write-MyOutput ('Menu: Cert PFX : {0}' -f $CertificatePath) }
            # Active switch labels (same computation the menu summary used)
            $menuSwitchMap = @{
                'Autopilot'=$Autopilot;'IncludeFixes'=$IncludeFixes;
                'PreflightOnly'=$PreflightOnly;'InstallWindowsUpdates'=$InstallWindowsUpdates;
                'GenerateDoc'=(-not [bool]$NoWordDoc);'German'=$German
            }
            $activeSwitches = ($menuSwitchMap.Keys | Where-Object { $menuSwitchMap[$_] }) -join ', '
            if ($activeSwitches) { Write-MyOutput ('Menu: Switches : {0}' -f $activeSwitches) }
            # Full cfg dump at verbose/debug tier (credentials redacted)
            $safeDump = @{}
            foreach ($k in $menuResult.Keys) {
                $safeDump[$k] = if ($k -in 'Credentials','CertificatePassword','AdminPassword') { '<redacted>' } else { $menuResult[$k] }
            }
            foreach ($k in ($safeDump.Keys | Sort-Object)) {
                Write-MyVerbose ('Menu: cfg[{0}] = {1}' -f $k, $safeDump[$k])
            }

            # Advanced Configuration — user configured via C in the main menu.
            # Seed $State['AdvancedFeatures'] now so Test-Feature picks it up in the
            # state-assignment block below.
            $State['AdvancedFeatures'] = if ($menuResult['AdvancedFeatures'] -is [hashtable]) { $menuResult['AdvancedFeatures'] } else { @{} }
        }
        elseif ($ConfigFile) {
            # Headless mode: load all parameters from a .psd1 config file.
            # The menu is automatically skipped when -ConfigFile is specified.
            Write-MyOutput "Loading configuration from $ConfigFile"
            $cfg = Import-PowerShellDataFile -Path $ConfigFile -ErrorAction Stop

            # Helper: read a value from the config, or keep the current parameter value
            function Get-CfgValue { param($Key, $Current) if ($cfg.ContainsKey($Key)) { $cfg[$Key] } else { $Current } }

            # Paths
            $SourcePath   = Get-CfgValue 'SourcePath'   $SourcePath
            $InstallPath  = if (Get-CfgValue 'InstallPath' $InstallPath) { Get-CfgValue 'InstallPath' $InstallPath } else { Split-Path $ScriptFullName -Parent }

            # Exchange config
            $Organization     = Get-CfgValue 'Organization'     $Organization
            $MDBName          = Get-CfgValue 'MDBName'          $MDBName
            $MDBDBPath        = Get-CfgValue 'MDBDBPath'        $MDBDBPath
            $MDBLogPath       = Get-CfgValue 'MDBLogPath'       $MDBLogPath
            $SCP              = Get-CfgValue 'SCP'              $SCP
            $TargetPath       = Get-CfgValue 'TargetPath'       $TargetPath
            $DAGName          = Get-CfgValue 'DAGName'          $DAGName
            $CopyServerConfig = Get-CfgValue 'CopyServerConfig' $CopyServerConfig
            $CertificatePath  = Get-CfgValue 'CertificatePath'  $CertificatePath
            $EdgeDNSSuffix    = Get-CfgValue 'EdgeDNSSuffix'    $EdgeDNSSuffix

            # Installation mode
            $InstallEdge                = [switch](Get-CfgValue 'InstallEdge'                ([bool]$InstallEdge))
            $Recover                    = [switch](Get-CfgValue 'Recover'                    ([bool]$Recover))
            $NoSetup                    = [switch](Get-CfgValue 'NoSetup'                    ([bool]$NoSetup))
            $InstallRecipientManagement = [switch](Get-CfgValue 'InstallRecipientManagement' ([bool]$InstallRecipientManagement))
            $InstallManagementTools     = [switch](Get-CfgValue 'InstallManagementTools'     ([bool]$InstallManagementTools))
            $RecipientMgmtCleanup       = [switch](Get-CfgValue 'RecipientMgmtCleanup'       ([bool]$RecipientMgmtCleanup))

            # Security / TLS switches
            $Autopilot      = [switch](Get-CfgValue 'Autopilot'      ([bool]$Autopilot))
            $IncludeFixes   = [switch](Get-CfgValue 'IncludeFixes'   ([bool]$IncludeFixes))
            $DisableSSL3    = [switch](Get-CfgValue 'DisableSSL3'    ([bool]$DisableSSL3))
            $DisableRC4     = [switch](Get-CfgValue 'DisableRC4'     ([bool]$DisableRC4))
            $EnableECC      = [switch](Get-CfgValue 'EnableECC'      ([bool]$EnableECC))
            $NoCBC          = [switch](Get-CfgValue 'NoCBC'          ([bool]$NoCBC))
            $EnableAMSI     = [switch](Get-CfgValue 'EnableAMSI'     ([bool]$EnableAMSI))
            $EnableTLS12    = [switch](Get-CfgValue 'EnableTLS12'    ([bool]$EnableTLS12))
            $EnableTLS13    = [switch](Get-CfgValue 'EnableTLS13'    ([bool]$EnableTLS13))
            $DoNotEnableEP  = [switch](Get-CfgValue 'DoNotEnableEP'  ([bool]$DoNotEnableEP))
            $DiagnosticData = [switch](Get-CfgValue 'DiagnosticData' ([bool]$DiagnosticData))

            # Options
            $Lock                 = [switch](Get-CfgValue 'Lock'                 ([bool]$Lock))
            $SkipRolesCheck       = [switch](Get-CfgValue 'SkipRolesCheck'       ([bool]$SkipRolesCheck))
            $PreflightOnly        = [switch](Get-CfgValue 'PreflightOnly'        ([bool]$PreflightOnly))
            $NoCheckpoint         = [switch](Get-CfgValue 'NoCheckpoint'         ([bool]$NoCheckpoint))
            $SkipHealthCheck      = [switch](Get-CfgValue 'SkipHealthCheck'      ([bool]$SkipHealthCheck))
            $NoNet481             = [switch](Get-CfgValue 'NoNet481'             ([bool]$NoNet481))
            $InstallWindowsUpdates = [switch](Get-CfgValue 'InstallWindowsUpdates' ([bool]$InstallWindowsUpdates))
            $SkipWindowsUpdates   = [switch](Get-CfgValue 'SkipWindowsUpdates'   ([bool]$SkipWindowsUpdates))
            $SkipSetupAssist      = [switch](Get-CfgValue 'SkipSetupAssist'       ([bool]$SkipSetupAssist))
            $Namespace            = Get-CfgValue 'Namespace'      $Namespace
            $MailDomain           = Get-CfgValue 'MailDomain'     $MailDomain
            $DownloadDomain       = Get-CfgValue 'DownloadDomain' $DownloadDomain
            $RunEOMT              = [switch](Get-CfgValue 'RunEOMT'              ([bool]$RunEOMT))
            $WaitForADSync        = [switch](Get-CfgValue 'WaitForADSync'        ([bool]$WaitForADSync))
            $LogRetentionDays     = Get-CfgValue 'LogRetentionDays' $LogRetentionDays
            $RelaySubnets         = Get-CfgValue 'RelaySubnets'         $RelaySubnets
            $ExternalRelaySubnets = Get-CfgValue 'ExternalRelaySubnets' $ExternalRelaySubnets
            $NoWordDoc            = [switch](Get-CfgValue 'NoWordDoc'        ([bool]$NoWordDoc))
            $CustomerDocument     = [switch](Get-CfgValue 'CustomerDocument' ([bool]$CustomerDocument))
            # Config-file back-compat: legacy 'Language=DE' still maps to $German; 'German=true' also accepted.
            $cfgLang   = Get-CfgValue 'Language' $null
            $cfgGerman = Get-CfgValue 'German'   $null
            if ($cfgGerman -ne $null) { $German = [switch][bool]$cfgGerman }
            elseif ($cfgLang)         { $German = [switch]($cfgLang -imatch '^DE$') }
            $DocumentScope        = Get-CfgValue 'DocumentScope'  $DocumentScope
            $IncludeServers       = @((Get-CfgValue 'IncludeServers' ($IncludeServers -join ',')) -split ',' | Where-Object { $_ })
            $TemplatePath         = Get-CfgValue 'TemplatePath'   $TemplatePath

            # MEAC passthroughs (v5.93)
            $MEACIgnoreHybridConfig       = [switch](Get-CfgValue 'MEACIgnoreHybridConfig'       ([bool]$MEACIgnoreHybridConfig))
            $MEACIgnoreUnreachableServers = [switch](Get-CfgValue 'MEACIgnoreUnreachableServers' ([bool]$MEACIgnoreUnreachableServers))
            $MEACNotificationEmail        = Get-CfgValue 'MEACNotificationEmail' $MEACNotificationEmail

            # -- Plain-text install-admin credential (UNATTENDED ONLY, v5.92) --------------
            # SECURITY: AdminUser/AdminPassword may be supplied in plain text to enable
            # fully unattended deployment with zero operator input. The config file MUST
            # be deleted or scrubbed immediately after the install completes. Plain-text
            # from the file is converted to PSCredential here; the rest of the script
            # never sees the literal string. State persistence remains DPAPI-encrypted
            # (user+machine bound) as before.
            $cfgAdminUser = Get-CfgValue 'AdminUser'     $null
            $cfgAdminPw   = Get-CfgValue 'AdminPassword' $null
            if ($cfgAdminUser -and $cfgAdminPw -and -not $Credentials) {
                $sec = ConvertTo-SecureString -String $cfgAdminPw -AsPlainText -Force
                $Credentials = New-Object System.Management.Automation.PSCredential($cfgAdminUser, $sec)
                Write-MyWarning ''
                Write-MyWarning '################################################################'
                Write-MyWarning '##  SECURITY WARNING: PLAIN-TEXT CREDENTIALS IN CONFIG FILE   ##'
                Write-MyWarning '################################################################'
                Write-MyWarning ('Config file: {0}' -f $ConfigFile)
                Write-MyWarning 'contains a plain-text AdminPassword.'
                Write-MyWarning 'Acceptable ONLY for short-lived, unattended installation runs.'
                Write-MyWarning 'DELETE OR SCRUB the config file IMMEDIATELY after install completes.'
                Write-MyWarning 'Do not archive, commit to version control, copy to a share, or email.'
                Write-MyWarning 'EXpress state file persists credentials as DPAPI (user+machine bound),'
                Write-MyWarning 'but the config file itself remains a plain-text artefact on disk.'
                Write-MyWarning '################################################################'
                Write-MyWarning ''
            }

            # Advanced Configuration — nested block in .psd1, with backwards-compat
            # for legacy top-level keys. Precedence: nested > top-level > catalog default.
            if (-not ($State['AdvancedFeatures'] -is [hashtable])) { $State['AdvancedFeatures'] = @{} }
            $catalogNames = (Get-AdvancedFeatureCatalog).Keys
            # Legacy top-level keys (backwards compat)
            foreach ($name in $catalogNames) {
                if ($cfg.ContainsKey($name) -and -not $State['AdvancedFeatures'].ContainsKey($name)) {
                    $State['AdvancedFeatures'][$name] = [bool]$cfg[$name]
                }
            }
            # Nested AdvancedFeatures = @{ ... } block (new canonical form; wins over top-level)
            if ($cfg.ContainsKey('AdvancedFeatures') -and $cfg['AdvancedFeatures'] -is [hashtable]) {
                foreach ($k in $cfg['AdvancedFeatures'].Keys) {
                    if ($catalogNames -contains $k) {
                        $State['AdvancedFeatures'][$k] = [bool]$cfg['AdvancedFeatures'][$k]
                    } else {
                        Write-MyWarning ("Config AdvancedFeatures.{0}: unknown feature name — ignored" -f $k)
                    }
                }
            }

            # Recalculate state file path with potentially overridden InstallPath
            $StateFile = "$InstallPath\$($env:computerName)_EXpress_State.xml"
            Write-MyOutput "Configuration loaded: mode=$(if ($InstallEdge){'Edge'}elseif($Recover){'Recovery'}else{'Mailbox'}), source=$SourcePath, org=$Organization"
        }
        elseif ( $($PsCmdlet.ParameterSetName) -eq "Autopilot") {
            Write-Error "Running in Autopilot mode but no state file present"
            exit $ERR_AUTOPILOTNOSTATEFILE
        }

        $State["InstallMailbox"] = $True
        $State["InstallEdge"] = $InstallEdge
        $State["InstallMDBDBPath"] = $MDBDBPath
        $State["InstallMDBLogPath"] = $MDBLogPath
        $State["InstallMDBName"] = $MDBName
        $State["InstallPhase"] = 0
        $State["InstallingUser"] = [Security.Principal.WindowsIdentity]::GetCurrent().Name
        $State["OrganizationName"] = $Organization
        $State["AdminAccount"] = if ($Credentials) { $Credentials.UserName } else { $null }
        $State["AdminPassword"] = if ($Credentials) { ($Credentials.Password | ConvertFrom-SecureString -ErrorAction SilentlyContinue) } else { $null }
        # MEAC Split-Permissions passthrough (v5.93). Persists DPAPI-encrypted so it
        # survives the Autopilot reboot chain between Phase 0 intake and Phase 6
        # Register-AuthCertificateRenewal. Not populated in standard deployments.
        $State["MEACAutomationUser"] = if ($MEACAutomationCredential) { $MEACAutomationCredential.UserName } else { $null }
        $State["MEACAutomationPW"]   = if ($MEACAutomationCredential) { ($MEACAutomationCredential.Password | ConvertFrom-SecureString) } else { $null }
        if ( Get-DiskImage -ImagePath $SourcePath -ErrorAction SilentlyContinue) {
            $State['SourceImage'] = $SourcePath
            # Unblock ISO before mounting: on WS2022+ Windows propagates MOTW from the ISO container
            # to all files executed from it. Zone.Identifier ADS on the ISO itself must be removed first
            # because files inside UDF (ISO9660) cannot carry ADS and cannot be unblocked after mounting.
            if ( Get-Item -Path $SourcePath -Stream 'Zone.Identifier' -ErrorAction SilentlyContinue) {
                Write-MyOutput "ISO source has Zone.Identifier — unblocking before mount to prevent MOTW propagation"
                Unblock-File -Path $SourcePath
            }
            $State["SourcePath"] = Resolve-SourcePath -SourceImage $SourcePath
        }
        else {
            if ( $State['SourceImage']) {
                $State["SourcePath"] = Resolve-SourcePath -SourceImage $State['SourceImage']
            }
            else {
                $State['SourceImage'] = $null
                $State["SourcePath"] = $SourcePath
            }
        }
        $State["SetupVersion"] = ( Get-DetectedFileVersion "$($State["SourcePath"])\setup.exe")
        $State["TargetPath"] = $TargetPath
        $State["Autopilot"] = $Autopilot
        $State["ConfigDriven"] = [bool]$ConfigFile
        # Persist absolute ConfigFile path so reboot-resume can still log which config drove the run
        $State["ConfigFile"] = if ($ConfigFile) {
            try { (Resolve-Path -Path $ConfigFile -ErrorAction Stop).Path } catch { [string]$ConfigFile }
        } else { $null }
        $State["IncludeFixes"] = $IncludeFixes
        $State["NoSetup"] = $NoSetup
        $State["Recover"] = $Recover
        $State["Upgrade"] = $false
        $State["Install481"] = $False
        $State["VCRedist2012"] = $False
        $State["VCRedist2013"] = $False

        # -- Advanced Configuration catalog ---------------------------------------
        # Merge explicitly cmdline-bound switches into $State['AdvancedFeatures']
        # (backwards compat for scripts still passing -DisableSSL3, -EnableECC …).
        # Precedence already established: menu/config > cmdline-bound > catalog default.
        if (-not ($State['AdvancedFeatures'] -is [hashtable])) { $State['AdvancedFeatures'] = @{} }
        $advCatalog = Get-AdvancedFeatureCatalog
        foreach ($name in $advCatalog.Keys) {
            if ($PSBoundParameters.ContainsKey($name) -and -not $State['AdvancedFeatures'].ContainsKey($name)) {
                $v = Get-Variable -Name $name -ValueOnly -ErrorAction SilentlyContinue
                $State['AdvancedFeatures'][$name] = [bool]$v
            }
        }
        # Project every catalog entry to its flat $State[Name] so the rest of the
        # script (Phase 5 hardening, reports, HealthChecker gates …) keeps reading
        # $State['DisableSSL3'] etc. unchanged. Test-Feature applies precedence.
        foreach ($name in $advCatalog.Keys) { $State[$name] = Test-Feature $name }
        # Derived: EnableCBC is the logical inverse of NoCBC.
        $State["EnableCBC"]     = -not (Test-Feature 'NoCBC')
        if ($State["EnableTLS13"] -and -not $State["EnableTLS12"]) {
            Write-MyWarning 'EnableTLS13 requires EnableTLS12; automatically enabling TLS 1.2 enforcement'
            $State["EnableTLS12"] = $true
            $State['AdvancedFeatures']['EnableTLS12'] = $true
        }
        $State["DoNotEnableEP_FEEWS"] = $DoNotEnableEP_FEEWS
        $State["SCP"] = $SCP
        $State["EdgeDNSSuffix"] = $EdgeDNSSuffix
        $State["InstallPath"]  = $InstallPath
        $State["ReportsPath"]  = Join-Path $InstallPath 'reports'
        if (-not (Test-Path $State["ReportsPath"])) { New-Item -Path $State["ReportsPath"] -ItemType Directory -Force | Out-Null }
        # TranscriptFile may have been pre-seeded by the early block; keep that so pre-menu
        # messages end up in the same single log file.
        if (-not $State["TranscriptFile"]) {
            $State["TranscriptFile"] = Join-Path $State["ReportsPath"] ('{0}_EXpress_Install_{1}.log' -f $env:computerName, (Get-Date -Format 'yyyyMMdd-HHmmss'))
        }
        $State["PreflightOnly"] = $PreflightOnly
        $State["CopyServerConfig"] = $CopyServerConfig
        $State["CertificatePath"] = $CertificatePath
        $State["CertificatePassword"] = $null
        $State["DAGName"] = $DAGName
        $State["ServerConfigExportPath"] = $null
        $State["InstallRecipientManagement"] = [bool]$InstallRecipientManagement
        $State["InstallManagementTools"] = [bool]$InstallManagementTools
        $State["RecipientMgmtCleanup"] = [bool]$RecipientMgmtCleanup
        if ([bool]$InstallWindowsUpdates -and [bool]$SkipWindowsUpdates) {
            Write-MyWarning '-InstallWindowsUpdates and -SkipWindowsUpdates are both set; updates will be skipped'
        }
        $State["InstallWindowsUpdates"] = [bool]$InstallWindowsUpdates -and -not [bool]$SkipWindowsUpdates
        $State["SkipSetupAssist"] = $SkipSetupAssist
        $State["Namespace"]     = $Namespace
        $State["MailDomain"]    = $MailDomain
        $State["DownloadDomain"] = $DownloadDomain
        $State["LogRetentionDays"] = $LogRetentionDays
        $State["RelaySubnets"]         = $RelaySubnets
        $State["ExternalRelaySubnets"] = $ExternalRelaySubnets
        $State["StandaloneOptimize"]  = [bool]$StandaloneOptimize
        $State["SkipInstallReport"]   = [bool]$SkipInstallReport
        $State["StandaloneDocument"]  = [bool]$StandaloneDocument
        $State["NoWordDoc"]           = [bool]$NoWordDoc
        $State["CustomerDocument"]    = [bool]$CustomerDocument
        # English is the default; -German is the only opt-in for German output.
        # $State['Language'] stays as the internal flag ('EN'|'DE') to avoid touching
        # the L helper in New-InstallationDocument that reads it.
        $State["Language"]            = if ($German) { 'DE' } else { 'EN' }
        $State["DocumentScope"]       = if ($DocumentScope) { $DocumentScope } else { 'All' }
        $State["IncludeServers"]      = if ($IncludeServers) { $IncludeServers -join ',' } else { '' }
        $State["TemplatePath"]        = $TemplatePath

        # Prompt for PFX password at startup if certificate path specified
        if ($CertificatePath) {
            Write-MyOutput 'Certificate import requested, prompting for PFX password'
            $pfxPwd = Read-Host -Prompt 'Enter PFX password' -AsSecureString
            # ConvertFrom-SecureString without -Key uses DPAPI (user+machine bound).
            # Safe here: PFX import happens in Phase 5 on the same machine/user.
            $State["CertificatePassword"] = ($pfxPwd | ConvertFrom-SecureString)
        }

        # Store Server Manager state
        $State['DoNotOpenServerManagerAtLogon'] = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\ServerManager' -Name DoNotOpenServerManagerAtLogon -ErrorAction SilentlyContinue).DoNotOpenServerManagerAtLogon

        $State["Verbose"] = $VerbosePreference

    }
    else {
        # Run from saved parameters
        # ISO is only needed for phases 1-4 (setup); skip remount for phase 5+ to allow dismount after phase 4
        if ( $State['SourceImage'] -and $State['InstallPhase'] -lt 4) {
            # Mount ISO image, and set SourcePath to actual mounted location to anticipate drive letter changes
            $State["SourcePath"] = Resolve-SourcePath -SourceImage $State['SourceImage']
        }
    }

    if ( $State["Lock"] ) {
        LockScreen
    }

    if ($State['InstallPhase'] -le 1) {
        Clear-DesktopBackground
    }

    if ( $State.containsKey("LastSuccessfulPhase")) {
        Write-MyVerbose "Continuing from last successful phase $($State["InstallPhase"])"
        $State["InstallPhase"] = $State["LastSuccessfulPhase"]
    }
    if ( $PSBoundParameters.ContainsKey('Phase')) {
        Write-MyVerbose "Phase manually set to $Phase"
        $State["InstallPhase"] = $Phase
    }
    else {
        $State["InstallPhase"]++
    }

    $VerbosePreference = 'SilentlyContinue'

    # When skipping setup, limit no. of steps
    if ( $State["NoSetup"]) {
        $MAX_PHASE = 3
    }
    elseif ( $State["InstallRecipientManagement"] -or $State["InstallManagementTools"]) {
        # Recipient Management and Management Tools modes use a 3-phase flow
        $MAX_PHASE = 3
    }
    elseif ( $State["StandaloneOptimize"]) {
        $MAX_PHASE = 1
    }
    elseif ( $State["StandaloneDocument"]) {
        $MAX_PHASE = 1
    }
    else {
        $MAX_PHASE = 6
    }

    $runMode = if ($State['ConfigDriven']) { 'Autopilot (fully automated)' } else { 'Copilot (interactive)' }
    Write-MyOutput ('Mode: {0}' -f $runMode)
    if ($State['ConfigDriven']) {
        # Resolve and log the configuration file actually used (absolute path + metadata)
        $cfgResolved = if ($ConfigFile) { $ConfigFile } else { $State['ConfigFile'] }
        if ($cfgResolved) {
            try {
                $cfgItem = Get-Item -Path $cfgResolved -ErrorAction Stop
                $State['ConfigFile'] = $cfgItem.FullName
                Write-MyOutput ('Configuration: {0}' -f $cfgItem.FullName)
                Write-MyVerbose ('Configuration details: size={0} bytes, modified={1:u}' -f $cfgItem.Length, $cfgItem.LastWriteTimeUtc)
            } catch {
                Write-MyWarning ('Configuration file cannot be resolved: {0} ({1})' -f $cfgResolved, $_.Exception.Message)
            }
        } else {
            Write-MyWarning 'Autopilot mode active but no configuration file path recorded.'
        }
    }

    if ( $Autopilot -and $State["InstallPhase"] -gt 1) {
        # Wait a little before proceeding
        Write-MyOutput "Will continue unattended installation of Exchange in $COUNTDOWN_TIMER seconds .."
        for ($i = $COUNTDOWN_TIMER; $i -gt 0; $i--) {
            Write-Progress -Id 2 -Activity 'Autopilot resume' -Status ('Continuing in {0}s ...' -f $i) -PercentComplete (($COUNTDOWN_TIMER - $i) * 100 / $COUNTDOWN_TIMER)
            Start-Sleep -Seconds 1
        }
        Write-Progress -Id 2 -Activity 'Autopilot resume' -Completed
    }

    # Generate Pre-Flight Report (only on first phase or PreflightOnly mode)
    if ($State['InstallPhase'] -le 1 -or $State['PreflightOnly']) {
        New-Item -Path $State['InstallPath'] -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
        $preflightFailures = New-PreflightReport
        if ($State['PreflightOnly']) {
            Write-MyOutput 'PreflightOnly mode - exiting after report generation'
            if ($preflightFailures -gt 0) {
                Write-MyWarning ('{0} preflight check(s) failed - review the report' -f $preflightFailures)
            }
            exit $ERR_OK
        }
    }

    Test-Preflight

    Write-MyVerbose "Logging to $($State["TranscriptFile"])"

    if ($State['LogDebug']) {
        Write-MyVerbose ('PS {0} on {1} (64-bit: {2}) - PID {3} - User {4}\{5}' -f $PSVersionTable.PSVersion, [Environment]::OSVersion.VersionString, [Environment]::Is64BitProcess, $PID, $env:USERDOMAIN, $env:USERNAME)
        Write-MyVerbose 'Suppressed errors (-ErrorAction SilentlyContinue, 2>$null) will be flushed to the log tagged [SUPPRESSED-ERROR].'
    }

    # Get rid of the security dialog when spawning exe's etc.
    Disable-OpenFileSecurityWarning

    # Always disable autologon allowing you to "fix" things and reboot intermediately
    Disable-AutoLogon

    Write-MyOutput "Checking for pending reboot .."
    if ( Test-RebootPending ) {
        $State["InstallPhase"]--
        if ( $State["Autopilot"]) {
            Write-MyWarning "Reboot pending, will reboot system and rerun phase"
        }
        else {
            Write-MyError "Reboot pending, please reboot system and restart script (parameters will be saved)"
        }
    }
    else {

        Write-MyVerbose "Current phase is $($State["InstallPhase"]) of $MAX_PHASE"

        Write-MyVerbose 'Disabling Server Manager at logon'
        Set-RegistryValue -Path 'HKCU:\Software\Microsoft\ServerManager' -Name DoNotOpenServerManagerAtLogon -Value 1

        # Create System Restore checkpoint before each phase.
        # Checkpoint-Computer is only supported on client OS (ProductType=1).
        # It exists as a cmdlet on Windows Server but throws at runtime — check OS type first.
        if (-not $State['NoCheckpoint']) {
            $isClientOS = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).ProductType -eq 1
            if ($isClientOS) {
                try {
                    Checkpoint-Computer -Description ('Exchange Install Phase {0}' -f $State['InstallPhase']) -RestorePointType MODIFY_SETTINGS -ErrorAction Stop
                    Write-MyOutput ('System Restore checkpoint created for Phase {0}' -f $State['InstallPhase'])
                }
                catch {
                    Write-MyWarning ('Could not create System Restore checkpoint: {0}' -f $_.Exception.Message)
                }
            }
            else {
                Write-MyVerbose 'System Restore not supported on Windows Server — skipping checkpoint'
            }
        }

        if ($State["InstallRecipientManagement"]) {
            switch ($State["InstallPhase"]) {
                1 {
                    Write-MyOutput 'Recipient Management Tools - Phase 1: Installing prerequisites'
                    Install-RecipientManagementPrereqs
                    if ( Test-RebootPending) {
                        if ($State['Autopilot']) { Write-MyWarning 'Reboot pending, will reboot and continue' }
                        else { Write-MyOutput 'Reboot pending, please reboot and restart script' }
                    }
                }
                2 {
                    Write-MyOutput 'Recipient Management Tools - Phase 2: Installing Exchange Management Tools'
                    Install-RecipientManagement
                }
                3 {
                    Write-MyOutput 'Recipient Management Tools - Phase 3: Post-install configuration'
                    New-RecipientManagementShortcut
                    if ($State['RecipientMgmtCleanup']) {
                        Invoke-RecipientManagementADCleanup
                    }
                    Write-MyOutput 'Recipient Management Tools installation complete'
                }
                default {
                    Write-MyError "Unknown phase ($($State["InstallPhase"])) in RecipientManagement mode"
                }
            }
        }
        elseif ($State["InstallManagementTools"]) {
            switch ($State["InstallPhase"]) {
                1 {
                    Write-MyOutput 'Exchange Management Tools - Phase 1: Installing Windows prerequisites'
                    Install-ManagementToolsPrereqs
                    if ( Test-RebootPending) {
                        if ($State['Autopilot']) { Write-MyWarning 'Reboot pending, will reboot and continue' }
                        else { Write-MyOutput 'Reboot pending, please reboot and restart script' }
                    }
                }
                2 {
                    Write-MyOutput 'Exchange Management Tools - Phase 2: Installing runtime prerequisites'
                    Install-ManagementToolsRuntimePrereqs
                }
                3 {
                    Write-MyOutput 'Exchange Management Tools - Phase 3: Running Exchange setup /roles:ManagementTools'
                    Install-ManagementToolsOnly
                    Write-MyOutput 'Exchange Management Tools installation complete'
                }
                default {
                    Write-MyError "Unknown phase ($($State["InstallPhase"])) in ManagementTools mode"
                }
            }
        }
        elseif ($State["StandaloneDocument"]) {
            switch ($State["InstallPhase"]) {
                1 {
                    Write-MyOutput 'Standalone Document — generating Word installation document for existing Exchange server'
                    Import-ExchangeModule
                    # Ensure Defender realtime is on before we snapshot server state into the report.
                    Enable-DefenderRealtimeMonitoring -Force
                    try { New-InstallationDocument } catch { Write-MyWarning ('Word document failed: {0}' -f $_.Exception.Message) }
                    Write-MyOutput 'Installation document generation complete.'
                }
                default {
                    Write-MyError "Unknown phase ($($State["InstallPhase"])) in StandaloneDocument mode"
                }
            }
        }
        elseif ($State["StandaloneOptimize"]) {
            switch ($State["InstallPhase"]) {
                1 {
                    Write-MyOutput 'Standalone Optimize — running post-install optimizations on existing Exchange server'
                    Import-ExchangeModule

                    if ($State['Namespace']) {
                        Write-MyOutput 'Configuring Virtual Directory URLs'
                        Set-VirtualDirectoryURLs
                    }

                    Write-MyOutput 'Running Exchange optimizations'
                    Invoke-ExchangeOptimizations

                    if ($State['CertificatePath']) {
                        Import-ExchangeCertificateFromPFX
                        Set-HSTSHeader
                    }

                    if ($State['RelaySubnets'] -or $State['ExternalRelaySubnets']) {
                        New-AnonymousRelayConnector
                    }

                    # Access Namespace — Accepted Domain + Email Address Policy (F26)
                    if ($State['AccessNamespaceMail'] -and $State['Namespace'] -and $State['NewExchangeOrg']) {
                        Enable-AccessNamespaceMailConfig
                    }

                    Test-DBLogPathSeparation

                    Get-RBACReport

                    if (-not $State['SkipHealthCheck']) {
                        Invoke-HealthChecker
                    }

                    if ($State['LogRetentionDays']) {
                        Register-ExchangeLogCleanup
                    }

                    Write-MyOutput 'Standalone optimization complete.'
                }
                default {
                    Write-MyError "Unknown phase ($($State["InstallPhase"])) in StandaloneOptimize mode"
                }
            }
        }
        else {
        # Loop so Phase 2 can flow directly into Phase 3 when no reboot is pending
        # (e.g. WS2025 + Exchange SE where nothing reboot-relevant was installed).
        do {
            $continueInProcess = $false
        switch ($State["InstallPhase"]) {
            1 {

                if ( [System.Version]$FullOSVersion -ge [System.Version]$WS2016_MAJOR) {

                    Write-MyOutput ('Exchange setup detected: {0}' -f (Get-SetupTextVersion $State['SetupVersion']))
                    Write-MyOutput ('Operating System detected: {0}' -f (Get-OSVersionText $FullOSVersion))

                    if ( $State["NoNet481"]) {
                        Write-MyOutput "NoNet481 specified, will not install .NET Framework 4.8.1"
                        $State["Install481"] = $False
                    }
                    else {
                        if ([System.Version]$FullOSVersion -lt [System.Version]$WS2022_PREFULL ) {
                            Write-MyOutput ".NET Framework 4.8 required for this OS — will install if not present"
                            $State["Install481"] = $False
                        }
                        else {
                            Write-MyOutput ".NET Framework 4.8.1 required for this OS — will install if not present"
                            $State["Install481"] = $True
                        }
                    }

                    Write-MyOutput "Will install Visual C++ 2012 Runtime"
                    $State["VCRedist2012"] = $True

                    Write-MyOutput "Will install Visual C++ 2013 Runtime"
                    $State["VCRedist2013"] = $True

                }
                else {
                    Write-MyError ('Operating System version {0} not supported' -f $FullOSVersion)
                    exit $ERR_UNEXPECTEDOS
                }
                $phSw = [Diagnostics.Stopwatch]::StartNew()
                Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 1 of 6: Windows Features + .NET' -PercentComplete 0
                Disable-IEESC
                Disable-ServerManagerAtLogon
                Disable-DefenderRealtimeMonitoring
                Write-MyOutput "Installing Operating System prerequisites"
                Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 1 of 6: Installing Windows Features' -PercentComplete 10
                Install-WindowsFeatures $MajorOSVersion

                if ($State['CopyServerConfig']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 1 of 6: Exporting source server config' -PercentComplete 80
                    Export-SourceServerConfig $State['CopyServerConfig']
                }

                # Install pending Windows Updates before rebooting (if requested)
                Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 1 of 6: Windows Updates' -PercentComplete 90
                Install-PendingWindowsUpdates
                Write-MyVerbose ('Phase 1 completed in {0:F1}s' -f $phSw.Elapsed.TotalSeconds)
                Write-PhaseProgress -Activity 'Exchange Installation' -Completed
            }

            2 {
                $phSw = [Diagnostics.Stopwatch]::StartNew()
                Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 2 of 6: Prerequisites' -PercentComplete 0
                Write-MyOutput "Installing BITS module"
                Import-Module BITSTransfer

                # Check .NET FrameWork 4.8.1 needs to be installed
                Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 2 of 6: .NET Framework' -PercentComplete 10
                if ( $State["Install481"]) {

                    Remove-NETFrameworkInstallBlock '4.8.1' '-' '481'
                    if ( (Get-NETVersion) -lt $NETVERSION_481) {
                        Install-MyPackage "-" "Microsoft .NET Framework 4.8.1" "NDP481-x86-x64-AllOS-ENU.exe" "https://download.microsoft.com/download/4/b/2/cd00d4ed-ebdd-49ee-8a33-eabc3d1030e3/NDP481-x86-x64-AllOS-ENU.exe" ("/q", "/norestart")
                    }
                    else {
                        Write-MyOutput ".NET Framework 4.8.1 or later detected"
                    }
                }
                else {
                    Write-MyOutput ('Keeping current .NET Framework ({0})' -f (Get-NETVersion))
                    Set-NETFrameworkInstallBlock '4.8.1' '-' '481'
                }

                # OS specific hotfixes

                if ( [System.Version]$FullOSVersion -ge [System.Version]$WS2016_MAJOR -and [System.Version]$FullOSVersion -lt [System.Version]$WS2019_PREFULL) {
                    # WS2016
                    Install-MyPackage "KB3206632" "Cumulative Update for Windows Server 2016 for x64-based Systems" "windows10.0-kb3206632-x64_b2e20b7e1aa65288007de21e88cd21c3ffb05110.msu" "http://download.windowsupdate.com/d/msdownload/update/software/secu/2016/12/windows10.0-kb3206632-x64_b2e20b7e1aa65288007de21e88cd21c3ffb05110.msu" ("/quiet", "/norestart")
                }
                if ( [System.Version]$FullOSVersion -ge [System.Version]$WS2019_PREFULL -and [System.Version]$FullOSVersion -lt [System.Version]$WS2022_PREFULL) {
                    # WS2019
                }
                if ( [System.Version]$FullOSVersion -ge [System.Version]$WS2022_PREFULL -and [System.Version]$FullOSVersion -lt [System.Version]$WS2025_PREFULL) {
                    # WS2022
                }
                if ( [System.Version]$FullOSVersion -ge [System.Version]$WS2025_PREFULL) {
                    # WS2025
                }

                # Check if need to install VC++ Runtimes
                Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 2 of 6: Visual C++ Runtimes' -PercentComplete 50
                # VC++ 2012 (v11.0): required for Exchange 2016 CU23, Exchange 2019, SE and Edge Transport role (flagged by HealthChecker)
                if ( -not (Get-VCRuntime -version '11.0') -and $State["VCRedist2012"] ) {
                    Install-MyPackage "" "Visual C++ 2012 Redistributable" "vcredist_x64_2012.exe" "https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe" ("/install", "/quiet", "/norestart")
                    if ( -not (Get-VCRuntime -version '11.0')) {
                        Write-MyError 'Visual C++ 2012 Redistributable installation could not be verified — Exchange Setup will fail. Check the installer manually.'
                        exit $ERR_PROBLEMPACKAGESETUP
                    }
                }

                # VC++ 2013 (v12.0): required for Exchange 2016 CU23, 2019 and SE; minimum 12.0.40664 (KB4538461)
                if ( -not (Get-VCRuntime -version '12.0' -MinBuild '12.0.40664') -and $State["VCRedist2013"] ) {
                    # VC++ 2013 x64 12.0.40664.0 — High-DPI aware build (aka.ms/highdpimfc2013x64enu)
                    # This is the version HC checks for; the older GUID CDN URL delivers 12.0.40660.
                    Install-MyPackage "" "Visual C++ 2013 Redistributable" "vcredist_x64_2013.exe" "https://aka.ms/highdpimfc2013x64enu" ("/install", "/quiet", "/norestart")
                    if ( -not (Get-VCRuntime -version '12.0')) {
                        Write-MyError 'Visual C++ 2013 Redistributable installation could not be verified — Exchange Setup will fail. Check the installer manually.'
                        exit $ERR_PROBLEMPACKAGESETUP
                    }
                }

                # URL Rewrite module
                Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 2 of 6: URL Rewrite Module' -PercentComplete 80
                Install-MyPackage "{9BCA2118-F753-4A1E-BCF3-5A820729965C}" "URL Rewrite Module 2.1" "rewrite_amd64_en-US.msi" "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_en-US.msi" ("/quiet", "/norestart")
                Write-MyVerbose ('Phase 2 completed in {0:F1}s' -f $phSw.Elapsed.TotalSeconds)
                Write-PhaseProgress -Activity 'Exchange Installation' -Completed

            }

            3 {
                $phSw = [Diagnostics.Stopwatch]::StartNew()
                Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 3 of 6: Prerequisites (continued)' -PercentComplete 0
                if ( !($State['InstallEdge'])) {
                    Write-MyOutput "Installing Exchange prerequisites (continued)"
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 3 of 6: UCMA Runtime' -PercentComplete 20
                    if ( [System.Version]$FullOSVersion -ge [System.Version]$WS2019_PREFULL -and (Test-ServerCore) ) {
                        Install-MyPackage "{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" "Unified Communications Managed API 4.0 Runtime (Core)" "Setup.exe" (Join-Path -Path $State['SourcePath'] -ChildPath 'UcmaRedist\Setup.exe') ("/passive", "/norestart") -NoDownload
                    }
                    else {
                        Install-MyPackage "{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" "Unified Communications Managed API 4.0 Runtime" "UcmaRuntimeSetup.exe" "https://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe" ("/passive", "/norestart")
                    }
                }
                else {
                    Write-MyOutput 'Setting Primary DNS Suffix'
                    Set-EdgeDNSSuffix -DNSSuffix $State['EdgeDNSSuffix']
                }
                if ($State["OrganizationName"]) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 3 of 6: Checking Active Directory' -PercentComplete 60
                    $adPrepRan = Initialize-Exchange
                    if ($adPrepRan) {
                        Wait-ADReplication
                    }
                }
                Write-MyVerbose ('Phase 3 completed in {0:F1}s' -f $phSw.Elapsed.TotalSeconds)
                Write-PhaseProgress -Activity 'Exchange Installation' -Completed
            }

            4 {
                $phSw = [Diagnostics.Stopwatch]::StartNew()
                Write-MyOutput "Installing Exchange"
                Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 4 of 6: Running Exchange Setup (this may take 30-60 min)' -PercentComplete 0

                switch ( $State["SCP"]) {
                    '' {
                        # Do nothing
                    }
                    '-' {
                        Clear-AutodiscoverServiceConnectionPoint $ENV:COMPUTERNAME -Wait
                    }
                    default {
                        Set-AutodiscoverServiceConnectionPoint $ENV:COMPUTERNAME $State['SCP'] -Wait
                    }
                }

                $null = Start-DisableMSExchangeAutodiscoverAppPoolJob

                Install-EXpress_

                # Cleanup any background jobs
                Stop-BackgroundJobs
                Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 4 of 6: Configuring transport services' -PercentComplete 95

                if ( Get-Service MSExchangeTransport -ErrorAction SilentlyContinue) {
                    Write-MyOutput "Configuring MSExchangeTransport startup to Manual"
                    Set-Service MSExchangeTransport -StartupType Manual
                }
                if ( Get-Service MSExchangeFrontEndTransport -ErrorAction SilentlyContinue) {
                    Write-MyOutput "Configuring MSExchangeFrontEndTransport startup to Manual"
                    Set-Service MSExchangeFrontEndTransport -StartupType Manual
                }
                # Dismount ISO after Exchange setup — no longer needed for phases 5+
                if ($State['SourceImage']) {
                    Dismount-DiskImage -ImagePath $State['SourceImage'] | Out-Null
                    Write-MyVerbose ('Exchange setup complete — dismounted ISO: {0}' -f $State['SourceImage'])
                }
                Write-MyVerbose ('Phase 4 completed in {0:F1}s' -f $phSw.Elapsed.TotalSeconds)
                Write-PhaseProgress -Activity 'Exchange Installation' -Completed
            }

            5 {
                Write-MyOutput "Post-configuring"
                $p5Steps = @(
                    'Windows Defender exclusions', 'Power plan', 'NIC power management', 'Page file',
                    'TCP settings', 'SMBv1', 'Windows Search', 'WDigest', 'HTTP/2', 'TCP offload',
                    'Credential Guard', 'LM compatibility', 'LSA Protection', 'RSS / NIC queues',
                    'IPv4 over IPv6', 'NetBIOS on NICs', 'LLMNR', 'mDNS',
                    'MaxConcurrentAPI', 'Disk allocation', 'Scheduled tasks', 'Server Manager',
                    'CRL timeout', 'TLS / Schannel', 'Root CA auto-update', 'Exchange module + search tuning',
                    'Security hardening', 'Org/Transport optimizations', 'IANA timezone mapping',
                    'SSL offloading', 'Extended Protection', 'MRS Proxy', 'MAPI encryption',
                    'Certificate', 'HSTS header', 'Exchange SU', 'Server config import', 'EOMT'
                )
                $p5Total = $p5Steps.Count; $p5Step = 0
                $p5Sw = [Diagnostics.Stopwatch]::new(); $p5LastDesc = $null
                $script:p5NeedsIisRestart = $false   # set to $true by Enable-ECC/CBC/AMSI if they change anything; cleared after batched restart at end of Phase 5
                function Step-P5($desc) {
                    if ($script:p5LastDesc) {
                        Write-MyVerbose ('{0} took {1:F1}s' -f $script:p5LastDesc, $script:p5Sw.Elapsed.TotalSeconds)
                    }
                    $script:p5LastDesc = $desc
                    $script:p5Sw.Restart()
                    $script:p5Step++
                    Write-PhaseProgress -Id 1 -Activity 'Phase 5 of 6: Post-configuration' -Status $desc -PercentComplete ($script:p5Step * 100 / $script:p5Total)
                }

                Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 5 of 6: Post-configuration' -PercentComplete 0
                Step-P5 'Windows Defender exclusions';  Enable-WindowsDefenderExclusions
                Step-P5 'Power plan';                   Register-ExecutedCommand -Category 'Hardening' -Command 'powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c  # High Performance';   Enable-HighPerformancePowerPlan
                Step-P5 'NIC power management';         Register-ExecutedCommand -Category 'Hardening' -Command 'Get-NetAdapterPowerManagement | Set-NetAdapterPowerManagement -WakeOnMagicPacket Disabled -WakeOnPattern Disabled';  Disable-NICPowerManagement
                Step-P5 'Page file';                    Set-Pagefile
                Step-P5 'TCP settings';                 Register-ExecutedCommand -Category 'Hardening' -Command 'Set-ItemProperty HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters KeepAliveTime 900000';  Set-TCPSettings
                Step-P5 'SMBv1';                        Register-ExecutedCommand -Category 'Hardening' -Command 'Disable-WindowsOptionalFeature -Online -FeatureName SMB1Protocol'; Register-ExecutedCommand -Category 'Hardening' -Command 'Set-SmbServerConfiguration -EnableSMB1Protocol $false';  Disable-SMBv1
                Step-P5 'Windows Search service';       Register-ExecutedCommand -Category 'Hardening' -Command 'Set-Service WSearch -StartupType Disabled'; Register-ExecutedCommand -Category 'Hardening' -Command 'Stop-Service WSearch';  Disable-WindowsSearchService
                Step-P5 'Unnecessary services';         Disable-UnnecessaryServices
                Step-P5 'Shutdown Event Tracker';       Register-ExecutedCommand -Category 'Hardening' -Command "Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Reliability' -Name ShutdownReasonOn -Value 0"; Register-ExecutedCommand -Category 'Hardening' -Command "Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Reliability' -Name ShutdownReasonUI -Value 0";  Disable-ShutdownEventTracker
                Step-P5 'WDigest caching';              Register-ExecutedCommand -Category 'Hardening' -Command "Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest' UseLogonCredential 0";  Disable-WDigestCredentialCaching
                Step-P5 'HTTP/2';                       Register-ExecutedCommand -Category 'Hardening' -Command "Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\HTTP\Parameters' -Name EnableHttp2Tls -Value 0"; Register-ExecutedCommand -Category 'Hardening' -Command "Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\HTTP\Parameters' -Name EnableHttp2Cleartext -Value 0";  Disable-HTTP2
                Step-P5 'TCP offload';                  $osBuildTcp = [int](Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name CurrentBuildNumber -EA SilentlyContinue).CurrentBuildNumber; if ($osBuildTcp -gt 0 -and $osBuildTcp -lt 17763) { Register-ExecutedCommand -Category 'Hardening' -Command 'netsh int tcp set global chimney=disabled autotuninglevel=restricted' } else { Register-ExecutedCommand -Category 'Hardening' -Command 'netsh int tcp set global autotuninglevel=restricted' }; Register-ExecutedCommand -Category 'Hardening' -Command 'Set-NetOffloadGlobalSetting -TaskOffload Disabled';  Disable-TCPOffload
                Step-P5 'Credential Guard';             Register-ExecutedCommand -Category 'Hardening' -Command "Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA' -Name LsaCfgFlags -Value 0"; Register-ExecutedCommand -Category 'Hardening' -Command "Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard' -Name EnableVirtualizationBasedSecurity -Value 0";  Disable-CredentialGuard
                Step-P5 'LM compatibility level';       Register-ExecutedCommand -Category 'Hardening' -Command "Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' LmCompatibilityLevel 5  # NTLMv2 only";  Set-LmCompatibilityLevel
                Step-P5 'LSA Protection (RunAsPPL)';    Register-ExecutedCommand -Category 'Hardening' -Command "Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' RunAsPPL 1  # effective after reboot";  Enable-LSAProtection
                Step-P5 'RSS / NIC queues';             Enable-RSSOnAllNICs  # registers internally (runtime queue count)
                Step-P5 'IPv4 over IPv6';               Register-ExecutedCommand -Category 'Hardening' -Command "Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters' DisabledComponents 0x20  # prefer IPv4, keep IPv6 loopback";  Set-IPv4OverIPv6Preference
                Step-P5 'NetBIOS on NICs';              Register-ExecutedCommand -Category 'Hardening' -Command 'Get-CimInstance Win32_NetworkAdapterConfiguration -Filter IPEnabled=True | Invoke-CimMethod SetTcpipNetbios @{TcpipNetbiosOptions=2}';  Disable-NetBIOSOnAllNICs
                Step-P5 'LLMNR';                        Register-ExecutedCommand -Category 'Hardening' -Command "Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient' EnableMulticast 0";  Disable-LLMNR
                Step-P5 'mDNS';                         Register-ExecutedCommand -Category 'Hardening' -Command "Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\Dnscache\Parameters' EnableMDNS 0";  Disable-MDNS
                Step-P5 'MaxConcurrentAPI';             Set-MaxConcurrentAPI  # registers internally (runtime processor count)
                Step-P5 'Disk allocation unit';         Test-DiskAllocationUnitSize
                Step-P5 'Scheduled tasks';              Disable-UnnecessaryScheduledTasks
                Step-P5 'CRL check timeout';            Register-ExecutedCommand -Category 'Hardening' -Command "Set-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Cryptography\OID\EncodingType 0\CertDllCreateCertificateChainEngine\Config' -Name ChainUrlRetrievalTimeoutMilliseconds -Value 15000";  Set-CRLCheckTimeout
                Step-P5 'TLS / Schannel'
                if ( $State["DisableSSL3"]) {
                    Register-ExecutedCommand -Category 'Hardening' -Command "Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server' -Name Enabled -Value 0"
                    Disable-SSL3
                }
                if ( $State["DisableRC4"]) {
                    Disable-RC4
                }
                Set-TLSSettings -TLS12 $State["EnableTLS12"] -TLS13 $State["EnableTLS13"]

                Step-P5 'Root CA auto-update';     Register-ExecutedCommand -Category 'Hardening' -Command "Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\SystemCertificates\AuthRoot' DisableRootAutoUpdate 0";  Enable-RootCertificateAutoUpdate

                Step-P5 'Exchange module + search tuning'
                Import-ExchangeModule
                Register-ExecutedCommand -Category 'ExchangeTuning' -Command "Set-ItemProperty 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Search\SystemParameters' CtsProcessorAffinityPercentage 0"
                Set-CtsProcessorAffinityPercentage
                Register-ExecutedCommand -Category 'ExchangeTuning' -Command "Set-ItemProperty 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Diagnostics' EnableSerializationDataSigning 1"
                Enable-SerializedDataSigning
                Register-ExecutedCommand -Category 'ExchangeTuning' -Command '(Select-Xml -Path "$ExchangePath\Bin\Search\Ceres\Runtime\1.0\noderunner.exe.config" -XPath "//nodeRunnerSettings").Node.memoryLimitMegabytes = "0"'
                Set-NodeRunnerMemoryLimit
                Register-ExecutedCommand -Category 'ExchangeTuning' -Command '(Select-Xml -Path "$ExchangePath\bin\MSExchangeMapiFrontEndAppPool_CLRConfig.config" -XPath "//gcServer").Node.enabled = "true"  # servers with >= 20 GB RAM'
                Enable-MAPIFrontEndServerGC

                if ( $State["EnableECC"]) {
                    Register-ExecutedCommand -Category 'Hardening' -Command "New-ItemProperty 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Diagnostics' -Name EnableEccCertificateSupport -Value 1 -Type String -Force"
                    Enable-ECC
                }
                if ( $State["EnableCBC"]) {
                    Register-ExecutedCommand -Category 'Hardening' -Command 'New-SettingOverride -Name "EnableEncryptionAlgorithmCBC" -Parameters @("Enabled=True") -Component Encryption -Section EnableEncryptionAlgorithmCBC -Reason "Enable CBC encryption"'
                    Enable-CBC
                }
                if ( $State["EnableAMSI"]) {
                    # HealthChecker always checks for the SettingOverride (Get-SettingOverride on
                    # AmsiRequestBodyScanning), regardless of Exchange version defaults.
                    # Apply the override for all versions so HC reports the correct state.
                    Register-ExecutedCommand -Category 'ExchangeTuning' -Command 'New-SettingOverride -Name "AMSI" -Component "Cafe" -Section "HttpRequestFiltering" -Parameters @("Enabled=True") -Reason "EXpress"'
                    Enable-AMSI
                }

                if ( $State["InstallMailbox"] ) {
                    # Insert your own Mailbox Server code here
                }
                if ( $State["InstallEdge"]) {
                    # Insert your own Edge Server code here
                }
                # Insert your own generic customizations here

                # Org / Transport optimizations (interactive menu or Autopilot defaults)
                if (-not $State['InstallEdge']) {
                    Step-P5 'Org/Transport optimizations'
                    Invoke-ExchangeOptimizations
                }

                # IANA timezone mapping check (Exchange 2019 CU14+ / SE)
                if (-not $State['InstallEdge']) {
                    Step-P5 'IANA timezone mapping'
                    Enable-IanaTimeZoneMappings
                }

                # SSL offloading, Extended Protection, MRS Proxy, MAPI encryption (F13, F6, F18, F19)
                if (-not $State['InstallEdge']) {
                    Step-P5 'SSL offloading'
                    Disable-SSLOffloading
                    Step-P5 'Extended Protection'
                    Enable-ExtendedProtection
                    Step-P5 'MRS Proxy'
                    Register-ExecutedCommand -Category 'ExchangeTuning' -Command "Get-WebServicesVirtualDirectory -Server $env:COMPUTERNAME | Set-WebServicesVirtualDirectory -MRSProxyEnabled `$false"
                    Disable-MRSProxy
                    Step-P5 'MAPI encryption'
                    Register-ExecutedCommand -Category 'ExchangeTuning' -Command "Set-MailboxServer -Identity '$env:COMPUTERNAME' -MAPIEncryptionRequired `$true"
                    Set-MAPIEncryptionRequired
                }

                # Import PFX certificate — must run BEFORE Exchange SU; the SU installer restarts
                # Exchange services (and possibly W3SVC) which kills the EMS session. Certificate
                # import uses Import-ExchangeCertificate / Enable-ExchangeCertificate (EMS cmdlets)
                # so it must complete while the session established above is still alive.
                Step-P5 'PFX certificate import'
                if ($State['CertificatePath']) {
                    Import-ExchangeCertificateFromPFX
                }

                # HSTS header — only when a certificate was imported (avoid browser lockout with self-signed cert)
                # Kept immediately after cert import; no EMS required (WebAdministration module).
                Step-P5 'HSTS header'
                if ($State['CertificatePath']) {
                    Set-HSTSHeader
                }
                else {
                    Write-MyVerbose 'No CertificatePath specified — skipping HSTS (requires valid certificate to avoid browser lockout)'
                }

                # Exchange Security Updates — installer restarts Exchange services (and may restart
                # W3SVC), which kills the EMS session. All EMS-dependent operations above must be
                # complete before this block. No EMS required after this point in Phase 5.
                # Capture W3SVC PID before SU so we can detect whether the installer restarted IIS.
                $w3svcPidBeforeSU = (Get-CimInstance Win32_Service -Filter "Name='W3SVC'" -ErrorAction SilentlyContinue).ProcessId
                Step-P5 'Exchange Security Updates'
                if ( $State["IncludeFixes"]) {
                    Write-MyOutput "Installing additional recommended hotfixes and security updates for Exchange"

                    $ImagePathVersion = Get-DetectedFileVersion ( (Get-CimInstance -Query 'SELECT * FROM win32_service WHERE name="MSExchangeServiceHost"').PathName.Trim('"') )
                    Write-MyVerbose ('Installed Exchange MSExchangeIS version {0}' -f $ImagePathVersion)

                    switch ( $State['ExSetupVersion']) {
                        $EX2019SETUPEXE_CU14 {
                            Install-MyPackage 'KB5049233' 'Security Update For Exchange Server 2019 CU14 SU3 V2' 'Exchange2019-KB5049233-x64-en.exe' 'https://download.microsoft.com/download/8/0/b/80b356e4-f7b1-4e11-9586-d3132a7a2fc3/Exchange2019-KB5049233-x64-en.exe' ('/passive')
                        }
                        $EX2019SETUPEXE_CU13 {
                            Install-MyPackage 'KB5049233' 'Security Update For Exchange Server 2019 CU13 SU7 V2' 'Exchange2019-KB5049233-x64-en.exe' 'https://download.microsoft.com/download/4/e/5/4e5cbbcc-5894-457d-88c4-c0b2ff7f208f/Exchange2019-KB5049233-x64-en.exe' ('/passive')
                        }
                        $EX2016SETUPEXE_CU23 {
                            Install-MyPackage 'KB5049233' 'Security Update For Exchange Server 2016 CU23 SU14 V2' 'Exchange2016-KB5049233-x64-en.exe' 'https://download.microsoft.com/download/0/9/9/0998c26c-8eb6-403a-b97a-ae44c4db5e20/Exchange2016-KB5049233-x64-en.exe' ('/passive')
                        }
                        default {

                        }
                    }

                }

                # Install Exchange Security Update if available and requested
                Install-ExchangeSecurityUpdate

                # Import server configuration from source server (no EMS required)
                Step-P5 'Server configuration import'
                if ($State['CopyServerConfig'] -and $State['ServerConfigExportPath']) {
                    Import-ServerConfig
                }

                # EOMT — optional CVE mitigation tool (no EMS required)
                Step-P5 'EOMT'
                Invoke-EOMT
                if ($p5LastDesc) { Write-MyVerbose ('{0} took {1:F1}s' -f $p5LastDesc, $p5Sw.Elapsed.TotalSeconds) }

                # ECC / CBC / AMSI SettingOverride changes need a W3SVC/WAS restart to take effect.
                # Skip if the SU installer already restarted IIS (detected via W3SVC PID change).
                if ($script:p5NeedsIisRestart) {
                    $w3svcPidAfterSU = (Get-CimInstance Win32_Service -Filter "Name='W3SVC'" -ErrorAction SilentlyContinue).ProcessId
                    $suRestartedIis  = $w3svcPidBeforeSU -and $w3svcPidAfterSU -and ($w3svcPidBeforeSU -ne $w3svcPidAfterSU)
                    $rebootPending   = $State['RebootRequired'] -or (Test-RebootPending)
                    if ($suRestartedIis) {
                        Write-MyVerbose 'W3SVC already restarted by SU installer — skipping IIS restart for ECC/CBC/AMSI'
                    }
                    elseif ($rebootPending) {
                        Write-MyVerbose 'Server reboot pending after Phase 5 — IIS will restart with the reboot, skipping explicit W3SVC restart'
                    }
                    else {
                        Write-MyOutput 'Restarting W3SVC and WAS to activate ECC/CBC/AMSI SettingOverride changes (may take up to ~60s)'
                        Restart-Service -Name W3SVC, WAS -Force -WarningAction SilentlyContinue
                        Write-MyOutput 'W3SVC and WAS restarted'
                    }
                    $script:p5NeedsIisRestart = $false
                }

                Write-PhaseProgress -Id 1 -Activity 'Phase 5 of 6: Post-configuration' -Completed
                Write-PhaseProgress -Activity 'Exchange Installation' -Completed
            }

            6 {
                Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: Finalizing' -PercentComplete 0
                Enable-DefenderRealtimeMonitoring
                if ( Get-Service MSExchangeTransport -ErrorAction SilentlyContinue) {
                    Write-MyOutput "Configuring MSExchangeTransport startup to Automatic"
                    Set-Service MSExchangeTransport -StartupType Automatic
                }
                if ( Get-Service MSExchangeFrontEndTransport -ErrorAction SilentlyContinue) {
                    Write-MyOutput "Configuring MSExchangeFrontEndTransport startup to Automatic"
                    Set-Service MSExchangeFrontEndTransport -StartupType Automatic
                    Start-Service MSExchangeFrontEndTransport -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                }

                $null = Enable-MSExchangeAutodiscoverAppPool

                # W3SVC/WAS was restarted at the end of Phase 5 (ECC/CBC/AMSI SettingOverride),
                # and MSExchangeFrontEndTransport was just started above — both kill any existing
                # Exchange implicit-remoting session. Reconnect before the first EMS cmdlet.
                if (-not $State['InstallEdge']) {
                    Reconnect-ExchangeSession
                }

                # Load Exchange PS module once for all Phase 6 operations
                if (-not $State['InstallEdge']) {
                    Import-ExchangeModule
                }

                # Install antispam agents (Mailbox role only)
                if ($State['InstallMailbox'] -and -not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: Antispam agents' -PercentComplete 8
                    Install-AntispamAgents
                }

                # Set Virtual Directory URLs
                if ($State['Namespace'] -and -not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: Virtual Directory URLs' -PercentComplete 15
                    Set-VirtualDirectoryURLs
                }

                # Join Database Availability Group
                if ($State['DAGName']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: Joining DAG' -PercentComplete 30
                    Join-DAG
                }

                # DAG replication health check (F8)
                if ($State['DAGName'] -and -not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: DAG replication health' -PercentComplete 33
                    Test-DAGReplicationHealth
                }

                # Add server to existing Send Connectors
                if (-not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: Send Connectors' -PercentComplete 35
                    Add-ServerToSendConnectors
                }

                # Server Manager stays disabled permanently on Exchange servers (set machine-wide in Phase 5)

                if ( !($State['InstallEdge'])) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: IIS health check' -PercentComplete 60
                    Write-MyVerbose 'Performing Health Monitor checks..'
                    # Warmup IIS
                    $hcPassed = 0
                    $hcFailed = 0
                    'OWA', 'ECP', 'EWS', 'Autodiscover', 'Microsoft-Server-ActiveSync', 'OAB', 'mapi', 'rpc' | ForEach-Object {
                        $url = 'https://localhost/{0}/healthcheck.htm' -f $_
                        try {
                            if ($PSVersionTable.PSVersion.Major -ge 6) {
                                $response = Invoke-WebRequest -Uri $url -SkipCertificateCheck -UseBasicParsing -ErrorAction Stop
                                $responseContent = $response.Content
                            } else {
                                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                                $prevCb = [Net.ServicePointManager]::ServerCertificateValidationCallback
                                [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
                                $wc = New-Object System.Net.WebClient
                                try { $responseContent = $wc.DownloadString($url) }
                                finally {
                                    $wc.Dispose()
                                    [Net.ServicePointManager]::ServerCertificateValidationCallback = $prevCb
                                }
                            }
                            Write-MyOutput ('Healthcheck {0}: {1}' -f $url, ($responseContent -split '<')[0])
                            $script:hcPassed++
                        }
                        catch {
                            Write-MyWarning ('Healthcheck {0}: {1}' -f $url, 'ERR')
                            $script:hcFailed++
                        }
                    }
                    Write-MyOutput ('Health Monitor summary: {0} passed, {1} failed out of 8 endpoints' -f $hcPassed, $hcFailed)
                    if ($hcFailed -gt 0) {
                        Write-MyWarning ('{0} health endpoint(s) failed - review above warnings' -f $hcFailed)
                    }
                }
                else {
                    Write-MyVerbose 'InstallEdge Mode, skipping IIS health monitor checks'
                }

                # Database / log path separation check
                if (-not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: DB path check' -PercentComplete 73
                    Test-DBLogPathSeparation
                }

                # Auth Certificate health check + MEAC auto-renewal task
                if (-not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: Auth Certificate check' -PercentComplete 75
                    Test-AuthCertificate
                    Register-AuthCertificateRenewal
                }

                # VSS writers, EEMS, Modern Auth checks (F9, F10, F11)
                if (-not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: VSS writers' -PercentComplete 76
                    Test-VSSWriters
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: EEMS status' -PercentComplete 77
                    Test-EEMSStatus
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: Modern Auth' -PercentComplete 77
                    Get-ModernAuthReport
                }

                # Anonymous relay connector
                if (($State['RelaySubnets'] -or $State['ExternalRelaySubnets']) -and -not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: Anonymous relay connector' -PercentComplete 78
                    New-AnonymousRelayConnector
                }

                # Access Namespace — Accepted Domain + Email Address Policy (F26)
                if ($State['AccessNamespaceMail'] -and $State['Namespace'] -and $State['NewExchangeOrg'] -and -not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: Access namespace mail config' -PercentComplete 79
                    Enable-AccessNamespaceMailConfig
                }

                # Exchange log cleanup scheduled task
                Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: Log cleanup task' -PercentComplete 76
                Register-ExchangeLogCleanup

                # RBAC role group membership report
                if (-not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: RBAC report' -PercentComplete 78
                    Get-RBACReport
                }

                # Run CSS-Exchange HealthChecker
                Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: HealthChecker' -PercentComplete 80
                if (-not $State['SkipHealthCheck']) {
                    Invoke-HealthChecker
                }

                # Re-enable UAC and IE ESC BEFORE report/document generation so the
                # captured security state reflects the final, hardened configuration.
                Enable-UAC
                Enable-IEESC

                # Installation Report
                if (-not $State['SkipInstallReport'] -and -not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: Installation Report' -PercentComplete 88
                    # B16: wrap in try/catch so a crash inside New-InstallationReport does not
                    # propagate to the global trap { break } and kill the script before the
                    # "We're good to go" message and phase-end reboot logic run.
                    try { New-InstallationReport } catch {
                        $ierr = $_
                        Write-MyWarning ('Installation Report failed: ' + $ierr.Exception.Message)
                        $sln = if ($ierr.InvocationInfo) { $ierr.InvocationInfo.ScriptLineNumber } else { '?' }
                        $sli = if ($ierr.InvocationInfo) { ($ierr.InvocationInfo.Line -replace '\s+', ' ').Trim() } else { '' }
                        Write-MyWarning ('  at line ' + $sln + ': ' + $sli)
                    }

                    if (-not $State['NoWordDoc']) {
                        Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 6: Word Installation Document' -PercentComplete 94
                        # Ensure Defender realtime is on before we snapshot server state into the report.
                        Enable-DefenderRealtimeMonitoring -Force
                        try { New-InstallationDocument } catch {
                            $derr = $_
                            Write-MyWarning ('Word document failed: ' + $derr.Exception.Message)
                            $dln = if ($derr.InvocationInfo) { $derr.InvocationInfo.ScriptLineNumber } else { '?' }
                            $dli = if ($derr.InvocationInfo) { ($derr.InvocationInfo.Line -replace '\s+', ' ').Trim() } else { '' }
                            Write-MyWarning ('  at line ' + $dln + ': ' + $dli)
                            if ($derr.ScriptStackTrace) { Write-MyVerbose ('  stack: ' + $derr.ScriptStackTrace) }
                        }
                    }
                }

                Write-PhaseProgress -Activity 'Exchange Installation' -Completed
                Write-MyOutput "Setup finished - We're good to go."
            }

            default {
                Write-MyError "Unknown phase ($($State["InstallPhase"]))"
                exit $ERR_UNEXPTECTEDPHASE
            }
        }

            # Skip reboot between Phase 2 and 3 if Windows doesn't signal a pending
            # reboot. Persist the advance immediately so a crash in Phase 3 doesn't
            # re-run Phase 2.
            if ( $State["Autopilot"] -and $State["InstallPhase"] -eq 2 -and -not (Test-RebootPending) ) {
                Write-MyOutput 'Phase 2 complete — no reboot pending, continuing directly with Phase 3 ..'
                $State["LastSuccessfulPhase"] = 2
                $State["InstallPhase"] = 3
                Save-State $State
                $continueInProcess = $true
            }

            # Skip reboot between Phase 5 and 6 unless an Exchange SU set RebootRequired
            # (exit code 3010) or Windows reports a pending reboot from any other source.
            # Phase 5 otherwise only changes registry/IIS settings that don't require reboot.
            if ( $State["Autopilot"] -and $State["InstallPhase"] -eq 5 -and -not $State['RebootRequired'] -and -not (Test-RebootPending) ) {
                Write-MyOutput 'Phase 5 complete — no SU reboot and no pending reboot, continuing directly with Phase 6 ..'
                $State["LastSuccessfulPhase"] = 5
                $State["InstallPhase"] = 6
                Save-State $State
                $continueInProcess = $true
            }
        } while ($continueInProcess)
        } # end else (standard Exchange install switch)
    }
    $State["LastSuccessfulPhase"] = $State["InstallPhase"]
    Enable-OpenFileSecurityWarning
    Save-State $State
    if ( $State['SourceImage']) {
        Dismount-DiskImage -ImagePath $State['SourceImage'] | Out-Null
        Write-MyVerbose ('Dismounted ISO: {0}' -f $State['SourceImage'])
    }

    if ( $State["Autopilot"]) {
        if ( $State["InstallPhase"] -lt $MAX_PHASE) {
            Write-MyVerbose "Preparing system for next phase"
            Disable-UAC
            Enable-AutoLogon
            Enable-RunOnce
        }
        else {
            Cleanup
        }
        Write-MyOutput "Rebooting in $COUNTDOWN_TIMER seconds .."
        for ($i = $COUNTDOWN_TIMER; $i -gt 0; $i--) {
            Write-Progress -Id 2 -Activity 'Reboot' -Status ('Rebooting in {0}s ...' -f $i) -PercentComplete (($COUNTDOWN_TIMER - $i) * 100 / $COUNTDOWN_TIMER)
            Start-Sleep -Seconds 1
        }
        Write-Progress -Id 2 -Activity 'Reboot' -Completed
        Restart-Computer -Force
    }

    exit $ERR_OK

} #Process

