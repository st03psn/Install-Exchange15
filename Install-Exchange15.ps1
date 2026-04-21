<#
    .SYNOPSIS
    Install-Exchange15

    Maintainer: st03ps

    Original author: Michel de Rooij (michel@eightwone.com)
    Many thanks to Michel de Rooij for the extensive prior work this fork
    is built upon. All original copyright and license notices are preserved.

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 5.79, April 21, 2026

    Thanks to Maarten Piederiet, Thomas Stensitzki, Brian Reid, Martin Sieber, Sebastiaan Brozius, Bobby West,`
    Pavel Andreev, Rob Whaley, Simon Poirier, Brenle, Eric Vegter and everyone else who provided feedback
    or contributed in other ways.

    .DESCRIPTION
    This script can install Exchange 2016/2019/SE prerequisites, optionally create the Exchange
    organization (prepares Active Directory) and installs Exchange Server. When the Autopilot switch is
    specified, it will do all the required rebooting and automatic logging on using provided credentials.
    To keep track of provided parameters and state, it uses an XML file; if this file is
    present, this information will be used to resume the process. Note that you can use a central
    location for Install (UNC path with proper permissions) to re-use additional downloads.

    Starting with v5.1, the script can also install Recipient Management Tools (-InstallRecipientManagement)
    and Exchange Management Tools only (-InstallManagementTools) on dedicated admin workstations. When
    started interactively without parameters, an installation menu is shown to configure all options.

    .LINK
    http://eightwone.com

    .NOTES

    Requirements:
    - Supported Operating Systems
      - Windows Server 2016 (Exchange 2016 CU23)
      - Windows Server 2019 (Desktop or Core, Exchange 2019/SE)
      - Windows Server 2022 (Exchange 2019 CU12+/SE)
      - Windows Server 2025 (Exchange 2019 CU15+/SE)
    - Domain-joined system, except for Edge Server Role
    - "Autopilot" mode requires account with elevated administrator privileges
    - When you let the script prepare AD, the account needs proper permissions

    .REVISIONS

    1.0     Initial community release
    1.01    Added logic to prepare AD when organization present
            Fixed checks and logic to prepare AD
            Added testing for domain mixed/native mode
            Added testing for forest functional level
    1.02    Fixed small typo in post-prepare AD function
    1.03    Replaced installing most OS features in favor of /InstallWindowsComponents
            Removed installation of Office Filtering Pack
    1.1     When used for AD preparation, RSAT-ADDS-Tools won't be uninstalled
            Pending reboot detection. In Autopilot, script will reboot and restart phase.
            Installs Server-Media-Foundation feature (UCMA 4.0 requirement)
            Validates provided credentials for Autopilot
            Check OS version as string (should accomodate non-US OS)
    1.5     Added support for WS2008R2 (i.e. added prereqs NET45, WMF3), IEESC toggling,
            KB974405, KB2619234, KB2758857 (supersedes KB2533623). Inserted phase for
            WS2008R2 to install hotfixes (+reboot); this phase is skipped for WS2012.
            Added InstallPath to Autopilot set (or default won't be set).
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
    1.71    Uncommented RunOnce line - Autopilot should work again
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
    4.30    Added post-config security hardening and performance optimizations:
            - Disable SMBv1 protocol (security best practice)
            - Disable Windows Search service (Exchange has own content indexing)
            - Disable WDigest credential caching (Mimikatz mitigation)
            - Disable HTTP/2 protocol (Exchange MAPI/RPC compatibility)
            - Disable TCP Chimney and Task Offload (performance)
            - Verify 64KB disk allocation unit sizes (Exchange best practice)
            - Disable unnecessary scheduled tasks (defrag)
            - Configure CRL check timeout to prevent startup delays
    4.31    Added CSS-Exchange HealthChecker recommendations:
            - Disable Credential Guard (performance issues, default on WS2025)
            - Set LAN Manager compatibility level to 5 (NTLMv2 only)
            - Enable Receive Side Scaling (RSS) on all NICs
            - Set CtsProcessorAffinityPercentage to 0 (Search performance)
            - Enable Serialized Data Signing (security hardening)
            - Set NodeRunner memory limit to 0 (Search performance)
            - Enable Server GC for MAPI Front End App Pool (20+ GB RAM)
            - Extended .NET strong crypto to v2.0 paths (HealthChecker requirement)
            - Fixed $Error[0] in Enable-MSExchangeAutodiscoverAppPool catch blocks
    5.0     Major feature release:
            - Pre-flight validation HTML report (-PreflightOnly)
            - Idempotency guards: Set-RegistryValue skips if value already set
            - Post-install CSS-Exchange HealthChecker auto-run (-SkipHealthCheck to skip)
            - Configuration export/import from source server (-CopyServerConfig <ServerName>)
            - PFX certificate import with IIS+SMTP binding (-CertificatePath)
            - DAG join automation (-DAGName)
            - System Restore checkpoints before each phase (-NoCheckpoint to skip)
            - Added Exchange Server SE RTM support (build 15.02.2562.017)
            - Exchange SE OS compatibility check (requires WS2019+)
            - Exchange SE coexistence warning (EX2016 must be decommissioned before SE CU2)
            - Exchange SE RTM Feb26SU (KB5074992) in IncludeFixes
            - Exchange SE IU registry path detection
    5.01    Bugfixes and robustness improvements:
            - Auto-elevation: script re-launches elevated when not running as Administrator
            - Auto-unblock: detect and remove Zone.Identifier on Exchange setup source files
              (prevents .NET assembly sandboxing errors from downloaded/extracted media)
            - Fixed Initialize-Exchange: $MinFFL/$MinDFL now set for new-org path
              (was unset, causing post-PrepareAD validation to compare against $null)
            - Fixed Initialize-Exchange: setup.exe /PrepareAD exit code is now checked
              (exit code 1 was silently ignored, causing script to advance to next phase)
            - Fixed FFL/DFL null check in Test-Preflight: $null -lt 17000 evaluated to $true
              in PowerShell, causing false abort when AD was not yet prepared
            - Pre-flight report now only generated on first phase (was repeated every phase)
            - System Restore checkpoint: detect if Checkpoint-Computer is available
              (not present on Windows Server, was producing warning every phase)
    5.1     Major feature release (maintainer: st03ps):
            - Interactive installation menu when started without parameters
              (numbered mode selection + letter toggles for switches, with greying of
              options not applicable to the selected mode; RawUI.ReadKey for instant
              toggle, falls back to Read-Host for RDP/PS2Exe/redirected-stdin compat)
            - Credential prompt with validation retry loop (max 3 attempts, interactive only)
              via new Get-ValidatedCredentials function
            - New mode: Recipient Management Tools (-InstallRecipientManagement)
              3-phase install flow for dedicated Exchange Recipient Admin workstations
              (Server/Client aware prerequisites, optional AD permission setup, desktop shortcut)
            - New mode: Exchange Management Tools only (-InstallManagementTools)
              3-phase install flow installing setup.exe /roles:ManagementTools
            - New Build.ps1 helper to compile the script into a single .exe via PS2Exe
            - Script self-detection when running as .exe (PS2Exe): RunOnce command is
              adjusted accordingly so Autopilot mode keeps working
            - Automatic Windows Update + Exchange Security Update handling (-InstallWindowsUpdates)
              via PSWindowsUpdate module with WUA COM fallback; known Exchange SU download list
              ($ExchangeSUMap: SE RTM/2019 CU13-15/2016 CU23, direct download.microsoft.com URLs)
            - Configuration file support (-ConfigFile) to load all parameters from a .psd1
            - Write-Progress indicators (Id 0 = overall phase, Id 1 = Phase 5 step counter)
            - Header/help documentation synchronized with all parameters
            - Enable-LSAProtection: LSA RunAsPPL=1 (Exchange 2019 CU12+/SE; reboot required)
            - Set-MaxConcurrentAPI: Netlogon MaxConcurrentApi = logical core count (min 10)
            - Enable-RSSOnAllNICs: additionally sets NumberOfReceiveQueues to physical cores
            - Clear-DesktopBackground: RUNDLL32-based (no Add-Type/C# compilation delay)
            - Get-DetectedFileVersion: FileVersionInfo API (no Get-Command PATH overhead)
            - Invoke-WebDownload: PS 5.1-compatible download helper (WebClient fallback)
            - IIS health check: PS 5.1/6+ split (WebClient.DownloadString vs Invoke-WebRequest)
            - Fixed $InstallWindowsUpdates not mapped from menu result (toggle R had no effect)
            - Fixed $Error[0] in autodiscover background job catch blocks
            - Fixed Get-WindowsFeature Bits check in Cleanup (missing .Installed)
            - Parameter block: removed ValueFromPipelineByPropertyName=$false (PS default)
            - Fixed Zone.Identifier check on UDF/ISO source paths: skip ADS query when source is a
              mounted ISO (UDF has no ADS support); added try/catch safety net for exotic filesystems
            - Fixed Server Manager and IE ESC reappearing after every reboot: Disable-ServerManagerAtLogon
              and Disable-IEESC moved to Phase 1 (called once; registry changes persist across reboots)
            - Install-PendingWindowsUpdates: per-update prompt (Y/N/A/S) in interactive mode;
              Autopilot installs all without prompting; download+install runs in background job
              with $WU_DOWNLOAD_TIMEOUT_SEC (300s) timeout — Exchange install continues on timeout
            - config.psd1 auto-detection: if found in script/exe folder on interactive start,
              user is asked whether to use it before the installation menu is shown
            - Install-PendingWindowsUpdates: prompt shown whenever console is interactive
              (Autopilot no longer suppresses prompt); timeout raised to 3600s (60 min)
            - Write-PhaseProgress: PS2Exe fallback — emits status via Write-MyOutput when
              Write-Progress is not rendered (Id 0 = phase transitions, Id 1 = Phase 5 steps)
    5.2     v5.2 feature release:
            - Set-HSTSHeader: configures Strict-Transport-Security (max-age=31536000) on OWA
              and ECP virtual directories in IIS; only runs when -CertificatePath is set to
              avoid browser lockout with self-signed certificates (Phase 5)
            - Invoke-EOMT: downloads and runs CSS-Exchange Emergency Mitigation Tool for
              CVE mitigation; new -RunEOMT switch; BITS download with Invoke-WebDownload
              fallback; SHA256 logged (Phase 5)
            - Test-DBLogPathSeparation: warns when database and transaction log paths share
              the same volume root; also emits DAG-aware database size guidance (Phase 6)
            - New-PreflightReport: added Exchange Database Sizing Best Practices section
              (DAG vs standalone max sizes, log/DB separation, allocation unit, free space)
            - Wait-ADReplication: polls repadmin /showrepl /errorsonly after PrepareAD until
              replication is error-free or 6-minute timeout; -WaitForADSync opt-in (Phase 3)
            - Register-ExchangeLogCleanup: creates \Exchange\Exchange Log Cleanup scheduled
              task (daily 02:00, SYSTEM, 1h limit); cleans IIS + transport + tracking logs
              older than -LogRetentionDays days (1–365); script stored in InstallPath (Phase 6)
            - New-AnonymousRelayConnector: two-connector design — "Anonymous Internal Relay"
              (-RelaySubnets, no external relay right) and "Anonymous External Relay"
              (-ExternalRelaySubnets, Ms-Exch-SMTP-Accept-Any-Recipient granted); on success
              removes AnonymousUsers from Default Frontend connector; idempotent (Phase 6)
            - StandaloneOptimize mode (-StandaloneOptimize): single-phase run of all
              post-install tasks on an existing Exchange server; shares -Namespace,
              -CertificatePath, -DAGName, -SkipHealthCheck, -RelaySubnets, -LogRetentionDays
            - Pester tests extended: Get-FullDomainAccount edge cases, DB/Log separation
              logic, HSTS header value validation (no includeSubDomains, min 1-year max-age)
            Bugfixes and code quality (2026-04-17):
            - ValidatePattern for -Organization: removed inline (?# ...) regex comment that
              caused a parse error (ArgumentException: Too many )'s) on script load
            - Install-PendingWindowsUpdates: installed count now filtered to approved KBs only;
              PSWindowsUpdate previously returned already-installed updates as 'Installed'
            - Disable-IEESC and Disable-ServerManagerAtLogon moved from Autopilot reboot block
              to Phase 1 (called once); registry changes persist across reboots — no need to
              repeat before each reboot
            - Removed dead code: DisableSharedCacheServiceProbe (function was defined but
              never called)
            - Added named constants: $ERR_SUS_NOT_APPLICABLE (-2145124329),
              $POWERPLAN_HIGH_PERFORMANCE (High Performance plan GUID)
            - TLS 1.3 menu default now uses [System.Version] comparison against $WS2022_PREFULL
              instead of hardcoded integer 20348
            - Write-Host in MSExchangeAutodiscoverAppPool ScriptBlock replaced with
              Write-MyVerbose (logging pattern consistency)
    5.5     Bugfixes and Phase 6 improvements (2026-04-19):
            Bugfixes:
            - New-InstallationReport: fixed FormatError crash — all multi-arg -f expressions
              inside .Add() method calls are now parenthesized to prevent PowerShell from
              splitting the -f arguments as method arguments
            - Invoke-HealthChecker: output suppressed in console; all lines written to
              transcript as verbose; brief message with report file path shown
            - Install-ExchangeSecurityUpdate / ExchangeSUMap: KB5074992 (Exchange SE RTM
              Feb26SU) corrected to .cab extension; Invoke-Process CAB handler extracts
              and installs the embedded .msp via MSIEXEC /update
            - ISO only remounted for phases 1-3; dismounted at end of Phase 4 (consistent)
            - Test-Preflight: heavy checks (setup path, AD, FFL, roles) now correctly
              skipped for Phase >= 5 (no Exchange session available before Phase 5)
            - Set-VirtualDirectoryURLs: MAPI -InternalAuthenticationMethods wrapped in
              separate try/catch (not all builds support the parameter); OWA now sets
              -LogonFormat UPN
            - Get-RBACReport: format string crash fixed — catch block uses string
              interpolation instead of -f operator to avoid argument ambiguity
            - Import-ExchangeModule: no longer emits WARNING when module already loaded;
              uses Get-ExchangeServer instead of Connect-ExchangeServer to detect session;
              Phase 6 loads module exactly once
            New features:
            - Install-AntispamAgents: runs Install-AntispamAgents.ps1 from Exchange
              bin folder, restarts transport service, then disables all anti-spam agents
              except RecipientFilter Agent; called in Phase 6
            - Add-ServerToSendConnectors: interactive Y/N prompt to add the newly
              installed server to all existing Send Connectors; Phase 6
            - Register-ExchangeLogCleanup: interactive prompt for script folder
              (default C:\#service, 2-min timeout); generated cleanup script logs to
              logs\ subfolder, cleans own logs > 30 days, covers IIS + Exchange
              transport + Monitoring logs
            - New-AnonymousRelayConnector: relay connectors now use -AuthMechanism Tls
              (STARTTLS offered) and -ProtocolLoggingLevel Verbose
            - Import-ExchangeCertificateFromPFX: wildcard detection (CN=* or SAN *.);
              non-wildcard certs additionally enable IMAP and POP services
            Report and installation fixes (2026-04-19 session 2):
            - New-InstallationReport: Serialized Data Signing registry value name corrected
              (EnableSerializationDataSigning, was EnableSigningVerification — never matched)
            - New-InstallationReport: Autodiscover SCP moved from Virtual Directory table
              to Exchange Server info section (CAS property, not a virtual directory)
            - New-InstallationReport: Mailbox Database query tries -Status first, falls back
              without -Status; distinguishes "not found" from query error
            - New-InstallationReport: volumes now exclude DVD-ROM and removable drives;
              DNS server addresses shown per NIC
            - New-InstallationReport: Security and Performance tables gain "Exchange
              Recommendation" column with best-practice values; UAC state row added
            - New-InstallationReport: HealthChecker HTML embedded as iframe in new section 7
            - New-InstallationReport: Edge headless stderr redirected to suppress
              Chrome DevTools renderer noise; --log-level=3 flag added
            - VC++ 2013 URL updated to stable Microsoft CDN URL (KB3138367)
            - VC++ 2012 now also installed for Exchange 2016 CU23 (was Edge-only)
    5.51    Bugfixes (2026-04-19):
            - Get-ValidatedCredentials: PSObject cast fixed — Get-Credential result unwrapped
              via PSObject.BaseObject before assigning to typed $Credentials variable; Read-Host
              fallback added for PS2Exe/compiled-exe environments where Get-Credential returns $null
            - Start-DisableMSExchangeAutodiscoverAppPoolJob: Test-Path 'IIS:\AppPools\...' replaces
              Get-WebAppPoolState — PathNotFound exception not suppressed by -ErrorAction SilentlyContinue
              (same fix already applied to Enable-MSExchangeAutodiscoverAppPool in v5.5)
            - Restart-Service W3SVC/WAS and MSExchangeTransport: -WarningAction SilentlyContinue
              added to suppress repetitive polling warnings
            - Install-AntispamAgents: Install-AntispamAgents.ps1 output redirected (3>$null);
              Enable/Disable-TransportAgent get -WarningAction SilentlyContinue
            - Set-VirtualDirectoryURLs: -Confirm:$false added to all Set-*VirtualDirectory calls
              to suppress "host can't be resolved" confirmation prompt
            - Register-ExchangeLogCleanup: FlushInputBuffer in own try/catch — exception no longer
              aborts entire 2-minute RawUI input loop; falls back to Read-Host correctly
            - VC++ 2012 (v11.0): condition extended to all Exchange versions (was Exchange 2016/Edge
              only); HealthChecker flags v11.0 as required for Exchange 2019/SE
            - ExchangeSUMap KB5074992: Windows Update Catalog CAB URL added — auto-download now
              works without manual file placement
            - Reconnect-ExchangeSession: new helper reconnects Exchange implicit-remoting PS session
              after W3SVC restarts caused by Enable-ECC/CBC/AMSI; waits up to 90s for endpoint;
              called before Invoke-ExchangeOptimizations when any of ECC/CBC/AMSI were enabled
    5.79    Bugfix (2026-04-21):
            - New-InstallationReport (B16): transcript file read with explicit UTF-8 encoding
              while PS 5.1 writes transcripts as UTF-16 LE — removed explicit encoding so .NET
              auto-detects the BOM; log section now wrapped in try/catch so an IOException does
              not propagate to the global trap and kill the script; log output capped to last
              2000 lines with truncation notice to prevent multi-MB HTML files when transcript
              has accumulated over multiple reboots; Phase 6 call site wrapped in try/catch so
              any crash inside New-InstallationReport no longer kills the entire script
    5.78    Bugfix (2026-04-21):
            - Install-ExchangeSecurityUpdate (B15): Exchange SU installer may call ExitWindowsEx
              internally and reboot the machine before the script's phase-end logic runs
              (LastSuccessfulPhase update + Enable-RunOnce). In Autopilot mode, RunOnce + state
              are now persisted immediately before launching the installer so the script always
              auto-resumes after an installer-triggered reboot. A per-KB flag
              (ExchangeSUInstalled_<KB>) is stored in state after a successful install (rc 0/3010)
              so phase-5 re-entry skips the SU even when Get-InstalledExchangeBuild still returns
              the pre-SU build number (service binary cache not yet flushed after reboot).
    5.77    Bugfix (2026-04-21):
            - Install-ExchangeSecurityUpdate (B14): removed /norestart from Exchange SU installer
              arguments — Exchange SU .exe only accepts /passive and /silent; /norestart caused
              "The following command line options are not recognized: /norestart" and aborted install
    5.76    Bugfixes (2026-04-21):
            - Test-AuthCertificate: added null-guard for $authConfig before property access;
              Get-AuthConfig can return $null when Exchange PS session is not fully initialized
              (observed in Phase 6 after IIS restart) — previously threw "You cannot call a
              method on a null-valued expression" as a terminating error caught by the catch block
            - New-AnonymousRelayConnector: fixed race condition where Get-ReceiveConnector failed
              immediately after New-ReceiveConnector because Exchange AD had not yet registered
              the new object; now captures the object returned by New-ReceiveConnector directly
              and uses it for Add-ADPermission; added 3-attempt retry fallback with 5s backoff
              for the edge case where New-ReceiveConnector returns null
    5.64    Report cleanup (2026-04-20):
            - New-InstallationReport: HealthChecker section (Section 7) removed from report;
              HC runs independently and produces its own HTML report in ReportsPath
    5.63    P6 — Dynamic SU detection (2026-04-20):
            - Get-LatestSUBuildFromHC: new function parses HC.ps1's GetExchangeBuildDictionary
              to find the latest known SU build for the installed Exchange CU (RTM/CU13–CU15/CU23)
            - Get-InstalledExchangeBuild: new helper reads installed build from MSExchangeServiceHost
              binary (actual running version, not setup media version)
            - Install-ExchangeSecurityUpdate: skips install when installed build already >= SU target;
              downloads HC.ps1 if not present; compares installed build with HC latest and warns if
              behind (newer SU may require ESU enrollment for Exchange 2016/2019)
            - Note: Feb 2026 SU for Exchange 2019 CU14/CU15 and Exchange 2016 CU23 require ESU
              enrollment (no public download URL); Exchange SE Feb 2026 SU is publicly available
    5.62    F14/F15 + VC++ fix (2026-04-20):
            - F14 (OWA Download Domains): -DownloadDomain parameter added; Set-VirtualDirectoryURLs
              sets ExternalDownloadHostName + InternalDownloadHostName on OWA VDir (CVE-2021-1730
              mitigation); prompted in interactive menu after Namespace
            - F15 (AMSI body scanning): Enable-AMSI skipped for Exchange SE — AMSI is enabled by
              default in Exchange SE (Aug 2025 SU era); SettingOverride only applied for Exchange
              2016/2019 where manual configuration is still required
            - VC++ 2013: updated from KB3138367 (12.0.40660) to KB4538461 (12.0.40664); Get-VCRuntime
              now accepts -MinBuild parameter and reinstalls if installed version is below minimum
            - Phase 6: MSExchangeFrontEndTransport now started (not just configured to Automatic)
              after startup type is set; service was left stopped after Exchange setup
    5.66    IPv4 preference and NetBIOS hardening (2026-04-20):
            - Set-IPv4OverIPv6Preference: sets DisabledComponents = 0x20 in
              HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters — prefers IPv4 over
              IPv6 on all non-loopback interfaces while keeping IPv6 loopback intact (required
              by Exchange internal components); Microsoft recommended setting for Exchange;
              flags RebootRequired when changed
            - Disable-NetBIOSOnAllNICs: calls SetTcpipNetbios(2) via WMI on all IP-enabled
              NICs; Exchange does not require NetBIOS; disabling reduces attack surface for
              LLMNR / NBT-NS poisoning and credential capture attacks
    5.65    Fix broken SU download hint URL (2026-04-20):
            - Replace aka.ms/ExchangeSU (resolves to Bing, not Exchange content) with the
              official Microsoft Learn build numbers page which lists all SUs with download links:
              https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates
    5.64    Register-ExchangeLogCleanup log path improvements (2026-04-20):
            - Invoke-ExchangeLogCleanup.ps1: Exchange log coverage expanded from 6 specific
              transport sub-paths to two recursive trees: V15\Logging\ (EWS, OWA, ECP,
              HttpProxy, RpcClientAccess, MAPI, Search, Audit, Migration, etc.) and
              V15\TransportRoles\Logs\ (Hub, FrontEnd, MessageTracking, DSN, Connectivity,
              Pickup, Replay, etc.) — prevents large log accumulation in non-transport areas
            - HTTPERR logs added: %SystemRoot%\System32\LogFiles\HTTPERR cleaned with same
              retention period
            - IIS log path now dynamically resolved from IIS metabase via
              Get-WebConfigurationProperty (handles non-default IIS log locations);
              falls back to inetpub\logs\LogFiles when WebAdministration module unavailable
            - All Get-ChildItem results wrapped in @() — prevents .Count failure on $null
              in PS 5.1 when no files match the filter
    5.63    Bugfixes (2026-04-20):
            - Install-AntispamAgents: changed & $installScript 3>$null to *>&1 | Out-Null —
              3>$null did not suppress implicit-remoting warnings from Enable-TransportAgent
              (PS 5.1 implicit-remoting warnings bypass stream 3 redirection and write directly
              to host); *>&1 | Out-Null suppresses all output including the agent table and the
              duplicate "service restart required / Please exit PowerShell" warnings
            - Exchange SE RTM SU (KB5074992): removed hardcoded Install-MyPackage switch entry —
              it was a duplicate of Install-ExchangeSecurityUpdate / $ExchangeSUMap and caused
              two consecutive failed install attempts; $ExchangeSUMap URL cleared (WU-catalog
              CAB format is not DISM-compatible); FileName changed to .exe so a manually placed
              EXE installer is picked up automatically; improved warning message with download hint
    5.62    Security hardening and Phase 6 health checks (2026-04-20):
            - F13 (Disable SSL Offloading): Disable-SSLOffloading sets Set-OutlookAnywhere
              -SSLOffloading $false — prerequisite for Extended Protection channel binding
            - F6 (Extended Protection): Enable-ExtendedProtection validates EP on CU14+/SE;
              downloads and runs ExchangeExtendedProtection.ps1 (CSS-Exchange) for 2016/pre-CU14
            - F17 (Root CA AutoUpdate): Enable-RootCertificateAutoUpdate re-enables root cert
              auto-update when disabled by policy (prevents Exchange Online / M365 connectivity break)
            - F18 (Disable MRS Proxy): Disable-MRSProxy sets MRSProxyEnabled $false on EWS VDir
              (re-enable manually when performing cross-forest migrations)
            - F19 (MAPI Encryption Required): Set-MAPIEncryptionRequired forces MAPI encryption
              on all Outlook connections (Set-MailboxServer -MAPIEncryptionRequired $true)
            - F8 (DAG Replication Health): Test-DAGReplicationHealth checks database copy status
              after DAG join in Phase 6; warns on non-Mounted/Healthy copies
            - F9 (VSS Writer Validation): Test-VSSWriters calls vssadmin list writers and warns
              on any writer not in a Stable state (broken VSS → failed Exchange backups)
            - F10 (EEMS Status Check): Test-EEMSStatus checks MSExchangeMitigation service and
              org-wide MitigationsEnabled / MitigationsBlocked settings (CU11+ / SE)
            - F11 (Modern Auth Report): Get-ModernAuthReport warns when OAuth2 is disabled
            - P7 (Compliance Mapping): Installation Report Security section gains CIS / BSI
              control-ID column: TLS entries reference CIS L1 / PCI-DSS 4.2.1; SMBv1 CIS §18.3;
              WDigest CIS §18.9.48 / DISA STIG; LSA CIS L2; LM compat CIS §2.3.11.7; UAC CIS L1
    5.61    Bugfixes (2026-04-20):
            - Set-VirtualDirectoryURLs: @forceParam removed from all six Set-*VirtualDirectory
              cmdlets; OWA -Force was ambiguous (matched ForceSave*/ForceWac* parameters) and
              threw ParameterBindingException; all cmdlets now use -Confirm:$false only
            - New-AnonymousRelayConnector: external relay placeholder warning corrected —
              192.0.2.1/32 → 192.0.2.2/32 (internal and external use different RFC 5737 addresses)
            - New-AnonymousRelayConnector: Add-ADPermission -WarningAction SilentlyContinue added;
              "access control entry already present" warning suppressed on re-run
            - Register-ExchangeLogCleanup: RawUI unavailable catch block simplified — default
              folder now accepted silently instead of blocking Read-Host fallback
            - Countdown progress bars (Write-Progress -Id 2) added to all timed prompts:
              log cleanup folder (2 min auto-accept), Windows Update per-update (2 min auto-No),
              Autopilot resume (10 s), reboot countdown (10 s)
            - New-InstallationReport: '-f (if ...)' expression caused CommandNotFoundException
              for 'if' in PS 5.1; replaced with intermediate variable
            - Invoke-HealthChecker / New-InstallationReport: HC report detection updated to
              match current filename (ExchangeAllServersReport-*.html, was HealthChecker*.html);
              detected path stored in $State['HCReportPath'] for reliable report embedding
    5.6     Report improvements and relay connector enhancements (2026-04-19):
            - New-InstallationReport: RBAC Role Group Membership section added (10 groups, queried
              live via Get-RoleGroupMember; members shown with RecipientType)
            - New-InstallationReport: Installation Log section added (transcript content embedded
              as scrollable dark <pre> block; HTML-escaped)
            - New-InstallationReport: Autodiscover SCP moved into Virtual Directory URLs table
              (first row; queried via Get-ClientAccessService)
            - New-InstallationReport: UAC now re-enabled before report generation (was after);
              report correctly shows UAC = Enabled
            - New-InstallationReport: HealthChecker section distinguishes -SkipHealthCheck
              (intentionally skipped) from report not found (HC failed)
            - Reports subfolder: all reports and logs written to <InstallPath>\reports\
              ($State['ReportsPath']); folder created on first run and on Autopilot resume
            - New-AnonymousRelayConnector: menu now creates both internal AND external relay
              connectors when [Y] selected; blank subnet entry uses RFC 5737 placeholder
              192.0.2.1/32 (never routable, no SMTP traffic matches until updated); Default
              Frontend AnonymousUsers hardening skipped when only placeholders are set
            - Register-ExchangeLogCleanup: interactive prompt skipped in Autopilot mode (uses
              default C:\#service silently)
    5.4     Installation Report + Verbose Logging (2026-04-18):
            - New-InstallationReport: comprehensive HTML report at Phase 6 completion;
              6 sections: Installation Parameters, System Info, Active Directory,
              Exchange Configuration (vdirs, DBs, connectors, certs), Security Settings,
              Performance & Tuning; modern sidebar navigation, status badges, print CSS
            - PDF export via Microsoft Edge headless (--print-to-pdf); automatic when
              msedge.exe is found; fallback message to browser print if not available
            - -SkipInstallReport switch to suppress report generation
            - Verbose messages always written to log file (transcript); console output
              suppressed ($VerbosePreference = SilentlyContinue)
    5.3     Code quality and robustness (2026-04-17):
            - Add-BackgroundJob: new helper that prunes Completed/Failed/Stopped entries
              from $Global:BackgroundJobs before appending; prevents unbounded list growth
            - New-LDAPSearch: new helper that encapsulates DirectorySearcher creation
              (SearchRoot + Filter); eliminates duplicated 3-line blocks in 4 functions
            - Registry idempotency: RunOnce, Disable/Enable-UAC, Enable-AutoLogon,
              Disable-OpenFileSecurityWarning, Set-NETFrameworkInstallBlock and
              Disable-ServerManagerAtLogon now all use Set-RegistryValue (idempotency guard)
            - BSTR zeroing: ZeroFreeBSTR called after PtrToStringAuto in Test-Credentials
              and Enable-AutoLogon; eliminates plaintext password residue in memory
            - Exit code checks: RUNDLL32 (Clear-DesktopBackground) and powercfg.exe
              (Enable-HighPerformancePowerPlan) now log Write-MyWarning on non-zero exit
            - Pester tests extended: Set-RegistryValue idempotency (5 cases),
              Add-BackgroundJob pruning (4 cases); 45 -> 54 tests total

    .PARAMETER Organization
    Specifies name of the Exchange organization to create. When omitted, the step
    to prepare Active Directory (PrepareAD) will be skipped.

    .PARAMETER InstallEdge
    Specifies you want to install the Edge server role  (Exchange 2016/2019/SE).

    .PARAMETER EdgeDNSSuffix
    Specifies the DNS suffix you want to use on your EDGE

    .PARAMETER MDBName (optional)
    Specifies name of the initially created database.

    .PARAMETER MDBDBPath (optional)
    Specifies database path of the initially created database. Requires MDBName.

    .PARAMETER MDBLogPath (optional)
    Specifies log path of the initially created database. Requires MDBName.

    .PARAMETER InstallPath (optional)
    Specifies (temporary) location of where to store prerequisites files, log
    files, etc. Default location is C:\Install.

    .PARAMETER NoSetup (optional)
    Specifies you don't want to setup Exchange (prepare/prerequisites only). Note that you
    still need to specify the location of Exchange setup, which is used to determine
    its version and which prerequisites should be installed.

    .PARAMETER SourcePath
    Specifies location of the Exchange installation files (setup.exe) or the location of
    the Exchange installation ISO. This ISO will be mounted during installation.

    .PARAMETER TargetPath
    Specifies the location where to install the Exchange binaries.

    .PARAMETER Autopilot (switch)
    Specifies you want to automatically restart and logon using Account specified. When
    not specified, you will need to restart, logon and start the script again manually.
    You also need to use the InstallPath parameter when used before, so the script knows where
    to pick up the state file.

    .PARAMETER Credentials
    Specifies credentials to use for automatic logon. Use DOMAIN\User or user@domain. When
    not specified, you will be prompted to enter credentials.

    .PARAMETER IncludeFixes (optional)
    Depending on operating system and detected Exchange version to install, will download
    and install additional recommended Exchange hotfixes.

    .PARAMETER SkipRolesCheck (optional)
    Instructs script not to check for Schema Admin and Enterprise Admin roles.

    .PARAMETER NONET481 (optional)
    Prevents installing .NET Framework 4.8.1 and uses 4.8 when deploying Exchange 2019 CU14+
    on supported Operating Systems (WS2016, WS2019). WS2022 only supports .NET Framework 4.8.1

    .PARAMETER DoNotEnableEP (optional)
    Do not enable Extended Protection on Exchange 2019 CU14+

    .PARAMETER DoNotEnableEP_FEEWS (optional)
    Do not enable Extended Protection on the Front-End EWS virtual directory on Exchange 2019 CU14+

    .PARAMETER DisableSSL3 (optional)
    Disables SSL3 after setup.

    .PARAMETER DisableRC4 (optional)
    Disables RC4 after setup.

    .PARAMETER EnableECC (optional)
    Configures Elliptic Curve Cryptography support after setup.

    .PARAMETER NoCBC (optional)
    Prevents configuring AES256-CBC-encrypted content support after setup.

    .PARAMETER EnableAMSI (optional)
    Configure AMSI body scanning for ECP, EWS, OWA and PowerShell (adjust as necessary in-code)

    .PARAMETER EnableTLS12 (optional)
    Enable or disable TLS12

    .PARAMETER EnableTLS13 (optional)
    Enable or disable TLS13 on WS2022/WS2025 for Exchange 2019 CU15+ (default: enable)

    .PARAMETER Recover
    Runs Exchange setup in RecoverServer mode.

    .PARAMETER SCP (optional)
    Reconfigures Autodiscover Service Connection Point record for this server post-setup, i.e.
    https://autodiscover.contoso.com/autodiscover/autodiscover.xml. If you want to remove
    the record, set it to '-'.

    .PARAMETER Lock (optional)
    Locks system when running script.

    .PARAMETER DiagnosticData (optional)
    Switch determines initial Data Collection mode for deploying Exchange 2019 CU11+ or Exchange 2016.

    .PARAMETER Phase
    Internal Use Only :)

    .PARAMETER PreflightOnly (optional)
    Runs only the preflight validation checks and generates the HTML report, then exits
    without performing any installation actions.

    .PARAMETER CopyServerConfig (optional)
    Specifies the name of a source Exchange server from which to export configuration
    (Virtual Directories, Transport, Receive Connectors) via Remote PowerShell. The
    exported configuration is applied post-setup.

    .PARAMETER CertificatePath (optional)
    Path to a PFX certificate file that should be imported and enabled for IIS + SMTP
    post-setup. You will be prompted for the PFX password.

    .PARAMETER DAGName (optional)
    Name of an existing Database Availability Group this server should join post-setup.

    .PARAMETER SkipHealthCheck (optional)
    Skips the automatic download and execution of the CSS-Exchange HealthChecker
    at the end of the installation.

    .PARAMETER NoCheckpoint (optional)
    Skips creation of System Restore checkpoints before each phase. Has no effect on
    Windows Server, where Checkpoint-Computer is not available.

    .PARAMETER InstallRecipientManagement (optional, v5.1)
    Activates the Recipient Management Tools installation mode (3-phase flow). Installs
    Exchange setup.exe /roles:ManagementTools on a dedicated admin workstation (Server
    or Client), runs Add-PermissionForEMT.ps1 and creates a desktop shortcut loading
    the *RecipientManagement PSSnapin.

    .PARAMETER InstallManagementTools (optional, v5.1)
    Activates the Exchange Management Tools installation mode (3-phase flow). Installs
    prerequisites and setup.exe /roles:ManagementTools only.

    .PARAMETER RecipientMgmtCleanup (optional, v5.1)
    In Recipient Management mode, performs optional Active Directory cleanup of legacy
    permissions after a successful upgrade install.

    .PARAMETER ConfigFile (optional, v5.1)
    Path to a PowerShell data file (.psd1) containing a hashtable with all parameters
    to use. Makes long command lines manageable for repeat deployments.

    .PARAMETER InstallWindowsUpdates (optional, v5.1)
    Checks for pending Windows Updates and applicable Exchange Security Updates (SUs)
    during phase 1 / post-setup, downloads and installs them. Reboots are integrated
    into the existing Autopilot phase flow.

    .PARAMETER SkipWindowsUpdates (optional, v5.1)
    Explicitly skips the Windows Update / Exchange SU check even when the menu or
    ConfigFile would otherwise enable it.

    .PARAMETER SkipSetupAssist (optional, v5.2)
    Skips the automatic download and execution of CSS-Exchange SetupAssist.ps1
    when Exchange Setup fails in Phase 4.

    .PARAMETER Namespace (optional, v5.2)
    External namespace (e.g. outlook.contoso.com) used to configure all Exchange
    Virtual Directory internal and external URLs in Phase 6. If omitted, Virtual
    Directory URLs are left at their defaults.

    .PARAMETER DownloadDomain (optional, v5.61)
    Separate FQDN for OWA attachment downloads (e.g. download.contoso.com). Configures
    ExternalDownloadHostName and InternalDownloadHostName on the OWA virtual directory to
    mitigate CVE-2021-1730 (attachment cookie theft). Must be a different hostname from
    -Namespace. Requires DNS record and certificate coverage for the specified domain.
    Requires -Namespace to be set.

    .PARAMETER RunEOMT (optional, v5.2)
    Downloads and runs the CSS-Exchange Emergency Mitigation Tool (EOMT) in Phase 5
    to apply Microsoft-recommended CVE mitigations. Use this when deploying a server
    that may have been exposed to publicly known vulnerabilities before patching.

    .PARAMETER WaitForADSync (optional, v5.2)
    After PrepareAD (Phase 3), polls repadmin /showrepl /errorsonly until all AD
    replication errors are resolved or a 6-minute timeout elapses. Useful in
    multi-site AD environments where schema changes need time to replicate.

    .PARAMETER LogRetentionDays (optional, v5.2)
    Registers a Windows Scheduled Task (Exchange Log Cleanup, daily at 02:00)
    that removes IIS log files and Exchange transport/tracking logs older than
    the specified number of days (1–365). Stored in \Exchange\ task folder.

    .PARAMETER RelaySubnets (optional, v5.2)
    IP ranges (e.g. '192.168.1.0/24','10.0.0.5') for anonymous SMTP relay to
    accepted domains only (internal relay). Creates "Anonymous Internal Relay"
    receive connector without Ms-Exch-SMTP-Accept-Any-Recipient. If successful,
    AnonymousUsers is removed from the Default Frontend receive connector.

    .PARAMETER ExternalRelaySubnets (optional, v5.2)
    IP ranges for anonymous SMTP relay to any recipient including external addresses.
    Creates "Anonymous External Relay" receive connector and grants
    Ms-Exch-SMTP-Accept-Any-Recipient to the ANONYMOUS LOGON account (SID S-1-5-7,
    resolved dynamically — language-independent for DE/EN and other OS languages).
    Use with extreme care — only trusted send systems (e.g. scanner/printer IPs).

    .PARAMETER SkipInstallReport (optional, v5.4)
    Suppresses generation of the HTML installation report at Phase 6 completion.
    By default a comprehensive report (and PDF if Microsoft Edge is available)
    is created in InstallPath for customer handover and audit purposes.

    .PARAMETER StandaloneOptimize (optional, v5.2)
    Runs all post-install optimizations (VDir URLs, Exchange optimizations,
    RBAC report, HealthChecker, DB path check, log cleanup task, anonymous relay)
    on an already-installed Exchange server without running the full install flow.
    Combine with -Namespace, -CertificatePath, -DAGName, -RelaySubnets,
    -LogRetentionDays, and -SkipHealthCheck as needed.

    .EXAMPLE
    # Start interactively — opens the installation menu (mode, toggles, inputs)
    .\Install-Exchange15.ps1

    .EXAMPLE
    # Load all parameters from a config file (skips the interactive menu)
    .\Install-Exchange15.ps1 -ConfigFile .\deploy-mbx01.psd1

    .EXAMPLE
    # Fully unattended mailbox install with Autopilot (automatic reboots through all phases)
    .\Install-Exchange15.ps1 -SourcePath D:\Exchange -Organization Contoso -Autopilot

    .EXAMPLE
    # Full install with custom DB paths, Autodiscover SCP, and certificate
    $Cred = Get-Credential
    .\Install-Exchange15.ps1 -SourcePath C:\Install\ExchangeServerSE-x64.iso -Organization Contoso `
        -MDBName MDB01 -MDBDBPath D:\MailboxData\MDB01\DB -MDBLogPath D:\MailboxData\MDB01\Log `
        -SCP https://autodiscover.contoso.com/autodiscover/autodiscover.xml `
        -CertificatePath C:\Certs\mail.pfx -Autopilot -Credentials $Cred

    .EXAMPLE
    # Swing migration: copy config from source server, import PFX, join DAG
    .\Install-Exchange15.ps1 -SourcePath D:\Exchange -Autopilot `
        -CopyServerConfig EX01 -CertificatePath D:\Certs\mail.pfx -DAGName DAG01

    .EXAMPLE
    # Generate pre-flight HTML report only (no installation)
    .\Install-Exchange15.ps1 -SourcePath D:\Exchange -PreflightOnly

    .EXAMPLE
    # Install prerequisites only, skip Exchange setup
    .\Install-Exchange15.ps1 -NoSetup -SourcePath D:\Exchange

    .EXAMPLE
    # Recover a failed server
    .\Install-Exchange15.ps1 -Recover -SourcePath D:\Exchange -Autopilot

    .EXAMPLE
    # Edge Transport role
    .\Install-Exchange15.ps1 -InstallEdge -SourcePath D:\Exchange -Autopilot

    .EXAMPLE
    # Install Recipient Management Tools on an admin workstation
    .\Install-Exchange15.ps1 -InstallRecipientManagement -SourcePath D:\Exchange -Autopilot

    .EXAMPLE
    # Install Exchange Management Tools only (Server OS)
    .\Install-Exchange15.ps1 -InstallManagementTools -SourcePath D:\Exchange

    .EXAMPLE
    # Run all post-install optimizations on an existing Exchange server (no setup required)
    .\Install-Exchange15.ps1 -StandaloneOptimize -Namespace mail.contoso.com `
        -CertificatePath C:\Certs\mail.pfx -LogRetentionDays 30 `
        -RelaySubnets '10.0.1.0/24' -ExternalRelaySubnets '10.0.2.5'

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
    [ValidateRange(1, 365)]
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
    [Switch]$SkipInstallReport
)

process {

    $ScriptVersion = '5.79'

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
        if (-not $State['TranscriptFile']) { return }
        $Location = Split-Path $State['TranscriptFile'] -Parent
        if ( Test-Path $Location) {
            "$(Get-Date -Format u): [$Level] $Text" | Out-File $State['TranscriptFile'] -Append -ErrorAction SilentlyContinue
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
                if (-not $downloaded) { $res = $false }
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
        if ($isExe) {
            $RunOnce = "`"$ScriptFullName`" -InstallPath `"$InstallPath`""
        }
        else {
            $PSExe = (Get-Process -Id $PID).Path
            $RunOnce = "`"$PSExe`" -NoProfile -ExecutionPolicy Unrestricted -Command `"& `'$ScriptFullName`' -InstallPath `'$InstallPath`'`""
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
        $maxAttempts = 3
        for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
            try {
                $defaultUser = if ($State['AdminAccount']) { $State['AdminAccount'] } else { [System.Security.Principal.WindowsIdentity]::GetCurrent().Name }
                $rawCred = Get-Credential -UserName $defaultUser -Message ('Enter credentials for Autopilot (attempt {0}/{1})' -f $attempt, $maxAttempts)
                # Get-Credential can return a PSObject wrapper in some terminal environments; unwrap before assigning to the typed $Credentials parameter variable.
                $Script:Credentials = if ($rawCred -is [pscredential]) { $rawCred }
                                      elseif ($rawCred -and $rawCred.PSObject.BaseObject -is [pscredential]) { $rawCred.PSObject.BaseObject }
                                      else { $null }
                # Fallback: Get-Credential returns $null silently in PS2Exe/compiled-exe and some RDP/Hyper-V console sessions.
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

    function Install-Exchange15_ {
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
            $Feats = 'ADLDS'
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
        $RunFrom = $State['InstallPath']
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

            if ( [System.Version]$MajorOSVersion -ge [System.Version]$WS2016_MAJOR ) {
                Write-MyOutput "Operating System is $($MajorOSVersion).$($MinorOSVersion)"
            }
            else {
                Write-MyError 'The following Operating Systems are supported: Windows Server 2019, Windows Server 2022 (Exchange 2019) or Windows Server 2025 (Exchange 2019 CU15+)'
                exit $ERR_UNEXPECTEDOS
            }
            Write-MyOutput ('Server core mode: {0}' -f (Test-ServerCore))

            $NetVersion = Get-NETVersion
            $NetVersionText = Get-NetVersionText $NetVersion
            Write-MyOutput ".NET Framework is $NetVersion ($NetVersionText)"
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

        if ( ($MajorSetupVersion -eq $EX2019_MAJOR -and [System.Version]$SetupVersion -lt [System.Version]$EX2019SETUPEXE_CU10) -or
            ($MajorSetupVersion -eq $EX2016_MAJOR -and [System.Version]$SetupVersion -lt [System.Version]$EX2016SETUPEXE_CU23) ) {
            Write-MyError 'Unsupported version of Exchange detected; only Exchange SE, Exchange 2019 CU10 or later, or Exchange 2016 CU23 are supported'
            exit $ERR_UNSUPPORTEDEX
        }

        if ( [System.Version]$SetupVersion -ge [System.Version]$EX2019SETUPEXE_CU15) {
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

        if ( [System.Version]$FullOSVersion -ge $WS2025_PREFULL -and [System.Version]$SetupVersion -lt $EX2019SETUPEXE_CU15) {
            Write-MyError 'Windows Server 2025 is only supported for Exchange 2019 CU15 or later, or Exchange Server SE'
            exit $ERR_UNEXPECTEDOS
        }

        if ( [System.Version]$SetupVersion -ge [System.Version]$EXSESETUPEXE_RTM -and [System.Version]$FullOSVersion -lt $WS2019_PREFULL) {
            Write-MyError 'Exchange Server SE requires Windows Server 2019, Windows Server 2022 or Windows Server 2025'
            exit $ERR_UNEXPECTEDOS
        }

        if ( [System.Version]$FullOSVersion -lt [System.Version]$WS2016_MAJOR -and $MajorSetupVersion -eq $EX2016_MAJOR) {
            Write-MyError 'Exchange 2016 requires Windows Server 2016 or later'
            exit $ERR_UNEXPECTEDOS
        }

        if ( [System.Version]$FullOSVersion -ge $WS2022_PREFULL -and [System.Version]$FullOSVersion -lt $WS2025_PREFULL -and [System.Version]$SetupVersion -lt $EX2019SETUPEXE_CU12) {
            Write-MyError 'Windows Server 2022 is only supported for Exchange Server 2019 CU12 or later, or Exchange Server SE'
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

        if ( $State["MajorSetupVersion"] -eq $EX2019_MAJOR -and [System.Version]$State["SetupVersion"] -lt [System.Version]$EX2019SETUPEXE_CU14 ) {
            if ( $State['DoNotEnableEP']) {
                Write-MyWarning 'DoNotEnableEP is not supported with this Exchange version, ignoring this switch'
                $State['DoNotEnableEP'] = $false
            }
            if ( $State['DoNotEnableEP_FEEWS']) {
                Write-MyWarning 'DoNotEnableEP_FEEWS is not supported with this Exchange version, ignoring this switch'
                $State['DoNotEnableEP_FEEWS'] = $false
            }
        }

        if ( ($State["MajorSetupVersion"] -eq $EX2019_MAJOR) -and [System.Version]$State["SetupVersion"] -ge [System.Version]$EX2019SETUPEXE_CU11 ) {
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
        $reportPath = Join-Path $State['ReportsPath'] ('PreflightReport_{0}.html' -f $env:COMPUTERNAME)
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
        $configPath = Join-Path $State['InstallPath'] ('ServerConfig_{0}.xml' -f $SourceServer)

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

        # Ask for script destination folder with 2-minute timeout via RawUI (same pattern as Show-InstallationMenu)
        $defaultScriptFolder = 'C:\#service'
        $scriptFolder = $defaultScriptFolder
        if ([Environment]::UserInteractive) {
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
# Exchange Log Cleanup Script — generated by Install-Exchange15.ps1
# Runs daily via Scheduled Task; retention: $days days for Exchange/IIS logs, 30 days for own logs

param([int]`$DaysToKeep = $days)

`$ScriptDir  = Split-Path -Path `$MyInvocation.MyCommand.Path
`$LogFolder  = Join-Path `$ScriptDir 'logs'
`$LogFile    = Join-Path `$LogFolder ('ExchangeLogCleanup-{0}.log' -f (Get-Date -Format 'yyyyMM'))
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
            Register-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Action $action `
                -Trigger $trigger -Settings $settings -Principal $principal -ErrorAction Stop | Out-Null
            Write-MyOutput ('Scheduled task "{0}" registered — runs daily at 02:00, retention {1} days' -f $taskName, $days)
        }
        catch {
            Write-MyWarning ('Failed to register log cleanup task: {0}' -f $_.Exception.Message)
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
            try {
                # $PSDefaultParameterValues['*:WarningAction'] has higher precedence than $WarningPreference
                # and overrides any internal $WarningPreference reset inside Install-AntispamAgents.ps1
                $savedWP  = $WarningPreference;  $WarningPreference = 'Ignore'
                $waKey    = '*:WarningAction'
                $savedWA  = $PSDefaultParameterValues[$waKey]
                $PSDefaultParameterValues[$waKey] = 'Ignore'
                & $installScript *>&1 | Out-Null
                $PSDefaultParameterValues[$waKey] = $savedWA
                if ($null -eq $savedWA) { $null = $PSDefaultParameterValues.Remove($waKey) }
                $WarningPreference = $savedWP
                Write-MyVerbose 'Restarting MSExchangeTransport after antispam agent install'
                Restart-Service MSExchangeTransport -Force -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                Write-MyVerbose 'MSExchangeTransport restarted after antispam agent install'
            }
            catch {
                Write-MyWarning ('Failed to install antispam agents: {0}' -f $_.Exception.Message)
                return
            }
        }
        else {
            Write-MyVerbose ('Antispam agents already installed ({0} found), skipping install script' -f $existingAgents.Count)
        }

        # Configure agents: only Recipient Filter enabled, all others disabled
        $needsTransportRestart = $false
        $allAgents = Get-TransportAgent -ErrorAction SilentlyContinue |
                     Where-Object { $_.Identity -like '*Filter*' -or $_.Identity -like '*Antispam*' -or $_.Identity -like '*Reputation*' }
        foreach ($agent in $allAgents) {
            $isRecipientFilter = $agent.Identity -like '*Recipient Filter*'
            if ($isRecipientFilter) {
                if (-not $agent.Enabled) {
                    Enable-TransportAgent -Identity $agent.Identity -Confirm:$false -WarningAction Ignore -ErrorAction SilentlyContinue *>&1 | Out-Null
                    Write-MyOutput ('Enabled: {0}' -f $agent.Identity)
                    $needsTransportRestart = $true
                }
                else {
                    Write-MyVerbose ('Already enabled: {0}' -f $agent.Identity)
                }
            }
            else {
                if ($agent.Enabled) {
                    Disable-TransportAgent -Identity $agent.Identity -Confirm:$false -WarningAction Ignore -ErrorAction SilentlyContinue *>&1 | Out-Null
                    Write-MyOutput ('Disabled: {0}' -f $agent.Identity)
                    $needsTransportRestart = $true
                }
                else {
                    Write-MyVerbose ('Already disabled: {0}' -f $agent.Identity)
                }
            }
        }
        if ($needsTransportRestart) {
            Write-MyVerbose 'Restarting MSExchangeTransport after agent configuration change'
            Restart-Service MSExchangeTransport -Force -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            Write-MyVerbose 'MSExchangeTransport restarted after agent configuration change'
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
                    Set-ReceiveConnector -Identity "$server\$intName" -RemoteIPRanges $State['RelaySubnets'] `
                        -AuthMechanism Tls -ProtocolLoggingLevel Verbose -ErrorAction Stop
                    Write-MyVerbose 'Internal relay connector already exists — RemoteIPRanges, TLS and logging updated'
                }
                else {
                    New-ReceiveConnector -Name $intName -Server $server -TransportRole FrontendTransport `
                        -RemoteIPRanges $State['RelaySubnets'] -Bindings '0.0.0.0:25' `
                        -PermissionGroups AnonymousUsers -AuthMechanism Tls `
                        -ProtocolLoggingLevel Verbose -ErrorAction Stop | Out-Null
                    Write-MyOutput 'Internal relay connector created (TLS offered, accepted domains only, no external relay right)'
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
                    Set-ReceiveConnector -Identity "$server\$extName" -RemoteIPRanges $State['ExternalRelaySubnets'] `
                        -AuthMechanism Tls -ProtocolLoggingLevel Verbose -ErrorAction Stop
                    Write-MyVerbose 'External relay connector already exists — RemoteIPRanges, TLS and logging updated'
                    $connObj = $existing
                }
                else {
                    # Capture the returned object directly to avoid a race condition where
                    # Get-ReceiveConnector fails immediately after creation (Exchange AD not yet updated).
                    $connObj = New-ReceiveConnector -Name $extName -Server $server -TransportRole FrontendTransport `
                        -RemoteIPRanges $State['ExternalRelaySubnets'] -Bindings '0.0.0.0:25' `
                        -PermissionGroups AnonymousUsers -AuthMechanism Tls `
                        -ProtocolLoggingLevel Verbose -ErrorAction Stop
                }
                # Fallback: if the object is somehow null, retry Get-ReceiveConnector with backoff
                if (-not $connObj) {
                    for ($retry = 1; $retry -le 3 -and -not $connObj; $retry++) {
                        Write-MyVerbose ('Waiting for external relay connector to register in Exchange (attempt {0}/3)...' -f $retry)
                        Start-Sleep -Seconds 5
                        $connObj = Get-ReceiveConnector -Identity "$server\$extName" -ErrorAction SilentlyContinue
                    }
                }
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

    function Invoke-EOMT {
        if (-not $State['RunEOMT']) {
            Write-MyVerbose 'RunEOMT not specified, skipping EOMT'
            return
        }
        Write-MyOutput 'Running CSS-Exchange Emergency Mitigation Tool (EOMT)'
        $eomtPath = Join-Path $State['InstallPath'] 'EOMT.ps1'
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
            $cert = Import-ExchangeCertificate -FileData ([System.IO.File]::ReadAllBytes($pfxPath)) -Password $secPwd -PrivateKeyExportable $true -ErrorAction Stop
            Write-MyOutput ('Certificate imported: {0} (Thumbprint: {1})' -f $cert.Subject, $cert.Thumbprint)

            # Detect wildcard certificate (CN=* or SAN with *.domain)
            $isWildcard = ($cert.Subject -match 'CN=\*') -or ($cert.SubjectAlternativeNames -match '^\*\.')
            if ($isWildcard) {
                # Wildcard: enable for IIS and SMTP only (IMAP/POP use specific SANs)
                Enable-ExchangeCertificate -Thumbprint $cert.Thumbprint -Services IIS,SMTP -Force -ErrorAction Stop
                Write-MyOutput ('Wildcard certificate enabled for IIS and SMTP services')
            }
            else {
                # Named certificate: also enable for IMAP and POP
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
        $dlDomain       = $State['DownloadDomain']
        $nsResolves     = $false
        $dlResolves     = $false
        try { [System.Net.Dns]::GetHostEntry($ns) | Out-Null; $nsResolves = $true } catch { }
        if ($dlDomain) { try { [System.Net.Dns]::GetHostEntry($dlDomain) | Out-Null; $dlResolves = $true } catch { } }
        if (-not $nsResolves) {
            Write-MyVerbose ('Namespace {0} not resolvable — adding temporary hosts entry to suppress VDir confirmation prompt' -f $ns)
            "`r`n127.0.0.1`t$ns`t$tempHostsMark" | Add-Content -Path $hostsFile -Encoding ASCII -ErrorAction SilentlyContinue
        }
        if ($dlDomain -and -not $dlResolves) {
            Write-MyVerbose ('Download domain {0} not resolvable — adding temporary hosts entry' -f $dlDomain)
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
                    Set-OwaVirtualDirectory -Identity "$server\owa (Default Web Site)" `
                        -ExternalDownloadHostName $dlDomain -InternalDownloadHostName $dlDomain `
                        -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
                    Write-MyVerbose ('OWA Download Domains configured: {0} (CVE-2021-1730 mitigation)' -f $dlDomain)
                    $changed++
                }
            }
            catch { Write-MyWarning ('OWA Download Domains: {0}' -f $_.Exception.Message); $errors++ }
        }

        # ECP
        try {
            $vd = Get-EcpVirtualDirectory -Identity "$server\ecp (Default Web Site)" -ADPropertiesOnly -ErrorAction Stop
            if ((Test-VdirUrl $vd.InternalUrl "https://$ns/ecp") -and (Test-VdirUrl $vd.ExternalUrl "https://$ns/ecp")) {
                Write-MyVerbose 'ECP: URLs already set, skipping'
            } else {
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
        }
        catch { Write-MyVerbose ('MAPI auth methods not supported on this build: {0}' -f $_.Exception.Message) }

        # Autodiscover SCP — always use autodiscover.<parent-domain>, not the namespace hostname
        try {
            $cas = Get-ClientAccessService -Identity $server -ErrorAction Stop
            $nsParts   = $ns -split '\.'
            $scpHost   = if ($nsParts[0] -eq 'autodiscover') { $ns } else { 'autodiscover.' + ($nsParts[1..($nsParts.Length-1)] -join '.') }
            $scpTarget = "https://$scpHost/Autodiscover/Autodiscover.xml"
            if ([string]$cas.AutoDiscoverServiceInternalUri -eq $scpTarget) {
                Write-MyVerbose 'Autodiscover SCP: already set, skipping'
            } else {
                Set-ClientAccessService -Identity $server `
                    -AutoDiscoverServiceInternalUri $scpTarget `
                    -ErrorAction Stop -WarningAction SilentlyContinue
                Write-MyVerbose 'Autodiscover SCP configured'
                $changed++
            }
        }
        catch { Write-MyWarning ('Autodiscover SCP: {0}' -f $_.Exception.Message); $errors++ }

        # Remove temporary hosts entries if any were added
        if (-not $nsResolves -or ($dlDomain -and -not $dlResolves)) {
            try {
                $hostsLines = Get-Content $hostsFile -ErrorAction Stop
                $hostsLines = $hostsLines | Where-Object { $_ -notlike "*$tempHostsMark*" }
                Set-Content $hostsFile $hostsLines -Encoding ASCII -ErrorAction Stop
                Write-MyVerbose 'Temporary hosts entries removed'
            }
            catch { Write-MyVerbose ('Could not remove temporary hosts entries: {0}' -f $_.Exception.Message) }
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
        $hcPath = Join-Path $State['InstallPath'] 'HealthChecker.ps1'
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
                    Where-Object { $_.LastWriteTime -ge $hcBefore -and $_.Extension -match '\.html?' -and $_.Name -match '^(ExchangeAllServersReport|HealthChecker)' } |
                    Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($hcReport) {
                    # Rename to include server name for easy identification
                    $newHcName = '{0}_{1}' -f $env:COMPUTERNAME, $hcReport.Name
                    $newHcPath = Join-Path $State['ReportsPath'] $newHcName
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
                Write-MyOutput 'NOTE: HealthChecker "Exchange Server Membership" may show blank/failed results in this run — the current process token was created before Exchange setup added the computer account to "Exchange Servers". Re-run HealthChecker after the next reboot for accurate membership results.'
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
        $saPath = Join-Path $State['InstallPath'] 'SetupAssist.ps1'
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
        $slrPath = Join-Path $State['InstallPath'] 'SetupLogReviewer.ps1'
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

    function New-InstallationReport {
        Write-MyOutput 'Generating Installation Report'
        $reportPath = Join-Path $State['ReportsPath'] ('{0}_InstallationReport_{1}.html' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMddHHmmss'))
        $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

        function Format-Badge($text, $type) {
            $colors = @{ ok='background:#107c10;color:#fff'; warn='background:#d83b01;color:#fff'; fail='background:#c50f1f;color:#fff'; info='background:#0078d4;color:#fff'; na='background:#8a8886;color:#fff' }
            '<span style="display:inline-block;padding:2px 10px;border-radius:12px;font-size:12px;font-weight:600;{0}">{1}</span>' -f $colors[$type.ToLower()], $text
        }
        function New-HtmlSection($id, $title, $content) {
            '<section id="{0}" class="section"><h2 class="section-title">{1}</h2><div class="section-body">{2}</div></section>' -f $id, $title, $content
        }

        # ── 1. INSTALLATION PARAMETERS ────────────────────────────────────────
        $instRows = [System.Collections.Generic.List[string]]::new()
        $instMode = if ($State['InstallEdge']) { 'Edge Transport' } elseif ($State['InstallRecipientManagement']) { 'Recipient Management Tools' } elseif ($State['InstallManagementTools']) { 'Management Tools' } elseif ($State['StandaloneOptimize']) { 'Standalone Optimize' } elseif ($State['NoSetup']) { 'Optimization Only' } else { 'Mailbox Server' }
        $instRows.Add('<tr><td>Script Version</td><td>Install-Exchange15.ps1 v{0}</td></tr>' -f $ScriptVersion)
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
        $sysContent = '<table class="data-table"><tr><th>Property</th><th>Value</th><th>Detail / Status</th></tr>{0}</table>' -f ($sysRows -join '')

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
        $adContent = '<table class="data-table"><tr><th>Property</th><th>Value</th><th>Status</th></tr>{0}</table>' -f ($adRows -join '')

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
                $daysLeft = ($_.NotAfter - (Get-Date)).Days
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
                $maxSendMB = [math]::Round($transCfg2.MaxSendSize.Value.ToBytes() / 1MB, 0)
                $maxRecvMB = [math]::Round($transCfg2.MaxReceiveSize.Value.ToBytes() / 1MB, 0)
                $sizeBadge = if ($maxSendMB -ge 50) { Format-Badge '✓' 'ok' } else { Format-Badge 'Default 25 MB' 'warn' }
                $exchOptRows.Add(('<tr><td>Max Message Size</td><td>Send: {0} MB / Recv: {1} MB</td><td>≥ 50 MB (modern workflow files)</td><td>{2}</td></tr>' -f $maxSendMB, $maxRecvMB, $sizeBadge))

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

        $exContent = @"
<table class="data-table"><tr><th>Property</th><th>Value</th><th>Status</th></tr>{0}</table>
<h3 class="subsection">Virtual Directory URLs</h3>
<table class="data-table">{1}</table>
<h3 class="subsection">Mailbox Databases</h3>
<table class="data-table">{2}</table>
<h3 class="subsection">Receive Connectors</h3>
<table class="data-table">{3}</table>
<h3 class="subsection">Certificates</h3>
<table class="data-table">{4}</table>
<h3 class="subsection">Exchange Optimizations</h3>
<table class="data-table">{5}</table>
"@ -f ($exRows -join ''), ($vdirRows -join ''), ($dbRows -join ''), ($connRows -join ''), ($certRows -join ''), ($exchOptRows -join '')

        # ── 5. SECURITY SETTINGS (with Exchange best-practice + reference column) ─
        $secRows = [System.Collections.Generic.List[string]]::new()
        function Get-SecRegVal($path, $name) { try { (Get-ItemProperty -Path $path -Name $name -ErrorAction Stop).$name } catch { $null } }
        function Format-RefLink($url, $label) { '<a href="{0}" target="_blank" style="font-size:0.85em;white-space:nowrap">{1} ↗</a>' -f $url, $label }

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
        try {
            $nbNics = @(Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True' -ErrorAction Stop)
            $nbDisabled = ($nbNics | Where-Object { $_.TcpipNetbiosOptions -eq 2 }).Count
            $nbBadge = if ($nbNics.Count -gt 0 -and $nbDisabled -eq $nbNics.Count) { Format-Badge 'Disabled ✓' 'ok' } else { Format-Badge ('{0}/{1} disabled' -f $nbDisabled, $nbNics.Count) 'warn' }
            $secRows.Add(('<tr><td>NetBIOS over TCP/IP</td><td>{0} of {1} NICs disabled</td><td>Disabled on all NICs (reduces LLMNR/NBT-NS attack surface)</td><td>{2}</td><td>{3}</td><td>CIS §18 / BSI</td></tr>' -f $nbDisabled, $nbNics.Count, $nbBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/disable-netbios-tcp-ip-using-dhcp' 'MS Learn')))
        } catch { }

        # Root CA auto-update
        $rootAU = Get-SecRegVal 'HKLM:\SOFTWARE\Policies\Microsoft\SystemCertificates\AuthRoot' 'DisableRootAutoUpdate'
        $rootAUBadge = if ($rootAU -ne 1) { Format-Badge 'Enabled ✓' 'ok' } else { Format-Badge 'Disabled by policy!' 'warn' }
        $secRows.Add(('<tr><td>Root CA Auto-Update</td><td>DisableRootAutoUpdate = {0}</td><td>Must not be disabled (Exchange Online / M365 connectivity)</td><td>{1}</td><td>{2}</td><td>MS Exchange</td></tr>' -f $rootAU, $rootAUBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/security/trusted-root/release-notes' 'MS Trusted Root')))

        # Extended Protection (OWA VDir)
        if (-not $State['InstallEdge']) {
            try {
                $owaVdir = Get-OwaVirtualDirectory -Server $env:COMPUTERNAME -ADPropertiesOnly -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($owaVdir) {
                    $epVal = $owaVdir.ExtendedProtectionTokenChecking
                    $epBadge = if ($epVal -in 'Require','Allow') { Format-Badge "$epVal ✓" 'ok' } else { Format-Badge "$epVal (risk)" 'warn' }
                    $secRows.Add(('<tr><td>Extended Protection (OWA)</td><td>{0}</td><td>Require or Allow</td><td>{1}</td><td>{2}</td><td>MS Exchange</td></tr>' -f $epVal, $epBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/exchange/plan-and-deploy/post-installation-tasks/security-best-practices/exchange-extended-protection' 'MS Learn')))
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

        $secContent = '<table class="data-table"><tr><th>Setting</th><th>Current Value</th><th>Exchange Recommendation</th><th>Status</th><th>Reference</th><th>CIS / BSI</th></tr>{0}</table>' -f ($secRows -join '')

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

        $perfContent = '<table class="data-table"><tr><th>Setting</th><th>Current Value</th><th>Exchange Recommendation</th><th>Status</th></tr>{0}</table>' -f ($perfRows -join '')

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
        $rbacContent = '<table class="data-table">{0}</table>' -f ($rbacRows -join '')

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
            $certs = @(Get-ExchangeCertificate -Server $env:COMPUTERNAME -ErrorAction Stop)
            $expiring = @($certs | Where-Object { ($_.NotAfter - (Get-Date)).Days -le 90 })
            if ($expiring.Count -gt 0) { Format-Badge ('{0} expiring within 90 days' -f $expiring.Count) 'warn' }
            else { Format-Badge ('{0} certificate(s) — all valid' -f $certs.Count) 'ok' }
        } catch { Format-Badge 'Could not query' 'na' }

        $vdirStatus = if ($State['Namespace']) { Format-Badge ('Configured ({0})' -f $State['Namespace']) 'ok' } else { Format-Badge 'Not configured' 'warn' }
        $suStatus   = if ($State['IncludeFixes']) { Format-Badge 'Enabled' 'ok' } else { Format-Badge 'Skipped' 'warn' }
        $hcStatus   = if ($State['SkipHealthCheck']) { Format-Badge 'Skipped' 'na' } elseif ($State['HCReportPath']) { Format-Badge 'Completed — see section below' 'ok' } else { Format-Badge 'Failed / not found' 'warn' }
        $modeStatus = if ($State['ConfigDriven']) { Format-Badge 'Autopilot (fully automated)' 'info' } else { Format-Badge 'Copilot (interactive)' 'info' }

        $mgmtRows.Add('<tr><td style="width:220px"><strong>Exchange Version</strong></td><td>{0}</td><td>{1}</td></tr>' -f $exVersion, (Format-Badge 'Installed' 'ok'))
        $mgmtRows.Add('<tr><td><strong>Server</strong></td><td>{0}</td><td>{1}</td></tr>' -f ('{0} ({1})' -f $env:COMPUTERNAME, (try { '{0}.{1}' -f (Get-CimInstance Win32_ComputerSystem -EA SilentlyContinue).DNSHostName, (Get-CimInstance Win32_ComputerSystem -EA SilentlyContinue).Domain } catch { '' })), (Format-Badge 'OK' 'ok'))
        $mgmtRows.Add('<tr><td><strong>Organization</strong></td><td>{0}</td><td></td></tr>' -f $State['OrganizationName'])
        $mgmtRows.Add('<tr><td><strong>Installation Mode</strong></td><td>{0}</td><td>{1}</td></tr>' -f $instMode, $modeStatus)
        $mgmtRows.Add('<tr><td><strong>Virtual Directory URLs</strong></td><td>{0}</td><td>{1}</td></tr>' -f $State['Namespace'], $vdirStatus)
        $mgmtRows.Add('<tr><td><strong>Security Hardening</strong></td><td>{0} OK / {1} warnings / {2} critical</td><td>{3}</td></tr>' -f $secOK, $secWarn, $secFail, $secStatusBadge)
        $mgmtRows.Add('<tr><td><strong>Performance Settings</strong></td><td>{0} OK / {1} to review</td><td>{2}</td></tr>' -f $perfOK, $perfWarn, $perfStatusBadge)
        $mgmtRows.Add('<tr><td><strong>Certificates</strong></td><td></td><td>{0}</td></tr>' -f $certStatus)
        $mgmtRows.Add('<tr><td><strong>Security Updates</strong></td><td></td><td>{0}</td></tr>' -f $suStatus)
        $mgmtRows.Add('<tr><td><strong>HealthChecker</strong></td><td></td><td>{0}</td></tr>' -f $hcStatus)
        $mgmtRows.Add('<tr><td><strong>Report Generated</strong></td><td>{0}</td><td></td></tr>' -f $reportDate)

        # Action items — surface any WARN/FAIL as bullet list
        $actionItems = [System.Collections.Generic.List[string]]::new()
        if ($secFail -gt 0)  { $actionItems.Add('<li><strong>Security:</strong> {0} critical finding(s) require immediate attention — see Security Settings section.</li>' -f $secFail) }
        if ($secWarn -gt 0)  { $actionItems.Add('<li><strong>Security:</strong> {0} warning(s) — review Security Settings section.</li>' -f $secWarn) }
        if ($perfWarn -gt 0) { $actionItems.Add('<li><strong>Performance:</strong> {0} setting(s) below recommendation — review Performance &amp; Tuning section.</li>' -f $perfWarn) }
        if (-not $State['Namespace']) { $actionItems.Add('<li><strong>Virtual Directories:</strong> No external namespace configured. OWA/ECP/EWS URLs may still point to server hostname.</li>') }
        if (-not $State['IncludeFixes']) { $actionItems.Add('<li><strong>Security Updates:</strong> Exchange Security Update installation was skipped. Apply the latest SU manually.</li>') }

        $actionHtml = if ($actionItems.Count -gt 0) {
            '<h3 class="subsection">Action Items</h3><ul style="margin:8px 0 0 20px;line-height:1.8;font-size:13px">' + ($actionItems -join '') + '</ul>'
        } else {
            '<p style="color:#107c10;font-size:13px;margin-top:12px">&#10003; No critical action items identified.</p>'
        }

        $mgmtContent = '<table class="data-table"><tr><th>Item</th><th>Detail</th><th>Status</th></tr>{0}</table>{1}' -f ($mgmtRows -join ''), $actionHtml

        $sections = @(
            (New-HtmlSection 'summary'      'Management Summary'        $mgmtContent)
            (New-HtmlSection 'params'       'Installation Parameters'   ('<table class="data-table">{0}</table>' -f ($instRows -join '')))
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
    <div style="font-size:12px;opacity:.65;margin-top:4px">Generated by Install-Exchange15.ps1 v$ScriptVersion</div>
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
<footer>Exchange Server Installation Report &bull; $env:COMPUTERNAME &bull; $reportDate &bull; Install-Exchange15.ps1 v$ScriptVersion &bull; promiseIT</footer>
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

    function Get-RBACReport {
        Write-MyOutput 'Generating RBAC role group membership report'
        $ts = Get-Date -Format 'yyyyMMddHHmmss'
        $reportPath = Join-Path $State['ReportsPath'] "${env:COMPUTERNAME}_RBACReport_${ts}.txt"

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
            Write-Host ('  Install-Exchange15 v{0} — Exchange Optimizations' -f $script:ScriptVersion) -ForegroundColor Cyan
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
                $state = if ($sel[$LastKey]) { 'ENABLED' } else { 'DISABLED' }
                Write-Host ('  [{0}] {1}  ({2})' -f $LastKey, $opt.Label, $state) -ForegroundColor Yellow
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

        # --- Phase 2: Per-update prompt (interactive only; Autopilot installs all) ---
        $approvedKBs = @()
        $installAll  = -not $isInteractive   # non-interactive session: approve everything immediately

        for ($idx = 0; $idx -lt $candidates.Count; $idx++) {
            $u = $candidates[$idx]
            $label = '[{0}/{1}] {2} — {3}' -f ($idx + 1), $candidates.Count, $u.Title, $(if ($u.Severity) { $u.Severity } else { 'Unknown' })

            if ($installAll) {
                Write-MyOutput ('Will install: {0}' -f $label)
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
                Write-Host ('  Install? [Y/N/A=all/S=skip] (auto-No in {0}s) ' -f $WU_PROMPT_TIMEOUT_SEC) -NoNewline -ForegroundColor DarkCyan
                while ($sw.Elapsed.TotalSeconds -lt $WU_PROMPT_TIMEOUT_SEC) {
                    $secsLeft = [int]($WU_PROMPT_TIMEOUT_SEC - $sw.Elapsed.TotalSeconds)
                    Write-Progress -Id 2 -Activity 'Windows Update' `
                        -Status ('Auto-No in {0}s  |  Y = install  |  N = skip  |  A = all  |  S = skip remaining' -f $secsLeft) `
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
                $ans = (Read-Host '  Install? [Y=yes / N=no / A=all / S=skip remaining] (default: Y)').Trim().ToUpper()
                if ($ans -eq '') { $ans = 'Y' }
            }
            switch ($ans) {
                'A' { $installAll = $true;  if ($u.KB) { $approvedKBs += $u.KB }; Write-MyOutput ('Approved (all): {0}' -f $u.Title) }
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
                $session  = New-Object -ComObject Microsoft.Update.Session
                $searcher = $session.CreateUpdateSearcher()
                $filter   = ($kbs | ForEach-Object { "KBArticleID='$_'" }) -join ' or '
                $found    = $searcher.Search("IsInstalled=0 and ($filter)")
                if ($found.Updates.Count -eq 0) { return @{ Installed=0; RebootRequired=$false } }
                $dl = $session.CreateUpdateDownloader()
                $dl.Updates = $found.Updates
                $dl.Download() | Out-Null
                $inst       = $session.CreateUpdateInstaller()
                $inst.Updates = $found.Updates
                $instResult = $inst.Install()
                @{ Installed = $found.Updates.Count; RebootRequired = $instResult.RebootRequired; ResultCode = $instResult.ResultCode }
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
            $statusText = 'Installing {0} update(s) — {1}s elapsed, {2}s remaining (X = cancel)' -f $approvedKBs.Count, $elapsed, $remaining
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
        # in InstallPath before running, or apply via Windows Update / WSUS.
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
        $hcPath = Join-Path $State['InstallPath'] 'HealthChecker.ps1'
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
                $suPath = Join-Path $State['InstallPath'] $su.FileName
                if (-not (Test-Path $suPath)) {
                    if ($su.URL) {
                        Write-MyOutput ('Downloading {0}' -f $su.KB)
                        $null = Get-MyPackage -Package $su.KB -URL $su.URL -FileName $su.FileName -InstallPath $State['InstallPath']
                    }
                    if (-not (Test-Path $suPath)) {
                        Write-MyWarning ('Exchange SU {0}: installer not available for automatic download.' -f $su.KB)
                        Write-MyOutput  ('  Download:  https://support.microsoft.com/help/{0}' -f ($su.KB -replace '^KB', ''))
                        Write-MyOutput  ('  Place EXE: {0}' -f $suPath)

                        # Interactive countdown — user has 5 min to place the file, then ENTER to install.
                        # Autopilot / non-interactive: skip silently (no file available, no reboot loop).
                        if ([Environment]::UserInteractive -and -not $State['ConfigDriven']) {
                            Write-MyOutput 'Place the installer and press ENTER to continue, or wait 5 min to skip:'
                            $suTotalSecs = 300
                            $suDeadline  = [DateTime]::Now.AddSeconds($suTotalSecs)
                            try {
                                try { $host.UI.RawUI.FlushInputBuffer() } catch { }
                                while ([DateTime]::Now -lt $suDeadline) {
                                    $secsLeft = [int]($suDeadline - [DateTime]::Now).TotalSeconds
                                    Write-Progress -Id 2 -Activity ('Exchange SU {0}' -f $su.KB) `
                                        -Status ('Place {0} in {1} then ENTER  |  auto-skip in {2}s' -f $su.FileName, $State['InstallPath'], $secsLeft) `
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
                    $rc = Invoke-Process -FilePath $State['InstallPath'] -FileName $su.FileName -ArgumentList '/passive'
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
        $hcPath = Join-Path $State['InstallPath'] 'HealthChecker.ps1'
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

        # Run CSS-Exchange Add-PermissionForEMT.ps1 if available
        $emtScript = Join-Path $State['InstallPath'] 'Add-PermissionForEMT.ps1'
        $emtUrl = 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/Add-PermissionForEMT.ps1'
        if (-not (Test-Path $emtScript)) {
            try {
                Write-MyVerbose ('Downloading Add-PermissionForEMT from {0}' -f $emtUrl)
                Start-BitsTransfer -Source $emtUrl -Destination $emtScript -ErrorAction Stop
            }
            catch {
                Write-MyWarning ('Could not download Add-PermissionForEMT.ps1: {0}' -f $_.Exception.Message)
            }
        }
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
                }
                Set-Service -Name $svc.Name -StartupType Disabled -ErrorAction SilentlyContinue
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
        # Microsoft recommends disabling TCP offload features on Exchange servers
        Write-MyOutput 'Disabling TCP Chimney and Task Offload settings'
        try {
            $null = netsh.exe int tcp set global chimney=disabled 2>&1
            if ($LASTEXITCODE -ne 0) { Write-MyWarning ('netsh chimney=disabled exited with code {0}' -f $LASTEXITCODE) }
            $null = netsh.exe int tcp set global autotuninglevel=restricted 2>&1
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
            $null = reg load $defaultHiveKey $defaultHive 2>$null
            if (Test-Path "Registry::$defaultHiveKey\Software\Microsoft\ServerManager") {
                Set-ItemProperty -Path "Registry::$defaultHiveKey\Software\Microsoft\ServerManager" -Name 'DoNotOpenServerManagerAtLogon' -Value 1 -Type DWord -ErrorAction SilentlyContinue
            }
            else {
                New-Item -Path "Registry::$defaultHiveKey\Software\Microsoft\ServerManager" -Force -ErrorAction SilentlyContinue | Out-Null
                New-ItemProperty -Path "Registry::$defaultHiveKey\Software\Microsoft\ServerManager" -Name 'DoNotOpenServerManagerAtLogon' -Value 1 -PropertyType DWord -Force -ErrorAction SilentlyContinue | Out-Null
            }
            $null = reg unload $defaultHiveKey 2>$null
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
            Write-MyOutput 'IPv4 over IPv6 preference set (DisabledComponents = 0x20) — effective after reboot'
            $State['RebootRequired'] = $true
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
                    Write-MyVerbose ('NetBIOS disable on {0}: reboot required' -f $nic.Description)
                    $changed++
                    $State['RebootRequired'] = $true
                } else {
                    Write-MyWarning ('NetBIOS disable on {0} returned code {1}' -f $nic.Description, $result)
                }
            }
            Write-MyVerbose ('NetBIOS disabled on {0} NIC(s)' -f $changed)
        } catch {
            Write-MyWarning ('Failed to disable NetBIOS: {0}' -f $_.Exception.Message)
        }
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
        }
    }

    function Enable-ECC {
        # https://learn.microsoft.com/en-us/exchange/architecture/client-access/certificates?view=exchserver-2019#elliptic-curve-cryptography-certificates-support-in-exchange-server
        Write-MyVerbose 'Enabling Elliptic Curve Cryptography support'

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
            Write-MyVerbose 'Restarting W3SVC and WAS'
            Restart-Service -Name W3SVC, WAS -Force -WarningAction SilentlyContinue
        }
        else {
            Write-MyVerbose ('No override configuration for ECC found')
        }
    }

    function Enable-CBC {
        # https://support.microsoft.com/en-us/topic/enable-support-for-aes256-cbc-encrypted-content-in-exchange-server-august-2023-su-add63652-ee17-4428-8928-ddc45339f99e
        Write-MyVerbose 'Enabling AES256-CBC mode of encryption support'

        $Override = Get-SettingOverride | Where-Object { ($_.SectionName -eq "EnableEncryptionAlgorithmCBC") -and ($_.Parameters -eq "Enabled=True") }
        if ( $Override) {
            Write-MyVerbose ('Configuration for CBC already configured')
        }
        else {
            New-SettingOverride -Name "EnableEncryptionAlgorithmCBC" -Parameters @("Enabled=True") -Component Encryption -Reason "Enable CBC encryption" -Section EnableEncryptionAlgorithmCBC | Out-Null
            Get-ExchangeDiagnosticInfo -Process Microsoft.Exchange.Directory.TopologyService -Component VariantConfiguration -Argument Refresh | Out-Null
            Write-MyVerbose 'Restarting W3SVC and WAS'
            Restart-Service -Name W3SVC, WAS -Force -WarningAction SilentlyContinue
        }
    }

    function Enable-AMSI {
        param(
            [string[]]$ConfigParam = @("EnabledEcp=True", "EnabledEws=True", "EnabledOwa=True", "EnabledPowerShell=True")
        )
        # https://learn.microsoft.com/en-us/exchange/antispam-and-antimalware/amsi-integration-with-exchange?view=exchserver-2019#enable-exchange-server-amsi-body-scanning
        # Edge Transport is not domain-joined and has no org connection; New-SettingOverride would fail.
        if ($State['InstallEdge']) { Write-MyVerbose 'Enable-AMSI: skipped (Edge Transport — no org connection)'; return }
        Write-MyVerbose 'Enabling AMSI body scanning for OWA, ECP, EWS and PowerShell'

        New-SettingOverride -Name "EnableAMSIBodyScan" -Component Cafe -Section AmsiRequestBodyScanning -Parameters $ConfigParam -Reason "Enabling AMSI body Scan"
        Get-ExchangeDiagnosticInfo -Process Microsoft.Exchange.Directory.TopologyService -Component VariantConfiguration -Argument Refresh
        Write-MyVerbose 'Restarting W3SVC and WAS'
        Restart-Service -Name W3SVC, WAS -Force -WarningAction SilentlyContinue
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
        # Exchange 2016 / 2019 pre-CU14: downloads and runs ExchangeExtendedProtection.ps1 from CSS-Exchange.
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

        # Exchange 2016 / 2019 pre-CU14: configure via CSS-Exchange ExchangeExtendedProtection.ps1
        Write-MyOutput 'Enabling Extended Protection via CSS-Exchange ExchangeExtendedProtection.ps1'
        $epPath = Join-Path $State['InstallPath'] 'ExchangeExtendedProtection.ps1'
        $epUrl  = 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/ExchangeExtendedProtection.ps1'

        if (-not (Test-Path $epPath)) {
            try {
                Invoke-WebDownload -Uri $epUrl -OutFile $epPath
                Write-MyVerbose ('ExchangeExtendedProtection.ps1 downloaded, SHA256: {0}' -f (Get-FileHash $epPath -Algorithm SHA256).Hash)
            }
            catch {
                Write-MyWarning ('Could not download ExchangeExtendedProtection.ps1: {0}' -f $_.Exception.Message)
                return
            }
        }

        try {
            $epArgs = @('-ExchangeServerNames', $env:computername)
            if ($State['DoNotEnableEP_FEEWS']) { $epArgs += '-SkipEWS' }
            & $epPath @epArgs *>&1 | ForEach-Object { Write-ToTranscript ([string]$_) }
        }
        catch {
            Write-MyWarning ('ExchangeExtendedProtection.ps1 failed: {0}' -f $_.Exception.Message)
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
        }
        else {
            Write-MyVerbose 'Windows Defender not installed'
        }
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
                else {
                    Write-Verbose ('MSExchangeAutodiscoverAppPool not found, waiting a bit ..')
                    Start-Sleep -Seconds 10
                }
            } while ($true)
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
        }

        # Toggle definitions: Key=letter, Name=parameter name, Default=initial state
        # TLS 1.3 requires Windows Server 2022 or later
        $tls13Default = [System.Version]$FullOSVersion -ge [System.Version]$WS2022_PREFULL

        # Name = parameter/cfg key; Label = display text shown in menu
        $toggleDefs = [ordered]@{
            'A' = @{ Name='Autopilot';             Label='Autopilot (auto-reboot)';      Default=$true  }
            'B' = @{ Name='IncludeFixes';           Label='Install Exchange SU';           Default=$true  }
            'C' = @{ Name='DisableSSL3';            Label='Disable SSL 3.0';               Default=$true  }
            'D' = @{ Name='DisableRC4';             Label='Disable RC4';                   Default=$true  }
            'E' = @{ Name='EnableECC';              Label='Enable ECC ciphers';            Default=$true  }
            'F' = @{ Name='NoCBC';                  Label='Disable CBC(not recommended)';  Default=$false }
            'G' = @{ Name='EnableAMSI';             Label='Enable AMSI';                   Default=$true  }
            'H' = @{ Name='EnableTLS12';            Label='Enforce TLS 1.2';               Default=$true  }
            'I' = @{ Name='DoNotEnableEP';          Label='No Extended Protection';        Default=$false }
            'J' = @{ Name='EnableTLS13';            Label='Enable TLS 1.3';                Default=$tls13Default }
            'K' = @{ Name='DiagnosticData';         Label='Send diagnostic data';          Default=$false }
            'L' = @{ Name='Lock';                   Label='Lock screen during install';    Default=$false }
            'M' = @{ Name='SkipRolesCheck';         Label='Skip AD roles check';           Default=$false }
            'N' = @{ Name='PreflightOnly';          Label='Preflight only (no install)';   Default=$false }
            'O' = @{ Name='NoCheckpoint';           Label='Skip restore checkpoints';      Default=$false }
            'P' = @{ Name='SkipHealthCheck';        Label='Skip HealthChecker';            Default=$false }
            'Q' = @{ Name='NoNet481';               Label='Skip .NET 4.8.1 install';       Default=$false }
            'R' = @{ Name='InstallWindowsUpdates';  Label='Install Windows Updates';       Default=$true  }
            'S' = @{ Name='RunEOMT';                Label='Run EOMT (CVE mitigations)';    Default=$false }
            'T' = @{ Name='WaitForADSync';          Label='Wait for AD replication';       Default=$false }
        }

        # Toggles disabled per mode (letters that cannot be toggled in that mode)
        # T=WaitForADSync only makes sense in modes that run PrepareAD (1 and 5)
        # S=RunEOMT only makes sense for modes that run Exchange post-config (1, 5, 6)
        $disabledToggles = @{
            1 = @()
            2 = @('I','G','S','T')
            3 = @('B','C','D','E','F','G','H','I','J','K','L','M','N','P','Q','R','S','T')
            4 = @('B','I','G','S','T')
            5 = @()
            6 = @('B','I','K','M','N','Q','R','T')                                        # Standalone: no setup, no PrepareAD
        }

        # Initialize toggle states from defaults
        $toggleState = @{}
        foreach ($k in $toggleDefs.Keys) { $toggleState[$k] = $toggleDefs[$k].Default }

        $selectedMode = 0

        # Returns extra letters that should be disabled based on current toggle state
        function Get-DynamicDisabled {
            param([hashtable]$TS)
            $extra = @()
            if (-not $TS['A']) { $extra += 'L' }                          # Lock requires Autopilot
            if (-not $TS['H']) { $extra += 'J'; $extra += 'F' }           # TLS 1.3 + CBC require TLS 1.2
            if ($TS['N'])      { $extra += @('B','O','P','Q','S','T') }   # PreflightOnly: post-install irrelevant
            return $extra
        }

        function Write-MenuLine {
            param([string]$Line, [System.ConsoleColor]$Color = [System.ConsoleColor]::White)
            Write-Host $Line -ForegroundColor $Color
        }

        function Draw-Menu {
            param([int]$Mode, [hashtable]$ToggState, [string]$StatusMsg = '', [array]$ExtraDisabled = @())
            Clear-Host
            Write-MenuLine ('=' * 60) Cyan
            Write-MenuLine ('  Install-Exchange15 v{0}  —  Copilot' -f $ScriptVersion) Cyan
            Write-MenuLine ('=' * 60) Cyan
            Write-Host ''
            Write-MenuLine '  Installation Mode:' Yellow
            for ($i = 1; $i -le 6; $i++) {
                $marker = if ($Mode -eq $i) { '>' } else { ' ' }
                $color  = if ($Mode -eq $i) { [System.ConsoleColor]::Green } else { [System.ConsoleColor]::Gray }
                Write-Host ('    [{0}] {1}  {2}' -f $i, $marker, $modes[$i]) -ForegroundColor $color
            }
            Write-Host ''
            Write-MenuLine '  Switches (toggle A-T, then ENTER to proceed to inputs):' Yellow

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

            Write-Host ''
            if ($StatusMsg) { Write-Host "  $StatusMsg" -ForegroundColor Yellow }
        }

        # --- Step 1: Mode selection ---
        while ($selectedMode -lt 1 -or $selectedMode -gt 6) {
            Draw-Menu -Mode $selectedMode -ToggState $toggleState
            $raw = Read-Host '  Mode [1-6]'
            if ($raw -match '^[1-6]$') {
                $selectedMode = [int]$raw
                # Apply mode-specific toggle defaults
                switch ($selectedMode) {
                    2 { $toggleState['G'] = $false; $toggleState['I'] = $false }
                    3 { foreach ($k in $disabledToggles[3]) { $toggleState[$k] = $false } }
                    6 { foreach ($k in $disabledToggles[6]) { $toggleState[$k] = $false } }
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
            Draw-Menu -Mode $selectedMode -ToggState $toggleState -StatusMsg $statusMsg -ExtraDisabled $dynDisabled
            $statusMsg = ''

            if ($useRawKey) {
                Write-Host '  Press A-T to toggle, ENTER to continue: ' -NoNewline -ForegroundColor Cyan
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
                $raw = (Read-Host '  Toggle [A-T] or ENTER to continue').Trim().ToUpper()
                if ($raw -eq '') { break }
            }

            if ($raw.Length -eq 1 -and $toggleDefs.Contains($raw)) {
                $dynNow = Get-DynamicDisabled $toggleState
                if (($disabledToggles[$selectedMode] -contains $raw) -or ($dynNow -contains $raw)) {
                    $statusMsg = "[$raw] is not available in this configuration"
                }
                else {
                    $toggleState[$raw] = -not $toggleState[$raw]
                    # Reset any toggles that became disabled by this change
                    $dynAfter = Get-DynamicDisabled $toggleState
                    foreach ($x in $dynAfter) {
                        if ($toggleState[$x]) { $toggleState[$x] = $false }
                    }
                }
            }
            elseif ($raw.Length -gt 0) {
                $statusMsg = "Unknown key '$raw' — press A-T to toggle or ENTER to continue"
            }
        }

        # --- Step 3: String inputs (context-dependent) ---
        Clear-Host
        Write-MenuLine ('=' * 60) Cyan
        Write-MenuLine ("  Install-Exchange15 v{0} - Mode: {1}" -f $ScriptVersion, $modes[$selectedMode]) Cyan
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
            $displayDefault = if ($Default) { "[$Default]" } else { '' }
            $full = if ($displayDefault) { "  $Prompt $displayDefault" } else { "  $Prompt" }
            while ($true) {
                $val = Read-Host $full
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
        $defaultIso = Join-Path (Split-Path $ScriptFullName -Parent) 'ExchangeServerSE-x64.iso'
        if ($selectedMode -ne 6) {
            $cfg['SourcePath'] = Read-MenuInput -Prompt 'Exchange source (folder or .iso)' -Default $defaultIso -Required $true
        }
        $cfg['InstallPath'] = Read-MenuInput -Prompt 'Working/log folder' -Default 'C:\Install'

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
            $cfg['DAGName']          = Read-MenuInput -Prompt 'DAG name               (blank = no DAG join) [not tested]'
            $cfg['CopyServerConfig'] = Read-MenuInput -Prompt 'Copy config from server (FQDN, blank = none) [not tested]' -Validate $validateFQDN -ValidateMessage 'Not a valid FQDN (e.g. ex01.contoso.com)'
            $cfg['CertificatePath']  = Read-MenuInput -Prompt 'PFX certificate path   (blank = none)        [not tested]'
            $cfg['Namespace']        = Read-MenuInput -Prompt 'External namespace      (e.g. mail.contoso.com, blank = skip URL config)' -Validate $validateFQDN -ValidateMessage 'Not a valid FQDN (e.g. mail.contoso.com)'
            if ($cfg['Namespace']) {
                $cfg['DownloadDomain'] = Read-MenuInput -Prompt 'OWA download domain     (e.g. download.contoso.com, blank = skip CVE-2021-1730)' -Validate $validateFQDN -ValidateMessage 'Not a valid FQDN (e.g. download.contoso.com)'
            }
            if ((Read-MenuInput -Prompt 'Enable log cleanup task? [Y/N]' -Default 'N') -imatch '^[Yy]$') {
                $retDays = Read-MenuInput -Prompt 'Log retention days' -Default '30' -Required $true
                $cfg['LogRetentionDays'] = [int]$retDays
            } else {
                $cfg['LogRetentionDays'] = 0
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
        elseif ($selectedMode -eq 6) {
            $cfg['Namespace']        = Read-MenuInput -Prompt 'External namespace      (e.g. mail.contoso.com, blank = skip URL config)' -Validate $validateFQDN -ValidateMessage 'Not a valid FQDN (e.g. mail.contoso.com)'
            if ($cfg['Namespace']) {
                $cfg['DownloadDomain'] = Read-MenuInput -Prompt 'OWA download domain     (e.g. download.contoso.com, blank = skip CVE-2021-1730)' -Validate $validateFQDN -ValidateMessage 'Not a valid FQDN (e.g. download.contoso.com)'
            }
            $cfg['CertificatePath']  = Read-MenuInput -Prompt 'PFX certificate path   (blank = none)        [not tested]'
            $cfg['DAGName']          = Read-MenuInput -Prompt 'DAG name               (blank = no DAG join) [not tested]'
            if ((Read-MenuInput -Prompt 'Enable log cleanup task? [Y/N]' -Default 'N') -imatch '^[Yy]$') {
                $retDays = Read-MenuInput -Prompt 'Log retention days' -Default '30' -Required $true
                $cfg['LogRetentionDays'] = [int]$retDays
            } else {
                $cfg['LogRetentionDays'] = 0
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

        # --- Step 4: Summary + confirmation ---
        while ($true) {
            Clear-Host
            Write-MenuLine ('=' * 60) Cyan
            Write-MenuLine '  Summary' Cyan
            Write-MenuLine ('=' * 60) Cyan
            Write-Host ''
            Write-Host ('  Mode    : {0}' -f $modes[$selectedMode]) -ForegroundColor Green
            Write-Host ('  Source  : {0}' -f $cfg['SourcePath'])
            Write-Host ('  Install : {0}' -f $cfg['InstallPath'])
            if ($cfg['Organization']) { Write-Host ('  Org     : {0}' -f $cfg['Organization']) }
            # Active switches
            $finalDisabled  = @($disabledToggles[$selectedMode]) + (Get-DynamicDisabled $toggleState)
            $activeToggles = ($toggleDefs.Keys | Where-Object { $toggleState[$_] -and ($finalDisabled -notcontains $_) }) -join ', '
            if ($activeToggles) { Write-Host ('  Switches: {0}' -f $activeToggles) }
            Write-Host ''
            $confirm = Read-Host '  Start installation? [Y=yes / N=back to menu / Q=quit]'
            if ($confirm -imatch '^[Yy]') { return $cfg }
            if ($confirm -imatch '^[Qq]') { return $null }
            # N or anything else = restart from mode selection
            $selectedMode = 0
            while ($selectedMode -lt 1 -or $selectedMode -gt 6) {
                Draw-Menu -Mode $selectedMode -ToggState $toggleState
                $raw = Read-Host '  Mode [1-6]'
                if ($raw -match '^[1-6]$') { $selectedMode = [int]$raw }
            }
        }
    }

    ########################################
    # MAIN
    ########################################

    #Requires -Version 5.1

    # When compiled with PS2Exe, MyInvocation.MyCommand.Path is empty — fall back to the process image path
    $ScriptFullName = if ($MyInvocation.MyCommand.Path) {
        $MyInvocation.MyCommand.Path
    } else {
        [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    }
    # Detect PS2Exe compiled run: MyCommand.Path is empty; Write-Progress is not rendered visually
    $IsPS2Exe = -not $MyInvocation.MyCommand.Path
    $ScriptName = $ScriptFullName.Split("\")[-1]
    $ParameterString = $PSBoundParameters.getEnumerator() -join " "
    $OSVersionParts = (Get-CimInstance -ClassName Win32_OperatingSystem).Version.Split('.')
    $MajorOSVersion = '{0}.{1}' -f $OSVersionParts[0], $OSVersionParts[1]
    $MinorOSVersion = $OSVersionParts[2]
    $FullOSVersion  = '{0}.{1}' -f $MajorOSVersion, $MinorOSVersion

    $State = @{}
    $StateFile = "$InstallPath\$($env:computerName)_$($ScriptName)_state.xml"
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

    if ($State.Count -gt 0 -and -not $ParameterString) {
        $ParameterString = '[resuming from phase {0}]' -f $State['InstallPhase']
    }
    Write-Output "Script $ScriptFullName v$ScriptVersion called using $ParameterString"
    Write-Verbose "Using parameterSet $($PsCmdlet.ParameterSetName)"
    Write-Output ('Running on OS build {0}' -f $FullOSVersion)

    if (! $State.Count) {
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
            $InstallPath     = if ($menuResult['InstallPath']) { $menuResult['InstallPath'] } else { 'C:\Install' }
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
            $DisableSSL3         = [switch]($menuResult['DisableSSL3'])
            $DisableRC4          = [switch]($menuResult['DisableRC4'])
            $EnableECC           = [switch]($menuResult['EnableECC'])
            $NoCBC               = [switch]($menuResult['NoCBC'])
            $EnableAMSI          = [switch]($menuResult['EnableAMSI'])
            $EnableTLS12         = [switch]($menuResult['EnableTLS12'])
            $EnableTLS13         = [switch]($menuResult['EnableTLS13'])
            $DoNotEnableEP       = [switch]($menuResult['DoNotEnableEP'])
            $DiagnosticData      = [switch]($menuResult['DiagnosticData'])
            $Lock                = [switch]($menuResult['Lock'])
            $SkipRolesCheck      = [switch]($menuResult['SkipRolesCheck'])
            $PreflightOnly       = [switch]($menuResult['PreflightOnly'])
            $NoCheckpoint        = [switch]($menuResult['NoCheckpoint'])
            $SkipHealthCheck         = [switch]($menuResult['SkipHealthCheck'])
            $NoNet481                = [switch]($menuResult['NoNet481'])
            $InstallWindowsUpdates   = [switch]($menuResult['InstallWindowsUpdates'])
            $RunEOMT             = [switch]($menuResult['RunEOMT'])
            $WaitForADSync       = [switch]($menuResult['WaitForADSync'])
            $InstallEdge         = [switch]($mode -eq 2)
            $Recover             = [switch]($mode -eq 5)
            $StandaloneOptimize  = [switch]($mode -eq 6)
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
            $StateFile = "$InstallPath\$($env:computerName)_$($ScriptName)_state.xml"
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
            $InstallPath  = if (Get-CfgValue 'InstallPath' $InstallPath) { Get-CfgValue 'InstallPath' $InstallPath } else { 'C:\Install' }

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
            $DownloadDomain       = Get-CfgValue 'DownloadDomain' $DownloadDomain
            $RunEOMT              = [switch](Get-CfgValue 'RunEOMT'              ([bool]$RunEOMT))
            $WaitForADSync        = [switch](Get-CfgValue 'WaitForADSync'        ([bool]$WaitForADSync))
            $LogRetentionDays     = Get-CfgValue 'LogRetentionDays' $LogRetentionDays
            $RelaySubnets         = Get-CfgValue 'RelaySubnets'         $RelaySubnets
            $ExternalRelaySubnets = Get-CfgValue 'ExternalRelaySubnets' $ExternalRelaySubnets

            # Recalculate state file path with potentially overridden InstallPath
            $StateFile = "$InstallPath\$($env:computerName)_$($ScriptName)_state.xml"
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
        $State["OrganizationName"] = $Organization
        $State["AdminAccount"] = if ($Credentials) { $Credentials.UserName } else { $null }
        $State["AdminPassword"] = if ($Credentials) { ($Credentials.Password | ConvertFrom-SecureString -ErrorAction SilentlyContinue) } else { $null }
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
        $State["IncludeFixes"] = $IncludeFixes
        $State["NoSetup"] = $NoSetup
        $State["Recover"] = $Recover
        $State["Upgrade"] = $false
        $State["Install481"] = $False
        $State["VCRedist2012"] = $False
        $State["VCRedist2013"] = $False
        $State["DisableSSL3"] = $DisableSSL3
        $State["DisableRC4"] = $DisableRC4
        $State["EnableECC"] = $EnableECC
        $State["EnableCBC"] = -not $NoCBC
        $State["EnableTLS12"] = $EnableTLS12
        $State["EnableTLS13"] = $EnableTLS13
        if ($State["EnableTLS13"] -and -not $State["EnableTLS12"]) {
            Write-MyWarning '-EnableTLS13 requires -EnableTLS12; automatically enabling TLS 1.2 enforcement'
            $State["EnableTLS12"] = $true
        }
        $State["DoNotEnableEP"] = $DoNotEnableEP
        $State["DoNotEnableEP_FEEWS"] = $DoNotEnableEP_FEEWS
        $State["SkipRolesCheck"] = $SkipRolesCheck
        $State["SCP"] = $SCP
        $State["DiagnosticData"] = $DiagnosticData
        $State["Lock"] = $Lock
        $State["EdgeDNSSuffix"] = $EdgeDNSSuffix
        $State["InstallPath"]  = $InstallPath
        $State["ReportsPath"]  = Join-Path $InstallPath 'reports'
        if (-not (Test-Path $State["ReportsPath"])) { New-Item -Path $State["ReportsPath"] -ItemType Directory -Force | Out-Null }
        $State["TranscriptFile"] = Join-Path $State["ReportsPath"] ('{0}_{1}_{2}.log' -f $env:computerName, $ScriptName, (Get-Date -Format 'yyyyMMddHHmmss'))
        $State["PreflightOnly"] = $PreflightOnly
        $State["CopyServerConfig"] = $CopyServerConfig
        $State["CertificatePath"] = $CertificatePath
        $State["CertificatePassword"] = $null
        $State["DAGName"] = $DAGName
        $State["SkipHealthCheck"] = $SkipHealthCheck
        $State["NoCheckpoint"] = $NoCheckpoint
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
        $State["DownloadDomain"] = $DownloadDomain
        $State["RunEOMT"]          = [bool]$RunEOMT
        $State["WaitForADSync"]    = [bool]$WaitForADSync
        $State["LogRetentionDays"] = $LogRetentionDays
        $State["RelaySubnets"]         = $RelaySubnets
        $State["ExternalRelaySubnets"] = $ExternalRelaySubnets
        $State["StandaloneOptimize"] = [bool]$StandaloneOptimize
        $State["SkipInstallReport"] = [bool]$SkipInstallReport

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
    else {
        $MAX_PHASE = 7
    }

    $runMode = if ($State['ConfigDriven']) { 'Autopilot (fully automated)' } else { 'Copilot (interactive)' }
    Write-MyOutput ('Mode: {0}' -f $runMode)

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
                    # VC++ 2013 x64 12.0.40664 (KB4538461, Jan 2020 update) — Microsoft Download Center
                    Install-MyPackage "" "Visual C++ 2013 Redistributable" "vcredist_x64_2013.exe" "https://download.microsoft.com/download/C/C/2/CC2DF5F8-4454-44B4-802D-5EA68D086676/vcredist_x64.exe" ("/install", "/quiet", "/norestart")
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

                Start-DisableMSExchangeAutodiscoverAppPoolJob

                Install-Exchange15_

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
                    'IPv4 over IPv6', 'NetBIOS on NICs',
                    'MaxConcurrentAPI', 'Disk allocation', 'Scheduled tasks', 'Server Manager',
                    'CRL timeout', 'TLS / Schannel', 'Root CA auto-update', 'Exchange module + search tuning',
                    'Security hardening', 'Org/Transport optimizations', 'IANA timezone mapping',
                    'SSL offloading', 'Extended Protection', 'MRS Proxy', 'MAPI encryption',
                    'Exchange SU', 'Server config import', 'Certificate', 'HSTS header', 'EOMT'
                )
                $p5Total = $p5Steps.Count; $p5Step = 0
                $p5Sw = [Diagnostics.Stopwatch]::new(); $p5LastDesc = $null
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
                Step-P5 'Power plan';                   Enable-HighPerformancePowerPlan
                Step-P5 'NIC power management';         Disable-NICPowerManagement
                Step-P5 'Page file';                    Set-Pagefile
                Step-P5 'TCP settings';                 Set-TCPSettings
                Step-P5 'SMBv1';                        Disable-SMBv1
                Step-P5 'Windows Search service';       Disable-WindowsSearchService
                Step-P5 'Unnecessary services';         Disable-UnnecessaryServices
                Step-P5 'Shutdown Event Tracker';       Disable-ShutdownEventTracker
                Step-P5 'WDigest caching';              Disable-WDigestCredentialCaching
                Step-P5 'HTTP/2';                       Disable-HTTP2
                Step-P5 'TCP offload';                  Disable-TCPOffload
                Step-P5 'Credential Guard';             Disable-CredentialGuard
                Step-P5 'LM compatibility level';       Set-LmCompatibilityLevel
                Step-P5 'LSA Protection (RunAsPPL)';   Enable-LSAProtection
                Step-P5 'RSS / NIC queues';             Enable-RSSOnAllNICs
                Step-P5 'IPv4 over IPv6';               Set-IPv4OverIPv6Preference
                Step-P5 'NetBIOS on NICs';              Disable-NetBIOSOnAllNICs
                Step-P5 'MaxConcurrentAPI';             Set-MaxConcurrentAPI
                Step-P5 'Disk allocation unit';         Test-DiskAllocationUnitSize
                Step-P5 'Scheduled tasks';              Disable-UnnecessaryScheduledTasks
                Step-P5 'CRL check timeout';            Set-CRLCheckTimeout
                Step-P5 'TLS / Schannel'
                if ( $State["DisableSSL3"]) {
                    Disable-SSL3
                }
                if ( $State["DisableRC4"]) {
                    Disable-RC4
                }
                Set-TLSSettings -TLS12 $State["EnableTLS12"] -TLS13 $State["EnableTLS13"]

                Step-P5 'Root CA auto-update';     Enable-RootCertificateAutoUpdate

                Step-P5 'Exchange module + search tuning'
                Import-ExchangeModule
                Set-CtsProcessorAffinityPercentage
                Enable-SerializedDataSigning
                Set-NodeRunnerMemoryLimit
                Enable-MAPIFrontEndServerGC

                if ( $State["EnableECC"]) {
                    Enable-ECC
                }
                if ( $State["EnableCBC"]) {
                    Enable-CBC
                }
                if ( $State["EnableAMSI"]) {
                    # HealthChecker always checks for the SettingOverride (Get-SettingOverride on
                    # AmsiRequestBodyScanning), regardless of Exchange version defaults.
                    # Apply the override for all versions so HC reports the correct state.
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
                    # ECC/CBC/AMSI above may have restarted W3SVC, killing the implicit-remoting session.
                    if ($State['EnableECC'] -or $State['EnableCBC'] -or $State['EnableAMSI']) {
                        Reconnect-ExchangeSession
                    }
                    Invoke-ExchangeOptimizations
                }

                # IANA timezone mapping check (Exchange 2019 CU14+)
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
                    Disable-MRSProxy
                    Step-P5 'MAPI encryption'
                    Set-MAPIEncryptionRequired
                }

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

                # Import server configuration from source server
                Step-P5 'Server configuration import'
                if ($State['CopyServerConfig'] -and $State['ServerConfigExportPath']) {
                    Import-ServerConfig
                }

                # Import PFX certificate
                Step-P5 'PFX certificate import'
                if ($State['CertificatePath']) {
                    Import-ExchangeCertificateFromPFX
                }

                # HSTS header — only when a certificate was imported (avoid browser lockout with self-signed cert)
                Step-P5 'HSTS header'
                if ($State['CertificatePath']) {
                    Set-HSTSHeader
                }
                else {
                    Write-MyVerbose 'No CertificatePath specified — skipping HSTS (requires valid certificate to avoid browser lockout)'
                }

                # EOMT — optional CVE mitigation tool
                Step-P5 'EOMT'
                Invoke-EOMT
                if ($p5LastDesc) { Write-MyVerbose ('{0} took {1:F1}s' -f $p5LastDesc, $p5Sw.Elapsed.TotalSeconds) }

                Write-PhaseProgress -Id 1 -Activity 'Phase 5 of 6: Post-configuration' -Completed
                Write-PhaseProgress -Activity 'Exchange Installation' -Completed
            }

            6 {
                Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 7: Finalizing' -PercentComplete 0
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

                # Load Exchange PS module once for all Phase 6 operations
                if (-not $State['InstallEdge']) {
                    Import-ExchangeModule
                }

                # Install antispam agents (Mailbox role only)
                if ($State['InstallMailbox'] -and -not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 7: Antispam agents' -PercentComplete 8
                    Install-AntispamAgents
                }

                # Set Virtual Directory URLs
                if ($State['Namespace'] -and -not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 7: Virtual Directory URLs' -PercentComplete 15
                    Set-VirtualDirectoryURLs
                }

                # Join Database Availability Group
                if ($State['DAGName']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 7: Joining DAG' -PercentComplete 30
                    Join-DAG
                }

                # DAG replication health check (F8)
                if ($State['DAGName'] -and -not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 7: DAG replication health' -PercentComplete 33
                    Test-DAGReplicationHealth
                }

                # Add server to existing Send Connectors
                if (-not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 7: Send Connectors' -PercentComplete 35
                    Add-ServerToSendConnectors
                }

                # Server Manager stays disabled permanently on Exchange servers (set machine-wide in Phase 5)

                if ( !($State['InstallEdge'])) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 7: IIS health check' -PercentComplete 60
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
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 7: DB path check' -PercentComplete 73
                    Test-DBLogPathSeparation
                }

                # Auth Certificate health check
                if (-not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 7: Auth Certificate check' -PercentComplete 75
                    Test-AuthCertificate
                }

                # VSS writers, EEMS, Modern Auth checks (F9, F10, F11)
                if (-not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 7: VSS writers' -PercentComplete 76
                    Test-VSSWriters
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 7: EEMS status' -PercentComplete 77
                    Test-EEMSStatus
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 7: Modern Auth' -PercentComplete 77
                    Get-ModernAuthReport
                }

                # Anonymous relay connector
                if (($State['RelaySubnets'] -or $State['ExternalRelaySubnets']) -and -not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 7: Anonymous relay connector' -PercentComplete 78
                    New-AnonymousRelayConnector
                }

                # Exchange log cleanup scheduled task
                Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 7: Log cleanup task' -PercentComplete 76
                Register-ExchangeLogCleanup

                # RBAC role group membership report
                if (-not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 6 of 7: RBAC report' -PercentComplete 78
                    Get-RBACReport
                }

                Enable-UAC
                Enable-IEESC

                Write-PhaseProgress -Activity 'Exchange Installation' -Completed
            }

            7 {
                # Phase 7 runs after the Phase 6 reboot so that:
                # - Exchange computer account group membership is in the process token
                # - All services have started cleanly from scratch
                # - HealthChecker results reflect the fully-configured server
                Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 7 of 7: HealthChecker &amp; Report' -PercentComplete 0

                # Reload Exchange PS module after reboot
                if (-not $State['InstallEdge']) {
                    Import-ExchangeModule
                }

                # Run CSS-Exchange HealthChecker
                Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 7 of 7: HealthChecker' -PercentComplete 20
                if (-not $State['SkipHealthCheck']) {
                    Invoke-HealthChecker
                }

                # Installation Report
                if (-not $State['SkipInstallReport'] -and -not $State['InstallEdge']) {
                    Write-PhaseProgress -Activity 'Exchange Installation' -Status 'Phase 7 of 7: Installation Report' -PercentComplete 70
                    # B16: wrap in try/catch so a crash inside New-InstallationReport does not
                    # propagate to the global trap { break } and kill the script before the
                    # "We're good to go" message and phase-end reboot logic run.
                    try { New-InstallationReport } catch { Write-MyWarning ('Installation Report failed: {0}' -f $_.Exception.Message) }
                }

                Write-PhaseProgress -Activity 'Exchange Installation' -Completed
                Write-MyOutput "Setup finished - We're good to go."
            }

            default {
                Write-MyError "Unknown phase ($($State["InstallPhase"]))"
                exit $ERR_UNEXPTECTEDPHASE
            }
        }
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

