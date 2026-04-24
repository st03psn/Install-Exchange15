#
# Install-Exchange15.ps1 — Configuration file example
#
# Usage:  .\Install-Exchange15.ps1 -ConfigFile .\deploy-mbx01.psd1
#
# The interactive menu is automatically skipped when -ConfigFile is used.
# Parameters specified on the command line take precedence over the config file.
# All keys are optional — omitted keys fall back to the script's built-in defaults.
#
@{
    # -------------------------------------------------------------------------
    # Source & paths
    # -------------------------------------------------------------------------

    # Exchange setup source: folder with setup.exe OR path to .iso file
    # The ISO is auto-mounted and unblocked (Zone.Identifier) before mounting.
    SourcePath  = 'C:\Install\Exchange-Server-Install.iso'

    # Working directory: state file, logs, and downloaded packages are stored here
    InstallPath = 'C:\Install'

    # -------------------------------------------------------------------------
    # Installation mode  (default: Mailbox if none specified)
    # -------------------------------------------------------------------------
    # InstallEdge                = $false   # Edge Transport role
    # Recover                    = $false   # Recovery installation
    # NoSetup                    = $false   # Skip Exchange setup (post-config only)
    # InstallRecipientManagement = $false   # Recipient Management Tools only
    # InstallManagementTools     = $false   # Exchange Management Tools only

    # -------------------------------------------------------------------------
    # Exchange configuration
    # -------------------------------------------------------------------------

    # Exchange organization name (blank = use existing org, required for new deployments)
    Organization = 'Contoso'

    # Mailbox database (all optional — Exchange defaults are used if omitted)
    MDBName    = 'MDB01'
    # MDBDBPath  = 'D:\MailboxData\MDB01\DB'    # custom DB path
    # MDBLogPath = 'D:\MailboxData\MDB01\Log'   # custom log path

    # Autodiscover SCP URL (blank = keep existing, '-' = remove SCP)
    # SCP = 'https://autodiscover.contoso.com/autodiscover/autodiscover.xml'

    # Custom Exchange install directory (blank = C:\Program Files\Microsoft\Exchange Server\V15)
    # TargetPath = 'D:\Exchange'

    # -------------------------------------------------------------------------
    # AutoPilot & credentials
    # -------------------------------------------------------------------------

    # Autopilot: automatic reboot + resume after each phase (fully unattended).
    # Set to $false (or omit) to use Copilot (interactive) mode.
    # Credentials must be entered interactively on first run (stored encrypted in state file).
    Autopilot = $true

    # -------------------------------------------------------------------------
    # Advanced Configuration  (v5.95+; ~55 hardening / tuning / policy knobs)
    # -------------------------------------------------------------------------
    #
    # The nested 'AdvancedFeatures' block replaces the old flat top-level keys
    # (DisableSSL3, EnableECC, …). Omitted entries keep their catalog default
    # — equivalent to current v5.x behaviour. See Get-AdvancedFeatureCatalog
    # in Install-Exchange15.ps1 for the full list.
    #
    # Precedence: AdvancedFeatures nested block > legacy top-level key
    #             > -<Name> cmdline switch > catalog default.

    AdvancedFeatures = @{
        # --- Security / TLS ---
        DisableSSL3         = $true    # POODLE / CVE-2014-3566
        DisableRC4          = $true    # Deprecated stream cipher
        EnableECC           = $true    # Prefer ECC key exchange
        NoCBC               = $false   # Disable CBC ciphers (not recommended)
        EnableAMSI          = $true
        EnableTLS12         = $true    # Disables TLS 1.0/1.1 + .NET StrongCrypto
        EnableTLS13         = $true    # WS2022+; ignored on older OS
        # DoNotEnableEP     = $false   # Opt-out of Extended Protection

        # --- Security / Hardening ---
        SMBv1Disable        = $true
        NetBIOSDisable      = $true
        LLMNRDisable        = $true
        MDNSDisable         = $true
        WDigestDisable      = $true
        LSAProtection       = $true
        LmCompat5           = $true
        SerializedDataSig   = $true
        ShutdownTrackerOff  = $true
        HTTP2Disable        = $true
        CredentialGuardOff  = $true
        UnnecessaryServices = $true
        WindowsSearchOff    = $true
        CRLTimeout          = $true
        RootCAAutoUpdate    = $true
        SMTPBannerHarden    = $true

        # --- Performance / Tuning ---
        MaxConcurrentAPI    = $true
        DiskAllocHint       = $true
        CtsProcAffinity     = $true
        NodeRunnerMemLimit  = $true
        MapiFeGC            = $true
        NICPowerMgmtOff     = $true
        RSSEnable           = $true
        TCPTuning           = $true
        TCPOffloadOff       = $true
        IPv4OverIPv6Off     = $true

        # --- Exchange Org Policy ---
        ModernAuth          = $true
        OWASessionTimeout6h = $true
        DisableTelemetry    = $true
        MapiHttp            = $true
        MaxMessageSize150MB = $true
        MessageExpiration7d = $true
        HtmlNDR             = $true
        # ShadowRedundancy  = $false  # DAG-only
        SafetyNet2d         = $true

        # --- Post-Config / Integration ---
        MECA                = $true
        AntispamAgents      = $true
        SSLOffloading       = $true
        MRSProxy            = $true
        IANATimezone        = $true
        # AnonymousRelay    = $true   # auto-enabled when RelaySubnets set
        # SkipHealthCheck   = $false
        RBACReport          = $true
        # RunEOMT           = $false  # legacy CUs only

        # --- Install-Flow / Debug (defaults $false) ---
        # DiagnosticData    = $false
        # Lock              = $false
        # SkipRolesCheck    = $false
        # NoCheckpoint      = $false
        # NoNet481          = $false
        # WaitForADSync     = $false
    }

    # -------------------------------------------------------------------------
    # Updates
    # -------------------------------------------------------------------------

    IncludeFixes          = $true   # Install latest Exchange Security Update (SU) after setup
    InstallWindowsUpdates = $true   # Install Windows security/critical updates in Phase 1

    # -------------------------------------------------------------------------
    # Optional features
    # -------------------------------------------------------------------------

    # DAGName          = 'DAG01'                    # Join this DAG after install
    # CopyServerConfig = 'exch01.contoso.com'       # Copy virtual dirs + connectors from this server
    # CertificatePath  = 'C:\Certs\exchange.pfx'    # Import PFX certificate (password prompted)

    # -------------------------------------------------------------------------
    # Behaviour flags
    # -------------------------------------------------------------------------

    PreflightOnly     = $false   # $true to generate pre-flight report and exit
    SkipInstallReport = $false   # $true to suppress HTML/PDF installation report at Phase 6
    SkipSetupAssist   = $false   # $true to skip CSS-Exchange SetupAssist on Phase 4 failure
    # SkipHealthCheck / NoCheckpoint / DiagnosticData moved to AdvancedFeatures above.

    # -------------------------------------------------------------------------
    # v5.82 / v5.84 — Word installation document (F22)
    # -------------------------------------------------------------------------

    # $true to skip Word (.docx) installation document generation after Phase 6
    # NoWordDoc = $false

    # Redact RFC1918 IPs, certificate thumbprints, and passwords in the document
    # (useful when sharing the document with external parties)
    # CustomerDocument = $false

    # Document language: 'DE' (default) or 'EN'
    # Language = 'DE'

    # Scope of the generated document (v5.84):
    #   All   — org-wide settings + all Exchange servers + local details (default)
    #   Org   — org-wide chapter only (no per-server hardware / VDir queries)
    #   Local — per-server sections only (no org-wide chapter)
    # DocumentScope = 'All'

    # Limit per-server documentation to specific server names (v5.84).
    # Local server is always included. Applies when DocumentScope is All or Local.
    # IncludeServers = @('EX01', 'EX02')

    # F24 (v1.0): path to a custom DOCX template for the installation document.
    # When supplied, the cover page and header come from the template; the chapter
    # body is generated by the script and injected into {{document_body}}.
    # Use tools\Build-InstallationTemplate.ps1 to generate the starter templates in
    # templates\Exchange-installation-document-{DE,EN}.docx.
    # TemplatePath = 'C:\Deploy\my-company-template.docx'

    # DoNotEnableEP / NoNet481 / SkipRolesCheck / Lock / RunEOMT / WaitForADSync
    # moved to AdvancedFeatures above.

    # -------------------------------------------------------------------------
    # Relay connectors (v5.2)
    # -------------------------------------------------------------------------

    # Anonymous internal relay: accepted domains only (no external relay right)
    # Source IPs resolved via SID S-1-5-7 — language-independent (DE/EN/FR/...)
    # RelaySubnets = @('192.168.10.0/24', '10.0.0.5')

    # Anonymous external relay: any recipient (Ms-Exch-SMTP-Accept-Any-Recipient)
    # SECURITY: restrict to trusted send systems (scanners, printers) only
    # ExternalRelaySubnets = @('10.0.1.100')

    # -------------------------------------------------------------------------
    # v5.2 — Log cleanup & namespace
    # -------------------------------------------------------------------------

    # Register daily scheduled task (02:00, SYSTEM) to delete logs older than N days
    # Cleans: IIS logs, Exchange transport logs, message tracking logs
    # LogRetentionDays = 30

    # External namespace for Virtual Directory URL configuration (Phase 6)
    # Namespace = 'mail.contoso.com'

    # OWA Download Domain — separate FQDN for attachment downloads (CVE-2021-1730 mitigation)
    # Must differ from Namespace; requires matching DNS record and certificate coverage
    # DownloadDomain = 'download.contoso.com'

    # Path to PFX certificate — also enables HSTS on OWA/ECP when set
    # CertificatePath = 'C:\Certs\exchange.pfx'
}
