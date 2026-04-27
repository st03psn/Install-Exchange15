#
# EXpress.ps1 — Configuration file example
#
# Usage:  .\EXpress.ps1 -ConfigFile .\deploy-mbx01.psd1
#
# The interactive menu is automatically skipped when -ConfigFile is used.
# Parameters specified on the command line take precedence over the config file.
# All keys are optional — omitted keys fall back to the script's built-in defaults.
#
# Fully unattended setup checklist:
#   - Using -ConfigFile already implies Autopilot (auto-reboot + resume).
#     Add Autopilot = $false only if you want Copilot prompts with a config file.
#   - AdminUser / AdminPassword set in this file      (or interactive prompt on first run)
#   - InstallWindowsUpdates = $true                   (Windows Updates installed in Phase 1)
#   - AutoApproveWindowsUpdates auto-defaults to $true for ConfigFile runs — no extra setup
#     needed. Set explicitly = $false to keep the prompt and skip auto-install.
#
@{
    # -------------------------------------------------------------------------
    # Source & paths
    # -------------------------------------------------------------------------

    # Exchange setup source: folder with setup.exe OR path to .iso file.
    # The ISO is auto-mounted and unblocked (Zone.Identifier) before mounting.
    # Convention: place install media under <InstallPath>\Sources\ — Copilot menu
    # auto-detects sources\ExchangeServerSE-x64.iso under InstallPath by default.
    SourcePath  = 'C:\Install\Sources\ExchangeServerSE-x64.iso'

    # Working directory: state file, install log, downloaded packages, reports/, Debug/.
    InstallPath = 'C:\Install'

    # -------------------------------------------------------------------------
    # Installation mode  (default: Mailbox if none specified)
    # -------------------------------------------------------------------------
    # InstallEdge                = $false   # Install Edge Transport role instead of Mailbox
    # EdgeDNSSuffix              = ''       # DNS suffix for Edge Transport (blank = auto-detect)
    # Recover                    = $false   # Recover-Server install (rebuild from AD object)
    # NoSetup                    = $false   # Skip Exchange setup entirely (post-config only)
    # InstallRecipientManagement = $false   # Install Recipient Management Tools only (no server role)
    # InstallManagementTools     = $false   # Install Exchange Management Tools only (no server role)

    # -------------------------------------------------------------------------
    # Exchange configuration
    # -------------------------------------------------------------------------

    # Exchange organization name (blank = use existing org, required for new deployments).
    Organization = 'Contoso'

    # Mailbox database name (blank = Exchange-generated default like 'Mailbox Database 1234567890').
    MDBName    = 'MDB01'
    # Custom mailbox database file path (blank = Exchange default under V15\Mailbox).
    # MDBDBPath  = 'D:\MailboxData\MDB01\DB'
    # Custom mailbox database log path (blank = same folder as DB).
    # MDBLogPath = 'D:\MailboxData\MDB01\Log'

    # Autodiscover SCP URL (blank = keep existing AD value, '-' = remove SCP entirely).
    # Use this to point internal clients at a load-balanced namespace instead of the server FQDN.
    # SCP = 'https://autodiscover.contoso.com/autodiscover/autodiscover.xml'

    # Custom Exchange install directory (blank = C:\Program Files\Microsoft\Exchange Server\V15).
    # TargetPath = 'D:\Exchange'

    # -------------------------------------------------------------------------
    # Credentials  (KEEP NEAR TOP — required for unattended setup)
    # -------------------------------------------------------------------------
    # Autopilot (auto-reboot + resume) is the default when using -ConfigFile.
    # To keep Copilot prompts while still loading from a config file, add:
    #   Autopilot = $false

    # ###################################################################
    # SECURITY: plain-text secrets below are ONLY for zero-touch pipelines.
    # Every run logs a SECURITY WARNING when these are read from the config.
    # AUTO-SCRUB: After the FIRST successful credential validation, EXpress
    # rewrites this file and replaces AdminPassword / CertificatePassword
    # values with empty strings (with a "scrubbed by EXpress on YYYY-MM-DD"
    # comment). The DPAPI-encrypted copy in the state file (user+machine
    # bound) is what subsequent phases read. You do NOT need to delete the
    # config file by hand — but its directory permissions still matter.
    # Without these keys, EXpress prompts interactively on first run.
    # ###################################################################

    # Domain admin account used for AD prep + Exchange setup (DOMAIN\username format).
    # AdminUser     = 'CONTOSO\svc-exchange-install'
    # Plain-text password matching AdminUser.
    # AdminPassword = 'P@ssw0rd!'

    # Path to PFX certificate to import + assign to IIS/SMTP services (Phase 5).
    # Without CertificatePassword set, the password is prompted interactively at
    # script startup — that breaks Autopilot. Provide CertificatePassword for
    # unattended runs.
    # CertificatePath     = 'C:\Certs\exchange.pfx'
    # CertificatePassword = 'pfxP@ss!'

    # -------------------------------------------------------------------------
    # Advanced Configuration  (~55 hardening / tuning / policy knobs)
    # -------------------------------------------------------------------------
    #
    # Omitted entries keep their catalog default. See Get-AdvancedFeatureCatalog
    # in EXpress.ps1 for the full list.

    AdvancedFeatures = @{
        # --- Security / TLS ---
        DisableSSL3         = $true    # Disable SSL 3.0 (POODLE / CVE-2014-3566)
        DisableRC4          = $true    # Disable deprecated RC4 stream cipher
        EnableECC           = $true    # Prefer ECC key exchange suites
        NoCBC               = $false   # Disable CBC ciphers (BREAKS Outlook Anywhere — keep $false)
        EnableAMSI          = $true    # Enable AMSI integration for Exchange transport (Defender script scanning)
        EnableTLS12         = $true    # Force TLS 1.2 + .NET StrongCrypto; disables TLS 1.0/1.1
        EnableTLS13         = $true    # Enable TLS 1.3 on WS2022+ (silently ignored on older OS)
        # DoNotEnableEP     = $false   # Opt-out of Extended Protection (NOT recommended; default enables EP)

        # --- Security / Hardening ---
        SMBv1Disable        = $true    # Remove SMBv1 client/server (WannaCry / EternalBlue mitigation)
        NetBIOSDisable      = $true    # Disable NetBIOS over TCP/IP on all NICs (reduces broadcast attack surface)
        LLMNRDisable        = $true    # Disable LLMNR multicast name resolution (responder.py mitigation)
        MDNSDisable         = $true    # Disable mDNS (port 5353) — same family as LLMNR
        WDigestDisable      = $true    # Block plaintext credential caching in WDigest (Mimikatz mitigation)
        LSAProtection       = $true    # Enable RunAsPPL — protects LSASS process from credential dumping
        LmCompat5           = $true    # LmCompatibilityLevel=5 (NTLMv2 only, refuse LM/NTLMv1)
        SerializedDataSig   = $true    # Force signed serialized data in Exchange (.NET deserialization hardening)
        ShutdownTrackerOff  = $true    # Suppress the "Why are you shutting down?" prompt on Server SKUs
        HSTS                = $true    # HSTS header on OWA/ECP (requires CertificatePath set)
        MAPIEncryption      = $true    # Force encrypted MAPI/RPC client connections
        HTTP2Disable        = $true    # Disable HTTP/2 in IIS (avoids known Exchange + HTTP/2 stability issues)
        CredentialGuardOff  = $true    # Turn off Credential Guard (blocks legacy SSPs Exchange depends on)
        UnnecessaryServices = $true    # Disable a curated list of services not needed by Exchange/AD
        WindowsSearchOff    = $true    # Disable Windows Search service (Exchange uses its own indexing engine)
        CRLTimeout          = $true    # Shorten CRL retrieval timeouts (faster cert chain validation under network failure)
        RootCAAutoUpdate    = $true    # Enable automatic Root CA program updates from Microsoft
        SMTPBannerHarden    = $true    # Replace SMTP receive-connector banner with non-disclosing string

        # --- Performance / Tuning ---
        MaxConcurrentAPI    = $true    # Raise MaxConcurrentApi to 150 (NTLM auth bottleneck for OWA/EWS)
        DiskAllocHint       = $true    # Set 64K NTFS allocation-unit hint on Exchange data volumes
        CtsProcAffinity     = $true    # Pin Content Transformation Service to specific CPU cores
        NodeRunnerMemLimit  = $true    # Cap noderunner.exe (Search Foundation) RAM growth
        MapiFeGC            = $true    # Tune .NET garbage collection for MapiFrontEndAppPool (latency)
        NICPowerMgmtOff     = $true    # Disable "Allow the computer to turn off this device" on all NICs
        RSSEnable           = $true    # Enable Receive-Side Scaling on all NICs (multi-core network throughput)
        TCPTuning           = $true    # Apply Exchange-recommended TCP autotuning + congestion-provider settings
        TCPOffloadOff       = $true    # Disable TCP Chimney/RSC offload (recommended by Microsoft for Exchange)
        IPv4OverIPv6Off     = $true    # Prefer IPv6 over IPv4 (default Windows order; reverses misconfigurations)

        # --- Exchange Org Policy (ORG-WIDE!) ---
        # ###################################################################
        # WARNING: These settings call Set-OrganizationConfig / Set-TransportConfig
        # which apply ORG-WIDE — they affect every Exchange server in the
        # organisation, not just the server being installed.
        #
        # Behaviour for existing organisations:
        #   The script auto-detects an existing org (ADSI probe + Initialize-Exchange)
        #   and FLIPS THE DEFAULT to $false for all settings in this block — so a
        #   second-server install never silently overwrites org-wide policy decisions
        #   the admin already made.
        #   Setting any value explicitly here always wins (including $true to force
        #   the rewrite, e.g. when this install is the deliberate policy update).
        #
        # All entries below stay commented by default. Enable individually only when
        # this install is meant to (re)define the org-wide policy.
        # ###################################################################
        # ModernAuth          = $true   # OAuth/Modern Authentication org-wide (Set-OrganizationConfig)
        # OWASessionTimeout6h = $true   # OWA private-computer session timeout 6h (org-wide OAuth)
        # DisableTelemetry    = $true   # CustomerFeedbackEnabled=$false (org-wide)
        # MapiHttp            = $true   # MAPI/HTTP enabled at org level (default for SE; legacy CUs)
        # MaxMessageSize150MB = $true   # Bump MaxSend/MaxReceive/Connector limits to 150MB (org-wide)
        # MessageExpiration7d = $true   # Transport message expiration 7 days (org-wide)
        # HtmlNDR             = $true   # HTML-formatted NDRs (org-wide)
        # ShadowRedundancy    = $false  # Shadow Redundancy on transport (DAG-only; auto-skipped without DAG)
        # SafetyNet2d         = $true   # SafetyNetHoldTime 2 days (org-wide)

        # --- Post-Config / Integration ---
        MECA                = $true    # Register CSS-Exchange MonitorExchangeAuthCertificate scheduled renewal task
        AntispamAgents      = $true    # Install built-in antispam agents on Mailbox role (no Edge in topology)
        SSLOffloading       = $true    # Configure OWA/ECP/EWS for SSL offloading at the load balancer
        MRSProxy            = $true    # Enable MRSProxy on EWS (cross-org / cross-site mailbox moves)
        IANATimezone        = $true    # Configure IANA ↔ Windows TZ mapping (skipped for existing orgs — one-way change)
        # AnonymousRelay is configured in the "Relay connectors" section below — not here.
        AccessNamespaceMail = $true    # Add Accepted Domain + EAP from MailDomain. Skipped automatically
                                       # when org already exists (only runs if EXpress created the org this run)
                                       # or when Namespace is blank.
        # SkipHealthCheck     = $false # Skip CSS-Exchange HealthChecker run at end of Phase 6
        RBACReport          = $true    # Generate RBAC permissions HTML report alongside installation report
        # RunEOMT             = $false # Run Exchange On-premises Mitigation Tool (legacy CUs / ProxyShell era only)

        # --- Install-Flow / Debug (defaults $false) ---
        # AutoApproveWindowsUpdates = $false  # Autopilot only: $true=install all WUs silently, $false=skip with warning.
                                              # Copilot ignores this (interactive Y/N/S prompt per update).
                                              # ConfigFile runs default to $true unless explicitly $false here.
        # DiagnosticData    = $false  # Enable extra diagnostic dumps in install log (verbose error context)
        # Lock              = $false  # Lock console after Autopilot reboot until next phase starts
        # SkipRolesCheck    = $false  # Bypass conflicting-roles check (lab/test scenarios only)
        # NoCheckpoint      = $false  # Disable Exchange setup recovery checkpointing (forces full re-run on failure)
    }

    # -------------------------------------------------------------------------
    # Updates
    # -------------------------------------------------------------------------

    # Install latest Exchange Security Update (SU) after setup completes (Phase 5).
    IncludeFixes          = $true
    # Install pending Windows security/critical updates in Phase 1 before Exchange setup.
    InstallWindowsUpdates = $true

    # -------------------------------------------------------------------------
    # Optional features
    # -------------------------------------------------------------------------

    # Join this DAG after install (DAG must already exist; FSW will be configured by AD).
    # DAGName          = 'DAG01'
    # Copy virtual-directory URLs + send/receive connectors from this existing Exchange server.
    # CopyServerConfig = 'exch01.contoso.com'
    # CertificatePath / CertificatePassword: see "AutoPilot & credentials" near the top.

    # -------------------------------------------------------------------------
    # Behaviour flags
    # -------------------------------------------------------------------------

    # $true to generate pre-flight HTML report and exit without installing anything.
    PreflightOnly      = $false
    # $true to suppress HTML/PDF installation report at end of Phase 6.
    SkipInstallReport  = $false
    # $true to skip CSS-Exchange SetupAssist auto-run when Phase 4 setup.exe fails.
    SkipSetupAssist    = $false
    # $true to remove RSAT tools after Recipient Management Tools install completes.
    # RecipientMgmtCleanup = $false

    # -------------------------------------------------------------------------
    # Word installation document
    # -------------------------------------------------------------------------

    # $true to skip Word (.docx) installation document generation after Phase 6.
    # NoWordDoc = $false

    # Redact RFC1918 IPs, certificate thumbprints, and passwords in the document.
    # Useful when sharing the document with external parties (auditors, customers).
    # CustomerDocument = $false

    # Document language as 2-letter ISO code: 'EN' (default), 'DE' (translated),
    # 'IT'/'FR'/'ES'/... reserved (fall back to 'EN' until translations land).
    # Legacy key 'German = $true' is still accepted (maps to Language='DE').
    # Language = 'DE'

    # Scope of the generated document:
    #   All   — org-wide settings + all Exchange servers + local details (default)
    #   Org   — org-wide chapter only (no per-server hardware / VDir queries)
    #   Local — per-server sections only (no org-wide chapter)
    # DocumentScope = 'All'

    # Limit per-server documentation to specific server names.
    # Local server is always included. Applies when DocumentScope is All or Local.
    # IncludeServers = @('EX01', 'EX02')

    # Path to a custom DOCX template for the installation document.
    # When supplied, the cover page and header come from the template; the chapter
    # body is generated by the script and injected into {{document_body}}.
    # Use tools\Build-InstallationTemplate.ps1 to generate the starter templates in
    # templates\Exchange-installation-document-{DE,EN}.docx.
    # TemplatePath = 'C:\Deploy\my-company-template.docx'

    # -------------------------------------------------------------------------
    # MEAC — MonitorExchangeAuthCertificate
    # -------------------------------------------------------------------------

    # Suppress hybrid-config check when registering the MEAC renewal task.
    # MEACIgnoreHybridConfig       = $false

    # Skip unreachable servers during MEAC setup (useful in DAG with offline members).
    # MEACIgnoreUnreachableServers = $false

    # Email address that receives expiry alerts 60 days before Auth Cert expiration.
    # MEACNotificationEmail = 'exchange-admin@contoso.com'

    # -------------------------------------------------------------------------
    # Relay connectors
    # -------------------------------------------------------------------------
    # AnonymousRelay = $true is the master switch. Without it no connectors are created,
    # even if subnet keys below are set.
    #
    # Modes:
    #   AnonymousRelay = $true + subnet keys set   → connectors created with those subnets
    #   AnonymousRelay = $true + no subnet keys    → connectors created with RFC 5737
    #                                                placeholders (192.0.2.1/32 internal,
    #                                                192.0.2.2/32 external) — visible in EAC
    #                                                for the admin to fill in real subnets
    #   AnonymousRelay omitted / $false            → no connectors created

    # Uncomment all three lines for a typical scanner/printer relay setup:
    # AnonymousRelay    = $true

    # Internal relay connector: anonymous SMTP relay restricted to accepted domains only.
    # Source IPs resolved via SID S-1-5-7 (language-independent).
    # RelaySubnets      = @('10.0.0.0/8', '172.16.0.0/12', '192.168.0.0/16')

    # External relay connector: relay to ANY recipient (Ms-Exch-SMTP-Accept-Any-Recipient).
    # SECURITY: restrict to specific trusted hosts (scanners, printers) — open relay = abuse.
    # ExternalRelaySubnets = @('10.0.1.100')

    # -------------------------------------------------------------------------
    # Log cleanup
    # -------------------------------------------------------------------------
    # Daily scheduled task (02:00, SYSTEM) to delete logs older than N days.
    # Cleans: IIS logs, Exchange transport logs, message tracking logs.
    # Default when using -ConfigFile: 30 days. Set to 0 to disable.
    # LogRetentionDays = 30

    # Folder where the cleanup script and its own log live (created if missing).
    # LogCleanupFolder = 'C:\#service'

    # -------------------------------------------------------------------------
    # Namespace & mail domain
    # -------------------------------------------------------------------------
    # Access namespace for Virtual Directory URL configuration (Phase 6).
    # Sets InternalUrl/ExternalUrl on OWA/ECP/EWS/OAB/MAPI/EAS/PowerShell/IMAP/POP3
    # to https://<Namespace>/... (IMAP/POP3 use <Namespace> directly).
    # Required for a functional deployment — EXpress aborts the config-file run when missing.
    # Namespace = 'mail.contoso.com'

    # Mail domain — root domain for Accepted Domain + Email Address Policy (e.g. @contoso.com).
    # Defaults to the parent of Namespace (mail.contoso.com → contoso.com) when omitted.
    # MailDomain = 'contoso.com'

    # OWA Download Domain — separate FQDN for attachment downloads (CVE-2021-1730 mitigation).
    # Must differ from Namespace; requires matching DNS record and certificate coverage.
    # DownloadDomain = 'download.contoso.com'

}
