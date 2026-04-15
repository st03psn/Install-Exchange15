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

    # AutoPilot: automatic reboot + resume after each phase
    # Credentials must be entered interactively on first run (stored encrypted in state file)
    AutoPilot = $true

    # -------------------------------------------------------------------------
    # Security hardening  (recommended settings shown below)
    # -------------------------------------------------------------------------

    DisableSSL3  = $true    # Disable SSL 3.0 (POODLE vulnerability)
    DisableRC4   = $true    # Disable RC4 cipher (deprecated)
    EnableECC    = $true    # Prefer ECC key exchange over RSA
    NoCBC        = $false   # Keep CBC enabled — Exchange requires it for compatibility
    EnableAMSI   = $true    # Enable Antimalware Scan Interface for Exchange
    EnableTLS12  = $true    # Enforce TLS 1.2 (disables TLS 1.0/1.1)
    EnableTLS13  = $true    # Enable TLS 1.3 (Windows Server 2022+ only, ignored on older OS)

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

    SkipHealthCheck = $false   # $true to skip CSS-Exchange HealthChecker at end
    NoCheckpoint    = $false   # $true to skip System Restore checkpoints
    PreflightOnly   = $false   # $true to generate pre-flight report and exit
    DiagnosticData  = $false   # $true = /IAcceptExchangeServerLicenseTerms_DiagnosticDataON

    # DoNotEnableEP = $false   # $true to skip Extended Protection configuration
    # NoNet481      = $false   # $true to skip .NET 4.8.1 installation
    # SkipRolesCheck = $false  # $true to skip Schema/Enterprise Admin membership check
    # Lock          = $false   # $true to lock screen during installation
}
