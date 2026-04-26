<#
.SYNOPSIS
    Creates a demo EXpress state file for resuming from Phase 5.

.DESCRIPTION
    Writes HOSTNAME_EXpress_State.xml into $InstallPath so EXpress picks it up
    automatically on the next run and begins at Phase 5 (post-install hardening).

    Edit the variables in the "CUSTOMISE" block below before running.

.EXAMPLE
    # Dry-run — show what would be written without creating the file
    .\New-DemoState.ps1 -WhatIf

    # Create the state file in D:\EXpress
    .\New-DemoState.ps1 -InstallPath D:\EXpress
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$InstallPath = (Split-Path $PSScriptRoot -Parent)
)

# ── CUSTOMISE ─────────────────────────────────────────────────────────────────
$OrganizationName   = 'Contoso'      # Exchange org name
$Namespace          = $null          # external namespace, e.g. 'mail.contoso.com' — leave $null to skip VDir URL config
$MailDomain         = $null          # mail domain, e.g. 'contoso.com'            — leave $null to use Exchange defaults
$DownloadDomain     = $null          # download domain for OWA attachments         — leave $null to skip

# Exchange version this state pretends was installed.
# Phase 5 SU selection uses ExSetupVersion — set this to match the target CU.
# Common values:
#   Exchange SE RTM   : '15.02.2562.017'
#   Exchange 2019 CU15: '15.02.1748.008'
#   Exchange 2016 CU23: '15.01.2507.006'
$ExSetupVersion     = '15.02.2562.017'    # Exchange SE RTM

# Certificate for Phase 5 import.  $null = skip cert import.
$CertificatePath    = $null          # e.g. 'C:\Certs\mail.contoso.com.pfx'

# Exchange installation target path (where Exchange landed).
$ExchangeInstallPath = 'C:\Program Files\Microsoft\Exchange Server\V15'

# DAG name — leave $null if standalone (no DAG join in Phase 6).
$DAGName            = $null
# ── END CUSTOMISE ─────────────────────────────────────────────────────────────

$ReportsPath   = Join-Path $InstallPath 'reports'
$TranscriptLog = Join-Path $ReportsPath ('{0}_EXpress_Install_DEMO_{1}.log' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd-HHmmss'))
$StateFile     = Join-Path $InstallPath ('{0}_EXpress_State.xml' -f $env:COMPUTERNAME)

if (-not (Test-Path $ReportsPath)) {
    New-Item -Path $ReportsPath -ItemType Directory -Force | Out-Null
}

# ── Advanced Feature defaults (mirrors Get-AdvancedFeatureCatalog) ─────────────
# All set to catalog defaults.  Toggle individual keys as needed for testing.
$AdvancedFeatures = @{
    # TLS
    DisableSSL3         = $true
    DisableRC4          = $true
    EnableECC           = $true
    NoCBC               = $false
    EnableAMSI          = $true
    EnableTLS12         = $true
    EnableTLS13         = $true     # gated by Condition in real run (WS2022+)
    DoNotEnableEP       = $false
    # Hardening
    SMBv1Disable        = $true
    NetBIOSDisable      = $true
    LLMNRDisable        = $true
    MDNSDisable         = $true
    WDigestDisable      = $true
    LSAProtection       = $true
    LmCompat5           = $true
    SerializedDataSig   = $true
    ShutdownTrackerOff  = $true
    HSTS                = $true
    MAPIEncryption      = $true
    HTTP2Disable        = $true
    CredentialGuardOff  = $true
    UnnecessaryServices = $true
    WindowsSearchOff    = $true
    CRLTimeout          = $true
    RootCAAutoUpdate    = $true
    SMTPBannerHarden    = $true
    # Performance
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
    # ExchangePolicy
    ModernAuth          = $true
    OWASessionTimeout6h = $true
    DisableTelemetry    = $true
    MapiHttp            = $true
    MaxMessageSize150MB = $true
    MessageExpiration7d = $true
    HtmlNDR             = $true
    ShadowRedundancy    = $false    # only valid with DAGName
    SafetyNet2d         = $true
    # PostConfig
    MECA                = $true
    AntispamAgents      = $true
    SSLOffloading       = $true
    MRSProxy            = $true
    IANATimezone        = $true
    AnonymousRelay      = $false    # only valid with RelaySubnets
    AccessNamespaceMail = $false    # only valid with Namespace + NewExchangeOrg
    SkipHealthCheck     = $false
    RBACReport          = $true
    RunEOMT             = $false
    # InstallFlow
    AutoApproveWindowsUpdates = $false
    DiagnosticData      = $false
    Lock                = $false
    SkipRolesCheck      = $false
    NoCheckpoint        = $false
}

$State = @{
    # ── Phase control ──────────────────────────────────────────────────────────
    InstallPhase          = 4    # +1 → Phase 5 on first run
    LastSuccessfulPhase   = 4

    # ── Paths ──────────────────────────────────────────────────────────────────
    InstallPath           = $InstallPath
    ReportsPath           = $ReportsPath
    TranscriptFile        = $TranscriptLog
    TargetPath            = $ExchangeInstallPath

    # ── Exchange identity ──────────────────────────────────────────────────────
    OrganizationName      = $OrganizationName
    ExSetupVersion        = $ExSetupVersion
    SetupVersion          = $ExSetupVersion

    # ── Install mode ───────────────────────────────────────────────────────────
    InstallMailbox        = $true
    InstallEdge           = $false
    Autopilot             = $false    # false = Copilot (interactive); no auto-reboot
    ConfigDriven          = $false
    ConfigFile            = $null
    Recover               = $false
    Upgrade               = $false
    NoSetup               = $false
    NewExchangeOrg        = $false    # set $true if EXpress created the org in Phase 3/4

    # ── Credentials (DPAPI-encrypted in real runs — empty in demo) ────────────
    InstallingUser        = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    AdminAccount          = $null
    AdminPassword         = $null
    MEACAutomationUser    = $null
    MEACAutomationPW      = $null

    # ── Namespace / Mail config ────────────────────────────────────────────────
    Namespace             = $Namespace
    MailDomain            = $MailDomain
    DownloadDomain        = $DownloadDomain
    SCP                   = $null
    EdgeDNSSuffix         = $null

    # ── Certificate ────────────────────────────────────────────────────────────
    CertificatePath       = $CertificatePath
    CertificatePassword   = $null    # will prompt at Phase 5 if CertificatePath is set

    # ── DAG / Connectors ───────────────────────────────────────────────────────
    DAGName               = $DAGName
    RelaySubnets          = $null
    ExternalRelaySubnets  = $null

    # ── Server config import ───────────────────────────────────────────────────
    CopyServerConfig      = $false
    ServerConfigExportPath = $null

    # ── Optional installs ──────────────────────────────────────────────────────
    IncludeFixes                  = $false   # skip SU download during Phase 5
    InstallWindowsUpdates         = $false
    InstallRecipientManagement    = $false
    InstallManagementTools        = $false
    RecipientMgmtCleanup          = $false

    # ── Report / Doc options ───────────────────────────────────────────────────
    SkipInstallReport     = $false
    NoWordDoc             = $false
    CustomerDocument      = $false
    Language              = 'EN'
    DocumentScope         = 'All'
    IncludeServers        = ''
    TemplatePath          = $null
    SkipSetupAssist       = $false
    PreflightOnly         = $false

    # ── Logging ────────────────────────────────────────────────────────────────
    LogVerbose            = $false
    LogDebug              = $false
    Verbose               = 'SilentlyContinue'

    # ── Source media (Phase 5 does not need the ISO) ───────────────────────────
    SourcePath            = $null
    SourceImage           = $null

    # ── Flags set during earlier phases ───────────────────────────────────────
    RebootRequired        = $false
    DoNotOpenServerManagerAtLogon = 1
    Install481            = $true    # .NET 4.8.1 was installed in Phase 2
    VCRedist2012          = $true
    VCRedist2013          = $true
    HCReportPath          = $null

    # ── MDB layout (Phase 1 init) ──────────────────────────────────────────────
    InstallMDBDBPath      = $null
    InstallMDBLogPath     = $null
    InstallMDBName        = $null

    # ── Advanced features ──────────────────────────────────────────────────────
    AdvancedFeatures      = $AdvancedFeatures
    SuppressAdvancedPrompt = $false

    # ── Flat projection of AdvancedFeatures (mirrors what EXpress builds at startup) ──
    DisableSSL3           = $AdvancedFeatures['DisableSSL3']
    DisableRC4            = $AdvancedFeatures['DisableRC4']
    EnableECC             = $AdvancedFeatures['EnableECC']
    NoCBC                 = $AdvancedFeatures['NoCBC']
    EnableCBC             = -not $AdvancedFeatures['NoCBC']    # derived
    EnableAMSI            = $AdvancedFeatures['EnableAMSI']
    EnableTLS12           = $AdvancedFeatures['EnableTLS12']
    EnableTLS13           = $AdvancedFeatures['EnableTLS13']
    DoNotEnableEP         = $AdvancedFeatures['DoNotEnableEP']
    SMBv1Disable          = $AdvancedFeatures['SMBv1Disable']
    NetBIOSDisable        = $AdvancedFeatures['NetBIOSDisable']
    LLMNRDisable          = $AdvancedFeatures['LLMNRDisable']
    MDNSDisable           = $AdvancedFeatures['MDNSDisable']
    WDigestDisable        = $AdvancedFeatures['WDigestDisable']
    LSAProtection         = $AdvancedFeatures['LSAProtection']
    LmCompat5             = $AdvancedFeatures['LmCompat5']
    SerializedDataSig     = $AdvancedFeatures['SerializedDataSig']
    ShutdownTrackerOff    = $AdvancedFeatures['ShutdownTrackerOff']
    HSTS                  = $AdvancedFeatures['HSTS']
    MAPIEncryption        = $AdvancedFeatures['MAPIEncryption']
    HTTP2Disable          = $AdvancedFeatures['HTTP2Disable']
    CredentialGuardOff    = $AdvancedFeatures['CredentialGuardOff']
    UnnecessaryServices   = $AdvancedFeatures['UnnecessaryServices']
    WindowsSearchOff      = $AdvancedFeatures['WindowsSearchOff']
    CRLTimeout            = $AdvancedFeatures['CRLTimeout']
    RootCAAutoUpdate      = $AdvancedFeatures['RootCAAutoUpdate']
    SMTPBannerHarden      = $AdvancedFeatures['SMTPBannerHarden']
    MaxConcurrentAPI      = $AdvancedFeatures['MaxConcurrentAPI']
    DiskAllocHint         = $AdvancedFeatures['DiskAllocHint']
    CtsProcAffinity       = $AdvancedFeatures['CtsProcAffinity']
    NodeRunnerMemLimit    = $AdvancedFeatures['NodeRunnerMemLimit']
    MapiFeGC              = $AdvancedFeatures['MapiFeGC']
    NICPowerMgmtOff       = $AdvancedFeatures['NICPowerMgmtOff']
    RSSEnable             = $AdvancedFeatures['RSSEnable']
    TCPTuning             = $AdvancedFeatures['TCPTuning']
    TCPOffloadOff         = $AdvancedFeatures['TCPOffloadOff']
    IPv4OverIPv6Off       = $AdvancedFeatures['IPv4OverIPv6Off']
    ModernAuth            = $AdvancedFeatures['ModernAuth']
    OWASessionTimeout6h   = $AdvancedFeatures['OWASessionTimeout6h']
    DisableTelemetry      = $AdvancedFeatures['DisableTelemetry']
    MapiHttp              = $AdvancedFeatures['MapiHttp']
    MaxMessageSize150MB   = $AdvancedFeatures['MaxMessageSize150MB']
    MessageExpiration7d   = $AdvancedFeatures['MessageExpiration7d']
    HtmlNDR               = $AdvancedFeatures['HtmlNDR']
    ShadowRedundancy      = $AdvancedFeatures['ShadowRedundancy']
    SafetyNet2d           = $AdvancedFeatures['SafetyNet2d']
    MECA                  = $AdvancedFeatures['MECA']
    AntispamAgents        = $AdvancedFeatures['AntispamAgents']
    SSLOffloading         = $AdvancedFeatures['SSLOffloading']
    MRSProxy              = $AdvancedFeatures['MRSProxy']
    IANATimezone          = $AdvancedFeatures['IANATimezone']
    AnonymousRelay        = $AdvancedFeatures['AnonymousRelay']
    AccessNamespaceMail   = $AdvancedFeatures['AccessNamespaceMail']
    SkipHealthCheck       = $AdvancedFeatures['SkipHealthCheck']
    RBACReport            = $AdvancedFeatures['RBACReport']
    RunEOMT               = $AdvancedFeatures['RunEOMT']
    AutoApproveWindowsUpdates = $AdvancedFeatures['AutoApproveWindowsUpdates']
    DiagnosticData        = $AdvancedFeatures['DiagnosticData']
    Lock                  = $AdvancedFeatures['Lock']
    SkipRolesCheck        = $AdvancedFeatures['SkipRolesCheck']
    NoCheckpoint          = $AdvancedFeatures['NoCheckpoint']
    DoNotEnableEP_FEEWS   = $false
}

if ($PSCmdlet.ShouldProcess($StateFile, 'Write EXpress demo state')) {
    $State | Export-Clixml -Path $StateFile -Force
    Write-Host "State written to: $StateFile"
    Write-Host "Phase resume: 4 + 1 = 5"
    Write-Host ""
    Write-Host "Run EXpress from the install directory to continue:"
    Write-Host "  pwsh -File '$InstallPath\EXpress.ps1'"
    Write-Host ""
    Write-Host "Or pass -Phase explicitly to override the state:"
    Write-Host "  pwsh -File '$InstallPath\EXpress.ps1' -Phase 5"
}
