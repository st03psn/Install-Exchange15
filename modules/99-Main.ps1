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
                    if ($State['AccessNamespaceMail'] -and $State['Namespace']) {
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

                # IANA timezone mapping check (Exchange 2019 CU14+)
                if (-not $State['InstallEdge']) {
                    Step-P5 'IANA timezone mapping'
                    Register-ExecutedCommand -Category 'ExchangeTuning' -Command 'Set-OrganizationConfig -UseIanaTimeZoneId $true  # Exchange 2019 CU14+'
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
                if ($State['AccessNamespaceMail'] -and $State['Namespace'] -and -not $State['InstallEdge']) {
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

