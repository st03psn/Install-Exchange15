    function Install-WindowsFeatures( $MajorOSVersion) {
        Write-MyVerbose 'Configuring Windows Features'

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
        Write-MyVerbose ('Checking {0} required Windows features ...' -f $Feats.Count)
        $allFeatureState = Get-WindowsFeature -Name $Feats -ErrorAction SilentlyContinue
        $missing = @($allFeatureState | Where-Object { -not $_.Installed } | ForEach-Object { $_.Name })

        if ($missing.Count -eq 0) {
            Write-MyStep -Label 'Windows Features' -Value ('all {0} already installed' -f $Feats.Count)
        }
        else {
            Write-MyStep -Label 'Windows Features' -Value ('installing {0} missing of {1}' -f $missing.Count, $Feats.Count) -Status Run
            Write-MyVerbose ('Missing features: {0}' -f ($missing -join ', '))
            Install-WindowsFeature $missing | Out-Null
            Write-MyStep -Label 'Windows Features' -Value 'installed successfully' -Status OK
        }

        foreach ( $Feat in $Feats) {
            if ( !( (Get-WindowsFeature -Name $Feat).Installed)) {
                Write-MyError "Feature $Feat appears not to be installed"
                exit $ERR_PROBLEMADDINGFEATURE
            }
        }

        'NET-WCF-MSMQ-Activation45', 'MSMQ' | ForEach-Object {
            if ( (Get-WindowsFeature -Name $_).Installed) {
                Write-MyStep -Label 'Feature' -Value ('removing obsolete: {0}' -f $_) -Status Run
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
            Write-MyVerbose "Processing $Package ($PackageID)"
            $PresenceKey = Test-MyPackage $PackageID
        }
        else {
            # Just install, don't detect
            Write-MyVerbose "Processing $Package"
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
                    Write-MyStep -Label $Package -Value 'extracting hotfix' -Status Run
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

            Write-MyStep -Label $Package -Value ('installing from {0}' -f $RunFrom) -Status Run
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
        Write-MyStep -Label '.NET install blockade' -Value ('set for {0} ({1})' -f $Version, $KB) -Status OK
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
            Write-MyStep -Label '.NET install blockade' -Value ('removed for {0} ({1})' -f $Version, $KB) -Status OK
            Remove-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue | Out-Null
        }
        if ( Get-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue) {
            Write-MyError "Unable to remove registry entry $RegKey\$RegName"
        }
    }

    function Test-Preflight {
        # Informational checks only on first run (phase 0/1); on resume these were already validated
        if ($State['InstallPhase'] -le 1) {
            Write-MySection 'Preflight Checks'

            $Computer = Get-LocalFQDNHostname
            if ( $Computer) {
                Write-MyStep -Label 'Computer name' -Value $Computer
            }

            Write-MyVerbose 'Checking temporary installation folder'
            if (-not (Test-Path $State['InstallPath'])) { New-Item -Path $State['InstallPath'] -ItemType Directory -Force | Out-Null }
            if ( !( Test-Path $State['InstallPath'])) {
                Write-MyError "Can't create temporary folder $($State['InstallPath'])"
                exit $ERR_CANTCREATETEMPFOLDER
            }

            # Downloads cache: all prerequisite packages (.NET, VC++, UCMA, URL Rewrite, hotfixes,
            # Exchange SUs) and CSS-Exchange scripts (HealthChecker, EOMT, SetupAssist,
            # ExchangeExtendedProtection, MEAC, etc.) land here. Pre-staging is automatic —
            # when a file is already present the download is skipped (air-gapped / proxy scenarios).
            $State['SourcesPath'] = Join-Path $State['InstallPath'] 'sources'
            if (-not (Test-Path $State['SourcesPath'])) { New-Item -Path $State['SourcesPath'] -ItemType Directory -Force | Out-Null }
            if ( !( Test-Path $State['SourcesPath'])) {
                Write-MyError "Can't create downloads folder $($State['SourcesPath'])"
                exit $ERR_CANTCREATETEMPFOLDER
            }
            Write-MyVerbose ('Downloads cache: {0}' -f $State['SourcesPath'])

            if ( [System.Version]$MajorOSVersion -ge [System.Version]$WS2016_MAJOR ) {
                Write-MyStep -Label 'Operating System' -Value ('{0}.{1}' -f $MajorOSVersion, $MinorOSVersion)
            }
            else {
                Write-MyError 'Supported operating systems: Windows Server 2016 (Exchange 2016 CU23), Windows Server 2019/2022/2025 (Exchange 2019 CU15+ or Exchange Server SE)'
                exit $ERR_UNEXPECTEDOS
            }
            Write-MyStep -Label 'Server Core' -Value ([string](Test-ServerCore)) -Status Info

            $NetVersion = Get-NETVersion
            $NetVersionText = Get-NetVersionText $NetVersion
            Write-MyStep -Label '.NET Framework' -Value ('{0} ({1})' -f $NetVersionText, $NetVersion)

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
            Write-MyStep -Label 'Elevation' -Value 'Administrator'
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
                Write-MyVerbose 'Checking provided credentials'
                if (Test-Credentials) {
                    Write-MyStep -Label 'Credentials' -Value 'Valid'
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

            # Credentials are now validated and persisted DPAPI-encrypted in state.
            # Plain-text password copies in the config file are no longer needed —
            # scrub them so the file is safe to leave on disk between phases/reboots.
            if ($State['ConfigFile']) {
                Remove-CredentialsFromConfig -Path $State['ConfigFile']
            }
        }

        # Checks below are only relevant before/during setup (phases 0-4); skip after Exchange is installed
        if ($State['InstallPhase'] -le 4) {

        if ( $State["SkipRolesCheck"] -or $State['InstallEdge']) {
            Write-MyStep -Label 'AD admin checks' -Value 'skipped (-SkipRolesCheck)' -Status Info
        }
        else {
            if (! ( Test-SchemaAdmin)) {
                Write-MyError 'Current user is not member of Schema Administrators'
                exit $ERR_RUNNINGNONSCHEMAADMIN
            }
            else {
                Write-MyStep -Label 'AD: Schema Admins' -Value 'Member'
            }

            if (! ( Test-EnterpriseAdmin)) {
                Write-MyError 'User is not member of Enterprise Administrators'
                exit $ERR_RUNNINGNONENTERPRISEADMIN
            }
            else {
                Write-MyStep -Label 'AD: Enterprise Admins' -Value 'Member'
            }
        }
        if (!$State['InstallEdge']) {
            $ADSite = Get-ADSite
            if ( $ADSite) {
                Write-MyStep -Label 'AD Site' -Value $ADSite
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
                Write-MyStep -Label 'Exchange Organisation' -Value ('{0} (existing)' -f $ExOrg)
            }
            else {
                if ( $State['OrganizationName']) {
                    Write-MyStep -Label 'Exchange Organisation' -Value ('{0} (new)' -f $State['OrganizationName'])
                }
                else {
                    Write-MyError 'OrganizationName not specified and no Exchange Organization discovered'
                    exit $ERR_MISSINGORGANIZATIONNAME
                }
            }
        }
        Write-MyVerbose 'Checking if we can access Exchange setup ..'

        if (! (Test-Path (Join-Path $State['SourcePath'] "setup.exe"))) {
            Write-MyError "Can't find Exchange setup at $($State['SourcePath'])"
            exit $ERR_MISSINGEXCHANGESETUP
        }
        else {
            Write-MyStep -Label 'Exchange Setup' -Value (Join-Path $State['SourcePath'] 'setup.exe')
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
                Write-MyStep -Label 'Source files' -Value 'unblocked' -Status OK
            }
        }

        $State['ExSetupVersion'] = Get-DetectedFileVersion "$($State['SourcePath'])\Setup\ServerRoles\Common\ExSetup.exe"
        $SetupVersion = $State['ExSetupVersion']
        $State['SetupVersionText'] = Get-SetupTextVersion $SetupVersion
        Write-MyStep -Label 'Setup Version' -Value $State['SetupVersionText']
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
            Write-MyVerbose 'Not checking roles (NoSetup, Recover or Upgrade mode)'
        }
        else {
            Write-MyVerbose 'Checking roles to install'
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
                Write-MyStep -Label 'Diagnostic Data' -Value 'enabled' -Status Info
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
                    Write-MyStep -Label 'Recovery mode' -Value 'Exchange server object found' -Status OK
                }
                else {
                    if ( Test-Path $EXCHANGEINSTALLKEY) {
                        Write-MyStep -Label 'Existing install' -Value 'detected — switching to Upgrade mode' -Status Info
                        $State['Upgrade'] = $true
                    }
                    else {
                        Write-MyError 'Existing Exchange server object found in Active Directory, but installation missing - please use Recover switch to recover a server'
                        exit $ERR_PROBLEMEXCHANGESERVEREXISTS
                    }
                }
            }

            Write-MyVerbose 'Checking domain membership status ..'
            if (! ( Get-CimInstance Win32_ComputerSystem).PartOfDomain) {
                Write-MyError 'System is not domain-joined'
                exit $ERR_NOTDOMAINJOINED
            }
        }
        Write-MyVerbose 'Checking NIC configuration ..'
        if (! (Get-CimInstance Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True and DHCPEnabled=False')) {
            $AzureHosted = Get-Service | Where-Object { $_.Name -ieq 'Windows Azure Guest Agent' -or $_.Name -ieq 'Windows Azure Network Agent' -or $_.Name -ieq 'Windows Azure Telemetry Service' }
            if ( $AzureHosted) {
                Write-MyError "System doesn't have a static IP addresses configured"
                exit $ERR_NOFIXEDIPADDRESS
            }
            else {
                Write-MyStep -Label 'Static IP check' -Value 'skipped (Azure detected)' -Status Info
            }
        }
        else {
            Write-MyVerbose 'Static IP address(es) assigned.'
        }
        if ( $State['TargetPath']) {
            $Location = Split-Path $State['TargetPath'] -Qualifier
            Write-MyVerbose 'Checking installation path ..'
            if ( !(Test-Path $Location)) {
                Write-MyError "MDB log location unavailable: ($Location)"
                exit $ERR_MDBDBLOGPATH
            }
        }
        if ( $State['InstallMDBLogPath']) {
            $Location = Split-Path $State['InstallMDBLogPath'] -Qualifier
            Write-MyVerbose 'Checking MDB log path ..'
            if ( !(Test-Path $Location)) {
                Write-MyError "MDB log location unavailable: ($Location)"
                exit $ERR_MDBDBLOGPATH
            }
        }
        if ( $State['InstallMDBDBPath']) {
            $Location = Split-Path $State['InstallMDBDBPath'] -Qualifier
            Write-MyVerbose 'Checking MDB database path ..'
            if ( !(Test-Path $Location)) {
                Write-MyError "MDB database location unavailable: ($Location)"
                exit $ERR_MDBDBLOGPATH
            }
        }
        if ( !($State['InstallEdge'])) {
            Write-MyVerbose 'Checking Exchange Forest Schema Version'
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
                Write-MyStep -Label 'Exchange Forest Schema' -Value $EFL
            }
            else {
                Write-MyStep -Label 'Exchange Forest Schema' -Value 'AD not prepared (will run PrepareAD)' -Status Info
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

            Write-MyVerbose 'Checking Exchange Domain Version'
            $EDV = Get-ExchangeDomainLevel
            if ( $EDV) {
                Write-MyStep -Label 'Exchange Domain Version' -Value $EDV
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

            Write-MyVerbose 'Checking domain mode'
            if ( Test-DomainNativeMode -eq $DOMAIN_MIXEDMODE) {
                Write-MyError 'Domain is in mixed mode, native mode is required'
                exit $ERR_ADMIXEDMODE
            }
            else {
                Write-MyStep -Label 'Domain mode' -Value 'native'
            }

            Write-MyVerbose 'Checking Forest Functional Level'
            $FFL = Get-ForestFunctionalLevel
            if ( $MajorSetupVersion -eq $EX2019_MAJOR) {
                if ( $FFL -lt $FOREST_LEVEL2012R2) {
                    Write-MyError ('Exchange Server 2019/SE requires Forest Functionality Level 2012R2 ({0}).' -f $FFL)
                    exit $ERR_ADFORESTLEVEL
                }
                else {
                    Write-MyStep -Label 'Forest Functional Level' -Value ('{0} ({1})' -f (Get-FFLText $FFL), $FFL)
                }
            }
            else {
                if ( $FFL -lt $FOREST_LEVEL2012) {
                    Write-MyError ('Exchange Server 2016 or later requires Forest Functionality Level 2012 ({0}).' -f $FFL)
                    exit $ERR_ADFORESTLEVEL
                }
                else {
                    Write-MyStep -Label 'Forest Functional Level' -Value ('OK ({0})' -f $FFL)
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
        Write-MyStep -Label 'Pre-Flight Report' -Value 'generating' -Status Run
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
        Write-MyStep -Label 'Pre-Flight Report' -Value $reportPath -Status OK
        return $failCount
    }

