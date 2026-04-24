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

