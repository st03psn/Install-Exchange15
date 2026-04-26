    function Install-EXpress_ {
        $ver = $State['MajorSetupVersion']
        Write-MyStep -Label 'Exchange Server' -Value ('installing ({0})' -f $ver) -Status Run
        $PresenceKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{CD981244-E9B8-405A-9026-6AEB9DCEF1F1}'

        if (Get-ItemProperty -Path $PresenceKey -Name InstallDate -ErrorAction SilentlyContinue) {
            Write-MyStep -Label 'Exchange Server' -Value 'already installed (setup skipped)' -Status OK
            return
        }

        if ( $State['Recover']) {
            Write-MyStep -Label 'Setup mode' -Value 'recover' -Status Info
            $Params = '/mode:RecoverServer', $State['IAcceptSwitch'], '/DoNotStartTransport', '/InstallWindowsComponents'
            if ( $State['TargetPath']) {
                $Params += "/TargetDir:`"$($State['TargetPath'])`""
            }
        }
        else {
            if ( $State['Upgrade']) {
                Write-MyStep -Label 'Setup mode' -Value 'upgrade' -Status Info
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

        Write-MyVerbose 'Checking whether Active Directory preparation is required'
        if ($null -ne (Test-ExchangeOrganization $State['OrganizationName'])) {
            Write-MyStep -Label 'Exchange organization' -Value ("'$($State['OrganizationName'])' does not exist — PrepareAD required") -Status Info
            $params += '/PrepareAD', "/OrganizationName:`"$($State['OrganizationName'])`""
            $State['NewExchangeOrg'] = $true   # org created by this run — Enable-AccessNamespaceMailConfig may run
            Save-State $State
        }
        else {
            # Org already exists in this forest — flag it so Phase 5 can adjust defaults for
            # org-wide settings (MaxMessageSize, IANA timezone) that risk overwriting admin choices.
            $State['ExistingOrg'] = $true
            Save-State $State
            $forestlvl = Get-ExchangeForestLevel
            $domainlvl = Get-ExchangeDomainLevel
            Write-MyStep -Label 'Forest Schema / Domain' -Value ('{0} (min {1}) / {2} (min {3})' -f $forestlvl, $MinFFL, $domainlvl, $MinDFL)
            if (($forestlvl -lt $MinFFL) -or ($domainlvl -lt $MinDFL)) {
                Write-MyStep -Label 'AD schema/domain' -Value 'below minimum — PrepareAD required' -Status Info
                $params += '/PrepareAD'
            }
            else {
                Write-MyStep -Label 'Active Directory' -Value 'already prepared (skipping PrepareAD)' -Status OK
                return $false
            }
        }

        Write-MyStep -Label 'PrepareAD' -Value ('org: {0}' -f $State['OrganizationName']) -Status Run
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

