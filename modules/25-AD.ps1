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
            Write-MyDebug "Can't find Exchange Organization object"
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
            Write-MyVerbose ('Started background job to clear AutodiscoverServiceConnectionPoint for {0} (Job ID: {1})' -f $Name, $Job.Id)
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
            catch { Write-MyVerbose ('IIS health probe error (transient, retrying): {0}' -f $_) }
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

