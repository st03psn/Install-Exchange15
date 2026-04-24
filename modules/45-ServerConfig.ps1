    function Export-SourceServerConfig {
        param([string]$SourceServer)
        Write-MyOutput ('Exporting configuration from source server {0}' -f $SourceServer)
        $configPath = Join-Path $State['InstallPath'] ('{0}_EXpress_Config.xml' -f $SourceServer)

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
# Exchange Log Cleanup Script — generated by EXpress.ps1
# Runs daily via Scheduled Task; retention: $days days for Exchange/IIS logs, 30 days for own logs

param([int]`$DaysToKeep = $days)

`$ScriptDir  = Split-Path -Path `$MyInvocation.MyCommand.Path
`$LogFolder  = Join-Path `$ScriptDir 'logs'
`$LogFile    = Join-Path `$LogFolder ('LogCleanup_{0}.log' -f (Get-Date -Format 'yyyyMM'))
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
            Register-ExecutedCommand -Category 'ScheduledTask' -Command ("Register-ScheduledTask -TaskName '$taskName' -TaskPath '$taskPath' -Action (New-ScheduledTaskAction …Clean-ExchangeLogs.ps1 -RetentionDays $days) -Trigger (Daily 02:00) -Principal SYSTEM -RunLevel Highest")
            Register-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Action $action `
                -Trigger $trigger -Settings $settings -Principal $principal -ErrorAction Stop | Out-Null
            Write-MyOutput ('Scheduled task "{0}" registered — runs daily at 02:00, retention {1} days' -f $taskName, $days)
        }
        catch {
            Write-MyWarning ('Failed to register log cleanup task: {0}' -f $_.Exception.Message)
        }
    }
