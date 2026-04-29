    function Get-OrganizationReportData {
        # Collects org-wide Exchange settings. No server-specific data here.
        # Safe to call from New-InstallationDocument in all scenarios.
        $org = @{}

        # Org config
        try { $org.OrgConfig = Get-OrganizationConfig -ErrorAction Stop } catch { $org.OrgConfig = $null }

        # Accepted / Remote Domains
        try { $org.AcceptedDomains = @(Get-AcceptedDomain -ErrorAction Stop) } catch { $org.AcceptedDomains = @() }
        try { $org.RemoteDomains   = @(Get-RemoteDomain   -ErrorAction Stop) } catch { $org.RemoteDomains   = @() }

        # Email Address Policies
        try { $org.EmailAddressPolicies = @(Get-EmailAddressPolicy -ErrorAction Stop) } catch { $org.EmailAddressPolicies = @() }

        # Transport
        try { $org.TransportConfig = Get-TransportConfig -ErrorAction Stop } catch { $org.TransportConfig = $null }
        try { $org.TransportRules  = @(Get-TransportRule  -ErrorAction Stop | Select-Object Name, State, Priority, Mode, Comments) } catch { $org.TransportRules = @() }

        # Journal / Retention / DLP
        try { $org.JournalRules     = @(Get-JournalRule      -ErrorAction Stop) } catch { $org.JournalRules     = @() }
        try { $org.RetentionPolicies   = @(Get-RetentionPolicy    -ErrorAction Stop) } catch { $org.RetentionPolicies   = @() }
        try { $org.RetentionPolicyTags = @(Get-RetentionPolicyTag -ErrorAction Stop) } catch { $org.RetentionPolicyTags = @() }
        try { $org.DlpPolicies      = @(Get-DlpPolicy        -ErrorAction Stop) } catch { $org.DlpPolicies      = @() }

        # Mobile / OWA policies
        try { $org.MobileDevicePolicies = @(Get-MobileDeviceMailboxPolicy -ErrorAction Stop) } catch { $org.MobileDevicePolicies = @() }
        try { $org.OwaPolicies          = @(Get-OwaMailboxPolicy          -ErrorAction Stop) } catch { $org.OwaPolicies          = @() }

        # DAGs (all)
        try {
            $org.DAGs = @(Get-DatabaseAvailabilityGroup -Status -ErrorAction Stop | ForEach-Object {
                $dag = $_
                $copies = @{}
                try {
                    Get-MailboxDatabaseCopyStatus -Server ($dag.Servers | Select-Object -First 1) -ErrorAction SilentlyContinue | ForEach-Object {
                        $copies[$_.DatabaseName] = $copies[$_.DatabaseName] + @($_)
                    }
                } catch { Write-MyVerbose ('Get-MailboxDatabaseCopyStatus failed for DAG {0}: {1}' -f $dag.Name, $_) }
                [pscustomobject]@{
                    DAG             = $dag
                    DatabaseCopies  = $copies
                }
            })
        } catch { $org.DAGs = @() }

        # Send Connectors (org-scoped, not per-server)
        try { $org.SendConnectors = @(Get-SendConnector -ErrorAction Stop) } catch { $org.SendConnectors = @() }

        # Federation
        try { $org.FederationTrust  = @(Get-FederationTrust  -ErrorAction Stop) } catch { $org.FederationTrust  = @() }
        try { $org.FederationOrg    = Get-FederatedOrganizationIdentifier -ErrorAction SilentlyContinue } catch { $org.FederationOrg = $null }

        # Hybrid
        try { $org.HybridConfig = Get-HybridConfiguration -ErrorAction Stop } catch { $org.HybridConfig = $null }

        # OAuth / AuthConfig
        try { $org.AuthConfig = Get-AuthConfig -ErrorAction Stop } catch { $org.AuthConfig = $null }
        try { $org.IntraOrgConnectors = @(Get-IntraOrganizationConnector -ErrorAction Stop) } catch { $org.IntraOrgConnectors = @() }

        # RBAC role groups (with members). Keep members as name/recipient-type only — full DN bloats the doc.
        $org.RoleGroups = @()
        try {
            $rgList = @(Get-RoleGroup -ErrorAction Stop | Sort-Object Name)
            foreach ($rg in $rgList) {
                $mem = @()
                try {
                    $mem = @(Get-RoleGroupMember -Identity $rg.Name -ErrorAction Stop |
                             Select-Object @{n='Name';e={$_.Name}}, @{n='Type';e={$_.RecipientType}})
                } catch { Write-MyVerbose ('Get-RoleGroupMember failed for {0}: {1}' -f $rg.Name, $_) }
                $org.RoleGroups += [pscustomobject]@{
                    Name        = $rg.Name
                    Description = $rg.Description
                    Members     = $mem
                    ManagedBy   = @($rg.ManagedBy | ForEach-Object { $_.ToString() })
                }
            }
        } catch { Write-MyVerbose ('Get-RoleGroup enumeration failed: {0}' -f $_) }

        # Admin Audit Log Config (org-wide; controls which cmdlets/parameters are recorded in the admin audit log)
        try { $org.AdminAuditLog = Get-AdminAuditLogConfig -ErrorAction Stop } catch { $org.AdminAuditLog = $null }

        # Anti-spam filter configuration (org-wide settings objects; only present when anti-spam agents are installed)
        try { $org.ContentFilterConfig   = Get-ContentFilterConfig   -ErrorAction Stop } catch { $org.ContentFilterConfig   = $null }
        try { $org.SenderFilterConfig    = Get-SenderFilterConfig    -ErrorAction Stop } catch { $org.SenderFilterConfig    = $null }
        try { $org.RecipientFilterConfig = Get-RecipientFilterConfig -ErrorAction Stop } catch { $org.RecipientFilterConfig = $null }
        try { $org.SenderIdConfig        = Get-SenderIdConfig        -ErrorAction Stop } catch { $org.SenderIdConfig        = $null }

        # Auth Certificate (current + next) — org-wide (replicated to all servers).
        try { $org.AuthCertCurrent = Get-AuthConfig -ErrorAction Stop |
                 Select-Object CurrentCertificateThumbprint, PreviousCertificateThumbprint, NextCertificateThumbprint,
                               ServiceName, Realm } catch { $org.AuthCertCurrent = $null }

        # Scheduled Tasks (Exchange-related: MEAC auth-cert renewal, EXpress log cleanup).
        # ServerManager auto-start disable is an OS-level hardening step documented in Chapter 8, not a scheduled
        # task worth listing here.
        $org.ScheduledTasks = @()
        try {
            $foundTasks = @{}
            # Direct name lookup for known task names (fast path)
            # CSS-Exchange MEAC task is named "Daily Auth Certificate Check" (Register-AuthCertificateRenewalTask.ps1 default)
            $knownNames = @('Daily Auth Certificate Check','MonitorExchangeAuthCertificate','Exchange Log Cleanup','EXpressLogCleanup')
            foreach ($tn in $knownNames) {
                try { Get-ScheduledTask -TaskName $tn -ErrorAction SilentlyContinue | ForEach-Object { $foundTasks[$_.TaskName] = $_ } } catch { Write-MyVerbose ('Get-ScheduledTask lookup for {0} failed: {1}' -f $tn, $_) }
            }
            # Broad pattern search — catches variants across CSS-Exchange releases
            try {
                Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
                    $_.TaskName -match 'Daily Auth Certificate|MonitorExchangeAuth|ExchangeLogClean|EXpressLog'
                } | ForEach-Object { $foundTasks[$_.TaskName] = $_ }
            } catch { Write-MyVerbose ('Get-ScheduledTask pattern search failed: {0}' -f $_) }
            foreach ($task in $foundTasks.Values) {
                $info = try { Get-ScheduledTaskInfo -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction SilentlyContinue } catch { $null }
                $org.ScheduledTasks += [pscustomobject]@{
                    Name      = $task.TaskName
                    Path      = $task.TaskPath
                    State     = $task.State
                    LastRun   = if ($info) { $info.LastRunTime }   else { $null }
                    NextRun   = if ($info) { $info.NextRunTime }   else { $null }
                    LastResult= if ($info) { $info.LastTaskResult } else { $null }
                    Actions   = @($task.Actions | ForEach-Object { if ($_.Execute) { "$($_.Execute) $($_.Arguments)".Trim() } })
                }
            }
        } catch { Write-MyVerbose ('Scheduled task enumeration failed: {0}' -f $_) }

        return $org
    }

    function Get-ServerReportData {
        param([Parameter(Mandatory)][string]$ServerName)

        $srv = @{ ServerName = $ServerName }

        # Exchange server object
        try { $srv.ExServer = Get-ExchangeServer -Identity $ServerName -ErrorAction Stop } catch { $srv.ExServer = $null }

        # Databases on this server
        try { $srv.Databases = @(Get-MailboxDatabase -Server $ServerName -Status -ErrorAction Stop) } catch { $srv.Databases = @() }

        # Virtual directories
        try { $srv.VDirOWA    = @(Get-OwaVirtualDirectory              -Server $ServerName -ADPropertiesOnly -ErrorAction Stop) } catch { $srv.VDirOWA    = @() }
        try { $srv.VDirECP    = @(Get-EcpVirtualDirectory              -Server $ServerName -ADPropertiesOnly -ErrorAction Stop) } catch { $srv.VDirECP    = @() }
        try { $srv.VDirEWS    = @(Get-WebServicesVirtualDirectory      -Server $ServerName -ADPropertiesOnly -ErrorAction Stop) } catch { $srv.VDirEWS    = @() }
        try { $srv.VDirAS     = @(Get-ActiveSyncVirtualDirectory       -Server $ServerName -ADPropertiesOnly -ErrorAction Stop) } catch { $srv.VDirAS     = @() }
        try { $srv.VDirOAB    = @(Get-OabVirtualDirectory              -Server $ServerName -ADPropertiesOnly -ErrorAction Stop) } catch { $srv.VDirOAB    = @() }
        try { $srv.VDirMAPI   = @(Get-MapiVirtualDirectory             -Server $ServerName -ADPropertiesOnly -ErrorAction Stop) } catch { $srv.VDirMAPI   = @() }
        try { $srv.VDirPW     = @(Get-PowerShellVirtualDirectory       -Server $ServerName -ADPropertiesOnly -ErrorAction Stop) } catch { $srv.VDirPW     = @() }
        try { $srv.AutodiscoverSCP = Get-ClientAccessService           -Identity $ServerName -ErrorAction Stop } catch { $srv.AutodiscoverSCP = $null }

        # Connectors (Receive only — Send is org-scoped)
        try { $srv.ReceiveConnectors = @(Get-ReceiveConnector -Server $ServerName -ErrorAction Stop) } catch { $srv.ReceiveConnectors = @() }

        # IMAP/POP settings (local only — remote Exchange management remoting requires a separate PS session)
        $srv.ImapSettings = $null
        $srv.PopSettings  = $null
        if ($ServerName -ieq $env:COMPUTERNAME) {
            try { $srv.ImapSettings = Get-ImapSettings -Server $ServerName -ErrorAction Stop } catch { Write-MyVerbose ('Get-ImapSettings failed for {0}: {1}' -f $ServerName, $_) }
            try { $srv.PopSettings  = Get-PopSettings  -Server $ServerName -ErrorAction Stop } catch { Write-MyVerbose ('Get-PopSettings failed for {0}: {1}' -f $ServerName, $_) }
        }

        # Certificates — query Cert:\LocalMachine\My directly (immune to implicit-remoting
        # re-entrant session interference that silently empties Get-ExchangeCertificate -Server results).
        # Get-ExchangeCertificate is called without -Server to build a thumbprint→Services map;
        # -Server variant is the fallback. Only Exchange-managed thumbprints are shown when the map
        # is populated; all non-phantom store certs are shown as fallback if both Exchange calls fail.
        $certStoreRaw = @()
        try {
            $certStoreRaw = @(Get-ChildItem Cert:\LocalMachine\My -ErrorAction Stop |
                Where-Object { $_.Thumbprint -and $_.NotAfter -gt [datetime]'1970-01-01' })
        } catch { Write-MyVerbose ('Cert:\LocalMachine\My query failed: {0}' -f $_) }

        $exSvcMap = @{}   # thumbprint -> Services string
        try {
            Get-ExchangeCertificate -ErrorAction Stop | ForEach-Object {
                if ($_.Thumbprint) { $exSvcMap[$_.Thumbprint] = "$($_.Services)" }
            }
        } catch {
            Write-MyVerbose ('Get-ExchangeCertificate (services) failed, trying -Server {0}: {1}' -f $ServerName, $_)
            try {
                Get-ExchangeCertificate -Server $ServerName -ErrorAction Stop | ForEach-Object {
                    if ($_.Thumbprint) { $exSvcMap[$_.Thumbprint] = "$($_.Services)" }
                }
            } catch { Write-MyVerbose ('Get-ExchangeCertificate -Server {0} (services) failed: {1}' -f $ServerName, $_) }
        }

        # Filter to Exchange-managed certs when possible; fall back to full store list.
        $certBase = if ($exSvcMap.Count -gt 0) {
            @($certStoreRaw | Where-Object { $exSvcMap.ContainsKey($_.Thumbprint) })
        } else {
            $certStoreRaw
        }
        $srv.Certificates = @($certBase | ForEach-Object {
            [PSCustomObject]@{
                Thumbprint   = $_.Thumbprint
                Subject      = $_.Subject
                NotAfter     = $_.NotAfter
                Services     = if ($exSvcMap.ContainsKey($_.Thumbprint)) { $exSvcMap[$_.Thumbprint] } else { '—' }
                IsSelfSigned = ($_.Subject -eq $_.Issuer)
            }
        })

        # Transport agents (only present on servers with Hub Transport)
        try { $srv.TransportAgents = @(Get-TransportAgent -ErrorAction Stop) } catch { $srv.TransportAgents = @() }

        # Database copy status (per server; runs against local server where available)
        try { $srv.DatabaseCopies = @(Get-MailboxDatabaseCopyStatus -Server $ServerName -ErrorAction Stop |
                                      Select-Object Name, DatabaseName, Status, ContentIndexState, CopyQueueLength, ReplayQueueLength, ActivationPreference, MailboxServer) } catch { $srv.DatabaseCopies = @() }

        # Defender preferences — only meaningful for local server (remote would need CIM/PSSession)
        $srv.DefenderExclusions = $null
        if ($ServerName -ieq $env:COMPUTERNAME) {
            try {
                $mp       = Get-MpPreference    -ErrorAction Stop
                $mpStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
                $srv.DefenderExclusions = [pscustomobject]@{
                    ExclusionPath      = @($mp.ExclusionPath)
                    ExclusionProcess   = @($mp.ExclusionProcess)
                    ExclusionExtension = @($mp.ExclusionExtension)
                    RealTimeEnabled    = if ($mpStatus) { $mpStatus.RealTimeProtectionEnabled } else { -not $mp.DisableRealtimeMonitoring }
                    AMRunningMode      = if ($mpStatus) { [string]$mpStatus.AMRunningMode } else { $null }
                }
            } catch { Write-MyVerbose ('Get-MpPreference failed for {0}: {1}' -f $ServerName, $_) }
        }

        # IIS log configuration (local only — remote IIS queries require WinRM/PSSession, out of scope)
        $srv.IISLogs = $null
        if ($ServerName -ieq $env:COMPUTERNAME) {
            try {
                Import-Module WebAdministration -ErrorAction SilentlyContinue
                $sites = @(Get-Website -ErrorAction SilentlyContinue | Where-Object { $_.Name -in 'Default Web Site','Exchange Back End' } | ForEach-Object {
                    [pscustomobject]@{
                        Name      = $_.Name
                        LogDir    = $_.LogFile.Directory
                        LogFormat = $_.LogFile.LogFormat
                        Period    = $_.LogFile.Period
                    }
                })
                $srv.IISLogs = [pscustomobject]@{
                    Sites = $sites
                    ExchangeLogPath = Join-Path (Split-Path $env:ExchangeInstallPath -Parent) 'Logging'
                }
            } catch { Write-MyVerbose ('IIS log configuration query failed for {0}: {1}' -f $ServerName, $_) }
        }

        # Hardware/OS data — direct CIM for local server, CIM/WSMan + prompt for remote
        if ($ServerName -ieq $env:COMPUTERNAME) {
            $srv.RemoteData = @{
                ComputerName    = $ServerName
                Reachable       = $true
                Error           = $null
                OS              = Get-CimInstance Win32_OperatingSystem          -ErrorAction SilentlyContinue
                CPU             = Get-CimInstance Win32_Processor                 -ErrorAction SilentlyContinue
                ComputerSys     = Get-CimInstance Win32_ComputerSystem            -ErrorAction SilentlyContinue
                PageFile        = Get-CimInstance Win32_PageFileSetting           -ErrorAction SilentlyContinue
                Volumes         = @(Get-CimInstance Win32_Volume -Filter 'DriveType=3'                          -ErrorAction SilentlyContinue)
                NICs            = @(Get-CimInstance Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=TRUE'  -ErrorAction SilentlyContinue)
                TimeZone        = Get-CimInstance Win32_TimeZone                  -ErrorAction SilentlyContinue
                NICDrivers      = @(Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object { $_.Status -ne 'Disconnected' } | Select-Object Name, DriverVersion, DriverDate, LinkSpeed, InterfaceDescription)
                TlsCipherSuites = @(Get-TlsCipherSuite -ErrorAction SilentlyContinue | Select-Object Name, Exchange, Hash, KeyExchange)
                VCRuntimes      = @(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like 'Microsoft Visual C++ *' } | Select-Object DisplayName, DisplayVersion, InstallDate | Sort-Object DisplayName)
            }
        } else {
            $srv.RemoteData = Invoke-RemoteQueryWithPrompt -ComputerName $ServerName
        }

        return $srv
    }

    function Get-InstallationReportData {
        param(
            [ValidateSet('All','Org','Local')][string]$Scope = 'All',
            [string[]]$IncludeServers = @()
        )

        $data = @{
            Org     = $null
            Servers = @()
            Local   = @{}
        }

        # Org-wide data
        if ($Scope -in 'All','Org') {
            Write-MyVerbose 'Collecting org-wide Exchange configuration'
            $data.Org = Get-OrganizationReportData
        }

        # Per-server data
        if ($Scope -in 'All','Local') {
            try {
                $allServers = @(Get-ExchangeServer -ErrorAction Stop | Sort-Object Name)
                if ($IncludeServers.Count -gt 0) {
                    $allServers = @($allServers | Where-Object { $_.Name -in $IncludeServers })
                }
                foreach ($srv in $allServers) {
                    Write-MyVerbose ('Collecting data for server {0}' -f $srv.Name)
                    $srvData = Get-ServerReportData -ServerName $srv.Name
                    $srvData.IsLocalServer = ($srv.Name -ieq $env:COMPUTERNAME)
                    $data.Servers += $srvData
                }
            } catch {
                Write-MyWarning ('Get-InstallationReportData: could not enumerate Exchange servers: {0}' -f $_.Exception.Message)
            }
        }

        # Local system data (always, for the server running EXpress)
        $data.Local.OS          = Get-CimInstance Win32_OperatingSystem        -ErrorAction SilentlyContinue
        $data.Local.CPU         = Get-CimInstance Win32_Processor               -ErrorAction SilentlyContinue
        $data.Local.ComputerSys = Get-CimInstance Win32_ComputerSystem          -ErrorAction SilentlyContinue
        $data.Local.PageFile    = Get-CimInstance Win32_PageFileSetting         -ErrorAction SilentlyContinue
        $data.Local.Volumes     = @(Get-CimInstance Win32_Volume -Filter 'DriveType=3' -ErrorAction SilentlyContinue)
        $data.Local.NICs        = @(Get-CimInstance Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=TRUE' -ErrorAction SilentlyContinue)

        return $data
    }

    function Get-RemoteServerData {
        <#
        .SYNOPSIS
            Collects hardware/OS/pagefile/volume/NIC data from a remote Exchange server via CIM/WSMan.
        .DESCRIPTION
            Uses CIM over WSMan (WinRM TCP 5985/5986, Kerberos). NOT WMI/DCOM.
            Returns a uniform hashtable; on failure sets Reachable = $false with Error text.
            Timeout 30 s; always disposes CimSession in finally.
            Pre-requisites on target: see tools\Enable-EXpressRemoteQuery.ps1 or docs\remote-query-setup.md.
        #>
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][string]$ComputerName,
            [int]$TimeoutSec = 30
        )

        $result = @{
            ComputerName = $ComputerName
            Reachable    = $false
            Error        = $null
            OS           = $null
            CPU          = $null
            ComputerSys  = $null
            PageFile     = $null
            Volumes      = @()
            NICs         = @()
        }

        $session = $null
        try {
            $opt = New-CimSessionOption -Protocol WSMan
            $session = New-CimSession -ComputerName $ComputerName -SessionOption $opt `
                                      -OperationTimeoutSec $TimeoutSec -ErrorAction Stop

            $result.OS          = Get-CimInstance -CimSession $session -ClassName Win32_OperatingSystem          -ErrorAction Stop
            $result.CPU         = Get-CimInstance -CimSession $session -ClassName Win32_Processor                -ErrorAction Stop
            $result.ComputerSys = Get-CimInstance -CimSession $session -ClassName Win32_ComputerSystem           -ErrorAction Stop
            $result.PageFile    = Get-CimInstance -CimSession $session -ClassName Win32_PageFileSetting          -ErrorAction SilentlyContinue
            $result.Volumes     = @(Get-CimInstance -CimSession $session -ClassName Win32_Volume -Filter 'DriveType=3' -ErrorAction SilentlyContinue)
            $result.NICs        = @(Get-CimInstance -CimSession $session -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=TRUE' -ErrorAction SilentlyContinue)
            $result.Reachable   = $true
        }
        catch {
            $result.Error = $_.Exception.Message
            Write-MyVerbose ('Get-RemoteServerData {0}: {1}' -f $ComputerName, $_.Exception.Message)
        }
        finally {
            if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
        }

        return $result
    }

    function Invoke-RemoteQueryWithPrompt {
        <#
        .SYNOPSIS
            Wraps Get-RemoteServerData with interactive retry/skip prompt on failure.
        .DESCRIPTION
            Copilot (interactive) mode: on failure, shows hint pointing to Enable-EXpressRemoteQuery.ps1
            and offers [R]etry / [S]kip with a 10-minute auto-skip timeout (Write-Progress -Id 2).
            Autopilot mode or non-interactive session: silent skip.
        #>
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][string]$ComputerName,
            [int]$TimeoutSec = 600
        )

        $data = Get-RemoteServerData -ComputerName $ComputerName
        if ($data.Reachable) { return $data }

        $nonInteractive = $State['Autopilot'] -or -not [Environment]::UserInteractive
        if ($nonInteractive) {
            Write-MyVerbose ('Remote query skipped (non-interactive) for {0}: {1}' -f $ComputerName, $data.Error)
            return $data
        }

        while (-not $data.Reachable) {
            Write-Host ''
            Write-MyWarning ('Remote query failed for {0}' -f $ComputerName)
            Write-Host ('    Error : {0}' -f $data.Error) -ForegroundColor Yellow
            Write-Host  '    Fix   : Run tools\Enable-EXpressRemoteQuery.ps1 on the target server,' -ForegroundColor Yellow
            Write-Host  '            or apply GPO per docs\remote-query-setup.md' -ForegroundColor Yellow
            Write-Host ''
            Write-Host '    [R] Retry    [S] Skip    (auto-skip in 10:00)' -ForegroundColor Cyan

            $choice = $null
            try { $host.UI.RawUI.FlushInputBuffer() } catch { } # intentional: RawUI unavailable in PS2Exe/redirected hosts
            $deadline = [DateTime]::Now.AddSeconds($TimeoutSec)
            while ([DateTime]::Now -lt $deadline -and -not $choice) {
                $secsLeft = [int]($deadline - [DateTime]::Now).TotalSeconds
                $mm = [int]($secsLeft / 60); $ss = $secsLeft % 60
                Write-Progress -Id 2 -Activity ('Remote query: {0}' -f $ComputerName) `
                    -Status ('Auto-skip in {0:D2}:{1:D2}  |  [R] Retry  |  [S] Skip' -f $mm, $ss) `
                    -PercentComplete (($TimeoutSec - $secsLeft) * 100 / $TimeoutSec)
                if ($host.UI.RawUI.KeyAvailable) {
                    $key = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                    switch ($key.Character.ToString().ToUpper()) {
                        'R' { $choice = 'Retry' }
                        'S' { $choice = 'Skip'  }
                    }
                }
                Start-Sleep -Milliseconds 100
            }
            Write-Progress -Id 2 -Activity ('Remote query: {0}' -f $ComputerName) -Completed

            if (-not $choice) {
                Write-MyOutput ('Auto-skip: {0}' -f $ComputerName)
                return $data
            }
            if ($choice -eq 'Skip') {
                Write-MyOutput ('Skipped remote query for {0}' -f $ComputerName)
                return $data
            }

            Write-MyOutput ('Retrying remote query for {0}...' -f $ComputerName)
            $data = Get-RemoteServerData -ComputerName $ComputerName
        }

        return $data
    }

