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
                } catch { }
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
                } catch { }
                $org.RoleGroups += [pscustomobject]@{
                    Name        = $rg.Name
                    Description = $rg.Description
                    Members     = $mem
                    ManagedBy   = @($rg.ManagedBy | ForEach-Object { $_.ToString() })
                }
            }
        } catch { }

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
                try { Get-ScheduledTask -TaskName $tn -ErrorAction SilentlyContinue | ForEach-Object { $foundTasks[$_.TaskName] = $_ } } catch { }
            }
            # Broad pattern search — catches variants across CSS-Exchange releases
            try {
                Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
                    $_.TaskName -match 'Daily Auth Certificate|MonitorExchangeAuth|ExchangeLogClean|EXpressLog'
                } | ForEach-Object { $foundTasks[$_.TaskName] = $_ }
            } catch { }
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
        } catch { }

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
            try { $srv.ImapSettings = Get-ImapSettings -Server $ServerName -ErrorAction Stop } catch { }
            try { $srv.PopSettings  = Get-PopSettings  -Server $ServerName -ErrorAction Stop } catch { }
        }

        # Certificates
        try { $srv.Certificates = @(Get-ExchangeCertificate -Server $ServerName -ErrorAction Stop) } catch { $srv.Certificates = @() }

        # Transport agents (only present on servers with Hub Transport)
        try { $srv.TransportAgents = @(Get-TransportAgent -ErrorAction Stop) } catch { $srv.TransportAgents = @() }

        # Database copy status (per server; runs against local server where available)
        try { $srv.DatabaseCopies = @(Get-MailboxDatabaseCopyStatus -Server $ServerName -ErrorAction Stop |
                                      Select-Object Name, DatabaseName, Status, ContentIndexState, CopyQueueLength, ReplayQueueLength, ActivationPreference, MailboxServer) } catch { $srv.DatabaseCopies = @() }

        # Defender preferences — only meaningful for local server (remote would need CIM/PSSession)
        $srv.DefenderExclusions = $null
        if ($ServerName -ieq $env:COMPUTERNAME) {
            try {
                $mp = Get-MpPreference -ErrorAction Stop
                $srv.DefenderExclusions = [pscustomobject]@{
                    ExclusionPath      = @($mp.ExclusionPath)
                    ExclusionProcess   = @($mp.ExclusionProcess)
                    ExclusionExtension = @($mp.ExclusionExtension)
                    RealTimeEnabled    = -not $mp.DisableRealtimeMonitoring
                }
            } catch { }
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
            } catch { }
        }

        # Hardware/OS data — direct CIM for local server, CIM/WSMan + prompt for remote
        if ($ServerName -ieq $env:COMPUTERNAME) {
            $srv.RemoteData = @{
                ComputerName = $ServerName
                Reachable    = $true
                Error        = $null
                OS           = Get-CimInstance Win32_OperatingSystem          -ErrorAction SilentlyContinue
                CPU          = Get-CimInstance Win32_Processor                 -ErrorAction SilentlyContinue
                ComputerSys  = Get-CimInstance Win32_ComputerSystem            -ErrorAction SilentlyContinue
                PageFile     = Get-CimInstance Win32_PageFileSetting           -ErrorAction SilentlyContinue
                Volumes      = @(Get-CimInstance Win32_Volume -Filter 'DriveType=3'                          -ErrorAction SilentlyContinue)
                NICs         = @(Get-CimInstance Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=TRUE'  -ErrorAction SilentlyContinue)
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

