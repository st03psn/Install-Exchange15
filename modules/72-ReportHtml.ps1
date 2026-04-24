    function New-InstallationReport {
        Write-MyOutput 'Generating Installation Report'
        $reportPath = Join-Path $State['ReportsPath'] ('{0}_EXpress_Report_{1}.html' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd-HHmmss'))
        $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

        function Format-Badge($text, $type) {
            $colors = @{ ok='background:#107c10;color:#fff'; warn='background:#d83b01;color:#fff'; fail='background:#c50f1f;color:#fff'; info='background:#0078d4;color:#fff'; na='background:#8a8886;color:#fff' }
            '<span style="display:inline-block;padding:2px 10px;border-radius:12px;font-size:12px;font-weight:600;' + $colors[$type.ToLower()] + '">' + $text + '</span>'
        }
        function New-HtmlSection($id, $title, $content) {
            '<section id="' + $id + '" class="section"><h2 class="section-title">' + $title + '</h2><div class="section-body">' + $content + '</div></section>'
        }

        # ── 1. INSTALLATION PARAMETERS ────────────────────────────────────────
        $instRows = [System.Collections.Generic.List[string]]::new()
        $instMode = if ($State['InstallEdge']) { 'Edge Transport' } elseif ($State['InstallRecipientManagement']) { 'Recipient Management Tools' } elseif ($State['InstallManagementTools']) { 'Management Tools' } elseif ($State['StandaloneOptimize']) { 'Standalone Optimize' } elseif ($State['NoSetup']) { 'Optimization Only' } else { 'Mailbox Server' }
        $instRows.Add('<tr><td>Script Version</td><td>EXpress v{0}</td></tr>' -f $ScriptVersion)
        $instRows.Add('<tr><td>Report Generated</td><td>{0}</td></tr>' -f $reportDate)
        $instRows.Add('<tr><td>Server</td><td>{0}</td></tr>' -f $env:COMPUTERNAME)
        $instRows.Add('<tr><td>Installation Mode</td><td>{0}</td></tr>' -f $instMode)
        $instRows.Add('<tr><td>Organization</td><td>{0}</td></tr>' -f $State['OrganizationName'])
        $instRows.Add(('<tr><td>Setup Version</td><td>{0} ({1})</td></tr>' -f $State['SetupVersion'], (Get-SetupTextVersion $State['SetupVersion'])))
        $instRows.Add('<tr><td>Install Path</td><td>{0}</td></tr>' -f $State['InstallPath'])
        if ($State['Namespace'])        { $instRows.Add('<tr><td>Namespace</td><td>{0}</td></tr>' -f $State['Namespace']) }
        if ($State['DownloadDomain'])   { $instRows.Add('<tr><td>OWA Download Domain</td><td>{0}</td></tr>' -f $State['DownloadDomain']) }
        if ($State['DAGName'])          { $instRows.Add('<tr><td>DAG</td><td>{0}</td></tr>' -f $State['DAGName']) }
        if ($State['CertificatePath'])  { $instRows.Add('<tr><td>Certificate Path</td><td>{0}</td></tr>' -f $State['CertificatePath']) }
        if ($State['CopyServerConfig']) { $instRows.Add('<tr><td>Source Server Config</td><td>{0}</td></tr>' -f $State['CopyServerConfig']) }
        if ($State['LogRetentionDays']) { $instRows.Add('<tr><td>Log Retention</td><td>{0} days</td></tr>' -f $State['LogRetentionDays']) }
        $modeText = if ($State['ConfigDriven']) { 'Autopilot (fully automated)' } else { 'Copilot (interactive)' }
        $instRows.Add('<tr><td>Mode</td><td>{0}</td></tr>' -f $modeText)
        $instRows.Add('<tr><td>TLS 1.2 Enforced</td><td>{0}</td></tr>' -f $State['EnableTLS12'])
        $instRows.Add('<tr><td>TLS 1.3 Enforced</td><td>{0}</td></tr>' -f $State['EnableTLS13'])
        $instRows.Add('<tr><td>SSL 3 Disabled</td><td>{0}</td></tr>' -f $State['DisableSSL3'])
        $instRows.Add('<tr><td>RC4 Disabled</td><td>{0}</td></tr>' -f $State['DisableRC4'])
        $instRows.Add('<tr><td>Log File</td><td><code>{0}</code></td></tr>' -f $State['TranscriptFile'])

        # ── 2. SYSTEM INFORMATION ─────────────────────────────────────────────
        $sysRows = [System.Collections.Generic.List[string]]::new()
        try {
            $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
            $sysRows.Add(('<tr><td>Operating System</td><td>{0}</td><td>{1}</td></tr>' -f $os.Caption, (Format-Badge 'OK' 'ok')))
            $sysRows.Add('<tr><td>OS Build</td><td>{0}</td><td></td></tr>' -f $os.Version)
            $sysRows.Add('<tr><td>Last Boot</td><td>{0}</td><td></td></tr>' -f $os.LastBootUpTime.ToString('yyyy-MM-dd HH:mm:ss'))
            $totalRAM = [math]::Round($os.TotalVisibleMemorySize / 1MB, 0)
            $sysRows.Add('<tr><td>Total RAM</td><td>{0} GB</td><td></td></tr>' -f $totalRAM)
        } catch { $sysRows.Add('<tr><td colspan="3">OS info unavailable</td></tr>') }
        try {
            $cpuList     = @(Get-CimInstance Win32_Processor -ErrorAction Stop)
            $cpu         = $cpuList[0]
            $totalCores  = ($cpuList | Measure-Object -Property NumberOfCores -Sum).Sum
            $totalLogical= ($cpuList | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
            $sysRows.Add(('<tr><td>CPU</td><td>{0}</td><td>{1} cores / {2} logical</td></tr>' -f $cpu.Name.Trim(), $totalCores, $totalLogical))
        } catch { }
        try {
            $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
            $sysRows.Add(('<tr><td>Computer Name (FQDN)</td><td>{0}.{1}</td><td></td></tr>' -f $cs.DNSHostName, $cs.Domain))
        } catch { }
        # Volumes — exclude DVD-ROM and removable drives
        try {
            Get-Volume -ErrorAction SilentlyContinue | Where-Object {
                $_.DriveLetter -and $_.DriveType -notin 'CD-ROM','Removable' -and $_.Size -gt 0
            } | ForEach-Object {
                $freeGB    = [math]::Round($_.SizeRemaining / 1GB, 1)
                $totalGB   = [math]::Round($_.Size / 1GB, 1)
                $freePct   = [math]::Round($_.SizeRemaining / $_.Size * 100, 0)
                $auBadge   = if ($_.AllocationUnitSize -eq 65536) { Format-Badge '64 KB ✓' 'ok' } elseif ($_.AllocationUnitSize) { Format-Badge ('{0} KB !' -f ($_.AllocationUnitSize / 1KB)) 'warn' } else { '' }
                $diskBadge = if ($freePct -lt 10) { Format-Badge ('Free {0}%' -f $freePct) 'fail' } elseif ($freePct -lt 20) { Format-Badge ('Free {0}%' -f $freePct) 'warn' } else { Format-Badge ('Free {0}%' -f $freePct) 'ok' }
                $sysRows.Add(('<tr><td>Volume {0}:</td><td>{1} GB free of {2} GB &nbsp; {3}</td><td>{4} &nbsp; Alloc: {5}</td></tr>' -f $_.DriveLetter, $freeGB, $totalGB, $diskBadge, $_.FileSystem, $auBadge))
            }
        } catch { }
        # Network adapters — IP address + DNS servers
        try {
            $nicIPs  = @{}
            $nicDNS  = @{}
            Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                Where-Object { $_.InterfaceAlias -notlike '*Loopback*' } |
                ForEach-Object { $nicIPs[$_.InterfaceAlias] = ('{0}/{1}' -f $_.IPAddress, $_.PrefixLength) }
            Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                Where-Object { $_.InterfaceAlias -notlike '*Loopback*' -and $_.ServerAddresses } |
                ForEach-Object { $nicDNS[$_.InterfaceAlias] = ($_.ServerAddresses -join ', ') }
            foreach ($nic in ($nicIPs.Keys | Sort-Object)) {
                $dns = if ($nicDNS[$nic]) { $nicDNS[$nic] } else { '<em>not set</em>' }
                $sysRows.Add(('<tr><td>NIC: {0}</td><td>{1}</td><td>DNS: {2}</td></tr>' -f $nic, $nicIPs[$nic], $dns))
            }
        } catch { }
        $sysContent = '<table class="data-table"><tr><th>Property</th><th>Value</th><th>Detail / Status</th></tr>' + ($sysRows -join '') + '</table>'

        # ── 3. ACTIVE DIRECTORY ───────────────────────────────────────────────
        $adRows = [System.Collections.Generic.List[string]]::new()
        try {
            $cs2 = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
            $adRows.Add('<tr><td>Domain</td><td>{0}</td><td></td></tr>' -f $cs2.Domain)
        } catch { }
        try {
            $ffl = Get-ForestFunctionalLevel
            $fflBadge = if ($ffl -ge $FOREST_LEVEL2012R2) { Format-Badge 'OK' 'ok' } else { Format-Badge 'WARN' 'warn' }
            $adRows.Add(('<tr><td>Forest Functional Level</td><td>{0} ({1})</td><td>{2}</td></tr>' -f $ffl, (Get-FFLText $ffl), $fflBadge))
        } catch { }
        try {
            $exOrg = Get-ExchangeOrganization
            if ($exOrg) { $adRows.Add('<tr><td>Exchange Organization</td><td>{0}</td><td></td></tr>' -f $exOrg) }
            $exFL = Get-ExchangeForestLevel
            $adRows.Add('<tr><td>Exchange Forest Schema</td><td>{0}</td><td></td></tr>' -f $exFL)
            $exDL = Get-ExchangeDomainLevel
            $adRows.Add('<tr><td>Exchange Domain Level</td><td>{0}</td><td></td></tr>' -f $exDL)
        } catch { }
        $adContent = '<table class="data-table"><tr><th>Property</th><th>Value</th><th>Status</th></tr>' + ($adRows -join '') + '</table>'

        # ── 4. EXCHANGE CONFIGURATION ─────────────────────────────────────────
        $exRows    = [System.Collections.Generic.List[string]]::new()
        $vdirRows  = [System.Collections.Generic.List[string]]::new()
        $connRows  = [System.Collections.Generic.List[string]]::new()
        $dbRows    = [System.Collections.Generic.List[string]]::new()
        $certRows  = [System.Collections.Generic.List[string]]::new()
        $exVersion = Get-SetupTextVersion $State['SetupVersion']

        try {
            $exSrv = Get-ExchangeServer $env:COMPUTERNAME -ErrorAction Stop
            $exVersion = $exSrv.AdminDisplayVersion.ToString()
            $exRows.Add('<tr><td>Exchange Version</td><td>{0}</td><td></td></tr>' -f $exSrv.AdminDisplayVersion)
            $exRows.Add('<tr><td>Server Role</td><td>{0}</td><td></td></tr>' -f ($exSrv.ServerRole -join ', '))
            $exRows.Add('<tr><td>Edition</td><td>{0}</td><td></td></tr>' -f $exSrv.Edition)
            $exRows.Add('<tr><td>AD Site</td><td>{0}</td><td></td></tr>' -f $exSrv.Site)
        } catch { $exRows.Add('<tr><td colspan="3">Exchange server query unavailable</td></tr>') }
        # Autodiscover SCP (Client Access Service, not a virtual directory)
        try {
            $cas = Get-ClientAccessService -Identity $env:COMPUTERNAME -ErrorAction SilentlyContinue
            if ($cas) { $exRows.Add('<tr><td>Autodiscover SCP</td><td>{0}</td><td></td></tr>' -f $cas.AutoDiscoverServiceInternalUri) }
        } catch { }
        try {
            $orgCfg = Get-OrganizationConfig -ErrorAction Stop
            $exRows.Add('<tr><td>Organization Name</td><td>{0}</td><td></td></tr>' -f $orgCfg.Name)
            $oauthBadge = if ($orgCfg.OAuth2ClientProfileEnabled) { Format-Badge 'Enabled' 'ok' } else { Format-Badge 'Disabled' 'warn' }
            $exRows.Add(('<tr><td>Modern Auth (OAuth2)</td><td>{0}</td><td>{1}</td></tr>' -f $orgCfg.OAuth2ClientProfileEnabled, $oauthBadge))
            $mapiBadge = if ($orgCfg.MapiHttpEnabled) { Format-Badge 'Enabled' 'ok' } else { Format-Badge 'Disabled' 'warn' }
            $exRows.Add(('<tr><td>MAPI/HTTP</td><td>{0}</td><td>{1}</td></tr>' -f $orgCfg.MapiHttpEnabled, $mapiBadge))
        } catch { }
        try {
            Get-AcceptedDomain -ErrorAction Stop | ForEach-Object {
                $exRows.Add(('<tr><td>Accepted Domain</td><td>{0}</td><td>{1}</td></tr>' -f $_.DomainName, (Format-Badge $_.DomainType 'info')))
            }
        } catch { }

        # Virtual directories — Autodiscover SCP + OWA, ECP, EWS, OAB, ActiveSync, MAPI
        $vdirRows.Add('<tr><th>Service</th><th>Internal URL</th><th>External URL</th></tr>')
        try {
            $cas = Get-ClientAccessService -Identity $env:COMPUTERNAME -ErrorAction Stop
            $scpUri = if ($cas.AutoDiscoverServiceInternalUri) { $cas.AutoDiscoverServiceInternalUri.AbsoluteUri } else { '<em>not set</em>' }
            $vdirRows.Add(('<tr><td>Autodiscover SCP</td><td>{0}</td><td><em>n/a (SCP)</em></td></tr>' -f $scpUri))
        } catch { }
        @(
            @{ Name='OWA';         Cmd={ Get-OwaVirtualDirectory         -Server $env:COMPUTERNAME -ADPropertiesOnly -ErrorAction SilentlyContinue | Select-Object -First 1 } }
            @{ Name='ECP';         Cmd={ Get-EcpVirtualDirectory         -Server $env:COMPUTERNAME -ADPropertiesOnly -ErrorAction SilentlyContinue | Select-Object -First 1 } }
            @{ Name='EWS';         Cmd={ Get-WebServicesVirtualDirectory  -Server $env:COMPUTERNAME -ADPropertiesOnly -ErrorAction SilentlyContinue | Select-Object -First 1 } }
            @{ Name='OAB';         Cmd={ Get-OabVirtualDirectory         -Server $env:COMPUTERNAME -ADPropertiesOnly -ErrorAction SilentlyContinue | Select-Object -First 1 } }
            @{ Name='ActiveSync';  Cmd={ Get-ActiveSyncVirtualDirectory   -Server $env:COMPUTERNAME -ADPropertiesOnly -ErrorAction SilentlyContinue | Select-Object -First 1 } }
            @{ Name='MAPI';        Cmd={ Get-MapiVirtualDirectory         -Server $env:COMPUTERNAME -ADPropertiesOnly -ErrorAction SilentlyContinue | Select-Object -First 1 } }
        ) | ForEach-Object {
            try {
                $vd = & $_.Cmd
                if ($vd) {
                    $int = if ($vd.InternalUrl) { $vd.InternalUrl.AbsoluteUri } else { '<em>not set</em>' }
                    $ext = if ($vd.ExternalUrl) { $vd.ExternalUrl.AbsoluteUri } else { '<em>not set</em>' }
                    $vdirRows.Add(('<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>' -f $_.Name, $int, $ext))
                }
            } catch { }
        }

        # Mailbox databases — try with status, fall back without
        $dbRows.Add('<tr><th>Database</th><th>DB Path</th><th>Log Path</th><th>Status</th></tr>')
        try {
            $dbs = Get-MailboxDatabase -Server $env:COMPUTERNAME -Status -ErrorAction SilentlyContinue
            if (-not $dbs) {
                $dbs = Get-MailboxDatabase -Server $env:COMPUTERNAME -ErrorAction Stop
            }
            if ($dbs) {
                $dbs | ForEach-Object {
                    $mounted = if ($null -ne $_.Mounted) { $_.Mounted } else { $null }
                    $mountedText  = if ($null -eq $mounted) { 'Unknown' } elseif ($mounted) { 'Mounted' } else { 'Dismounted' }
                    $mountedBadge = if ($null -eq $mounted) { Format-Badge 'Unknown' 'na' } elseif ($mounted) { Format-Badge 'Mounted' 'ok' } else { Format-Badge 'Dismounted' 'warn' }
                    $dbRows.Add(('<tr><td>{0}</td><td><code>{1}</code></td><td><code>{2}</code></td><td>{3}</td></tr>' -f $_.Name, $_.EdbFilePath, $_.LogFolderPath, $mountedBadge))
                }
            } else {
                $dbRows.Add('<tr><td colspan="4"><em>No mailbox databases found on this server</em></td></tr>')
            }
        } catch { $dbRows.Add(('<tr><td colspan="4">Query failed: {0}</td></tr>' -f $_.Exception.Message)) }

        # Receive connectors
        $connRows.Add('<tr><th>Connector</th><th>Bindings</th><th>Remote IP Ranges</th><th>Auth</th><th>Permission Groups</th></tr>')
        try {
            Get-ReceiveConnector -Server $env:COMPUTERNAME -ErrorAction Stop | ForEach-Object {
                $connRows.Add(('<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td></tr>' -f $_.Name, ($_.Bindings -join '<br>'), ($_.RemoteIPRanges -join '<br>'), $_.AuthMechanism, $_.PermissionGroups))
            }
        } catch { $connRows.Add('<tr><td colspan="5">Receive connector query failed</td></tr>') }

        # Certificates
        $certRows.Add('<tr><th>Subject</th><th>Expiry</th><th>Services</th><th>Thumbprint</th></tr>')
        try {
            Get-ExchangeCertificate -Server $env:COMPUTERNAME -ErrorAction Stop | ForEach-Object {
                # Skip phantom entries: no subject, no thumbprint, or NotAfter = DateTime.MinValue
                if ([string]::IsNullOrEmpty($_.Thumbprint) -or $_.NotAfter -le [datetime]'1970-01-01') { return }
                $daysLeft = [int][Math]::Floor(($_.NotAfter - (Get-Date)).TotalDays)
                $expiryBadge = if ($daysLeft -lt 30) { Format-Badge ('Expires {0}d!' -f $daysLeft) 'fail' } elseif ($daysLeft -lt 90) { Format-Badge ('Expires {0}d' -f $daysLeft) 'warn' } else { Format-Badge ('{0} ({1}d)' -f $_.NotAfter.ToString('yyyy-MM-dd'), $daysLeft) 'ok' }
                $certRows.Add(('<tr><td>{0}</td><td>{1}</td><td>{2}</td><td><code>{3}</code></td></tr>' -f $_.Subject, $expiryBadge, $_.Services, $_.Thumbprint))
            }
        } catch { $certRows.Add('<tr><td colspan="4">Certificate query failed</td></tr>') }

        # Exchange Optimizations — org/transport level settings
        $exchOptRows = [System.Collections.Generic.List[string]]::new()
        $exchOptRows.Add('<tr><th>Setting</th><th>Current Value</th><th>Recommendation</th><th>Status</th></tr>')
        try {
            $orgCfg2 = Get-OrganizationConfig -ErrorAction SilentlyContinue
            if ($orgCfg2) {
                $toEnabled  = $orgCfg2.ActivityBasedAuthenticationTimeoutEnabled
                $toInterval = $orgCfg2.ActivityBasedAuthenticationTimeoutInterval
                $toBadge    = if ($toEnabled -and $toInterval -le [TimeSpan]'06:00:00') { Format-Badge '✓' 'ok' } else { Format-Badge 'Review' 'warn' }
                $exchOptRows.Add(('<tr><td>OWA/ECP Session Timeout</td><td>Enabled: {0} / Interval: {1}</td><td>Enabled, ≤ 6 h (security compliance)</td><td>{2}</td></tr>' -f $toEnabled, $toInterval, $toBadge))

                $ceipBadge = if (-not $orgCfg2.CustomerFeedbackEnabled) { Format-Badge 'Disabled ✓' 'ok' } else { Format-Badge 'Enabled' 'warn' }
                $exchOptRows.Add(('<tr><td>CEIP / Telemetry</td><td>{0}</td><td>Disabled (privacy / GDPR)</td><td>{1}</td></tr>' -f $orgCfg2.CustomerFeedbackEnabled, $ceipBadge))
            }
        } catch { }
        try {
            $transCfg2 = Get-TransportConfig -ErrorAction SilentlyContinue
            if ($transCfg2) {
                # MaxSendSize / MaxReceiveSize may be Unlimited ($null .Value) on fresh orgs.
                $maxSendMB = if ($transCfg2.MaxSendSize    -and $transCfg2.MaxSendSize.Value)    { [math]::Round($transCfg2.MaxSendSize.Value.ToBytes()    / 1MB, 0) } else { $null }
                $maxRecvMB = if ($transCfg2.MaxReceiveSize -and $transCfg2.MaxReceiveSize.Value) { [math]::Round($transCfg2.MaxReceiveSize.Value.ToBytes() / 1MB, 0) } else { $null }
                $maxSendDisp = if ($null -ne $maxSendMB) { ('{0} MB' -f $maxSendMB) } else { 'Unlimited / not set' }
                $maxRecvDisp = if ($null -ne $maxRecvMB) { ('{0} MB' -f $maxRecvMB) } else { 'Unlimited / not set' }
                $sizeBadge = if ($null -ne $maxSendMB -and $maxSendMB -ge 50) { Format-Badge '✓' 'ok' } else { Format-Badge 'Default 25 MB' 'warn' }
                $exchOptRows.Add(('<tr><td>Max Message Size</td><td>Send: {0} / Recv: {1}</td><td>≥ 50 MB (modern workflow files)</td><td>{2}</td></tr>' -f $maxSendDisp, $maxRecvDisp, $sizeBadge))

                $ndrBadge = if ($transCfg2.InternalDsnSendHtml -and $transCfg2.ExternalDsnSendHtml) { Format-Badge 'Enabled ✓' 'ok' } else { Format-Badge 'Plain text' 'warn' }
                $exchOptRows.Add(('<tr><td>HTML Non-Delivery Reports</td><td>Internal: {0} / External: {1}</td><td>Enabled (improves end-user NDR readability)</td><td>{2}</td></tr>' -f $transCfg2.InternalDsnSendHtml, $transCfg2.ExternalDsnSendHtml, $ndrBadge))

                $snBadge = if ($transCfg2.SafetyNetHoldTime -ge [TimeSpan]'2.00:00:00') { Format-Badge '✓' 'ok' } else { Format-Badge 'Short' 'warn' }
                $exchOptRows.Add(('<tr><td>Safety Net Hold Time</td><td>{0}</td><td>≥ 2 days (message redelivery after DB failover)</td><td>{1}</td></tr>' -f $transCfg2.SafetyNetHoldTime, $snBadge))
            }
        } catch { }
        try {
            $transSvc2 = Get-TransportService -Identity $env:COMPUTERNAME -ErrorAction SilentlyContinue
            if ($transSvc2) {
                $expBadge = if ($transSvc2.MessageExpirationTimeout -ge [TimeSpan]'7.00:00:00') { Format-Badge '✓' 'ok' } else { Format-Badge 'Default 2d' 'warn' }
                $exchOptRows.Add(('<tr><td>Message Expiration Timeout</td><td>{0}</td><td>7 days (delay NDRs during multi-day outages)</td><td>{1}</td></tr>' -f $transSvc2.MessageExpirationTimeout, $expBadge))
            }
        } catch { }

        $exContent = '<table class="data-table"><tr><th>Property</th><th>Value</th><th>Status</th></tr>' + ($exRows -join '') + '</table>' +
            '<h3 class="subsection">Virtual Directory URLs</h3><table class="data-table">' + ($vdirRows -join '') + '</table>' +
            '<h3 class="subsection">Mailbox Databases</h3><table class="data-table">' + ($dbRows -join '') + '</table>' +
            '<h3 class="subsection">Receive Connectors</h3><table class="data-table">' + ($connRows -join '') + '</table>' +
            '<h3 class="subsection">Certificates</h3><table class="data-table">' + ($certRows -join '') + '</table>' +
            '<h3 class="subsection">Exchange Optimizations</h3><table class="data-table">' + ($exchOptRows -join '') + '</table>'

        # ── 5. SECURITY SETTINGS (with Exchange best-practice + reference column) ─
        $secRows = [System.Collections.Generic.List[string]]::new()
        function Get-SecRegVal($path, $name) { try { (Get-ItemProperty -Path $path -Name $name -ErrorAction Stop).$name } catch { $null } }
        function Format-RefLink($url, $label) { '<a href="' + $url + '" target="_blank" style="font-size:0.85em;white-space:nowrap">' + $label + ' ↗</a>' }

        # TLS protocols — show current value + Exchange recommendation
        @(
            @{ Proto='1.0'; Rec='Disabled'; LegacyRisk=$true;  CisId='CIS L1 / PCI-DSS 4.2.1'; RefUrl='https://techcommunity.microsoft.com/t5/exchange-team-blog/exchange-server-tls-guidance-part-1-getting-ready-for-tls-1-2/ba-p/607649'; RefLabel='Exchange TLS Guide' }
            @{ Proto='1.1'; Rec='Disabled'; LegacyRisk=$true;  CisId='CIS L1 / PCI-DSS 4.2.1'; RefUrl='https://techcommunity.microsoft.com/t5/exchange-team-blog/exchange-server-tls-guidance-part-1-getting-ready-for-tls-1-2/ba-p/607649'; RefLabel='Exchange TLS Guide' }
            @{ Proto='1.2'; Rec='Enabled';  LegacyRisk=$false; CisId='CIS L1 / PCI-DSS 4.2.1'; RefUrl='https://techcommunity.microsoft.com/blog/exchange/exchange-server-tls-guidance-part-2-enabling-tls-1-2-and-identifying-clients-not/607761'; RefLabel='Exchange TLS Guide' }
            @{ Proto='1.3'; Rec='Enabled (Exchange SE / 2019 CU15+ on WS2022+)'; LegacyRisk=$false; CisId='Best practice'; RefUrl='https://support.microsoft.com/en-us/topic/partial-tls-1-3-support-for-exchange-server-2019-5f4058f5-b288-4859-9a85-9aac680f50fe'; RefLabel='Exchange Blog' }
        ) | ForEach-Object {
            $proto      = $_.Proto
            $srvEnabled = Get-SecRegVal "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS $proto\Server" 'Enabled'
            $cliEnabled = Get-SecRegVal "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS $proto\Client" 'Enabled'
            $isEnabled  = -not ($srvEnabled -eq 0 -or $cliEnabled -eq 0)
            $valText    = if ($null -eq $srvEnabled -and $null -eq $cliEnabled) { 'OS Default' } else { "Srv=$srvEnabled / Cli=$cliEnabled" }
            $label      = if ($isEnabled) { 'Enabled' } else { 'Disabled' }
            $badgeType  = if ($_.LegacyRisk) { if ($isEnabled) { 'warn' } else { 'ok' } } else { if ($isEnabled) { 'ok' } else { 'warn' } }
            $secRows.Add(('<tr><td>TLS {0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td></tr>' -f $proto, $valText, $_.Rec, (Format-Badge $label $badgeType), (Format-RefLink $_.RefUrl $_.RefLabel), $_.CisId))
        }
        $strongCrypto = Get-SecRegVal 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319' 'SchUseStrongCrypto'
        $strongBadge  = if ($strongCrypto -eq 1) { Format-Badge 'Enabled' 'ok' } else { Format-Badge 'Not set' 'warn' }
        $secRows.Add(('<tr><td>.NET Strong Crypto</td><td>SchUseStrongCrypto = {0}</td><td>1 (required)</td><td>{1}</td><td>{2}</td><td>CIS L1 §18.3</td></tr>' -f $strongCrypto, $strongBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/dotnet/framework/network-programming/tls' 'MS Learn')))
        try {
            $smb1 = (Get-SmbServerConfiguration -ErrorAction Stop).EnableSMB1Protocol
            $smb1Badge = if ($smb1) { Format-Badge 'Enabled (risk)' 'warn' } else { Format-Badge 'Disabled' 'ok' }
            $secRows.Add(('<tr><td>SMBv1</td><td>{0}</td><td>Disabled</td><td>{1}</td><td>{2}</td><td>CIS L1 §18.3 / BSI SYS.1</td></tr>' -f $smb1, $smb1Badge, (Format-RefLink 'https://techcommunity.microsoft.com/t5/storage-at-microsoft/stop-using-smb1/ba-p/425858' 'Microsoft Blog')))
        } catch { }
        $wdigest = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest' 'UseLogonCredential'
        $wdigestBadge = if ($wdigest -eq 0) { Format-Badge 'Disabled' 'ok' } else { Format-Badge 'Enabled (risk)' 'warn' }
        $secRows.Add(('<tr><td>WDigest Caching</td><td>UseLogonCredential = {0}</td><td>0 = Disabled</td><td>{1}</td><td>{2}</td><td>CIS L1 §18.9.48 / DISA STIG</td></tr>' -f $wdigest, $wdigestBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/windows-server/security/credentials-protection-and-management/configuring-additional-lsa-protection' 'MS Learn')))
        $lsaPPL = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' 'RunAsPPL'
        $lsaBadge = if ($lsaPPL -eq 1) { Format-Badge 'Enabled' 'ok' } else { Format-Badge 'Not set' 'warn' }
        $secRows.Add(('<tr><td>LSA Protection (RunAsPPL)</td><td>{0}</td><td>1 = Enabled (Ex2019 CU12+/SE)</td><td>{1}</td><td>{2}</td><td>CIS L2 §2.3.11 / DISA STIG</td></tr>' -f $lsaPPL, $lsaBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/windows-server/security/credentials-protection-and-management/configuring-additional-lsa-protection' 'MS Learn')))
        $lmLevel = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' 'LmCompatibilityLevel'
        $lmBadge = if ($lmLevel -ge 5) { Format-Badge "Level $lmLevel ✓" 'ok' } else { Format-Badge "Level $lmLevel" 'warn' }
        $secRows.Add(('<tr><td>LM Compatibility Level</td><td>{0}</td><td>5 = NTLMv2 only</td><td>{1}</td><td>{2}</td><td>CIS L1 §2.3.11.7 / BSI</td></tr>' -f $lmLevel, $lmBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/windows/security/threat-protection/security-policy-settings/network-security-lan-manager-authentication-level' 'MS Learn')))
        $credGuard = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard' 'EnableVirtualizationBasedSecurity'
        $cgBadge = if ($credGuard -eq 0) { Format-Badge 'Disabled' 'ok' } else { Format-Badge 'Enabled (review)' 'warn' }
        $secRows.Add(('<tr><td>Credential Guard</td><td>EnableVBS = {0}</td><td>0 = Disabled (Exchange servers)</td><td>{1}</td><td>{2}</td><td>CIS L2</td></tr>' -f $credGuard, $cgBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/exchange/plan-and-deploy/virtualization' 'Exchange Virtualization')))
        $http2 = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\HTTP\Parameters' 'EnableHttp2Tls'
        $http2Badge = if ($http2 -eq 0) { Format-Badge 'Disabled' 'ok' } else { Format-Badge 'Enabled' 'warn' }
        $secRows.Add(('<tr><td>HTTP/2 over TLS</td><td>EnableHttp2Tls = {0}</td><td>0 = Disabled (MAPI/RPC compat)</td><td>{1}</td><td>{2}</td><td>MS Exchange</td></tr>' -f $http2, $http2Badge, (Format-RefLink 'https://techcommunity.microsoft.com/blog/exchange/released-2022-h1-cumulative-updates-for-exchange-server/3285026' 'Exchange Blog')))
        # Serialized Data Signing — registry value name: EnableSerializationDataSigning
        $serialSign = Get-SecRegVal 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Diagnostics' 'EnableSerializationDataSigning'
        $serialBadge = if ($serialSign -eq 1) { Format-Badge 'Enabled' 'ok' } else { Format-Badge 'Not set' 'warn' }
        $secRows.Add(('<tr><td>Serialized Data Signing</td><td>{0}</td><td>1 = Enabled</td><td>{1}</td><td>{2}</td><td>MS Exchange</td></tr>' -f $serialSign, $serialBadge, (Format-RefLink 'https://techcommunity.microsoft.com/blog/exchange/released-2022-h1-cumulative-updates-for-exchange-server/3285026' 'Exchange Blog')))
        $uacVal = Get-SecRegVal 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' 'EnableLUA'
        $uacBadge = if ($uacVal -eq 1 -or $null -eq $uacVal) { Format-Badge 'Enabled' 'ok' } else { Format-Badge 'Disabled!' 'fail' }
        $secRows.Add(('<tr><td>UAC (EnableLUA)</td><td>{0}</td><td>1 = Enabled (re-enabled after setup)</td><td>{1}</td><td>{2}</td><td>CIS L1 §17.2 / BSI SYS.2.1</td></tr>' -f $uacVal, $uacBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/windows/security/application-security/application-control/user-account-control/' 'MS Learn')))

        # IPv4 over IPv6 preference
        $ipv4Comp = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters' 'DisabledComponents'
        $ipv4ValText = if ($null -eq $ipv4Comp) { 'Not set (OS default)' } else { '0x{0:X}' -f [int]$ipv4Comp }
        $ipv4Badge = if ($ipv4Comp -eq 0x20) { Format-Badge '0x20 ✓' 'ok' } else { Format-Badge 'Not configured' 'warn' }
        $secRows.Add(('<tr><td>IPv4 over IPv6 preference</td><td>{0}</td><td>0x20 = prefer IPv4 (keep IPv6 loopback)</td><td>{1}</td><td>{2}</td><td>MS Exchange</td></tr>' -f $ipv4ValText, $ipv4Badge, (Format-RefLink 'https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/configure-ipv6-in-windows' 'Exchange Blog')))

        # NetBIOS over TCP/IP
        # SetTcpipNetbios(2) may return 1 (pending reboot), leaving the live CIM value stale.
        # Cross-check the registry (NetbiosOptions=2) so the report reflects the pending state.
        try {
            $nbNics = @(Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True' -ErrorAction Stop)
            $nbDisabled = 0
            foreach ($nbNic in $nbNics) {
                if ($nbNic.TcpipNetbiosOptions -eq 2) { $nbDisabled++; continue }
                $nbReg = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\NetBT\Parameters\Interfaces\Tcpip_$($nbNic.SettingID)" -Name 'NetbiosOptions' -ErrorAction SilentlyContinue).NetbiosOptions
                if ($nbReg -eq 2) { $nbDisabled++ }
            }
            $nbBadge = if ($nbNics.Count -gt 0 -and $nbDisabled -eq $nbNics.Count) { Format-Badge 'Disabled ✓' 'ok' } else { Format-Badge ('{0}/{1} disabled' -f $nbDisabled, $nbNics.Count) 'warn' }
            $secRows.Add(('<tr><td>NetBIOS over TCP/IP</td><td>{0} of {1} NICs disabled</td><td>Disabled on all NICs (reduces LLMNR/NBT-NS attack surface)</td><td>{2}</td><td>{3}</td><td>CIS §18 / BSI</td></tr>' -f $nbDisabled, $nbNics.Count, $nbBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/disable-netbios-tcp-ip-using-dhcp' 'MS Learn')))
        } catch { }

        # Root CA auto-update
        $rootAU = Get-SecRegVal 'HKLM:\SOFTWARE\Policies\Microsoft\SystemCertificates\AuthRoot' 'DisableRootAutoUpdate'
        $rootAUBadge = if ($rootAU -ne 1) { Format-Badge 'Enabled ✓' 'ok' } else { Format-Badge 'Disabled by policy!' 'warn' }
        $rootAUDisplay = if ($null -eq $rootAU) { '(not set — default enabled)' } else { 'DisableRootAutoUpdate = {0}' -f $rootAU }
        $secRows.Add(('<tr><td>Root CA Auto-Update</td><td>{0}</td><td>Must not be disabled (Exchange Online / M365 connectivity)</td><td>{1}</td><td>{2}</td><td>MS Exchange</td></tr>' -f $rootAUDisplay, $rootAUBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/security/trusted-root/release-notes' 'MS Trusted Root')))

        # Extended Protection (OWA VDir — evaluate Frontend only; Back End is internal and EP=None by design)
        if (-not $State['InstallEdge']) {
            try {
                $owaVdirs = @(Get-OwaVirtualDirectory -Server $env:COMPUTERNAME -ADPropertiesOnly -ErrorAction SilentlyContinue)
                $owaFe    = $owaVdirs | Where-Object { $_.Name -notlike '*Back End*' -and $_.WebSiteName -notlike '*Back End*' } | Select-Object -First 1
                if (-not $owaFe) { $owaFe = $owaVdirs | Select-Object -First 1 }
                if ($owaFe) {
                    $siteLabel = if ($owaFe.WebSiteName) { $owaFe.WebSiteName } else { $owaFe.Name }
                    $epVal = [string]$owaFe.ExtendedProtectionTokenChecking
                    if ([string]::IsNullOrEmpty($epVal)) { $epVal = 'None' }
                    # Normalize integer forms returned when deserializing AD properties
                    if ($epVal -eq '2') { $epVal = 'Require' } elseif ($epVal -eq '1') { $epVal = 'Allow' } elseif ($epVal -eq '0') { $epVal = 'None' }
                    $epBadge = if ($epVal -in 'Require','Allow') { Format-Badge "$epVal ✓" 'ok' } else { Format-Badge "$epVal (risk)" 'warn' }
                    $secRows.Add(('<tr><td>Extended Protection (OWA)</td><td>{0} — {1}</td><td>Require or Allow</td><td>{2}</td><td>{3}</td><td>MS Exchange</td></tr>' -f $siteLabel, $epVal, $epBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/exchange/plan-and-deploy/post-installation-tasks/security-best-practices/exchange-extended-protection' 'MS Learn')))
                }
            } catch { }
        }

        # SSL Offloading (Outlook Anywhere) — must be off for Extended Protection
        if (-not $State['InstallEdge']) {
            try {
                $oaVdir = Get-OutlookAnywhere -Server $env:COMPUTERNAME -ErrorAction SilentlyContinue
                if ($oaVdir) {
                    $oaBadge = if (-not $oaVdir.SSLOffloading) { Format-Badge 'Disabled ✓' 'ok' } else { Format-Badge 'Enabled (blocks EP)' 'warn' }
                    $secRows.Add(('<tr><td>SSL Offloading (Outlook Anywhere)</td><td>{0}</td><td>False (required for Extended Protection channel binding)</td><td>{1}</td><td>{2}</td><td>MS Exchange</td></tr>' -f $oaVdir.SSLOffloading, $oaBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/exchange/plan-and-deploy/post-installation-tasks/security-best-practices/exchange-extended-protection' 'MS Learn')))
                }
            } catch { }
        }

        # MRS Proxy (EWS)
        if ($State['InstallMailbox'] -and -not $State['InstallEdge']) {
            try {
                $ewsVdir = Get-WebServicesVirtualDirectory -Server $env:COMPUTERNAME -ADPropertiesOnly -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($ewsVdir) {
                    $mrsBadge = if (-not $ewsVdir.MRSProxyEnabled) { Format-Badge 'Disabled ✓' 'ok' } else { Format-Badge 'Enabled (review)' 'warn' }
                    $secRows.Add(('<tr><td>MRS Proxy (EWS)</td><td>{0}</td><td>False (enable only for cross-forest migrations)</td><td>{1}</td><td>{2}</td><td>MS Exchange</td></tr>' -f $ewsVdir.MRSProxyEnabled, $mrsBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/exchange/architecture/mailbox-servers/mrs-proxy-endpoint' 'MS Learn')))
                }
            } catch { }
        }

        # MAPI Encryption Required
        if ($State['InstallMailbox'] -and -not $State['InstallEdge']) {
            try {
                $mbxSrv = Get-MailboxServer -Identity $env:COMPUTERNAME -ErrorAction SilentlyContinue
                if ($mbxSrv) {
                    $mapiEncBadge = if ($mbxSrv.MAPIEncryptionRequired) { Format-Badge 'Required ✓' 'ok' } else { Format-Badge 'Not required' 'warn' }
                    $secRows.Add(('<tr><td>MAPI Encryption Required</td><td>{0}</td><td>True (forces encrypted Outlook MAPI connections)</td><td>{1}</td><td>{2}</td><td>MS Exchange</td></tr>' -f $mbxSrv.MAPIEncryptionRequired, $mapiEncBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/exchange/clients/mapi-over-http/configure-mapi-over-http' 'MS Learn')))
                }
            } catch { }
        }

        # SMTP Banner hardening
        if (-not $State['InstallEdge']) {
            try {
                $feBannerConns = @(Get-ReceiveConnector -Server $env:COMPUTERNAME -ErrorAction SilentlyContinue | Where-Object { $_.TransportRole -eq 'FrontendTransport' })
                if ($feBannerConns.Count -gt 0) {
                    $bannersHardened = ($feBannerConns | Where-Object { $_.Banner -and $_.Banner -notlike '*Microsoft ESMTP*' }).Count
                    $bannerBadge = if ($bannersHardened -eq $feBannerConns.Count) { Format-Badge 'Hardened ✓' 'ok' } else { Format-Badge ('{0}/{1} hardened' -f $bannersHardened, $feBannerConns.Count) 'warn' }
                    $secRows.Add(('<tr><td>SMTP Banner</td><td>{0}/{1} Frontend connectors hardened</td><td>Generic banner (hides Exchange version from attackers)</td><td>{2}</td><td>{3}</td><td>CIS / DISA STIG</td></tr>' -f $bannersHardened, $feBannerConns.Count, $bannerBadge, (Format-RefLink 'https://learn.microsoft.com/en-us/exchange/mail-flow/connectors/receive-connectors' 'MS Learn')))
                }
            } catch { }
        }

        $secContent = '<table class="data-table"><tr><th>Setting</th><th>Current Value</th><th>Exchange Recommendation</th><th>Status</th><th>Reference</th><th>CIS / BSI</th></tr>' + ($secRows -join '') + '</table>'

        # ── 6. PERFORMANCE SETTINGS (with best-practice column) ───────────────
        $perfRows = [System.Collections.Generic.List[string]]::new()
        try {
            $plan = Get-CimInstance -Namespace 'root\cimv2\power' -ClassName Win32_PowerPlan -Filter 'IsActive=True' -ErrorAction Stop
            $isHighPerf = $plan.InstanceID -like "*$POWERPLAN_HIGH_PERFORMANCE*"
            $planBadge  = if ($isHighPerf) { Format-Badge 'High Performance ✓' 'ok' } else { Format-Badge 'Not High Performance' 'warn' }
            $perfRows.Add(('<tr><td>Power Plan</td><td>{0}</td><td>High Performance</td><td>{1}</td></tr>' -f $plan.ElementName, $planBadge))
        } catch { }
        try {
            $pf = Get-CimInstance Win32_PageFileSetting -ErrorAction Stop | Select-Object -First 1
            if ($pf) {
                $ramMB        = [math]::Round((Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue).TotalPhysicalMemory / 1MB, 0)
                $recMB        = if ($State['MajorSetupVersion'] -ge $EX2019_MAJOR) { [int]($ramMB * 0.25) } else { [math]::Min($ramMB + 10, 32768 + 10) }
                $pfOk         = $pf.InitialSize -ge $recMB -and $pf.MaximumSize -ge $recMB
                $pfBadge      = if ($pfOk) { Format-Badge '✓' 'ok' } else { Format-Badge 'Below recommendation' 'warn' }
                $perfRows.Add(('<tr><td>Pagefile</td><td>{0} — Init: {1} MB / Max: {2} MB</td><td>≥ {3} MB</td><td>{4}</td></tr>' -f $pf.Name, $pf.InitialSize, $pf.MaximumSize, $recMB, $pfBadge))
            }
        } catch { }
        $keepAlive = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters' 'KeepAliveTime'
        if ($keepAlive) {
            $kaBadge = if ($keepAlive -le 900000) { Format-Badge '✓' 'ok' } else { Format-Badge 'High' 'warn' }
            $perfRows.Add(('<tr><td>TCP KeepAliveTime</td><td>{0} ms ({1} min)</td><td>900000 ms (15 min)</td><td>{2}</td></tr>' -f $keepAlive, [math]::Round($keepAlive / 60000, 0), $kaBadge))
        }
        try {
            Get-NetAdapterRss -ErrorAction SilentlyContinue | Where-Object { $_.Enabled } | ForEach-Object {
                $perfRows.Add(('<tr><td>RSS: {0}</td><td>Enabled — Queues: {1}</td><td>Enabled; Queues = physical cores</td><td>{2}</td></tr>' -f $_.Name, $_.NumberOfReceiveQueues, (Format-Badge 'Enabled ✓' 'ok')))
            }
        } catch { }
        $maxConcAPI = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters' 'MaxConcurrentApi'
        if ($maxConcAPI) {
            $mcaBadge = if ($maxConcAPI -ge 10) { Format-Badge '✓' 'ok' } else { Format-Badge 'Low' 'warn' }
            $perfRows.Add(('<tr><td>Netlogon MaxConcurrentApi</td><td>{0}</td><td>&ge; logical cores (min 10)</td><td>{1}</td></tr>' -f $maxConcAPI, $mcaBadge))
        }
        $ctsPct = Get-SecRegVal 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Search\SystemParameters' 'CtsProcessorAffinityPercentage'
        if ($null -ne $ctsPct) {
            $ctsBadge = if ($ctsPct -eq 0) { Format-Badge '0% ✓' 'ok' } else { Format-Badge "$ctsPct%" 'warn' }
            $perfRows.Add(('<tr><td>CTS Processor Affinity %</td><td>{0}</td><td>0 (Exchange Search best practice)</td><td>{1}</td></tr>' -f $ctsPct, $ctsBadge))
        }
        $tcpOffload = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters' 'DisableTaskOffload'
        if ($null -ne $tcpOffload) {
            $toBadge = if ($tcpOffload -eq 1) { Format-Badge 'Disabled ✓' 'ok' } else { Format-Badge 'Enabled' 'warn' }
            $perfRows.Add(('<tr><td>TCP Task Offload</td><td>DisableTaskOffload = {0}</td><td>1 = Disabled</td><td>{1}</td></tr>' -f $tcpOffload, $toBadge))
        }
        $wSearch = (Get-Service -Name WSearch -ErrorAction SilentlyContinue)
        if ($wSearch) {
            $wsBadge = if ($wSearch.StartType -eq 'Disabled') { Format-Badge 'Disabled ✓' 'ok' } else { Format-Badge $wSearch.StartType 'warn' }
            $perfRows.Add(('<tr><td>Windows Search Service</td><td>{0}</td><td>Disabled (Exchange uses own indexing)</td><td>{1}</td></tr>' -f $wSearch.StartType, $wsBadge))
        }

        # RPC Minimum Connection Timeout
        $rpcTO = Get-SecRegVal 'HKLM:\SOFTWARE\Microsoft\Rpc' 'MinimumConnectionTimeout'
        if ($null -ne $rpcTO) {
            $rpcBadge = if ($rpcTO -ge 120) { Format-Badge "✓ ${rpcTO}s" 'ok' } else { Format-Badge "${rpcTO}s (low)" 'warn' }
            $perfRows.Add(('<tr><td>RPC Min Connection Timeout</td><td>{0} s</td><td>120 s (prevents premature RPC timeouts under load)</td><td>{1}</td></tr>' -f $rpcTO, $rpcBadge))
        }

        # TCP Chimney Offload
        $tcpChim = Get-SecRegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters' 'EnableTCPChimney'
        if ($null -ne $tcpChim) {
            $chimBadge = if ($tcpChim -eq 0) { Format-Badge 'Disabled ✓' 'ok' } else { Format-Badge 'Enabled' 'warn' }
            $perfRows.Add(('<tr><td>TCP Chimney Offload</td><td>EnableTCPChimney = {0}</td><td>0 = Disabled (Microsoft recommendation for Exchange)</td><td>{1}</td></tr>' -f $tcpChim, $chimBadge))
        }

        # NodeRunner Max Memory
        $nodeRunMem = Get-SecRegVal 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Search\SystemParameters' 'NodeRunnerMaxMemory'
        if ($null -ne $nodeRunMem) {
            $nrBadge = if ($nodeRunMem -eq 0) { Format-Badge 'Unlimited ✓' 'ok' } else { Format-Badge "${nodeRunMem} MB (capped)" 'warn' }
            $perfRows.Add(('<tr><td>NodeRunner Max Memory</td><td>{0}</td><td>0 = unlimited (Exchange Search best practice)</td><td>{1}</td></tr>' -f $nodeRunMem, $nrBadge))
        }

        $perfContent = '<table class="data-table"><tr><th>Setting</th><th>Current Value</th><th>Exchange Recommendation</th><th>Status</th></tr>' + ($perfRows -join '') + '</table>'

        # ── 7. RBAC ROLE GROUP MEMBERSHIP ─────────────────────────────────────
        $rbacGroups = @(
            'Organization Management', 'Server Management', 'Recipient Management',
            'Help Desk', 'Hygiene Management', 'Compliance Management',
            'Records Management', 'Discovery Management',
            'Public Folder Management', 'View-Only Organization Management'
        )
        $rbacRows = [System.Collections.Generic.List[string]]::new()
        $rbacRows.Add('<tr><th>Role Group</th><th>Members</th></tr>')
        foreach ($group in $rbacGroups) {
            try {
                $members = @(Get-RoleGroupMember -Identity $group -ErrorAction Stop)
                $memberHtml = if ($members.Count -gt 0) {
                    ($members | ForEach-Object { '<code>{0}</code> <span style="color:#888;font-size:11px">({1})</span>' -f [string]$_.Name, [string]$_.RecipientType }) -join '<br>'
                } else { '<em style="color:#888">no members</em>' }
                $rbacRows.Add(('<tr><td>{0}</td><td>{1}</td></tr>' -f $group, $memberHtml))
            }
            catch {
                $rbacRows.Add(('<tr><td>{0}</td><td style="color:#c50f1f"><em>Query failed: {1}</em></td></tr>' -f $group, $_.Exception.Message))
            }
        }
        $rbacContent = '<table class="data-table">' + ($rbacRows -join '') + '</table>'

        # ── 9. INSTALLATION LOG ───────────────────────────────────────────────
        # B16 fixes:
        # 1. ReadAllText was called with UTF-8 encoding; PS 5.1 transcripts are UTF-16 LE —
        #    removed explicit encoding so .NET auto-detects the BOM (handles both UTF-8/UTF-16).
        # 2. Wrapped in try/catch — an IOExceptionon a large/locked transcript file propagated
        #    to the global trap { break } and killed the entire script.
        # 3. Capped to the last 2000 lines — a transcript accumulated over multiple reboots
        #    (30h+ install) can be several MB; embedding the full file makes the HTML report
        #    unusably large and strains memory during the regex-replace operations.
        $logContent = if ($State['TranscriptFile'] -and (Test-Path $State['TranscriptFile'])) {
            try {
                $logLines   = [System.IO.File]::ReadAllLines($State['TranscriptFile'])
                $totalLines = $logLines.Count
                $maxLines   = 2000
                $truncated  = $totalLines -gt $maxLines
                $logLines   = if ($truncated) { $logLines[($totalLines - $maxLines)..($totalLines - 1)] } else { $logLines }
                $logText    = $logLines -join "`n"
                $logEscaped = $logText -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
                $truncNote  = if ($truncated) { '<div style="color:#ff9800;font-size:11px;margin-bottom:6px">&#9888; Log truncated — showing last {0} of {1} lines. Full log: <code>{2}</code></div>' -f $maxLines, $totalLines, $State['TranscriptFile'] } else { '' }
                $truncNote + '<pre style="font-family:Consolas,monospace;font-size:12px;line-height:1.5;background:#1e1e1e;color:#d4d4d4;padding:16px;border-radius:4px;overflow:auto;max-height:600px;white-space:pre-wrap;word-break:break-all">' + $logEscaped + '</pre>'
            } catch {
                '<p style="color:#8a8886"><em>Log file could not be read ({0}): {1}</em></p>' -f $State['TranscriptFile'], $_.Exception.Message
            }
        } else {
            '<p style="color:#8a8886"><em>Log file not found: {0}</em></p>' -f $State['TranscriptFile']
        }

        # ── BUILD HTML ────────────────────────────────────────────────────────
        $css = @'
:root{--primary:#1a2332;--accent:#0078d4;--ok:#107c10;--warn:#d83b01;--fail:#c50f1f;--bg:#f3f2f1;--card:#fff;--border:#e1dfdd}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:var(--bg);color:#252423;font-size:14px}
header{background:var(--primary);color:#fff;padding:24px 40px;display:flex;align-items:center;gap:16px}
header h1{font-size:22px;font-weight:300;letter-spacing:.5px}
.logo{width:44px;height:44px;background:var(--accent);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:18px;font-weight:700;flex-shrink:0}
.summary-bar{background:var(--accent);color:#fff;padding:12px 40px;display:flex;gap:40px;font-size:13px;flex-wrap:wrap}
.summary-bar span{opacity:.8}.summary-bar strong{opacity:1;font-weight:600}
.container{display:flex}
.toc{width:210px;min-width:210px;background:var(--primary);padding:20px 0;position:sticky;top:0;height:100vh;overflow-y:auto;flex-shrink:0}
.toc-title{color:#888;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:1px;padding:16px 20px 6px}
.toc a{display:block;padding:9px 20px;color:#c8c8c8;text-decoration:none;font-size:13px;border-left:3px solid transparent;transition:all .15s}
.toc a:hover{color:#fff;background:rgba(255,255,255,.08);border-left-color:var(--accent)}
main{flex:1;padding:32px 36px;max-width:1200px;overflow-x:auto}
.section{background:var(--card);border-radius:8px;margin-bottom:24px;box-shadow:0 1px 4px rgba(0,0,0,.08);overflow:hidden}
.section-title{background:var(--primary);color:#fff;padding:12px 20px;font-size:15px;font-weight:400;letter-spacing:.3px}
.section-body{padding:20px}
.subsection{margin:22px 0 10px;font-size:13px;font-weight:700;color:var(--primary);border-bottom:2px solid var(--border);padding-bottom:6px;text-transform:uppercase;letter-spacing:.4px}
.data-table{width:100%;border-collapse:collapse;font-size:13px}
.data-table th{background:#f3f2f1;font-weight:600;text-align:left;padding:8px 12px;border-bottom:2px solid var(--border);color:var(--primary)}
.data-table td{padding:7px 12px;border-bottom:1px solid #f0efee;vertical-align:middle}
.data-table tr:last-child td{border-bottom:none}
.data-table tr:hover td{background:#faf9f8}
code{font-family:Consolas,monospace;font-size:12px;background:#f0f0f0;padding:1px 5px;border-radius:3px;word-break:break-all}
footer{background:var(--primary);color:#888;padding:16px 40px;font-size:12px;text-align:center;margin-top:8px}
@media print{
  @page{size:A4 portrait;margin:1.8cm 1.5cm 2cm 1.5cm}
  @page:first{margin-top:1cm}
  *{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}
  .toc,.summary-bar{display:none!important}
  body{background:#fff;font-size:10pt}
  header{padding:12px 20px;break-after:avoid}
  header h1{font-size:14pt}
  .container{display:block}
  main{padding:0;max-width:100%}
  .section{background:#fff;box-shadow:none;border:1pt solid #ccc;margin-bottom:14pt;page-break-inside:avoid;break-inside:avoid}
  .section-title{padding:7pt 12pt;font-size:11pt}
  .section-body{padding:10pt 12pt}
  .subsection{font-size:9pt;margin:14pt 0 6pt}
  .data-table{font-size:8.5pt;width:100%;table-layout:fixed}
  .data-table th,.data-table td{padding:4pt 6pt;word-break:break-word}
  .data-table tr{page-break-inside:avoid;break-inside:avoid}
  code{font-size:7.5pt;word-break:break-all}
  #installlog pre{max-height:none;font-size:6.5pt;page-break-inside:auto}
  #healthchecker iframe{display:none}
  #healthchecker::after{content:"HealthChecker report is embedded in the HTML version of this document.";font-style:italic;color:#666;font-size:10pt}
  a{color:inherit;text-decoration:none}
  footer{font-size:8pt;padding:8pt 20pt;margin-top:6pt}
  /* Print document header on continuation pages via border-top trick */
  .section:nth-child(n+2){border-top:2pt solid #1a2332}
}
'@

        # ── 9. HEALTHCHECKER (embedded via srcdoc — no external file dependency) ──
        $hcContent = if ($State['SkipHealthCheck']) {
            '<p style="color:#666;font-size:13px;">HealthChecker was skipped (<code>-SkipHealthCheck</code> specified).</p>'
        } elseif ($State['HCReportPath'] -and (Test-Path $State['HCReportPath'])) {
            try {
                $hcRaw     = [System.IO.File]::ReadAllText($State['HCReportPath'], [System.Text.Encoding]::UTF8)
                # Escape for HTML attribute value: & → &amp;  " → &quot;
                $hcSrcdoc  = $hcRaw -replace '&', '&amp;' -replace '"', '&quot;'
                '<iframe srcdoc="' + $hcSrcdoc + '" style="width:100%;height:860px;border:1px solid #e1dfdd;border-radius:4px;" title="HealthChecker Report" sandbox="allow-same-origin allow-scripts allow-popups"></iframe>'
            }
            catch {
                '<p style="color:#d83b01;font-size:13px;">HealthChecker report could not be embedded: {0}</p>' -f $_.Exception.Message
            }
        } else {
            '<p style="color:#d83b01;font-size:13px;">HealthChecker report not found — HC may have failed to run. Check the installation log for details.</p>'
        }

        # ── 0. MANAGEMENT SUMMARY ─────────────────────────────────────────────
        $mgmtRows = [System.Collections.Generic.List[string]]::new()

        # Count security badge types from built HTML string
        $secOK   = ([regex]::Matches($secContent,   'background:#107c10')).Count
        $secWarn = ([regex]::Matches($secContent,   'background:#d83b01')).Count
        $secFail = ([regex]::Matches($secContent,   'background:#c50f1f')).Count
        $perfOK  = ([regex]::Matches($perfContent,  'background:#107c10')).Count
        $perfWarn= ([regex]::Matches($perfContent,  'background:#d83b01')).Count

        $secStatusBadge  = if ($secFail -gt 0)  { Format-Badge "$secFail critical, $secWarn warnings" 'fail' }
                           elseif ($secWarn -gt 0) { Format-Badge "$secWarn warnings" 'warn' }
                           else { Format-Badge "All $secOK items OK" 'ok' }
        $perfStatusBadge = if ($perfWarn -gt 0) { Format-Badge "$perfWarn items to review" 'warn' } else { Format-Badge "All $perfOK items OK" 'ok' }

        $certStatus = try {
            $certs = @(Get-ExchangeCertificate -Server $env:COMPUTERNAME -ErrorAction Stop | Where-Object { $_.Thumbprint -and $_.NotAfter -gt [datetime]'1970-01-01' })
            $expiring = @($certs | Where-Object { [int][Math]::Floor(($_.NotAfter - (Get-Date)).TotalDays) -le 90 })
            if ($expiring.Count -gt 0) { Format-Badge ('{0} expiring within 90 days' -f $expiring.Count) 'warn' }
            else { Format-Badge ('{0} certificate(s) — all valid' -f $certs.Count) 'ok' }
        } catch { Format-Badge 'Could not query' 'na' }

        $vdirStatus = if ($State['Namespace']) { Format-Badge ('Configured ({0})' -f $State['Namespace']) 'ok' } else { Format-Badge 'Not configured' 'warn' }
        $suStatus   = if ($State['IncludeFixes']) { Format-Badge 'Enabled' 'ok' } else { Format-Badge 'Skipped' 'warn' }
        $hcStatus   = if ($State['SkipHealthCheck']) { Format-Badge 'Skipped' 'na' } elseif ($State['HCReportPath']) { Format-Badge 'Completed — see section below' 'ok' } else { Format-Badge 'Failed / not found' 'warn' }
        $modeStatus = if ($State['ConfigDriven']) { Format-Badge 'Autopilot (fully automated)' 'info' } else { Format-Badge 'Copilot (interactive)' 'info' }

        $mgmtRows.Add(('<tr><td style="width:220px"><strong>Exchange Version</strong></td><td>{0}</td><td>{1}</td></tr>' -f $exVersion, (Format-Badge 'Installed' 'ok')))
        $serverFqdn = try { '{0}.{1}' -f (Get-CimInstance Win32_ComputerSystem -EA SilentlyContinue).DNSHostName, (Get-CimInstance Win32_ComputerSystem -EA SilentlyContinue).Domain } catch { '' }
        $mgmtRows.Add(('<tr><td><strong>Server</strong></td><td>{0}</td><td>{1}</td></tr>' -f ('{0} ({1})' -f $env:COMPUTERNAME, $serverFqdn), (Format-Badge 'OK' 'ok')))
        $mgmtRows.Add('<tr><td><strong>Organization</strong></td><td>{0}</td><td></td></tr>' -f $State['OrganizationName'])
        $mgmtRows.Add(('<tr><td><strong>Installation Mode</strong></td><td>{0}</td><td>{1}</td></tr>' -f $instMode, $modeStatus))
        $mgmtRows.Add(('<tr><td><strong>Virtual Directory URLs</strong></td><td>{0}</td><td>{1}</td></tr>' -f $State['Namespace'], $vdirStatus))
        $mgmtRows.Add(('<tr><td><strong>Security Hardening</strong></td><td>{0} OK / {1} warnings / {2} critical</td><td>{3}</td></tr>' -f $secOK, $secWarn, $secFail, $secStatusBadge))
        $mgmtRows.Add(('<tr><td><strong>Performance Settings</strong></td><td>{0} OK / {1} to review</td><td>{2}</td></tr>' -f $perfOK, $perfWarn, $perfStatusBadge))
        $mgmtRows.Add('<tr><td><strong>Certificates</strong></td><td></td><td>{0}</td></tr>' -f $certStatus)
        $mgmtRows.Add('<tr><td><strong>Security Updates</strong></td><td></td><td>{0}</td></tr>' -f $suStatus)
        $mgmtRows.Add('<tr><td><strong>HealthChecker</strong></td><td></td><td>{0}</td></tr>' -f $hcStatus)
        $mgmtRows.Add('<tr><td><strong>Report Generated</strong></td><td>{0}</td><td></td></tr>' -f $reportDate)

        # Action items — surface any WARN/FAIL as bullet list
        $actionItems = [System.Collections.Generic.List[string]]::new()
        if ($secFail -gt 0)  { $actionItems.Add('<li><strong>Security:</strong> {0} critical finding(s) require immediate attention — see Security Settings section.</li>' -f $secFail) }
        if ($secWarn -gt 0)  { $actionItems.Add('<li><strong>Security:</strong> {0} warning(s) — review Security Settings section.</li>' -f $secWarn) }
        if ($perfWarn -gt 0) { $actionItems.Add('<li><strong>Performance:</strong> {0} setting(s) below recommendation — review Performance &amp; Tuning section.</li>' -f $perfWarn) }
        if (-not $State['Namespace']) { $actionItems.Add('<li><strong>Virtual Directories:</strong> No access namespace configured. OWA/ECP/EWS URLs may still point to server hostname.</li>') }
        if (-not $State['IncludeFixes']) { $actionItems.Add('<li><strong>Security Updates:</strong> Exchange Security Update installation was skipped. Apply the latest SU manually.</li>') }

        $actionHtml = if ($actionItems.Count -gt 0) {
            '<h3 class="subsection">Action Items</h3><ul style="margin:8px 0 0 20px;line-height:1.8;font-size:13px">' + ($actionItems -join '') + '</ul>'
        } else {
            '<p style="color:#107c10;font-size:13px;margin-top:12px">&#10003; No critical action items identified.</p>'
        }

        $mgmtContent = '<table class="data-table"><tr><th>Item</th><th>Detail</th><th>Status</th></tr>' + ($mgmtRows -join '') + '</table>' + $actionHtml

        $sections = @(
            (New-HtmlSection 'summary'      'Management Summary'        $mgmtContent)
            (New-HtmlSection 'params'       'Installation Parameters'   ('<table class="data-table">' + ($instRows -join '') + '</table>'))
            (New-HtmlSection 'system'       'System Information'        $sysContent)
            (New-HtmlSection 'ad'           'Active Directory'          $adContent)
            (New-HtmlSection 'exchange'     'Exchange Configuration'    $exContent)
            (New-HtmlSection 'security'     'Security Settings'         $secContent)
            (New-HtmlSection 'performance'  'Performance &amp; Tuning'  $perfContent)
            (New-HtmlSection 'rbac'         'RBAC Role Group Membership' $rbacContent)
            (New-HtmlSection 'installlog'   'Installation Log'          $logContent)
            (New-HtmlSection 'healthchecker' 'HealthChecker'            $hcContent)
        )

        $toc = @(
            '<div class="toc-title">Contents</div>'
            '<a href="#summary">Management Summary</a>'
            '<a href="#params">Installation Parameters</a>'
            '<a href="#system">System Information</a>'
            '<a href="#ad">Active Directory</a>'
            '<a href="#exchange">Exchange Configuration</a>'
            '<a href="#security">Security Settings</a>'
            '<a href="#performance">Performance &amp; Tuning</a>'
            '<a href="#rbac">RBAC Role Groups</a>'
            '<a href="#installlog">Installation Log</a>'
            '<a href="#healthchecker">HealthChecker</a>'
        )

        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Exchange Installation Report — $env:COMPUTERNAME</title>
<style>$css</style>
</head>
<body>
<header>
  <div class="logo">Ex</div>
  <div>
    <h1>Exchange Server Installation Report</h1>
    <div style="font-size:12px;opacity:.65;margin-top:4px">Generated by EXpress v$ScriptVersion</div>
  </div>
</header>
<div class="summary-bar">
  <div><span>Server: </span><strong>$env:COMPUTERNAME</strong></div>
  <div><span>Exchange: </span><strong>$exVersion</strong></div>
  <div><span>Organization: </span><strong>$($State['OrganizationName'])</strong></div>
  <div><span>Report Date: </span><strong>$reportDate</strong></div>
</div>
<div class="container">
  <nav class="toc">$($toc -join '')</nav>
  <main>$($sections -join '')</main>
</div>
<footer>Exchange Server Installation Report &bull; $env:COMPUTERNAME &bull; $reportDate &bull; EXpress v$ScriptVersion</footer>
</body>
</html>
"@

        try {
            $html | Out-File -FilePath $reportPath -Encoding utf8 -ErrorAction Stop
            Write-MyOutput ('Installation Report saved to {0}' -f $reportPath)
        }
        catch {
            Write-MyWarning ('Could not write Installation Report: {0}' -f $_.Exception.Message)
            return
        }

        # Optional PDF via Microsoft Edge headless
        $edgeExe = @(
            "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
            "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
        ) | Where-Object { Test-Path $_ } | Select-Object -First 1

        if ($edgeExe) {
            $pdfPath    = $reportPath -replace '\.html$', '.pdf'
            $edgeStdErr = Join-Path $env:TEMP 'edge_headless_stderr.txt'
            Write-MyVerbose ('Generating PDF via Edge headless: {0}' -f $pdfPath)
            try {
                $fileUri  = 'file:///{0}' -f ($reportPath -replace '\\', '/')
                $edgeArgs = '--headless', '--disable-gpu', '--run-all-compositor-stages-before-draw',
                            '--log-level=3',        # suppress DevTools/renderer noise
                            '--disable-extensions', # no extension interference
                            '--print-to-pdf-no-header',  # remove browser URL/date chrome
                            '--virtual-time-budget=8000', # allow srcdoc iframe time to render
                            "--print-to-pdf=`"$pdfPath`"", "`"$fileUri`""
                $proc = Start-Process -FilePath $edgeExe -ArgumentList $edgeArgs -NoNewWindow -Wait -PassThru `
                            -RedirectStandardError $edgeStdErr -ErrorAction Stop
                Remove-Item $edgeStdErr -ErrorAction SilentlyContinue
                if ($proc.ExitCode -eq 0 -and (Test-Path $pdfPath)) {
                    Write-MyOutput ('Installation Report PDF saved to {0}' -f $pdfPath)
                }
                else {
                    Write-MyVerbose ('Edge PDF export exit code: {0}' -f $proc.ExitCode)
                }
            }
            catch {
                Write-MyVerbose ('Edge PDF generation failed: {0}' -f $_.Exception.Message)
            }
        }
        else {
            Write-MyVerbose 'Microsoft Edge not found — skipping PDF export. Open the HTML report in a browser and use File > Print > Save as PDF.'
        }
    }

    # ── OpenXML Engine (shared with tools/Build-ConceptTemplate.ps1) ─────────────
    # Pure PowerShell, no Office/COM required. PS2Exe-safe.

