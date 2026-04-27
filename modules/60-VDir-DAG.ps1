    function Set-VirtualDirectoryURLs {
        if (-not $State['Namespace']) {
            Write-MyVerbose 'No Namespace specified, skipping Virtual Directory URL configuration'
            return
        }

        $ns     = $State['Namespace']
        $server = $env:COMPUTERNAME
        $errors = 0
        $changed = 0
        Write-MyStep -Label 'VDir URLs' -Value ('configuring for {0}' -f $ns) -Status Run

        # Exchange VDir cmdlets call ShouldContinue("host can't be resolved") when the namespace
        # doesn't resolve in DNS — ShouldContinue cannot be suppressed by -Confirm:$false or
        # preference variables. Add a temporary hosts entry if needed and remove it afterwards.
        $hostsFile      = "$env:SystemRoot\System32\drivers\etc\hosts"
        $tempHostsMark  = '# EXpress-temp-vdir'
        $hostsBackup    = $null
        $dlDomain       = $State['DownloadDomain']
        $nsResolves     = $false
        $dlResolves     = $false
        try { [System.Net.Dns]::GetHostEntry($ns) | Out-Null; $nsResolves = $true } catch { } # intentional: exception = not resolvable; handled below
        if ($dlDomain) { try { [System.Net.Dns]::GetHostEntry($dlDomain) | Out-Null; $dlResolves = $true } catch { } } # intentional: same
        if (-not $nsResolves) {
            Write-MyVerbose ('Namespace {0} not resolvable — adding temporary hosts entry to suppress VDir confirmation prompt' -f $ns)
            $hostsBackup = [System.IO.File]::ReadAllBytes($hostsFile)
            "`r`n127.0.0.1`t$ns`t$tempHostsMark" | Add-Content -Path $hostsFile -Encoding ASCII -ErrorAction SilentlyContinue
        }
        if ($dlDomain -and -not $dlResolves) {
            Write-MyVerbose ('Download domain {0} not resolvable — adding temporary hosts entry' -f $dlDomain)
            if (-not $hostsBackup) { $hostsBackup = [System.IO.File]::ReadAllBytes($hostsFile) }
            "`r`n127.0.0.1`t$dlDomain`t$tempHostsMark" | Add-Content -Path $hostsFile -Encoding ASCII -ErrorAction SilentlyContinue
        }

        # Helper: compare a vdir URL property (Uri object or string) to a target string
        function Test-VdirUrl($current, $target) {
            if (-not $current) { return $false }
            return ([string]$current -eq $target)
        }

        # OWA — set URL and UPN logon format
        try {
            $vd = Get-OwaVirtualDirectory -Identity "$server\owa (Default Web Site)" -ADPropertiesOnly -ErrorAction Stop
            $urlOk    = (Test-VdirUrl $vd.InternalUrl "https://$ns/owa") -and (Test-VdirUrl $vd.ExternalUrl "https://$ns/owa")
            $formatOk = ([string]$vd.LogonFormat -eq 'PrincipalName')
            if ($urlOk -and $formatOk) {
                Write-MyVerbose 'OWA: URLs and logon format already set, skipping'
            } else {
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-OwaVirtualDirectory -Identity '$server\owa (Default Web Site)' -InternalUrl 'https://$ns/owa' -ExternalUrl 'https://$ns/owa' -LogonFormat PrincipalName -DefaultDomain ''")
                Set-OwaVirtualDirectory -Identity "$server\owa (Default Web Site)" `
                    -InternalUrl "https://$ns/owa" -ExternalUrl "https://$ns/owa" `
                    -LogonFormat PrincipalName -DefaultDomain '' `
                    -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
                Write-MyVerbose 'OWA virtual directory configured (UPN logon)'
                $changed++
            }
        }
        catch { Write-MyWarning ('OWA: {0}' -f $_.Exception.Message); $errors++ }

        # OWA Download Domains — CVE-2021-1730 mitigation (isolates attachment downloads to a separate hostname)
        if ($dlDomain) {
            try {
                $vd = Get-OwaVirtualDirectory -Identity "$server\owa (Default Web Site)" -ADPropertiesOnly -ErrorAction Stop
                $dlOk = ([string]$vd.ExternalDownloadHostName -eq $dlDomain) -and ([string]$vd.InternalDownloadHostName -eq $dlDomain)
                if ($dlOk) {
                    Write-MyVerbose ('OWA Download Domains already set to {0}, skipping' -f $dlDomain)
                } else {
                    Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-OwaVirtualDirectory -Identity '$server\owa (Default Web Site)' -ExternalDownloadHostName '$dlDomain' -InternalDownloadHostName '$dlDomain'")
                    Set-OwaVirtualDirectory -Identity "$server\owa (Default Web Site)" `
                        -ExternalDownloadHostName $dlDomain -InternalDownloadHostName $dlDomain `
                        -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
                    Write-MyVerbose ('OWA Download Domains configured: {0} (CVE-2021-1730 mitigation)' -f $dlDomain)
                    $changed++
                }
            }
            catch { Write-MyWarning ('OWA Download Domains: {0}' -f $_.Exception.Message); $errors++ }
            # EnableDownloadDomains must be set at org level for CVE-2021-1730 mitigation to take effect
            try {
                $tc = Get-OrganizationConfig -ErrorAction Stop
                if (-not $tc.EnableDownloadDomains) {
                    Register-ExecutedCommand -Category 'VirtualDirectories' -Command 'Set-OrganizationConfig -EnableDownloadDomains $true'
                    Set-OrganizationConfig -EnableDownloadDomains $true -ErrorAction Stop
                    Write-MyVerbose 'EnableDownloadDomains enabled at org level (CVE-2021-1730)'
                    $changed++
                } else {
                    Write-MyVerbose 'EnableDownloadDomains already enabled at org level'
                }
            }
            catch { Write-MyWarning ('EnableDownloadDomains: {0}' -f $_.Exception.Message); $errors++ }
        }

        # ECP
        try {
            $vd = Get-EcpVirtualDirectory -Identity "$server\ecp (Default Web Site)" -ADPropertiesOnly -ErrorAction Stop
            if ((Test-VdirUrl $vd.InternalUrl "https://$ns/ecp") -and (Test-VdirUrl $vd.ExternalUrl "https://$ns/ecp")) {
                Write-MyVerbose 'ECP: URLs already set, skipping'
            } else {
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-EcpVirtualDirectory -Identity '$server\ecp (Default Web Site)' -InternalUrl 'https://$ns/ecp' -ExternalUrl 'https://$ns/ecp'")
                Set-EcpVirtualDirectory -Identity "$server\ecp (Default Web Site)" `
                    -InternalUrl "https://$ns/ecp" -ExternalUrl "https://$ns/ecp" `
                    -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
                Write-MyVerbose 'ECP virtual directory configured'
                $changed++
            }
        }
        catch { Write-MyWarning ('ECP: {0}' -f $_.Exception.Message); $errors++ }

        # EWS
        try {
            $vd = Get-WebServicesVirtualDirectory -Identity "$server\EWS (Default Web Site)" -ADPropertiesOnly -ErrorAction Stop
            if ((Test-VdirUrl $vd.InternalUrl "https://$ns/EWS/Exchange.asmx") -and (Test-VdirUrl $vd.ExternalUrl "https://$ns/EWS/Exchange.asmx")) {
                Write-MyVerbose 'EWS: URLs already set, skipping'
            } else {
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-WebServicesVirtualDirectory -Identity '$server\EWS (Default Web Site)' -InternalUrl 'https://$ns/EWS/Exchange.asmx' -ExternalUrl 'https://$ns/EWS/Exchange.asmx'")
                Set-WebServicesVirtualDirectory -Identity "$server\EWS (Default Web Site)" `
                    -InternalUrl "https://$ns/EWS/Exchange.asmx" -ExternalUrl "https://$ns/EWS/Exchange.asmx" `
                    -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
                Write-MyVerbose 'EWS virtual directory configured'
                $changed++
            }
        }
        catch { Write-MyWarning ('EWS: {0}' -f $_.Exception.Message); $errors++ }

        # OAB
        try {
            $vd = Get-OabVirtualDirectory -Identity "$server\OAB (Default Web Site)" -ADPropertiesOnly -ErrorAction Stop
            if ((Test-VdirUrl $vd.InternalUrl "https://$ns/OAB") -and (Test-VdirUrl $vd.ExternalUrl "https://$ns/OAB")) {
                Write-MyVerbose 'OAB: URLs already set, skipping'
            } else {
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-OabVirtualDirectory -Identity '$server\OAB (Default Web Site)' -InternalUrl 'https://$ns/OAB' -ExternalUrl 'https://$ns/OAB'")
                Set-OabVirtualDirectory -Identity "$server\OAB (Default Web Site)" `
                    -InternalUrl "https://$ns/OAB" -ExternalUrl "https://$ns/OAB" `
                    -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
                Write-MyVerbose 'OAB virtual directory configured'
                $changed++
            }
        }
        catch { Write-MyWarning ('OAB: {0}' -f $_.Exception.Message); $errors++ }

        # ActiveSync
        try {
            $vd = Get-ActiveSyncVirtualDirectory -Identity "$server\Microsoft-Server-ActiveSync (Default Web Site)" -ADPropertiesOnly -ErrorAction Stop
            if ((Test-VdirUrl $vd.InternalUrl "https://$ns/Microsoft-Server-ActiveSync") -and (Test-VdirUrl $vd.ExternalUrl "https://$ns/Microsoft-Server-ActiveSync")) {
                Write-MyVerbose 'ActiveSync: URLs already set, skipping'
            } else {
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-ActiveSyncVirtualDirectory -Identity '$server\Microsoft-Server-ActiveSync (Default Web Site)' -InternalUrl 'https://$ns/Microsoft-Server-ActiveSync' -ExternalUrl 'https://$ns/Microsoft-Server-ActiveSync'")
                Set-ActiveSyncVirtualDirectory -Identity "$server\Microsoft-Server-ActiveSync (Default Web Site)" `
                    -InternalUrl "https://$ns/Microsoft-Server-ActiveSync" -ExternalUrl "https://$ns/Microsoft-Server-ActiveSync" `
                    -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
                Write-MyVerbose 'ActiveSync virtual directory configured'
                $changed++
            }
        }
        catch { Write-MyWarning ('ActiveSync: {0}' -f $_.Exception.Message); $errors++ }

        # MAPI — URL first, auth methods in a separate try (not available on all builds)
        try {
            $vd = Get-MapiVirtualDirectory -Identity "$server\mapi (Default Web Site)" -ADPropertiesOnly -ErrorAction Stop
            if ((Test-VdirUrl $vd.InternalUrl "https://$ns/mapi") -and (Test-VdirUrl $vd.ExternalUrl "https://$ns/mapi")) {
                Write-MyVerbose 'MAPI: URLs already set, skipping'
            } else {
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-MapiVirtualDirectory -Identity '$server\mapi (Default Web Site)' -InternalUrl 'https://$ns/mapi' -ExternalUrl 'https://$ns/mapi'")
                Set-MapiVirtualDirectory -Identity "$server\mapi (Default Web Site)" `
                    -InternalUrl "https://$ns/mapi" -ExternalUrl "https://$ns/mapi" `
                    -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
                Write-MyVerbose 'MAPI virtual directory URL configured'
                $changed++
            }
        }
        catch { Write-MyWarning ('MAPI URL: {0}' -f $_.Exception.Message); $errors++ }

        # InternalAuthenticationMethods was removed from Set-MapiVirtualDirectory in Exchange SE RTM.
        # Skip the attempt entirely on SE to avoid the misleading ParameterBindingException in the log.
        if ($State['ExSetupVersion'] -and ([System.Version]$State['ExSetupVersion'] -lt [System.Version]$EXSESETUPEXE_RTM)) {
            try {
                Set-MapiVirtualDirectory -Identity "$server\mapi (Default Web Site)" `
                    -InternalAuthenticationMethods NTLM,Negotiate,OAuth `
                    -ExternalAuthenticationMethods NTLM,Negotiate,OAuth `
                    -ErrorAction Stop -WarningAction SilentlyContinue
                Write-MyVerbose 'MAPI authentication methods configured'
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-MapiVirtualDirectory -Identity '$server\mapi (Default Web Site)' -InternalAuthenticationMethods NTLM,Negotiate,OAuth -ExternalAuthenticationMethods NTLM,Negotiate,OAuth")
            }
            catch { Write-MyVerbose ('MAPI auth methods not supported on this build: {0}' -f $_.Exception.Message) }
        } else {
            Write-MyVerbose 'MAPI auth methods: skipped — InternalAuthenticationMethods removed in Exchange SE RTM'
        }

        # PowerShell — ExternalUrl only; InternalUrl stays http (Exchange internal services use http by default)
        try {
            $vd = Get-PowerShellVirtualDirectory -Identity "$server\PowerShell (Default Web Site)" -ADPropertiesOnly -ErrorAction Stop
            if (Test-VdirUrl $vd.ExternalUrl "https://$ns/powershell") {
                Write-MyVerbose 'PowerShell: ExternalUrl already set, skipping'
            } else {
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-PowerShellVirtualDirectory -Identity '$server\PowerShell (Default Web Site)' -ExternalUrl 'https://$ns/powershell'")
                Set-PowerShellVirtualDirectory -Identity "$server\PowerShell (Default Web Site)" `
                    -ExternalUrl "https://$ns/powershell" `
                    -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
                Write-MyVerbose 'PowerShell virtual directory ExternalUrl configured'
                $changed++
            }
        }
        catch { Write-MyWarning ('PowerShell URL: {0}' -f $_.Exception.Message); $errors++ }

        # Autodiscover SCP — always use autodiscover.<parent-domain>, not the namespace hostname
        try {
            $cas = Get-ClientAccessService -Identity $server -ErrorAction Stop
            $nsParts   = $ns -split '\.'
            $scpHost   = if ($nsParts[0] -eq 'autodiscover') { $ns } else { 'autodiscover.' + ($nsParts[1..($nsParts.Length-1)] -join '.') }
            $scpTarget = "https://$scpHost/Autodiscover/Autodiscover.xml"
            if ([string]$cas.AutoDiscoverServiceInternalUri -eq $scpTarget) {
                Write-MyVerbose 'Autodiscover SCP: already set, skipping'
            } else {
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-ClientAccessService -Identity '$server' -AutoDiscoverServiceInternalUri '$scpTarget'")
                Set-ClientAccessService -Identity $server `
                    -AutoDiscoverServiceInternalUri $scpTarget `
                    -ErrorAction Stop -WarningAction SilentlyContinue
                Write-MyVerbose 'Autodiscover SCP configured'
                $changed++
            }
        }
        catch { Write-MyWarning ('Autodiscover SCP: {0}' -f $_.Exception.Message); $errors++ }

        # IMAP4 — ExternalConnectionSettings / InternalConnectionSettings
        try {
            $imapTarget = "$ns`:993:SSL"
            $imap = Get-ImapSettings -Server $server -ErrorAction Stop
            $extOk = [bool]($imap.ExternalConnectionSettings | Where-Object { [string]$_ -eq $imapTarget })
            $intOk = [bool]($imap.InternalConnectionSettings | Where-Object { [string]$_ -eq $imapTarget })
            if ($extOk -and $intOk) {
                Write-MyVerbose 'IMAP4: connection settings already set, skipping'
            } else {
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-ImapSettings -Server '$server' -ExternalConnectionSettings @('$imapTarget') -InternalConnectionSettings @('$imapTarget')")
                Set-ImapSettings -Server $server `
                    -ExternalConnectionSettings @($imapTarget) `
                    -InternalConnectionSettings @($imapTarget) `
                    -ErrorAction Stop
                Write-MyVerbose ('IMAP4 connection settings configured: {0}' -f $imapTarget)
                $changed++
            }
        }
        catch { Write-MyWarning ('IMAP4: {0}' -f $_.Exception.Message); $errors++ }

        # POP3 — ExternalConnectionSettings / InternalConnectionSettings
        try {
            $popTarget = "$ns`:995:SSL"
            $pop = Get-PopSettings -Server $server -ErrorAction Stop
            $extOk = [bool]($pop.ExternalConnectionSettings | Where-Object { [string]$_ -eq $popTarget })
            $intOk = [bool]($pop.InternalConnectionSettings | Where-Object { [string]$_ -eq $popTarget })
            if ($extOk -and $intOk) {
                Write-MyVerbose 'POP3: connection settings already set, skipping'
            } else {
                Register-ExecutedCommand -Category 'VirtualDirectories' -Command ("Set-PopSettings -Server '$server' -ExternalConnectionSettings @('$popTarget') -InternalConnectionSettings @('$popTarget')")
                Set-PopSettings -Server $server `
                    -ExternalConnectionSettings @($popTarget) `
                    -InternalConnectionSettings @($popTarget) `
                    -ErrorAction Stop
                Write-MyVerbose ('POP3 connection settings configured: {0}' -f $popTarget)
                $changed++
            }
        }
        catch { Write-MyWarning ('POP3: {0}' -f $_.Exception.Message); $errors++ }

        # Restore hosts file to exact pre-modification state using the binary backup.
        if ($hostsBackup) {
            try {
                [System.IO.File]::WriteAllBytes($hostsFile, $hostsBackup)
                Write-MyVerbose 'Temporary hosts entries removed (hosts file restored from backup)'
            }
            catch { Write-MyVerbose ('Could not restore hosts file: {0}' -f $_.Exception.Message) }
        }

        if ($errors -eq 0) {
            if ($changed -gt 0) {
                Write-MyStep -Label 'VDir URLs' -Value ('https://{0} (OWA: UPN)' -f $ns) -Status OK
            } else {
                Write-MyStep -Label 'VDir URLs' -Value ('https://{0} (already correct)' -f $ns) -Status OK
            }
        }
        else {
            Write-MyWarning ('{0} virtual directory(s) could not be configured — check warnings above' -f $errors)
        }
    }

    function Join-DAG {
        if (-not $State['DAGName']) {
            return
        }

        Write-MyStep -Label 'DAG join' -Value $State['DAGName'] -Status Run

        # Ensure Exchange module is loaded
        Import-ExchangeModule

        try {
            $dag = Get-DatabaseAvailabilityGroup -Identity $State['DAGName'] -ErrorAction Stop
            if ($null -eq $dag) {
                Write-MyError ('DAG {0} not found' -f $State['DAGName'])
                exit $ERR_DAGJOIN
            }
            if ($dag.Servers -contains $env:COMPUTERNAME) {
                Write-MyStep -Label 'DAG' -Value ('{0} (already member)' -f $State['DAGName']) -Status OK
                return
            }

            Register-ExecutedCommand -Category 'DAG' -Command ("Add-DatabaseAvailabilityGroupServer -Identity '$($State['DAGName'])' -MailboxServer '$env:COMPUTERNAME'")
            Add-DatabaseAvailabilityGroupServer -Identity $State['DAGName'] -MailboxServer $env:COMPUTERNAME -ErrorAction Stop
            Write-MyStep -Label 'DAG join' -Value ('{0} (joined)' -f $State['DAGName']) -Status OK
        }
        catch {
            Write-MyError ('Failed to join DAG {0}: {1}' -f $State['DAGName'], $_.Exception.Message)
            exit $ERR_DAGJOIN
        }
    }

    function Invoke-HealthChecker {
        if ($State['SkipHealthCheck']) {
            Write-MyVerbose 'SkipHealthCheck specified, skipping HealthChecker'
            return
        }

        Write-MyStep -Label 'HealthChecker' -Value 'running' -Status Run
        $hcPath = Join-Path $State['SourcesPath'] 'HealthChecker.ps1'
        $hcUrl = 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/HealthChecker.ps1'

        # Download if not present
        if (-not (Test-Path $hcPath)) {
            $downloaded = $false
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                try {
                    Write-MyVerbose ('Downloading HealthChecker from {0} (attempt {1}/3)' -f $hcUrl, $attempt)
                    Start-BitsTransfer -Source $hcUrl -Destination $hcPath -ErrorAction Stop
                    $downloaded = $true
                    break
                }
                catch {
                    if ($attempt -eq 3) {
                        try {
                            Invoke-WebDownload -Uri $hcUrl -OutFile $hcPath
                            $downloaded = $true
                        }
                        catch {
                            Write-MyWarning ('Could not download HealthChecker after 3 attempts: {0}' -f $_.Exception.Message)
                        }
                    }
                    else {
                        Start-Sleep -Seconds ($attempt * 5)
                    }
                }
            }
            if ($downloaded -and (Test-Path $hcPath)) {
                $hash = (Get-FileHash -Path $hcPath -Algorithm SHA256).Hash
                Write-MyVerbose ('HealthChecker downloaded, SHA256: {0}' -f $hash)
            }
            elseif (-not $downloaded) {
                return
            }
        }

        if (Test-Path $hcPath) {
            try {
                # HC writes ExchangeAllServersReport-*.html to the *current directory*, not -OutputFilePath.
                # Push-Location so both the XML (-OutputFilePath) and the HTML land in ReportsPath.
                Push-Location $State['ReportsPath']
                $hcBefore = [datetime]::Now
                & $hcPath -OutputFilePath $State['ReportsPath'] -SkipVersionCheck *>&1 | Out-Null
                & $hcPath -BuildHtmlServersReport -SkipVersionCheck *>&1 | Out-Null
                Pop-Location
                $hcReport = Get-ChildItem -Path $State['ReportsPath'] -ErrorAction SilentlyContinue |
                    Where-Object { $_.LastWriteTime -ge $hcBefore -and $_.Extension -match '\.html?' -and $_.Name -match '^(ExchangeAllServersReport|HealthChecker|HCExchangeServerReport)' } |
                    Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($hcReport) {
                    # Rename to SERVER_HCExchangeServerReport-<timestamp>.html
                    $hcTimestamp = $hcReport.Name -replace '^(?:ExchangeAllServersReport|HealthChecker|HCExchangeServerReport)-', ''
                    $newHcName   = '{0}_HCExchangeServerReport-{1}' -f $env:COMPUTERNAME, $hcTimestamp
                    $newHcPath   = Join-Path $State['ReportsPath'] $newHcName
                    try {
                        Rename-Item -Path $hcReport.FullName -NewName $newHcName -ErrorAction Stop
                        $State['HCReportPath'] = $newHcPath
                        Write-MyStep -Label 'HC report' -Value $newHcPath -Status OK
                    }
                    catch {
                        $State['HCReportPath'] = $hcReport.FullName
                        Write-MyStep -Label 'HC report' -Value $hcReport.FullName -Status OK
                    }
                } else {
                    Write-MyStep -Label 'HealthChecker' -Value ('completed (reports in {0})' -f $State['ReportsPath']) -Status OK
                }
                # On Domain Controllers there are no local security groups — the SAM database is
                # replaced by AD. HC's "Exchange Server Membership" check enumerates Win32_GroupUser
                # for the local "Exchange Servers" / "Exchange Trusted Subsystem" groups, which don't
                # exist on DCs, so it always reports "failed/blank" regardless of AD group membership.
                # This is a HC limitation; the server IS a member via the domain group.
                $dcRole = try { (Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue).DomainRole } catch { 3 }
                if ($dcRole -ge 4) {
                    Write-MyWarning 'NOTE: This server is a Domain Controller. HC "Exchange Server Membership" will show failed/blank — DCs have no local security groups. Exchange group membership is via AD domain groups and is correct.'
                }
            }
            catch {
                Pop-Location -ErrorAction SilentlyContinue
                Write-MyWarning ('HealthChecker execution failed: {0}' -f $_.Exception.Message)
            }
        }
    }

    function Invoke-SetupAssist {
        if ($State['SkipSetupAssist']) {
            Write-MyVerbose 'SkipSetupAssist specified, skipping SetupAssist'
            return
        }

        Write-MyStep -Label 'SetupAssist' -Value 'running (diagnosing setup failure)' -Status Run
        $saPath = Join-Path $State['SourcesPath'] 'SetupAssist.ps1'
        $saUrl  = 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/SetupAssist.ps1'

        if (-not (Test-Path $saPath)) {
            $downloaded = $false
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                try {
                    Write-MyVerbose ('Downloading SetupAssist from {0} (attempt {1}/3)' -f $saUrl, $attempt)
                    Start-BitsTransfer -Source $saUrl -Destination $saPath -ErrorAction Stop
                    $downloaded = $true
                    break
                }
                catch {
                    if ($attempt -eq 3) {
                        try {
                            Invoke-WebDownload -Uri $saUrl -OutFile $saPath
                            $downloaded = $true
                        }
                        catch {
                            Write-MyWarning ('Could not download SetupAssist after 3 attempts: {0}' -f $_.Exception.Message)
                        }
                    }
                    else {
                        Start-Sleep -Seconds ($attempt * 5)
                    }
                }
            }
            if ($downloaded -and (Test-Path $saPath)) {
                Write-MyVerbose ('SetupAssist downloaded, SHA256: {0}' -f (Get-FileHash -Path $saPath -Algorithm SHA256).Hash)
            }
            elseif (-not $downloaded) {
                return
            }
        }

        if (Test-Path $saPath) {
            try {
                & $saPath
            }
            catch {
                Write-MyWarning ('SetupAssist execution failed: {0}' -f $_.Exception.Message)
            }
        }

        # SetupLogReviewer — additional log analysis tool
        $slrPath = Join-Path $State['SourcesPath'] 'SetupLogReviewer.ps1'
        $slrUrl  = 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/SetupLogReviewer.ps1'

        if (-not (Test-Path $slrPath)) {
            $downloaded = $false
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                try {
                    Write-MyVerbose ('Downloading SetupLogReviewer from {0} (attempt {1}/3)' -f $slrUrl, $attempt)
                    Start-BitsTransfer -Source $slrUrl -Destination $slrPath -ErrorAction Stop
                    $downloaded = $true
                    break
                }
                catch {
                    if ($attempt -eq 3) {
                        try {
                            Invoke-WebDownload -Uri $slrUrl -OutFile $slrPath
                            $downloaded = $true
                        }
                        catch {
                            Write-MyWarning ('Could not download SetupLogReviewer after 3 attempts: {0}' -f $_.Exception.Message)
                        }
                    }
                    else {
                        Start-Sleep -Seconds ($attempt * 5)
                    }
                }
            }
            if ($downloaded -and (Test-Path $slrPath)) {
                Write-MyVerbose ('SetupLogReviewer downloaded, SHA256: {0}' -f (Get-FileHash -Path $slrPath -Algorithm SHA256).Hash)
            }
        }

        if (Test-Path $slrPath) {
            try {
                Write-MyStep -Label 'SetupLogReviewer' -Value 'running (analyzing setup logs)' -Status Run
                & $slrPath
            }
            catch {
                Write-MyWarning ('SetupLogReviewer execution failed: {0}' -f $_.Exception.Message)
            }
        }
    }

    function Test-AuthCertificate {
        try {
            $authConfig = Get-AuthConfig -ErrorAction Stop
            if (-not $authConfig) {
                Write-MyVerbose 'Test-AuthCertificate: Get-AuthConfig returned null — Exchange PS session may not be fully initialized'
                return
            }
            $thumbprint = $authConfig.CurrentCertificateThumbprint
            if (-not $thumbprint) {
                Write-MyWarning 'Exchange Auth Certificate: no thumbprint configured in AuthConfig'
                return
            }
            $cert = Get-ExchangeCertificate -Thumbprint $thumbprint -ErrorAction SilentlyContinue
            if (-not $cert) {
                Write-MyWarning ('Exchange Auth Certificate (thumbprint {0}) not found on this server' -f $thumbprint)
                return
            }
            # Exchange Auth Certificate (self-signed virtual cert) may return DateTime.MinValue for NotAfter.
            # Guard before arithmetic and .ToString() to avoid "cannot call method on null-valued expression".
            $notAfter = $cert.NotAfter
            if (-not $notAfter -or $notAfter -le [datetime]'1970-01-01') {
                Write-MyWarning ('Exchange Auth Certificate (thumbprint {0}) has no valid expiry date — verify with Get-AuthConfig / Get-ExchangeCertificate' -f $thumbprint)
                return
            }
            $daysLeft = ($notAfter - (Get-Date)).Days
            if ($daysLeft -le 0) {
                Write-MyWarning ('Exchange Auth Certificate EXPIRED {0} day(s) ago (expires {1}, thumbprint {2}). Renew: New-ExchangeCertificate, then Set-AuthConfig -NewCertificateThumbprint / -PublishCertificate' -f [Math]::Abs($daysLeft), $notAfter.ToString('yyyy-MM-dd'), $thumbprint)
            }
            elseif ($daysLeft -le 60) {
                Write-MyWarning ('Exchange Auth Certificate expires in {0} days on {1} (thumbprint {2}). Renew soon: New-ExchangeCertificate, then Set-AuthConfig -NewCertificateThumbprint / -PublishCertificate' -f $daysLeft, $notAfter.ToString('yyyy-MM-dd'), $thumbprint)
            }
            else {
                Write-MyStep -Label 'Auth Cert' -Value ('valid {0}d (expires {1})' -f $daysLeft, $notAfter.ToString('yyyy-MM-dd'))
            }
        }
        catch {
            Write-MyVerbose ('Test-AuthCertificate: {0}' -f $_.Exception.Message)
        }
    }

    function Test-DAGReplicationHealth {
        # F8: Validates mailbox database copy replication after DAG join.
        if (-not $State['DAGName']) { Write-MyVerbose 'No DAG configured, skipping replication health check'; return }
        Write-MyVerbose ('Checking DAG database copy replication health on {0}' -f $env:computername)
        try {
            $copies = @(Get-MailboxDatabaseCopyStatus -Server $env:computername -ErrorAction Stop)
            if ($copies.Count -eq 0) { Write-MyVerbose 'No mailbox database copies found on this server'; return }
            $warns = 0
            foreach ($copy in $copies) {
                $ok  = $copy.Status -in 'Mounted', 'Healthy'
                $msg = 'DB copy {0}: Status={1}, CopyQueue={2}, ReplayQueue={3}' -f $copy.DatabaseName, $copy.Status, $copy.CopyQueueLength, $copy.ReplayQueueLength
                if ($ok) { Write-MyVerbose $msg } else { Write-MyWarning $msg; $warns++ }
            }
            if ($warns -eq 0) {
                Write-MyStep -Label 'DAG replication' -Value ('{0} copy/copies OK' -f $copies.Count) -Status OK
            }
            else {
                Write-MyWarning ('{0} database copy/copies not healthy — review replication status' -f $warns)
            }
        }
        catch {
            Write-MyWarning ('DAG replication health check failed: {0}' -f $_.Exception.Message)
        }
    }

    function Test-VSSWriters {
        # F9: Checks all VSS writers are in a stable state. Unstable writers can break Exchange online backup.
        Write-MyVerbose 'Checking VSS writer health'
        try {
            $output = & vssadmin.exe list writers 2>&1
            $currentWriter = ''
            $warns = 0
            foreach ($line in $output) {
                if ($line -match "Writer name:\s+'(.+)'") { $currentWriter = $Matches[1] }
                elseif ($line -match 'State:\s*\[\d+\]\s+(.+)') {
                    $stateText = $Matches[1].Trim()
                    if ($stateText -notmatch '^Stable') {
                        Write-MyWarning ('VSS Writer "{0}": {1}' -f $currentWriter, $stateText)
                        $warns++
                    }
                }
            }
            if ($warns -eq 0) { Write-MyVerbose 'All VSS writers are stable' }
            else { Write-MyWarning ('{0} VSS writer(s) not stable — check Volume Shadow Copy Service' -f $warns) }
        }
        catch {
            Write-MyWarning ('VSS writer check failed: {0}' -f $_.Exception.Message)
        }
    }

    function Test-EEMSStatus {
        # F10: Exchange Emergency Mitigation Service (EEMS) — available from Exchange 2019 CU11+ and SE.
        # EEMS applies automatic security mitigations for critical CVEs before patches are available.
        $svc = Get-Service MSExchangeMitigation -ErrorAction SilentlyContinue
        if (-not $svc) { Write-MyVerbose 'EEMS service not present (Exchange 2016 or 2019 pre-CU11)'; return }
        $statusLabel = if ($svc.Status -eq 'Running') { 'Running (OK)' } else { $svc.Status.ToString() }
        Write-MyStep -Label 'EEMS (Emergency Mitigation)' -Value $statusLabel
        if ($svc.Status -ne 'Running') {
            Write-MyWarning 'EEMS is not running — automatic CVE mitigations will not be applied'
        }
        try {
            $orgCfg = Get-OrganizationConfig -ErrorAction Stop
            if ($orgCfg.PSObject.Properties['MitigationsEnabled']) {
                if (-not $orgCfg.MitigationsEnabled) {
                    Write-MyWarning 'EEMS mitigations disabled org-wide (Set-OrganizationConfig -MitigationsEnabled $true to re-enable)'
                }
                else {
                    Write-MyVerbose ('EEMS mitigations enabled: {0}' -f $orgCfg.MitigationsEnabled)
                }
                $blocked = $orgCfg.MitigationsBlocked
                if ($blocked) {
                    Write-MyWarning ('EEMS blocked mitigations: {0}' -f ($blocked -join ', '))
                }
            }
        }
        catch {
            Write-MyVerbose ('EEMS org config check: {0}' -f $_.Exception.Message)
        }
    }

    function Get-ModernAuthReport {
        # F11: Verifies Modern Authentication (OAuth2) is enabled. Required for Outlook 2016+,
        # Microsoft Teams, mobile clients, and any Hybrid / Azure AD configuration.
        Write-MyVerbose 'Checking Modern Authentication (OAuth2) configuration'
        try {
            $orgCfg = Get-OrganizationConfig -ErrorAction Stop
            if ($orgCfg.OAuth2ClientProfileEnabled) {
                Write-MyStep -Label 'Modern Authentication' -Value 'enabled' -Status OK
            }
            else {
                Write-MyWarning 'Modern Authentication (OAuth2) is DISABLED — required for Outlook 2016+, Teams, mobile clients, and Hybrid. Enable: Set-OrganizationConfig -OAuth2ClientProfileEnabled $true'
            }
        }
        catch {
            Write-MyVerbose ('Modern Auth report: {0}' -f $_.Exception.Message)
        }
    }

