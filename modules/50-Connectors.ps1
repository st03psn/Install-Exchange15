    function Get-MEACAutomationCredentialFromState {
        # Rehydrates -MEACAutomationCredential across the Autopilot reboot chain.
        # Only used for AD Split-Permissions deployments where a Domain Admin has
        # pre-created the SystemMailbox{b963af59-…} account via -MEACPrepareADOnly
        # and the credential needs to survive reboots between Phase 0 and Phase 6.
        # Standard (non-Split) deployments never populate this — MEAC self-provisions.
        if (-not $State['MEACAutomationUser'] -or -not $State['MEACAutomationPW']) { return $null }
        try {
            $sec = ConvertTo-SecureString $State['MEACAutomationPW'] -ErrorAction Stop
            return New-Object System.Management.Automation.PSCredential($State['MEACAutomationUser'], $sec)
        }
        catch {
            Write-MyWarning ('MEAC: could not decrypt stored automation credential (DPAPI mismatch?): {0}' -f $_.Exception.Message)
            return $null
        }
    }

    function Register-AuthCertificateRenewal {
        # MEAC — CSS-Exchange MonitorExchangeAuthCertificate.ps1. Creates a daily
        # scheduled task that auto-renews the Exchange Auth Certificate 60 days
        # before expiry. Without it, Auth Cert expiry causes a full outage
        # (OAuth, Hybrid, EWS). Skip on Edge and Management-only installs.
        #
        # v5.93: MEAC self-provisions SystemMailbox{b963af59-…}, the Auth Certificate
        # Management role group, and the batch-logon right. EXpress does not layer a
        # credential system on top — but MEAC still needs a password supplied at
        # registration time so Task Scheduler can run the daily task AS that user
        # (Task Scheduler refuses to register a task-under-user-identity without a
        # credential). EXpress generates a strong random password inline and passes
        # it via MEAC -Password; it is never persisted.
        #
        # v5.93: hybrid-aware + documented passthroughs. Detects hybrid via
        # Get-HybridConfiguration; in hybrid mode MEAC refuses to renew by default
        # (to avoid breaking HCW federation silently) — operator opts into renewal
        # via -MEACIgnoreHybridConfig with the implicit promise to rerun HCW.
        # Split-Permissions path: DA pre-creates the account (separate run with
        # -MEACPrepareADOnly); Exchange admin passes the resulting credential here
        # as -MEACAutomationCredential, forwarded to MEAC -AutomationAccountCredential.
        if ($State['InstallEdge'] -or $State['InstallManagementTools'] -or $State['InstallRecipientManagement']) {
            Write-MyVerbose 'Auth Certificate renewal task: not applicable to this install mode'
            return
        }
        # Idempotency: on re-run after a failure, skip MEAC entirely when both
        # the scheduled task and the auto-provisioned automation account already
        # exist. Re-invoking MEAC in this state is redundant and can surface
        # spurious errors from the registration path.
        $existingTask = Get-ScheduledTask -TaskName 'Daily Auth Certificate Check' -ErrorAction SilentlyContinue
        $existingUser = $null
        if (Get-Command Get-User -ErrorAction SilentlyContinue) {
            $existingUser = Get-User -Identity 'SystemMailbox{b963af59-3975-4f92-9d58-ad0b1fe3a1a3}' -ErrorAction SilentlyContinue
        }
        if ($existingTask -and $existingUser) {
            Write-MyOutput 'Auth Certificate renewal already registered (scheduled task + automation account present) - skipping MEAC'
            return
        }
        Write-MyOutput 'Registering Auth Certificate renewal (MEAC / CSS-Exchange MonitorExchangeAuthCertificate.ps1)'
        $meacPath = Join-Path $State['SourcesPath'] 'MonitorExchangeAuthCertificate.ps1'
        $meacUrl  = 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/MonitorExchangeAuthCertificate.ps1'
        if (-not (Test-Path $meacPath)) {
            try {
                Invoke-WebDownload -Uri $meacUrl -OutFile $meacPath
                Write-MyVerbose ('MEAC downloaded, SHA256: {0}' -f (Get-FileHash $meacPath -Algorithm SHA256).Hash)
            }
            catch {
                Write-MyWarning ('Could not download MonitorExchangeAuthCertificate.ps1: {0}' -f $_.Exception.Message)
                return
            }
        }

        # --- Hybrid detection (transparent) ---------------------------------------
        # Get-HybridConfiguration returns an object only when HCW has been run.
        # Blank / non-existent object → not hybrid. Errors are treated as not-hybrid
        # rather than blocking, because MEAC itself re-detects authoritatively.
        $hybridDetected = $false
        if (Get-Command Get-HybridConfiguration -ErrorAction SilentlyContinue) {
            try {
                $hc = Get-HybridConfiguration -ErrorAction Stop
                $hybridDetected = [bool]($hc -and ($hc.Domains -or $hc.Features -or $hc.ClientAccessServers -or $hc.TransportServers))
            } catch { $hybridDetected = $false }
        }

        # --- Build parameter set for MEAC -----------------------------------------
        $meacParams = @{
            ConfigureScriptToRunViaScheduledTask = $true
            Confirm                              = $false
        }

        # Split-Permissions: DA pre-created the automation account; pass it through.
        # Prefer state (survives Autopilot reboot chain) over current-run CLI parameter.
        $autoCred = if ($State['MEACAutomationUser']) { Get-MEACAutomationCredentialFromState } else { $MEACAutomationCredential }
        if ($autoCred) {
            $meacParams.AutomationAccountCredential = $autoCred
            Write-MyOutput ('MEAC: using pre-created automation account {0} (AD Split-Permissions passthrough)' -f $autoCred.UserName)
        }
        else {
            # Standard (non-Split) deployment: MEAC self-provisions the
            # SystemMailbox{b963af59-…} account, but Task Scheduler still needs a
            # password to register a task that runs as that user. MEAC accepts
            # either -Password <SecureString> or -AutomationAccountCredential; we
            # take the simpler path and generate a strong random password inline.
            #
            # The password is transient: once MEAC sets it on the account AND
            # registers the scheduled task, Windows stores the credential in the
            # task's (DPAPI-protected) credential store. EXpress never needs it
            # again — if re-registration becomes necessary (e.g. password policy
            # rotation), re-running this function generates a fresh password and
            # MEAC resets both the account and the task atomically.
            #
            # Therefore: not persisted to state, not logged, not returned.
            $charset = [char[]](
                (65..90) +             # A-Z
                (97..122) +            # a-z
                (50..57) +             # 2-9  (skip 0/1 for clarity)
                @(33, 35, 36, 37, 38, 42, 43, 45, 61, 63, 64)   # ! # $ % & * + - = ? @
            )
            $pwBytes = [byte[]]::new(32)
            [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($pwBytes)
            $pwChars = foreach ($b in $pwBytes) { $charset[$b % $charset.Length] }
            $pwSecure = ConvertTo-SecureString -String (-join $pwChars) -AsPlainText -Force
            $meacParams.Password = $pwSecure
            Remove-Variable pwBytes, pwChars -ErrorAction SilentlyContinue
            Write-MyVerbose 'MEAC: generated transient 32-char password for SystemMailbox{b963af59-…} automation account (not persisted; Task Scheduler stores it DPAPI-protected)'
        }

        if ($MEACIgnoreHybridConfig) {
            $meacParams.IgnoreHybridConfig = $true
            Write-MyWarning 'MEAC: -IgnoreHybridConfig enabled. MEAC will renew the Auth Certificate when due; you MUST rerun the Hybrid Configuration Wizard afterwards or OAuth/federation with Exchange Online will break.'
        }
        if ($MEACIgnoreUnreachableServers) {
            $meacParams.IgnoreUnreachableServers = $true
            Write-MyVerbose 'MEAC: -IgnoreUnreachableServers enabled'
        }
        if ($MEACNotificationEmail) {
            $meacParams.SendEmailNotificationTo = $MEACNotificationEmail
            Write-MyOutput ('MEAC: renewal notifications will be sent to {0}' -f $MEACNotificationEmail)
        }

        # Hybrid advisory — transparent to the operator.
        if ($hybridDetected -and -not $MEACIgnoreHybridConfig) {
            Write-MyOutput ''
            Write-MyOutput 'MEAC: Hybrid configuration detected.'
            Write-MyOutput '      Task registered in hybrid-safe mode — MEAC will REFUSE to renew the'
            Write-MyOutput '      Auth Certificate until -MEACIgnoreHybridConfig is passed (which also'
            Write-MyOutput '      obliges you to rerun HCW afterwards to re-federate with Exchange Online).'
            Write-MyOutput '      Without the flag, daily checks still run; wire up -MEACNotificationEmail'
            Write-MyOutput '      to receive an alert 60 days before expiry.'
            Write-MyOutput ''
        }

        # --- Run MEAC -------------------------------------------------------------
        try {
            Push-Location $State['InstallPath']
            $meacArgStr = ($meacParams.GetEnumerator() | ForEach-Object {
                $v = if ($_.Value -is [System.Security.SecureString]) { '<SecureString>' }
                     elseif ($_.Value -is [System.Management.Automation.PSCredential]) { '<PSCredential>' }
                     elseif ($_.Value -is [bool] -or $_.Value -is [switch]) { '' }
                     else { "'$($_.Value)'" }
                if ($v) { "-$($_.Key) $v" } else { "-$($_.Key)" }
            }) -join ' '
            Register-ExecutedCommand -Category 'ScheduledTask' -Command (".\MonitorExchangeAuthCertificate.ps1 $meacArgStr")
            & $meacPath @meacParams *>&1 | ForEach-Object { Write-MyVerbose ('MEAC: {0}' -f $_) }
        }
        catch {
            Write-MyWarning ('MEAC registration failed: {0}' -f $_.Exception.Message)
            return
        }
        finally {
            Pop-Location
        }
        $meacTask = Get-ScheduledTask -TaskName 'Daily Auth Certificate Check' -ErrorAction SilentlyContinue
        if ($meacTask) {
            Write-MyOutput 'MEAC scheduled task registered — auth cert will auto-renew 60 days before expiry'
        }
        else {
            Write-MyWarning 'MEAC: task "Daily Auth Certificate Check" not found after registration — check MEAC log in Exchange Logging\AuthCertificateMonitoring\ for details'
        }
    }

    function Add-ServerToSendConnectors {
        if ($State['InstallEdge']) {
            Write-MyVerbose 'Edge role — skipping Send Connector update'
            return
        }
        try {
            $sendConnectors = Get-SendConnector -ErrorAction Stop | Where-Object {
                $srvList = @($_.SourceTransportServers | ForEach-Object { $_.Name })
                $srvList -notcontains $env:COMPUTERNAME
            }
            if (-not $sendConnectors -or $sendConnectors.Count -eq 0) {
                Write-MyVerbose 'All Send Connectors already include this server'
                return
            }
            Write-MyOutput ('{0} Send Connector(s) do not include this server:' -f $sendConnectors.Count)
            foreach ($sc in $sendConnectors) {
                Write-MyOutput ('  - {0}' -f $sc.Name)
            }
            $answer = 'Y'
            if ([Environment]::UserInteractive) {
                Write-MyOutput 'Add this server as source transport server? [Y/N/S=skip] (default: Y):'
                try {
                    $key = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                    $answer = $key.Character.ToString().ToUpper()
                }
                catch { $answer = 'Y' }
                Write-MyOutput $answer
            }
            if ($answer -notin @('N', 'S')) {
                foreach ($sc in $sendConnectors) {
                    $sources = [System.Collections.Generic.List[object]]($sc.SourceTransportServers)
                    $sources.Add($env:COMPUTERNAME) | Out-Null
                    Register-ExecutedCommand -Category 'SendConnector' -Command ("Set-SendConnector -Identity '$($sc.Identity)' -SourceTransportServers $($sources -join ',')")
                    Set-SendConnector -Identity $sc.Identity -SourceTransportServers $sources -ErrorAction Stop
                    Write-MyOutput ('Added {0} to Send Connector: {1}' -f $env:COMPUTERNAME, $sc.Name)
                }
            }
            else {
                Write-MyVerbose 'Send Connector update skipped by user'
            }
        }
        catch {
            Write-MyWarning ('Failed to update Send Connectors: {0}' -f $_.Exception.Message)
        }
    }

    function Install-AntispamAgents {
        if ($State['InstallEdge']) {
            Write-MyVerbose 'Edge role has antispam agents built-in, skipping'
            return
        }
        $exSetup = (Get-ItemProperty -Path $EXCHANGEINSTALLKEY -ErrorAction SilentlyContinue).MsiInstallPath
        if (-not $exSetup) {
            Write-MyWarning 'Exchange install path not found — cannot install antispam agents'
            return
        }
        $installScript = Join-Path $exSetup 'Scripts\Install-AntispamAgents.ps1'
        if (-not (Test-Path $installScript)) {
            Write-MyWarning ('Install-AntispamAgents.ps1 not found at: {0}' -f $installScript)
            return
        }

        # Check if antispam agents are already installed by looking for a known agent
        $existingAgents = Get-TransportAgent -ErrorAction SilentlyContinue |
                          Where-Object { $_.Identity -like '*Filter*' -or $_.Identity -like '*Antispam*' }
        if (-not $existingAgents) {
            Write-MyOutput 'Installing Exchange antispam agents'
            # Capture everything that bypasses the pipeline (Write-Host / $host.UI.WriteWarningLine /
            # [Console]::WriteLine) by redirecting Console.Out + Console.Error to StringWriters.
            # Combined with *>&1 this catches both stream-based and host-UI-based output.
            $capOut    = [System.IO.StringWriter]::new()
            $capErr    = [System.IO.StringWriter]::new()
            $origOut   = [System.Console]::Out
            $origErr   = [System.Console]::Error
            $records   = $null
            # Suppress warnings in the current session scope: the Exchange PSSnapin autoload
            # ("Please exit Windows PowerShell", "restart MSExchangeTransport") fires in the
            # caller's session — not in the child script's streams — so *>&1 alone cannot
            # capture it. $WarningPreference = SilentlyContinue mutes this session-level noise.
            $savedWP = $WarningPreference
            $WarningPreference = 'SilentlyContinue'
            try {
                [System.Console]::SetOut($capOut)
                [System.Console]::SetError($capErr)
                $records = & $installScript *>&1
            }
            catch {
                [System.Console]::SetOut($origOut)
                [System.Console]::SetError($origErr)
                Write-MyWarning ('Failed to install antispam agents: {0}' -f $_.Exception.Message)
                return
            }
            finally {
                [System.Console]::SetOut($origOut)
                [System.Console]::SetError($origErr)
                $WarningPreference = $savedWP
            }

            # Route pipeline records: warnings/verbose to debug log, Exchange agent objects
            # collected into a single summary list, everything else to stdout.
            # Suppresses the well-known "Please restart MSExchangeTransport" warning from
            # the Exchange-shipped Install-AntispamAgents.ps1 (we restart the service below).
            $agentSummary = [System.Collections.Generic.List[string]]::new()
            foreach ($r in $records) {
                if ($null -eq $r) { continue }
                if ($r -is [System.Management.Automation.WarningRecord]) {
                    $msg = $r.Message
                    # We restart MSExchangeTransport below, so suppress the "restart required"
                    # and PSSnapin-autoload "exit Windows PowerShell" warnings shipped with
                    # Install-AntispamAgents.ps1.
                    if ($msg -match '(restart is required|restart the Microsoft Exchange Transport|exit Windows PowerShell to complete)') { continue }
                    Write-MyDebug ('[antispam] WARN: {0}' -f $msg)
                }
                elseif ($r -is [System.Management.Automation.ErrorRecord]) {
                    Write-MyWarning ('[antispam] {0}' -f $r.Exception.Message)
                }
                elseif ($r -is [System.Management.Automation.VerboseRecord] -or
                        $r -is [System.Management.Automation.DebugRecord]) {
                    Write-MyDebug ('[antispam] {0}' -f $r.Message)
                }
                elseif ($r -is [System.Management.Automation.InformationRecord]) {
                    Write-MyDebug ('[antispam] {0}' -f $r.MessageData)
                }
                else {
                    # TransportAgent objects — each would otherwise render as its own table.
                    # Collect a compact one-liner per agent instead.
                    $idName = if ($r.PSObject.Properties['Identity']) { [string]$r.Identity } else { '' }
                    if ($idName) {
                        $prio    = if ($r.PSObject.Properties['Priority']) { $r.Priority } else { '-' }
                        $enabled = if ($r.PSObject.Properties['Enabled'])  { $r.Enabled  } else { '-' }
                        $agentSummary.Add(('  {0,-32} Priority={1,-3} Enabled={2}' -f $idName, $prio, $enabled))
                    }
                    else {
                        $line = ($r | Out-String).TrimEnd()
                        if ($line) { Write-MyDebug ('[antispam] {0}' -f $line) }
                    }
                }
            }
            if ($agentSummary.Count -gt 0) {
                Write-MyOutput ('[antispam] Installed {0} transport agents:' -f $agentSummary.Count)
                foreach ($s in $agentSummary) { Write-MyOutput $s }
            }

            # Route host-UI output captured via Console redirection — demote to debug.
            foreach ($captured in @($capOut.ToString(), $capErr.ToString())) {
                if (-not $captured) { continue }
                foreach ($line in ($captured -split "`r?`n")) {
                    if (-not $line) { continue }
                    if ($line -match 'restart the Microsoft Exchange Transport') { continue }
                    Write-MyDebug ('[antispam] {0}' -f $line.TrimEnd())
                }
            }

            $installDidRun = $true
        }
        else {
            Write-MyVerbose ('Antispam agents already installed ({0} found), skipping install script' -f $existingAgents.Count)
            $installDidRun = $false
        }

        # Configure agents: only Recipient Filter Agent enabled, all other antispam
        # agents shipped by Install-AntispamAgents.ps1 disabled.
        #
        # Enumerate via Get-TransportAgent instead of hard-coded names — the exact
        # Identity casing varies across CUs/locales ("Sender Id Agent" vs
        # "Sender ID Agent"); a hard-coded lookup left Sender ID / Protocol Analysis
        # enabled in the field. Match by regex against all antispam agent names
        # and pass the actual Identity back to Enable/Disable-TransportAgent.
        $antispamRegex  = '(Content Filter|Sender\s*Id|Sender Filter|Recipient Filter|Protocol Analysis)'
        $recipientRegex = 'Recipient Filter'

        $configChanged = $false
        $savedWP       = $WarningPreference
        $WarningPreference = 'SilentlyContinue'   # mutes session-level "restart required" /
                                                  # "exit Windows PowerShell" host-UI warnings
                                                  # emitted by Enable-/Disable-TransportAgent
        try {
            $allAgents = @(Get-TransportAgent -ErrorAction SilentlyContinue)
            foreach ($agent in $allAgents) {
                $id = [string]$agent.Identity
                if ($id -notmatch $antispamRegex) { continue }
                $wantEnabled = ($id -match $recipientRegex)
                if ($agent.Enabled -eq $wantEnabled) {
                    Write-MyVerbose ('Already {0}: {1}' -f ($(if ($wantEnabled) {'enabled'} else {'disabled'})), $id)
                    continue
                }
                if ($wantEnabled) {
                    Register-ExecutedCommand -Category 'Antispam' -Command ("Enable-TransportAgent -Identity '$id'")
                    Enable-TransportAgent  -Identity $id -Confirm:$false -WarningAction SilentlyContinue -ErrorAction SilentlyContinue *>&1 | Out-Null
                    Write-MyOutput ('Enabled: {0}' -f $id)
                }
                else {
                    Register-ExecutedCommand -Category 'Antispam' -Command ("Disable-TransportAgent -Identity '$id'")
                    Disable-TransportAgent -Identity $id -Confirm:$false -WarningAction SilentlyContinue -ErrorAction SilentlyContinue *>&1 | Out-Null
                    Write-MyOutput ('Disabled: {0}' -f $id)
                }
                $configChanged = $true
            }
        }
        finally {
            $WarningPreference = $savedWP
        }

        # Enable the recipient lookup against the GAL. Without this the Recipient
        # Filter Agent only applies block-list / tarpit logic — the main value
        # (rejecting mail for non-existent recipients on Authoritative domains)
        # stays off. Accepted Domains with DomainType=Authoritative are covered
        # implicitly; Internal/External Relay domains are skipped by design.
        try {
            $rfc = Get-RecipientFilterConfig -ErrorAction Stop
            if (-not $rfc.RecipientValidationEnabled) {
                Register-ExecutedCommand -Category 'Antispam' -Command 'Set-RecipientFilterConfig -RecipientValidationEnabled $true'
                Set-RecipientFilterConfig -RecipientValidationEnabled $true -Confirm:$false -ErrorAction Stop
                Write-MyOutput 'Enabled recipient lookup (RecipientFilterConfig.RecipientValidationEnabled = True)'
                $configChanged = $true
            }
            else {
                Write-MyVerbose 'Recipient lookup already enabled'
            }
        }
        catch {
            Write-MyWarning ('Could not configure RecipientFilterConfig: {0}' -f $_.Exception.Message)
        }

        # One restart, at the end — covers both the install (if it ran) and any
        # enable/disable changes. Nothing downstream looks at the agents before
        # this point, so a single bounce is sufficient.
        if ($installDidRun -or $configChanged) {
            Write-MyOutput 'Restarting MSExchangeTransport (may take ~30s)'
            Restart-Service MSExchangeTransport -Force -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            Write-MyOutput 'MSExchangeTransport restarted'
        }
        Write-MyOutput 'Antispam agents configured: only Recipient Filter Agent is enabled'
    }

    function New-AnonymousRelayConnector {
        $hasInternal = $State['RelaySubnets']         -and $State['RelaySubnets'].Count         -gt 0
        $hasExternal = $State['ExternalRelaySubnets'] -and $State['ExternalRelaySubnets'].Count -gt 0

        if (-not $hasInternal -and -not $hasExternal) {
            Write-MyVerbose 'No RelaySubnets or ExternalRelaySubnets specified, skipping relay connector setup'
            return
        }
        if ($State['InstallEdge']) {
            Write-MyVerbose 'Edge role — skipping relay connector setup (use EdgeSync for relay)'
            return
        }

        $server  = $env:COMPUTERNAME
        $success = $true

        # --- Internal relay connector (no Ms-Exch-SMTP-Accept-Any-Recipient) ---
        # Anonymous senders can deliver to accepted domains only; cannot relay externally.
        # RFC 5737 TEST-NET placeholder — never routable, used when no subnets were specified.
        # Prevents the relay connector from matching real SMTP traffic until the admin sets proper IPs.
        $RELAY_PLACEHOLDER          = '192.0.2.1/32'
        $RELAY_PLACEHOLDER_EXTERNAL = '192.0.2.2/32'   # Different RFC 5737 address — avoids Bindings+RemoteIPRanges conflict
        $internalIsPlaceholder = ($State['RelaySubnets'].Count -eq 1 -and $State['RelaySubnets'][0] -eq $RELAY_PLACEHOLDER)
        $externalIsPlaceholder = ($State['ExternalRelaySubnets'].Count -eq 1 -and $State['ExternalRelaySubnets'][0] -in @($RELAY_PLACEHOLDER, $RELAY_PLACEHOLDER_EXTERNAL))

        if ($hasInternal) {
            $intName    = ('Anonymous Internal Relay - {0}' -f $server)
            $subnetList = $State['RelaySubnets'] -join ', '
            if ($internalIsPlaceholder) {
                Write-MyWarning 'Internal relay connector: no subnets specified — using placeholder IP (192.0.2.1/32, RFC 5737).'
                Write-MyWarning 'No real SMTP traffic will match this connector until you set RemoteIPRanges to your actual relay sources.'
            }
            Write-MyOutput ('Configuring internal relay connector "{0}" — subnets: {1}' -f $intName, $subnetList)
            try {
                $existing = Get-ReceiveConnector -Identity "$server\$intName" -ErrorAction SilentlyContinue
                if ($existing) {
                    Register-ExecutedCommand -Category 'ReceiveConnector' -Command ("Set-ReceiveConnector -Identity '$server\$intName' -RemoteIPRanges $($State['RelaySubnets'] -join ',') -AuthMechanism Tls -ProtocolLoggingLevel Verbose -Banner '220 Mail Service'")
                    Set-ReceiveConnector -Identity "$server\$intName" -RemoteIPRanges $State['RelaySubnets'] `
                        -AuthMechanism Tls -ProtocolLoggingLevel Verbose -Banner '220 Mail Service' -ErrorAction Stop
                    Write-MyVerbose 'Internal relay connector already exists — RemoteIPRanges, TLS, logging and banner updated'
                }
                else {
                    Register-ExecutedCommand -Category 'ReceiveConnector' -Command ("New-ReceiveConnector -Name '$intName' -Server '$server' -TransportRole FrontendTransport -RemoteIPRanges $($State['RelaySubnets'] -join ',') -Bindings '0.0.0.0:25' -PermissionGroups AnonymousUsers -AuthMechanism Tls -ProtocolLoggingLevel Verbose -Banner '220 Mail Service'")
                    New-ReceiveConnector -Name $intName -Server $server -TransportRole FrontendTransport `
                        -RemoteIPRanges $State['RelaySubnets'] -Bindings '0.0.0.0:25' `
                        -PermissionGroups AnonymousUsers -AuthMechanism Tls `
                        -ProtocolLoggingLevel Verbose -Banner '220 Mail Service' -ErrorAction Stop | Out-Null
                    Write-MyOutput 'Internal relay connector created (TLS offered, accepted domains only, no external relay right, hardened banner)'
                }
            }
            catch {
                Write-MyWarning ('Failed to configure internal relay connector: {0}' -f $_.Exception.Message)
                $success = $false
            }
        }

        # --- External relay connector (Ms-Exch-SMTP-Accept-Any-Recipient granted) ---
        # Anonymous senders from these IPs can relay to any recipient including external.
        if ($hasExternal) {
            $extName = ('Anonymous External Relay - {0}' -f $server)
            Write-MyWarning 'SECURITY: External relay connector allows anonymous relay to ANY recipient.'
            if ($externalIsPlaceholder) {
                Write-MyWarning 'External relay connector: no subnets specified — using placeholder IP (192.0.2.2/32, RFC 5737).'
                Write-MyWarning 'No real SMTP traffic will match this connector until you set RemoteIPRanges to your actual relay sources.'
            }
            Write-MyWarning ('         External relay subnets: {0}' -f ($State['ExternalRelaySubnets'] -join ', '))
            try {
                # Resolve ANONYMOUS LOGON by SID (S-1-5-7) — language-independent
                $anonLogon = ([System.Security.Principal.SecurityIdentifier]'S-1-5-7').Translate(
                    [System.Security.Principal.NTAccount]).Value
                Write-MyVerbose ('Resolved ANONYMOUS LOGON account: {0}' -f $anonLogon)

                $existing = Get-ReceiveConnector -Identity "$server\$extName" -ErrorAction SilentlyContinue
                $connObj = $null
                if ($existing) {
                    Register-ExecutedCommand -Category 'ReceiveConnector' -Command ("Set-ReceiveConnector -Identity '$server\$extName' -RemoteIPRanges $($State['ExternalRelaySubnets'] -join ',') -AuthMechanism Tls -ProtocolLoggingLevel Verbose -Banner '220 Mail Service'")
                    Set-ReceiveConnector -Identity "$server\$extName" -RemoteIPRanges $State['ExternalRelaySubnets'] `
                        -AuthMechanism Tls -ProtocolLoggingLevel Verbose -Banner '220 Mail Service' -ErrorAction Stop
                    Write-MyVerbose 'External relay connector already exists — RemoteIPRanges, TLS, logging and banner updated'
                    $connObj = $existing
                }
                else {
                    # Capture the returned object directly to avoid a race condition where
                    # Get-ReceiveConnector fails immediately after creation (Exchange AD not yet updated).
                    Register-ExecutedCommand -Category 'ReceiveConnector' -Command ("New-ReceiveConnector -Name '$extName' -Server '$server' -TransportRole FrontendTransport -RemoteIPRanges $($State['ExternalRelaySubnets'] -join ',') -Bindings '0.0.0.0:25' -PermissionGroups AnonymousUsers -AuthMechanism Tls -ProtocolLoggingLevel Verbose -Banner '220 Mail Service'")
                    $connObj = New-ReceiveConnector -Name $extName -Server $server -TransportRole FrontendTransport `
                        -RemoteIPRanges $State['ExternalRelaySubnets'] -Bindings '0.0.0.0:25' `
                        -PermissionGroups AnonymousUsers -AuthMechanism Tls `
                        -ProtocolLoggingLevel Verbose -Banner '220 Mail Service' -ErrorAction Stop
                }
                # Fallback: if the object is somehow null, retry Get-ReceiveConnector with backoff
                if (-not $connObj) {
                    for ($retry = 1; $retry -le 3 -and -not $connObj; $retry++) {
                        Write-MyVerbose ('Waiting for external relay connector to register in Exchange (attempt {0}/3)...' -f $retry)
                        Start-Sleep -Seconds 5
                        $connObj = Get-ReceiveConnector -Identity "$server\$extName" -ErrorAction SilentlyContinue
                    }
                }
                Register-ExecutedCommand -Category 'ReceiveConnector' -Command ("Get-ReceiveConnector '$server\$extName' | Add-ADPermission -User '$anonLogon' -ExtendedRights 'Ms-Exch-SMTP-Accept-Any-Recipient'")
                $connObj | Add-ADPermission -User $anonLogon `
                    -ExtendedRights 'Ms-Exch-SMTP-Accept-Any-Recipient' -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
                Write-MyOutput ('External relay connector created with Ms-Exch-SMTP-Accept-Any-Recipient for {0}' -f $anonLogon)
            }
            catch {
                Write-MyWarning ('Failed to configure external relay connector: {0}' -f $_.Exception.Message)
                $success = $false
            }
        }

        # --- Remove AnonymousUsers from Default Frontend connector ---
        # Only when at least one dedicated relay connector was configured successfully.
        # This prevents unauthenticated inbound from arbitrary IPs while keeping
        # relay restricted to the explicitly defined subnets above.
        # Skip Default Frontend hardening when only placeholder IPs are set — the relay connector
        # won't match real traffic yet, so removing AnonymousUsers would break inbound mail.
        $onlyPlaceholders = ($hasInternal -and $internalIsPlaceholder) -and (-not $hasExternal -or $externalIsPlaceholder)
        if ($success -and -not $onlyPlaceholders) {
            $defaultName = ('Default Frontend {0}' -f $server)
            try {
                $rc = Get-ReceiveConnector -Identity "$server\$defaultName" -ErrorAction SilentlyContinue
                if ($rc -and ($rc.PermissionGroups -match 'AnonymousUsers')) {
                    $pgList  = ($rc.PermissionGroups.ToString() -split ',\s*') | Where-Object { $_.Trim() -ne 'AnonymousUsers' }
                    Register-ExecutedCommand -Category 'ReceiveConnector' -Command ("Set-ReceiveConnector -Identity '$server\$defaultName' -PermissionGroups '$($pgList -join ',')'  # AnonymousUsers removed")
                    Set-ReceiveConnector -Identity "$server\$defaultName" -PermissionGroups ($pgList -join ',') -ErrorAction Stop
                    Write-MyOutput ('Removed AnonymousUsers from "{0}" receive connector' -f $defaultName)
                }
                else {
                    Write-MyVerbose ('AnonymousUsers already absent from "{0}"' -f $defaultName)
                }
            }
            catch {
                Write-MyWarning ('Failed to modify Default Frontend connector: {0}' -f $_.Exception.Message)
            }
        }
        elseif ($onlyPlaceholders) {
            Write-MyWarning ('Default Frontend connector NOT modified — relay connector uses placeholder IPs only. Set real RemoteIPRanges first, then remove AnonymousUsers from "{0}" manually.' -f ('Default Frontend {0}' -f $server))
        }
        else {
            Write-MyWarning 'One or more relay connectors failed — Default Frontend connector was NOT modified'
        }
    }

    function Enable-AccessNamespaceMailConfig {
        # F26 — Configure the Access Namespace as an Accepted Domain and update the
        # default Email Address Policy so that mailboxes get a primary SMTP address
        # @<AccessNamespace> (e.g. @mail.contoso.com).
        #
        # Steps:
        #  1. Add the Access Namespace as an Authoritative Accepted Domain (skip if present).
        #  2. Update the default Email Address Policy to make @<namespace> the primary
        #     SMTP template; retain the AD/internal domain as a secondary address.
        #  3. Remove pure-internal domains (.local / nonroutable) from the policy
        #     templates — they serve no purpose as email addresses.
        #  4. Apply the updated policy to all mailboxes via Update-EmailAddressPolicy.
        #
        # Safety: the function is idempotent.  Running it a second time re-checks each
        # step and skips anything already in place.
        #
        if ($State['InstallEdge']) { Write-MyVerbose 'Enable-AccessNamespaceMailConfig: skipped (Edge Transport)'; return }
        if (-not $State['Namespace']) { Write-MyVerbose 'Enable-AccessNamespaceMailConfig: no namespace set — skipping'; return }

        # MailDomain is the root domain used for email addresses (e.g. contoso.com).
        # It defaults to the parent of the access namespace (drop leftmost label).
        # e.g. Namespace=outlook.domain.de → MailDomain=domain.de
        $ns = if ($State['MailDomain']) {
            $State['MailDomain']
        } else {
            $part = ($State['Namespace'] -split '\.', 2)[1]
            if ($part -match '\.') { $part } else { $State['Namespace'] }
        }

        Write-MyOutput ('Configuring access namespace mail settings — mail domain: {0}' -f $ns)

        # ── 1. Accepted Domain ──────────────────────────────────────────────────
        try {
            $existing = Get-AcceptedDomain -ErrorAction Stop | Where-Object { $_.DomainName -eq $ns }
            if ($existing) {
                Write-MyVerbose ('Accepted domain already present: {0} ({1})' -f $ns, $existing.DomainType)
            }
            else {
                New-AcceptedDomain -Name $ns -DomainName $ns -DomainType Authoritative -ErrorAction Stop | Out-Null
                Register-ExecutedCommand -Category 'ExchangePolicy' -Command ("New-AcceptedDomain -Name '{0}' -DomainName '{0}' -DomainType Authoritative" -f $ns)
                Write-MyOutput ('Accepted domain added: {0} (Authoritative)' -f $ns)
            }
        }
        catch {
            Write-MyWarning ('Could not create accepted domain {0}: {1}' -f $ns, $_.Exception.Message)
            return
        }

        # ── 2. Email Address Policy — primary SMTP template ─────────────────────
        try {
            $policy = Get-EmailAddressPolicy -ErrorAction Stop | Sort-Object Priority | Select-Object -First 1
            if (-not $policy) {
                Write-MyWarning 'No Email Address Policy found — skipping policy update'
                return
            }

            $currentTemplates = @($policy.EnabledEmailAddressTemplates)
            $nsTemplate        = "SMTP:@$ns"   # uppercase SMTP = primary

            # Build non-internal templates (remove .local / nonroutable suffixes as primary/secondary)
            # but keep any routable domain that is NOT the access namespace as secondary smtp:
            $keepTemplates = $currentTemplates | Where-Object {
                $t = $_ -replace '^smtp:|^SMTP:',''
                # Drop .local, .internal, .lan addresses — these are not routable
                ($t -notmatch '\.(local|internal|lan|corp)$') -and ($t -ne "@$ns")
            } | ForEach-Object {
                # Downcase existing primary markers so we don't have two primaries
                $_ -replace '^SMTP:', 'smtp:'
            }

            # Check whether we dropped any internal templates (for logging)
            $droppedCount = $currentTemplates.Count - $keepTemplates.Count - 1  # -1 for namespace slot
            if ($droppedCount -gt 0) {
                Write-MyOutput ("Removing {0} internal-only address template(s) from policy '{1}'" -f $droppedCount, $policy.Name)
            }

            # Build new list: access namespace first (primary), remaining routable domains second
            $newTemplates = @($nsTemplate) + @($keepTemplates | Where-Object { $_ })

            # Only update if something actually changed
            $alreadyPrimary = ($currentTemplates | Select-Object -First 1) -ieq $nsTemplate
            $sameCount      = ($newTemplates.Count -eq $currentTemplates.Count) -and
                              (($newTemplates | Sort-Object) -join '|') -eq (($currentTemplates | Sort-Object) -join '|')

            if ($alreadyPrimary -and $sameCount) {
                Write-MyVerbose ("Email Address Policy '{0}' already configured correctly — no change needed" -f $policy.Name)
            }
            else {
                Set-EmailAddressPolicy -Identity $policy.Identity -EnabledEmailAddressTemplates $newTemplates -ErrorAction Stop
                Register-ExecutedCommand -Category 'ExchangePolicy' -Command ("Set-EmailAddressPolicy -Identity '{0}' -EnabledEmailAddressTemplates @('{1}')" -f $policy.Name, ($newTemplates -join "','"))
                Write-MyOutput ("Email Address Policy '{0}' updated — primary SMTP: @{1}" -f $policy.Name, $ns)

                # ── 3. Apply policy to all mailboxes ────────────────────────────
                Write-MyOutput ('Applying Email Address Policy to all mailboxes (this may take a moment)...')
                Update-EmailAddressPolicy -Identity $policy.Identity -ErrorAction Stop
                Register-ExecutedCommand -Category 'ExchangePolicy' -Command ("Update-EmailAddressPolicy -Identity '{0}'" -f $policy.Name)
                Write-MyOutput ('Email Address Policy applied.')
            }
        }
        catch {
            Write-MyWarning ('Email Address Policy update failed: {0}' -f $_.Exception.Message)
        }
    }

