    function Get-RBACReport {
        Write-MyOutput 'Generating RBAC role group membership report'
        $reportPath = Join-Path $State['ReportsPath'] ('{0}_EXpress_RBAC_{1}.txt' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd-HHmmss'))

        $roleGroups = @(
            'Organization Management',
            'Server Management',
            'Recipient Management',
            'Help Desk',
            'Hygiene Management',
            'Compliance Management',
            'Records Management',
            'Discovery Management',
            'Public Folder Management',
            'View-Only Organization Management'
        )

        $lines = [System.Collections.Generic.List[string]]::new()
        $lines.Add('Exchange RBAC Role Group Membership Report')
        $lines.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        $lines.Add("Server: $env:COMPUTERNAME")
        $lines.Add('-' * 60)

        foreach ($group in $roleGroups) {
            try {
                $members = @(Get-RoleGroupMember -Identity $group -ErrorAction Stop)
                $lines.Add('')
                $lines.Add("[$group]")
                if ($members.Count -gt 0) {
                    foreach ($member in $members) {
                        try {
                            $memberName = [string]$member.Name
                            $memberType = [string]$member.RecipientType
                            $lines.Add("  $memberName ($memberType)")
                        }
                        catch {
                            $lines.Add('  (could not display member)')
                        }
                    }
                }
                else {
                    $lines.Add('  (no members)')
                }
            }
            catch {
                $errMsg = if ($null -ne $_.Exception -and $_.Exception.Message) { $_.Exception.Message } else { $_.ToString() }
                $lines.Add('')
                $lines.Add("[$group] - could not retrieve: $errMsg")
            }
        }

        try {
            $lines | Set-Content -Path $reportPath -Encoding UTF8 -ErrorAction Stop
            Write-MyOutput "RBAC report saved to $reportPath"
        }
        catch {
            Write-MyWarning "Could not save RBAC report: $($_.Exception.Message)"
            $lines | ForEach-Object { Write-MyOutput $_ }
        }
    }

    function Get-OptimizationCatalog {
        # ─── Exchange Optimization Catalog ────────────────────────────────────────
        # To add a new optimization: append a hashtable to the array below.
        # Required fields:
        #   Key         – Single letter (A–Z) used as menu toggle key
        #   Name        – Unique identifier (used internally)
        #   Label       – Short display name shown in the menu (max 26 chars)
        #   Hint        – One-liner shown alongside the toggle (max 28 chars)
        #   Description – Full explanation shown in the description panel
        #   Default     – $true = selected by default, $false = opt-in
        #   Action      – ScriptBlock executed when the optimization is applied
        # Optional:
        #   Condition   – ScriptBlock returning $true if this entry is applicable.
        #                 If omitted the entry is always shown.
        # ──────────────────────────────────────────────────────────────────────────
        return @(
            @{
                Key         = 'A'
                Name        = 'ModernAuth'
                Label       = 'Modern Authentication'
                Hint        = 'Outlook 2016+, Teams, mobile'
                Description = 'Enables OAuth2 / Modern Authentication org-wide (Set-OrganizationConfig -OAuth2ClientProfileEnabled $true). Required for Outlook 2016+, Microsoft Teams, all mobile clients and any Hybrid / Azure AD configuration. Safe to enable on all Exchange 2016 / 2019 / SE installations. Without this, Outlook falls back to Basic Auth which Microsoft is deprecating.'
                Default     = $true
                Action      = {
                    Write-MyOutput 'Enabling Modern Authentication (OAuth2)'
                    Set-OrganizationConfig -OAuth2ClientProfileEnabled $true -ErrorAction Stop -WarningAction SilentlyContinue
                }
            }
            @{
                Key         = 'B'
                Name        = 'SessionTimeout'
                Label       = 'OWA Session Timeout (6h)'
                Hint        = 'Auto-logout after inactivity'
                Description = 'Sets activity-based OWA/ECP session timeout to 6 hours (Set-OrganizationConfig -ActivityBasedAuthenticationTimeoutEnabled $true -ActivityBasedAuthenticationTimeoutInterval 06:00:00). After 6 hours of inactivity the browser session is automatically logged out. Recommended for open-plan or shared workstation environments and for compliance requirements that mandate session expiry.'
                Default     = $true
                Action      = {
                    Write-MyOutput 'Configuring OWA/ECP session timeout (6 hours inactivity)'
                    Set-OrganizationConfig -ActivityBasedAuthenticationTimeoutEnabled $true -ActivityBasedAuthenticationTimeoutInterval '06:00:00' -ErrorAction Stop -WarningAction SilentlyContinue
                }
            }
            @{
                Key         = 'C'
                Name        = 'DisableTelemetry'
                Label       = 'Disable Telemetry (CEIP)'
                Hint        = 'Privacy / DSGVO: no Watson'
                Description = 'Disables the Microsoft Customer Experience Improvement Program (CEIP) and Watson crash telemetry (Set-OrganizationConfig -CustomerFeedbackEnabled $false). Prevents Exchange from sending diagnostic and usage data to Microsoft. Recommended for environments with strict data-privacy requirements (GDPR / DSGVO) or where external telemetry is blocked by policy.'
                Default     = $true
                Action      = {
                    Write-MyOutput 'Disabling CEIP / telemetry'
                    Set-OrganizationConfig -CustomerFeedbackEnabled $false -ErrorAction Stop -WarningAction SilentlyContinue
                }
            }
            @{
                Key         = 'D'
                Name        = 'MapiHttp'
                Label       = 'MAPI over HTTP (explicit)'
                Hint        = 'Replaces legacy RPC/HTTP'
                Description = 'Explicitly enables MAPI over HTTP (Set-OrganizationConfig -MapiHttpEnabled $true). MAPI/HTTP replaces the older Outlook Anywhere (RPC/HTTP), offering faster failover, better behaviour across NAT and load balancers, and improved Outlook startup performance. Enabled by default since Exchange 2016, but explicit activation avoids edge cases after upgrades or migrations.'
                Default     = $true
                Action      = {
                    Write-MyOutput 'Enabling MAPI over HTTP'
                    Set-OrganizationConfig -MapiHttpEnabled $true -ErrorAction Stop -WarningAction SilentlyContinue
                }
            }
            @{
                Key         = 'E'
                Name        = 'MaxMessageSize'
                Label       = 'Max Message Size (150MB)'
                Hint        = 'Org-wide send/receive limit'
                Description = 'Raises the organisation-wide maximum message size to 150MB for both send and receive, and limits recipients per message to 500 (Set-TransportConfig -MaxSendSize/-MaxReceiveSize/-MaxRecipientEnvelopeLimit). The Exchange default of 25MB is often too restrictive for modern file-sharing workflows. Frontend Receive Connectors are updated consistently. Adjust to match your storage capacity and bandwidth.'
                Default     = $true
                Action      = {
                    Write-MyOutput 'Setting org-wide max message size to 150MB'
                    Set-TransportConfig -MaxSendSize 150MB -MaxReceiveSize 150MB -MaxRecipientEnvelopeLimit 500 -ErrorAction Stop -WarningAction SilentlyContinue
                    Get-ReceiveConnector | Where-Object { $_.TransportRole -eq 'FrontendTransport' } | ForEach-Object {
                        Set-ReceiveConnector -Identity $_.Identity -MaxMessageSize 150MB -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                    }
                }
            }
            @{
                Key         = 'F'
                Name        = 'MessageExpiration'
                Label       = 'Message expiration 7 days'
                Hint        = 'Delay NDRs on outage (default: 2d)'
                Description = 'Extends message expiration timeout from the default 2 days to 7 days (Set-TransportService -MessageExpirationTimeout 7.00:00:00). During an outage or connectivity loss, messages remain in the queue for up to 7 days before an NDR is generated. Recommended for environments where short outages should not immediately result in delivery failure notifications. Skipped when CopyServerConfig is active (value is imported from source server).'
                Default     = $true
                Condition   = { -not $State['CopyServerConfig'] }
                Action      = {
                    $current = (Get-TransportService -Identity $env:COMPUTERNAME).MessageExpirationTimeout
                    if ($current -ne [TimeSpan]'7.00:00:00') {
                        Write-MyOutput 'Setting message expiration timeout to 7 days'
                        Set-TransportService -Identity $env:COMPUTERNAME -MessageExpirationTimeout 7.00:00:00 -ErrorAction Stop -WarningAction SilentlyContinue
                    }
                    else {
                        Write-MyVerbose 'Message expiration timeout already set to 7 days'
                    }
                }
            }
            @{
                Key         = 'G'
                Name        = 'ConnectorBanner'
                Label       = 'Harden SMTP Banner'
                Hint        = 'Remove Exchange version info'
                Description = 'Replaces the default SMTP greeting banner on all Frontend Receive Connectors with a generic "220 Mail Service" message (Set-ReceiveConnector -Banner). The default banner discloses the exact Exchange version, which helps attackers identify applicable CVEs. This is a low-effort hardening step recommended by security benchmarks (CIS, DISA STIG).'
                Default     = $true
                Action      = {
                    Write-MyOutput 'Hardening SMTP banner on Frontend Receive Connectors'
                    Get-ReceiveConnector | Where-Object { $_.TransportRole -eq 'FrontendTransport' } | ForEach-Object {
                        Set-ReceiveConnector -Identity $_.Identity -Banner '220 Mail Service' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                    }
                }
            }
            @{
                Key         = 'H'
                Name        = 'HtmlNDR'
                Label       = 'HTML Non-Delivery Reports'
                Hint        = 'Readable bounce messages'
                Description = 'Configures Exchange to generate HTML-formatted Non-Delivery Reports for both internal and external messages (Set-TransportConfig -InternalDsnSendHtml $true -ExternalDsnSendHtml $true). Plain-text NDRs are difficult for end users to interpret. HTML NDRs include formatted error descriptions and suggested next steps, reducing helpdesk escalations.'
                Default     = $true
                Action      = {
                    Write-MyOutput 'Enabling HTML-formatted Non-Delivery Reports'
                    Set-TransportConfig -InternalDsnSendHtml $true -ExternalDsnSendHtml $true -ErrorAction Stop -WarningAction SilentlyContinue
                }
            }
            @{
                Key         = 'I'
                Name        = 'ShadowRedundancy'
                Label       = 'Shadow Redundancy (DAG)'
                Hint        = 'Prefer remote shadow copy'
                Description = 'Configures Shadow Message Redundancy to prefer a remote DAG member as the shadow server (Set-TransportConfig -ShadowMessagePreferenceSetting PreferRemote). In a DAG, this ensures the redundant copy of each in-flight message is held on a different physical server than the primary, improving resilience against single-server failure during transport. Only effective in a DAG deployment.'
                Default     = $false
                Condition   = { $State['DAGName'] }
                Action      = {
                    Write-MyOutput 'Configuring Shadow Redundancy to prefer remote DAG member'
                    Set-TransportConfig -ShadowMessagePreferenceSetting PreferRemote -ErrorAction Stop -WarningAction SilentlyContinue
                }
            }
            @{
                Key         = 'J'
                Name        = 'SafetyNet'
                Label       = 'Safety Net Hold Time (2d)'
                Hint        = 'Explicit redelivery hold time'
                Description = 'Explicitly sets the Safety Net message hold time to 2 days (Set-TransportConfig -SafetyNetHoldTime 2.00:00:00). Safety Net retains a redundant copy of successfully delivered messages, enabling redelivery after a database failure or mailbox switchover. The 2-day default is appropriate for most environments; adjust to match your backup and recovery SLA.'
                Default     = $true
                Action      = {
                    Write-MyOutput 'Setting Safety Net hold time to 2 days'
                    Set-TransportConfig -SafetyNetHoldTime '2.00:00:00' -ErrorAction Stop -WarningAction SilentlyContinue
                }
            }
        )
    }

    function Invoke-SingleOptimization {
        param($Opt)
        try {
            & $Opt.Action
        }
        catch {
            Write-MyWarning ('Optimization [{0}] {1} failed: {2}' -f $Opt.Key, $Opt.Label, $_.Exception.Message)
        }
    }

    function Invoke-ExchangeOptimizations {
        $catalog    = Get-OptimizationCatalog
        $applicable = @($catalog | Where-Object { -not $_.ContainsKey('Condition') -or (& $_.Condition) })

        if ($applicable.Count -eq 0) {
            Write-MyVerbose 'No applicable Exchange org/transport optimizations for this configuration'
            return
        }

        # Selection state: Key -> bool
        $sel = @{}
        foreach ($opt in $applicable) { $sel[$opt.Key] = $opt.Default }

        # ── Autopilot / non-interactive: apply defaults without menu ──────────
        if ($State['Autopilot'] -or -not [Environment]::UserInteractive) {
            $defaults = @($applicable | Where-Object { $sel[$_.Key] })
            Write-MyOutput ('Applying Exchange optimizations — {0} of {1} selected (defaults)' -f $defaults.Count, $applicable.Count)
            foreach ($opt in $defaults) { Invoke-SingleOptimization $opt }
            return
        }

        # ── Interactive menu ──────────────────────────────────────────────────
        $byKey    = @{}
        foreach ($opt in $applicable) { $byKey[$opt.Key] = $opt }
        $keys     = @($applicable | ForEach-Object { $_.Key })
        $half     = [Math]::Ceiling($keys.Count / 2)
        $lastKey  = ''
        $statusMsg = ''

        $useRawKey = $false
        try { $null = $host.UI.RawUI.KeyAvailable; $useRawKey = $true } catch { }

        function Draw-OptimizationMenu {
            param([string]$Status = '', [string]$LastKey = '')
            Clear-Host
            Write-Host ('=' * 62) -ForegroundColor Cyan
            Write-Host ('  EXpress v{0} — Exchange Optimizations' -f $script:ScriptVersion) -ForegroundColor Cyan
            Write-Host ('=' * 62) -ForegroundColor Cyan
            Write-Host ''
            Write-Host '  Toggle optimizations to apply in Phase 5:' -ForegroundColor Yellow
            Write-Host ''

            # Two-column toggle list
            for ($r = 0; $r -lt $half; $r++) {
                $lk = $keys[$r]
                $rk = if (($r + $half) -lt $keys.Count) { $keys[$r + $half] } else { $null }
                $lo = $byKey[$lk]

                $lv   = if ($sel[$lk]) { 'X' } else { ' ' }
                $lStr = '  [{0}] [{1}] {2,-26}' -f $lk, $lv, $lo.Label

                $lColor = if ($lk -eq $LastKey) { [System.ConsoleColor]::Yellow } else { [System.ConsoleColor]::White }
                Write-Host $lStr -ForegroundColor $lColor -NoNewline

                if ($rk) {
                    $ro    = $byKey[$rk]
                    $rv    = if ($sel[$rk]) { 'X' } else { ' ' }
                    $rStr  = '   [{0}] [{1}] {2,-26}' -f $rk, $rv, $ro.Label
                    $rColor = if ($rk -eq $LastKey) { [System.ConsoleColor]::Yellow } else { [System.ConsoleColor]::White }
                    Write-Host $rStr -ForegroundColor $rColor
                } else {
                    Write-Host ''
                }
            }

            Write-Host ''

            # Description panel — shows full description of last-toggled option
            Write-Host ('  ' + [string]::new([char]0x2500, 58)) -ForegroundColor DarkGray
            if ($LastKey -and $byKey.ContainsKey($LastKey)) {
                $opt  = $byKey[$LastKey]
                $optState = if ($sel[$LastKey]) { 'ENABLED' } else { 'DISABLED' }  # NOTE: not $state — shadows outer $State hashtable
                Write-Host ('  [{0}] {1}  ({2})' -f $LastKey, $opt.Label, $optState) -ForegroundColor Yellow
                Write-Host ''
                # Word-wrap description at 58 chars
                $words  = ($opt.Description -replace '\s+', ' ').Trim() -split ' '
                $line   = '  '
                foreach ($w in $words) {
                    if (($line + $w).Length -gt 60) {
                        Write-Host $line
                        $line = '  ' + $w + ' '
                    }
                    else { $line += $w + ' ' }
                }
                if ($line.Trim()) { Write-Host $line }
            }
            else {
                Write-Host '  Press a letter key to see a detailed description.' -ForegroundColor DarkGray
            }
            Write-Host ('  ' + [string]::new([char]0x2500, 58)) -ForegroundColor DarkGray
            Write-Host ''

            if ($Status) { Write-Host "  $Status" -ForegroundColor Yellow; Write-Host '' }
        }

        while ($true) {
            Draw-OptimizationMenu -Status $statusMsg -LastKey $lastKey
            $statusMsg = ''

            if ($useRawKey) {
                Write-Host ('  Press {0} to toggle  |  ENTER = apply  |  S = skip all: ' -f ($keys -join '/')) -NoNewline -ForegroundColor Cyan
                try {
                    $keyInfo = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                    $vk  = $keyInfo.VirtualKeyCode
                    $raw = $keyInfo.Character.ToString().ToUpper()
                    Write-Host $raw
                    if ($vk -eq 13)    { break }        # Enter = apply
                    if ($vk -eq 27 -or $raw -eq 'S') { Write-MyOutput 'Exchange optimizations skipped'; return }
                }
                catch {
                    $useRawKey = $false
                    $raw = (Read-Host '').Trim().ToUpper()
                    if ($raw -eq '')   { break }
                    if ($raw -eq 'S')  { Write-MyOutput 'Exchange optimizations skipped'; return }
                }
            }
            else {
                $raw = (Read-Host ('  Toggle [{0}]  |  ENTER = apply  |  S = skip all' -f ($keys -join '/'))).Trim().ToUpper()
                if ($raw -eq '')  { break }
                if ($raw -eq 'S') { Write-MyOutput 'Exchange optimizations skipped'; return }
            }

            if ($raw.Length -eq 1 -and $byKey.ContainsKey($raw)) {
                $sel[$raw] = -not $sel[$raw]
                $lastKey   = $raw
            }
            elseif ($raw.Length -gt 0) {
                $statusMsg = "Unknown key '$raw' — use the listed letters, ENTER or S"
            }
        }

        # Apply selected optimizations
        $applied = 0
        foreach ($opt in $applicable | Where-Object { $sel[$_.Key] }) {
            Invoke-SingleOptimization $opt
            $applied++
        }
        Write-MyOutput ('{0} of {1} Exchange optimization(s) applied' -f $applied, $applicable.Count)
    }

