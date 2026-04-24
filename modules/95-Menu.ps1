    function Get-AdvancedFeatureCatalog {
        # Advanced Configuration catalog. Each entry:
        #   Name        — unique key persisted in $State['AdvancedFeatures'] and config file
        #   Label       — short display text in the Advanced menu (max ~30 chars)
        #   Description — one-line explanation shown in description panel
        #   Default     — $true/$false; matches current behaviour unless noted
        #   Category    — 'TLS', 'Hardening', 'Performance', 'ExchangePolicy', 'PostConfig', 'InstallFlow'
        #   Condition   — (optional) scriptblock; entry hidden when it returns $false
        return [ordered]@{
            # ─── Security / TLS ──────────────────────────────────────────────
            DisableSSL3         = @{ Category='TLS'; Label='Disable SSL 3.0';                Default=$true;  Description='Disable legacy SSL 3.0 (POODLE, CVE-2014-3566).' }
            DisableRC4          = @{ Category='TLS'; Label='Disable RC4 cipher';             Default=$true;  Description='Disable deprecated RC4 stream cipher.' }
            EnableECC           = @{ Category='TLS'; Label='Prefer ECC key exchange';        Default=$true;  Description='Enable ECC cipher suites and prefer over RSA.' }
            NoCBC               = @{ Category='TLS'; Label='Disable CBC ciphers';            Default=$false; Description='Disables CBC cipher suites. Not recommended — breaks compatibility with several clients.' }
            EnableAMSI          = @{ Category='TLS'; Label='Enable AMSI';                    Default=$true;  Description='Antimalware Scan Interface for Exchange transport and OWA.' }
            EnableTLS12         = @{ Category='TLS'; Label='Enforce TLS 1.2';                Default=$true;  Description='Enforce TLS 1.2; disables TLS 1.0/1.1 on SChannel and .NET StrongCrypto.' }
            EnableTLS13         = @{ Category='TLS'; Label='Enable TLS 1.3';                 Default=$true;  Description='Enable TLS 1.3 (Windows Server 2022+).'; Condition={ [System.Version]$script:FullOSVersion -ge [System.Version]$script:WS2022_PREFULL } }
            DoNotEnableEP       = @{ Category='TLS'; Label='Opt-out: Extended Protection';   Default=$false; Description='Skip Extended Protection configuration. Required for Hybrid + Modern Hybrid Topology where EP is incompatible.' }

            # ─── Security / Hardening ────────────────────────────────────────
            SMBv1Disable        = @{ Category='Hardening'; Label='Disable SMBv1';             Default=$true;  Description='Remove SMBv1 (WannaCry mitigation, MS17-010).' }
            NetBIOSDisable      = @{ Category='Hardening'; Label='Disable NetBIOS/TCP';       Default=$true;  Description='Disable NetBIOS over TCP/IP on all NICs (reduces attack surface).' }
            LLMNRDisable        = @{ Category='Hardening'; Label='Disable LLMNR';             Default=$true;  Description='Disable Link-Local Multicast Name Resolution (CIS L1 §18.5.4.2).' }
            MDNSDisable         = @{ Category='Hardening'; Label='Disable mDNS';              Default=$true;  Description='Disable Multicast DNS responder (WS2022+).' }
            WDigestDisable      = @{ Category='Hardening'; Label='Disable WDigest caching';   Default=$true;  Description='Prevent plaintext credentials in LSASS memory.' }
            LSAProtection       = @{ Category='Hardening'; Label='Enable LSA Protection';     Default=$true;  Description='RunAsPPL for LSASS to prevent credential dumping.' }
            LmCompat5           = @{ Category='Hardening'; Label='LmCompatibilityLevel=5';    Default=$true;  Description='Enforce NTLMv2, refuse LM/NTLMv1.' }
            SerializedDataSig   = @{ Category='Hardening'; Label='SerializedDataSigning';     Default=$true;  Description='Exchange SerializedDataSigning (MS-mandatory post CVE-2023-21529).' }
            ShutdownTrackerOff  = @{ Category='Hardening'; Label='Disable Shutdown Tracker';  Default=$true;  Description='Suppress the Shutdown Event Tracker reason dialog on server shutdowns.' }
            HSTS                = @{ Category='Hardening'; Label='HSTS on OWA/ECP';           Default=$true;  Description='HTTP Strict-Transport-Security header on OWA/ECP virtual directories.' }
            MAPIEncryption      = @{ Category='Hardening'; Label='Required MAPI encryption';  Default=$true;  Description='Set-RpcClientAccess -EncryptionRequired $true.' }
            HTTP2Disable        = @{ Category='Hardening'; Label='Disable HTTP/2';            Default=$true;  Description='Workaround for Exchange compatibility issues with HTTP/2.' }
            CredentialGuardOff  = @{ Category='Hardening'; Label='Disable Credential Guard';  Default=$true;  Description='Exchange is incompatible with Credential Guard; disable if enabled.' }
            UnnecessaryServices = @{ Category='Hardening'; Label='Disable unneeded services'; Default=$true;  Description='Disable Print Spooler, Xbox, Geolocation and other unneeded services on Exchange servers.' }
            WindowsSearchOff    = @{ Category='Hardening'; Label='Disable Windows Search';    Default=$true;  Description='Disable Windows Search service (not used by Exchange; uses CPU/IO).' }
            CRLTimeout          = @{ Category='Hardening'; Label='CRL Check Timeout';         Default=$true;  Description='Tune CRL retrieval timeout to avoid slow startup when OCSP/CRL endpoints are unreachable.' }
            RootCAAutoUpdate    = @{ Category='Hardening'; Label='Root CA Auto-Update';       Default=$true;  Description='Keep Automatic Root Certificates Update enabled (required for Modern Auth / O365 Hybrid).' }
            SMTPBannerHarden    = @{ Category='Hardening'; Label='Harden SMTP banner';        Default=$true;  Description='Replace Exchange version banner on Frontend Receive Connectors with "220 Mail Service".' }

            # ─── Performance / Tuning ────────────────────────────────────────
            MaxConcurrentAPI    = @{ Category='Performance'; Label='MaxConcurrentAPI';        Default=$true;  Description='MS KB 2688798 — raise MaxConcurrentApi to prevent NTLM auth bottlenecks.' }
            DiskAllocHint       = @{ Category='Performance'; Label='Disk allocation hint';    Default=$true;  Description='Emit warning when DB/log volumes are not formatted with 64K NTFS cluster size.' }
            CtsProcAffinity     = @{ Category='Performance'; Label='Content conv. affinity';  Default=$true;  Description='Limit Content Conversion processor affinity to stabilise CPU load.' }
            NodeRunnerMemLimit  = @{ Category='Performance'; Label='NodeRunner RAM cap';      Default=$true;  Description='Cap Exchange Search NodeRunner memory to prevent runaway allocations.' }
            MapiFeGC            = @{ Category='Performance'; Label='MAPI FrontEnd Server GC'; Default=$true;  Description='Enable Server GC mode for MAPI FrontEnd AppPool.' }
            NICPowerMgmtOff     = @{ Category='Performance'; Label='NIC Power Management';    Default=$true;  Description='Disable "Allow computer to turn off this device" on all NICs.' }
            RSSEnable           = @{ Category='Performance'; Label='Receive Side Scaling';    Default=$true;  Description='Enable RSS on all NICs for multi-core packet processing.' }
            TCPTuning           = @{ Category='Performance'; Label='TCP tuning';              Default=$true;  Description='Autotuning, Chimney offload and related TCP stack tweaks for Exchange workloads.' }
            TCPOffloadOff       = @{ Category='Performance'; Label='Disable TCP offload';     Default=$true;  Description='Disable TCP checksum/segmentation offload (avoids driver bugs on Exchange).' }
            IPv4OverIPv6Off     = @{ Category='Performance'; Label='Prefer IPv4 over IPv6';    Default=$true;  Description='Prefer IPv4 over IPv6 (DisabledComponents=0x20) — avoids Exchange DNS-lookup delays on IPv6-only hosts.' }

            # ─── Exchange Org Policy (current Optimization Catalog A–J) ──────
            ModernAuth          = @{ Category='ExchangePolicy'; Label='Modern Auth (OAuth2)';    Default=$true;  Description='Org-wide OAuth2 / Modern Authentication. Required for Outlook 2016+, Teams, mobile.' }
            OWASessionTimeout6h = @{ Category='ExchangePolicy'; Label='OWA Session Timeout 6h';  Default=$true;  Description='Activity-based OWA/ECP session timeout at 6h inactivity.' }
            DisableTelemetry    = @{ Category='ExchangePolicy'; Label='Disable CEIP telemetry';  Default=$true;  Description='Set-OrganizationConfig -CustomerFeedbackEnabled $false (GDPR/DSGVO).' }
            MapiHttp            = @{ Category='ExchangePolicy'; Label='MAPI over HTTP';          Default=$true;  Description='Explicit MapiHttpEnabled — replaces legacy RPC/HTTP.' }
            MaxMessageSize150MB = @{ Category='ExchangePolicy'; Label='Max message size 150MB';  Default=$true;  Description='Raise org-wide + Frontend receive connector max message size to 150 MB.' }
            MessageExpiration7d = @{ Category='ExchangePolicy'; Label='Expiration 7 days';       Default=$true;  Description='Extend transport message expiration to 7 days. Condition: not CopyServerConfig.'; Condition={ -not $script:State['CopyServerConfig'] } }
            HtmlNDR             = @{ Category='ExchangePolicy'; Label='HTML NDR formatting';     Default=$true;  Description='Set-TransportConfig -InternalDsnSendHtml / -ExternalDsnSendHtml.' }
            ShadowRedundancy    = @{ Category='ExchangePolicy'; Label='Shadow Redundancy';       Default=$false; Description='Prefer remote DAG member for shadow copies. DAG-only.'; Condition={ [bool]$script:State['DAGName'] } }
            SafetyNet2d         = @{ Category='ExchangePolicy'; Label='Safety Net 2d hold';      Default=$true;  Description='Safety Net hold time set to 2 days.' }

            # ─── Post-Config / Integration ───────────────────────────────────
            MECA                = @{ Category='PostConfig'; Label='MECA Auth Cert Renewal';  Default=$true;  Description='Register CSS-Exchange MonitorExchangeAuthCertificate scheduled task for automatic renewal.' }
            AntispamAgents      = @{ Category='PostConfig'; Label='Install Antispam Agents'; Default=$true;  Description='Install built-in antispam agents (Mailbox role only; no effect on Edge).' }
            SSLOffloading       = @{ Category='PostConfig'; Label='SSL Offloading tuning';   Default=$true;  Description='IIS/OWA SSL offload settings for load-balanced deployments.' }
            MRSProxy            = @{ Category='PostConfig'; Label='Enable MRS Proxy';        Default=$true;  Description='Enable MRS Proxy on EWS for cross-forest/cross-org mailbox moves.' }
            IANATimezone        = @{ Category='PostConfig'; Label='IANA timezone mapping';   Default=$true;  Description='Configure IANA ↔ Windows timezone mapping (iCal interop).' }
            AnonymousRelay      = @{ Category='PostConfig'; Label='Anonymous relay connector'; Default=$true; Description='Create anonymous internal/external relay connector if RelaySubnets is configured.'; Condition={ [bool]$script:State['RelaySubnets'] -or [bool]$script:State['ExternalRelaySubnets'] } }
            AccessNamespaceMail = @{ Category='PostConfig'; Label='Access Namespace mail config'; Default=$true; Description='Add Access Namespace as Authoritative Accepted Domain and set it as primary SMTP in the default Email Address Policy. Removes .local/nonroutable templates. Only available when EXpress created the Exchange org.'; Condition={ [bool]$script:State['Namespace'] -and [bool]$script:State['NewExchangeOrg'] } }
            SkipHealthCheck     = @{ Category='PostConfig'; Label='Opt-out: HealthChecker';  Default=$false; Description='Skip CSS-Exchange HealthChecker run at end of Phase 6.' }
            RBACReport          = @{ Category='PostConfig'; Label='RBAC Report';             Default=$true;  Description='Generate RBAC (role assignments / role groups) HTML report.' }
            RunEOMT             = @{ Category='PostConfig'; Label='Run EOMT';                Default=$false; Description='Run CSS-Exchange Emergency Mitigation Tool (legacy CUs; no-op on current CUs).' }

            # ─── Install-Flow / Debug ────────────────────────────────────────
            AutoApproveWindowsUpdates = @{ Category='InstallFlow'; Label='Auto-approve Windows Updates'; Default=$false; Description='Autopilot: approve all pending Security/Critical Windows Updates without prompting. Off by default — deliberate opt-in required.' }
            DiagnosticData      = @{ Category='InstallFlow'; Label='Send diagnostic data';   Default=$false; Description='/IAcceptExchangeServerLicenseTerms_DiagnosticDataON — share setup telemetry with Microsoft.' }
            Lock                = @{ Category='InstallFlow'; Label='Lock screen during run'; Default=$false; Description='Lock the console while the installation is in progress (Autopilot only).' }
            SkipRolesCheck      = @{ Category='InstallFlow'; Label='Skip AD roles check';    Default=$false; Description='Skip Schema/Enterprise/Domain Admin membership check (use with caution).' }
            NoCheckpoint        = @{ Category='InstallFlow'; Label='Skip System Restore';    Default=$false; Description='Skip pre-install System Restore checkpoints.' }
            NoNet481            = @{ Category='InstallFlow'; Label='Skip .NET 4.8.1';        Default=$false; Description='Skip .NET 4.8.1 install (debug only — may break Exchange setup).' }
            WaitForADSync       = @{ Category='InstallFlow'; Label='Wait for AD replication'; Default=$false; Description='After PrepareAD, wait up to 6 min for error-free AD replication before continuing.' }
        }
    }

    function Show-AdvancedMenu {
        # Interactive Advanced Configuration menu. 2 categories per page (~3 pages total).
        # Navigation uses Enter / Backspace / 0 / Esc — never letter keys — so all
        # A-Z letters are available for toggling items without conflicts.
        # Returns a hashtable @{Name=bool} of all toggle states, or $null on cancel.
        # $InitialValues: pre-seed toggle state (e.g. from a previous C-press in the main menu).
        param([hashtable]$InitialValues = $null)

        $catalog = Get-AdvancedFeatureCatalog

        $categoryDefs = @(
            @{ Key='TLS';            Title='Security / TLS' }
            @{ Key='Hardening';      Title='Security / Hardening' }
            @{ Key='Performance';    Title='Performance / Tuning' }
            @{ Key='ExchangePolicy'; Title='Exchange Org Policy' }
            @{ Key='PostConfig';     Title='Post-Config / Integration' }
            @{ Key='InstallFlow';    Title='Install-Flow / Debug' }
        )

        # Initialize selection state; filter entries whose Condition is $false.
        $sel     = @{}
        $visible = @{}
        foreach ($cat in $categoryDefs) { $visible[$cat.Key] = @() }
        $existing = if ($InitialValues -is [hashtable])                { $InitialValues }
                    elseif ($State['AdvancedFeatures'] -is [hashtable]) { $State['AdvancedFeatures'] }
                    else                                                 { @{} }
        foreach ($name in $catalog.Keys) {
            $entry = $catalog[$name]
            if ($entry.ContainsKey('Condition')) {
                try { if (-not (& $entry.Condition)) { continue } } catch { continue }
            }
            $sel[$name] = if ($existing.ContainsKey($name)) { [bool]$existing[$name] } else { [bool]$entry.Default }
            $visible[$entry.Category] += ,$name
        }

        # Build pages: 2 non-empty categories per page.
        $activeCats = @($categoryDefs | Where-Object { $visible[$_.Key].Count -gt 0 })
        if ($activeCats.Count -eq 0) { Write-MyVerbose 'No advanced features applicable'; return $sel }
        $pages = @()
        for ($i = 0; $i -lt $activeCats.Count; $i += 2) {
            $pg = @($activeCats[$i])
            if ($i + 1 -lt $activeCats.Count) { $pg += $activeCats[$i + 1] }
            $pages += ,@{ Cats = $pg }
        }

        $useRawKey = $false
        try { $null = $host.UI.RawUI.KeyAvailable; $useRawKey = $true } catch { }

        $pageIdx   = 0
        $lastName  = ''
        $statusMsg = ''

        while ($true) {
            # Flatten all visible names on this page in category order (for letter assignment).
            $pageNames = @()
            foreach ($cat in $pages[$pageIdx].Cats) { $pageNames += $visible[$cat.Key] }
            $count = $pageNames.Count

            Clear-Host
            Write-Host ('=' * 70) -ForegroundColor Cyan
            Write-Host ('  EXpress v{0} — Advanced Configuration  (page {1}/{2})' -f $script:ScriptVersion, ($pageIdx + 1), $pages.Count) -ForegroundColor Cyan
            Write-Host ('=' * 70) -ForegroundColor Cyan

            # Render each category section with its own 2-column block.
            $offset = 0
            foreach ($cat in $pages[$pageIdx].Cats) {
                $catNames = @($visible[$cat.Key])
                if ($catNames.Count -eq 0) { continue }
                $sep = [string]::new([char]0x2500, [Math]::Max(0, 52 - $cat.Title.Length))
                Write-Host ''
                Write-Host ('  -- {0} {1}' -f $cat.Title, $sep) -ForegroundColor Yellow
                Write-Host ''
                $half = [int][Math]::Ceiling($catNames.Count / 2)
                for ($r = 0; $r -lt $half; $r++) {
                    $li      = $r
                    $ri      = $r + $half
                    $lName   = $catNames[$li]
                    $lLetter = [char]([int][char]'A' + $offset + $li)
                    $lEntry  = $catalog[$lName]
                    $lMark   = if ($sel[$lName]) { 'X' } else { ' ' }
                    $lColor  = if ($lName -eq $lastName) { [System.ConsoleColor]::Yellow } else { [System.ConsoleColor]::White }
                    Write-Host ('  [{0}] [{1}] {2,-28}' -f $lLetter, $lMark, $lEntry.Label) -ForegroundColor $lColor -NoNewline
                    if ($ri -lt $catNames.Count) {
                        $rName   = $catNames[$ri]
                        $rLetter = [char]([int][char]'A' + $offset + $ri)
                        $rEntry  = $catalog[$rName]
                        $rMark   = if ($sel[$rName]) { 'X' } else { ' ' }
                        $rColor  = if ($rName -eq $lastName) { [System.ConsoleColor]::Yellow } else { [System.ConsoleColor]::White }
                        Write-Host ('   [{0}] [{1}] {2,-28}' -f $rLetter, $rMark, $rEntry.Label) -ForegroundColor $rColor
                    } else {
                        Write-Host ''
                    }
                }
                $offset += $catNames.Count
            }

            # Description panel
            Write-Host ''
            Write-Host ('  ' + [string]::new([char]0x2500, 66)) -ForegroundColor DarkGray
            if ($lastName -and $catalog.Contains($lastName)) {
                $opt       = $catalog[$lastName]
                $dispState = if ($sel[$lastName]) { 'ENABLED' } else { 'DISABLED' }
                Write-Host ('  {0}  ({1})' -f $opt.Label, $dispState) -ForegroundColor Yellow
                Write-Host ''
                $words = ($opt.Description -replace '\s+', ' ').Trim() -split ' '
                $line  = '  '
                foreach ($w in $words) {
                    if (($line + $w).Length -gt 68) { Write-Host $line; $line = '  ' + $w + ' ' }
                    else { $line += $w + ' ' }
                }
                if ($line.Trim()) { Write-Host $line }
            } else {
                Write-Host '  Press a letter to toggle it and see its description.' -ForegroundColor DarkGray
            }
            Write-Host ('  ' + [string]::new([char]0x2500, 66)) -ForegroundColor DarkGray
            Write-Host ''

            if ($statusMsg) { Write-Host ('  ' + $statusMsg) -ForegroundColor Yellow; Write-Host ''; $statusMsg = '' }

            $lastLetter = [char]([byte][char]'A' + $count - 1)
            $navFwd  = if ($pageIdx -lt $pages.Count - 1) { 'Enter=Next' } else { 'Enter=Apply' }
            $navBack = if ($pageIdx -gt 0) { '  Back=Prev' } else { '' }
            Write-Host ('  [A-{0}]=Toggle  |  {1}{2}  |  0=Skip-all  |  Esc=Cancel: ' -f $lastLetter, $navFwd, $navBack) -NoNewline -ForegroundColor Cyan

            # Read one key. Nav is encoded as a sentinel string so no letter key is
            # ever reserved for navigation (fixes A/N/P/S conflicts in the old design).
            $action = 'UNKNOWN'
            if ($useRawKey) {
                try {
                    $keyInfo = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                    $vk = $keyInfo.VirtualKeyCode
                    $ch = $keyInfo.Character.ToString().ToUpper()
                    if     ($vk -eq 27)           { Write-Host ''; $action = 'CANCEL' }
                    elseif ($vk -eq 13)           { Write-Host ''; $action = 'NEXT'   }
                    elseif ($vk -in @(8, 37))     { Write-Host ''; $action = 'PREV'   }   # Backspace or Left arrow
                    elseif ($ch -eq '0')          { Write-Host '0'; $action = 'SKIP'  }
                    elseif ($ch -match '^[A-Z]$') { Write-Host $ch; $action = $ch     }
                    else                          { Write-Host '' }
                } catch {
                    $useRawKey = $false
                }
            }
            if (-not $useRawKey -and $action -eq 'UNKNOWN') {
                $typed = (Read-Host '').Trim().ToUpper()
                if     ($typed -eq '')             { $action = 'NEXT'   }
                elseif ($typed -in @('-','<','B')) { $action = 'PREV'   }
                elseif ($typed -eq '0')            { $action = 'SKIP'   }
                elseif ($typed -eq 'Q')            { $action = 'CANCEL' }
                elseif ($typed -match '^[A-Z]$')   { $action = $typed   }
            }

            switch ($action) {
                'NEXT' {
                    if ($pageIdx -lt $pages.Count - 1) { $pageIdx++; $lastName = '' }
                    else { return $sel }
                }
                'PREV' {
                    if ($pageIdx -gt 0) { $pageIdx--; $lastName = '' }
                    else { $statusMsg = 'Already on first page' }
                }
                'SKIP' {
                    foreach ($n in $sel.Keys.Clone()) { $sel[$n] = [bool]$catalog[$n].Default }
                    Write-MyOutput 'Advanced configuration skipped — using defaults'
                    return $sel
                }
                'CANCEL' {
                    Write-MyOutput 'Advanced configuration cancelled — continuing with defaults'
                    return $null
                }
                default {
                    if ($action -match '^[A-Z]$') {
                        $idx = [byte][char]$action - [byte][char]'A'
                        if ($idx -ge 0 -and $idx -lt $count) {
                            $targetName = $pageNames[$idx]
                            $sel[$targetName] = -not $sel[$targetName]
                            $lastName = $targetName
                        } else {
                            $statusMsg = "No item on key '$action' — valid range A-$lastLetter"
                        }
                    }
                }
            }
        }
    }

    function Invoke-AdvancedConfigurationPrompt {
        # Offers the Advanced Configuration menu with a 60-second auto-skip (default = skip).
        # Autopilot / non-interactive: returns immediately without prompting.
        # Returns $true if the menu was shown and settings saved, $false if skipped.
        if ($State['Autopilot'] -or -not [Environment]::UserInteractive) { return $false }
        if ($State.ContainsKey('SuppressAdvancedPrompt') -and $State['SuppressAdvancedPrompt']) { return $false }

        $timeoutSec = 60
        Write-Host ''
        Write-Host ('  Configure advanced options? [y/N] (auto-skip in {0}s) ' -f $timeoutSec) -NoNewline -ForegroundColor Cyan

        $deadline = (Get-Date).AddSeconds($timeoutSec)
        $answer   = ''
        while ((Get-Date) -lt $deadline) {
            if ($host.UI.RawUI.KeyAvailable) {
                $k = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                if ($k.VirtualKeyCode -eq 13 -or $k.VirtualKeyCode -eq 27) { break }
                $answer = $k.Character.ToString().ToUpper()
                Write-Host $answer
                break
            }
            Start-Sleep -Milliseconds 200
            $remaining = [int]([Math]::Ceiling(($deadline - (Get-Date)).TotalSeconds))
            Write-Progress -Id 2 -Activity 'Advanced configuration prompt' -Status ('Auto-skip in {0}s — press Y to configure, N/Enter to skip' -f $remaining) -SecondsRemaining $remaining
        }
        Write-Progress -Id 2 -Activity 'Advanced configuration prompt' -Completed

        if ($answer -ne 'Y') {
            Write-MyOutput 'Advanced configuration skipped — continuing with defaults'
            return $false
        }
        Write-Host ''

        $result = Show-AdvancedMenu
        if ($null -eq $result) {
            Write-MyOutput 'Advanced configuration cancelled — continuing with defaults'
            return $false
        }

        $State['AdvancedFeatures'] = $result
        Save-State $State
        $changed = @($result.Keys | Where-Object { $result[$_] -ne [bool](Get-AdvancedFeatureCatalog)[$_].Default }).Count
        Write-MyOutput ('Advanced configuration applied — {0} setting(s) differ from defaults' -f $changed)
        return $true
    }

    function Test-Feature {
        # Returns $true when an Advanced feature is enabled.
        # Precedence: $State['AdvancedFeatures'][Name] > catalog default.
        # Condition scriptblock (if any) is evaluated first — returns $false when not met,
        # regardless of the stored or default value. Prevents config-file bypass of
        # runtime-gated features (e.g. ShadowRedundancy without DAG, EnableTLS13 on WS2019).
        # Unknown names return $false (fail closed) and log a verbose warning.
        param([Parameter(Mandatory)][string]$Name)

        $catalog = Get-AdvancedFeatureCatalog
        if (-not $catalog.Contains($Name)) {
            Write-MyVerbose ("Test-Feature: unknown feature name '{0}' — returning `$false" -f $Name)
            return $false
        }

        $entry = $catalog[$Name]
        if ($entry.ContainsKey('Condition')) {
            try   { if (-not (& $entry.Condition)) { return $false } }
            catch { return $false }
        }

        $features = $State['AdvancedFeatures']
        if ($features -is [hashtable] -and $features.ContainsKey($Name)) {
            return [bool]$features[$Name]
        }
        return [bool]$entry.Default
    }

    function Show-InstallationMenu {
        # Interactive console menu. Returns a hashtable of all chosen settings, or $null if user cancelled.
        # Uses Read-Host for all input so it works reliably over RDP, Hyper-V console and Windows Terminal.

        $modes = @{
            1 = 'Exchange Server (Mailbox)'
            2 = 'Exchange Server (Edge Transport)  [not tested]'
            3 = 'Recipient Management Tools         [not tested]'
            4 = 'Exchange Management Tools only     [not tested]'
            5 = 'Recovery Mode                      [not tested]'
            6 = 'Standalone Optimize                [not tested]'
            7 = 'Generate Installation Document     [not tested]'
        }

        # Toggle definitions: Key=letter, Name=parameter name, Default=initial state
        # Main menu exposes installation-flow toggles only; ~55 hardening/tuning options
        # live in the Advanced Configuration menu (see Get-AdvancedFeatureCatalog).

        # Name = parameter/cfg key; Label = display text shown in menu
        $toggleDefs = [ordered]@{
            'A' = @{ Name='Autopilot';             Label='Autopilot (auto-reboot)';        Default=$true  }
            'B' = @{ Name='IncludeFixes';          Label='Install Exchange SU';            Default=$true  }
            'N' = @{ Name='PreflightOnly';         Label='Preflight only (no install)';    Default=$false }
            'R' = @{ Name='InstallWindowsUpdates'; Label='Install Windows Updates';        Default=$true  }
            'U' = @{ Name='GenerateDoc';           Label='Generate Installation Document'; Default=$false }
            'V' = @{ Name='German';                Label='Language:  DE (default EN)';     Default=$false }
        }

        # Toggles disabled per mode (letters that cannot be toggled in that mode)
        $disabledToggles = @{
            1 = @()
            2 = @('U','V')                                    # Edge: no installation doc
            3 = @('B','N','R','U','V')                        # Recipient Mgmt: only Autopilot
            4 = @('B','U','V')                                # Mgmt Tools: no setup, no doc
            5 = @()
            6 = @('B','N','R')                                # Standalone Optimize: no setup, no WU, no preflight
            7 = @('A','B','N','R','U')                        # Document-only: only language matters
        }

        # Initialize toggle states from defaults
        $toggleState = @{}
        foreach ($k in $toggleDefs.Keys) { $toggleState[$k] = $toggleDefs[$k].Default }

        $selectedMode = 0

        # Returns extra letters that should be disabled based on current toggle state
        function Get-DynamicDisabled {
            param([hashtable]$TS)
            $extra = @()
            if ($TS['N'])      { $extra += @('B','R') }   # PreflightOnly: SU/WU irrelevant
            if (-not $TS['U']) { $extra += 'V' }          # V (language) only meaningful when doc is generated
            return $extra
        }

        function Write-MenuLine {
            param([string]$Line, [System.ConsoleColor]$Color = [System.ConsoleColor]::White)
            Write-Host $Line -ForegroundColor $Color
        }

        function Draw-Menu {
            param([int]$Mode, [hashtable]$ToggState, [string]$StatusMsg = '', [array]$ExtraDisabled = @(), [int]$AdvCount = 0)
            Clear-Host
            Write-MenuLine ('=' * 60) Cyan
            Write-MenuLine ('  EXpress v{0}  —  Copilot' -f $ScriptVersion) Cyan
            Write-MenuLine ('=' * 60) Cyan
            Write-Host ''
            Write-MenuLine '  Installation Mode:' Yellow
            for ($i = 1; $i -le 7; $i++) {
                $marker = if ($Mode -eq $i) { '>' } else { ' ' }
                $color  = if ($Mode -eq $i) { [System.ConsoleColor]::Green } else { [System.ConsoleColor]::Gray }
                Write-Host ('    [{0}] {1}  {2}' -f $i, $marker, $modes[$i]) -ForegroundColor $color
            }
            Write-Host ''
            Write-MenuLine '  Switches (press letter to toggle, C=Advanced, then ENTER to start):' Yellow

            $disabled = @(if ($Mode -gt 0) { $disabledToggles[$Mode] } else { @() }) + $ExtraDisabled
            $letters  = @($toggleDefs.Keys)
            # Render two columns
            for ($r = 0; $r -lt [Math]::Ceiling($letters.Count / 2); $r++) {
                $left  = $letters[$r]
                $right = $letters[$r + [Math]::Ceiling($letters.Count / 2)]
                $leftDis  = $disabled -contains $left
                $rightDis = $right -and ($disabled -contains $right)
                $leftVal  = if ($ToggState[$left])  { 'X' } else { ' ' }
                $rightVal = if ($right -and $ToggState[$right]) { 'X' } else { ' ' }
                $leftStr  = '  [{0}] {1,-28} [{2}]' -f $left,  $toggleDefs[$left].Label,  $leftVal
                $rightStr = if ($right) { '   [{0}] {1,-28} [{2}]' -f $right, $toggleDefs[$right].Label, $rightVal } else { '' }
                $lColor = if ($leftDis)  { [System.ConsoleColor]::DarkGray } else { [System.ConsoleColor]::White }
                $rColor = if ($rightDis) { [System.ConsoleColor]::DarkGray } else { [System.ConsoleColor]::White }
                Write-Host $leftStr  -ForegroundColor $lColor -NoNewline
                Write-Host $rightStr -ForegroundColor $rColor
            }
            # Advanced Configuration shortcut
            $advStatus = if ($AdvCount -gt 0) { "($AdvCount customized)" } else { '(defaults)' }
            Write-Host ('  [C] Advanced Configuration...          {0}' -f $advStatus) -ForegroundColor Cyan

            Write-Host ''
            if ($StatusMsg) { Write-Host "  $StatusMsg" -ForegroundColor Yellow }
        }

        Write-MyVerbose 'Menu: Show-InstallationMenu started'

        # Advanced Configuration state — populated when user presses C.
        # Starts empty (@{}) so Test-Feature falls back to catalog defaults.
        $advancedFeatures = @{}

        # --- Step 1: Mode selection ---
        while ($selectedMode -lt 1 -or $selectedMode -gt 7) {
            Draw-Menu -Mode $selectedMode -ToggState $toggleState -AdvCount $advancedFeatures.Count
            $raw = Read-Host '  Mode [1-7]'
            if ($raw -match '^[1-7]$') {
                $selectedMode = [int]$raw
                Write-MyVerbose ('Menu: Mode {0} selected ({1})' -f $selectedMode, $modes[$selectedMode])
                # Apply mode-specific toggle defaults
                switch ($selectedMode) {
                    2 { $toggleState['G'] = $false; $toggleState['I'] = $false }
                    3 { foreach ($k in $disabledToggles[3]) { $toggleState[$k] = $false } }
                    6 { foreach ($k in $disabledToggles[6]) { $toggleState[$k] = $false } }
                    7 { foreach ($k in $disabledToggles[7]) { $toggleState[$k] = $false } }
                }
            }
        }

        # --- Step 2: Toggle switches ---
        # Try RawUI.ReadKey (no Enter needed); fall back to Read-Host if console is not interactive
        # (e.g. stdin redirected, PS2Exe without console, or restricted host).
        $useRawKey = $false
        try {
            $null = $host.UI.RawUI.KeyAvailable  # throws if RawUI is not available
            $useRawKey = $true
        } catch { }

        $statusMsg = ''
        while ($true) {
            $dynDisabled = Get-DynamicDisabled $toggleState
            Draw-Menu -Mode $selectedMode -ToggState $toggleState -StatusMsg $statusMsg -ExtraDisabled $dynDisabled -AdvCount $advancedFeatures.Count
            $statusMsg = ''

            if ($useRawKey) {
                Write-Host '  Press letter to toggle, C=Advanced, ENTER to start: ' -NoNewline -ForegroundColor Cyan
                try {
                    $keyInfo = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                    $vk  = $keyInfo.VirtualKeyCode
                    $raw = $keyInfo.Character.ToString().ToUpper()
                    Write-Host $raw  # echo the pressed key
                    if ($vk -eq 13) { break }                          # Enter
                    if ($vk -eq 27) { return $null }                   # Escape = cancel
                } catch {
                    # RawUI failed mid-session — fall back
                    $useRawKey = $false
                    $raw = (Read-Host '').Trim().ToUpper()
                    if ($raw -eq '') { break }
                }
            }
            else {
                $raw = (Read-Host '  Toggle letter / C=Advanced / ENTER to start').Trim().ToUpper()
                if ($raw -eq '') { break }
            }

            if ($raw -eq 'C') {
                # Open Advanced Configuration menu; preserve state across multiple C-presses.
                $advResult = Show-AdvancedMenu -InitialValues $advancedFeatures
                if ($null -ne $advResult) {
                    $advancedFeatures = $advResult
                    $changed = ($advancedFeatures.Keys | Where-Object { $advancedFeatures[$_] -ne [bool](Get-AdvancedFeatureCatalog)[$_].Default }).Count
                    Write-MyVerbose ('Menu: Advanced Configuration applied — {0} setting(s) differ from defaults' -f $changed)
                }
            }
            elseif ($raw.Length -eq 1 -and $toggleDefs.Contains($raw)) {
                $dynNow = Get-DynamicDisabled $toggleState
                if (($disabledToggles[$selectedMode] -contains $raw) -or ($dynNow -contains $raw)) {
                    $statusMsg = "[$raw] is not available in this configuration"
                }
                else {
                    $toggleState[$raw] = -not $toggleState[$raw]
                    $toggState = if ($toggleState[$raw]) { 'ON' } else { 'OFF' }
                    Write-MyVerbose ('Menu: Toggle [{0}] {1} -> {2}' -f $raw, $toggleDefs[$raw].Label, $toggState)
                    # Reset any toggles that became disabled by this change
                    $dynAfter = Get-DynamicDisabled $toggleState
                    foreach ($x in $dynAfter) {
                        if ($toggleState[$x]) {
                            $toggleState[$x] = $false
                            Write-MyVerbose ('Menu: Toggle [{0}] auto-cleared (now disabled)' -f $x)
                        }
                    }
                }
            }
            elseif ($raw.Length -gt 0) {
                $statusMsg = "Unknown key '$raw' — press a listed letter, C=Advanced, or ENTER to start"
            }
        }

        # --- Step 3: String inputs (context-dependent) ---
        Clear-Host
        Write-MenuLine ('=' * 60) Cyan
        Write-MenuLine ("  EXpress v{0} - Mode: {1}" -f $ScriptVersion, $modes[$selectedMode]) Cyan
        Write-MenuLine ('=' * 60) Cyan
        Write-Host ''
        Write-MenuLine '  Enter values (leave blank for default, shown in [brackets]):' Yellow
        Write-Host ''

        function Read-MenuInput {
            param(
                [string]$Prompt,
                [string]$Default = '',
                [bool]$Required = $false,
                [scriptblock]$Validate = $null,
                [string]$ValidateMessage = 'Invalid input — please try again'
            )
            while ($true) {
                if ($Default) {
                    Write-Host -NoNewline ("  {0} " -f $Prompt)
                    Write-Host -NoNewline ("[{0}]: " -f $Default) -ForegroundColor Green
                    $val = Read-Host
                } else {
                    $val = Read-Host ("  {0}" -f $Prompt)
                }
                if ($val -eq '') { $val = $Default }
                if ($Required -and -not $val) {
                    Write-Host '  (required — cannot be empty)' -ForegroundColor Yellow
                }
                elseif ($val -and $Validate -and -not (& $Validate $val)) {
                    Write-Host "  $ValidateMessage" -ForegroundColor Yellow
                }
                else { return $val }
            }
        }

        $validateFQDN = { param($v) $v -match '^[a-zA-Z0-9]([a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?(\.[a-zA-Z0-9]([a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?)+$' }
        $validateCIDRList = {
            param($v)
            ($v -split '\s*,\s*') | Where-Object { $_ } | ForEach-Object {
                $_ -match '^\d{1,3}(\.\d{1,3}){3}(/([0-9]|[12]\d|3[0-2]))?$'
            } | Where-Object { -not $_ } | Measure-Object | Select-Object -ExpandProperty Count | ForEach-Object { $_ -eq 0 }
        }

        $cfg = @{}
        $cfg['Mode']       = $selectedMode
        $cfg['InstallPath'] = if ($ScriptFullName) { Split-Path $ScriptFullName -Parent } else { $PWD.Path }
        if ($selectedMode -notin @(6, 7)) {
            $defaultIso = Join-Path $cfg['InstallPath'] 'sources\ExchangeServerSE-x64.iso'
            $srcTry = 0
            while ($true) {
                $srcPath = Read-MenuInput -Prompt 'Exchange source (folder or .iso)' -Default $defaultIso -Required $true
                if (Test-Path $srcPath) { $cfg['SourcePath'] = $srcPath; break }
                $srcTry++
                Write-Host ("  Path not found: {0}" -f $srcPath) -ForegroundColor Yellow
                if ($srcTry -ge 3) {
                    Write-Host '  3 failed attempts — returning to main menu.' -ForegroundColor Red
                    return $null
                }
                Write-Host ("  Attempt {0}/3 — verify the path and try again." -f $srcTry) -ForegroundColor Yellow
            }
        }

        if ($selectedMode -eq 1) {
            # Detect existing Exchange organisation from AD (requires domain connectivity)
            $detectedOrg = ''
            try {
                $configNC  = ([ADSI]'LDAP://RootDSE').configurationNamingContext
                $searcher  = New-Object System.DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC")
                $searcher.Filter = '(objectClass=msExchOrganizationContainer)'
                $searcher.PropertiesToLoad.Add('name') | Out-Null
                $result = $searcher.FindOne()
                if ($result) { $detectedOrg = $result.Properties['name'][0] }
            } catch { }

            if ($detectedOrg) {
                Write-Host ("  Existing Exchange organisation detected: {0}" -f $detectedOrg) -ForegroundColor Green
                $cfg['Organization'] = Read-MenuInput -Prompt 'Organization name      (ENTER = keep existing)' -Default $detectedOrg
            } else {
                Write-Host '  No existing Exchange organisation found in AD.' -ForegroundColor Yellow
                # Require an org name — cannot install into a new org without a name
                $orgInput = ''
                while (-not $orgInput) {
                    $orgInput = (Read-Host '  Organization name      (required for new org)').Trim()
                    if (-not $orgInput) {
                        Write-Host '  Organisation name is required when no existing organisation is found. Enter Q to quit.' -ForegroundColor Yellow
                        if ($orgInput -imatch '^[Qq]$') { return $null }
                    }
                }
                $cfg['Organization'] = $orgInput
            }

            # Detect current Autodiscover SCP URL from AD
            $currentSCP = ''
            try {
                $configNC2 = ([ADSI]'LDAP://RootDSE').configurationNamingContext
                $scpSearch = New-Object System.DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC2")
                $scpSearch.Filter = "(&(cn=$($env:COMPUTERNAME))(objectClass=serviceConnectionPoint)(serviceClassName=ms-Exchange-AutoDiscover-Service))"
                $scpSearch.PropertiesToLoad.Add('serviceBindingInformation') | Out-Null
                $scpResult = $scpSearch.FindOne()
                if ($scpResult) { $currentSCP = $scpResult.Properties['serviceBindingInformation'][0] }
            } catch { }

            $cfg['MDBName']          = Read-MenuInput -Prompt 'Mailbox DB name        (blank = default name)'
            $cfg['MDBDBPath']        = Read-MenuInput -Prompt 'Mailbox DB path        (blank = Exchange default)'
            $cfg['MDBLogPath']       = Read-MenuInput -Prompt 'Mailbox log path       (blank = Exchange default)'
            if ($currentSCP) {
                Write-Host ("  Current Autodiscover SCP: {0}" -f $currentSCP) -ForegroundColor DarkGray
                $cfg['SCP']          = Read-MenuInput -Prompt 'Autodiscover SCP URL   (ENTER = keep current, - = remove)' -Default $currentSCP
            } else {
                $cfg['SCP']          = Read-MenuInput -Prompt 'Autodiscover SCP URL   (blank = let Setup set, - = remove)'
            }
            $cfg['TargetPath']       = Read-MenuInput -Prompt 'Exchange install path  (blank = C:\Program Files\Microsoft\Exchange Server\V15)'
            $knownDAGs = Get-ExchangeDAGNames
            if ($knownDAGs.Count -gt 0) {
                Write-Host ("  DAGs found in AD: {0}" -f ($knownDAGs -join ', ')) -ForegroundColor DarkGray
                $cfg['DAGName'] = Read-MenuInput -Prompt ('DAG name               ({0}, blank = no DAG join)' -f ($knownDAGs -join ' / ')) -Default ($knownDAGs[0])
            } else {
                $cfg['DAGName'] = Read-MenuInput -Prompt 'DAG name               (blank = no DAG join)'
            }
            $cfg['CopyServerConfig'] = Read-MenuInput -Prompt 'Copy config from server (FQDN, blank = none) [not tested]' -Validate $validateFQDN -ValidateMessage 'Not a valid FQDN (e.g. ex01.contoso.com)'
            $cfg['CertificatePath']  = Read-MenuInput -Prompt 'PFX certificate path   (blank = none)        [not tested]'
            $cfg['Namespace']        = Read-MenuInput -Prompt 'Access namespace       (e.g. mail.contoso.com, blank = skip URL config)' -Validate $validateFQDN -ValidateMessage 'Not a valid FQDN (e.g. mail.contoso.com)'
            if ($cfg['Namespace']) {
                # Default mail domain = parent of access namespace (drop leftmost label)
                $defaultMailDomain = ($cfg['Namespace'] -split '\.', 2)[1]
                if ($defaultMailDomain -notmatch '\.') { $defaultMailDomain = $cfg['Namespace'] }
                $cfg['MailDomain']     = Read-MenuInput -Prompt 'Mail domain             (e.g. contoso.com — for Accepted Domain + email addresses)' -Default $defaultMailDomain -Validate $validateFQDN -ValidateMessage 'Not a valid domain (e.g. contoso.com)'
                $cfg['DownloadDomain'] = Read-MenuInput -Prompt 'OWA download domain     (e.g. download.contoso.com, blank = skip CVE-2021-1730)' -Validate $validateFQDN -ValidateMessage 'Not a valid FQDN (e.g. download.contoso.com)'
            }
            if ((Read-MenuInput -Prompt 'Enable log cleanup task? [Y/N]' -Default 'Y') -imatch '^[Yy]$') {
                $retDays = Read-MenuInput -Prompt 'Log retention days' -Default '30' -Required $true
                $cfg['LogRetentionDays'] = [int]$retDays
                $cfg['LogCleanupFolder'] = Read-MenuInput -Prompt 'Log cleanup script folder' -Default 'C:\#service'
            } else {
                $cfg['LogRetentionDays'] = 0
                $cfg['LogCleanupFolder'] = ''
            }
            if ((Read-MenuInput -Prompt 'Create relay connectors? [Y/N]' -Default 'N') -imatch '^[Yy]$') {
                $relay = Read-MenuInput -Prompt 'Internal relay subnets  (comma-separated CIDRs, blank = placeholder)' -Validate $validateCIDRList -ValidateMessage 'Invalid format — use e.g. 192.168.1.0/24,10.0.0.5'
                $cfg['RelaySubnets'] = if ($relay) { $relay -split '\s*,\s*' | Where-Object { $_ } } else { @('192.0.2.1/32') }
                $extRelay = Read-MenuInput -Prompt 'External relay subnets  (comma-separated CIDRs, blank = placeholder)' -Validate $validateCIDRList -ValidateMessage 'Invalid format — use e.g. 192.168.2.0/24,10.0.1.5'
                $cfg['ExternalRelaySubnets'] = if ($extRelay) { $extRelay -split '\s*,\s*' | Where-Object { $_ } } else { @('192.0.2.2/32') }
            } else {
                $cfg['RelaySubnets'] = @()
                $cfg['ExternalRelaySubnets'] = @()
            }
        }
        elseif ($selectedMode -eq 2) {
            $cfg['EdgeDNSSuffix'] = Read-MenuInput -Prompt 'Edge DNS suffix (e.g. edge.contoso.com)' -Required $true -Validate $validateFQDN -ValidateMessage 'Not a valid FQDN (e.g. edge.contoso.com)'
            $cfg['TargetPath']    = Read-MenuInput -Prompt 'Exchange install path  (blank = Exchange default)'
        }
        elseif ($selectedMode -eq 3) {
            $cfg['RecipientMgmtCleanup'] = (Read-MenuInput -Prompt 'Run AD cleanup after install? [Y/N]' -Default 'N') -imatch '^[Yy]'
        }
        elseif ($selectedMode -eq 7) {
            # Mode 7 always generates the doc; language is picked via toggle V (see toggleDefs above).
            $custInput = Read-MenuInput -Prompt 'Redact sensitive values for customer? [Y/N]' -Default 'N'
            $cfg['CustomerDocument'] = ($custInput -imatch '^[Yy]$')
        }
        elseif ($selectedMode -eq 6) {
            $cfg['Namespace']        = Read-MenuInput -Prompt 'Access namespace       (e.g. mail.contoso.com, blank = skip URL config)' -Validate $validateFQDN -ValidateMessage 'Not a valid FQDN (e.g. mail.contoso.com)'
            if ($cfg['Namespace']) {
                $defaultMailDomain2 = ($cfg['Namespace'] -split '\.', 2)[1]
                if ($defaultMailDomain2 -notmatch '\.') { $defaultMailDomain2 = $cfg['Namespace'] }
                $cfg['MailDomain']     = Read-MenuInput -Prompt 'Mail domain             (e.g. contoso.com — for Accepted Domain + email addresses)' -Default $defaultMailDomain2 -Validate $validateFQDN -ValidateMessage 'Not a valid domain (e.g. contoso.com)'
                $cfg['DownloadDomain'] = Read-MenuInput -Prompt 'OWA download domain     (e.g. download.contoso.com, blank = skip CVE-2021-1730)' -Validate $validateFQDN -ValidateMessage 'Not a valid FQDN (e.g. download.contoso.com)'
            }
            $cfg['CertificatePath']  = Read-MenuInput -Prompt 'PFX certificate path   (blank = none)        [not tested]'
            $knownDAGs2 = Get-ExchangeDAGNames
            if ($knownDAGs2.Count -gt 0) {
                Write-Host ("  DAGs found in AD: {0}" -f ($knownDAGs2 -join ', ')) -ForegroundColor DarkGray
                $cfg['DAGName'] = Read-MenuInput -Prompt ('DAG name               ({0}, blank = no DAG join)' -f ($knownDAGs2 -join ' / ')) -Default ($knownDAGs2[0])
            } else {
                $cfg['DAGName'] = Read-MenuInput -Prompt 'DAG name               (blank = no DAG join)'
            }
            if ((Read-MenuInput -Prompt 'Enable log cleanup task? [Y/N]' -Default 'Y') -imatch '^[Yy]$') {
                $retDays = Read-MenuInput -Prompt 'Log retention days' -Default '30' -Required $true
                $cfg['LogRetentionDays'] = [int]$retDays
                $cfg['LogCleanupFolder'] = Read-MenuInput -Prompt 'Log cleanup script folder' -Default 'C:\#service'
            } else {
                $cfg['LogRetentionDays'] = 0
                $cfg['LogCleanupFolder'] = ''
            }
            if ((Read-MenuInput -Prompt 'Create relay connectors? [Y/N]' -Default 'N') -imatch '^[Yy]$') {
                $relay = Read-MenuInput -Prompt 'Internal relay subnets  (comma-separated CIDRs, blank = placeholder)' -Validate $validateCIDRList -ValidateMessage 'Invalid format — use e.g. 192.168.1.0/24,10.0.0.5'
                $cfg['RelaySubnets'] = if ($relay) { $relay -split '\s*,\s*' | Where-Object { $_ } } else { @('192.0.2.1/32') }
                $extRelay = Read-MenuInput -Prompt 'External relay subnets  (comma-separated CIDRs, blank = placeholder)' -Validate $validateCIDRList -ValidateMessage 'Invalid format — use e.g. 192.168.2.0/24,10.0.1.5'
                $cfg['ExternalRelaySubnets'] = if ($extRelay) { $extRelay -split '\s*,\s*' | Where-Object { $_ } } else { @('192.0.2.2/32') }
            } else {
                $cfg['RelaySubnets'] = @()
                $cfg['ExternalRelaySubnets'] = @()
            }
        }

        # Copy toggle values into cfg
        foreach ($k in $toggleDefs.Keys) {
            $cfg[$toggleDefs[$k].Name] = $toggleState[$k]
        }
        # Advanced Configuration — persisted so state-assignment can pick it up via Test-Feature.
        $cfg['AdvancedFeatures'] = $advancedFeatures

        # --- Step 4: Summary + confirmation ---
        # Build ordered list of editable fields per mode for the E=Edit path.
        # Each entry: Key (cfg hashtable key), Label (display), Prompt (Read-Host text),
        #             Validate (scriptblock or $null), ValidateMsg, Required.
        $editFields = [System.Collections.Generic.List[hashtable]]::new()
        if ($selectedMode -in @(1, 6)) {
            if ($selectedMode -eq 1) {
                $editFields.Add(@{ Key='SourcePath';    Label='Exchange source';      Prompt='Exchange source (folder or .iso)';                               Required=$true;  Validate={ param($v) Test-Path $v }; ValidateMsg='Path not found — enter a valid folder or .iso file path' })
                $editFields.Add(@{ Key='Organization';  Label='Organization name';    Prompt='Organization name';                                              Required=$false; Validate=$null;         ValidateMsg='' })
                $editFields.Add(@{ Key='MDBName';       Label='Mailbox DB name';      Prompt='Mailbox DB name        (blank = default name)';                  Required=$false; Validate=$null;         ValidateMsg='' })
                $editFields.Add(@{ Key='MDBDBPath';     Label='Mailbox DB path';      Prompt='Mailbox DB path        (blank = Exchange default)';              Required=$false; Validate=$null;         ValidateMsg='' })
                $editFields.Add(@{ Key='MDBLogPath';    Label='Mailbox log path';     Prompt='Mailbox log path       (blank = Exchange default)';              Required=$false; Validate=$null;         ValidateMsg='' })
                $editFields.Add(@{ Key='SCP';           Label='Autodiscover SCP URL'; Prompt='Autodiscover SCP URL   (blank = let Setup set, - = remove)';    Required=$false; Validate=$null;         ValidateMsg='' })
                $editFields.Add(@{ Key='TargetPath';    Label='Exchange install path';Prompt='Exchange install path  (blank = C:\Program Files\Microsoft\Exchange Server\V15)'; Required=$false; Validate=$null; ValidateMsg='' })
                $editFields.Add(@{ Key='DAGName';       Label='DAG name';             Prompt='DAG name               (blank = no DAG join)';                  Required=$false; Validate=$validateFQDN; ValidateMsg='Not a valid FQDN (e.g. dag01.contoso.com)' })
                $editFields.Add(@{ Key='CertificatePath'; Label='PFX certificate';   Prompt='PFX certificate path   (blank = none)';                          Required=$false; Validate=$null;         ValidateMsg='' })
            }
            $editFields.Add(@{ Key='Namespace';      Label='Access Namespace';        Prompt='Access namespace       (e.g. mail.contoso.com, blank = skip URL config)'; Required=$false; Validate=$validateFQDN; ValidateMsg='Not a valid FQDN (e.g. mail.contoso.com)' })
            $editFields.Add(@{ Key='MailDomain';     Label='Mail domain';             Prompt='Mail domain             (e.g. contoso.com — for Accepted Domain + email addresses)'; Required=$false; Validate=$validateFQDN; ValidateMsg='Not a valid domain (e.g. contoso.com)' })
            $editFields.Add(@{ Key='DownloadDomain'; Label='OWA download domain';     Prompt='OWA download domain    (e.g. download.contoso.com, blank = skip CVE-2021-1730)'; Required=$false; Validate=$validateFQDN; ValidateMsg='Not a valid FQDN (e.g. download.contoso.com)' })
        }

        while ($true) {
            Clear-Host
            Write-MenuLine ('=' * 60) Cyan
            Write-MenuLine '  Summary' Cyan
            Write-MenuLine ('=' * 60) Cyan
            Write-Host ''
            Write-Host ('  Mode    : {0}' -f $modes[$selectedMode]) -ForegroundColor Green
            if ($cfg['SourcePath'])    { Write-Host ('  Source  : {0}' -f $cfg['SourcePath']) }
            Write-Host                   ('  Install : {0}' -f $cfg['InstallPath'])
            if ($cfg['Organization'])  { Write-Host ('  Org     : {0}' -f $cfg['Organization']) }
            if ($cfg['MDBName'])       { Write-Host ('  MDB     : {0}' -f $cfg['MDBName']) }
            if ($cfg['MDBDBPath'])     { Write-Host ('  DB Path : {0}' -f $cfg['MDBDBPath']) }
            if ($cfg['MDBLogPath'])    { Write-Host ('  Log Path: {0}' -f $cfg['MDBLogPath']) }
            if ($cfg['SCP'])           { Write-Host ('  SCP     : {0}' -f $cfg['SCP']) }
            if ($cfg['TargetPath'])    { Write-Host ('  Target  : {0}' -f $cfg['TargetPath']) }
            if ($cfg['DAGName'])       { Write-Host ('  DAG     : {0}' -f $cfg['DAGName']) }
            if ($cfg['Namespace'])     { Write-Host ('  Namespace: {0}' -f $cfg['Namespace']) -ForegroundColor Cyan }
            if ($cfg['MailDomain'])    { Write-Host ('  MailDomain: {0}' -f $cfg['MailDomain']) -ForegroundColor Cyan }
            if ($cfg['DownloadDomain']){ Write-Host ('  DL Domain: {0}' -f $cfg['DownloadDomain']) }
            if ($cfg['CertificatePath']){ Write-Host ('  Cert    : {0}' -f $cfg['CertificatePath']) }
            if ($cfg['EdgeDNSSuffix']) { Write-Host ('  Edge DNS: {0}' -f $cfg['EdgeDNSSuffix']) }
            # Active switches
            $finalDisabled  = @($disabledToggles[$selectedMode]) + (Get-DynamicDisabled $toggleState)
            $activeToggles = ($toggleDefs.Keys | Where-Object { $toggleState[$_] -and ($finalDisabled -notcontains $_) }) -join ', '
            if ($activeToggles) { Write-Host ('  Switches: {0}' -f $activeToggles) }
            Write-Host ''

            $editHint = if ($editFields.Count -gt 0) { ' / E=edit a field' } else { '' }
            $confirm = Read-Host ("  Start? [Y=yes{0} / N=back to menu / Q=quit]" -f $editHint)

            if ($confirm -imatch '^[Yy]') { return $cfg }
            if ($confirm -imatch '^[Qq]') { return $null }

            if ($confirm -imatch '^[Ee]' -and $editFields.Count -gt 0) {
                # Show numbered list of editable fields with current values
                Write-Host ''
                Write-Host '  Edit a field — current values:' -ForegroundColor Cyan
                for ($fi = 0; $fi -lt $editFields.Count; $fi++) {
                    $fld  = $editFields[$fi]
                    $fval = if ($cfg[$fld.Key]) { $cfg[$fld.Key] } else { '(empty)' }
                    Write-Host ('  {0,2}.  {1,-24} : {2}' -f ($fi + 1), $fld.Label, $fval)
                }
                Write-Host ''
                $pick = (Read-Host '  Field number (ENTER = cancel)').Trim()
                if ($pick -match '^\d+$') {
                    $idx = [int]$pick - 1
                    if ($idx -ge 0 -and $idx -lt $editFields.Count) {
                        $fld      = $editFields[$idx]
                        $curVal   = if ($cfg[$fld.Key]) { $cfg[$fld.Key] } else { '' }
                        $valMsg   = if ($fld.ValidateMsg) { $fld.ValidateMsg } else { 'Invalid input' }
                        $newVal   = Read-MenuInput -Prompt $fld.Prompt -Default $curVal -Required $fld.Required -Validate $fld.Validate -ValidateMessage $valMsg
                        $cfg[$fld.Key] = $newVal
                        # Clear DownloadDomain if Namespace was cleared
                        if ($fld.Key -eq 'Namespace' -and -not $newVal) { $cfg['DownloadDomain'] = '' }
                    }
                }
                continue
            }

            Write-MyVerbose 'Menu: Back to mode selection'
            # N or anything else = restart from mode selection
            $selectedMode = 0
            while ($selectedMode -lt 1 -or $selectedMode -gt 7) {
                Draw-Menu -Mode $selectedMode -ToggState $toggleState
                $raw = Read-Host '  Mode [1-7]'
                if ($raw -match '^[1-7]$') { $selectedMode = [int]$raw }
            }
        }
    }
