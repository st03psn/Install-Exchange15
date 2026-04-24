    function Install-PendingWindowsUpdates {
        # Installs pending Windows security and critical updates.
        # Interactive mode: prompts per update (Y/N/A=all/S=skip rest).
        # Autopilot mode:   installs all without prompting.
        # Download + install runs in a background job with $WU_DOWNLOAD_TIMEOUT_SEC timeout;
        # on timeout the update step is skipped and Exchange installation continues.
        # Uses PSWindowsUpdate module when available; falls back to Windows Update Agent COM API.
        # Sets $State['RebootRequired'] = $true when a reboot is needed after updates.

        if (-not $State['InstallWindowsUpdates']) {
            Write-MyVerbose 'InstallWindowsUpdates not set, skipping Windows Update check'
            return
        }

        # Interactive prompts whenever a real console is available.
        # Autopilot does NOT suppress the prompt — if someone is at the keyboard they can still
        # review each update. In a truly headless run [Environment]::UserInteractive is $false.
        $isInteractive = [Environment]::UserInteractive

        Write-MyOutput 'Checking for pending Windows Updates (Security + Critical)'

        # --- Detect PSWindowsUpdate module ---
        $useModule = $false
        if (Get-Module -ListAvailable -Name PSWindowsUpdate -ErrorAction SilentlyContinue) {
            $useModule = $true
        }
        else {
            Write-MyVerbose 'PSWindowsUpdate module not found, attempting to install from PSGallery'
            try {
                # Ensure NuGet provider present unattended — without this Install-Module
                # prompts interactively even in non-interactive/Autopilot sessions.
                # Install-PackageProvider may fail to reach the provider index URI but
                # Install-Module -ForceBootstrap handles NuGet bootstrap itself without prompting.
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null
                Install-Module -Name PSWindowsUpdate -Scope CurrentUser -Force -AllowClobber -ForceBootstrap -ErrorAction Stop
                $useModule = $true
                Write-MyOutput 'PSWindowsUpdate module installed'
            }
            catch {
                Write-MyWarning ('Could not install PSWindowsUpdate: {0}. Falling back to WUA COM API' -f $_.Exception.Message)
            }
        }

        # --- Phase 1: Search (fast, runs in main thread) ---
        $candidates = @()   # [PSCustomObject]@{ Title; KB; Severity }

        if ($useModule) {
            try {
                Import-Module PSWindowsUpdate -ErrorAction Stop
                $wuList = Get-WindowsUpdate -Category 'Security Updates','Critical Updates' -NotTitle 'Preview' -ErrorAction Stop
                $candidates = @($wuList | ForEach-Object {
                    [PSCustomObject]@{ Title = $_.Title; KB = $_.KB; Severity = $_.MsrcSeverity }
                })
            }
            catch {
                Write-MyWarning ('PSWindowsUpdate search error: {0}' -f $_.Exception.Message)
            }
        }
        else {
            try {
                $wuaSession  = New-Object -ComObject Microsoft.Update.Session
                $wuaSearcher = $wuaSession.CreateUpdateSearcher()
                $wuaResult   = $wuaSearcher.Search('IsInstalled=0 and IsHidden=0 and BrowseOnly=0')
                $candidates  = @(foreach ($u in $wuaResult.Updates) {
                    if ($u.MsrcSeverity -in @('Critical','Important') -or $u.AutoSelectOnWebSites) {
                        [PSCustomObject]@{ Title = $u.Title; KB = ($u.KBArticleIDs | Select-Object -First 1); Severity = $u.MsrcSeverity }
                    }
                })
            }
            catch {
                Write-MyWarning ('WUA COM API search error: {0}' -f $_.Exception.Message)
            }
        }

        if ($candidates.Count -eq 0) {
            Write-MyOutput 'No pending Windows security/critical updates found'
            return
        }

        Write-MyOutput ('{0} update(s) found' -f $candidates.Count)

        # --- Phase 2: Per-update prompt ---
        # Autopilot auto-approves only when AutoApproveWindowsUpdates is explicitly set in
        # the Advanced Configuration — security updates are a deliberate opt-in decision.
        $approvedKBs     = @()
        $autoApproveAll  = (-not $isInteractive) -and $State['AutoApproveWindowsUpdates']

        if ((-not $isInteractive) -and (-not $State['AutoApproveWindowsUpdates'])) {
            Write-MyWarning ('Found {0} pending Windows update(s) — skipping in Autopilot because AutoApproveWindowsUpdates is not set. Enable it in Advanced Configuration to install automatically.' -f $candidates.Count)
            $candidates | ForEach-Object { Write-MyVerbose ('  Pending: {0} ({1})' -f $_.Title, $_.Severity) }
            return
        }

        for ($idx = 0; $idx -lt $candidates.Count; $idx++) {
            $u = $candidates[$idx]
            $label = '[{0}/{1}] {2} — {3}' -f ($idx + 1), $candidates.Count, $u.Title, $(if ($u.Severity) { $u.Severity } else { 'Unknown' })

            if ($autoApproveAll) {
                Write-MyOutput ('Auto-approved: {0}' -f $label)
                if ($u.KB) { $approvedKBs += $u.KB }
                continue
            }

            # Timed prompt: auto-skip (N) after 120 seconds with no keypress.
            # Uses RawUI.ReadKey so no Enter is required; falls back to Read-Host
            # (blocking, no timeout) when console is unavailable (redirected stdin).
            $WU_PROMPT_TIMEOUT_SEC = 120
            Write-Host ('{0}' -f $label) -ForegroundColor Cyan
            $ans = ''
            if ($host.UI.RawUI -and $host.UI.RawUI.KeyAvailable -ne $null) {
                # Flush any buffered keystrokes (e.g. from credential prompts or prior Read-Host
                # calls) so a stale keystroke doesn't immediately resolve the prompt as 'N'.
                try { $host.UI.RawUI.FlushInputBuffer() } catch { }
                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                Write-Host ('  Install? [Y/N/S=skip remaining] (auto-No in {0}s) ' -f $WU_PROMPT_TIMEOUT_SEC) -NoNewline -ForegroundColor DarkCyan
                while ($sw.Elapsed.TotalSeconds -lt $WU_PROMPT_TIMEOUT_SEC) {
                    $secsLeft = [int]($WU_PROMPT_TIMEOUT_SEC - $sw.Elapsed.TotalSeconds)
                    Write-Progress -Id 2 -Activity 'Windows Update' `
                        -Status ('Auto-No in {0}s  |  Y = install  |  N = skip  |  S = skip remaining' -f $secsLeft) `
                        -PercentComplete ($sw.Elapsed.TotalSeconds * 100 / $WU_PROMPT_TIMEOUT_SEC)
                    if ($host.UI.RawUI.KeyAvailable) {
                        $key = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                        $ans = $key.Character.ToString().ToUpper()
                        Write-Host $ans
                        break
                    }
                    Start-Sleep -Milliseconds 200
                }
                Write-Progress -Id 2 -Activity 'Windows Update' -Completed
                if ($ans -eq '') {
                    Write-Host 'N (timeout)'
                    $ans = 'N'
                }
            }
            else {
                $ans = (Read-Host '  Install? [Y=yes / N=no / S=skip remaining] (default: Y)').Trim().ToUpper()
                if ($ans -eq '') { $ans = 'Y' }
            }
            switch ($ans) {
                'S' { Write-MyOutput 'Skipping all remaining updates'; $idx = $candidates.Count; continue }
                'N' { Write-MyOutput ('Skipping: {0}' -f $u.Title) }
                default { if ($u.KB) { $approvedKBs += $u.KB }; Write-MyOutput ('Approved: {0}' -f $u.Title) }
            }
        }

        if ($approvedKBs.Count -eq 0) {
            Write-MyOutput 'No updates approved for installation — skipping Windows Update step'
            return
        }

        # --- Phase 3: Download + Install in background job with timeout ---
        Write-MyOutput ('Installing {0} approved update(s) (timeout: {1}s) ...' -f $approvedKBs.Count, $WU_DOWNLOAD_TIMEOUT_SEC)

        if ($useModule) {
            $wuJob = Start-Job -ScriptBlock {
                param([string[]]$kbs)
                Import-Module PSWindowsUpdate -ErrorAction Stop
                $result = Install-WindowsUpdate -KBArticleID $kbs -AcceptAll -IgnoreReboot -ErrorAction Stop
                $result | Select-Object Title, KB, Result, RebootRequired
            } -ArgumentList (,$approvedKBs)
        }
        else {
            $wuJob = Start-Job -ScriptBlock {
                param([string[]]$kbs)
                $session  = New-Object -ComObject Microsoft.Update.Session
                $searcher = $session.CreateUpdateSearcher()
                $filter   = ($kbs | ForEach-Object { "KBArticleID='$_'" }) -join ' or '
                $found    = $searcher.Search("IsInstalled=0 and ($filter)")
                if ($found.Updates.Count -eq 0) { return @{ Installed=0; RebootRequired=$false } }
                $dl = $session.CreateUpdateDownloader()
                $dl.Updates = $found.Updates
                $dl.Download() | Out-Null
                $inst       = $session.CreateUpdateInstaller()
                $inst.Updates = $found.Updates
                $instResult = $inst.Install()
                @{ Installed = $found.Updates.Count; RebootRequired = $instResult.RebootRequired; ResultCode = $instResult.ResultCode }
            } -ArgumentList (,$approvedKBs)
        }

        # --- Polling loop: show progress + allow keyboard cancellation ---
        $pollInterval = 5   # seconds between status checks
        $elapsed      = 0
        $cancelled    = $false
        Write-Host '  Press X to cancel Windows Update installation at any time.' -ForegroundColor DarkGray

        while ($wuJob.State -eq 'Running') {
            Start-Sleep -Seconds $pollInterval
            $elapsed += $pollInterval

            $remaining  = $WU_DOWNLOAD_TIMEOUT_SEC - $elapsed
            $pct        = [Math]::Min(99, [int]($elapsed * 100 / $WU_DOWNLOAD_TIMEOUT_SEC))
            $statusText = 'Installing {0} update(s) — {1}s elapsed  |  auto-abort in {2}s  |  X = cancel' -f $approvedKBs.Count, $elapsed, $remaining
            Write-Progress -Activity 'Windows Updates' -Status $statusText -PercentComplete $pct

            # Non-blocking key check for cancellation
            if ([Console]::KeyAvailable) {
                $key = [Console]::ReadKey($true)
                if ($key.Key -in @([ConsoleKey]::X, [ConsoleKey]::Q)) {
                    Write-Progress -Activity 'Windows Updates' -Completed
                    Write-MyWarning 'Windows Update installation cancelled by user — continuing Exchange installation without updates'
                    Stop-Job  $wuJob -ErrorAction SilentlyContinue
                    Remove-Job $wuJob -Force -ErrorAction SilentlyContinue
                    $cancelled = $true
                    break
                }
            }

            if ($elapsed -ge $WU_DOWNLOAD_TIMEOUT_SEC) {
                Write-Progress -Activity 'Windows Updates' -Completed
                Stop-Job  $wuJob -ErrorAction SilentlyContinue
                Remove-Job $wuJob -Force -ErrorAction SilentlyContinue
                Write-MyWarning ('Windows Update timed out after {0}s — continuing Exchange installation without updates' -f $WU_DOWNLOAD_TIMEOUT_SEC)
                $cancelled = $true
                break
            }
        }
        Write-Progress -Activity 'Windows Updates' -Completed
        if ($cancelled) { return }

        $jobOut    = Receive-Job $wuJob -ErrorVariable wuErrors
        Remove-Job $wuJob -Force -ErrorAction SilentlyContinue

        if ($wuErrors) {
            Write-MyWarning ('Windows Update error: {0}' -f $wuErrors[0].Exception.Message)
        }

        $rebootNeeded = $false
        if ($useModule) {
            $installed    = @($jobOut | Where-Object { $_.Result -eq 'Installed' -and $_.KB -and ($approvedKBs -contains $_.KB) }).Count
            $rebootNeeded = ($jobOut | Where-Object { $_.RebootRequired }) -as [bool]
            Write-MyOutput ('{0} update(s) installed' -f $installed)
        }
        else {
            $rebootNeeded = $jobOut.RebootRequired
            Write-MyOutput ('{0} update(s) installed, WUA result code: {1}' -f $jobOut.Installed, $jobOut.ResultCode)
        }

        if ($rebootNeeded) {
            Write-MyWarning 'Windows Updates require a reboot'
            $State['RebootRequired'] = $true
        }
        else {
            Write-MyOutput 'Windows Updates installed, no reboot required'
        }
    }

    # Known Exchange Security Updates (SU): hashtable of SetupVersion -> SU info
    # Format: @{ '<ExSetup build>' = @{ KB='KBxxxxxxx'; URL='<msp url>'; TargetVersion='<build after SU>' } }
    # Maps RTM setup.exe version -> latest known Security Update.
    # Keys are ExSetup.exe ProductVersion strings (from Get-DetectedFileVersion on setup.exe).
    # FileName must be the .exe installer name; URL must be a direct download link.
    # Update this table whenever Microsoft releases a new Exchange Security Update.
    $ExchangeSUMap = @{
        # Exchange SE RTM (15.02.2562.017) -> Feb 2026 SU (KB5074992)
        # No URL: the WU-catalog CAB is not installable via DISM/expand.exe.
        # Place ExchangeSubscriptionEdition-KB5074992-x64-en.exe (from Microsoft Download Center)
        # in <InstallPath>\sources\ before running, or apply via Windows Update / WSUS.
        '15.02.2562.017' = @{
            KB            = 'KB5074992'
            FileName      = 'ExchangeSubscriptionEdition-KB5074992-x64-en.exe'
            URL           = $null
            TargetVersion = '15.02.2562.037'
        }
        # Exchange 2019 CU15 (15.02.1748.008) -> Jan 2025 SU (KB5049233 SU3 V2)
        '15.02.1748.008' = @{
            KB            = 'KB5049233'
            FileName      = 'Exchange2019-KB5049233-x64-en.exe'
            URL           = 'https://download.microsoft.com/download/8/0/b/80b356e4-f7b1-4e11-9586-d3132a7a2fc3/Exchange2019-KB5049233-x64-en.exe'
            TargetVersion = '15.02.1748.016'
        }
        # Exchange 2019 CU14 (15.02.1544.004) -> Jan 2025 SU (KB5049233 SU3 V2)
        '15.02.1544.004' = @{
            KB            = 'KB5049233'
            FileName      = 'Exchange2019-KB5049233-x64-en.exe'
            URL           = 'https://download.microsoft.com/download/8/0/b/80b356e4-f7b1-4e11-9586-d3132a7a2fc3/Exchange2019-KB5049233-x64-en.exe'
            TargetVersion = '15.02.1544.014'
        }
        # Exchange 2019 CU13 (15.02.1258.012) -> Jan 2025 SU (KB5049233 SU7 V2)
        '15.02.1258.012' = @{
            KB            = 'KB5049233'
            FileName      = 'Exchange2019-KB5049233-x64-en.exe'
            URL           = 'https://download.microsoft.com/download/4/e/5/4e5cbbcc-5894-457d-88c4-c0b2ff7f208f/Exchange2019-KB5049233-x64-en.exe'
            TargetVersion = '15.02.1258.032'
        }
        # Exchange 2016 CU23 (15.01.2507.006) -> Jan 2025 SU (KB5049233 SU14 V2)
        '15.01.2507.006' = @{
            KB            = 'KB5049233'
            FileName      = 'Exchange2016-KB5049233-x64-en.exe'
            URL           = 'https://download.microsoft.com/download/0/9/9/0998c26c-8eb6-403a-b97a-ae44c4db5e20/Exchange2016-KB5049233-x64-en.exe'
            TargetVersion = '15.01.2507.043'
        }
    }

    function Get-LatestExchangeSecurityUpdate {
        # Returns SU info hashtable for the currently installed Exchange setup version, or $null if up to date / not applicable.
        $currentBuild = $State['SetupVersion']
        if (-not $currentBuild) { return $null }
        if ($ExchangeSUMap.ContainsKey($currentBuild)) {
            return $ExchangeSUMap[$currentBuild]
        }
        return $null
    }

    function Get-InstalledExchangeBuild {
        # Returns the installed Exchange build from the MSExchangeServiceHost service binary.
        try {
            $svcPath = (Get-CimInstance -Query 'SELECT * FROM win32_service WHERE name="MSExchangeServiceHost"' -ErrorAction Stop).PathName
            if ($svcPath) { return Get-DetectedFileVersion $svcPath.Trim('"') }
        }
        catch { }
        return $null
    }

    function Get-LatestSUBuildFromHC {
        # Parses HealthChecker.ps1's GetExchangeBuildDictionary to find the latest known SU
        # build for the installed Exchange CU. Returns a version string ('15.02.1748.043') or $null.
        $hcPath = Join-Path $State['SourcesPath'] 'HealthChecker.ps1'
        if (-not (Test-Path $hcPath)) { return $null }

        # Map setup.exe version to HC CU key
        $cuLookup = @{
            $EXSESETUPEXE_RTM    = 'RTM'
            $EX2019SETUPEXE_CU15 = 'CU15'
            $EX2019SETUPEXE_CU14 = 'CU14'
            $EX2019SETUPEXE_CU13 = 'CU13'
            $EX2016SETUPEXE_CU23 = 'CU23'
        }
        $cu = $cuLookup[$State['ExSetupVersion']]
        if (-not $cu) { return $null }

        try { $hcContent = Get-Content $hcPath -Raw -ErrorAction Stop }
        catch { return $null }

        # Find the CU block in GetExchangeBuildDictionary:
        #   "RTM"|"CUxx" = (NewCUAndSUObject "base.build" @{ "FebxxSU" = "x.x.x.x" ... })
        $cuPattern = '"' + [regex]::Escape($cu) + '"\s*=\s*\(NewCUAndSUObject\s+"[\d.]+"\s+@\{([^}]+)\}\)'
        $cuMatch   = [regex]::Match($hcContent, $cuPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
        if (-not $cuMatch.Success) { return $null }

        # Extract all SU version strings and pick the highest
        $builds = [regex]::Matches($cuMatch.Groups[1].Value, '"[\w]+"\s*=\s*"(\d+\.\d+\.\d+\.\d+)"') |
                  ForEach-Object { [System.Version]$_.Groups[1].Value } |
                  Sort-Object -Descending
        if (-not $builds -or $builds.Count -eq 0) { return $null }

        # Normalise from HC format (15.2.1748.43) to script format (15.02.1748.043)
        $b = $builds[0]
        return '{0}.{1:D2}.{2}.{3:D3}' -f $b.Major, $b.Minor, $b.Build, $b.Revision
    }

    function Install-ExchangeSecurityUpdate {
        # Downloads and installs an Exchange Security Update (.exe, .cab, or .msp).
        # P6: also does a dynamic gap-check against HealthChecker.ps1's build dictionary.
        if (-not $State['IncludeFixes']) {
            Write-MyVerbose 'IncludeFixes not set, skipping Exchange SU check'
            return
        }

        # Get the currently installed Exchange build; skip redundant reinstalls
        $installedBuild = Get-InstalledExchangeBuild
        if ($installedBuild) { Write-MyVerbose ('Installed Exchange build: {0}' -f $installedBuild) }

        $su = Get-LatestExchangeSecurityUpdate
        # B15: skip if we already installed this exact KB in a previous phase-5 run.
        # Exchange SU installers may trigger their own system reboot before the script's
        # phase-end logic runs (Enable-RunOnce / LastSuccessfulPhase update). On the next
        # run, the build version reported by Get-InstalledExchangeBuild may still show the
        # pre-SU value (service binary cache / timing), causing an endless install loop.
        # Persisting a per-KB flag in State prevents the reinstall.
        if ($su) {
            $suFlag = 'ExchangeSUInstalled_{0}' -f $su.KB
            if ($State[$suFlag]) {
                Write-MyVerbose ('Exchange SU {0} already installed in a previous run — skipping' -f $su.KB)
                return
            }
        }

        if (-not $su) {
            Write-MyOutput 'No known Exchange Security Update applicable for this build'
        }
        else {
            $targetVer    = try { [System.Version]$su.TargetVersion } catch { $null }
            $installedVer = if ($installedBuild) { try { [System.Version]$installedBuild } catch { $null } } else { $null }

            if ($installedVer -and $targetVer -and $installedVer -ge $targetVer) {
                Write-MyOutput ('Exchange build {0} already at or above SU target {1} ({2}), skipping install' -f $installedBuild, $su.TargetVersion, $su.KB)
            }
            else {
                Write-MyOutput ('Exchange Security Update {0} available for build {1} -> {2}' -f $su.KB, $State['SetupVersion'], $su.TargetVersion)
                $suPath = Join-Path $State['SourcesPath'] $su.FileName
                if (-not (Test-Path $suPath)) {
                    if ($su.URL) {
                        Write-MyOutput ('Downloading {0}' -f $su.KB)
                        $null = Get-MyPackage -Package $su.KB -URL $su.URL -FileName $su.FileName -InstallPath $State['SourcesPath']
                    }
                    if (-not (Test-Path $suPath)) {
                        Write-MyWarning ('Exchange SU {0}: installer not available for automatic download.' -f $su.KB)
                        Write-MyOutput  ('  Download:  https://support.microsoft.com/help/{0}' -f ($su.KB -replace '^KB', ''))
                        Write-MyOutput  ('  Place EXE: {0}' -f $suPath)

                        # Interactive countdown — user has 5 min to place the file, then ENTER to install.
                        # Autopilot / non-interactive: skip silently (no file available, no reboot loop).
                        if ([Environment]::UserInteractive -and -not $State['ConfigDriven']) {
                            Write-MyOutput 'Place the installer, then press ENTER — or skip now with ENTER / auto-skip after 5 min:'
                            $suTotalSecs = 300
                            $suDeadline  = [DateTime]::Now.AddSeconds($suTotalSecs)
                            try {
                                try { $host.UI.RawUI.FlushInputBuffer() } catch { }
                                while ([DateTime]::Now -lt $suDeadline) {
                                    $secsLeft = [int]($suDeadline - [DateTime]::Now).TotalSeconds
                                    Write-Progress -Id 2 -Activity ('Exchange SU {0}' -f $su.KB) `
                                        -Status ('Place {0} in {1} then ENTER  |  auto-skip in {2}s' -f $su.FileName, $State['SourcesPath'], $secsLeft) `
                                        -PercentComplete ([int](($suTotalSecs - $secsLeft) * 100 / $suTotalSecs))
                                    if ($host.UI.RawUI.KeyAvailable) {
                                        $key = $host.UI.RawUI.ReadKey('IncludeKeyDown,NoEcho')
                                        Write-Host ''
                                        if ($key.VirtualKeyCode -in 13, 27) { break }
                                    }
                                    Start-Sleep -Milliseconds 100
                                }
                                Write-Progress -Id 2 -Activity ('Exchange SU {0}' -f $su.KB) -Completed
                            }
                            catch { }
                        }
                    }
                }
                if (Test-Path $suPath) {
                    Write-MyOutput ('Installing Exchange SU {0}' -f $su.KB)
                    # B15: In Autopilot mode, pre-set RunOnce + save state before launching the
                    # installer. Exchange SU installers (.exe) may call ExitWindowsEx internally
                    # and reboot the machine before this script's phase-end logic runs, leaving
                    # LastSuccessfulPhase = 4 and no RunOnce set — so the script would not
                    # auto-resume. Pre-setting RunOnce here ensures the script always restarts.
                    if ($State['Autopilot']) {
                        Disable-UAC
                        Enable-AutoLogon
                        Enable-RunOnce
                        Save-State $State
                    }
                    # Exchange SU installers only accept /passive or /silent — /norestart is not supported.
                    # Exit code 3010 = success + reboot required; handled below.
                    $rc = Invoke-Process -FilePath $State['SourcesPath'] -FileName $su.FileName -ArgumentList '/passive'
                    if ($rc -eq 0 -or $rc -eq 3010) {
                        Write-MyOutput ('Exchange SU {0} installed successfully' -f $su.KB)
                        # Persist a per-KB installed flag immediately so phase-5 re-entry after
                        # the reboot skips the SU (build version check alone is unreliable when
                        # the service binary cache has not yet been flushed after the SU reboot).
                        $State['ExchangeSUInstalled_{0}' -f $su.KB] = $true
                        Save-State $State
                        if ($rc -eq 3010) {
                            Write-MyWarning 'Exchange SU requires a reboot'
                            $State['RebootRequired'] = $true
                        }
                    }
                    else {
                        Write-MyWarning ('Exchange SU {0} install failed (exit code {1}). Try applying via Windows Update or see https://support.microsoft.com/help/{2}' -f $su.KB, $rc, ($su.KB -replace '^KB', ''))
                    }
                }
            }
        }

        # P6 — Dynamic gap-check: download HC.ps1 if not present and compare installed
        # build against HC's GetExchangeBuildDictionary (single attempt, non-blocking).
        $hcPath = Join-Path $State['SourcesPath'] 'HealthChecker.ps1'
        if (-not (Test-Path $hcPath)) {
            try {
                Write-MyVerbose 'Downloading HealthChecker.ps1 for Exchange SU version check'
                Invoke-WebDownload -Uri 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/HealthChecker.ps1' -OutFile $hcPath
            }
            catch { Write-MyVerbose ('Could not download HealthChecker.ps1 for SU check: {0}' -f $_.Exception.Message) }
        }

        $hcLatest = Get-LatestSUBuildFromHC
        if ($hcLatest) {
            $hcLatestVer  = try { [System.Version]$hcLatest } catch { $null }
            # Re-query installed build after potential SU install above
            $currentBuild = Get-InstalledExchangeBuild
            $currentVer   = if ($currentBuild) { try { [System.Version]$currentBuild } catch { $null } } else { $null }
            if ($currentVer -and $hcLatestVer) {
                if ($currentVer -lt $hcLatestVer) {
                    Write-MyWarning ('Exchange build {0} is behind latest known SU {1} (per HealthChecker). Newer SU may require ESU enrollment — see https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates for the latest update.' -f $currentBuild, $hcLatest)
                }
                else {
                    Write-MyOutput ('Exchange build {0} is current per HealthChecker (latest known: {1})' -f $currentBuild, $hcLatest)
                }
            }
        }
    }

