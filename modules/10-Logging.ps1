    # ── Modern step/section/phase output ─────────────────────────────────────
    # Console rendering with consistent symbols + colors. Log files receive plain
    # text via Write-ToTranscript so machine readability is unchanged.

    function Write-MySection {
        # Group header — separates logical blocks (Preflight / Post-Config / etc.).
        param([Parameter(Mandatory)][string]$Title)
        $width = 64
        $line  = '=' * 3 + ' ' + $Title + ' '
        if ($line.Length -lt $width) { $line += '=' * ($width - $line.Length) }
        Write-Host ''
        Write-Host $line -ForegroundColor Cyan
        Write-ToTranscript 'INFO' ('=== {0} ===' -f $Title)
    }

    function Write-MyStep {
        # Two-column status line: [TAG] Label                Value
        # Status: OK (green), Warn (yellow), Fail (red), Run (cyan), Info (gray).
        param(
            [Parameter(Mandatory)][string]$Label,
            [string]$Value = '',
            [ValidateSet('OK','Warn','Fail','Run','Info')]
            [string]$Status = 'OK'
        )
        $tag = switch ($Status) {
            'OK'   { '[OK]'   }
            'Warn' { '[!] '   }
            'Fail' { '[X] '   }
            'Run'  { '[->]'   }
            'Info' { '[ i]'   }
        }
        $color = switch ($Status) {
            'OK'   { 'Green'    }
            'Warn' { 'Yellow'   }
            'Fail' { 'Red'      }
            'Run'  { 'DarkCyan' }
            'Info' { 'Gray'     }
        }
        $labelPad = $Label.PadRight(28)
        Write-Host -NoNewline '  '
        Write-Host -NoNewline $tag -ForegroundColor $color
        Write-Host -NoNewline ('  {0}' -f $labelPad)
        if ($Value) { Write-Host $Value -ForegroundColor DarkGray } else { Write-Host '' }
        $logBody = if ($Value) { '{0}: {1}' -f $Label, $Value } else { $Label }
        $logLine = '{0}  {1}' -f $tag.Trim(), $logBody
        $logLevel = switch ($Status) { 'Warn' { 'WARNING' } 'Fail' { 'ERROR' } default { 'INFO' } }
        Write-ToTranscript $logLevel $logLine
    }

    function Write-MyPhase {
        # Boxed phase banner. Replaces the old "Exchange Installation\n  Phase X of N: ..." pair.
        param(
            [Parameter(Mandatory)][int]$Number,
            [Parameter(Mandatory)][int]$Total,
            [Parameter(Mandatory)][string]$Title
        )
        $width = 64
        $inner = $width - 4
        $text  = ('Phase {0} / {1}   -   {2}' -f $Number, $Total, $Title)
        if ($text.Length -gt $inner) { $text = $text.Substring(0, $inner) }
        $padded = $text.PadRight($inner)
        $border = '+' + ('-' * ($width - 2)) + '+'
        Write-Host ''
        Write-Host $border -ForegroundColor Cyan
        Write-Host -NoNewline '|  ' -ForegroundColor Cyan
        Write-Host -NoNewline $padded -ForegroundColor White
        Write-Host '|' -ForegroundColor Cyan
        Write-Host $border -ForegroundColor Cyan
        Write-ToTranscript 'INFO' ('=== Phase {0} of {1}: {2} ===' -f $Number, $Total, $Title)
    }

    function Show-EXpressBanner {
        # Startup banner. Pure ASCII so every codepage renders cleanly
        # (CP437 / CP1252 / UTF-8). Skipped if no console is attached.
        param([string]$Version = '')
        try { if (-not $Host.UI.RawUI.WindowSize) { return } } catch { return }

        $width = 64
        $bar   = '#' * $width

        function Format-Centered {
            param([string]$Text, [int]$Width)
            $pad = [Math]::Max(0, $Width - $Text.Length)
            $left  = [int]($pad / 2)
            $right = $pad - $left
            (' ' * $left) + $Text + (' ' * $right)
        }

        $verTxt = if ($Version) { "v$Version" } else { '' }
        $title  = ('EXpress  {0}' -f $verTxt).Trim()
        $tag    = 'Unattended Exchange Server installation'
        $by     = 'github.com/st03psn/EXpress  -  st03psn'

        Write-Host ''
        Write-Host $bar -ForegroundColor Cyan
        Write-Host (Format-Centered '' $width) -ForegroundColor Cyan
        Write-Host (Format-Centered $title $width) -ForegroundColor Cyan
        Write-Host (Format-Centered $tag   $width) -ForegroundColor Gray
        Write-Host (Format-Centered $by    $width) -ForegroundColor DarkGray
        Write-Host (Format-Centered '' $width) -ForegroundColor Cyan
        Write-Host $bar -ForegroundColor Cyan
        Write-Host ''
    }

    function Write-ToTranscript( $Level, $Text) {
        # Single log file with tier filtering — driven by $State['LogVerbose'/'LogDebug']:
        #   Default       : INFO / WARNING / ERROR / EXE
        #   -Verbose      : + VERBOSE
        #   -Debug        : + VERBOSE + DEBUG + CMD + SUPPRESSED-ERROR diff scan
        # Encoding note: pinned UTF-8 (no BOM) via [IO.File]::AppendAllText so every
        # line shares the same encoding (avoids the PS 5.1 Out-File UTF-16 mix problem).
        if (-not $State['TranscriptFile']) { return }
        $Location = Split-Path $State['TranscriptFile'] -Parent
        if (-not (Test-Path $Location)) { return }
        $verboseOn = [bool]$State['LogVerbose']
        $debugOn   = [bool]$State['LogDebug']
        $shouldWrite = switch ($Level) {
            'VERBOSE' { $verboseOn -or $debugOn }
            'DEBUG'   { $debugOn }
            'CMD'     { $debugOn }
            default   { $true }
        }
        $utf8 = [System.Text.UTF8Encoding]::new($false)
        if ($shouldWrite) {
            try {
                [System.IO.File]::AppendAllText($State['TranscriptFile'], ("{0}: [{1}] {2}`r`n" -f (Get-Date -Format u), $Level, $Text), $utf8)
            } catch { Write-Verbose "Write-ToTranscript: AppendAllText failed: $_" }
        }
        # SUPPRESSED-ERROR diff scan: only in -Debug mode, appended to the same install log.
        if ($debugOn) {
            try {
                $cur = $Error.Count
                if ($cur -gt $script:lastErrorCount) {
                    $newCount = $cur - $script:lastErrorCount
                    for ($i = $newCount - 1; $i -ge 0; $i--) {
                        $e = $Error[$i]
                        if (-not $e) { continue }
                        $inv = $e.InvocationInfo
                        $ln  = if ($inv) { $inv.ScriptLineNumber } else { '?' }
                        $cmd = if ($inv) { ($inv.Line -replace '\s+', ' ').Trim() } else { '' }
                        $typ = if ($e.Exception) { $e.Exception.GetType().FullName } else { 'Error' }
                        $msg = if ($e.Exception) { $e.Exception.Message } else { [string]$e }
                        $line = '{0}: [SUPPRESSED-ERROR] ({1}) at line {2}: {3} :: {4}' -f (Get-Date -Format u), $typ, $ln, $cmd, $msg
                        [System.IO.File]::AppendAllText($State['TranscriptFile'], ($line + "`r`n"), $utf8)
                    }
                    $script:lastErrorCount = $cur
                }
            } catch { Write-Verbose "Write-ToTranscript: SUPPRESSED-ERROR scan failed: $_" }
        }
    }

    function Write-MyOutput( $Text) {
        Write-Output $Text
        Write-ToTranscript 'INFO' $Text
    }

    function Write-MyWarning( $Text) {
        Write-Warning $Text
        Write-ToTranscript 'WARNING' $Text
        $script:nonFatalErrorCount++
    }

    function Write-MyError( $Text) {
        Write-Error $Text
        Write-ToTranscript 'ERROR' $Text
        $script:nonFatalErrorCount++
    }

    function Write-MyVerbose( $Text) {
        # Console: render in the same indented + tagged style as Write-MyStep so
        # verbose output sits visually next to OK/Run/Info lines instead of breaking
        # alignment. Tag is `[..]` in DarkGray, text DarkGray. Only rendered when
        # $State['LogVerbose'] is on (set by -Verbose or -Debug).
        if ($State['LogVerbose']) {
            Write-Host -NoNewline '  '
            Write-Host -NoNewline '[..]' -ForegroundColor DarkGray
            Write-Host ('  ' + $Text) -ForegroundColor DarkGray
        }
        Write-ToTranscript 'VERBOSE' $Text
    }

    function Write-MyDebug( $Text) {
        # Console stays silent; log line appears only when -Debug tier active.
        Write-ToTranscript 'DEBUG' $Text
    }

    # Records configuration-level commands the script actually executed, so
    # chapter 14 of the Installation Document ("Executed Cmdlets") can list them
    # chronologically with exact syntax. Call sites pass the same command line
    # they are about to run — the helper does not re-execute, only records.
    function Register-ExecutedCommand {
        param(
            [Parameter(Mandatory)][string]$Command,
            [string]$Category = ''
        )
        # After Import-Clixml (Restore-State across a reboot) the list comes back as a
        # frozen Deserialized.* type with no .Add() method. Rehydrate it into a live
        # List, preserving any entries recorded before the reboot, on first post-reboot call.
        if (-not $State.ContainsKey('ExecutedCommands') -or $null -eq $State['ExecutedCommands']) {
            $State['ExecutedCommands'] = [System.Collections.Generic.List[object]]::new()
        }
        elseif ($State['ExecutedCommands'] -isnot [System.Collections.Generic.List[object]]) {
            $rehydrated = [System.Collections.Generic.List[object]]::new()
            foreach ($item in @($State['ExecutedCommands'])) { $rehydrated.Add($item) }
            $State['ExecutedCommands'] = $rehydrated
        }
        # Skip exact duplicates that accumulate when a phase re-runs after a reboot
        # (same Phase+Category+Command already in list from prior run).
        $alreadyRecorded = $State['ExecutedCommands'] | Where-Object {
            $_.Phase -eq [int]($State['InstallPhase']) -and $_.Category -eq $Category -and $_.Command -eq $Command
        }
        if ($alreadyRecorded) { return }
        $State['ExecutedCommands'].Add([pscustomobject]@{
            Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            Phase     = [int]($State['InstallPhase'])
            Category  = $Category
            Command   = $Command
        })
        Write-ToTranscript 'CMD' $Command
    }

    function Write-DebugSnapshot {
        # Per-phase point-in-time snapshot of state + system info under .\Debug\.
        # Complementary to the install log (which carries the streaming narrative
        # with DEBUG/CMD/SUPPRESSED-ERROR added when -Debug is on).
        # No-op unless $State['LogDebug'] is true.
        param([string]$Marker = '')
        if (-not $State['LogDebug']) { return }

        $debugDir = Join-Path $State['InstallPath'] 'Debug'
        try { if (-not (Test-Path $debugDir)) { New-Item -Path $debugDir -ItemType Directory -Force | Out-Null } } catch { return }

        $ts    = Get-Date -Format 'yyyyMMdd-HHmmss'
        $phase = if ($State['InstallPhase']) { $State['InstallPhase'] } else { 0 }
        $tag   = 'Phase{0}_{1}' -f $phase, $ts
        if ($Marker) { $tag += '_' + ($Marker -replace '[^a-zA-Z0-9_-]', '_') }

        # 1. State snapshot
        try {
            $State | Export-Clixml -Path (Join-Path $debugDir "$tag`_State.xml") -Force
        } catch { Write-MyVerbose ("DebugSnapshot: state export failed: {0}" -f $_) }

        # 2. System info dump
        $sysFile = Join-Path $debugDir "$tag`_System.txt"
        try {
            $utf8   = [System.Text.UTF8Encoding]::new($false)
            $lines  = [System.Collections.Generic.List[string]]::new()

            $lines.Add('=== EXpress Debug Snapshot ===')
            $lines.Add(('Phase: {0} | Marker: {1} | Time: {2}' -f $phase, $Marker, (Get-Date -Format u)))
            $lines.Add(('Host: {0} | User: {1}\{2} | PID: {3}' -f $env:COMPUTERNAME, $env:USERDOMAIN, $env:USERNAME, $PID))
            $lines.Add('')

            # PowerShell
            $lines.Add('--- PowerShell ---')
            $PSVersionTable.GetEnumerator() | ForEach-Object { $lines.Add(('  {0} = {1}' -f $_.Key, $_.Value)) }
            $lines.Add('')

            # OS
            $lines.Add('--- Operating System ---')
            $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
            if ($os) {
                $lines.Add(('  Caption  : {0}' -f $os.Caption))
                $lines.Add(('  Version  : {0} | Build: {1}' -f $os.Version, $os.BuildNumber))
                $freeGB  = [math]::Round($os.FreePhysicalMemory / 1MB, 1)
                $totalGB = [math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
                $lines.Add(('  RAM      : {0} GB free of {1} GB' -f $freeGB, $totalGB))
            }
            $lines.Add('')

            # .NET
            $lines.Add('--- .NET Framework ---')
            $netReg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue
            if ($netReg) { $lines.Add(('  Release: {0} | Version: {1}' -f $netReg.Release, $netReg.Version)) }
            $lines.Add('')

            # Disk
            $lines.Add('--- Disk Free Space ---')
            Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | ForEach-Object {
                $total = if (($_.Used + $_.Free) -gt 0) { '{0:F1} GB' -f (($_.Used + $_.Free) / 1GB) } else { '?' }
                $lines.Add(('  {0}: {1:F1} GB free / {2}' -f $_.Name, ($_.Free / 1GB), $total))
            }
            $lines.Add('')

            # Pending Reboot
            $lines.Add('--- Pending Reboot Flags ---')
            $cbsPend  = Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending' -ErrorAction SilentlyContinue
            $wuPend   = Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired' -ErrorAction SilentlyContinue
            $pfrVal   = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue).PendingFileRenameOperations
            $lines.Add(('  CBS: {0} | WindowsUpdate: {1} | PendingFileRename: {2}' -f $cbsPend, $wuPend, [bool]$pfrVal))
            $lines.Add('')

            # Exchange Registry
            $lines.Add('--- Exchange Setup Registry ---')
            $exReg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup' -ErrorAction SilentlyContinue
            if ($exReg) {
                $exReg.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' } |
                    ForEach-Object { $lines.Add(('  {0} = {1}' -f $_.Name, $_.Value)) }
            } else { $lines.Add('  (key not found — Exchange not yet installed)') }
            $lines.Add('')

            # Windows Features (Exchange prereqs)
            $lines.Add('--- Windows Features (Exchange prereqs) ---')
            $featNames = @('Web-Server','Web-Mgmt-Console','Web-Http-Redirect','Web-Dyn-Compression',
                           'Web-Filtering','Web-Windows-Auth','Web-Asp-Net45','Web-Isapi-Filter',
                           'Web-Isapi-Ext','NET-Framework-45-Core','NET-WCF-HTTP-Activation45',
                           'NET-WCF-TCP-Activation45','RSAT-ADDS','NET-Framework-Features',
                           'Windows-Identity-Foundation','Server-Media-Foundation')
            try {
                Get-WindowsFeature $featNames -ErrorAction SilentlyContinue |
                    ForEach-Object { $lines.Add(('  {0}: {1}' -f $_.Name, $_.InstallState)) }
            } catch { $lines.Add(('  (Get-WindowsFeature failed: {0})' -f $_)) }
            $lines.Add('')

            # MSExchange Services
            $lines.Add('--- MSExchange Services ---')
            Get-Service 'MSEX*' -ErrorAction SilentlyContinue | Sort-Object Name |
                ForEach-Object { $lines.Add(('  {0}: {1} [{2}]' -f $_.Name, $_.Status, $_.StartType)) }
            $lines.Add('')

            # Installed packages (Exchange / prereqs)
            $lines.Add('--- Installed Packages (Exchange / prereqs) ---')
            @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
              'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*') | ForEach-Object {
                Get-ItemProperty $_ -ErrorAction SilentlyContinue |
                    Where-Object { $_.DisplayName -match 'Exchange|UCMA|Visual C\+\+|URL Rewrite|\.NET Framework' -and $_.DisplayName } |
                    ForEach-Object { $lines.Add(('  {0} {1}' -f $_.DisplayName, $_.DisplayVersion)) }
            }
            $lines.Add('')

            # Application Event Log (Exchange errors, last hour)
            $lines.Add('--- Application Event Log (Errors/Warnings, Exchange-related, last 1h) ---')
            try {
                Get-EventLog -LogName Application -EntryType Error,Warning -Newest 200 -ErrorAction Stop |
                    Where-Object { ($_.TimeGenerated -gt (Get-Date).AddHours(-1)) -and
                                   ($_.Source -match 'MSExchange|Exchange|MsiInstaller|Windows Installer') } |
                    ForEach-Object {
                        $msg = $_.Message -replace '\s+', ' '
                        $lines.Add(('  {0} [{1}] {2}' -f $_.TimeGenerated.ToString('HH:mm:ss'), $_.Source, $msg.Substring(0, [math]::Min(300, $msg.Length))))
                    }
            } catch { $lines.Add(('  (EventLog query failed: {0})' -f $_)) }
            $lines.Add('')

            # Exchange Setup log tail
            $lines.Add('--- ExchangeSetup.log (last 100 lines) ---')
            $exSetupLog = 'C:\ExchangeSetupLogs\ExchangeSetup.log'
            if (Test-Path $exSetupLog) {
                Get-Content $exSetupLog -Tail 100 | ForEach-Object { $lines.Add($_) }
            } else { $lines.Add('  (not found — Phase 4 not yet run)') }
            $lines.Add('')

            # IIS Application Pools (if IIS installed)
            $lines.Add('--- IIS Application Pools (Exchange) ---')
            try {
                Import-Module WebAdministration -ErrorAction Stop
                Get-WebConfiguration 'system.applicationHost/applicationPools/add' -ErrorAction Stop |
                    Where-Object { $_.name -match 'Exchange|MSExchange' } |
                    ForEach-Object { $lines.Add(('  {0}: {1} (identity: {2})' -f $_.name, $_.state, $_.processModel.userName)) }
            } catch { $lines.Add('  (IIS not available or no Exchange pools found)') }
            $lines.Add('')

            # EXpress State (key fields)
            $lines.Add('--- EXpress State (key fields) ---')
            $stateKeys = @('InstallPhase','LastSuccessfulPhase','InstallMailbox','InstallEdge','Autopilot',
                           'OrganizationName','SetupVersion','ExSetupVersion','RebootRequired',
                           'ExistingOrg','NewExchangeOrg','Install481','VCRedist2012','VCRedist2013',
                           'Namespace','DAGName','CopyServerConfig','LogVerbose','LogDebug',
                           'TranscriptFile','InstallPath','ReportsPath')
            foreach ($k in $stateKeys) { $lines.Add(('  {0} = {1}' -f $k, $State[$k])) }

            [System.IO.File]::WriteAllLines($sysFile, $lines, $utf8)
            Write-MyVerbose (('DebugSnapshot written: {0}' -f $debugDir))
        } catch {
            Write-MyVerbose ('DebugSnapshot: system info collection failed: {0}' -f $_)
        }
    }

    function Write-PhaseProgress {
        # Lightweight wrapper: Write-Progress for phase-level and step-level feedback.
        # Id 0 = overall install progress (Phase X of 6)
        # Id 1 = current-phase step progress (used in Phase 5 only)
        # PS2Exe does not render Write-Progress visually — fall back to Write-MyOutput for
        # meaningful milestones so progress is still visible in the console window.
        param(
            [int]$Id = 0,
            [string]$Activity,
            [string]$Status,
            [int]$PercentComplete = -1,
            [switch]$Completed
        )
        if ($Completed) {
            Write-Progress -Id $Id -Activity $Activity -Completed
        }
        elseif ($PercentComplete -ge 0) {
            Write-Progress -Id $Id -Activity $Activity -Status $Status -PercentComplete $PercentComplete
        }
        else {
            Write-Progress -Id $Id -Activity $Activity -Status $Status
        }

        # PS2Exe fallback: emit status as plain output so progress is not lost
        if ($IsPS2Exe -and -not $Completed -and $Status) {
            if ($Id -eq 0) {
                # Phase-level: only log when status changes (major transitions)
                Write-MyOutput ('[{0}] {1}' -f $Activity, $Status)
            }
            elseif ($Id -eq 1) {
                # Step-level (Phase 5): log every step
                Write-MyOutput ('  -> {0}' -f $Status)
            }
        }
    }

    # Native-exe invoker that preserves stdout+stderr for the log. In normal mode output
    # is discarded (same as the old `$null = … 2>$null` pattern). With -Debug, the merged
    # output is written to the main log tagged [EXE] so nothing is hidden.
