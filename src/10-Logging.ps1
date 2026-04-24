    function Write-ToTranscript( $Level, $Text) {
        # Three tiers (single log file):
        #   Default     : INFO / WARNING / ERROR / EXE
        #   -Verbose    : + VERBOSE
        #   -Debug      : + DEBUG + SUPPRESSED-ERROR diff from $Error
        # Encoding note: PS 5.1 `Out-File` defaults to Unicode (UTF-16LE w/ BOM); mixing that
        # with the UTF-8 header produces "strange font" output in viewers. We pin UTF-8 (no BOM)
        # via [IO.File]::AppendAllText so every line in the file has the same encoding.
        if (-not $State['TranscriptFile']) { return }
        $Location = Split-Path $State['TranscriptFile'] -Parent
        if (-not (Test-Path $Location)) { return }
        $verboseOn = [bool]$State['LogVerbose']
        $debugOn   = [bool]$State['LogDebug']
        $shouldWrite = switch ($Level) {
            'VERBOSE' { $verboseOn -or $debugOn }
            'DEBUG'   { $debugOn }
            default   { $true }
        }
        $utf8 = [System.Text.UTF8Encoding]::new($false)
        if ($shouldWrite) {
            try {
                [System.IO.File]::AppendAllText($State['TranscriptFile'], ("{0}: [{1}] {2}`r`n" -f (Get-Date -Format u), $Level, $Text), $utf8)
            } catch { }
        }
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
            } catch { }
        }
    }

    function Write-MyOutput( $Text) {
        Write-Output $Text
        Write-ToTranscript 'INFO' $Text
    }

    function Write-MyWarning( $Text) {
        Write-Warning $Text
        Write-ToTranscript 'WARNING' $Text
    }

    function Write-MyError( $Text) {
        Write-Error $Text
        Write-ToTranscript 'ERROR' $Text
    }

    function Write-MyVerbose( $Text) {
        Write-Verbose $Text
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
        $State['ExecutedCommands'].Add([pscustomobject]@{
            Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            Phase     = [int]($State['InstallPhase'])
            Category  = $Category
            Command   = $Command
        })
        Write-ToTranscript 'CMD' $Command
    }

    # Native-exe invoker that preserves stdout+stderr for the log. In normal mode output
    # is discarded (same as the old `$null = … 2>$null` pattern). With -Debug, the merged
    # output is written to the main log tagged [EXE] so nothing is hidden.
