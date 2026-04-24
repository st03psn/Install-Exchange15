    function Save-State( $State) {
        Write-MyVerbose "Saving state information to $StateFile"
        Export-Clixml -InputObject $State -Path $StateFile
    }

    function Restore-State() {
        $State = @{}
        if (Test-Path $StateFile) {
            try {
                $State = Import-Clixml -Path $StateFile -ErrorAction Stop
                # Validate essential state properties
                if ($State -isnot [hashtable]) {
                    Write-MyWarning 'State file is corrupt (not a hashtable), starting fresh'
                    $State = @{}
                }
                else {
                    Write-Verbose "State information loaded from $StateFile"
                }
            }
            catch {
                Write-MyWarning ('Failed to load state file, starting fresh: {0}' -f $_.Exception.Message)
                $State = @{}
            }
        }
        else {
            Write-Verbose "No state file found at $StateFile"
        }
        return $State
    }


    function Get-OSVersionText( $OSVersion) {
        # Maps Windows build numbers to human-readable product names
        $builds = @{
            '10.0.14393' = 'Windows Server 2016'
            '10.0.17763' = 'Windows Server 2019'
            '10.0.20348' = 'Windows Server 2022'
            '10.0.26100' = 'Windows Server 2025'
        }
        $text = $builds[$OSVersion]
        if (-not $text) {
            # Unknown build — fall back to closest known version
            $text = ($builds.GetEnumerator() |
                Where-Object { [System.Version]$_.Key -le [System.Version]$OSVersion } |
                Sort-Object { [System.Version]$_.Key } |
                Select-Object -Last 1).Value
            if (-not $text) { $text = 'Windows Server (unknown)' }
        }
        return '{0} (build {1})' -f $text, $OSVersion
    }

    function Get-SetupTextVersion( $FileVersion) {
        $Versions = @{
            $EX2016SETUPEXE_CU23 = 'Exchange Server 2016 Cumulative Update 23'
            $EX2019SETUPEXE_CU10 = 'Exchange Server 2019 CU10'
            $EX2019SETUPEXE_CU11 = 'Exchange Server 2019 CU11'
            $EX2019SETUPEXE_CU12 = 'Exchange Server 2019 CU12'
            $EX2019SETUPEXE_CU13 = 'Exchange Server 2019 CU13'
            $EX2019SETUPEXE_CU14 = 'Exchange Server 2019 CU14'
            $EX2019SETUPEXE_CU15 = 'Exchange Server 2019 CU15'
            $EXSESETUPEXE_RTM    = 'Exchange Server SE RTM'
        }
        # Direct lookup first (exact CU build match)
        if ($Versions.ContainsKey($FileVersion)) {
            return '{0} (build {1})' -f $Versions[$FileVersion], $FileVersion
        }
        # Fallback: highest known CU version <= FileVersion (covers SU builds)
        $res = "Unsupported version (build $FileVersion)"
        $Versions.GetEnumerator() | Sort-Object -Property { [System.Version]$_.Key } | ForEach-Object {
            if ( [System.Version]$FileVersion -ge [System.Version]$_.Key) {
                $res = '{0} (build {1})' -f $_.Value, $FileVersion
            }
        }
        return $res
    }

    function Get-DetectedFileVersion( $File) {
        # Use FileVersionInfo directly — Get-Command triggers PowerShell command discovery
        # (PATH lookup, module analysis) which adds unnecessary overhead on ISO-mounted paths.
        if ( Test-Path $File) {
            return [System.Diagnostics.FileVersionInfo]::GetVersionInfo($File).ProductVersion
        }
        return 0
    }

