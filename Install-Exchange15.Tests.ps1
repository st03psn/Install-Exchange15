#Requires -Version 5.1
<#
.SYNOPSIS
    Pester unit tests for Install-Exchange15.ps1 helper functions.

.NOTES
    Run with: Invoke-Pester .\Install-Exchange15.Tests.ps1 -Output Detailed
    Requires Pester 5.x: Install-Module Pester -Force -SkipPublisherCheck

    These tests exercise pure-logic helpers extracted from the main script.
    They do NOT require domain membership, admin rights, or Exchange binaries.
#>

BeforeAll {
    # ---- Reproduce the constants and helper functions under test ----
    # (These are defined inside the main script's process{} block and cannot be
    #  dot-sourced without invoking the whole script. We replicate them here.)

    $EX2016SETUPEXE_CU23 = '15.01.2507.006'
    $EX2019SETUPEXE_CU10 = '15.02.0922.007'
    $EX2019SETUPEXE_CU11 = '15.02.0986.005'
    $EX2019SETUPEXE_CU12 = '15.02.1118.007'
    $EX2019SETUPEXE_CU13 = '15.02.1258.012'
    $EX2019SETUPEXE_CU14 = '15.02.1544.004'
    $EX2019SETUPEXE_CU15 = '15.02.1748.008'
    $EXSESETUPEXE_RTM    = '15.02.2562.017'

    function Get-SetupTextVersion($FileVersion) {
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
        if ($Versions.ContainsKey($FileVersion)) {
            return '{0} (build {1})' -f $Versions[$FileVersion], $FileVersion
        }
        $res = "Unsupported version (build $FileVersion)"
        $Versions.GetEnumerator() | Sort-Object { [System.Version]$_.Key } | ForEach-Object {
            if ([System.Version]$FileVersion -ge [System.Version]$_.Key) {
                $res = '{0} (build {1})' -f $_.Value, $FileVersion
            }
        }
        return $res
    }

    function Get-OSVersionText($OSVersion) {
        $builds = @{
            '10.0.14393' = 'Windows Server 2016'
            '10.0.17763' = 'Windows Server 2019'
            '10.0.20348' = 'Windows Server 2022'
            '10.0.26100' = 'Windows Server 2025'
        }
        $text = $builds[$OSVersion]
        if (-not $text) {
            $text = ($builds.GetEnumerator() |
                Where-Object { [System.Version]$_.Key -le [System.Version]$OSVersion } |
                Sort-Object { [System.Version]$_.Key } | Select-Object -Last 1).Value
            if (-not $text) { $text = 'Windows Server (unknown)' }
        }
        return '{0} (build {1})' -f $text, $OSVersion
    }

    function Get-FullDomainAccount([string]$Domain, [string]$Account) {
        return '{0}\{1}' -f $Domain, $Account
    }
}

# ---------------------------------------------------------------------------
Describe 'Get-SetupTextVersion' {

    It 'Identifies Exchange SE RTM by exact build' {
        $result = Get-SetupTextVersion '15.02.2562.017'
        $result | Should -BeLike '*Exchange Server SE RTM*'
        $result | Should -BeLike '*15.02.2562.017*'
    }

    It 'Identifies Exchange 2019 CU14 by exact build' {
        $result = Get-SetupTextVersion '15.02.1544.004'
        $result | Should -BeLike '*CU14*'
    }

    It 'Identifies Exchange 2019 CU15 by exact build' {
        $result = Get-SetupTextVersion '15.02.1748.008'
        $result | Should -BeLike '*CU15*'
    }

    It 'Identifies Exchange 2016 CU23 by exact build' {
        $result = Get-SetupTextVersion '15.01.2507.006'
        $result | Should -BeLike '*2016*'
        $result | Should -BeLike '*CU23*'
    }

    It 'Handles SU build after CU14 (fallback path)' {
        # A post-SU build that is higher than CU14 but lower than CU15
        $result = Get-SetupTextVersion '15.02.1544.009'
        $result | Should -BeLike '*CU14*'
    }

    It 'Handles SU build for Exchange SE (fallback path)' {
        $result = Get-SetupTextVersion '15.02.2562.024'
        $result | Should -BeLike '*SE RTM*'
    }

    It 'Returns unsupported for unknown old version' {
        $result = Get-SetupTextVersion '15.00.1000.000'
        $result | Should -BeLike '*Unsupported*'
    }
}

# ---------------------------------------------------------------------------
Describe 'Get-OSVersionText' {

    It 'Identifies Windows Server 2025 exactly' {
        $result = Get-OSVersionText '10.0.26100'
        $result | Should -BeLike '*2025*'
        $result | Should -BeLike '*10.0.26100*'
    }

    It 'Identifies Windows Server 2022 exactly' {
        $result = Get-OSVersionText '10.0.20348'
        $result | Should -BeLike '*2022*'
    }

    It 'Identifies Windows Server 2019 exactly' {
        $result = Get-OSVersionText '10.0.17763'
        $result | Should -BeLike '*2019*'
    }

    It 'Identifies Windows Server 2016 exactly' {
        $result = Get-OSVersionText '10.0.14393'
        $result | Should -BeLike '*2016*'
    }

    It 'Falls back to nearest known version for patch-level build' {
        # A patched WS2025 build (higher minor)
        $result = Get-OSVersionText '10.0.26100.1234'
        $result | Should -BeLike '*2025*'
    }

    It 'Returns unknown for very old build' {
        $result = Get-OSVersionText '6.3.9600'
        $result | Should -BeLike '*unknown*'
    }
}

# ---------------------------------------------------------------------------
Describe 'Get-FullDomainAccount' {

    It 'Combines domain and account with backslash' {
        Get-FullDomainAccount 'CONTOSO' 'Administrator' | Should -Be 'CONTOSO\Administrator'
    }

    It 'Handles subdomain format' {
        Get-FullDomainAccount 'int.promiseit' 'svc_exchange' | Should -Be 'int.promiseit\svc_exchange'
    }
}

# ---------------------------------------------------------------------------
Describe 'ExchangeSUMap structure' {
    # Validates the map entries are well-formed without making network calls.

    BeforeAll {
        $ExchangeSUMap = @{
            '15.02.2562.017' = @{ KB='KB5074992'; FileName='ExchangeSE-KB5074992-x64-en.exe';      URL='https://download.microsoft.com/download/f/0/3/f03a5dab-40cd-44c4-97d4-2cee29064561/ExchangeSE-KB5074992-x64-en.exe';      TargetVersion='15.02.2562.024' }
            '15.02.1748.008' = @{ KB='KB5049233'; FileName='Exchange2019-KB5049233-x64-en.exe';    URL='https://download.microsoft.com/download/8/0/b/80b356e4-f7b1-4e11-9586-d3132a7a2fc3/Exchange2019-KB5049233-x64-en.exe';    TargetVersion='15.02.1748.016' }
            '15.02.1544.004' = @{ KB='KB5049233'; FileName='Exchange2019-KB5049233-x64-en.exe';    URL='https://download.microsoft.com/download/8/0/b/80b356e4-f7b1-4e11-9586-d3132a7a2fc3/Exchange2019-KB5049233-x64-en.exe';    TargetVersion='15.02.1544.014' }
            '15.02.1258.012' = @{ KB='KB5049233'; FileName='Exchange2019-KB5049233-x64-en.exe';    URL='https://download.microsoft.com/download/4/e/5/4e5cbbcc-5894-457d-88c4-c0b2ff7f208f/Exchange2019-KB5049233-x64-en.exe';    TargetVersion='15.02.1258.032' }
            '15.01.2507.006' = @{ KB='KB5049233'; FileName='Exchange2016-KB5049233-x64-en.exe';    URL='https://download.microsoft.com/download/0/9/9/0998c26c-8eb6-403a-b97a-ae44c4db5e20/Exchange2016-KB5049233-x64-en.exe';    TargetVersion='15.01.2507.043' }
        }
    }

    It 'Has entries for all supported CU versions' {
        $ExchangeSUMap.Keys | Should -Contain '15.02.2562.017'   # Exchange SE RTM
        $ExchangeSUMap.Keys | Should -Contain '15.02.1748.008'   # Exchange 2019 CU15
        $ExchangeSUMap.Keys | Should -Contain '15.02.1544.004'   # Exchange 2019 CU14
        $ExchangeSUMap.Keys | Should -Contain '15.02.1258.012'   # Exchange 2019 CU13
        $ExchangeSUMap.Keys | Should -Contain '15.01.2507.006'   # Exchange 2016 CU23
    }

    It 'Every entry has required fields KB, FileName, URL, TargetVersion' {
        foreach ($key in $ExchangeSUMap.Keys) {
            $entry = $ExchangeSUMap[$key]
            $entry.KB            | Should -Not -BeNullOrEmpty -Because "entry $key missing KB"
            $entry.FileName      | Should -Not -BeNullOrEmpty -Because "entry $key missing FileName"
            $entry.URL           | Should -Not -BeNullOrEmpty -Because "entry $key missing URL"
            $entry.TargetVersion | Should -Not -BeNullOrEmpty -Because "entry $key missing TargetVersion"
        }
    }

    It 'All FileName entries end in .exe (not .msp)' {
        foreach ($key in $ExchangeSUMap.Keys) {
            $ExchangeSUMap[$key].FileName | Should -BeLike '*.exe' -Because "SU packages are .exe, not .msp"
        }
    }

    It 'All URL entries point to download.microsoft.com' {
        foreach ($key in $ExchangeSUMap.Keys) {
            $ExchangeSUMap[$key].URL | Should -BeLike 'https://download.microsoft.com/*'
        }
    }

    It 'TargetVersion is higher than source version for each entry' {
        foreach ($key in $ExchangeSUMap.Keys) {
            $src    = [System.Version]$key
            $target = [System.Version]$ExchangeSUMap[$key].TargetVersion
            $target | Should -BeGreaterThan $src -Because "SU must produce a higher build than RTM for key $key"
        }
    }
}
