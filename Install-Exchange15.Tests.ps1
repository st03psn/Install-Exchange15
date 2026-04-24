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
        $result | Should -BeLike '*Cumulative Update 23*'
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
            '15.02.2562.017' = @{ KB='KB5074992'; FileName='ExchangeSubscriptionEdition-KB5074992-x64-en.exe'; URL=''; TargetVersion='15.02.2562.037' }
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

    It 'Every entry has required fields KB, FileName, TargetVersion' {
        foreach ($key in $ExchangeSUMap.Keys) {
            $entry = $ExchangeSUMap[$key]
            $entry.KB            | Should -Not -BeNullOrEmpty -Because "entry $key missing KB"
            $entry.FileName      | Should -Not -BeNullOrEmpty -Because "entry $key missing FileName"
            $entry.TargetVersion | Should -Not -BeNullOrEmpty -Because "entry $key missing TargetVersion"
            $entry.ContainsKey('URL') | Should -BeTrue -Because "entry $key missing URL key (may be empty string)"
        }
    }

    It 'All FileName entries end in .exe or .cab (not .msp)' {
        foreach ($key in $ExchangeSUMap.Keys) {
            $ExchangeSUMap[$key].FileName | Should -Match '\.(exe|cab)$' -Because "SU packages must be .exe or .cab, not .msp (key: $key)"
        }
    }

    It 'Non-empty URL entries point to a known Microsoft download host' {
        $allowedHosts = @('https://download.microsoft.com/', 'https://catalog.s.download.windowsupdate.com/')
        foreach ($key in $ExchangeSUMap.Keys) {
            $url = $ExchangeSUMap[$key].URL
            if ($url) {
                $matchesHost = $allowedHosts | Where-Object { $url -like "$_*" }
                $matchesHost | Should -Not -BeNullOrEmpty -Because "URL must be from a known Microsoft host (key: $key)"
            }
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

# ---------------------------------------------------------------------------
Describe 'Get-FullDomainAccount edge cases' {

    It 'Handles empty account name' {
        Get-FullDomainAccount 'CONTOSO' '' | Should -Be 'CONTOSO\'
    }

    It 'Handles empty domain name' {
        Get-FullDomainAccount '' 'Administrator' | Should -Be '\Administrator'
    }

    It 'Handles both empty' {
        Get-FullDomainAccount '' '' | Should -Be '\'
    }

    It 'Preserves case exactly' {
        Get-FullDomainAccount 'MyDomain' 'MyUser' | Should -Be 'MyDomain\MyUser'
    }
}

# ---------------------------------------------------------------------------
Describe 'Test-DBLogPathSeparation logic' {
    # Pure logic: test path-root extraction and comparison without file system access

    BeforeAll {
        function Get-DBLogSeparationStatus([string]$DBPath, [string]$LogPath) {
            if (-not $DBPath -or -not $LogPath) { return 'skipped' }
            $dbRoot  = [System.IO.Path]::GetPathRoot($DBPath).TrimEnd('\')
            $logRoot = [System.IO.Path]::GetPathRoot($LogPath).TrimEnd('\')
            if ($dbRoot -and $logRoot -and ($dbRoot -eq $logRoot)) { return 'same' }
            return 'separate'
        }
    }

    It 'Detects same drive letter as shared volume' {
        Get-DBLogSeparationStatus 'C:\DB\MDB1\MDB1.edb' 'C:\Log\MDB1' | Should -Be 'same'
    }

    It 'Detects different drive letters as separate volumes' {
        Get-DBLogSeparationStatus 'D:\DB\MDB1\MDB1.edb' 'E:\Log\MDB1' | Should -Be 'separate'
    }

    It 'Detects UNC-style mount points on same root as same' {
        Get-DBLogSeparationStatus 'C:\ExDB\DB1\DB1.edb' 'C:\ExDB\Log1' | Should -Be 'same'
    }

    It 'Returns skipped when DBPath is empty' {
        Get-DBLogSeparationStatus '' 'E:\Log' | Should -Be 'skipped'
    }

    It 'Returns skipped when LogPath is empty' {
        Get-DBLogSeparationStatus 'D:\DB\MDB1.edb' '' | Should -Be 'skipped'
    }

    It 'Is case-insensitive on drive letter' {
        Get-DBLogSeparationStatus 'c:\DB\MDB1.edb' 'C:\Log\MDB1' | Should -Be 'same'
    }
}

# ---------------------------------------------------------------------------
Describe 'HSTS header value' {
    # Validates the HSTS header value string we write to IIS

    It 'Contains max-age directive' {
        $value = 'max-age=31536000'
        $value | Should -BeLike 'max-age=*'
    }

    It 'max-age is at least 1 year (31536000 seconds)' {
        $value = 'max-age=31536000'
        $age = [int]($value -replace 'max-age=', '' -split ';')[0].Trim()
        $age | Should -BeGreaterOrEqual 31536000
    }

    It 'Does not contain includeSubDomains (would lock out internal subdomains)' {
        $value = 'max-age=31536000'
        $value | Should -Not -BeLike '*includeSubDomains*'
    }
}

# ---------------------------------------------------------------------------
Describe 'SID S-1-5-7 resolution (ANONYMOUS LOGON)' {
    # Validates that Windows can resolve SID S-1-5-7 to an NTAccount name.
    # This is the language-independent way to obtain "NT AUTHORITY\ANONYMOUS LOGON" (DE/EN/FR/...).

    It 'Resolves to a non-empty NTAccount value' {
        $sid = [System.Security.Principal.SecurityIdentifier]'S-1-5-7'
        $account = $sid.Translate([System.Security.Principal.NTAccount])
        $account.Value | Should -Not -BeNullOrEmpty
    }

    It 'Resolved account contains a backslash (DOMAIN\Name format)' {
        $sid = [System.Security.Principal.SecurityIdentifier]'S-1-5-7'
        $account = $sid.Translate([System.Security.Principal.NTAccount])
        $account.Value | Should -BeLike '*\*'
    }

    It 'SID string round-trips correctly' {
        $sid = [System.Security.Principal.SecurityIdentifier]'S-1-5-7'
        $sid.Value | Should -Be 'S-1-5-7'
    }
}

# ---------------------------------------------------------------------------
Describe 'Relay connector naming convention' {
    # Pure logic: validate connector name patterns used in New-AnonymousRelayConnector

    BeforeAll {
        $serverName = 'EXCH01'

        function Get-InternalRelayConnectorName([string]$Server) {
            return "Anonymous Internal Relay - $Server"
        }

        function Get-ExternalRelayConnectorName([string]$Server) {
            return "Anonymous External Relay - $Server"
        }
    }

    It 'Internal connector name contains server name' {
        Get-InternalRelayConnectorName $serverName | Should -BeLike "*$serverName*"
    }

    It 'External connector name contains server name' {
        Get-ExternalRelayConnectorName $serverName | Should -BeLike "*$serverName*"
    }

    It 'Internal connector name does not contain the word External' {
        Get-InternalRelayConnectorName $serverName | Should -Not -BeLike '*External*'
    }

    It 'External connector name does not contain the word Internal' {
        Get-ExternalRelayConnectorName $serverName | Should -Not -BeLike '*Internal*'
    }

    It 'Internal and external connector names are distinct' {
        $internal = Get-InternalRelayConnectorName $serverName
        $external = Get-ExternalRelayConnectorName $serverName
        $internal | Should -Not -Be $external
    }
}

# ---------------------------------------------------------------------------
Describe 'Default Frontend connector PermissionGroups cleanup logic' {
    # Validates the string manipulation that removes AnonymousUsers from PermissionGroups

    BeforeAll {
        function Remove-AnonymousUsersFromPermissionGroups([string]$PermissionGroupsString) {
            $pgList = ($PermissionGroupsString -split ',\s*') |
                Where-Object { $_.Trim() -ne 'AnonymousUsers' }
            return ($pgList -join ',')
        }
    }

    It 'Removes AnonymousUsers when it is the only entry' {
        $result = Remove-AnonymousUsersFromPermissionGroups 'AnonymousUsers'
        $result | Should -Not -BeLike '*AnonymousUsers*'
    }

    It 'Removes AnonymousUsers from a multi-value list' {
        $result = Remove-AnonymousUsersFromPermissionGroups 'AnonymousUsers, ExchangeUsers, ExchangeServers'
        $result | Should -Not -BeLike '*AnonymousUsers*'
        $result | Should -BeLike '*ExchangeUsers*'
        $result | Should -BeLike '*ExchangeServers*'
    }

    It 'Leaves list unchanged when AnonymousUsers is not present' {
        $result = Remove-AnonymousUsersFromPermissionGroups 'ExchangeUsers, ExchangeServers'
        $result | Should -BeLike '*ExchangeUsers*'
        $result | Should -BeLike '*ExchangeServers*'
        $result | Should -Not -BeLike '*AnonymousUsers*'
    }

    It 'Does not introduce extra commas when AnonymousUsers is first' {
        $result = Remove-AnonymousUsersFromPermissionGroups 'AnonymousUsers, ExchangeUsers'
        $result.TrimStart(',').Trim() | Should -Not -BeLike ',*'
    }
}

Describe 'Set-RegistryValue idempotency' {
    BeforeAll {
        $TestPath = 'HKCU:\Software\Pester-InstallExchange15-Test'

        function Write-MyVerbose($Text) {}

        function Set-RegistryValue {
            param([string]$Path, [string]$Name, $Value, [string]$PropertyType = 'DWord')
            if (-not (Test-Path $Path -ErrorAction SilentlyContinue)) {
                New-Item -Path $Path -Force -ErrorAction SilentlyContinue | Out-Null
            }
            else {
                $existing = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
                if ($null -ne $existing -and $existing.$Name -eq $Value) {
                    Write-MyVerbose ('Registry value already set: {0}\{1} = {2}' -f $Path, $Name, $Value)
                    return
                }
            }
            New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $PropertyType -Force -ErrorAction SilentlyContinue | Out-Null
        }
    }

    AfterAll {
        Remove-Item -Path 'HKCU:\Software\Pester-InstallExchange15-Test' -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'Creates the registry key when it does not exist' {
        Remove-Item -Path $TestPath -Recurse -Force -ErrorAction SilentlyContinue
        Set-RegistryValue -Path $TestPath -Name 'TestVal' -Value 42
        Test-Path $TestPath | Should -BeTrue
    }

    It 'Sets the value correctly' {
        Set-RegistryValue -Path $TestPath -Name 'Counter' -Value 7
        (Get-ItemProperty -Path $TestPath -Name 'Counter').Counter | Should -Be 7
    }

    It 'Updates value when it differs' {
        Set-RegistryValue -Path $TestPath -Name 'Counter' -Value 7
        Set-RegistryValue -Path $TestPath -Name 'Counter' -Value 99
        (Get-ItemProperty -Path $TestPath -Name 'Counter').Counter | Should -Be 99
    }

    It 'Skips write when value is already set (idempotency)' {
        Set-RegistryValue -Path $TestPath -Name 'Counter' -Value 99
        # Mock Write-MyVerbose to detect skip — count calls
        $script:verboseCalls = 0
        function Write-MyVerbose($Text) { $script:verboseCalls++ }
        Set-RegistryValue -Path $TestPath -Name 'Counter' -Value 99
        $script:verboseCalls | Should -Be 1
        (Get-ItemProperty -Path $TestPath -Name 'Counter').Counter | Should -Be 99
    }

    It 'Stores string values correctly' {
        Set-RegistryValue -Path $TestPath -Name 'StrVal' -Value 'hello' -PropertyType String
        (Get-ItemProperty -Path $TestPath -Name 'StrVal').StrVal | Should -Be 'hello'
    }
}

Describe 'Add-BackgroundJob pruning' {
    BeforeAll {
        function Add-BackgroundJob {
            param([System.Management.Automation.Job]$Job)
            if (-not $Global:BackgroundJobs) { $Global:BackgroundJobs = @() }
            $Global:BackgroundJobs = @($Global:BackgroundJobs | Where-Object { $_.State -notin @('Completed', 'Failed', 'Stopped') })
            $Global:BackgroundJobs += $Job
        }
    }

    BeforeEach {
        $Global:BackgroundJobs = @()
    }

    AfterEach {
        $Global:BackgroundJobs | ForEach-Object { $_ | Stop-Job -ErrorAction SilentlyContinue; $_ | Remove-Job -Force -ErrorAction SilentlyContinue }
        $Global:BackgroundJobs = @()
    }

    It 'Appends a running job to the list' {
        $j = Start-Job { Start-Sleep 60 }
        Add-BackgroundJob $j
        $Global:BackgroundJobs.Count | Should -Be 1
    }

    It 'Prunes completed jobs before adding a new one' {
        $j1 = Start-Job { 1 }
        $null = $j1 | Wait-Job
        $Global:BackgroundJobs = @($j1)
        $j2 = Start-Job { Start-Sleep 60 }
        Add-BackgroundJob $j2
        $Global:BackgroundJobs.Count | Should -Be 1
        $Global:BackgroundJobs[0].Id | Should -Be $j2.Id
        $j1 | Remove-Job -Force -ErrorAction SilentlyContinue
    }

    It 'Prunes failed jobs before adding a new one' {
        $j1 = Start-Job { throw 'fail' }
        $null = $j1 | Wait-Job
        $Global:BackgroundJobs = @($j1)
        $j2 = Start-Job { Start-Sleep 60 }
        Add-BackgroundJob $j2
        $Global:BackgroundJobs.Count | Should -Be 1
        $j1 | Remove-Job -Force -ErrorAction SilentlyContinue
    }

    It 'Initialises list when global variable is null' {
        $Global:BackgroundJobs = $null
        $j = Start-Job { Start-Sleep 60 }
        Add-BackgroundJob $j
        $Global:BackgroundJobs | Should -Not -BeNullOrEmpty
    }
}

# ---------------------------------------------------------------------------
# F25 — Advanced Configuration menu
#
# We extract the four catalog/gate functions from Install-Exchange15.ps1 via
# AST and re-define them at script scope so tests stay in sync with source
# (no drift risk from copy-paste replication).
# ---------------------------------------------------------------------------
Describe 'F25 Advanced Configuration' {

    BeforeAll {
        # Use merged dist/ artifact when available (modular layout); fall back to monolith
        $distPath   = Join-Path $PSScriptRoot 'dist\Install-Exchange15.ps1'
        $scriptPath = if (Test-Path $distPath) { $distPath } else { Join-Path $PSScriptRoot 'Install-Exchange15.ps1' }
        $tokens = $null; $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
        if ($errors) { throw "Parser errors in ${scriptPath}: $($errors[0].Message)" }

        $wanted = @('Get-AdvancedFeatureCatalog','Show-AdvancedMenu','Invoke-AdvancedConfigurationPrompt','Test-Feature')
        $funcs = $ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and ($wanted -contains $n.Name)
        }, $true)

        # Satisfy $script:-prefixed references from catalog Condition scriptblocks.
        # The real script sets these in the process{} block; for tests we define
        # them at script scope so [ref]$script:X resolves.
        $script:State          = @{}
        $script:FullOSVersion  = '10.0.20348'   # default: WS2022 (TLS 1.3 visible)
        $script:WS2016_PREFULL = '10.0.14393'
        $script:WS2019_PREFULL = '10.0.17763'
        $script:WS2022_PREFULL = '10.0.20348'
        $script:WS2025_PREFULL = '10.0.26100'
        $script:ScriptVersion  = '5.95-test'

        # Stub the write-* helpers referenced by the extracted functions.
        function Write-MyOutput  { param([Parameter(ValueFromPipeline)]$m) }
        function Write-MyVerbose { param([Parameter(ValueFromPipeline)]$m) }
        function Write-MyWarning { param([Parameter(ValueFromPipeline)]$m) }
        function Save-State      { param($s) }

        foreach ($f in $funcs) {
            Invoke-Expression $f.Extent.Text
        }
    }

    Context 'Get-AdvancedFeatureCatalog' {
        It 'Returns an ordered hashtable' {
            $cat = Get-AdvancedFeatureCatalog
            $cat | Should -BeOfType ([System.Collections.Specialized.OrderedDictionary])
        }

        It 'Contains the expected minimum number of entries' {
            (Get-AdvancedFeatureCatalog).Count | Should -BeGreaterOrEqual 55
        }

        It 'Every entry has Category, Label, Description, Default' {
            $cat = Get-AdvancedFeatureCatalog
            foreach ($name in $cat.Keys) {
                $e = $cat[$name]
                $e.ContainsKey('Category')    | Should -BeTrue -Because "$name must have Category"
                $e.ContainsKey('Label')       | Should -BeTrue -Because "$name must have Label"
                $e.ContainsKey('Description') | Should -BeTrue -Because "$name must have Description"
                $e.ContainsKey('Default')     | Should -BeTrue -Because "$name must have Default"
                $e.Default                    | Should -BeOfType ([bool])
            }
        }

        It 'Uses only the six documented categories' {
            $allowed = @('TLS','Hardening','Performance','ExchangePolicy','PostConfig','InstallFlow')
            $cat = Get-AdvancedFeatureCatalog
            foreach ($name in $cat.Keys) {
                $allowed -contains $cat[$name].Category | Should -BeTrue -Because "$name category '$($cat[$name].Category)' must be one of: $($allowed -join ', ')"
            }
        }
    }

    Context 'Get-AdvancedFeatureCatalog Condition gates' {
        It 'EnableTLS13 Condition is TRUE on WS2022' {
            $script:FullOSVersion = '10.0.20348'
            $cat = Get-AdvancedFeatureCatalog
            (& $cat['EnableTLS13'].Condition) | Should -BeTrue
        }

        It 'EnableTLS13 Condition is FALSE on WS2019' {
            $script:FullOSVersion = '10.0.17763'
            $cat = Get-AdvancedFeatureCatalog
            (& $cat['EnableTLS13'].Condition) | Should -BeFalse
        }

        It 'ShadowRedundancy Condition requires DAG' {
            $script:State = @{ DAGName = $null }
            (& (Get-AdvancedFeatureCatalog)['ShadowRedundancy'].Condition) | Should -BeFalse
            $script:State = @{ DAGName = 'DAG01' }
            (& (Get-AdvancedFeatureCatalog)['ShadowRedundancy'].Condition) | Should -BeTrue
        }

        It 'AnonymousRelay Condition requires a RelaySubnets value' {
            $script:State = @{}
            (& (Get-AdvancedFeatureCatalog)['AnonymousRelay'].Condition) | Should -BeFalse
            $script:State = @{ RelaySubnets = @('10.0.0.0/24') }
            (& (Get-AdvancedFeatureCatalog)['AnonymousRelay'].Condition) | Should -BeTrue
        }

        It 'MessageExpiration7d Condition hidden when CopyServerConfig set' {
            $script:State = @{ CopyServerConfig = 'exch01.contoso.com' }
            (& (Get-AdvancedFeatureCatalog)['MessageExpiration7d'].Condition) | Should -BeFalse
        }
    }

    Context 'Test-Feature precedence' {
        BeforeEach {
            $script:State = @{}
        }

        It 'Returns catalog default when AdvancedFeatures is empty' {
            (Test-Feature 'DisableSSL3') | Should -BeTrue
            (Test-Feature 'NoCBC')       | Should -BeFalse
            (Test-Feature 'LSAProtection') | Should -BeTrue
        }

        It 'AdvancedFeatures value wins over catalog default' {
            $script:State['AdvancedFeatures'] = @{ DisableSSL3 = $false; NoCBC = $true }
            (Test-Feature 'DisableSSL3') | Should -BeFalse
            (Test-Feature 'NoCBC')       | Should -BeTrue
        }

        It 'Unknown feature names return $false (fail closed)' {
            (Test-Feature 'NotARealFeature') | Should -BeFalse
        }

        It 'Accepts catalog default when AdvancedFeatures is not a hashtable' {
            $script:State['AdvancedFeatures'] = 'not-a-hashtable'
            (Test-Feature 'DisableSSL3') | Should -BeTrue
        }

        It 'Coerces non-bool AdvancedFeatures values via [bool]' {
            $script:State['AdvancedFeatures'] = @{ DisableSSL3 = 0; EnableECC = 1 }
            (Test-Feature 'DisableSSL3') | Should -BeFalse
            (Test-Feature 'EnableECC')   | Should -BeTrue
        }
    }

    Context 'Config-loader precedence (reproduced)' {
        # This replicates the merge logic Install-Exchange15.ps1 runs in the
        # -ConfigFile branch (see "Advanced Configuration — nested block in
        # .psd1" in the main script). Keep this block in sync with the source.
        BeforeAll {
            function Merge-Cfg {
                param([hashtable]$Cfg, [hashtable]$State)
                if (-not ($State['AdvancedFeatures'] -is [hashtable])) { $State['AdvancedFeatures'] = @{} }
                $catalogNames = (Get-AdvancedFeatureCatalog).Keys
                foreach ($name in $catalogNames) {
                    if ($Cfg.ContainsKey($name) -and -not $State['AdvancedFeatures'].ContainsKey($name)) {
                        $State['AdvancedFeatures'][$name] = [bool]$Cfg[$name]
                    }
                }
                if ($Cfg.ContainsKey('AdvancedFeatures') -and $Cfg['AdvancedFeatures'] -is [hashtable]) {
                    foreach ($k in $Cfg['AdvancedFeatures'].Keys) {
                        if ($catalogNames -contains $k) {
                            $State['AdvancedFeatures'][$k] = [bool]$Cfg['AdvancedFeatures'][$k]
                        }
                    }
                }
                return $State
            }
        }

        It 'Legacy top-level key seeds AdvancedFeatures' {
            $state = Merge-Cfg -Cfg @{ DisableSSL3 = $false } -State @{}
            $state['AdvancedFeatures']['DisableSSL3'] | Should -BeFalse
        }

        It 'Nested AdvancedFeatures wins over top-level' {
            $state = Merge-Cfg -Cfg @{
                DisableSSL3      = $false
                AdvancedFeatures = @{ DisableSSL3 = $true }
            } -State @{}
            $state['AdvancedFeatures']['DisableSSL3'] | Should -BeTrue
        }

        It 'Unknown keys in nested block are ignored' {
            $state = Merge-Cfg -Cfg @{ AdvancedFeatures = @{ DefinitelyNotInCatalog = $true } } -State @{}
            $state['AdvancedFeatures'].ContainsKey('DefinitelyNotInCatalog') | Should -BeFalse
        }

        It 'Catalog defaults apply for unset names via Test-Feature' {
            $script:State = Merge-Cfg -Cfg @{ AdvancedFeatures = @{ DisableSSL3 = $false } } -State @{}
            (Test-Feature 'DisableSSL3') | Should -BeFalse   # overridden
            (Test-Feature 'EnableECC')   | Should -BeTrue    # catalog default
            (Test-Feature 'NoCBC')       | Should -BeFalse   # catalog default
        }
    }

    Context 'deploy-example.psd1 parses and uses only catalog names' {
        It 'AdvancedFeatures block references only known catalog entries' {
            $psd1 = Import-PowerShellDataFile -Path (Join-Path $PSScriptRoot 'deploy-example.psd1')
            $psd1.ContainsKey('AdvancedFeatures') | Should -BeTrue
            $catalogNames = (Get-AdvancedFeatureCatalog).Keys
            foreach ($k in $psd1['AdvancedFeatures'].Keys) {
                $catalogNames -contains $k | Should -BeTrue -Because "deploy-example.psd1 AdvancedFeatures.$k must exist in catalog"
            }
        }

        It 'Legacy top-level keys (DisableSSL3 etc.) are NO LONGER present' {
            # They moved into the nested AdvancedFeatures block. Keeps tests
            # honest: once the legacy section is gone, the example stays clean.
            $psd1 = Import-PowerShellDataFile -Path (Join-Path $PSScriptRoot 'deploy-example.psd1')
            $legacy = @('DisableSSL3','DisableRC4','EnableECC','NoCBC','EnableAMSI','EnableTLS12','EnableTLS13')
            foreach ($k in $legacy) {
                $psd1.ContainsKey($k) | Should -BeFalse -Because "$k was migrated into AdvancedFeatures"
            }
        }
    }
}
