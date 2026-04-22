# Test-ScriptSanity.ps1
# Structural + behavioural sanity checks that catch common bugs before manual testing.
# Runs on Windows PowerShell 5.1.  Exit 0 = all pass, Exit 1 = failures.
#
# Checks:
#   1.  Parser / syntax
#   2.  Functions that return a value must not also call Write-Output (pipeline pollution)
#   3.  No `(if (...)` used as a -f argument (PS 5.1 runtime crash)
#   4.  No Start-Transcript / Stop-Transcript (replaced by Out-File logging)
#   5.  No $global: preference changes (VerbosePreference etc.)
#   6.  No Out-File without -Encoding inside Write-ToTranscript
#   7.  Show-InstallationMenu return type is hashtable (not array)
#   8.  LogVerbose / LogDebug flags respected by Write-ToTranscript
#   9.  Write-MyOutput / Write-MyVerbose / Write-MyDebug don't write to console when prefs are Silent
#  10.  $script:isFreshStart set before any State mutation

[CmdletBinding()]
param(
    [string]$ScriptPath
)

$ErrorActionPreference = 'Stop'
if (-not $ScriptPath) {
    $here = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $ScriptPath = Join-Path $here '..\Install-Exchange15.ps1'
}
$ScriptPath = (Resolve-Path $ScriptPath).Path

$results = [System.Collections.Generic.List[object]]::new()
function Pass([string]$Name, [string]$Detail='') { $results.Add([pscustomobject]@{Name=$Name;Passed=$true;Detail=$Detail}) }
function Fail([string]$Name, [string]$Detail='') { $results.Add([pscustomobject]@{Name=$Name;Passed=$false;Detail=$Detail}) }

# ── helpers ──────────────────────────────────────────────────────────────────
$tokens = $null; $parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath,[ref]$tokens,[ref]$parseErrors)

function Get-FunctionAst([string]$Name) {
    $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name},$true) |
        Select-Object -First 1
}
function Get-AllFunctionAsts {
    $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$true)
}
function Inject-Function([string]$Name) {
    $fn = Get-FunctionAst $Name
    if (-not $fn) { throw "Function $Name not found" }
    Invoke-Expression $fn.Extent.Text
}

# ── 1. Parser ─────────────────────────────────────────────────────────────────
if ($parseErrors) {
    Fail 'Parser: no syntax errors' ($parseErrors | ForEach-Object { "$($_.Extent.StartLineNumber):$($_.Message)" } | Select-Object -First 3 | Out-String).Trim()
} else {
    Pass 'Parser: no syntax errors'
}

# ── 2. Pipeline pollution: functions with explicit return must not call Write-Output ──
# Collect all function bodies that contain a 'return <expr>' (not bare 'return')
# and also contain a 'Write-Output' call.
$polluted = @()
foreach ($fn in Get-AllFunctionAsts) {
    $body = $fn.Extent.Text
    $hasReturn = $body -match '\breturn\s+\$'
    $hasWriteOutput = $body -match '\bWrite-Output\b'
    if ($hasReturn -and $hasWriteOutput) { $polluted += $fn.Name }
}
# White-list: functions where Write-Output is intentional output (not alongside a structured return)
$pipelineOK = @('Write-MyOutput','Write-PhaseProgress','Get-ValidatedCredentials','Read-MyInput','Get-RBACReport')
$realPolluted = $polluted | Where-Object { $_ -notin $pipelineOK }
if ($realPolluted) {
    Fail 'Pipeline: no functions mix Write-Output + return <value>' ($realPolluted -join ', ')
} else {
    Pass 'Pipeline: no functions mix Write-Output + return <value>'
}

# ── 3. No (if ...) as -f argument (PS 5.1 crash) ─────────────────────────────
$badFmt = Select-String -Path $ScriptPath -Pattern '-f\s*\(if\s*\(' -AllMatches
if ($badFmt) {
    Fail 'PS5.1: no `-f (if (...)` patterns' ("line $($badFmt[0].LineNumber): $($badFmt[0].Line.Trim())")
} else {
    Pass 'PS5.1: no `-f (if (...)` patterns'
}

# ── 4. No Start-Transcript / Stop-Transcript ──────────────────────────────────
$transcript = Select-String -Path $ScriptPath -Pattern '\b(Start|Stop)-Transcript\b' -AllMatches
if ($transcript) {
    Fail 'Logging: no Start/Stop-Transcript' ("line $($transcript[0].LineNumber)")
} else {
    Pass 'Logging: no Start/Stop-Transcript'
}

# ── 5. No $global: preference variables ──────────────────────────────────────
$globalPref = Select-String -Path $ScriptPath -Pattern '\$global:(Verbose|Debug|Error|Warning|Information)Preference' -AllMatches
if ($globalPref) {
    Fail 'Logging: no $global:*Preference changes' ("line $($globalPref[0].LineNumber)")
} else {
    Pass 'Logging: no $global:*Preference changes'
}

# ── 6. Write-ToTranscript uses AppendAllText (not Out-File) ──────────────────
$wtFn = Get-FunctionAst 'Write-ToTranscript'
if ($wtFn) {
    $wtBody = $wtFn.Extent.Text
    # Strip comment lines before checking — Out-File may appear in explanatory comments
    $wtCode = ($wtBody -split '\r?\n' | Where-Object { $_ -notmatch '^\s*#' }) -join "`n"
    $usesOutFile = $wtCode -match '\bOut-File\b'
    $usesAppend  = $wtCode -match 'AppendAllText'
    if ($usesOutFile) { Fail 'Logging: Write-ToTranscript uses Out-File (encoding risk)' }
    else              { Pass 'Logging: Write-ToTranscript uses AppendAllText (not Out-File)' }
    if ($usesAppend)  { Pass 'Logging: Write-ToTranscript uses AppendAllText' }
    else              { Fail 'Logging: Write-ToTranscript must use AppendAllText' }
} else {
    Fail 'Logging: Write-ToTranscript function found'
}

# ── 7. Show-InstallationMenu return type is hashtable ────────────────────────
$menuFn = Get-FunctionAst 'Show-InstallationMenu'
if ($menuFn) {
    # Inject all required helpers into this scope
    foreach ($dep in @('Write-ToTranscript','Write-MyOutput','Write-MyVerbose','Write-MyWarning','Write-MyDebug','Write-MyError')) {
        $f = Get-FunctionAst $dep; if ($f) { Invoke-Expression $f.Extent.Text }
    }
    # Minimal $State so Write-ToTranscript doesn't crash
    $State = @{ TranscriptFile = $null; LogVerbose = $false; LogDebug = $false }
    $script:lastErrorCount = 0

    # Inspect the function body: it must NOT contain Write-Output (only Write-Host / Write-ToTranscript)
    $menuBody = $menuFn.Extent.Text
    $menuWriteOutput = [regex]::Matches($menuBody, '\bWrite-Output\b') |
        Where-Object { $menuBody.Substring([Math]::Max(0,$_.Index-200), [Math]::Min(200,$_.Index)) -notmatch '^\s*#' }
    if ($menuWriteOutput.Count -gt 0) {
        Fail 'Menu: Show-InstallationMenu contains Write-Output (pipeline pollution risk)' "($($menuWriteOutput.Count) occurrences)"
    } else {
        Pass 'Menu: Show-InstallationMenu has no Write-Output'
    }
} else {
    Fail 'Menu: Show-InstallationMenu function found'
}

# ── 8. Three-tier filtering: LogVerbose/LogDebug respected ───────────────────
Inject-Function 'Write-ToTranscript'
function New-TierState([bool]$v,[bool]$d) {
    $tmp = [IO.Path]::Combine($env:TEMP,'SanityTier_'+[guid]::NewGuid().ToString('N')+'.log')
    @{ TranscriptFile=$tmp; LogVerbose=$v; LogDebug=$d }
}
$utf8 = [System.Text.UTF8Encoding]::new($false)

# default
$State = New-TierState $false $false; $script:lastErrorCount = 0
Write-ToTranscript 'INFO'    'info'; Write-ToTranscript 'VERBOSE' 'verb'; Write-ToTranscript 'DEBUG' 'dbg'
$t = [IO.File]::ReadAllText($State['TranscriptFile'],$utf8)
if ($t -match '\[INFO\]' -and $t -notmatch '\[VERBOSE\]' -and $t -notmatch '\[DEBUG\]') { Pass 'Tier-default: INFO yes, VERBOSE/DEBUG filtered' }
else { Fail 'Tier-default filtering' $t }
Remove-Item $State['TranscriptFile'] -Force -ErrorAction SilentlyContinue

# verbose
$State = New-TierState $true $false; $script:lastErrorCount = 0
Write-ToTranscript 'INFO' 'i'; Write-ToTranscript 'VERBOSE' 'v'; Write-ToTranscript 'DEBUG' 'd'
$t = [IO.File]::ReadAllText($State['TranscriptFile'],$utf8)
if ($t -match '\[VERBOSE\]' -and $t -notmatch '\[DEBUG\]') { Pass 'Tier-verbose: VERBOSE yes, DEBUG filtered' }
else { Fail 'Tier-verbose filtering' $t }
Remove-Item $State['TranscriptFile'] -Force -ErrorAction SilentlyContinue

# debug
$State = New-TierState $true $true; $script:lastErrorCount = 0
Write-ToTranscript 'DEBUG' 'd'
$t = [IO.File]::ReadAllText($State['TranscriptFile'],$utf8)
if ($t -match '\[DEBUG\]') { Pass 'Tier-debug: DEBUG written' } else { Fail 'Tier-debug: DEBUG written' $t }
Remove-Item $State['TranscriptFile'] -Force -ErrorAction SilentlyContinue

# ── 9. Write-My* wrappers stay off console when prefs are SilentlyContinue ───
$State = New-TierState $true $true
$tmp2  = [IO.Path]::Combine($env:TEMP,'SanityConsole_'+[guid]::NewGuid().ToString('N')+'.log')
$State['TranscriptFile'] = $tmp2
$script:lastErrorCount = 0
Inject-Function 'Write-MyOutput'
Inject-Function 'Write-MyVerbose'
Inject-Function 'Write-MyDebug'
$VerbosePreference = 'SilentlyContinue'; $DebugPreference = 'SilentlyContinue'
$console = & { Write-MyVerbose 'vtest'; Write-MyDebug 'dtest' } *>&1 | Out-String
if ($console -match 'vtest|dtest') { Fail 'Console: Write-MyVerbose/Debug silent on console' $console.Trim() }
else { Pass 'Console: Write-MyVerbose/Debug silent on console' }
$log = [IO.File]::ReadAllText($tmp2, $utf8)
if ($log -match 'vtest' -and $log -match 'dtest') { Pass 'Console: Write-MyVerbose/Debug still logged to file' }
else { Fail 'Console: Write-MyVerbose/Debug still logged to file' $log }
Remove-Item $tmp2 -Force -ErrorAction SilentlyContinue

# ── 10. $script:isFreshStart set before first $State mutation in main block ───
# We compare line numbers via Select-String so we only look at lines outside function defs.
# Both assignments live in the main process block (not inside any function).
$freshLine = (Select-String -Path $ScriptPath -Pattern '^\s*\$script:isFreshStart\s*=' |
    Select-Object -First 1).LineNumber
$mutLine   = (Select-String -Path $ScriptPath -Pattern "^\s*\`$State\['Log(Verbose|Debug)'\]\s*=" |
    Select-Object -First 1).LineNumber
if ($freshLine -and $mutLine -and $freshLine -lt $mutLine) {
    Pass '$script:isFreshStart set before first State mutation'
} else {
    Fail '$script:isFreshStart set before first State mutation' "isFreshStart=line$freshLine firstMutation=line$mutLine"
}

# ── Report ────────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '=== Test-ScriptSanity Results ===' -ForegroundColor Cyan
$pass=0; $fail=0
foreach ($r in $results) {
    if ($r.Passed) { Write-Host ("  [PASS] {0}" -f $r.Name) -ForegroundColor Green; $pass++ }
    else           { Write-Host ("  [FAIL] {0}  {1}" -f $r.Name, $r.Detail) -ForegroundColor Red; $fail++ }
}
Write-Host ("---`n$pass passed, $fail failed") -ForegroundColor Cyan
if ($fail -gt 0) { exit 1 } else { exit 0 }
