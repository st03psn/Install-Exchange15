# Self-contained test harness for Write-ToTranscript.
# Extracts the function definition from Install-Exchange15.ps1 via the PS AST and runs it
# against a temp file under all three tiers. Verifies filtering + UTF-8 encoding + suppressed-error diff.

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
function Add-Result([string]$Name, [bool]$Passed, [string]$Detail = '') {
    $results.Add([pscustomobject]@{ Name = $Name; Passed = $Passed; Detail = $Detail })
}

# --- 1. Parse script and extract Write-ToTranscript ---
$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)
if ($errors) { throw "Parser errors: $($errors -join '; ')" }

$fn = $ast.FindAll({
    param($n)
    $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Write-ToTranscript'
}, $true) | Select-Object -First 1
if (-not $fn) { throw 'Write-ToTranscript not found in script.' }

# Inject it into this scope.
Invoke-Expression $fn.Extent.Text
Add-Result 'Function extracted from script' $true

# --- 2. Helper: fresh $State for each tier test ---
function New-TestState([bool]$LogVerbose, [bool]$LogDebug) {
    $tmp = [System.IO.Path]::Combine($env:TEMP, 'ExLogTest_' + [guid]::NewGuid().ToString('N') + '.log')
    @{
        TranscriptFile = $tmp
        LogVerbose     = $LogVerbose
        LogDebug       = $LogDebug
    }
}

function Read-LogText([string]$Path) {
    if (-not (Test-Path $Path)) { return '' }
    [System.IO.File]::ReadAllText($Path, [System.Text.UTF8Encoding]::new($false))
}

function Test-IsUtf8NoBomNoNulls([string]$Path) {
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    # Reject UTF-16LE (lots of null bytes in ASCII text)
    $nullCount = 0
    for ($i = 0; $i -lt $bytes.Length; $i++) { if ($bytes[$i] -eq 0) { $nullCount++ } }
    $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
    return @{ NoNulls = ($nullCount -eq 0); NoBom = (-not $hasBom); Bytes = $bytes.Length }
}

# Drive scoping for $State / $script:lastErrorCount access inside the injected function.
$script:lastErrorCount = 0

# --- 3. Tier: DEFAULT (no verbose/no debug) ---
$State = New-TestState $false $false
Write-ToTranscript 'INFO'    'info-line'
Write-ToTranscript 'WARNING' 'warn-line'
Write-ToTranscript 'ERROR'   'err-line'
Write-ToTranscript 'EXE'     'exe-line'
Write-ToTranscript 'VERBOSE' 'verbose-line-should-NOT-appear'
Write-ToTranscript 'DEBUG'   'debug-line-should-NOT-appear'
$txt = Read-LogText $State['TranscriptFile']
Add-Result 'Default tier: INFO written'     ($txt -match '\[INFO\] info-line')
Add-Result 'Default tier: WARNING written'  ($txt -match '\[WARNING\] warn-line')
Add-Result 'Default tier: ERROR written'    ($txt -match '\[ERROR\] err-line')
Add-Result 'Default tier: EXE written'      ($txt -match '\[EXE\] exe-line')
Add-Result 'Default tier: VERBOSE filtered' (-not ($txt -match 'verbose-line-should-NOT-appear'))
Add-Result 'Default tier: DEBUG filtered'   (-not ($txt -match 'debug-line-should-NOT-appear'))
$enc = Test-IsUtf8NoBomNoNulls $State['TranscriptFile']
Add-Result 'Default tier: UTF-8 (no UTF-16 nulls)' $enc.NoNulls ("bytes=$($enc.Bytes)")
Add-Result 'Default tier: No UTF-8 BOM'            $enc.NoBom
Remove-Item $State['TranscriptFile'] -Force -ErrorAction SilentlyContinue

# --- 4. Tier: VERBOSE ---
$State = New-TestState $true $false
Write-ToTranscript 'INFO'    'info-line'
Write-ToTranscript 'VERBOSE' 'verbose-line-should-APPEAR'
Write-ToTranscript 'DEBUG'   'debug-line-should-NOT-appear'
$txt = Read-LogText $State['TranscriptFile']
Add-Result 'Verbose tier: INFO written'     ($txt -match '\[INFO\] info-line')
Add-Result 'Verbose tier: VERBOSE written'  ($txt -match 'verbose-line-should-APPEAR')
Add-Result 'Verbose tier: DEBUG filtered'   (-not ($txt -match 'debug-line-should-NOT-appear'))
$enc = Test-IsUtf8NoBomNoNulls $State['TranscriptFile']
Add-Result 'Verbose tier: UTF-8 (no UTF-16 nulls)' $enc.NoNulls
Remove-Item $State['TranscriptFile'] -Force -ErrorAction SilentlyContinue

# --- 5. Tier: DEBUG + SUPPRESSED-ERROR diff ---
$State = New-TestState $true $true
$script:lastErrorCount = $Error.Count  # baseline
Write-ToTranscript 'INFO'    'info-line'
Write-ToTranscript 'VERBOSE' 'verbose-line-APPEARS'
Write-ToTranscript 'DEBUG'   'debug-line-APPEARS'
# Provoke a suppressed error so the $Error diff should be captured on the next call.
Get-ItemProperty -Path 'HKLM:\SOFTWARE\__this_path_does_not_exist__' -ErrorAction SilentlyContinue | Out-Null
Write-ToTranscript 'INFO' 'trigger-suppressed-error-flush'
$txt = Read-LogText $State['TranscriptFile']
Add-Result 'Debug tier: INFO written'    ($txt -match '\[INFO\] info-line')
Add-Result 'Debug tier: VERBOSE written' ($txt -match 'verbose-line-APPEARS')
Add-Result 'Debug tier: DEBUG written'   ($txt -match 'debug-line-APPEARS')
Add-Result 'Debug tier: SUPPRESSED-ERROR diff captured' ($txt -match '\[SUPPRESSED-ERROR\].*__this_path_does_not_exist__')
$enc = Test-IsUtf8NoBomNoNulls $State['TranscriptFile']
Add-Result 'Debug tier: UTF-8 (no UTF-16 nulls)' $enc.NoNulls

# --- 6. Umlaut round-trip (encoding sanity) ---
Write-ToTranscript 'INFO' 'Namenspräfix Grüße'
$txt = Read-LogText $State['TranscriptFile']
Add-Result 'Umlauts round-trip cleanly' ($txt -match 'Namenspräfix Grüße')
Remove-Item $State['TranscriptFile'] -Force -ErrorAction SilentlyContinue

# --- 7. Pre-menu bootstrap: verify Write-MyOutput / Write-MyVerbose also write to the log. ---
# Extract the wrappers so we can replicate the bootstrap path without launching the full script.
foreach ($fname in 'Write-MyOutput','Write-MyVerbose','Write-MyWarning') {
    $wf = $ast.FindAll({
        param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $fname
    }, $true) | Select-Object -First 1
    if ($wf) { Invoke-Expression $wf.Extent.Text }
}

$State = New-TestState $true $true
# Simulate the bootstrap ordering: $VerbosePreference/$DebugPreference get pinned off before the wrappers run.
$VerbosePreference = 'SilentlyContinue'
$DebugPreference   = 'SilentlyContinue'

# Capture console output so we can assert Write-Verbose stays silent.
$consoleOut = & {
    Write-MyOutput  'Script called using ...'
    Write-MyVerbose 'Using parameterSet Autopilot'
    Write-MyOutput  'Running on OS build 10.0.26100'
} 4>&1 6>&1 *>&1 | Out-String

$txt = Read-LogText $State['TranscriptFile']
Add-Result 'Pre-menu: Write-MyOutput reached log'  ($txt -match '\[INFO\] Script called')
Add-Result 'Pre-menu: Write-MyVerbose reached log' ($txt -match '\[VERBOSE\] Using parameterSet')
Add-Result 'Pre-menu: OS build line in log'        ($txt -match '\[INFO\] Running on OS build')
Add-Result 'Console: Verbose NOT on console (no "Using parameterSet")' `
    (-not ($consoleOut -match 'Using parameterSet'))
Remove-Item $State['TranscriptFile'] -Force -ErrorAction SilentlyContinue

# --- 7. Report ---
Write-Host ''
Write-Host '=== Test-Logging Results ===' -ForegroundColor Cyan
$pass = 0; $fail = 0
foreach ($r in $results) {
    if ($r.Passed) {
        Write-Host ('  [PASS] {0}' -f $r.Name) -ForegroundColor Green
        $pass++
    } else {
        Write-Host ('  [FAIL] {0}  {1}' -f $r.Name, $r.Detail) -ForegroundColor Red
        $fail++
    }
}
Write-Host ('---' + [Environment]::NewLine + "$pass passed, $fail failed") -ForegroundColor Cyan
if ($fail -gt 0) { exit 1 } else { exit 0 }
