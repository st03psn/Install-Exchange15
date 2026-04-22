<#
.SYNOPSIS
    Multi-layer quality check for Install-Exchange15.ps1 (or any PS script).

.DESCRIPTION
    Runs four classes of checks and produces a single grouped report:

    1) Parse errors (terminating)         — System.Management.Automation.Language.Parser::ParseFile
    2) PSScriptAnalyzer                   — standard PS linter (installs on demand if missing)
    3) Project-specific AST detectors     — bug classes we hit repeatedly:
        a) `(if ...)` as command-mode argument            → RuntimeError in PS 5.1
        b) `-f` as first arg of a .Add()/method call with
           additional comma-separated args                 → FormatException
        c) PS 7-only ternary `?:`                          → ParseError in PS 5.1
        d) `Restart-Service W3SVC,WAS` outside Phase-5
           batched-restart gate                            → EMS session churn
    4) File-level checks                  — UTF-8 BOM, line endings, trailing whitespace

    Exits with non-zero code if any Error-severity issue is found. Warnings do not
    fail the build but are listed.

.PARAMETER Path
    Path to the .ps1 file. Defaults to the repo's main script.

.PARAMETER SkipAnalyzer
    Skip PSScriptAnalyzer (useful when offline or it's too slow).

.PARAMETER IncludeInformation
    Also list Information-severity findings (PSScriptAnalyzer; off by default — too noisy).

.EXAMPLE
    .\tools\Test-ScriptQuality.ps1
    .\tools\Test-ScriptQuality.ps1 -SkipAnalyzer
#>
[CmdletBinding()]
param(
    [string] $Path = (Join-Path (Split-Path -Parent $PSScriptRoot) 'Install-Exchange15.ps1'),
    [switch] $SkipAnalyzer,
    [switch] $IncludeInformation
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path $Path)) { Write-Error "File not found: $Path"; exit 2 }

$findings = [System.Collections.Generic.List[pscustomobject]]::new()
function Add-Finding {
    param($Category, $Severity, $Line, $Rule, $Message)
    $null = $findings.Add([pscustomobject]@{
        Category = $Category; Severity = $Severity; Line = $Line; Rule = $Rule; Message = $Message
    })
}

# ────────────────────────────────────────────────────────────────────────────
# 1) Parse errors
# ────────────────────────────────────────────────────────────────────────────
Write-Host '── Parsing AST ────────────────────────────────────────────' -ForegroundColor Cyan
$parseErrors = $null
$tokens      = $null
$ast         = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors) {
    foreach ($e in $parseErrors) {
        Add-Finding 'Parse' 'Error' $e.Extent.StartLineNumber 'ParseError' $e.Message
    }
    Write-Host ('  {0} parse error(s)' -f $parseErrors.Count) -ForegroundColor Red
} else {
    Write-Host '  no parse errors' -ForegroundColor Green
}

# ────────────────────────────────────────────────────────────────────────────
# 2) PSScriptAnalyzer
# ────────────────────────────────────────────────────────────────────────────
if (-not $SkipAnalyzer) {
    Write-Host '── PSScriptAnalyzer ───────────────────────────────────────' -ForegroundColor Cyan
    if (-not (Get-Module -ListAvailable PSScriptAnalyzer)) {
        Write-Host '  PSScriptAnalyzer not installed — attempting Install-Module CurrentUser' -ForegroundColor Yellow
        try {
            Install-Module PSScriptAnalyzer -Scope CurrentUser -Force -ErrorAction Stop
        } catch {
            Write-Host ('  install failed: {0} — skipping analyzer' -f $_.Exception.Message) -ForegroundColor Yellow
            $SkipAnalyzer = $true
        }
    }
}
if (-not $SkipAnalyzer) {
    Import-Module PSScriptAnalyzer -ErrorAction SilentlyContinue
    # Exclusions: rules that are noisy / not applicable in this script's style.
    $excludeRules = @(
        'PSAvoidUsingWriteHost'                   # intentional console output in menus
        'PSUseShouldProcessForStateChangingFunctions' # many helpers are internal
        'PSUseSingularNouns'                      # internal helpers with plural names are fine
        'PSAvoidGlobalVars'                       # state hashtable pattern
        'PSReviewUnusedParameter'                 # too many false positives with splatting
    )
    $results = Invoke-ScriptAnalyzer -Path $Path -ExcludeRule $excludeRules -Severity Error,Warning,Information
    if (-not $IncludeInformation) { $results = $results | Where-Object Severity -ne 'Information' }
    foreach ($r in $results) {
        Add-Finding 'Analyzer' $r.Severity $r.Line $r.RuleName $r.Message
    }
    Write-Host ('  {0} finding(s)' -f @($results).Count) -ForegroundColor Green
}

# ────────────────────────────────────────────────────────────────────────────
# 3) Project-specific AST detectors
# ────────────────────────────────────────────────────────────────────────────
Write-Host '── Custom detectors ───────────────────────────────────────' -ForegroundColor Cyan

# 3a) Control-flow statement in `( ... )` used as a command-mode argument. Both PS 5.1 and PS 7
# throw at runtime: "The term 'if' is not recognized as a name of a cmdlet ...". Inside $(...)
# or @(...) it's fine. A `( Command-Name args )` paren (CommandAst) is also fine — only
# keyword-driven statements (if/for/foreach/while/switch/try/do) in parens fail.
$controlFlowTypes = @(
    [System.Management.Automation.Language.IfStatementAst]
    [System.Management.Automation.Language.ForStatementAst]
    [System.Management.Automation.Language.ForEachStatementAst]
    [System.Management.Automation.Language.WhileStatementAst]
    [System.Management.Automation.Language.SwitchStatementAst]
    [System.Management.Automation.Language.TryStatementAst]
    [System.Management.Automation.Language.DoWhileStatementAst]
    [System.Management.Automation.Language.DoUntilStatementAst]
)
function Test-IsControlFlowParen {
    param($Paren)
    if (-not ($Paren -is [System.Management.Automation.Language.ParenExpressionAst])) { return $false }
    $p = $Paren.Pipeline
    # Direct: paren wraps the statement directly
    foreach ($t in $controlFlowTypes) { if ($p -is $t) { return $true } }
    # Wrapped: paren wraps a PipelineAst whose elements contain control-flow (rare with parser)
    if ($p -is [System.Management.Automation.Language.PipelineAst]) {
        foreach ($el in $p.PipelineElements) {
            foreach ($t in $controlFlowTypes) { if ($el -is $t) { return $true } }
        }
    }
    return $false
}

$cmdAsts = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)
foreach ($c in $cmdAsts) {
    for ($i = 1; $i -lt $c.CommandElements.Count; $i++) {
        $el = $c.CommandElements[$i]
        if (Test-IsControlFlowParen $el) {
            Add-Finding 'Custom' 'Error' $el.Extent.StartLineNumber 'IfAsCommandArg' `
                ("Control-flow statement in `(...)` used as command arg to '{0}'. Use `$(...)` or a helper." -f $c.CommandElements[0].Extent.Text)
        }
    }
}

# 3b) `-f` as first argument of a method call when sibling args exist.
# .Add('fmt {0}{1}' -f $a, $b)   parses as   .Add(('fmt {0}{1}' -f $a), $b)
# → FormatException because -f only gets one arg but template expects two.
$methodCalls = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.InvokeMemberExpressionAst] }, $true)
foreach ($m in $methodCalls) {
    if ($m.Arguments.Count -lt 2) { continue }
    $first = $m.Arguments[0]
    if ($first -is [System.Management.Automation.Language.BinaryExpressionAst] -and
        $first.Operator -eq 'Format') {
        Add-Finding 'Custom' 'Error' $m.Extent.StartLineNumber 'FormatInMethodCall' `
            ("'-f' as first method arg with sibling comma args — parser binds comma to method, '-f' sees only one RHS. Wrap in extra parens: ({0}) or use `-f @(...)`." -f $first.Extent.Text.Substring(0, [Math]::Min(60, $first.Extent.Text.Length)))
    }
}

# 3c) PS 7-only ternary operator `?:`
$ternaryAstType = [Type]::GetType('System.Management.Automation.Language.TernaryExpressionAst, System.Management.Automation')
if ($ternaryAstType) {
    $ternaries = $ast.FindAll({ $args[0].GetType() -eq $ternaryAstType }, $true)
    foreach ($t in $ternaries) {
        Add-Finding 'Custom' 'Error' $t.Extent.StartLineNumber 'PS7TernaryInPS51Script' `
            "PS 7+ ternary '?:' — script #Requires -Version 5.1. Use 'if (...) {...} else {...}'."
    }
} else {
    # Running under PS 5.1 — the parser would have already produced a parse error for any ternary,
    # caught in section 1. No additional detection needed.
}

# 3d) `Restart-Service W3SVC, WAS` outside the batched-restart gate in Phase 5
$restartCalls = $ast.FindAll({
    param($n)
    $n -is [System.Management.Automation.Language.CommandAst] -and
    $n.CommandElements[0].Extent.Text -eq 'Restart-Service'
}, $true)
foreach ($r in $restartCalls) {
    $txt = $r.Extent.Text
    if ($txt -match 'W3SVC' -or $txt -match 'WAS') {
        # Check the enclosing function: if it's inside Enable-ECC / Enable-CBC / Enable-AMSI
        # the rule from this session forbids it — use the $script:p5NeedsIisRestart flag instead.
        $parent = $r.Parent
        while ($parent -and -not ($parent -is [System.Management.Automation.Language.FunctionDefinitionAst])) {
            $parent = $parent.Parent
        }
        $funcName = if ($parent) { $parent.Name } else { '<top-level>' }
        if ($funcName -in @('Enable-ECC','Enable-CBC','Enable-AMSI')) {
            Add-Finding 'Custom' 'Error' $r.Extent.StartLineNumber 'DirectIisRestart' `
                ("Direct W3SVC/WAS restart in {0} — violates the Phase-5 batched-restart contract. Set `$script:p5NeedsIisRestart = `$true` instead." -f $funcName)
        }
    }
}

# ────────────────────────────────────────────────────────────────────────────
# 4) File-level checks
# ────────────────────────────────────────────────────────────────────────────
Write-Host '── File checks ────────────────────────────────────────────' -ForegroundColor Cyan
$bytes = [System.IO.File]::ReadAllBytes($Path)
if ($bytes.Length -lt 3 -or -not ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)) {
    Add-Finding 'File' 'Warning' 0 'MissingUtf8Bom' `
        'File is not UTF-8 with BOM. PS 5.1 on a non-UTF-8 system code page will misread non-ASCII characters (em-dash, umlauts).'
}
# Trailing CRLF / mixed line endings
$raw = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
$hasLF = $raw -match "(?<!`r)`n"
$hasCRLF = $raw -match "`r`n"
if ($hasLF -and $hasCRLF) {
    Add-Finding 'File' 'Warning' 0 'MixedLineEndings' 'Mixed CRLF and LF line endings detected.'
}

# ────────────────────────────────────────────────────────────────────────────
# Report
# ────────────────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '═══ Report ════════════════════════════════════════════════' -ForegroundColor Cyan
$errCount  = @($findings | Where-Object Severity -eq 'Error').Count
$warnCount = @($findings | Where-Object Severity -eq 'Warning').Count
$infoCount = @($findings | Where-Object Severity -eq 'Information').Count
Write-Host ('  Errors:       {0}' -f $errCount)   -ForegroundColor (@{$true='Red';$false='Green'}[[bool]$errCount])
Write-Host ('  Warnings:     {0}' -f $warnCount)  -ForegroundColor (@{$true='Yellow';$false='Green'}[[bool]$warnCount])
Write-Host ('  Information:  {0}' -f $infoCount)  -ForegroundColor Gray
Write-Host ''

if ($findings.Count) {
    $findings |
        Sort-Object @{e='Severity';desc=$true}, Line |
        Format-Table -AutoSize -Wrap Category, Severity, Line, Rule, Message |
        Out-Host
}

# Exit non-zero on Error only; Warning and Information are informational.
exit ([int][bool]$errCount)
