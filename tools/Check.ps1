#Requires -Version 5.1
<#
.SYNOPSIS
    Single-command pre-commit quality gate for Install-Exchange15.ps1.

.DESCRIPTION
    Runs two check suites without requiring an Exchange environment:

      1. Test-ScriptQuality  — parse errors, PS 5.1 anti-patterns (ControlFlowInParen,
                               FormatInMethodCall, ternary), PSScriptAnalyzer (optional),
                               file encoding.
      2. Test-ScriptSanity   — logging tier filters, Write-My* wrappers, pipeline
                               pollution, menu structure.

    Exit code 0 = all clear.  Non-zero = at least one Error-severity finding.

.PARAMETER SkipAnalyzer
    Skip PSScriptAnalyzer (faster; useful offline or on first run before module install).

.EXAMPLE
    .\tools\Check.ps1
    .\tools\Check.ps1 -SkipAnalyzer
#>
[CmdletBinding()]
param(
    [switch]$SkipAnalyzer
)

$here    = $PSScriptRoot                          # ...\tools
$repo    = Split-Path -Parent $here               # repo root
$script  = Join-Path $repo 'Install-Exchange15.ps1'
$tools   = $here

$exitCode = 0

Write-Host ''
Write-Host ('═' * 60) -ForegroundColor Cyan
Write-Host '  EXpress — quality check' -ForegroundColor Cyan
Write-Host ('═' * 60) -ForegroundColor Cyan

# ── 1. Test-ScriptQuality ────────────────────────────────────────────────────
Write-Host ''
Write-Host '▶ Test-ScriptQuality' -ForegroundColor Cyan
& "$tools\Test-ScriptQuality.ps1" -Path $script -SkipAnalyzer:$SkipAnalyzer
if ($LASTEXITCODE -ne 0) { $exitCode = 1 }

# ── 2. Test-ScriptSanity ─────────────────────────────────────────────────────
Write-Host ''
Write-Host '▶ Test-ScriptSanity' -ForegroundColor Cyan
& "$tools\Test-ScriptSanity.ps1" -ScriptPath $script
if ($LASTEXITCODE -ne 0) { $exitCode = 1 }

# ── Summary ──────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host ('═' * 60) -ForegroundColor Cyan
if ($exitCode -eq 0) {
    Write-Host '  ALL CHECKS PASSED' -ForegroundColor Green
} else {
    Write-Host '  ONE OR MORE CHECKS FAILED' -ForegroundColor Red
}
Write-Host ('═' * 60) -ForegroundColor Cyan
Write-Host ''

exit $exitCode
