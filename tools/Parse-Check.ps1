#Requires -Version 5.1
<#
.SYNOPSIS
    Parses a PowerShell script and reports syntax errors.

.DESCRIPTION
    Uses the PowerShell AST parser to validate a .ps1 file.
    Exits with code 0 on success, code 1 if parse errors are found.
    Used as a CI/pre-commit guard after Merge-Source builds the release artifact.

.PARAMETER Path
    Path to the .ps1 file to check.
    Defaults to dist/Install-Exchange15.ps1.

.EXAMPLE
    .\tools\Parse-Check.ps1
    .\tools\Parse-Check.ps1 -Path .\Install-Exchange15.ps1
#>
[CmdletBinding()]
param(
    [string]$Path = (Join-Path (Split-Path $PSScriptRoot) 'dist\Install-Exchange15.ps1')
)
$ErrorActionPreference = 'Stop'

if (-not (Test-Path $Path)) {
    Write-Error "File not found: $Path"
    exit 1
}

$src    = [System.IO.File]::ReadAllText($Path)
$errors = [System.Management.Automation.Language.ParseError[]]@()
$tokens = [System.Management.Automation.Language.Token[]]@()
$null   = [System.Management.Automation.Language.Parser]::ParseInput($src, [ref]$tokens, [ref]$errors)

if ($errors.Count -eq 0) {
    Write-Host "PARSE OK  $Path" -ForegroundColor Green
    exit 0
} else {
    Write-Host "PARSE ERRORS  $Path" -ForegroundColor Red
    $errors | ForEach-Object { Write-Host "  Line $($_.Extent.StartLineNumber): $($_.Message)" -ForegroundColor Red }
    exit 1
}
