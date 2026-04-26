#Requires -Version 5.1
<#
    .SYNOPSIS
    Build.ps1 - Merge src/ modules and compile EXpress.ps1 into a standalone .exe via PS2Exe

    Maintainer: st03psn

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    .DESCRIPTION
    1. Runs tools\Merge-Source.ps1 to build dist\EXpress.ps1 from src\*.ps1 (unless -SkipMerge).
    2. Uses the PS2Exe module to compile dist\EXpress.ps1 into a self-contained Windows executable.
    The compiled .exe:
      - Requests elevation via a UAC manifest (-requireAdmin)
      - Carries the same version string as the source script
      - Supports all original parameters (PS2Exe preserves the param() block)
      - Writes RunOnce entries pointing to the .exe (handled in Enable-RunOnce)

    .PARAMETER OutputPath
    Directory where the compiled .exe will be placed. Defaults to the script directory.

    .PARAMETER IconPath
    Optional path to a .ico file to embed in the executable.

    .PARAMETER SkipMerge
    Skip the Merge-Source step (use existing dist\EXpress.ps1 as-is).

    .PARAMETER SkipModuleInstall
    Skip automatic installation of PS2Exe if it is not already present.

    .EXAMPLE
    .\Build.ps1

    .EXAMPLE
    .\Build.ps1 -OutputPath C:\Dist -IconPath .\exchange.ico

    .EXAMPLE
    .\Build.ps1 -SkipMerge
#>
[CmdletBinding()]
param(
    [string]$OutputPath = $PSScriptRoot,
    [string]$IconPath   = '',
    [switch]$SkipMerge,
    [switch]$SkipModuleInstall
)

$ErrorActionPreference = 'Stop'

# ── Step 1: Merge src/ modules into dist\EXpress.ps1 ──────────────────────────
if (-not $SkipMerge) {
    $mergeTool = Join-Path $PSScriptRoot 'tools\Merge-Source.ps1'
    if (-not (Test-Path $mergeTool)) {
        Write-Error "Merge-Source.ps1 not found: $mergeTool"
        exit 1
    }
    Write-Host 'Merging src/ modules into dist\EXpress.ps1 ...' -ForegroundColor Cyan
    & $mergeTool
    Write-Host 'Merge complete.' -ForegroundColor Green
}

# ── Step 2: Compile dist\EXpress.ps1 → EXpress.exe ───────────────────────────
$sourceScript = Join-Path $PSScriptRoot 'dist\EXpress.ps1'
if (-not (Test-Path $sourceScript)) {
    Write-Error "Source script not found: $sourceScript"
    exit 1
}

# Determine version from script
$versionLine = Select-String -Path $sourceScript -Pattern "ScriptVersion\s*=\s*'([^']+)'" | Select-Object -First 1
$version = if ($versionLine) { $versionLine.Matches[0].Groups[1].Value } else { '1.0' }

$exeName   = 'EXpress.exe'
$outputExe = Join-Path $OutputPath $exeName

Write-Host "Building EXpress v$version -> $outputExe" -ForegroundColor Cyan

# Ensure PS2Exe is available
if (-not (Get-Module -ListAvailable -Name PS2Exe)) {
    if ($SkipModuleInstall) {
        Write-Error 'PS2Exe module not found. Install it with: Install-Module PS2Exe -Scope CurrentUser'
        exit 1
    }
    Write-Host 'PS2Exe module not found, installing from PSGallery...' -ForegroundColor Yellow
    try {
        Install-Module -Name PS2Exe -Scope CurrentUser -Force -AllowClobber
        Write-Host 'PS2Exe installed successfully' -ForegroundColor Green
    }
    catch {
        Write-Error ("Failed to install PS2Exe: {0}" -f $_.Exception.Message)
        exit 1
    }
}

Import-Module PS2Exe -ErrorAction Stop

# Build argument hashtable for Invoke-PS2Exe
$ps2exeArgs = @{
    InputFile    = $sourceScript
    OutputFile   = $outputExe
    RequireAdmin = $true
    NoConsole    = $false     # Keep console window — the script is interactive/transcript-based
    Title        = 'EXpress'
    Product      = 'EXpress'
    Description  = 'Unattended Exchange Server Installation and Configuration'
    Version      = $version
    Company      = 'st03psn'
    Copyright    = 'Original author: Michel de Rooij (michel@eightwone.com)'
    Verbose      = $VerbosePreference -eq 'Continue'
}

if ($IconPath -and (Test-Path $IconPath)) {
    $ps2exeArgs['IconFile'] = $IconPath
}

if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath | Out-Null
}

try {
    Invoke-PS2Exe @ps2exeArgs
    if (Test-Path $outputExe) {
        $size = (Get-Item $outputExe).Length / 1KB
        Write-Host ("Build successful: {0} ({1:N0} KB)" -f $outputExe, $size) -ForegroundColor Green
    }
    else {
        Write-Error 'Build completed but output file not found'
        exit 1
    }
}
catch {
    Write-Error ("PS2Exe compilation failed: {0}" -f $_.Exception.Message)
    exit 1
}
