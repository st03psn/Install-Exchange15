#Requires -Version 5.1
<#
    .SYNOPSIS
    Build.ps1 - Compile Install-Exchange15.ps1 into a standalone .exe via PS2Exe

    Maintainer: st03ps

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    .DESCRIPTION
    Uses the PS2Exe module to compile Install-Exchange15.ps1 into a self-contained
    Windows executable that requires no separate PowerShell installation command.
    The compiled .exe:
      - Requests elevation via a UAC manifest (-requireAdmin)
      - Carries the same version string as the source script
      - Supports all original parameters (PS2Exe preserves the param() block)
      - Writes RunOnce entries pointing to the .exe (handled in Enable-RunOnce)

    .PARAMETER OutputPath
    Directory where the compiled .exe will be placed. Defaults to the script directory.

    .PARAMETER IconPath
    Optional path to a .ico file to embed in the executable.

    .PARAMETER SkipModuleInstall
    Skip automatic installation of PS2Exe if it is not already present.

    .EXAMPLE
    .\Build.ps1

    .EXAMPLE
    .\Build.ps1 -OutputPath C:\Dist -IconPath .\exchange.ico
#>
[CmdletBinding()]
param(
    [string]$OutputPath = $PSScriptRoot,
    [string]$IconPath   = '',
    [switch]$SkipModuleInstall
)

$ErrorActionPreference = 'Stop'

$sourceScript = Join-Path $PSScriptRoot 'Install-Exchange15.ps1'
if (-not (Test-Path $sourceScript)) {
    Write-Error "Source script not found: $sourceScript"
    exit 1
}

# Determine version from script
$versionLine = Select-String -Path $sourceScript -Pattern "ScriptVersion\s*=\s*'([^']+)'" | Select-Object -First 1
$version = if ($versionLine) { $versionLine.Matches[0].Groups[1].Value } else { '5.1' }

$exeName   = 'Install-Exchange15.exe'
$outputExe = Join-Path $OutputPath $exeName

Write-Host "Building Install-Exchange15 v$version -> $outputExe" -ForegroundColor Cyan

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
    Title        = "Install-Exchange15"
    Product      = "Install-Exchange15"
    Description  = "Unattended Exchange Server Installation Script"
    Version      = $version
    Company      = 'st03ps'
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
