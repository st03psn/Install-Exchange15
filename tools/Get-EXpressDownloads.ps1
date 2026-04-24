#Requires -Version 5.1
<#
.SYNOPSIS
    Pre-stages all EXpress prerequisite packages and CSS-Exchange scripts into a local
    sources\ folder so that EXpress can run without internet access (air-gapped / proxy).

.DESCRIPTION
    Downloads every file that EXpress would otherwise fetch on-demand during installation:
      - VC++ 2012 and 2013 Redistributable
      - URL Rewrite Module 2.1
      - Microsoft .NET Framework 4.8 and 4.8.1
      - UCMA 4.0 Runtime
      - CSS-Exchange tools: HealthChecker, EOMT, SetupAssist, SetupLogReviewer,
        ExchangeExtendedProtection, MonitorExchangeAuthCertificate, Add-PermissionForEMT

    Files already present are skipped (idempotent — run again to refresh individual files
    by deleting them first). UCMA offline setup (UcmaRedist\Setup.exe) must come from the
    Exchange installation media and is NOT downloaded here.

.PARAMETER OutputPath
    Destination folder. Defaults to .\sources\ relative to the script.
    Create the folder if it does not exist.

.PARAMETER SkipDotNet
    Skip the large .NET 4.8 / 4.8.1 downloads (~100 MB each) — useful when the target
    machines already have the correct .NET version installed.

.EXAMPLE
    .\tools\Get-EXpressDownloads.ps1
    Downloads all files into .\sources\.

.EXAMPLE
    .\tools\Get-EXpressDownloads.ps1 -SkipDotNet
    Downloads everything except .NET installers.
#>
param(
    [string]$OutputPath = (Join-Path (Split-Path $PSScriptRoot) 'sources'),
    [switch]$SkipDotNet
)
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Get-Prerequisite {
    param([string]$Label, [string]$Url, [string]$FileName)
    $dest = Join-Path $OutputPath $FileName
    if (Test-Path $dest) {
        Write-Host ("  [SKIP] {0} (already present)" -f $FileName) -ForegroundColor DarkGray
        return
    }
    Write-Host ("  [DOWN] {0} ..." -f $Label) -ForegroundColor Cyan
    $pp = $ProgressPreference; $ProgressPreference = 'SilentlyContinue'
    try {
        # Try BITS first; fall back to WebClient for environments where BITS is restricted
        try {
            Start-BitsTransfer -Source $Url -Destination $dest -ErrorAction Stop
        }
        catch {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $wc = New-Object System.Net.WebClient
            $wc.DownloadFile($Url, $dest)
        }
        $size = (Get-Item $dest).Length
        Write-Host ("         OK ({0:N1} MB)" -f ($size / 1MB)) -ForegroundColor Green
    }
    catch {
        Write-Host ("         FAILED: {0}" -f $_.Exception.Message) -ForegroundColor Red
        Remove-Item -Path $dest -ErrorAction SilentlyContinue
    }
    finally {
        $ProgressPreference = $pp
    }
}

# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------
if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory | Out-Null
}
Write-Host "`nEXpress pre-staging downloads → $OutputPath`n" -ForegroundColor White

# ---------------------------------------------------------------------------
# Prerequisites
# ---------------------------------------------------------------------------
Write-Host "Prerequisites" -ForegroundColor Yellow

Get-Prerequisite `
    -Label    'Visual C++ 2012 Redistributable (x64)' `
    -Url      'https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe' `
    -FileName 'vcredist_x64_2012.exe'

Get-Prerequisite `
    -Label    'Visual C++ 2013 Redistributable (x64, >=12.0.40664)' `
    -Url      'https://aka.ms/highdpimfc2013x64enu' `
    -FileName 'vcredist_x64_2013.exe'

Get-Prerequisite `
    -Label    'URL Rewrite Module 2.1' `
    -Url      'https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_en-US.msi' `
    -FileName 'rewrite_amd64_en-US.msi'

Get-Prerequisite `
    -Label    'UCMA 4.0 Runtime' `
    -Url      'https://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe' `
    -FileName 'UcmaRuntimeSetup.exe'

if (-not $SkipDotNet) {
    Write-Host "`n.NET Framework (large downloads)" -ForegroundColor Yellow

    Get-Prerequisite `
        -Label    '.NET Framework 4.8' `
        -Url      'https://go.microsoft.com/fwlink/?linkid=2088631' `
        -FileName 'NDP48-x86-x64-AllOS-ENU.exe'

    Get-Prerequisite `
        -Label    '.NET Framework 4.8.1' `
        -Url      'https://download.microsoft.com/download/4/b/2/cd00d4ed-ebdd-49ee-8a33-eabc3d1030e3/NDP481-x86-x64-AllOS-ENU.exe' `
        -FileName 'NDP481-x86-x64-AllOS-ENU.exe'
}
else {
    Write-Host "`n.NET Framework" -ForegroundColor Yellow
    Write-Host "  [SKIP] -SkipDotNet specified" -ForegroundColor DarkGray
}

# ---------------------------------------------------------------------------
# CSS-Exchange tools
# ---------------------------------------------------------------------------
Write-Host "`nCSS-Exchange tools" -ForegroundColor Yellow
$cssBase = 'https://github.com/microsoft/CSS-Exchange/releases/latest/download'

foreach ($script in @(
    'HealthChecker.ps1',
    'EOMT.ps1',
    'SetupAssist.ps1',
    'SetupLogReviewer.ps1',
    'ExchangeExtendedProtectionManagement.ps1',   # renamed from ExchangeExtendedProtection.ps1 in 2024
    'MonitorExchangeAuthCertificate.ps1'
    # Add-PermissionForEMT.ps1 removed from CSS-Exchange releases; pre-stage manually if needed
)) {
    Get-Prerequisite -Label $script -Url "$cssBase/$script" -FileName $script
}

Write-Host "`nDone. Files in: $OutputPath`n" -ForegroundColor White
Write-Host "Note: UcmaRedist\Setup.exe (Server Core offline installer) must come from the" -ForegroundColor DarkYellow
Write-Host "      Exchange installation media — it is not available as a standalone download.`n" -ForegroundColor DarkYellow
