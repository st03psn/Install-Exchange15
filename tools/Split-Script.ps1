#Requires -Version 5.1
<#
.SYNOPSIS
    One-time script: splits Install-Exchange15.ps1 into src/ modules.
    Run from repo root: .\tools\Split-Script.ps1
    After running, verify with:
        .\tools\Merge-Source.ps1 -ReferenceFile .\dist\Install-Exchange15.ps1
        .\tools\Parse-Check.ps1 -Path .\dist\Install-Exchange15.ps1
#>
[CmdletBinding()]
param(
    [switch]$Force
)
$ErrorActionPreference = 'Stop'

$repoRoot  = Split-Path $PSScriptRoot
$srcScript = Join-Path $repoRoot 'Install-Exchange15.ps1'
$srcDir    = Join-Path $repoRoot 'src'

# Guard: abort if already split
if (-not $Force) {
    $peek = [System.IO.File]::ReadAllText($srcScript)
    if ($peek -match '#region SOURCE-LOADER') {
        Write-Error 'Entry script already contains SOURCE-LOADER. Already split. Use -Force to override.'
        exit 1
    }
}

$enc = [System.Text.UTF8Encoding]::new($true)   # UTF-8 with BOM

Write-Host "Reading $srcScript ..."
$lines = [System.IO.File]::ReadAllLines($srcScript, $enc)
Write-Host "  $($lines.Count) lines"

$origHash = (Get-FileHash $srcScript -Algorithm SHA256).Hash
Write-Host "  SHA256: $origHash"

# Module definitions: [Name, StartIdx0, EndIdx0] (0-based inclusive)
$modules = @(
    @( '00-Constants',        1038,  1138 ),
    @( '05-State',            1139,  1223 ),
    @( '10-Logging',          1224,  1326 ),
    @( '15-Helpers',          1327,  1856 ),
    @( '25-AD',               1857,  2208 ),
    @( '35-Exchange',         2209,  2323 ),
    @( '40-Preflight',        2324,  3097 ),
    @( '45-ServerConfig',     3098,  3443 ),
    @( '50-Connectors',       3444,  4071 ),
    @( '55-Security',         4072,  4194 ),
    @( '60-VDir-DAG',         4195,  4871 ),
    @( '70-ReportData',       4872,  5142 ),
    @( '72-ReportHtml',       5143,  5888 ),
    @( '74-OpenXml',          5889,  6251 ),
    @( '76-InstallDoc',       6252,  7633 ),
    @( '78-PostConfig',       7634,  7996 ),
    @( '85-WU-SU',            7997,  8476 ),
    @( '88-RecipientMgmt',    8477,  8635 ),
    @( '90-Hardening',        8636,  9960 ),
    @( '95-Menu',             9961, 10690 ),
    @( '99-Main',            10691, 12102 )
)

# Verify contiguous coverage 1038..12102
$exp = 1038
foreach ($m in $modules) {
    $name, $s, $e = $m
    if ($s -ne $exp) { throw "Gap before '$name': expected $exp, got $s" }
    $exp = $e + 1
}
if (($exp - 1) -ne 12102) { throw "Coverage ends at $($exp-1), expected 12102" }
Write-Host "Coverage OK: indices 1038..12102" -ForegroundColor Green

# Structural anchor checks
if ($lines[1037]  -notmatch '^\s*process\s*\{')  { throw "Expected 'process {' at index 1037, got: [$($lines[1037])]" }
if ($lines[12103] -notmatch '^\}\s*#Process')     { throw "Expected '} #Process' at index 12103, got: [$($lines[12103])]" }
Write-Host "Anchors OK (process{ @1037, } #Process @12103)" -ForegroundColor Green

if (-not (Test-Path $srcDir)) { $null = New-Item $srcDir -ItemType Directory }

# Write modules
Write-Host 'Extracting modules...'
foreach ($m in $modules) {
    $name, $s, $e = $m
    $outPath = Join-Path $srcDir "$name.ps1"
    [System.IO.File]::WriteAllLines($outPath, $lines[$s..$e], $enc)
    Write-Host "  $name.ps1  ($($e-$s+1) lines)"
}

# Write new entry script
Write-Host 'Writing entry script...'
$entry = [System.Collections.Generic.List[string]]::new()
foreach ($l in $lines[0..1037])      { $entry.Add($l) }   # header + param() + "process {"
$entry.Add('#region SOURCE-LOADER')
$entry.Add("    foreach (`$m in (Get-ChildItem (Join-Path `$PSScriptRoot 'src') -Filter '*.ps1' | Sort-Object Name)) { . `$m.FullName }")
$entry.Add('#endregion SOURCE-LOADER')
foreach ($l in $lines[12103..12104]) { $entry.Add($l) }   # "} #Process" + trailing blank
[System.IO.File]::WriteAllLines($srcScript, $entry, $enc)
Write-Host "  Entry script: $($entry.Count) lines"

# Run merge and verify
Write-Host 'Running merge + hash verification...'
& (Join-Path $PSScriptRoot 'Merge-Source.ps1') -Quiet:$false

$mergedPath = Join-Path $repoRoot 'dist\Install-Exchange15.ps1'
$mergedHash = (Get-FileHash $mergedPath -Algorithm SHA256).Hash
if ($mergedHash -eq $origHash) {
    Write-Host "BYTE-IDENTICAL: $mergedHash" -ForegroundColor Green
} else {
    Write-Host 'HASH MISMATCH' -ForegroundColor Red
    Write-Host "  Original: $origHash"
    Write-Host "  Merged:   $mergedHash"
}
