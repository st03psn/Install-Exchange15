#Requires -Version 5.1
param(
    [string]$EntryScript   = (Join-Path (Split-Path $PSScriptRoot) 'EXpress.ps1'),
    [string]$SrcDir        = (Join-Path (Split-Path $PSScriptRoot) 'src'),
    [string]$Output        = (Join-Path (Split-Path $PSScriptRoot) 'dist\EXpress.ps1'),
    [string]$ReferenceFile = '',
    [switch]$Quiet
)
$ErrorActionPreference = 'Stop'
$enc = [System.Text.UTF8Encoding]::new($true)

$entryLines = [System.IO.File]::ReadAllLines($EntryScript, $enc)
if (-not $Quiet) { Write-Host ('Entry: ' + $entryLines.Count + ' lines') }

$rStart = $null; $rEnd = $null
for ($i = 0; $i -lt $entryLines.Count; $i++) {
    if ($entryLines[$i] -match '^#region SOURCE-LOADER$')    { $rStart = $i }
    if ($entryLines[$i] -match '^#endregion SOURCE-LOADER$') { $rEnd   = $i }
}
if ($null -eq $rStart -or $null -eq $rEnd) { throw 'SOURCE-LOADER region not found' }
if (-not $Quiet) { Write-Host ('Region: lines ' + ($rStart+1) + '..' + ($rEnd+1)) }

$moduleFiles = Get-ChildItem $SrcDir -Filter '*.ps1' | Sort-Object Name
if ($moduleFiles.Count -eq 0) { throw ('No .ps1 files in ' + $SrcDir) }
$moduleLines = [System.Collections.Generic.List[string]]::new()
foreach ($f in $moduleFiles) {
    $ml = [System.IO.File]::ReadAllLines($f.FullName, $enc)
    $moduleLines.AddRange([string[]]$ml)
    if (-not $Quiet) { Write-Host ('  + ' + $f.Name + ': ' + $ml.Count + ' lines') }
}

$bc = $rStart
$as = $rEnd + 1
$ac = $entryLines.Count - $as
$merged = [string[]]::new($bc + $moduleLines.Count + $ac)
[System.Array]::Copy($entryLines, 0,   $merged, 0,                           $bc)
$moduleLines.CopyTo($merged,                                                   $bc)
[System.Array]::Copy($entryLines, $as, $merged, $bc + $moduleLines.Count,    $ac)
if (-not $Quiet) { Write-Host ('Merged: ' + $bc + ' + ' + $moduleLines.Count + ' + ' + $ac + ' = ' + $merged.Count + ' lines') }

$outDir = Split-Path $Output
if ($outDir -and -not (Test-Path $outDir)) { $null = New-Item $outDir -ItemType Directory }
[System.IO.File]::WriteAllLines($Output, $merged, $enc)
if (-not $Quiet) { Write-Host ('Output: ' + $Output) }

if ($ReferenceFile) {
    $rh = (Get-FileHash $ReferenceFile -Algorithm SHA256).Hash
    $oh = (Get-FileHash $Output        -Algorithm SHA256).Hash
    if ($rh -eq $oh) {
        Write-Host ('BYTE-IDENTICAL: ' + $oh) -ForegroundColor Green
    } else {
        Write-Host 'HASH MISMATCH' -ForegroundColor Red
        Write-Host ('  Reference: ' + $rh)
        Write-Host ('  Output:    ' + $oh)
        exit 1
    }
}
return $Output
