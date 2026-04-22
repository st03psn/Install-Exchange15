$errs = $null
[void][System.Management.Automation.Language.Parser]::ParseFile('D:\DEV\Install-Exchange15\Install-Exchange15.ps1', [ref]$null, [ref]$errs)
if ($errs) {
    $errs | ForEach-Object { Write-Host ("{0}:{1} {2}" -f $_.Extent.StartLineNumber, $_.Extent.StartColumnNumber, $_.Message) }
    exit 1
} else {
    Write-Host 'OK'
}
