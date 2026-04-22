$f = 'D:\DEV\Install-Exchange15\Install-Exchange15.ps1'
$t = [IO.File]::ReadAllText($f, [Text.UTF8Encoding]::new($true))
$n = ([regex]::Matches($t, 'Phase (\d) of 6')).Count
$t = [regex]::Replace($t, 'Phase (\d) of 6', 'Phase $1 of 7')
[IO.File]::WriteAllText($f, $t, [Text.UTF8Encoding]::new($true))
Write-Host "Replaced $n occurrences"
