$file = 'D:\DEV\Install-Exchange15\Install-Exchange15.ps1'
$text = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)

$before = ([regex]::Matches($text, '\(\$DE \? ')).Count
Write-Host "Ternary patterns found: $before"

# Replace ($DE ? 'X' : 'Y') -> (if ($DE) { 'X' } else { 'Y' })
# Capture groups include the surrounding single-quotes
$text = [regex]::Replace(
    $text,
    '\(\$DE \? (''[^'']*'') : (''[^'']*'')\)',
    '(if ($DE) { $1 } else { $2 })'
)

$after = ([regex]::Matches($text, '\(\$DE \? ')).Count
Write-Host "Remaining simple ternary patterns: $after"

[System.IO.File]::WriteAllText($file, $text, (New-Object System.Text.UTF8Encoding($true)))
Write-Host "Done."
