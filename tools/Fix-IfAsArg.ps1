# Replaces `(if ($DE) { 'X' } else { 'Y' })` with `(L 'X' 'Y')` inside New-InstallationDocument
# and `(if ($cust) { 'X' } else { 'Y' })` with `(Lc $cust 'X' 'Y')`.
# PS 5.1 parser rejects `(if ...)` as a command argument; the helper functions L/Lc
# (added near the top of New-InstallationDocument) accept ordinary expression args.

$file = 'D:\DEV\Install-Exchange15\Install-Exchange15.ps1'
$text = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)

$before = ([regex]::Matches($text, '\(if \(\$(DE|cust)\) \{')).Count
Write-Host "`(if ($DE|$cust) { ... } else { ... }) occurrences before: $before"

# Simple single-quoted string literal → L
$text = [regex]::Replace(
    $text,
    '\(if \(\$DE\) \{ (''[^'']*'') \} else \{ (''[^'']*'') \}\)',
    '(L $1 $2)'
)

# Simple single-quoted string literal → Lc for $cust
$text = [regex]::Replace(
    $text,
    '\(if \(\$cust\) \{ (''[^'']*'') \} else \{ (''[^'']*'') \}\)',
    '(Lc $cust $1 $2)'
)

$after = ([regex]::Matches($text, '\(if \(\$(DE|cust)\) \{')).Count
Write-Host "`(if ($DE|$cust) { ... } else { ... }) occurrences after:  $after"

[System.IO.File]::WriteAllText($file, $text, (New-Object System.Text.UTF8Encoding($true)))
Write-Host "Done."
