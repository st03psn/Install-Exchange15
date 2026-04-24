#Requires -Version 5.1
<#
    .SYNOPSIS
    Generates Exchange Server installation-document templates (F24 DE + EN).

    .DESCRIPTION
    Run by the maintainer to produce:
        templates/Exchange-installation-document-DE.docx
        templates/Exchange-installation-document-EN.docx

    The resulting .docx files contain a cover page with {{token}} placeholders and
    a {{document_body}} anchor paragraph.  At runtime, New-InstallationDocument calls
    Write-WdFromTemplate to inject the generated chapter XML and fill the tokens.

    Relies on the OpenXML helpers in src/74-OpenXml.ps1 and the version constant in
    src/00-Constants.ps1.  Uses pure PowerShell OOXML-ZIP — no Office/COM dependency.

    .EXAMPLE
    .\tools\Build-InstallationTemplate.ps1
#>
$ErrorActionPreference = 'Stop'
Set-StrictMode -Off

$root = Split-Path $PSScriptRoot
# Load version constant and OpenXML helpers from src/
. (Join-Path $root 'src\00-Constants.ps1')
. (Join-Path $root 'src\74-OpenXml.ps1')

$outDir = Join-Path $root 'templates'
if (-not (Test-Path $outDir)) { $null = New-Item $outDir -ItemType Directory }

foreach ($lang in @('EN', 'DE')) {
    $isDE = ($lang -eq 'DE')

    # Language helper — inline to avoid function redefinition in loop
    $tl = { param([string]$e, [string]$d) if ($isDE) { $d } else { $e } }

    $docTitle    = & $tl 'Exchange Server Installation Documentation' 'Exchange Server Installationsdokumentation'
    $headerToken = '{{HeaderLabel}}'
    $coverSub    = & $tl 'Installation, Hybrid deployment, Mail flow' 'Installation, Hybridbereitstellung, Mailflow'
    $outPath     = Join-Path $outDir ('Exchange-installation-document-' + $lang + '.docx')

    $instModeLabel = & $tl 'Installation mode' 'Installationsmodus'
    $versionLabel  = & $tl 'Version'            'Versionsnummer'
    $dateLabel     = & $tl 'Date'               'Datum'
    $authorLabel   = & $tl 'Author'             'Autor'

    # ── Cover page parts (tokens instead of dynamic values) ────────────────────
    $coverParts = [System.Collections.Generic.List[string]]::new()

    # Logo placeholder — the template ships without a real logo.
    # Customers add their own logo to word/media/ and update word/_rels/document.xml.rels.
    $null = $coverParts.Add((New-WdSpacer 1440))
    $null = $coverParts.Add((New-WdCentered -Text '[LOGO]' -SizeHalfPt 20 -Color '808080'))
    $null = $coverParts.Add((New-WdCentered -Text 'Microsoft Exchange Server SE' -SizeHalfPt 40 -Bold $true -Color '1F3864'))
    $null = $coverParts.Add('<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t xml:space="preserve">{{DocTitle}}</w:t></w:r></w:p>')
    $null = $coverParts.Add('<w:p><w:pPr><w:pStyle w:val="Subtitle"/></w:pPr><w:r><w:t xml:space="preserve">{{CoverSub}}</w:t></w:r></w:p>')
    $null = $coverParts.Add((New-WdSpacer 1200))
    $null = $coverParts.Add((New-WdCentered -Text '{{Organization}}' -SizeHalfPt 28 -Bold $true -Color '1F3864'))
    $null = $coverParts.Add((New-WdCentered -Text '{{ServerName}}' -SizeHalfPt 24 -Color '404040'))
    $null = $coverParts.Add((New-WdCentered -Text '{{Scenario}}' -SizeHalfPt 22 -Color '404040'))
    $null = $coverParts.Add((New-WdCentered -Text ($instModeLabel + ': {{InstallMode}}') -SizeHalfPt 22 -Color '404040'))
    $null = $coverParts.Add((New-WdSpacer 1440))
    $null = $coverParts.Add((New-WdCentered -Text ($versionLabel + ': {{Version}}') -SizeHalfPt 22 -Color '404040'))
    $null = $coverParts.Add((New-WdCentered -Text ($dateLabel    + ': {{DateLong}}') -SizeHalfPt 22 -Color '404040'))
    $null = $coverParts.Add((New-WdCentered -Text ($authorLabel  + ': {{Author}}') -SizeHalfPt 22 -Color '404040'))
    $null = $coverParts.Add((New-WdCentered -Text '{{Company}}' -SizeHalfPt 22 -Color '404040'))
    $null = $coverParts.Add((New-WdSpacer 600))
    $null = $coverParts.Add((New-WdCentered -Text '{{Classification}}' -SizeHalfPt 22 -Bold $true -Color 'C00000'))
    $null = $coverParts.Add((New-WdPageBreak))

    # ── document_body anchor ────────────────────────────────────────────────────
    # Write-WdFromTemplate replaces this entire paragraph with the generated chapter XML.
    $null = $coverParts.Add('<w:p><w:r><w:t>{{document_body}}</w:t></w:r></w:p>')

    # Header token — replaced in word/header1.xml by Write-WdFromTemplate at runtime.
    # Passing '{{HeaderLabel}}' lands the token in the generated header1.xml.
    New-WdFile -OutputPath $outPath -BodyParts $coverParts.ToArray() -DocTitle $docTitle -HeaderLabel $headerToken

    Write-Host ('Generated: ' + $outPath)
}

Write-Host 'Done.'
