    function Invoke-XmlEscape { param([string]$Text) [Security.SecurityElement]::Escape([string]$Text) }

    function New-WdHeading {
        param([string]$Text, [int]$Level = 1)
        '<w:p><w:pPr><w:pStyle w:val="Heading{0}"/></w:pPr><w:r><w:t xml:space="preserve">{1}</w:t></w:r></w:p>' -f $Level, (Invoke-XmlEscape $Text)
    }
    function New-WdParagraph {
        param([string]$Text)
        if (-not $Text) { return '<w:p/>' }
        '<w:p><w:r><w:t xml:space="preserve">{0}</w:t></w:r></w:p>' -f (Invoke-XmlEscape $Text)
    }
    function New-WdPageBreak { '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' }
    function New-WdCentered {
        # Centered paragraph with configurable size (half-points) and optional bold.
        param([string]$Text, [int]$SizeHalfPt = 22, [bool]$Bold = $false, [string]$Color = '1F3864')
        $boldTag = if ($Bold) { '<w:b/>' } else { '' }
        $sb = '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:after="120"/></w:pPr><w:r><w:rPr>{0}<w:color w:val="{1}"/><w:sz w:val="{2}"/></w:rPr><w:t xml:space="preserve">{3}</w:t></w:r></w:p>' -f $boldTag, $Color, $SizeHalfPt, (Invoke-XmlEscape $Text)
        return $sb
    }
    function New-WdSpacer {
        # Vertical spacer paragraph (empty paragraph with configurable top spacing in twentieths of a point).
        param([int]$SpaceBefore = 240)
        '<w:p><w:pPr><w:spacing w:before="{0}" w:after="0"/></w:pPr></w:p>' -f $SpaceBefore
    }
    function New-WdToc {
        # Dynamic Table of Contents field. Word shows "Right-click → Update Field" or F9 after opening.
        # Levels 1-3 covers Heading1/2/3; \h = hyperlinks, \z = hide tab in web view, \u = use outline levels.
        param([string]$Title = 'Inhaltsverzeichnis')
        $titlePara = '<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t xml:space="preserve">{0}</w:t></w:r></w:p>' -f (Invoke-XmlEscape $Title)
        $tocField  = '<w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \o &quot;1-3&quot; \h \z \u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:i/><w:color w:val="808080"/></w:rPr><w:t xml:space="preserve">(Rechtsklick → Felder aktualisieren bzw. F9, um das Inhaltsverzeichnis zu aktualisieren)</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
        return $titlePara + $tocField
    }
    function New-WdBullet {
        param([string]$Text, [int]$Level = 0)
        '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="{0}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">{1}</w:t></w:r></w:p>' -f $Level, (Invoke-XmlEscape $Text)
    }
    function New-WdCode {
        param([string]$Text)
        '<w:p><w:pPr><w:pStyle w:val="Code"/><w:spacing w:after="120"/></w:pPr><w:r><w:t xml:space="preserve">{0}</w:t></w:r></w:p>' -f (Invoke-XmlEscape $Text)
    }
    function New-WdTable {
        # -Compact: shrinks runs to 8pt (half-point size = 16) for wide tables.
        # Word auto-layout distributes columns proportionally across the page width; with
        # long content in 6+ columns each cell wraps aggressively at 11pt default. 8pt
        # gives ~40% more horizontal characters per line and lifts most wrap to a single
        # break instead of cascading wraps on every column.
        param([string[]]$Headers, [object[]]$Rows, [switch]$Compact)
        $sb = [System.Text.StringBuilder]::new()
        $null = $sb.Append('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>')
        $colCount = if ($Headers) { $Headers.Count } else { 0 }
        # Font-size half-points: 22 = 11pt (default), 16 = 8pt (compact).
        $szHalfPt  = if ($Compact) { 16 } else { 22 }
        $cellRPr   = if ($Compact) { '<w:rPr><w:sz w:val="{0}"/></w:rPr>' -f $szHalfPt } else { '' }
        $headerRPr = if ($Compact) { '<w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="{0}"/></w:rPr>' -f $szHalfPt } else { '<w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>' }
        if ($Headers) {
            $null = $sb.Append('<w:tr><w:trPr><w:tblHeader/></w:trPr>')
            foreach ($h in $Headers) {
                $null = $sb.Append('<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="2F5496"/></w:tcPr>')
                $null = $sb.Append(('<w:p><w:r>{0}<w:t xml:space="preserve">{1}</w:t></w:r></w:p></w:tc>' -f $headerRPr, (Invoke-XmlEscape $h)))
            }
            $null = $sb.Append('</w:tr>')
        }
        # PS 5.1 flattens `@( @(a,b), @(c,d) )` literals to `@(a,b,c,d)` before the array is
        # bound to this parameter. Detect that case by scanning for any array-typed element; if
        # all elements are scalars and the total count is a multiple of $colCount, reshape into
        # rows of $colCount cells. Callers who pass `List[object[]].ToArray()` or use `,@(...)`
        # per row are unaffected because their rows remain array-typed.
        if ($Rows -and $colCount -gt 1 -and $Rows.Count -gt 0) {
            $anyArrayRow = $false
            foreach ($r in $Rows) {
                if ($null -ne $r -and -not ($r -is [string]) -and ($r -is [System.Collections.IEnumerable])) { $anyArrayRow = $true; break }
            }
            if (-not $anyArrayRow -and ($Rows.Count % $colCount -eq 0)) {
                $reshaped = New-Object 'System.Collections.Generic.List[object[]]'
                for ($i = 0; $i -lt $Rows.Count; $i += $colCount) {
                    $buf = New-Object 'object[]' $colCount
                    for ($j = 0; $j -lt $colCount; $j++) { $buf[$j] = $Rows[$i + $j] }
                    $reshaped.Add($buf)
                }
                $Rows = $reshaped.ToArray()
            }
        }
        foreach ($row in $Rows) {
            # Callers that forget the `,@(...)` prefix on literal jagged arrays cause PS 5.1
            # to flatten the outer @(...), so each $row arrives as a scalar string instead of
            # a row array. Normalize both cases here to avoid emitting ragged tables, which
            # some Word versions flag as invalid and refuse to render past that point.
            $cells = @($row)
            $null = $sb.Append('<w:tr>')
            foreach ($cell in $cells) {
                $cellStr = [string]$cell
                if ($cellStr -match "`n") {
                    # Multi-line cell: split at newlines, render each line as a separate run
                    # with <w:br/> between them. Font shrunk to 9pt (18 half-pt) so long paths
                    # fit on one line each.
                    $mlSz  = if ($Compact) { $szHalfPt } else { 18 }
                    $mlRPr = '<w:rPr><w:sz w:val="{0}"/></w:rPr>' -f $mlSz
                    $brRun = '<w:r>{0}<w:br/></w:r>' -f $mlRPr
                    $lines = $cellStr -split "`n"
                    $runs  = ($lines | ForEach-Object { '<w:r>{0}<w:t xml:space="preserve">{1}</w:t></w:r>' -f $mlRPr, (Invoke-XmlEscape $_) }) -join $brRun
                    $null  = $sb.Append(('<w:tc><w:p>{0}</w:p></w:tc>' -f $runs))
                }
                else {
                    $null = $sb.Append(('<w:tc><w:p><w:r>{0}<w:t xml:space="preserve">{1}</w:t></w:r></w:p></w:tc>' -f $cellRPr, (Invoke-XmlEscape $cellStr)))
                }
            }
            # Pad short rows to header width so cell counts match the first/header row.
            for ($pad = $cells.Count; $pad -lt $colCount; $pad++) {
                $null = $sb.Append(('<w:tc><w:p><w:r>{0}<w:t xml:space="preserve"></w:t></w:r></w:p></w:tc>' -f $cellRPr))
            }
            $null = $sb.Append('</w:tr>')
        }
        $null = $sb.Append('</w:tbl>')
        $sb.ToString()
    }
    function New-WdDocumentXml {
        param([string[]]$BodyParts)
        $body = $BodyParts -join "`n"
        @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
$body
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId3"/>
      <w:footerReference w:type="default" r:id="rId4"/>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1800" w:header="709" w:footer="709" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"@
    }
    function New-WdFile {
        param([string]$OutputPath, [string[]]$BodyParts, [string]$DocTitle = '', [string]$HeaderLabel = '', [string]$LogoPath = '')
        Add-Type -AssemblyName System.IO.Compression
        $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
        $hasLogo = $LogoPath -and (Test-Path $LogoPath -PathType Leaf)
        $fs  = [System.IO.File]::Open($OutputPath, [System.IO.FileMode]::Create)
        $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        function Add-ZipEntry([string]$name, [string]$content) {
            $entry  = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
            $stream = $entry.Open()
            $bytes  = $utf8NoBom.GetBytes($content)
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Dispose()
        }
        function Add-ZipBinaryEntry([string]$name, [byte[]]$bytes) {
            $entry  = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
            $stream = $entry.Open()
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Dispose()
        }
        $d    = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ssZ')
        $te   = Invoke-XmlEscape $DocTitle
        $heSrc = if ($HeaderLabel) { $HeaderLabel } else { $DocTitle }
        $he   = Invoke-XmlEscape $heSrc
        $pngCT = if ($hasLogo) { "`n  <Default Extension=`"png`" ContentType=`"image/png`"/>" } else { '' }
        Add-ZipEntry '[Content_Types].xml' @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>$pngCT
  <Override PartName="/word/document.xml"  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/header1.xml"   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml"   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/docProps/core.xml"  ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>
"@
        Add-ZipEntry '_rels/.rels' @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>
'@
        Add-ZipEntry 'docProps/core.xml' @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>$te</dc:title>
  <dc:creator>EXpress v$ScriptVersion</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">$d</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">$d</dcterms:modified>
</cp:coreProperties>
"@
        $logoRel = if ($hasLogo) { "`n  <Relationship Id=`"rId5`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`" Target=`"media/logo.png`"/>" } else { '' }
        Add-ZipEntry 'word/_rels/document.xml.rels' @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"    Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"    Target="header1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"    Target="footer1.xml"/>$logoRel
</Relationships>
"@
        Add-ZipEntry 'word/styles.xml' @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
      <w:sz w:val="22"/><w:szCs w:val="22"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:pageBreakBefore/><w:keepNext/><w:keepLines/><w:spacing w:before="480" w:after="80"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:b/><w:color w:val="2F5496"/><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:jc w:val="center"/><w:spacing w:before="240" w:after="120"/><w:contextualSpacing/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:b/><w:color w:val="1F3864"/><w:sz w:val="72"/><w:szCs w:val="72"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle">
    <w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:after="120"/><w:contextualSpacing/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:i/><w:color w:val="2E74B5"/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="TOCHeading">
    <w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/><w:next w:val="Normal"/>
    <w:pPr><w:pageBreakBefore/><w:outlineLvl w:val="9"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:b/><w:color w:val="2F5496"/><w:sz w:val="40"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="TOC1">
    <w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:spacing w:before="120" w:after="0"/></w:pPr>
    <w:rPr><w:b/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="TOC2">
    <w:name w:val="toc 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:spacing w:after="0"/><w:ind w:left="220"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="TOC3">
    <w:name w:val="toc 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:spacing w:after="0"/><w:ind w:left="440"/></w:pPr>
  </w:style>
  <w:style w:type="character" w:styleId="PlaceholderText">
    <w:name w:val="Placeholder Text"/>
    <w:rPr><w:color w:val="808080"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="40"/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:b/><w:color w:val="2E74B5"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="40"/><w:outlineLvl w:val="2"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:b/><w:color w:val="1F3864"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading4">
    <w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="20"/><w:outlineLvl w:val="3"/></w:pPr>
    <w:rPr><w:i/><w:color w:val="2E74B5"/><w:sz w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Code">
    <w:name w:val="Code"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="0" w:after="0"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Courier New"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:ind w:left="720"/></w:pPr>
  </w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal">
    <w:name w:val="Normal Table"/>
    <w:tblPr><w:tblCellMar>
      <w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
      <w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
    </w:tblCellMar></w:tblPr>
  </w:style>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/>
    <w:tblPr><w:tblBorders>
      <w:top    w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:left   w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:right  w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    </w:tblBorders></w:tblPr>
  </w:style>
</w:styles>
'@
        Add-ZipEntry 'word/numbering.xml' @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#x2022;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/><w:numFmt w:val="bullet"/>
      <w:lvlText w:val="o"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:hint="default"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>
'@
        Add-ZipEntry 'word/header1.xml' @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:jc w:val="right"/>
      <w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="2F5496"/></w:pBdr>
      <w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr>
    </w:pPr>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr>
      <w:t>$he</w:t>
    </w:r>
  </w:p>
</w:hdr>
"@
        Add-ZipEntry 'word/footer1.xml' @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pBdr><w:top w:val="single" w:sz="6" w:space="1" w:color="2F5496"/></w:pBdr>
      <w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs>
      <w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr>
    </w:pPr>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:t>INTERN</w:t></w:r>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:tab/></w:r>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve"> / </w:t></w:r>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:ftr>
'@
        if ($hasLogo) { Add-ZipBinaryEntry 'word/media/logo.png' ([System.IO.File]::ReadAllBytes($LogoPath)) }
        Add-ZipEntry 'word/document.xml' (New-WdDocumentXml $BodyParts)
        $zip.Dispose()
        $fs.Dispose()
    }

    # ── New-InstallationDocument (F22) ────────────────────────────────────────────
