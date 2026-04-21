#Requires -Version 5.1
<#
    .SYNOPSIS
    Generates Exchange Server Konzept-/Freigabedokument templates (F23 DE + EN).

    .DESCRIPTION
    Run by the maintainer to produce templates/Exchange-Konzept-Vorlage-DE.docx and -EN.docx.
    Outputs are committed to the repository and distributed with EXpress.
    Uses pure PowerShell OpenXML-ZIP — no Office/COM dependencies.

    .EXAMPLE
    .\tools\Build-KonzeptTemplate.ps1
#>
[CmdletBinding()]
param()

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

process {

# ── OpenXML Engine ─────────────────────────────────────────────────────────────

function Protect-Xml { param([string]$Text) [Security.SecurityElement]::Escape([string]$Text) }

function Get-ContentTypesXml {
@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/word/document.xml"  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/header1.xml"   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml"   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/docProps/core.xml"  ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>
'@
}

function Get-RootRelsXml {
@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>
'@
}

function Get-DocRelsXml {
@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"    Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"    Target="header1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"    Target="footer1.xml"/>
</Relationships>
'@
}

function Get-StylesXml {
@'
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
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="480" w:after="80"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:b/><w:color w:val="2F5496"/><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="40"/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:b/><w:color w:val="2E74B5"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="40"/><w:outlineLvl w:val="2"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:b/><w:color w:val="1F3864"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading4">
    <w:name w:val="heading 4"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="20"/><w:outlineLvl w:val="3"/></w:pPr>
    <w:rPr><w:i/><w:color w:val="2E74B5"/><w:sz w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Code">
    <w:name w:val="Code"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="0" w:after="0"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Courier New"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:ind w:left="720"/></w:pPr>
  </w:style>
  <w:style w:type="character" w:styleId="PlaceholderText">
    <w:name w:val="Placeholder Text"/>
    <w:rPr><w:color w:val="808080"/><w:i/></w:rPr>
  </w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal">
    <w:name w:val="Normal Table"/>
    <w:tblPr><w:tblCellMar>
      <w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
      <w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
    </w:tblCellMar></w:tblPr>
  </w:style>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
    <w:basedOn w:val="TableNormal"/>
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
}

function Get-NumberingXml {
@'
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
}

function Get-HeaderXml {
    param([string]$Title)
    $t = Protect-Xml $Title
    @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:jc w:val="right"/>
      <w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="2F5496"/></w:pBdr>
      <w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr>
    </w:pPr>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr>
      <w:t>[LOGO]&#x2003;$t</w:t>
    </w:r>
  </w:p>
</w:hdr>
"@
}

function Get-FooterXml {
    param([string]$ClassLabel = 'INTERN')
    $c = Protect-Xml $ClassLabel
    @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pBdr><w:top w:val="single" w:sz="6" w:space="1" w:color="2F5496"/></w:pBdr>
      <w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs>
      <w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr>
    </w:pPr>
    <w:r><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr><w:t>$c</w:t></w:r>
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
"@
}

function Get-CorePropsXml {
    param([string]$Title, [string]$Creator = 'EXpress')
    $d = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ssZ')
    $t = Protect-Xml $Title
    $cr = Protect-Xml $Creator
    @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>$t</dc:title>
  <dc:creator>$cr</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">$d</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">$d</dcterms:modified>
</cp:coreProperties>
"@
}

# Body fragment helpers

function Add-WordHeading {
    param([string]$Text, [int]$Level = 1)
    '<w:p><w:pPr><w:pStyle w:val="Heading{0}"/></w:pPr><w:r><w:t xml:space="preserve">{1}</w:t></w:r></w:p>' -f $Level, (Protect-Xml $Text)
}

function Add-WordParagraph {
    param([string]$Text)
    if (-not $Text) { return '<w:p/>' }
    '<w:p><w:r><w:t xml:space="preserve">{0}</w:t></w:r></w:p>' -f (Protect-Xml $Text)
}

function Add-WordBullet {
    param([string]$Text, [int]$Level = 0)
    '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="{0}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">{1}</w:t></w:r></w:p>' -f $Level, (Protect-Xml $Text)
}

function Add-WordCodeBlock {
    param([string]$Text)
    '<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr><w:r><w:t xml:space="preserve">{0}</w:t></w:r></w:p>' -f (Protect-Xml $Text)
}

function Add-WordPageBreak {
    '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
}

function Add-WordTable {
    param([string[]]$Headers, [object[][]]$Rows)
    $sb = [System.Text.StringBuilder]::new()
    $null = $sb.Append('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>')
    if ($Headers) {
        $null = $sb.Append('<w:tr><w:trPr><w:tblHeader/></w:trPr>')
        foreach ($h in $Headers) {
            $null = $sb.Append('<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="2F5496"/></w:tcPr>')
            $null = $sb.Append('<w:p><w:pPr><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr></w:pPr>')
            $null = $sb.Append('<w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t xml:space="preserve">{0}</w:t></w:r></w:p></w:tc>' -f (Protect-Xml $h))
        }
        $null = $sb.Append('</w:tr>')
    }
    foreach ($row in $Rows) {
        $null = $sb.Append('<w:tr>')
        foreach ($cell in $row) {
            $null = $sb.Append('<w:tc><w:p><w:r><w:t xml:space="preserve">{0}</w:t></w:r></w:p></w:tc>' -f (Protect-Xml ([string]$cell)))
        }
        $null = $sb.Append('</w:tr>')
    }
    $null = $sb.Append('</w:tbl>')
    $sb.ToString()
}

function Add-WordContentControl {
    param([string]$Tag, [string]$PlaceholderText = 'Eingabe...', [string]$Alias = '')
    $aliasXml = if ($Alias) { '<w:alias w:val="{0}"/>' -f (Protect-Xml $Alias) } else { '' }
    '<w:sdt><w:sdtPr>{0}<w:tag w:val="{1}"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>{2}</w:t></w:r></w:p></w:sdtContent></w:sdt>' -f $aliasXml, (Protect-Xml $Tag), (Protect-Xml $PlaceholderText)
}

# Questionnaire table: Nr. | Frage | Antwort (SDT content control)
function Add-WordQuestionnaireTable {
    param([string]$ColNr, [string]$ColFrage, [string]$ColAntwort, [object[][]]$Questions)
    $sb = [System.Text.StringBuilder]::new()
    $null = $sb.Append('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>')
    # Header row
    $null = $sb.Append('<w:tr><w:trPr><w:tblHeader/></w:trPr>')
    foreach ($h in @($ColNr, $ColFrage, $ColAntwort)) {
        $null = $sb.Append('<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="2F5496"/></w:tcPr>')
        $null = $sb.Append('<w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t xml:space="preserve">{0}</w:t></w:r></w:p></w:tc>' -f (Protect-Xml $h))
    }
    $null = $sb.Append('</w:tr>')
    # Data rows: $Questions is array of [nr, question, tag]
    foreach ($q in $Questions) {
        $nr  = Protect-Xml ([string]$q[0])
        $txt = Protect-Xml ([string]$q[1])
        $tag = [string]$q[2]
        $ph  = if ($q.Length -ge 4) { [string]$q[3] } else { 'Antwort eingeben...' }
        $null = $sb.Append('<w:tr>')
        $null = $sb.Append('<w:tc><w:tcPr><w:tcW w:w="600" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>{0}</w:t></w:r></w:p></w:tc>' -f $nr)
        $null = $sb.Append('<w:tc><w:p><w:r><w:t xml:space="preserve">{0}</w:t></w:r></w:p></w:tc>' -f $txt)
        # SDT inside table cell
        $null = $sb.Append('<w:tc><w:sdt><w:sdtPr><w:tag w:val="{0}"/><w:showingPlcHdr/><w:text/></w:sdtPr>' -f (Protect-Xml $tag))
        $null = $sb.Append('<w:sdtContent><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>{0}</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc>' -f (Protect-Xml $ph))
        $null = $sb.Append('</w:tr>')
    }
    $null = $sb.Append('</w:tbl>')
    $sb.ToString()
}

# Approval table (Chapter 16): 4 columns, 3 rows: role header / Name SDT / Datum SDT / Unterschrift label
function Add-WordApprovalTable {
    param([string[]]$Roles, [string]$LabelName = 'Name:', [string]$LabelDate = 'Datum:', [string]$LabelSig = 'Unterschrift:')
    $sb = [System.Text.StringBuilder]::new()
    $null = $sb.Append('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>')
    # Role headers
    $null = $sb.Append('<w:tr>')
    foreach ($role in $Roles) {
        $null = $sb.Append('<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="2F5496"/></w:tcPr>')
        $null = $sb.Append('<w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t xml:space="preserve">{0}</w:t></w:r></w:p></w:tc>' -f (Protect-Xml $role))
    }
    $null = $sb.Append('</w:tr>')
    # Name row
    $null = $sb.Append('<w:tr>')
    $idx = 0
    foreach ($role in $Roles) {
        $tag = 'approval_name_{0}' -f $idx
        $null = $sb.Append('<w:tc><w:p><w:r><w:t xml:space="preserve">{0} </w:t></w:r></w:p>' -f (Protect-Xml $LabelName))
        $null = $sb.Append('<w:sdt><w:sdtPr><w:tag w:val="{0}"/><w:showingPlcHdr/><w:text/></w:sdtPr>' -f $tag)
        $null = $sb.Append('<w:sdtContent><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Name eingeben...</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc>')
        $idx++
    }
    $null = $sb.Append('</w:tr>')
    # Date row
    $null = $sb.Append('<w:tr>')
    $idx = 0
    foreach ($role in $Roles) {
        $tag = 'approval_date_{0}' -f $idx
        $null = $sb.Append('<w:tc><w:p><w:r><w:t xml:space="preserve">{0} </w:t></w:r></w:p>' -f (Protect-Xml $LabelDate))
        $null = $sb.Append('<w:sdt><w:sdtPr><w:tag w:val="{0}"/><w:showingPlcHdr/><w:text/></w:sdtPr>' -f $tag)
        $null = $sb.Append('<w:sdtContent><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>TT.MM.JJJJ</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc>')
        $idx++
    }
    $null = $sb.Append('</w:tr>')
    # Signature row (tall blank row)
    $null = $sb.Append('<w:tr>')
    foreach ($role in $Roles) {
        $null = $sb.Append('<w:tc><w:tcPr><w:tcH w:val="1440" w:hRule="atLeast"/></w:tcPr>')
        $null = $sb.Append('<w:p><w:r><w:t xml:space="preserve">{0}</w:t></w:r></w:p></w:tc>' -f (Protect-Xml $LabelSig))
    }
    $null = $sb.Append('</w:tr>')
    $null = $sb.Append('</w:tbl>')
    $sb.ToString()
}

function Get-DocumentXml {
    param([string[]]$BodyParts)
    $body = $BodyParts -join "`n"
    @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
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

function New-WordDocument {
    param(
        [string]$OutputPath,
        [string[]]$BodyParts,
        [string]$Title = '',
        [string]$HeaderTitle = '',
        [string]$Creator = 'EXpress',
        [string]$ClassLabel = 'INTERN'
    )
    Add-Type -AssemblyName System.IO.Compression
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    $fs  = [System.IO.File]::Open($OutputPath, [System.IO.FileMode]::Create)
    $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    function Write-ZipEntry([string]$name, [string]$content) {
        $entry  = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
        $stream = $entry.Open()
        $bytes  = $utf8NoBom.GetBytes($content)
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Dispose()
    }
    Write-ZipEntry '[Content_Types].xml'        (Get-ContentTypesXml)
    Write-ZipEntry '_rels/.rels'                (Get-RootRelsXml)
    Write-ZipEntry 'docProps/core.xml'          (Get-CorePropsXml $Title $Creator)
    Write-ZipEntry 'word/_rels/document.xml.rels' (Get-DocRelsXml)
    Write-ZipEntry 'word/styles.xml'            (Get-StylesXml)
    Write-ZipEntry 'word/numbering.xml'         (Get-NumberingXml)
    Write-ZipEntry 'word/header1.xml'           (Get-HeaderXml ($HeaderTitle ? $HeaderTitle : $Title))
    Write-ZipEntry 'word/footer1.xml'           (Get-FooterXml $ClassLabel)
    Write-ZipEntry 'word/document.xml'          (Get-DocumentXml $BodyParts)
    $zip.Dispose()
    $fs.Dispose()
}

# ── F23 Content Generator ───────────────────────────────────────────────────────

function Get-F23Parts {
    param([string]$Language = 'DE')

    $DE = $Language -eq 'DE'

    $parts = [System.Collections.Generic.List[string]]::new()

    # ── Titelseite ──────────────────────────────────────────────────────────────
    $null = $parts.Add((Add-WordHeading ($DE ? 'Exchange Server Konzept- und Freigabedokument' : 'Exchange Server Concept and Approval Document') 1))
    $null = $parts.Add((Add-WordParagraph ''))

    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Eigenschaft' : 'Property'), ($DE ? 'Wert' : 'Value')) -Rows @(
        @(($DE ? 'Dokumentversion' : 'Document version'),   (Add-WordContentControl 'doc_version' ($DE ? 'z.B. 1.0' : 'e.g. 1.0') ($DE ? 'Dokumentversion' : 'Document version')))
        @(($DE ? 'Erstellt von'    : 'Created by'),         (Add-WordContentControl 'doc_author'  ($DE ? 'Vorname Nachname' : 'First Last') ($DE ? 'Erstellt von' : 'Created by')))
        @(($DE ? 'Datum'           : 'Date'),                (Add-WordContentControl 'doc_date'    ($DE ? 'TT.MM.JJJJ' : 'DD/MM/YYYY') ($DE ? 'Datum' : 'Date')))
        @(($DE ? 'Status'          : 'Status'),             (Add-WordContentControl 'doc_status'  ($DE ? 'Entwurf / Freigegeben' : 'Draft / Approved') ($DE ? 'Status' : 'Status')))
        @(($DE ? 'Klassifizierung' : 'Classification'),     'INTERN')
        @('EXpress Version',                                 (Add-WordContentControl 'doc_express_version' 'z.B. 5.78' 'EXpress Version'))
    )))

    $null = $parts.Add((Add-WordParagraph ''))

    # ── 1. Projektrahmen ────────────────────────────────────────────────────────
    $null = $parts.Add((Add-WordHeading ($DE ? '1. Projektrahmen' : '1. Project Framework') 1))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Eigenschaft' : 'Property'), ($DE ? 'Wert' : 'Value')) -Rows @(
        @(($DE ? 'Projektnummer / -name' : 'Project number / name'), (Add-WordContentControl 'proj_name' ($DE ? 'Projektnummer eingeben' : 'Enter project number')))
        @(($DE ? 'Auftraggeber'          : 'Sponsor'),               (Add-WordContentControl 'proj_sponsor' ($DE ? 'Organisation / Person' : 'Organisation / Person')))
        @(($DE ? 'Auftragnehmer'         : 'Contractor'),            (Add-WordContentControl 'proj_contractor' ($DE ? 'Name der IT-Abteilung oder des Dienstleisters' : 'IT department or service provider name')))
        @(($DE ? 'Projektleiter'         : 'Project manager'),       (Add-WordContentControl 'proj_manager' ($DE ? 'Vorname Nachname' : 'First Last')))
        @(($DE ? 'Erstellt am'           : 'Created on'),            (Add-WordContentControl 'proj_created' ($DE ? 'TT.MM.JJJJ' : 'DD/MM/YYYY')))
        @(($DE ? 'Geplante Umsetzung'    : 'Planned deployment'),    (Add-WordContentControl 'proj_date' ($DE ? 'TT.MM.JJJJ' : 'DD/MM/YYYY')))
    )))
    $null = $parts.Add((Add-WordParagraph ''))
    $null = $parts.Add((Add-WordHeading ($DE ? '1.1 Scope' : '1.1 Scope') 2))
    $null = $parts.Add((Add-WordContentControl 'proj_scope' ($DE ? 'Scope des Projekts beschreiben (Neuinstallation, Migration, DAG-Aufbau, Hybrid-Konfiguration ...)' : 'Describe project scope (new installation, migration, DAG setup, hybrid configuration ...)')))
    $null = $parts.Add((Add-WordHeading ($DE ? '1.2 Beteiligte' : '1.2 Participants') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Rolle' : 'Role'), ($DE ? 'Name' : 'Name'), ($DE ? 'Kontakt' : 'Contact')) -Rows @(
        @(($DE ? 'Projektleiter' : 'Project manager'), '', '')
        @(($DE ? 'Exchange-Administrator' : 'Exchange administrator'), '', '')
        @(($DE ? 'Active Directory-Administrator' : 'Active Directory administrator'), '', '')
        @(($DE ? 'Netzwerk-Administrator' : 'Network administrator'), '', '')
        @(($DE ? 'Backup-Verantwortlicher' : 'Backup administrator'), '', '')
        @(($DE ? 'Sicherheitsbeauftragter' : 'Security officer'), '', '')
    )))

    # ── 2. Lizenzbedingungen & Stand der Technik ────────────────────────────────
    $null = $parts.Add((Add-WordHeading ($DE ? '2. Lizenzbedingungen und Stand der Technik' : '2. Licensing and State of the Art') 1))
    $null = $parts.Add((Add-WordHeading ($DE ? '2.1 Unterstützte Versionen' : '2.1 Supported Versions') 2))
    if ($DE) {
        $null = $parts.Add((Add-WordParagraph 'Dieses Dokument beschreibt die Einführung von Exchange Server Subscription Edition (SE). Exchange Server 2016 und 2019 sind seit dem 14. Oktober 2025 außerhalb des Mainstream-Supports (End-of-Support). Ein Einsatz dieser Versionen in neuen Projekten wird ausdrücklich nicht empfohlen.'))
    } else {
        $null = $parts.Add((Add-WordParagraph 'This document describes the deployment of Exchange Server Subscription Edition (SE). Exchange Server 2016 and 2019 reached End-of-Support on October 14, 2025. Deploying these versions in new projects is explicitly not recommended.'))
    }
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Exchange-Version' : 'Exchange Version'), ($DE ? 'Windows Server' : 'Windows Server'), ($DE ? 'Support-Status' : 'Support Status')) -Rows @(
        @('Exchange SE RTM',  'WS 2022, WS 2025', ($DE ? 'Aktuell empfohlen' : 'Currently recommended'))
        @('Exchange SE CU1+', 'WS 2022, WS 2025', ($DE ? 'Aktuell empfohlen' : 'Currently recommended'))
        @('Exchange 2019 CU14/CU15', 'WS 2019, WS 2022', ($DE ? 'End-of-Support 14.10.2025 — kein Neueinsatz' : 'End-of-Support 14/10/2025 — no new deployments'))
        @('Exchange 2016 CU23', 'WS 2016', ($DE ? 'End-of-Support 14.10.2025 — kein Neueinsatz' : 'End-of-Support 14/10/2025 — no new deployments'))
    )))
    $null = $parts.Add((Add-WordHeading ($DE ? '2.2 Lizenzmodell — Subscription Edition' : '2.2 Licence Model — Subscription Edition') 2))
    if ($DE) {
        $null = $parts.Add((Add-WordParagraph 'Exchange Server SE erfordert eine Jahres-Subscription. Das Modell umfasst:'))
        $null = $parts.Add((Add-WordBullet 'Server-Lizenz (Subscription) pro Server'))
        $null = $parts.Add((Add-WordBullet 'Client Access Licences (CAL Subscription) pro Postfach ODER Inklusion über Microsoft 365 Business/Enterprise-Pläne'))
        $null = $parts.Add((Add-WordBullet 'Perpetual-Lizenzen (Kauflizenzen) für Exchange SE werden nicht angeboten'))
        $null = $parts.Add((Add-WordBullet 'Preisdetails: aktueller Microsoft-Preisleitfaden; keine kundenspezifischen Preise in diesem Dokument'))
    } else {
        $null = $parts.Add((Add-WordParagraph 'Exchange Server SE requires an annual subscription. The model includes:'))
        $null = $parts.Add((Add-WordBullet 'Server licence (subscription) per server'))
        $null = $parts.Add((Add-WordBullet 'Client Access Licences (CAL Subscription) per mailbox OR inclusion via Microsoft 365 Business/Enterprise plans'))
        $null = $parts.Add((Add-WordBullet 'Perpetual licences for Exchange SE are not available'))
        $null = $parts.Add((Add-WordBullet 'Pricing: current Microsoft price guide; no customer-specific prices in this document'))
    }
    $null = $parts.Add((Add-WordHeading ($DE ? '2.3 CU/SU-Kadenz' : '2.3 CU/SU Cadence') 2))
    if ($DE) {
        $null = $parts.Add((Add-WordParagraph 'Exchange SE wird im gleichen Quartals-CU-Rhythmus wie Exchange 2019 weiterentwickelt. Security Updates (SU) erscheinen monatlich (Patch Tuesday) für aktuelle CUs. EXpress installiert das jeweils aktuelle SU automatisch (Parameter -IncludeFixes).'))
    } else {
        $null = $parts.Add((Add-WordParagraph 'Exchange SE follows the same quarterly CU cadence as Exchange 2019. Security Updates (SU) are released monthly (Patch Tuesday) for current CUs. EXpress installs the latest SU automatically (-IncludeFixes parameter).'))
    }

    # ── 3. IST-Aufnahme ─────────────────────────────────────────────────────────
    $null = $parts.Add((Add-WordHeading ($DE ? '3. IST-Aufnahme' : '3. Current State Assessment') 1))
    $null = $parts.Add((Add-WordHeading ($DE ? '3.1 Active Directory' : '3.1 Active Directory') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Eigenschaft' : 'Property'), ($DE ? 'Wert' : 'Value')) -Rows @(
        @(($DE ? 'Forest Root Domain' : 'Forest Root Domain'), (Add-WordContentControl 'ad_forest' 'contoso.com'))
        @(($DE ? 'Forest Functional Level' : 'Forest Functional Level'), (Add-WordContentControl 'ad_ffl' ($DE ? 'z.B. Windows Server 2016' : 'e.g. Windows Server 2016')))
        @(($DE ? 'Domain Functional Level' : 'Domain Functional Level'), (Add-WordContentControl 'ad_dfl' ($DE ? 'z.B. Windows Server 2016' : 'e.g. Windows Server 2016')))
        @('FSMO — Schema Master', (Add-WordContentControl 'ad_fsmo_schema' ($DE ? 'FQDN des Servers' : 'Server FQDN')))
        @('FSMO — Domain Naming Master', (Add-WordContentControl 'ad_fsmo_naming' ($DE ? 'FQDN des Servers' : 'Server FQDN')))
        @('FSMO — PDC Emulator', (Add-WordContentControl 'ad_fsmo_pdc' ($DE ? 'FQDN des Servers' : 'Server FQDN')))
        @('FSMO — RID Master', (Add-WordContentControl 'ad_fsmo_rid' ($DE ? 'FQDN des Servers' : 'Server FQDN')))
        @('FSMO — Infrastructure Master', (Add-WordContentControl 'ad_fsmo_infra' ($DE ? 'FQDN des Servers' : 'Server FQDN')))
        @(($DE ? 'AD-Sites (Exchange-relevant)' : 'AD Sites (Exchange-relevant)'), (Add-WordContentControl 'ad_sites' ($DE ? 'Standortnamen eingeben' : 'Enter site names')))
        @(($DE ? 'Exchange-Schema-Version (objectVersion)' : 'Exchange schema version (objectVersion)'), (Add-WordContentControl 'ad_ex_schema' ($DE ? 'z.B. 17003 (Exchange SE RTM)' : 'e.g. 17003 (Exchange SE RTM)')))
    )))
    $null = $parts.Add((Add-WordHeading ($DE ? '3.2 Bestehende Exchange-Umgebung' : '3.2 Existing Exchange Environment') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Server' : 'Server'), ($DE ? 'Version' : 'Version'), ($DE ? 'Rolle' : 'Role'), ($DE ? 'Postfächer' : 'Mailboxes'), ($DE ? 'Anmerkung' : 'Note')) -Rows @(
        @((Add-WordContentControl 'ex_srv1' ($DE ? 'Servername' : 'Server name')), (Add-WordContentControl 'ex_ver1' ($DE ? 'z.B. Exchange 2019 CU14' : 'e.g. Exchange 2019 CU14')), (Add-WordContentControl 'ex_role1' ($DE ? 'Mailbox / Edge' : 'Mailbox / Edge')), (Add-WordContentControl 'ex_mbox1' ($DE ? 'Anzahl' : 'Count')), (Add-WordContentControl 'ex_note1' ''))
        @('', '', '', '', '')
    )))
    $null = $parts.Add((Add-WordHeading ($DE ? '3.3 Berechtigungen und Delegation' : '3.3 Permissions and Delegation') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Gruppe/Konto' : 'Group/Account'), ($DE ? 'Berechtigungsstufe' : 'Permission level'), ($DE ? 'Anmerkung' : 'Note')) -Rows @(
        @('Organization Management', ($DE ? 'Vollzugriff Exchange' : 'Full Exchange access'), ($DE ? 'Setup-Konto' : 'Setup account'))
        @('Domain Admins', ($DE ? 'AD-Vollzugriff' : 'AD full access'), '')
        @((Add-WordContentControl 'perm_svc1' ($DE ? 'Dienstkonto' : 'Service account')), '', '')
    )))
    $null = $parts.Add((Add-WordHeading ($DE ? '3.4 Netzwerk und DNS' : '3.4 Network and DNS') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Eigenschaft' : 'Property'), ($DE ? 'Wert' : 'Value')) -Rows @(
        @(($DE ? 'DNS-Server (intern)' : 'DNS server (internal)'), (Add-WordContentControl 'net_dns_int' ($DE ? 'IP-Adressen' : 'IP addresses')))
        @(($DE ? 'DNS-Suffix intern' : 'Internal DNS suffix'), (Add-WordContentControl 'net_dns_suffix' 'contoso.com'))
        @(($DE ? 'Externer DNS-Zonenbetreiber' : 'External DNS zone operator'), (Add-WordContentControl 'net_dns_ext' ($DE ? 'Hoster / selbst' : 'Hoster / self-managed')))
        @(($DE ? 'Exchange-Subnetz(e)' : 'Exchange subnet(s)'), (Add-WordContentControl 'net_subnets' ($DE ? 'z.B. 10.0.1.0/24' : 'e.g. 10.0.1.0/24')))
        @(($DE ? 'SMTP-Relay-Subnetz(e)' : 'SMTP relay subnet(s)'), (Add-WordContentControl 'net_relay_subnets' ($DE ? 'z.B. 10.0.2.0/24' : 'e.g. 10.0.2.0/24')))
        @(($DE ? 'IP-Adresse(n) Exchange-Server' : 'Exchange server IP address(es)'), (Add-WordContentControl 'net_ex_ips' ($DE ? 'IP-Adressen (CIDR)' : 'IP addresses (CIDR)')))
        @(($DE ? 'IP-Adresse(n) Load Balancer (VIP)' : 'Load balancer VIP(s)'), (Add-WordContentControl 'net_lb_vip' ($DE ? 'IP-Adressen' : 'IP addresses')))
        @(($DE ? 'Proxy / Smarthost (ausgehend)' : 'Proxy / Smarthost (outbound)'), (Add-WordContentControl 'net_smarthost' ($DE ? 'FQDN oder IP + Port' : 'FQDN or IP + port')))
    )))

    # ── 4. Sizing & Kapazitätsplanung ────────────────────────────────────────────
    $null = $parts.Add((Add-WordHeading ($DE ? '4. Sizing und Kapazitätsplanung' : '4. Sizing and Capacity Planning') 1))
    $null = $parts.Add((Add-WordHeading ($DE ? '4.1 Postfach-Profil' : '4.1 Mailbox Profile') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Parameter' : 'Parameter'), ($DE ? 'Wert' : 'Value')) -Rows @(
        @(($DE ? 'Anzahl Postfächer (gesamt)' : 'Total mailboxes'), (Add-WordContentControl 'sz_mbox_total' ($DE ? 'Anzahl' : 'Count')))
        @(($DE ? 'Durchschnittliche Postfachgröße' : 'Average mailbox size'), (Add-WordContentControl 'sz_mbox_avg' ($DE ? 'z.B. 5 GB' : 'e.g. 5 GB')))
        @(($DE ? 'Maximale Postfachgröße' : 'Maximum mailbox size'), (Add-WordContentControl 'sz_mbox_max' ($DE ? 'z.B. 50 GB' : 'e.g. 50 GB')))
        @(($DE ? 'Erwartetes Wachstum pro Jahr' : 'Expected annual growth'), (Add-WordContentControl 'sz_mbox_growth' ($DE ? 'z.B. 10%' : 'e.g. 10%')))
        @(($DE ? 'Nachrichtenprofil (IOPS)' : 'Message profile (IOPS)'), (Add-WordContentControl 'sz_iops' ($DE ? 'z.B. 0,1 IOPS/Postfach (niedrig)' : 'e.g. 0.1 IOPS/mailbox (low)')))
        @(($DE ? 'Concurrent Connections' : 'Concurrent connections'), (Add-WordContentControl 'sz_concurrent' ($DE ? 'Schätzwert' : 'Estimated value')))
        @(($DE ? 'Exchange Calculator-Referenz' : 'Exchange Calculator reference'), ($DE ? 'https://exdeploy.microsoft.com' : 'https://exdeploy.microsoft.com'))
    )))
    $null = $parts.Add((Add-WordHeading ($DE ? '4.2 CPU/RAM/Disk-Dimensionierung' : '4.2 CPU/RAM/Disk Sizing') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Komponente' : 'Component'), ($DE ? 'Minimum' : 'Minimum'), ($DE ? 'Empfohlen' : 'Recommended'), ($DE ? 'Geplant' : 'Planned')) -Rows @(
        @('CPU Kerne', '8', '16+', (Add-WordContentControl 'sz_cpu' ($DE ? 'Anzahl Kerne' : 'Core count')))
        @('RAM (GB)', '64', '128+', (Add-WordContentControl 'sz_ram' ($DE ? 'GB' : 'GB')))
        @(($DE ? 'DB-Disk (TB)' : 'DB disk (TB)'), ($DE ? 'Berechnet' : 'Calculated'), ($DE ? 'JBOD / RAID 5' : 'JBOD / RAID 5'), (Add-WordContentControl 'sz_db_disk' ($DE ? 'TB' : 'TB')))
        @(($DE ? 'Log-Disk (TB)' : 'Log disk (TB)'), ($DE ? 'Berechnet' : 'Calculated'), ($DE ? 'Separat' : 'Separate'), (Add-WordContentControl 'sz_log_disk' ($DE ? 'TB' : 'TB')))
        @('OS + Binaries (GB)', '80', '200', (Add-WordContentControl 'sz_os_disk' ($DE ? 'GB' : 'GB')))
    )))
    $null = $parts.Add((Add-WordHeading ($DE ? '4.3 Storage-Layout' : '4.3 Storage Layout') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Volume' : 'Volume'), ($DE ? 'Inhalt' : 'Content'), ($DE ? 'Cluster-Größe' : 'Allocation unit'), ($DE ? 'Dateisystem' : 'File system')) -Rows @(
        @('C:', ($DE ? 'OS, Exchange Binaries' : 'OS, Exchange Binaries'), '4 KB', 'NTFS')
        @((Add-WordContentControl 'sz_vol_db' ($DE ? 'z.B. E:' : 'e.g. E:')), ($DE ? 'Datenbank-Dateien (.edb)' : 'Database files (.edb)'), '64 KB', ($DE ? 'NTFS (ReFS: Prüfung erforderlich)' : 'NTFS (ReFS: verification required)'))
        @((Add-WordContentControl 'sz_vol_log' ($DE ? 'z.B. F:' : 'e.g. F:')), ($DE ? 'Transaktions-Logs' : 'Transaction logs'), '64 KB', 'NTFS')
        @((Add-WordContentControl 'sz_vol_queue' ($DE ? 'z.B. G:' : 'e.g. G:')), ($DE ? 'Transport-Queue' : 'Transport queue'), '4 KB', 'NTFS')
    )))

    # ── 5. SOLL-Architektur ──────────────────────────────────────────────────────
    $null = $parts.Add((Add-WordHeading ($DE ? '5. SOLL-Architektur' : '5. Target Architecture') 1))
    $null = $parts.Add((Add-WordHeading ($DE ? '5.1 Namensräume' : '5.1 Namespaces') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Dienst' : 'Service'), ($DE ? 'Intern (URL)' : 'Internal (URL)'), ($DE ? 'Extern (URL)' : 'External (URL)')) -Rows @(
        @('Autodiscover', (Add-WordContentControl 'ns_autodiscover_int' 'https://autodiscover.contoso.com/Autodiscover/Autodiscover.xml'), (Add-WordContentControl 'ns_autodiscover_ext' 'https://autodiscover.contoso.com/Autodiscover/Autodiscover.xml'))
        @('OWA', (Add-WordContentControl 'ns_owa_int' 'https://mail.contoso.com/owa'), (Add-WordContentControl 'ns_owa_ext' 'https://mail.contoso.com/owa'))
        @('ECP', (Add-WordContentControl 'ns_ecp_int' 'https://mail.contoso.com/ecp'), (Add-WordContentControl 'ns_ecp_ext' 'https://mail.contoso.com/ecp'))
        @('EWS', (Add-WordContentControl 'ns_ews_int' 'https://mail.contoso.com/EWS/Exchange.asmx'), (Add-WordContentControl 'ns_ews_ext' 'https://mail.contoso.com/EWS/Exchange.asmx'))
        @('OAB', (Add-WordContentControl 'ns_oab_int' 'https://mail.contoso.com/OAB'), (Add-WordContentControl 'ns_oab_ext' 'https://mail.contoso.com/OAB'))
        @('ActiveSync', (Add-WordContentControl 'ns_activesync_int' 'https://mail.contoso.com/Microsoft-Server-ActiveSync'), (Add-WordContentControl 'ns_activesync_ext' 'https://mail.contoso.com/Microsoft-Server-ActiveSync'))
        @('MAPI/HTTP', (Add-WordContentControl 'ns_mapi_int' 'https://mail.contoso.com/mapi'), (Add-WordContentControl 'ns_mapi_ext' 'https://mail.contoso.com/mapi'))
        @('OWA Download Domain', (Add-WordContentControl 'ns_dd_int' ($DE ? 'Optional: download.contoso.com' : 'Optional: download.contoso.com')), (Add-WordContentControl 'ns_dd_ext' ($DE ? 'Optional: download.contoso.com' : 'Optional: download.contoso.com')))
    )))
    $null = $parts.Add((Add-WordParagraph ($DE ? 'Split-DNS: Interne DNS-Zone für mail.contoso.com auf interne IP zeigen lassen; externe Zone auf öffentliche IP oder LB-VIP.' : 'Split-DNS: point internal DNS zone for mail.contoso.com to internal IP; external zone to public IP or LB VIP.')))
    $null = $parts.Add((Add-WordHeading ($DE ? '5.2 Server-Topologie' : '5.2 Server Topology') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Server' : 'Server'), ($DE ? 'Rolle' : 'Role'), ($DE ? 'Betriebssystem' : 'OS'), ($DE ? 'DAG-Mitglied' : 'DAG member'), ($DE ? 'Anmerkung' : 'Note')) -Rows @(
        @((Add-WordContentControl 'topo_srv1' ($DE ? 'Servername' : 'Server name')), ($DE ? 'Mailbox' : 'Mailbox'), (Add-WordContentControl 'topo_os1' 'WS 2025'), ($DE ? 'Ja / Nein' : 'Yes / No'), '')
        @((Add-WordContentControl 'topo_srv2' ($DE ? 'Servername' : 'Server name')), ($DE ? 'Mailbox' : 'Mailbox'), (Add-WordContentControl 'topo_os2' 'WS 2025'), ($DE ? 'Ja / Nein' : 'Yes / No'), '')
        @((Add-WordContentControl 'topo_srv_edge' ($DE ? 'Edge (optional)' : 'Edge (optional)')), ($DE ? 'Edge Transport' : 'Edge Transport'), (Add-WordContentControl 'topo_os_edge' 'WS 2025'), 'N/A', '')
    )))
    $null = $parts.Add((Add-WordHeading ($DE ? '5.3 DAG-Design' : '5.3 DAG Design') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Parameter' : 'Parameter'), ($DE ? 'Wert' : 'Value')) -Rows @(
        @(($DE ? 'DAG-Name' : 'DAG name'), (Add-WordContentControl 'dag_name' ($DE ? 'z.B. DAG01' : 'e.g. DAG01')))
        @(($DE ? 'DAG-IP (bei Mehrheit-Node)' : 'DAG IP (majority node)'), (Add-WordContentControl 'dag_ip' ($DE ? 'IP-Adresse oder keiner bei DHCP' : 'IP address or none for DHCP')))
        @(($DE ? 'File Share Witness (FSW)' : 'File Share Witness (FSW)'), (Add-WordContentControl 'dag_fsw' ($DE ? 'FQDN\\Freigabe' : 'FQDN\\share')))
        @(($DE ? 'Alternate FSW' : 'Alternate FSW'), (Add-WordContentControl 'dag_fsw_alt' ($DE ? 'FQDN\\Freigabe (optional)' : 'FQDN\\share (optional)')))
        @('DAC Mode', (Add-WordContentControl 'dag_dac' ($DE ? 'Enabled / Disabled' : 'Enabled / Disabled')))
        @(($DE ? 'Replikations-Netzwerk' : 'Replication network'), (Add-WordContentControl 'dag_repl_net' ($DE ? 'Subnetz (dediziert)' : 'Subnet (dedicated)')))
        @(($DE ? 'MAPI-Netzwerk' : 'MAPI network'), (Add-WordContentControl 'dag_mapi_net' ($DE ? 'Subnetz (Client-Zugriff)' : 'Subnet (client access)')))
        @(($DE ? 'Datenbank-Kopien pro DB' : 'Database copies per DB'), (Add-WordContentControl 'dag_copies' ($DE ? 'z.B. 2 (1 aktiv + 1 passiv)' : 'e.g. 2 (1 active + 1 passive)')))
        @(($DE ? 'Lag Copies' : 'Lag copies'), (Add-WordContentControl 'dag_lag' ($DE ? 'Ja / Nein — ReplayLagTime' : 'Yes / No — ReplayLagTime')))
        @(($DE ? 'Activation Preference' : 'Activation preference'), (Add-WordContentControl 'dag_actpref' ($DE ? 'Bevorzugte Aktivierungsreihenfolge' : 'Preferred activation order')))
    )))
    $null = $parts.Add((Add-WordHeading ($DE ? '5.4 Netzwerk und Load Balancer' : '5.4 Network and Load Balancer') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Parameter' : 'Parameter'), ($DE ? 'Wert' : 'Value')) -Rows @(
        @(($DE ? 'Load-Balancer-Produkt' : 'Load balancer product'), (Add-WordContentControl 'lb_product' ($DE ? 'z.B. F5, HAProxy, Windows NLB, keiner' : 'e.g. F5, HAProxy, Windows NLB, none')))
        @(($DE ? 'Persistence-Methode' : 'Persistence method'), (Add-WordContentControl 'lb_persistence' ($DE ? 'Source IP / Cookie / NTLM-passthrough' : 'Source IP / Cookie / NTLM passthrough')))
        @(($DE ? 'Health Probe' : 'Health probe'), ($DE ? 'HTTPS GET /owa/healthcheck.htm — erwartet HTTP 200' : 'HTTPS GET /owa/healthcheck.htm — expects HTTP 200'))
        @('SNAT', (Add-WordContentControl 'lb_snat' ($DE ? 'Ja / Nein' : 'Yes / No')))
        @(($DE ? 'LB High Availability' : 'LB high availability'), (Add-WordContentControl 'lb_ha' ($DE ? 'Aktiv/Passiv, Aktiv/Aktiv' : 'Active/Passive, Active/Active')))
    )))
    $null = $parts.Add((Add-WordHeading ($DE ? '5.5 Firewall-Matrix' : '5.5 Firewall Matrix') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Quelle' : 'Source'), ($DE ? 'Ziel' : 'Destination'), ($DE ? 'Port(s)' : 'Port(s)'), ($DE ? 'Protokoll' : 'Protocol'), ($DE ? 'Zweck' : 'Purpose')) -Rows @(
        @(($DE ? 'Clients (intern)' : 'Clients (internal)'), ($DE ? 'Exchange Mailbox' : 'Exchange Mailbox'), '443', 'TCP/HTTPS', ($DE ? 'OWA, EWS, MAPI, EAS, ECP, OAB' : 'OWA, EWS, MAPI, EAS, ECP, OAB'))
        @(($DE ? 'Exchange Mailbox' : 'Exchange Mailbox'), ($DE ? 'Exchange Mailbox (Replikation)' : 'Exchange Mailbox (replication)'), '64327', 'TCP', ($DE ? 'DAG-Datenbankreplikation' : 'DAG database replication'))
        @(($DE ? 'Exchange Mailbox' : 'Exchange Mailbox'), 'AD Domain Controller', '389/636/3268', 'TCP/UDP', ($DE ? 'LDAP / LDAPS / GC' : 'LDAP / LDAPS / GC'))
        @(($DE ? 'Exchange Mailbox' : 'Exchange Mailbox'), 'AD Domain Controller', '88/464', 'TCP/UDP', 'Kerberos')
        @(($DE ? 'Exchange Mailbox' : 'Exchange Mailbox'), 'AD Domain Controller', '445/135/49152–65535', 'TCP', ($DE ? 'RPC (NetLogon, SYSVOL)' : 'RPC (NetLogon, SYSVOL)'))
        @(($DE ? 'Exchange Mailbox' : 'Exchange Mailbox'), ($DE ? 'Internet / Smarthost' : 'Internet / Smarthost'), '25', 'TCP/SMTP', ($DE ? 'Ausgehende E-Mail' : 'Outbound email'))
        @(($DE ? 'Internet / MTA' : 'Internet / MTA'), ($DE ? 'Exchange Mailbox (oder Edge)' : 'Exchange Mailbox (or Edge)'), '25', 'TCP/SMTP', ($DE ? 'Eingehende E-Mail' : 'Inbound email'))
        @(($DE ? 'Exchange Edge' : 'Exchange Edge'), ($DE ? 'Exchange Mailbox' : 'Exchange Mailbox'), '50636', 'TCP/LDAPS', ($DE ? 'EdgeSync' : 'EdgeSync'))
        @(($DE ? 'Exchange Mailbox' : 'Exchange Mailbox'), ($DE ? 'FSW (File Share Witness)' : 'FSW (File Share Witness)'), '445', 'TCP', ($DE ? 'DAG-FSW (SMB)' : 'DAG FSW (SMB)'))
        @(($DE ? 'SMTP-Relay-Clients' : 'SMTP relay clients'), ($DE ? 'Exchange Mailbox' : 'Exchange Mailbox'), '587', 'TCP/SMTP', ($DE ? 'Authentifiziertes SMTP-Relay' : 'Authenticated SMTP relay'))
        @(($DE ? 'Anonym-Relay-Clients' : 'Anonymous relay clients'), ($DE ? 'Exchange Mailbox' : 'Exchange Mailbox'), '25', 'TCP/SMTP', ($DE ? 'Anonymes SMTP-Relay (Drucker, Scanner, Applikationen)' : 'Anonymous SMTP relay (printers, scanners, applications)'))
    )))
    $null = $parts.Add((Add-WordHeading ($DE ? '5.6 Datenbanken und Disk-Layout' : '5.6 Databases and Disk Layout') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Datenbank' : 'Database'), ($DE ? 'DB-Pfad' : 'DB path'), ($DE ? 'Log-Pfad' : 'Log path'), ($DE ? 'Größe (geplant)' : 'Size (planned)'), ($DE ? 'Kopien' : 'Copies')) -Rows @(
        @((Add-WordContentControl 'db_name1' ($DE ? 'z.B. MDB01' : 'e.g. MDB01')), (Add-WordContentControl 'db_path1' ($DE ? 'E:\MDB01\MDB01.edb' : 'E:\MDB01\MDB01.edb')), (Add-WordContentControl 'db_log1' ($DE ? 'F:\MDB01\' : 'F:\MDB01\')), (Add-WordContentControl 'db_size1' ($DE ? 'z.B. 500 GB' : 'e.g. 500 GB')), (Add-WordContentControl 'db_copies1' '2'))
        @('', '', '', '', '')
    )))
    $null = $parts.Add((Add-WordHeading ($DE ? '5.7 Konnektoren' : '5.7 Connectors') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Typ' : 'Type'), ($DE ? 'Name' : 'Name'), ($DE ? 'Scope' : 'Scope'), ($DE ? 'Anmerkung' : 'Note')) -Rows @(
        @(($DE ? 'Receive Connector' : 'Receive connector'), ($DE ? 'Default Frontend' : 'Default Frontend'), ($DE ? 'Eingehend SMTP 25' : 'Inbound SMTP 25'), ($DE ? 'Wird durch Exchange-Setup erstellt' : 'Created by Exchange setup'))
        @(($DE ? 'Receive Connector' : 'Receive connector'), ($DE ? 'Anonym-Relay' : 'Anonymous relay'), ($DE ? 'Relay-Subnetze' : 'Relay subnets'), ($DE ? 'EXpress: -RelaySubnets' : 'EXpress: -RelaySubnets'))
        @(($DE ? 'Send Connector' : 'Send connector'), (Add-WordContentControl 'con_send1' ($DE ? 'z.B. To Internet' : 'e.g. To Internet')), ($DE ? 'Alle Domänen (*)' : 'All domains (*)'), (Add-WordContentControl 'con_send1_note' ($DE ? 'Smarthost oder direkt' : 'Smarthost or direct')))
    )))
    $null = $parts.Add((Add-WordHeading ($DE ? '5.8 Zertifikatskonzept' : '5.8 Certificate Concept') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Eigenschaft' : 'Property'), ($DE ? 'Wert' : 'Value')) -Rows @(
        @(($DE ? 'Zertifikatstyp' : 'Certificate type'), (Add-WordContentControl 'cert_type' ($DE ? 'SAN/UCC — empfohlen' : 'SAN/UCC — recommended')))
        @(($DE ? 'Zertifizierungsstelle' : 'Certificate authority'), (Add-WordContentControl 'cert_ca' ($DE ? 'Öffentliche CA (z.B. DigiCert, Sectigo) oder interne CA' : 'Public CA (e.g. DigiCert, Sectigo) or internal CA')))
        @(($DE ? 'SAN-Einträge' : 'SAN entries'), (Add-WordContentControl 'cert_san' ($DE ? 'mail.contoso.com, autodiscover.contoso.com, ...' : 'mail.contoso.com, autodiscover.contoso.com, ...')))
        @(($DE ? 'Auth-Zertifikat (separat)' : 'Auth certificate (separate)'), ($DE ? 'Self-signed, automatisch durch Exchange erstellt — separater Rotationsprozess (60-Tage-Warnung)' : 'Self-signed, created automatically by Exchange — separate rotation process (60-day warning)'))
        @(($DE ? 'Gültigkeitsdauer' : 'Validity period'), (Add-WordContentControl 'cert_validity' ($DE ? 'z.B. 2 Jahre' : 'e.g. 2 years')))
        @(($DE ? 'Erneuerungsverantwortlicher' : 'Renewal owner'), (Add-WordContentControl 'cert_owner' ($DE ? 'Name / Team' : 'Name / Team')))
        @(($DE ? 'Rotationsstrategie' : 'Rotation strategy'), (Add-WordContentControl 'cert_rotation' ($DE ? '60 Tage vor Ablauf — Ticket in IT-System' : '60 days before expiry — ticket in IT system')))
    )))

    # ── 6. Sicherheits- & Härtungskonzept ────────────────────────────────────────
    $null = $parts.Add((Add-WordHeading ($DE ? '6. Sicherheits- und Härtungskonzept' : '6. Security and Hardening') 1))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Maßnahme' : 'Measure'), ($DE ? 'Einstellung' : 'Setting'), ($DE ? 'Status' : 'Status'), ($DE ? 'Referenz' : 'Reference')) -Rows @(
        @('TLS 1.0', ($DE ? 'Deaktiviert (SCHANNEL)' : 'Disabled (SCHANNEL)'), ($DE ? 'EXpress: -EnableTLS12' : 'EXpress: -EnableTLS12'), 'CIS L1 / PCI-DSS 4.2.1')
        @('TLS 1.1', ($DE ? 'Deaktiviert (SCHANNEL)' : 'Disabled (SCHANNEL)'), ($DE ? 'EXpress: -EnableTLS12' : 'EXpress: -EnableTLS12'), 'CIS L1 / PCI-DSS 4.2.1')
        @('TLS 1.2', ($DE ? 'Erzwungen (SCHANNEL + .NET)' : 'Enforced (SCHANNEL + .NET)'), ($DE ? 'EXpress: -EnableTLS12' : 'EXpress: -EnableTLS12'), 'CIS L1 / PCI-DSS 4.2.1')
        @('TLS 1.3', ($DE ? 'Aktiviert (WS 2022+)' : 'Enabled (WS 2022+)'), ($DE ? 'EXpress: -EnableTLS13' : 'EXpress: -EnableTLS13'), ($DE ? 'Best Practice' : 'Best practice'))
        @('SSL 3.0', ($DE ? 'Deaktiviert' : 'Disabled'), ($DE ? 'EXpress: -DisableSSL3' : 'EXpress: -DisableSSL3'), 'CIS L1')
        @('RC4', ($DE ? 'Deaktiviert' : 'Disabled'), ($DE ? 'EXpress: -DisableRC4' : 'EXpress: -DisableRC4'), 'CIS L1')
        @('SMBv1', ($DE ? 'Deaktiviert' : 'Disabled'), ($DE ? 'EXpress automatisch' : 'EXpress automatic'), 'CIS L1 §18.3 / BSI SYS.1')
        @('WDigest-Caching', ($DE ? 'Deaktiviert (UseLogonCredential=0)' : 'Disabled (UseLogonCredential=0)'), ($DE ? 'EXpress automatisch' : 'EXpress automatic'), 'CIS L1 §18.9.48 / DISA STIG')
        @('LSA Protection (RunAsPPL)', ($DE ? 'Aktiviert (WS 2019+/SE)' : 'Enabled (WS 2019+/SE)'), ($DE ? 'EXpress automatisch' : 'EXpress automatic'), 'CIS L2 §2.3.11')
        @('LM Compatibility Level', ($DE ? 'Level 5 (NTLMv2 only)' : 'Level 5 (NTLMv2 only)'), ($DE ? 'EXpress automatisch' : 'EXpress automatic'), 'CIS L1 §2.3.11.7 / BSI')
        @('Credential Guard', ($DE ? 'Deaktiviert (Exchange-Server)' : 'Disabled (Exchange servers)'), ($DE ? 'EXpress automatisch' : 'EXpress automatic'), 'CIS L2')
        @('AMSI', ($DE ? 'Aktiviert (2016/2019; SE default)' : 'Enabled (2016/2019; SE default)'), ($DE ? 'EXpress: -EnableAMSI' : 'EXpress: -EnableAMSI'), 'MS Security Best Practice')
        @('Extended Protection (EP)', ($DE ? 'Aktiviert (CU14+/SE)' : 'Enabled (CU14+/SE)'), ($DE ? 'EXpress automatisch' : 'EXpress automatic'), 'MS Security Advisory')
        @('Serialized Data Signing (SDS)', ($DE ? 'Aktiviert' : 'Enabled'), ($DE ? 'EXpress automatisch' : 'EXpress automatic'), 'Exchange Security Hardening')
        @('HTTP/2', ($DE ? 'Deaktiviert (Exchange-Kompatibilität)' : 'Disabled (Exchange compatibility)'), ($DE ? 'EXpress automatisch' : 'EXpress automatic'), 'MS Exchange Guidance')
        @('OWA Download Domain', ($DE ? 'Separat (Malware-Sandbox)' : 'Separate (malware sandbox)'), ($DE ? 'EXpress: -DownloadDomain' : 'EXpress: -DownloadDomain'), 'Exchange Security Hardening')
        @('Defender Exclusions', ($DE ? 'Exchange-Pfade ausgeschlossen' : 'Exchange paths excluded'), ($DE ? 'EXpress automatisch' : 'EXpress automatic'), 'MS Exchange Guidance')
        @('SSL Offloading', ($DE ? 'Deaktiviert (EP-Voraussetzung)' : 'Disabled (EP prerequisite)'), ($DE ? 'EXpress automatisch' : 'EXpress automatic'), 'MS Exchange Guidance')
        @('MAPI Encryption Required', ($DE ? 'Aktiviert' : 'Enabled'), ($DE ? 'EXpress automatisch' : 'EXpress automatic'), ($DE ? 'Interne Verschlüsselung' : 'Internal encryption'))
        @('Root Certificate AutoUpdate', ($DE ? 'Aktiviert' : 'Enabled'), ($DE ? 'EXpress automatisch' : 'EXpress automatic'), 'MS Guidance')
        @('MRS Proxy', ($DE ? 'Deaktiviert' : 'Disabled'), ($DE ? 'EXpress automatisch' : 'EXpress automatic'), ($DE ? 'Angriffsfläche reduzieren' : 'Reduce attack surface'))
    )))
    $null = $parts.Add((Add-WordParagraph ($DE ? 'Hinweis: Alle genannten Härtungsmaßnahmen werden durch EXpress automatisch konfiguriert. Die Compliance-Mapping-Spalte verweist auf CIS Benchmark (L1/L2), BSI IT-Grundschutz und DISA STIG.' : 'Note: All hardening measures listed above are configured automatically by EXpress. The compliance column references CIS Benchmark (L1/L2), BSI IT-Grundschutz and DISA STIG.')))

    # ── 7. Message Hygiene ────────────────────────────────────────────────────────
    $null = $parts.Add((Add-WordHeading ($DE ? '7. Message Hygiene' : '7. Message Hygiene') 1))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Lösung' : 'Solution'), ($DE ? 'Beschreibung' : 'Description'), ($DE ? 'Geplant' : 'Planned')) -Rows @(
        @(($DE ? 'Exchange Anti-Spam-Agents' : 'Exchange Anti-Spam Agents'), ($DE ? 'Integrierte Content/Sender/Recipient-Filter; keine Cloud-Anbindung erforderlich' : 'Built-in content/sender/recipient filters; no cloud connectivity required'), (Add-WordContentControl 'hygiene_builtin' ($DE ? 'Ja / Nein' : 'Yes / No')))
        @(($DE ? 'Edge Transport Server' : 'Edge Transport Server'), ($DE ? 'Dedizierter Edge-Server in DMZ; verhindert Mailbox-Kontakt mit Internet-SMTP' : 'Dedicated Edge server in DMZ; prevents mailbox contact with Internet SMTP'), (Add-WordContentControl 'hygiene_edge' ($DE ? 'Ja / Nein' : 'Yes / No')))
        @('Hornetsecurity', ($DE ? 'Cloud-basierter Dienst; MX auf Hornetsecurity zeigen' : 'Cloud-based service; MX pointed at Hornetsecurity'), (Add-WordContentControl 'hygiene_hornetsec' ($DE ? 'Ja / Nein' : 'Yes / No')))
        @('Proofpoint', ($DE ? 'Kommerzieller Gateway-Dienst' : 'Commercial gateway service'), (Add-WordContentControl 'hygiene_proofpoint' ($DE ? 'Ja / Nein' : 'Yes / No')))
        @('Mimecast', ($DE ? 'Kommerzieller Gateway-Dienst' : 'Commercial gateway service'), (Add-WordContentControl 'hygiene_mimecast' ($DE ? 'Ja / Nein' : 'Yes / No')))
        @(($DE ? 'Sonstiger Drittanbieter' : 'Other third-party'), (Add-WordContentControl 'hygiene_other_desc' ($DE ? 'Beschreibung' : 'Description')), (Add-WordContentControl 'hygiene_other' ($DE ? 'Ja / Nein' : 'Yes / No')))
    )))
    $null = $parts.Add((Add-WordContentControl 'hygiene_notes' ($DE ? 'Zusätzliche Hinweise zur Message-Hygiene-Strategie eingeben...' : 'Enter additional notes on message hygiene strategy...')))

    # ── 8. Backup, Recovery & Disaster Recovery ────────────────────────────────
    $null = $parts.Add((Add-WordHeading ($DE ? '8. Backup, Recovery und Disaster Recovery' : '8. Backup, Recovery and Disaster Recovery') 1))
    $null = $parts.Add((Add-WordHeading ($DE ? '8.1 Backup-Strategie und VSS-Integration' : '8.1 Backup Strategy and VSS Integration') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Parameter' : 'Parameter'), ($DE ? 'Wert' : 'Value')) -Rows @(
        @(($DE ? 'Backup-Lösung' : 'Backup solution'), (Add-WordContentControl 'bkp_solution' ($DE ? 'z.B. Veeam, Windows Server Backup, Commvault' : 'e.g. Veeam, Windows Server Backup, Commvault')))
        @(($DE ? 'Backup-Typ' : 'Backup type'), (Add-WordContentControl 'bkp_type' ($DE ? 'VSS-aware Full + Incremental; oder DB-Snapshot' : 'VSS-aware full + incremental; or DB snapshot')))
        @(($DE ? 'Backup-Intervall' : 'Backup interval'), (Add-WordContentControl 'bkp_interval' ($DE ? 'z.B. täglich, stündlich' : 'e.g. daily, hourly')))
        @(($DE ? 'Aufbewahrungsdauer' : 'Retention period'), (Add-WordContentControl 'bkp_retention' ($DE ? 'z.B. 30 Tage + 52 Wochen' : 'e.g. 30 days + 52 weeks')))
        @(($DE ? 'Backup-Ziel' : 'Backup target'), (Add-WordContentControl 'bkp_target' ($DE ? 'Lokales Tape / NAS / Cloud (Azure Backup)' : 'Local tape / NAS / cloud (Azure Backup)')))
        @(($DE ? 'Restore-Test-Kadenz' : 'Restore test cadence'), (Add-WordContentControl 'bkp_test' ($DE ? 'z.B. monatlich — isolierte Test-DB auf Lab-Server' : 'e.g. monthly — isolated test DB on lab server')))
        @(($DE ? 'Circular Logging' : 'Circular logging'), (Add-WordContentControl 'bkp_circular' ($DE ? 'Nein — empfohlen nur in DAG mit mehreren Kopien' : 'No — recommended only in DAG with multiple copies')))
        @('Recovery Database (RDB)', (Add-WordContentControl 'bkp_rdb' ($DE ? 'Vorgehen bei Item-Level-Recovery aus Backup beschreiben' : 'Describe item-level recovery process from backup')))
    )))
    $null = $parts.Add((Add-WordHeading ($DE ? '8.2 Transaktions-Log-Verwaltung und Truncation' : '8.2 Transaction Log Management and Truncation') 2))
    $null = $parts.Add((Add-WordParagraph ($DE ? 'In einer DAG: Logs werden nach dem Commit auf alle passiven Kopien automatisch per Log-Truncation entfernt (Continuous Log Truncation). Dies setzt voraus, dass alle Kopien den Log empfangen haben. Circular Logging ist in dieser Konstellation nur erlaubt, wenn alle Kopien gesund sind und ein VSS-Backup läuft.' : 'In a DAG: logs are automatically removed after committing to all passive copies via Continuous Log Truncation. This requires all copies to have received the log. Circular logging in this configuration is only permitted when all copies are healthy and a VSS backup is running.')))
    $null = $parts.Add((Add-WordHeading ($DE ? '8.3 DR-Szenarien' : '8.3 DR Scenarios') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Szenario' : 'Scenario'), ($DE ? 'Lösung' : 'Solution'), ($DE ? 'RTO / RPO' : 'RTO / RPO')) -Rows @(
        @(($DE ? 'FSW-Ausfall' : 'FSW failure'), ($DE ? 'Alternate FSW (Pre-Staged) wird automatisch aktiviert; oder manuelle Umstellung per Set-DatabaseAvailabilityGroup' : 'Alternate FSW (pre-staged) activates automatically; or manual switch via Set-DatabaseAvailabilityGroup'), (Add-WordContentControl 'dr_rto_fsw' ($DE ? 'RTO: <1 Min.' : 'RTO: <1 min')))
        @(($DE ? 'Server-Ausfall (DAG-Mitglied)' : 'Server failure (DAG member)'), ($DE ? 'Automatisches Failover auf passive Kopie; Activation Preference bestimmt Ziel-Server' : 'Automatic failover to passive copy; Activation Preference determines target server'), (Add-WordContentControl 'dr_rto_server' ($DE ? 'RTO: <5 Min., RPO: nahezu 0' : 'RTO: <5 min, RPO: near 0')))
        @(($DE ? 'Server-Ausfall (ohne DAG)' : 'Server failure (no DAG)'), ($DE ? 'setup.exe /m:RecoverServer; Restore Exchange DB aus Backup' : 'setup.exe /m:RecoverServer; restore Exchange DB from backup'), (Add-WordContentControl 'dr_rto_nodag' ($DE ? 'RTO: abhängig von Backup-Größe' : 'RTO: depends on backup size')))
        @(($DE ? 'Split-Brain / DAC-Modus' : 'Split-brain / DAC mode'), ($DE ? 'DAC-Modus verhindert doppeltes Mounten. Manuell: Stop-DatabaseAvailabilityGroup + Restore-DatabaseAvailabilityGroup' : 'DAC mode prevents double-mounting. Manual: Stop-DatabaseAvailabilityGroup + Restore-DatabaseAvailabilityGroup'), '')
        @(($DE ? 'Namespace-Failover' : 'Namespace failover'), ($DE ? 'DNS A-Record auf sekundären Server/LB umbiegen; oder LB-Failover-Gruppe' : 'Redirect DNS A record to secondary server/LB; or LB failover group'), (Add-WordContentControl 'dr_rto_ns' ($DE ? 'RTO: TTL-abhängig' : 'RTO: TTL-dependent')))
        @(($DE ? 'Gesamtausfall / Desaster' : 'Total outage / disaster'), ($DE ? 'Exchange SE kann auf neuem Windows-Server per RecoverServer neu installiert werden; DB aus Backup einspielen' : 'Exchange SE can be reinstalled on new Windows Server via RecoverServer; restore DB from backup'), (Add-WordContentControl 'dr_rto_full' ($DE ? 'RTO: geplante Zeit lt. Backup-Strategie' : 'RTO: planned time per backup strategy')))
    )))

    # ── 9. Monitoring-Konzept ─────────────────────────────────────────────────────
    $null = $parts.Add((Add-WordHeading ($DE ? '9. Monitoring-Konzept' : '9. Monitoring Concept') 1))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Eigenschaft' : 'Property'), ($DE ? 'Wert' : 'Value')) -Rows @(
        @(($DE ? 'Monitoring-Lösung' : 'Monitoring solution'), (Add-WordContentControl 'mon_solution' ($DE ? 'z.B. PRTG, Checkmk, SCOM, Zabbix' : 'e.g. PRTG, Checkmk, SCOM, Zabbix')))
        @(($DE ? 'Exchange HealthChecker' : 'Exchange HealthChecker'), ($DE ? 'EXpress führt HealthChecker automatisch in Phase 6 aus; Ergebnis-HTML im Report-Ordner' : 'EXpress runs HealthChecker automatically in Phase 6; result HTML in reports folder'))
        @(($DE ? 'Managed Availability' : 'Managed availability'), ($DE ? 'Exchange-internes Self-Healing; Event-Log-Kanal: Microsoft-Exchange-ManagedAvailability/Monitoring' : 'Exchange-internal self-healing; event log channel: Microsoft-Exchange-ManagedAvailability/Monitoring'))
        @(($DE ? 'Alarmierung' : 'Alerting'), (Add-WordContentControl 'mon_alerting' ($DE ? 'E-Mail / SMS / Teams-Kanal' : 'Email / SMS / Teams channel')))
        @(($DE ? 'Schwellenwerte' : 'Thresholds'), (Add-WordContentControl 'mon_thresholds' ($DE ? 'CPU >80%, RAM >90%, Disk <10%, Queue >500 Messages' : 'CPU >80%, RAM >90%, disk <10%, queue >500 messages')))
        @(($DE ? 'Event-Log-Aufbewahrung' : 'Event log retention'), (Add-WordContentControl 'mon_evtlog' ($DE ? 'z.B. 4 Wochen lokal, 6 Monate SIEM' : 'e.g. 4 weeks local, 6 months SIEM')))
        @(($DE ? 'Performance-Baseline' : 'Performance baseline'), ($DE ? 'Perfmon-Baseline innerhalb der ersten 4 Wochen nach Go-Live aufzeichnen' : 'Record Perfmon baseline within the first 4 weeks after go-live'))
    )))
    $null = $parts.Add((Add-WordHeading ($DE ? '9.1 Wichtige Event-IDs' : '9.1 Key Event IDs') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Event-ID' : 'Event ID'), ($DE ? 'Quelle' : 'Source'), ($DE ? 'Bedeutung' : 'Meaning')) -Rows @(
        @('1000/1002', 'MSExchangeIS', ($DE ? 'Information Store Start/Stop' : 'Information Store start/stop'))
        @('9001', 'MSExchangeIS Mailbox Store', ($DE ? 'DB-Mount fehlgeschlagen' : 'DB mount failed'))
        @('106', 'FailoverClustering', ($DE ? 'Cluster-Ressource offline' : 'Cluster resource offline'))
        @('4999', 'MSExchange Common', ($DE ? 'Application-Crash/Watson-Report' : 'Application crash/Watson report'))
        @('8026', 'MSExchangeTransport', ($DE ? 'NDR-Storm' : 'NDR storm'))
        @('15002', 'MSExchangeTransport', ($DE ? 'Transport-Backpressure aktiv' : 'Transport back pressure active'))
        @('2008', 'MSExchangeRepl', ($DE ? 'Replikationsunterbrechung (DAG)' : 'Replication interruption (DAG)'))
    )))

    # ── 10. Migration / Koexistenz ────────────────────────────────────────────────
    $null = $parts.Add((Add-WordHeading ($DE ? '10. Migration und Koexistenz (konditional)' : '10. Migration and Coexistence (conditional)') 1))
    $null = $parts.Add((Add-WordParagraph ($DE ? 'Dieses Kapitel ist nur relevant, wenn eine bestehende Exchange-Organisation (Exchange 2016/2019) migriert wird. Ohne Migration-Szenario kann dieses Kapitel entfernt oder als "Nicht zutreffend" markiert werden.' : 'This chapter is only relevant if an existing Exchange organisation (Exchange 2016/2019) is being migrated. Without a migration scenario this chapter may be removed or marked "Not applicable".')))
    $null = $parts.Add((Add-WordHeading ($DE ? '10.1 Legacy-2016/2019 zu Exchange SE' : '10.1 Legacy 2016/2019 to Exchange SE') 2))
    $null = $parts.Add((Add-WordParagraph ($DE ? 'Exchange 2016 und 2019 sind seit dem 14. Oktober 2025 out-of-support. Die Migration zu Exchange SE sollte prioritär behandelt werden. Koexistenz ist möglich, aber der Legacy-Server ist nach dem Datum ein Sicherheitsrisiko.' : 'Exchange 2016 and 2019 reached End-of-Support on October 14, 2025. Migration to Exchange SE should be treated as a priority. Coexistence is possible but the legacy server is a security risk after that date.')))
    $null = $parts.Add((Add-WordHeading ($DE ? '10.2 Cutover vs. schrittweise Koexistenz' : '10.2 Cutover vs. Staged Coexistence') 2))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Strategie' : 'Strategy'), ($DE ? 'Beschreibung' : 'Description'), ($DE ? 'Empfohlen für' : 'Recommended for')) -Rows @(
        @(($DE ? 'Cutover (Wochenende)' : 'Cutover (weekend)'), ($DE ? 'Alle Postfächer werden in einem Wartungsfenster migriert; kein Koexistenz-Zeitraum' : 'All mailboxes migrated in one maintenance window; no coexistence period'), ($DE ? 'Kleine Umgebungen (<500 Postfächer)' : 'Small environments (<500 mailboxes)'))
        @(($DE ? 'Schrittweise Koexistenz' : 'Staged coexistence'), ($DE ? 'Postfächer werden in Gruppen migriert; Legacy-Server bleibt bis zum Ende aktiv' : 'Mailboxes migrated in groups; legacy server stays active until end'), ($DE ? 'Größere Umgebungen oder kritische Zeitfenster' : 'Larger environments or critical time windows'))
    )))
    $null = $parts.Add((Add-WordHeading ($DE ? '10.3 Public-Folder-Migration (Legacy zu Modern)' : '10.3 Public Folder Migration (Legacy to Modern)') 2))
    $null = $parts.Add((Add-WordParagraph ($DE ? 'Public Folders aus Exchange 2016/2019 müssen manuell zu Modern Public Folders migriert werden, bevor das Legacy-System dekommissioniert wird. Verwende Get-PublicFolderMigrationRequestStatistics für die Überwachung.' : 'Public folders from Exchange 2016/2019 must be manually migrated to Modern Public Folders before the legacy system is decommissioned. Use Get-PublicFolderMigrationRequestStatistics for monitoring.')))
    $null = $parts.Add((Add-WordHeading ($DE ? '10.4 Namespace-Migration' : '10.4 Namespace Migration') 2))
    $null = $parts.Add((Add-WordParagraph ($DE ? 'Namespaces (DNS-Einträge) sollten erst umgebogen werden, wenn alle VDir-URLs auf dem neuen Exchange SE-Server konfiguriert sind. EXpress konfiguriert alle VDir-URLs automatisch (-Namespace).' : 'Namespaces (DNS records) should only be redirected after all VDir URLs are configured on the new Exchange SE server. EXpress configures all VDir URLs automatically (-Namespace).')))

    # ── 11. Hybrid / M365-Integration ─────────────────────────────────────────────
    $null = $parts.Add((Add-WordHeading ($DE ? '11. Hybrid und M365-Integration' : '11. Hybrid and M365 Integration') 1))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Parameter' : 'Parameter'), ($DE ? 'Wert' : 'Value')) -Rows @(
        @(($DE ? 'Hybrid-Konfiguration geplant?' : 'Hybrid configuration planned?'), (Add-WordContentControl 'hyb_planned' ($DE ? 'Ja / Nein' : 'Yes / No')))
        @(($DE ? 'HCW-Typ' : 'HCW type'), (Add-WordContentControl 'hyb_type' ($DE ? 'Full Hybrid / Minimal Hybrid / keiner' : 'Full Hybrid / Minimal Hybrid / none')))
        @('OAuth / Modern Auth', (Add-WordContentControl 'hyb_oauth' ($DE ? 'Konfiguriert über HCW' : 'Configured via HCW')))
        @(($DE ? 'Mail-Flow' : 'Mail flow'), (Add-WordContentControl 'hyb_mailflow' ($DE ? 'Zentral (über Exchange) / Direkt (Exchange Online)' : 'Centralised (via Exchange) / Direct (Exchange Online)')))
        @('Free/Busy', (Add-WordContentControl 'hyb_freebusy' ($DE ? 'Organisation Relationship / OAuth' : 'Organisation Relationship / OAuth')))
        @('Federation Trust', (Add-WordContentControl 'hyb_fed' ($DE ? 'Wird durch HCW erstellt' : 'Created by HCW')))
    )))

    # ── 12. Public Folders ────────────────────────────────────────────────────────
    $null = $parts.Add((Add-WordHeading ($DE ? '12. Public Folders und moderne Alternativen' : '12. Public Folders and Modern Alternatives') 1))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Entscheidung' : 'Decision'), ($DE ? 'Begründung' : 'Rationale')) -Rows @(
        @(($DE ? '&#x2610; Public Folders werden eingesetzt (Modern Mailbox-basiert)' : '&#x2610; Public folders will be used (Modern Mailbox-based)'), (Add-WordContentControl 'pf_design' ($DE ? 'Hierarchie, Anzahl PF, Größen beschreiben' : 'Describe hierarchy, count, sizes')))
        @(($DE ? '&#x2610; Public Folders werden NICHT eingesetzt' : '&#x2610; Public folders will NOT be used'), ($DE ? 'Public Folders werden nicht eingesetzt. Moderne Ablösung über Shared Mailboxes (gemeinsamer Posteingang, Kalender) und Microsoft Teams (Dokumentenablage, Zusammenarbeit). Begründung: geringerer Administrationsaufwand, Cloud-native, bessere Mobile-/Web-Experience, keine DAG-Replikations-Abhängigkeit.' : 'Public folders will not be deployed. Modern replacement via Shared Mailboxes (shared inbox, calendar) and Microsoft Teams (document storage, collaboration). Rationale: lower administrative overhead, cloud-native, better mobile/web experience, no DAG replication dependency.'))
    )))

    # ── 13. Compliance / eDiscovery / Journaling ──────────────────────────────────
    $null = $parts.Add((Add-WordHeading ($DE ? '13. Compliance, eDiscovery und Journaling' : '13. Compliance, eDiscovery and Journaling') 1))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Feature' : 'Feature'), ($DE ? 'Geplant' : 'Planned'), ($DE ? 'Konfiguration' : 'Configuration')) -Rows @(
        @('Litigation Hold', (Add-WordContentControl 'comp_lhold' ($DE ? 'Ja / Nein' : 'Yes / No')), (Add-WordContentControl 'comp_lhold_conf' ($DE ? 'Dauer, betroffene Postfächer' : 'Duration, affected mailboxes')))
        @('In-Place Archive', (Add-WordContentControl 'comp_archive' ($DE ? 'Ja / Nein' : 'Yes / No')), (Add-WordContentControl 'comp_archive_conf' ($DE ? 'Auto-Archivierungsrichtlinie' : 'Auto-archive policy')))
        @('Retention Policies', (Add-WordContentControl 'comp_retention' ($DE ? 'Ja / Nein' : 'Yes / No')), (Add-WordContentControl 'comp_retention_conf' ($DE ? 'Aufbewahrungsfristen (DSGVO: i.d.R. 6-10 Jahre)' : 'Retention periods (GDPR: typically 6-10 years)')))
        @('Journal Rules', (Add-WordContentControl 'comp_journal' ($DE ? 'Ja / Nein' : 'Yes / No')), (Add-WordContentControl 'comp_journal_conf' ($DE ? 'Journalziel, interne/externe Empfänger' : 'Journal target, internal/external recipients')))
        @('DLP (Data Loss Prevention)', (Add-WordContentControl 'comp_dlp' ($DE ? 'Ja / Nein' : 'Yes / No')), (Add-WordContentControl 'comp_dlp_conf' ($DE ? 'Regelwerk, Ausnahmen' : 'Rule set, exceptions')))
        @(($DE ? 'DSGVO-Auskunftspflicht' : 'GDPR data subject access'), (Add-WordContentControl 'comp_dsgvo' ($DE ? 'Prozess definiert?' : 'Process defined?')), (Add-WordContentControl 'comp_dsgvo_conf' ($DE ? 'eDiscovery-Suche + Export-Prozess' : 'eDiscovery search + export process')))
    )))

    # ── 14. Mobile & ActiveSync ────────────────────────────────────────────────────
    $null = $parts.Add((Add-WordHeading ($DE ? '14. Mobile und ActiveSync' : '14. Mobile and ActiveSync') 1))
    $null = $parts.Add((Add-WordTable -Headers @(($DE ? 'Parameter' : 'Parameter'), ($DE ? 'Wert' : 'Value')) -Rows @(
        @(($DE ? 'MDM/EMM-Lösung' : 'MDM/EMM solution'), (Add-WordContentControl 'mob_mdm' ($DE ? 'z.B. Intune, MobileIron, keiner' : 'e.g. Intune, MobileIron, none')))
        @(($DE ? 'ActiveSync-Policy (Gerätezugriffsregeln)' : 'ActiveSync policy (device access rules)'), (Add-WordContentControl 'mob_as_policy' ($DE ? 'Allow All / Block / Quarantine' : 'Allow All / Block / Quarantine')))
        @(($DE ? 'Quarantäne-Genehmigungsprozess' : 'Quarantine approval process'), (Add-WordContentControl 'mob_quarantine' ($DE ? 'Wer genehmigt neue Geräte?' : 'Who approves new devices?')))
        @(($DE ? 'Passwort-Anforderungen' : 'Password requirements'), (Add-WordContentControl 'mob_password' ($DE ? 'min. 6-stellig, Sperrzeit, Remote-Wipe' : 'min. 6 characters, lock time, remote wipe')))
        @(($DE ? 'Intune-Conditional-Access' : 'Intune conditional access'), (Add-WordContentControl 'mob_intune_ca' ($DE ? 'Ja / Nein — ActiveSync-Compliance-Anforderung' : 'Yes / No — ActiveSync compliance requirement')))
    )))

    # ── 15. Fragenkatalog ──────────────────────────────────────────────────────────
    $null = $parts.Add((Add-WordHeading ($DE ? '15. Fragenkatalog' : '15. Questionnaire') 1))
    $null = $parts.Add((Add-WordParagraph ($DE ? 'Alle offenen Parameter aus den Kapiteln 3-14. Bitte vor der Freigabe vollständig ausfüllen.' : 'All open parameters from Chapters 3-14. Please complete fully before approval.')))
    $questions = @(
        @('1',  ($DE ? 'Wie lautet der interne Exchange-Namespace (z.B. mail.contoso.com)?' : 'What is the internal Exchange namespace (e.g. mail.contoso.com)?'), 'q_ns_internal', ($DE ? 'Namespace eingeben...' : 'Enter namespace...'))
        @('2',  ($DE ? 'Wie lautet der externe Exchange-Namespace?' : 'What is the external Exchange namespace?'), 'q_ns_external', ($DE ? 'Namespace eingeben...' : 'Enter namespace...'))
        @('3',  ($DE ? 'Wie viele Mailbox-Server werden eingesetzt?' : 'How many mailbox servers will be deployed?'), 'q_srv_count', ($DE ? 'Anzahl...' : 'Count...'))
        @('4',  ($DE ? 'Wird ein DAG konfiguriert? Wenn ja: wie viele Kopien pro Datenbank?' : 'Will a DAG be configured? If yes: how many copies per database?'), 'q_dag', ($DE ? 'Ja/Nein, Anzahl Kopien...' : 'Yes/No, copy count...'))
        @('5',  ($DE ? 'Auf welchem Windows Server läuft Exchange SE? (WS 2022 / WS 2025)' : 'Which Windows Server version for Exchange SE? (WS 2022 / WS 2025)'), 'q_os_version', ($DE ? 'Betriebssystem eingeben...' : 'Enter OS version...'))
        @('6',  ($DE ? 'Wird ein Load Balancer eingesetzt? Produkt und Persistence-Methode?' : 'Will a load balancer be used? Product and persistence method?'), 'q_lb', ($DE ? 'LB-Typ und Methode eingeben...' : 'Enter LB type and method...'))
        @('7',  ($DE ? 'Wird ein Edge Transport Server in der DMZ eingesetzt?' : 'Will an Edge Transport server be deployed in the DMZ?'), 'q_edge', ($DE ? 'Ja/Nein...' : 'Yes/No...'))
        @('8',  ($DE ? 'Von welcher CA wird das SSL-Zertifikat bezogen? (öffentlich/intern)' : 'Which CA will issue the SSL certificate? (public/internal)'), 'q_cert_ca', ($DE ? 'CA-Name eingeben...' : 'Enter CA name...'))
        @('9',  ($DE ? 'Welche SAN-Einträge sind erforderlich?' : 'Which SAN entries are required?'), 'q_cert_san', ($DE ? 'SANs eingeben...' : 'Enter SANs...'))
        @('10', ($DE ? 'Welches Backup-Produkt wird eingesetzt? VSS-fähig?' : 'Which backup product will be used? VSS-capable?'), 'q_backup', ($DE ? 'Produkt und VSS-Support eingeben...' : 'Enter product and VSS support...'))
        @('11', ($DE ? 'Wie lange werden Backup-Daten aufbewahrt?' : 'How long will backup data be retained?'), 'q_backup_ret', ($DE ? 'Aufbewahrungsdauer eingeben...' : 'Enter retention period...'))
        @('12', ($DE ? 'Welche Monitoring-Lösung wird verwendet?' : 'Which monitoring solution will be used?'), 'q_monitoring', ($DE ? 'Monitoring-Produkt eingeben...' : 'Enter monitoring product...'))
        @('13', ($DE ? 'Wird Hybrid (M365) konfiguriert?' : 'Will Hybrid (M365) be configured?'), 'q_hybrid', ($DE ? 'Ja/Nein, Full/Minimal...' : 'Yes/No, Full/Minimal...'))
        @('14', ($DE ? 'Werden Public Folders eingesetzt?' : 'Will Public Folders be used?'), 'q_pf', ($DE ? 'Ja/Nein, ggf. Anzahl und Größe...' : 'Yes/No, count and size if applicable...'))
        @('15', ($DE ? 'Welche Anti-Spam-Lösung wird eingesetzt?' : 'Which anti-spam solution will be used?'), 'q_hygiene', ($DE ? 'Lösung eingeben...' : 'Enter solution...'))
        @('16', ($DE ? 'Sind Compliance-Anforderungen vorhanden? (Litigation Hold, Journaling, DSGVO)' : 'Are compliance requirements in place? (Litigation Hold, journaling, GDPR)'), 'q_compliance', ($DE ? 'Anforderungen eingeben...' : 'Enter requirements...'))
        @('17', ($DE ? 'Welche MDM/EMM-Lösung wird für mobile Endgeräte eingesetzt?' : 'Which MDM/EMM solution will be used for mobile devices?'), 'q_mdm', ($DE ? 'Lösung oder "keiner"...' : 'Solution or "none"...'))
        @('18', ($DE ? 'Welche SMTP-Relay-Subnetze senden anonym über Exchange?' : 'Which SMTP relay subnets send anonymously through Exchange?'), 'q_relay_subnets', ($DE ? 'Subnetze (CIDR) eingeben...' : 'Enter subnets (CIDR)...'))
        @('19', ($DE ? 'Gibt es einen ausgehenden Smarthost (Proxy für ausgehende E-Mail)?' : 'Is there an outbound smarthost (proxy for outbound email)?'), 'q_smarthost', ($DE ? 'FQDN/IP oder "keiner"...' : 'FQDN/IP or "none"...'))
        @('20', ($DE ? 'Welche Exchange-Organisation (Organisations-Name) wird eingesetzt?' : 'What Exchange organisation name will be used?'), 'q_org_name', ($DE ? 'Organisationsname eingeben...' : 'Enter organisation name...'))
        @('21', ($DE ? 'Sind Migration-Schritte erforderlich? (Kapitel 10 ausfüllen?)' : 'Are migration steps required? (Complete Chapter 10?)'), 'q_migration', ($DE ? 'Ja/Nein...' : 'Yes/No...'))
        @('22', ($DE ? 'Offene Punkte und offene Fragen (Freitext)' : 'Open items and questions (free text)'), 'q_open_items', ($DE ? 'Offene Punkte eingeben...' : 'Enter open items...'))
    )
    $null = $parts.Add((Add-WordQuestionnaireTable -ColNr 'Nr.' -ColFrage ($DE ? 'Frage' : 'Question') -ColAntwort ($DE ? 'Antwort' : 'Answer') -Questions $questions))

    # ── 16. Freigabeseite ──────────────────────────────────────────────────────────
    $null = $parts.Add((Add-WordHeading ($DE ? '16. Freigabeseite' : '16. Approval Page') 1))
    $null = $parts.Add((Add-WordParagraph ($DE ? 'Mit Unterzeichnung bestätigen die Beteiligten, dass sie das Konzept- und Freigabedokument geprüft haben und der Umsetzung zustimmen.' : 'By signing, the parties confirm that they have reviewed this concept and approval document and agree to proceed.')))
    $null = $parts.Add((Add-WordParagraph ''))
    $roles = if ($DE) { @('Ersteller', 'Prüfer', 'Freigeber', 'Kunde') } else { @('Author', 'Reviewer', 'Approver', 'Customer') }
    $labelName = if ($DE) { 'Name:' } else { 'Name:' }
    $labelDate = if ($DE) { 'Datum:' } else { 'Date:' }
    $labelSig  = if ($DE) { 'Unterschrift:' } else { 'Signature:' }
    $null = $parts.Add((Add-WordApprovalTable -Roles $roles -LabelName $labelName -LabelDate $labelDate -LabelSig $labelSig))

    $parts.ToArray()
}

# ── Main ────────────────────────────────────────────────────────────────────────

$repoRoot     = Split-Path $PSScriptRoot -Parent
$templatesDir = Join-Path $repoRoot 'templates'
if (-not (Test-Path $templatesDir)) { New-Item $templatesDir -ItemType Directory -Force | Out-Null }

Write-Host 'Generating Exchange-Konzept-Vorlage-DE.docx...'
$pathDE = Join-Path $templatesDir 'Exchange-Konzept-Vorlage-DE.docx'
New-WordDocument -OutputPath $pathDE `
    -BodyParts (Get-F23Parts 'DE') `
    -Title 'Exchange Server Konzept- und Freigabedokument' `
    -HeaderTitle 'EXCHANGE SERVER KONZEPT- UND FREIGABEDOKUMENT' `
    -Creator 'EXpress'
Write-Host "  -> $pathDE"

Write-Host 'Generating Exchange-Konzept-Vorlage-EN.docx...'
$pathEN = Join-Path $templatesDir 'Exchange-Konzept-Vorlage-EN.docx'
New-WordDocument -OutputPath $pathEN `
    -BodyParts (Get-F23Parts 'EN') `
    -Title 'Exchange Server Concept and Approval Document' `
    -HeaderTitle 'EXCHANGE SERVER CONCEPT AND APPROVAL DOCUMENT' `
    -Creator 'EXpress'
Write-Host "  -> $pathEN"

Write-Host ''
Write-Host 'Done. Open both files in Word/LibreOffice to verify structure.'
Write-Host 'Verify: headings, questionnaire content controls (Ch. 15), approval table (Ch. 16).'

} # end process
