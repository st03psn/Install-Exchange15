    function New-InstallationDocument {
        # Default to EN unless the state explicitly requests DE. Previous logic
        # ($State['Language'] -ne 'EN') flipped to DE whenever the key was missing
        # (e.g. Phase-5 re-entry against a pre-5.93 state file without the key).
        $DE   = ($State['Language'] -eq 'DE')
        $cust = [bool]$State['CustomerDocument']
        $lang = if ($DE) { 'DE' } else { 'EN' }
        $scope         = if ($State['DocumentScope']) { $State['DocumentScope'] } else { 'All' }
        $includeFilter = if ($State['IncludeServers']) { @($State['IncludeServers'] -split ',') } else { @() }
        $isAdHoc       = [bool]$State['StandaloneDocument'] -and -not $State['InstallPhase']
        $docStem  = if ($DE) { 'ExchangeServer-Dokumentation' } else { 'ExchangeServer-Documentation' }
        $docPath  = Join-Path $State['ReportsPath'] ('{0}_EXpress_{1}_{2}_{3}.docx' -f $env:COMPUTERNAME, $docStem, $lang, (Get-Date -Format 'yyyyMMdd-HHmmss'))
        $docTitle = if ($DE) { 'Exchange Server Installationsdokumentation' } else { 'Exchange Server Installation Documentation' }
        Write-MyStep -Label 'Word Document' -Value ('generating ({0})' -f $lang) -Status Run

        Write-MyVerbose 'Collecting installation report data'
        $rd = Get-InstallationReportData -Scope $scope -IncludeServers $includeFilter

        function Mask-Ip([string]$text) {
            if (-not $cust) { return $text }
            $text -replace '\b(10|172\.(1[6-9]|2[0-9]|3[01])|192\.168)\.\d{1,3}\.\d{1,3}\b', 'x.x.x.x'
        }
        function Mask-Val([string]$text) { if ($cust -and $text) { '[redacted]' } else { $text } }
        function SafeVal([object]$v, [string]$fallback = '') { if ($null -eq $v -or "$v" -eq '') { $fallback } else { "$v" } }
        # L / Lc: language helper. PS 5.1 cannot use (if ...) as a command argument; these helpers keep call sites compact.
        function L([string]$d, [string]$e) { if ($DE) { $d } else { $e } }
        function Lc([bool]$c, [string]$a, [string]$b) { if ($c) { $a } else { $b } }
        function Get-SecReg($path, $name) { try { (Get-ItemProperty -Path $path -Name $name -ErrorAction Stop).$name } catch { $null } }
        # Format-RegBool: translate registry 0/1 (or $false/$true) to localised enabled/disabled text.
        function Format-RegBool($v) {
            if ($null -eq $v -or "$v" -eq '') { return (L '(nicht gesetzt)' '(not set)') }
            # Use [bool] instead of [int]: Exchange cmdlet properties can return SwitchParameter,
            # which [int] cannot cast in PS 5.1 but [bool] handles via implicit conversion.
            if ([bool]$v) { return (L 'aktiviert' 'enabled') }
            return (L 'deaktiviert' 'disabled')
        }
        function Format-RemoteSysRows($remData) {
            $rows = [System.Collections.Generic.List[object[]]]::new()
            if (-not $remData -or -not $remData.Reachable) {
                $errMsg = if ($remData -and $remData.Error) { $remData.Error } else { (L 'WinRM nicht erreichbar' 'WinRM not reachable') }
                $rows.Add(@((L 'Systemdetails' 'System details'), (L ('Nicht abrufbar: {0} — Abhilfe: tools\Enable-EXpressRemoteQuery.ps1' -f $errMsg) ('Not available: {0} — Fix: tools\Enable-EXpressRemoteQuery.ps1' -f $errMsg))))
                return ,$rows
            }
            if ($remData.ComputerSys) {
                $cs = $remData.ComputerSys
                $hwType = if ($cs.Manufacturer -like '*VMware*' -or $cs.Model -like '*VMware*') { 'Virtual — VMware' }
                          elseif ($cs.Manufacturer -like '*Microsoft*' -and $cs.Model -like '*Virtual*') { 'Virtual — Hyper-V' }
                          elseif ($cs.Manufacturer -like '*QEMU*' -or $cs.Model -like '*KVM*') { 'Virtual — KVM/QEMU' }
                          elseif ($cs.HypervisorPresent) { (L 'Virtual — unbekannter Hypervisor' 'Virtual — unknown hypervisor') }
                          else { 'Physical ({0} {1})' -f $cs.Manufacturer.Trim(), $cs.Model.Trim() }
                $rows.Add(@((L 'Hardwaretyp' 'Hardware type'), $hwType))
                $rows.Add(@((L 'Computername (FQDN)' 'Computer name (FQDN)'), ('{0}.{1}' -f $cs.DNSHostName, $cs.Domain)))
            }
            if ($remData.OS) {
                $rows.Add(@((L 'Betriebssystem' 'Operating system'), $remData.OS.Caption))
                $rows.Add(@((L 'OS-Build' 'OS build'), $remData.OS.Version))
                $lastBoot = $remData.OS.LastBootUpTime
                $uptimeDays = if ($lastBoot) { [int][Math]::Floor(((Get-Date) - $lastBoot).TotalDays) } else { $null }
                $uptimeStr = if ($null -ne $uptimeDays) { '{0} — {1}d {2}' -f $lastBoot.ToString('yyyy-MM-dd HH:mm:ss'), $uptimeDays, (L 'Betriebszeit' 'uptime') } else { '?' }
                $rows.Add(@((L 'Letzter Neustart' 'Last boot'), $uptimeStr))
                $rows.Add(@((L 'RAM gesamt' 'Total RAM'), ('{0} GB' -f [math]::Round($remData.OS.TotalVisibleMemorySize / 1MB, 0))))
            }
            if ($remData.TimeZone) {
                $rows.Add(@((L 'Zeitzone' 'Time zone'), ('{0} ({1})' -f $remData.TimeZone.StandardName, $remData.TimeZone.Caption)))
            }
            if ($remData.CPU) {
                $cpuList = @($remData.CPU)
                $totalCores   = ($cpuList | Measure-Object NumberOfCores -Sum).Sum
                $totalLogical = ($cpuList | Measure-Object NumberOfLogicalProcessors -Sum).Sum
                $rows.Add(@('CPU', ('{0} — {1} {2} / {3} {4}' -f $cpuList[0].Name.Trim(), $totalCores, (L 'Kerne' 'cores'), $totalLogical, (L 'logisch' 'logical'))))
            }
            foreach ($vol in $remData.Volumes) {
                if ($vol.DriveLetter -and $vol.Capacity -gt 0) {
                    $freeGB = [math]::Round($vol.FreeSpace / 1GB, 1)
                    $totGB  = [math]::Round($vol.Capacity / 1GB, 1)
                    $pct    = [math]::Round($vol.FreeSpace / $vol.Capacity * 100, 0)
                    $au     = if ($vol.BlockSize) { '{0} KB' -f ($vol.BlockSize / 1KB) } else { '?' }
                    $rows.Add(@(('Volume {0}:' -f $vol.DriveLetter), ('{0} GB {1} / {2} GB ({3}% free) — AU: {4}' -f $freeGB, $vol.FileSystem, $totGB, $pct, $au)))
                }
            }
            if ($remData.PageFile) {
                $pf    = $remData.PageFile
                $ramMB = if ($remData.OS) { [math]::Round($remData.OS.TotalVisibleMemorySize / 1KB, 0) } else { 0 }
                $recMB = $ramMB + 10
                $rows.Add(@((L 'Auslagerungsdatei' 'Page file'), ('{0} — Init: {1} MB / Max: {2} MB — {3}: {4} MB' -f $pf.Name, $pf.InitialSize, $pf.MaximumSize, (L 'Empfehlung RAM+10MB' 'Recommended RAM+10MB'), $recMB)))
            }
            foreach ($nic in $remData.NICs) {
                $ips = if ($nic.IPAddress) { (Mask-Ip ($nic.IPAddress -join ', ')) } else { (L '(keine IP)' '(no IP)') }
                $dns = if ($nic.DNSServerSearchOrder) { (Mask-Ip ($nic.DNSServerSearchOrder -join ', ')) } else { (L '(nicht gesetzt)' '(not set)') }
                $rows.Add(@(('NIC: {0}' -f $nic.Description), ('{0} — DNS: {1}' -f $ips, $dns)))
            }
            if ($remData.NICDrivers -and $remData.NICDrivers.Count -gt 0) {
                foreach ($drv in $remData.NICDrivers) {
                    $drvDate = if ($drv.DriverDate) { try { ([datetime]$drv.DriverDate).ToString('yyyy-MM-dd') } catch { '?' } } else { '?' }
                    $speed   = if ($drv.LinkSpeed) { ('{0}' -f $drv.LinkSpeed) } else { '?' }
                    $rows.Add(@((L ('NIC-Treiber: {0}' -f $drv.Name) ('NIC driver: {0}' -f $drv.Name)), ('{0} — v{1} — {2} — {3}' -f $drv.InterfaceDescription, $drv.DriverVersion, $drvDate, $speed)))
                }
            }
            return ,$rows
        }

        $parts = [System.Collections.Generic.List[string]]::new()

        # ── Template check (F24) ─────────────────────────────────────────────────
        # When -TemplatePath is supplied and valid, the cover page is driven by the
        # template DOCX; $parts contains only the chapter body XML.
        $tplPath = $State['TemplatePath']
        $useTpl  = $tplPath -and (Test-Path $tplPath -PathType Leaf)
        if ($useTpl) {
            $tplCheck = Test-WdTemplate -Path $tplPath -RequiredTags @('document_body')
            if (-not $tplCheck.Valid) {
                Write-MyWarning ('Template missing required tokens: ' + ($tplCheck.Missing -join ', ') + ' — falling back to built-in cover page.')
                $useTpl = $false
            } else {
                Write-MyVerbose ('Using custom template: ' + $tplPath)
            }
        }

        $instMode = if ($isAdHoc) { (L 'Ad-hoc-Inventar' 'Ad-hoc Inventory') } elseif ($State['InstallEdge']) { 'Edge Transport' } elseif ($State['InstallRecipientManagement']) { 'Recipient Management Tools' } elseif ($State['InstallManagementTools']) { 'Management Tools' } elseif ($State['StandaloneOptimize']) { 'Standalone Optimize' } elseif ($State['NoSetup']) { 'Optimization Only' } else { 'Mailbox Server' }
        $scenario = if ($isAdHoc) { (L 'Ad-hoc-Inventar (vorhandene Umgebung)' 'Ad-hoc inventory (existing environment)') } elseif ($rd.Servers.Count -le 1) { (L 'Neue Exchange-Umgebung' 'New Exchange environment') } else { (L 'Server-Ergänzung zu bestehender Umgebung' 'Server added to existing environment') }
        $classification = (Lc $cust 'CUSTOMER' 'INTERN')

        # Cover page variables — needed both for built-in cover page and template tokens.
        $company  = SafeVal $State['CompanyName'] ''
        $author   = SafeVal $State['Author']      ''
        $coverSub = (L 'Installation, Hybridbereitstellung, Mailflow' 'Installation, Hybrid deployment, Mail flow')
        # Logo probe: sources\logo.png (user-placed) → beside the script → assets\logo.png (repo default)
        $_logoRoot = if ($PSScriptRoot) { $PSScriptRoot } else { $State['InstallPath'] }
        $logoFile = @(
            (Join-Path $State['SourcesPath'] 'logo.png'),
            (Join-Path $_logoRoot 'logo.png'),
            (Join-Path $_logoRoot 'assets\logo.png')
        ) | Where-Object { Test-Path $_ -PathType Leaf } | Select-Object -First 1
        if (-not $logoFile) { $logoFile = Join-Path $State['SourcesPath'] 'logo.png' }   # fallback path (will fail Test-Path gracefully)

        if (-not $useTpl) {
        # ── Deckblatt (Cover Page) ───────────────────────────────────────────────
        # Layout nach Referenzvorlage: Produkt (groß) / Titel (XXL) / Untertitel / Version+Datum+Autor.
        # Company/Author sind State-gesteuert ($State['CompanyName'], $State['Author']) ohne Default-Branding.
        $null = $parts.Add((New-WdSpacer 1440))
        if (Test-Path $logoFile -PathType Leaf) {
            # Logo centered, 6 cm wide (2160000 EMU), proportional height for 400×80 source: 432000 EMU
            $null = $parts.Add('<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="240"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="2160000" cy="432000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="logo"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="logo"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2160000" cy="432000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')
        }
        $null = $parts.Add((New-WdCentered -Text 'Microsoft Exchange Server SE' -SizeHalfPt 40 -Bold $true -Color '1F3864'))
        $null = $parts.Add(('<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t xml:space="preserve">{0}</w:t></w:r></w:p>' -f (Invoke-XmlEscape (L 'Installation & Konfiguration' 'Installation & Configuration'))))
        $null = $parts.Add(('<w:p><w:pPr><w:pStyle w:val="Subtitle"/></w:pPr><w:r><w:t xml:space="preserve">{0}</w:t></w:r></w:p>' -f (Invoke-XmlEscape $coverSub)))
        $null = $parts.Add((New-WdSpacer 1200))
        $orgLine = SafeVal $State['OrganizationName'] ''
        if ($orgLine) { $null = $parts.Add((New-WdCentered -Text $orgLine -SizeHalfPt 28 -Bold $true -Color '1F3864')) }
        $null = $parts.Add((New-WdCentered -Text $env:COMPUTERNAME -SizeHalfPt 24 -Color '404040'))
        $null = $parts.Add((New-WdCentered -Text $scenario          -SizeHalfPt 22 -Color '404040'))
        $null = $parts.Add((New-WdCentered -Text (('{0}: {1}' -f (L 'Installationsmodus' 'Installation mode'), $instMode)) -SizeHalfPt 22 -Color '404040'))
        $null = $parts.Add((New-WdSpacer 1440))
        $null = $parts.Add((New-WdCentered -Text (('{0}: {1}' -f (L 'Versionsnummer' 'Version'), ('{0} / EXpress v{1}' -f (Get-Date -Format 'yyyy-MM-dd'), $ScriptVersion))) -SizeHalfPt 22 -Color '404040'))
        $null = $parts.Add((New-WdCentered -Text (('{0}: {1}' -f (L 'Datum' 'Date'), (Get-Date -Format 'dd.MM.yyyy')))                       -SizeHalfPt 22 -Color '404040'))
        if ($author)  { $null = $parts.Add((New-WdCentered -Text (('{0}: {1}' -f (L 'Autor' 'Author'), $author))                              -SizeHalfPt 22 -Color '404040')) }
        if ($company) { $null = $parts.Add((New-WdCentered -Text $company                                                                      -SizeHalfPt 22 -Color '404040')) }
        $null = $parts.Add((New-WdSpacer 600))
        $null = $parts.Add((New-WdCentered -Text $classification -SizeHalfPt 22 -Bold $true -Color 'C00000'))
        $null = $parts.Add((New-WdPageBreak))
        } # end if (-not $useTpl)

        # ── Hinweise zu diesem Dokument ─────────────────────────────────────────
        # Struktur nach Referenzvorlage: Anpassungsvorbehalt, Genderhinweis, Warenzeichen,
        # Screenshots/Mockups, Copyright. Firmenname aus $State['CompanyName'] — kein Default.
        $companyRef = if ($company) { $company } else { (L 'der Hersteller dieses Dokuments' 'the publisher of this document') }
        $null = $parts.Add((New-WdHeading (L 'Hinweise zu diesem Dokument' 'Notes on this document') 1))
        $null = $parts.Add((New-WdParagraph (L ('{0} behält sich vor, den beschriebenen Funktionsumfang jederzeit an neue Anforderungen und Erkenntnisse anzupassen. Dadurch kann es gegebenenfalls zu Abweichungen zwischen diesem Dokument und der ausgelieferten Software kommen.' -f $companyRef) ('{0} reserves the right to adapt the functional scope described herein to new requirements and insights at any time. This may result in deviations between this document and the delivered software.' -f $companyRef))))
        $null = $parts.Add((New-WdParagraph (L 'Genderhinweis: Aus Gründen der besseren Lesbarkeit wird auf eine geschlechtsneutrale Differenzierung verzichtet. Entsprechende Begriffe gelten im Sinne der Gleichbehandlung grundsätzlich für alle Geschlechter. Die verkürzte Sprachform beinhaltet keine Wertung.' 'Gender note: For better readability, gender-neutral differentiation is omitted. Corresponding terms apply to all genders in the sense of equal treatment. The abbreviated language form does not imply a value judgement.')))
        $null = $parts.Add((New-WdParagraph (L 'Die hier genannten Produkte und Namen sind eingetragene Warenzeichen und/oder geschützte Marken und damit Eigentum der jeweiligen Rechteinhaber, u. a. der Microsoft Corporation (Microsoft, Exchange Server, Windows Server, Active Directory, Microsoft 365, Intune), Intel Corporation und weiterer.' 'The products and names mentioned here are registered trademarks and/or protected brands and therefore the property of the respective rights holders, including Microsoft Corporation (Microsoft, Exchange Server, Windows Server, Active Directory, Microsoft 365, Intune), Intel Corporation and others.')))
        $null = $parts.Add((New-WdParagraph (L 'Bitte beachten Sie: Teilweise zeigen dargestellte Ausgaben und Tabellen eine beispielhafte Konfiguration, um die beschriebenen Prozesse und Funktionalitäten zu dokumentieren. In Abstimmung mit dem Auftraggeber werden in der Vorbereitungsphase offene Fragen für die konkrete Umsetzung besprochen.' 'Please note: Some of the outputs and tables shown depict an exemplary configuration to document the described processes and functionality. Open questions regarding the concrete implementation are discussed with the contracting party during the preparation phase.')))
        $copyrightHolder = if ($company) { $company } else { (L '(Hersteller)' '(publisher)') }
        $null = $parts.Add((New-WdParagraph (L ('© Copyright {0}. Alle Rechte vorbehalten. Die Weitergabe und Vervielfältigung dieser Publikation oder von Teilen daraus sind, zu welchem Zweck und in welcher Form auch immer, ohne ausdrückliche schriftliche Genehmigung nicht gestattet. In dieser Publikation enthaltene Informationen können ohne vorherige Ankündigung geändert werden.' -f $copyrightHolder) ('© Copyright {0}. All rights reserved. Reproduction or distribution of this publication or parts thereof, for any purpose and in any form, is not permitted without express written approval. Information contained in this publication may be changed without prior notice.' -f $copyrightHolder))))
        $null = $parts.Add((New-WdParagraph (L 'Dieses Dokument wurde automatisch durch EXpress (Install-Exchange15.ps1) generiert und spiegelt die Konfiguration der Exchange-Umgebung zum Erstellungszeitpunkt wider. Spätere Änderungen sind nicht berücksichtigt. EXpress wird "wie besehen" ohne Gewährleistung bereitgestellt; die Verantwortung für die Einhaltung organisatorischer, rechtlicher sowie regulatorischer Vorgaben (z. B. DSGVO, GoBD, BAIT/VAIT, ISO 27001) liegt beim Betreiber.' 'This document was generated automatically by EXpress (Install-Exchange15.ps1) and reflects the Exchange environment configuration at the time of generation. Subsequent changes are not reflected. EXpress is provided "as is" without warranty; responsibility for compliance with organisational, legal and regulatory requirements (e.g. GDPR, SOX, ISO 27001) lies with the operator.')))
        $null = $parts.Add((New-WdHeading (L 'Versionshistorie' 'Revision History') 2))
        $revAuthor = if ($author) { $author } else { ('EXpress v{0}' -f $ScriptVersion) }
        $null = $parts.Add((New-WdTable -Headers @((L 'Version' 'Version'), (L 'Datum' 'Date'), (L 'Autor' 'Author'), (L 'Änderung' 'Change')) -Rows @(
            @('1.0', (Get-Date -Format 'dd.MM.yyyy'), $revAuthor, (L 'Automatische Erstgenerierung' 'Automatic initial generation'))
        )))

        # ── Dynamisches Inhaltsverzeichnis ───────────────────────────────────────
        $null = $parts.Add((New-WdToc (L 'Inhaltsverzeichnis' 'Table of Contents')))

        # ── 1. Dokumenteigenschaften ─────────────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '1. Dokumenteigenschaften' '1. Document Properties') 1))
        $null = $parts.Add((New-WdTable -Compact -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows @(
            @((L 'Dokument' 'Document'), $docTitle)
            @('EXpress Version', "v$ScriptVersion")
            @((L 'Erstellt auf Server' 'Generated on server'), $env:COMPUTERNAME)
            @((L 'Exchange-Organisation' 'Exchange Organisation'), (SafeVal $State['OrganizationName'] (L '(nicht gesetzt)' '(not set)')))
            @((L 'Szenario' 'Scenario'), $scenario)
            @((L 'Installationsmodus' 'Installation mode'), $instMode)
            @((L 'Installiert durch' 'Installed by'), (SafeVal $State['InstallingUser'] (L '(unbekannt)' '(unknown)')))
            @((L 'Erstellt am' 'Generated on'), (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
            @((L 'Klassifizierung' 'Classification'), $classification)
        )))

        # ── 1.1 Freigabe und Change-Management ───────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '1.1 Freigabe und Change-Management' '1.1 Sign-off and Change Management') 2))
        $null = $parts.Add((New-WdParagraph (L 'Die folgende Tabelle dient als formaler Freigabenachweis dieser Installation. Bitte nach Abschluss der Installation und Durchführung der Abnahmetests ausfüllen (siehe auch Kapitel 16 Abnahmetest).' 'The table below serves as formal sign-off evidence for this installation. Please complete after finishing the installation and acceptance tests (see also chapter 16 Acceptance Testing).')))
        $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows @(
            ,@((L 'Change-Request-Nr.' 'Change request no.'), '')
            ,@((L 'Genehmigt von' 'Approved by'), '')
            ,@((L 'Genehmigungsdatum' 'Approval date'), '')
            ,@((L 'Abnahme durch' 'Accepted by'), '')
            ,@((L 'Abnahmedatum' 'Acceptance date'), '')
            ,@((L 'Bemerkungen' 'Remarks'), '')
        )))

        # ── 2. Installationsparameter (nur bei tatsächlichem Setup-Lauf) ─────────
        if (-not $isAdHoc) {
            $null = $parts.Add((New-WdHeading (L '2. Installationsparameter' '2. Installation Parameters') 1))
            $null = $parts.Add((New-WdParagraph (L 'Die folgende Tabelle dokumentiert die bei der Installation verwendeten Parameter. Sie dient als Nachweis der gewählten Konfiguration und als Referenz für spätere Änderungen oder eine Neuinstallation. Im Autopilot-Modus wurden alle Parameter aus einer Konfigurationsdatei geladen; im Copilot-Modus wurden sie während des Installationslaufs interaktiv abgefragt.' 'The table below documents the parameters used during installation. It serves as evidence of the chosen configuration and as a reference for later changes or a reinstallation. In Autopilot mode, all parameters were loaded from a configuration file; in Copilot mode they were interactively collected during the installation run.')))
            $modeText = if ($State['ConfigDriven']) { (L 'Autopilot (vollautomatisch)' 'Autopilot (fully automated)') } else { (L 'Copilot (interaktiv)' 'Copilot (interactive)') }
            $paramRows = [System.Collections.Generic.List[object[]]]::new()
            $paramRows.Add(@((L 'Setup-Version' 'Setup version'), (SafeVal (& { try { (Get-SetupTextVersion $State['SetupVersion']) } catch { $State['SetupVersion'] } }))))
            $paramRows.Add(@((L 'Installationspfad' 'Install path'), (SafeVal $State['InstallPath'])))
            $paramRows.Add(@('Namespace',                        (SafeVal $State['Namespace']       '—')))
            $paramRows.Add(@('OWA Download Domain',              (SafeVal $State['DownloadDomain']  '—')))
            $paramRows.Add(@('DAG',                              (SafeVal $State['DAGName']         '—')))
            $paramRows.Add(@((L 'Zertifikatspfad' 'Certificate path'), (Mask-Val (SafeVal $State['CertificatePath'] '—'))))
            $logRet = if ($State['LogRetentionDays']) { '{0} {1}' -f $State['LogRetentionDays'], (L 'Tage' 'days') } else { '—' }
            $paramRows.Add(@((L 'Log-Aufbewahrung' 'Log retention'), $logRet))
            $relayStr = if ($State['RelaySubnets']) { Mask-Ip (($State['RelaySubnets'] -join ', ')) } else { '—' }
            $paramRows.Add(@((L 'Relay-Subnetze' 'Relay subnets'), $relayStr))
            $paramRows.Add(@((L 'Modus' 'Mode'), $modeText))
            $paramRows.Add(@((L 'Logdatei' 'Log file'), (SafeVal $State['TranscriptFile'])))
            $null = $parts.Add((New-WdTable -Headers @((L 'Parameter' 'Parameter'), (L 'Wert' 'Value')) -Rows $paramRows.ToArray()))
            $null = $parts.Add((New-WdParagraph (L 'TLS-Protokoll- und Schannel-Einstellungen (TLS 1.2/1.3, TLS 1.0/1.1, Cipher Suites) sind in Kapitel 8 dokumentiert.' 'TLS protocol and Schannel settings (TLS 1.2/1.3, TLS 1.0/1.1, cipher suites) are documented in Chapter 8.')))
        }

        # ── 3. IST-Aufnahme Active Directory ─────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '3. Active Directory — Voraussetzungen und Status' '3. Active Directory — Prerequisites and Status') 1))
        $null = $parts.Add((New-WdParagraph (L 'Exchange Server SE ist vollständig von Active Directory abhängig: Alle Konfigurationsdaten werden im AD gespeichert, die Authentifizierung erfolgt über Kerberos/NTLM gegen AD-Domänencontroller, und der Transport-Dienst nutzt AD-Standortinformationen für die Nachrichtenweiterleitung. Im Rahmen der Preflight-Prüfung wurden die AD-Voraussetzungen verifiziert. Die folgende Tabelle zeigt den ermittelten AD-Status zum Zeitpunkt der Installation.' 'Exchange Server SE is fully dependent on Active Directory: all configuration data is stored in AD, authentication is handled via Kerberos/NTLM against AD domain controllers, and the transport service uses AD site information for message routing. The AD prerequisites were verified during the preflight check. The table below shows the AD status at the time of installation.')))
        $adRows = [System.Collections.Generic.List[object[]]]::new()
        try { $localCS = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue; if ($localCS) { $adRows.Add(@((L 'Domäne' 'Domain'), $localCS.Domain)) } } catch { Write-MyVerbose ('Get-CimInstance Win32_ComputerSystem failed: {0}' -f $_) }
        try { $ffl = Get-ForestFunctionalLevel; $adRows.Add(@((L 'Forest Functional Level' 'Forest functional level'), ('{0} ({1})' -f $ffl, (Get-FFLText $ffl)))) } catch { Write-MyVerbose ('Get-ForestFunctionalLevel failed: {0}' -f $_) }
        try {
            $exOrg = Get-ExchangeOrganization
            if ($exOrg) { $adRows.Add(@((L 'Exchange-Organisation' 'Exchange organisation'), $exOrg)) }
            $adRows.Add(@((L 'Exchange Forest Schema (rangeUpper)' 'Exchange forest schema (rangeUpper)'), (SafeVal (Get-ExchangeForestLevel))))
            $adRows.Add(@((L 'Exchange Domain Level' 'Exchange domain level'), (SafeVal (Get-ExchangeDomainLevel))))
        } catch { Write-MyVerbose ('Get-ExchangeOrganization / Exchange schema level failed: {0}' -f $_) }
        try {
            $fsmoRoles = @{}
            $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            $fsmoRoles[(L 'Schema Master' 'Schema Master')] = $forest.SchemaRoleOwner.Name
            $fsmoRoles[(L 'Domain Naming Master' 'Domain Naming Master')] = $forest.NamingRoleOwner.Name
            $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $fsmoRoles[(L 'PDC Emulator' 'PDC Emulator')] = $domain.PdcRoleOwner.Name
            $fsmoRoles[(L 'RID Master' 'RID Master')] = $domain.RidRoleOwner.Name
            $fsmoRoles[(L 'Infrastructure Master' 'Infrastructure Master')] = $domain.InfrastructureRoleOwner.Name
            foreach ($role in $fsmoRoles.Keys) { $adRows.Add(@($role, (Mask-Ip $fsmoRoles[$role]))) }
        } catch { Write-MyVerbose ('FSMO role query failed: {0}' -f $_) }
        if ($adRows.Count -eq 0) { $adRows.Add(@((L '(AD-Daten nicht abrufbar)' '(AD data not available)'), '')) }
        $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $adRows.ToArray()))
        $null = $parts.Add((New-WdParagraph (L 'Hinweis: Ein Forest Functional Level von mindestens Windows Server 2012 R2 (Level 6) ist für Exchange SE erforderlich. Schema- und Domänenerweiterungen (PrepareSchema / PrepareAD / PrepareDomain) wurden von EXpress automatisch durchgeführt.' 'Note: A Forest Functional Level of at least Windows Server 2012 R2 (Level 6) is required for Exchange SE. Schema and domain extensions (PrepareSchema / PrepareAD / PrepareDomain) were performed automatically by EXpress.')))

        # ── 4. Organisation — übergreifende Konfiguration ────────────────────────
        $orgD = $rd.Org
        if ($scope -in 'All','Org') {
            Invoke-DocSection-Organisation -Parts $parts -ReportData $rd -DE $DE -Cust $cust
        }

        # ── 5. Server in der Organisation ────────────────────────────────────────
        if ($scope -in 'All','Local') {
            $null = $parts.Add((New-WdHeading (L '5. Server in der Organisation' '5. Servers in the Organisation') 1))
            $null = $parts.Add((New-WdParagraph (L 'Die folgenden Abschnitte dokumentieren jeden Exchange-Server in der Organisation. Der neu installierte Server ist mit dem Hinweis "← Neu installiert" gekennzeichnet. Systemdetails (Hardware, Volumes, NICs) werden über WinRM/CIM abgefragt — bei nicht erreichbaren Servern erscheint ein entsprechender Hinweis.' 'The following sections document each Exchange server in the organisation. The newly installed server is marked with "← Newly installed". System details (hardware, volumes, NICs) are retrieved via WinRM/CIM — for unreachable servers a corresponding note is shown.')))
            if ($rd.Servers.Count -eq 0) {
                $null = $parts.Add((New-WdParagraph (L '(Keine Exchange-Server abfragbar)' '(No Exchange servers available)')))
            }
            $srvCounter = 0
            foreach ($srvD in $rd.Servers) {
                $srvCounter++
                $srvName   = $srvD.ServerName
                $isLocal   = $srvD.IsLocalServer
                $exSrv2    = $srvD.ExServer
                $srvLabel  = if ($isLocal) { ('{0} ← {1}' -f $srvName, (L 'Neu installiert / lokaler Server' 'Newly installed / local server')) } else { $srvName }
                $null = $parts.Add((New-WdHeading ('5.{0} {1}' -f $srvCounter, $srvLabel) 2))

                # 5.x.1 Identität
                $null = $parts.Add((New-WdHeading (L 'Identität' 'Identity') 3))
                $idRows = [System.Collections.Generic.List[object[]]]::new()
                if ($exSrv2) {
                    $exReadable2 = try { Get-SetupTextVersion $State['SetupVersion'] } catch { '' }
                    $exVerStr2 = if ($exReadable2) { '{0} ({1})' -f $exSrv2.AdminDisplayVersion.ToString(), $exReadable2 } else { $exSrv2.AdminDisplayVersion.ToString() }
                    $idRows.Add(@((L 'Exchange-Version' 'Exchange version'), $exVerStr2))
                    $idRows.Add(@('FQDN', (SafeVal $exSrv2.Fqdn)))
                    $idRows.Add(@((L 'Serverrolle' 'Server role'), ($exSrv2.ServerRole -join ', ')))
                    $editionStr2 = $exSrv2.Edition.ToString()
                    if ($editionStr2 -like '*Evaluation*' -or $editionStr2 -like '*Trial*') { $editionStr2 += (L ' — ACHTUNG: Testlizenz, vor Produktiveinsatz in Standardlizenz umwandeln' ' — WARNING: evaluation licence, convert to standard before production use') }
                    $idRows.Add(@((L 'Edition' 'Edition'), $editionStr2))
                    $idRows.Add(@((L 'AD-Standort' 'AD site'), $exSrv2.Site.ToString()))
                    $idRows.Add(@((L 'Installiert am' 'Installed on'), (SafeVal $exSrv2.WhenCreated)))
                }
                if ($srvD.AutodiscoverSCP) { $idRows.Add(@('Autodiscover SCP', (SafeVal $srvD.AutodiscoverSCP.AutoDiscoverServiceInternalUri))) }
                $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $idRows.ToArray()))

                # 5.x.2 Systemdetails (CIM)
                $null = $parts.Add((New-WdHeading (L 'Systemdetails' 'System Details') 3))
                $sysDetailRows = Format-RemoteSysRows $srvD.RemoteData
                $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $sysDetailRows.ToArray()))

                # 5.x.3 Datenbanken
                $null = $parts.Add((New-WdHeading (L 'Postfachdatenbanken' 'Mailbox Databases') 3))
                $dbRows2 = [System.Collections.Generic.List[object[]]]::new()
                foreach ($db2 in $srvD.Databases) {
                    $mounted2 = if ($null -ne $db2.Mounted) { if ($db2.Mounted) { (L 'Eingehängt' 'Mounted') } else { (L 'Ausgehängt' 'Dismounted') } } else { (L 'Unbekannt' 'Unknown') }
                    $dbRows2.Add(@($db2.Name, (SafeVal $db2.EdbFilePath), (SafeVal $db2.LogFolderPath), $mounted2))
                }
                if ($dbRows2.Count -eq 0) { $dbRows2.Add(@((L '(keine Datenbank auf diesem Server)' '(no database on this server)'), '', '', '')) }
                $null = $parts.Add((New-WdTable -Headers @((L 'Datenbank' 'Database'), (L 'DB-Pfad' 'DB path'), (L 'Log-Pfad' 'Log path'), (L 'Status' 'Status')) -Rows $dbRows2.ToArray()))

                # 5.x.4 Virtuelle Verzeichnisse → konsolidierte Übersicht in Abschnitt 4.13
                # Virtual directories → consolidated view in section 4.13

                # 5.x.5 Receive Connectors — split into two tables (network / security)
                # A single 8-column table wraps every cell in portrait Word; splitting into
                # 4 + 5 columns keeps each row legible. Name repeats as the join key.
                $null = $parts.Add((New-WdHeading (L 'Receive Connectors' 'Receive Connectors') 3))
                $rcNetRows = [System.Collections.Generic.List[object[]]]::new()
                $rcSecRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($rc in $srvD.ReceiveConnectors) {
                    $reqTlsRc = Lc ([bool]$rc.RequireTLS) (L 'ja' 'yes') (L 'nein' 'no')
                    $maxMsgRc = if ($rc.MaxMessageSize) { $rc.MaxMessageSize.ToString() } else { '—' }
                    $rcNetRows.Add(@($rc.Name, (Mask-Ip ($rc.Bindings -join ', ')), (Mask-Ip ($rc.RemoteIPRanges -join ', ')), (SafeVal $rc.Fqdn '—')))
                    $rcSecRows.Add(@($rc.Name, $rc.AuthMechanism, $rc.PermissionGroups, $reqTlsRc, $maxMsgRc))
                }
                if ($rcNetRows.Count -eq 0) {
                    $rcNetRows.Add(@((L '(keine)' '(none)'), '', '', ''))
                    $rcSecRows.Add(@((L '(keine)' '(none)'), '', '', '', ''))
                }
                $null = $parts.Add((New-WdParagraph (L 'Netzwerk:' 'Network:')))
                $null = $parts.Add((New-WdTable -Compact -Headers @((L 'Connector' 'Connector'), 'Bindings', (L 'Remote-IPs' 'Remote IPs'), 'FQDN') -Rows $rcNetRows.ToArray()))
                $null = $parts.Add((New-WdParagraph (L 'Sicherheit und Limits:' 'Security and limits:')))
                $null = $parts.Add((New-WdTable -Compact -Headers @((L 'Connector' 'Connector'), 'Auth', (L 'Berechtigungen' 'Permissions'), 'TLS', (L 'Max. Größe' 'Max size')) -Rows $rcSecRows.ToArray()))

                # 5.x.6 IMAP/POP3-Konfiguration (nur lokaler Server)
                if ($srvD.IsLocalServer -and ($srvD.ImapSettings -or $srvD.PopSettings)) {
                    $null = $parts.Add((New-WdHeading (L 'IMAP/POP3-Konfiguration' 'IMAP/POP3 Configuration') 3))
                    $protoSrvRows = [System.Collections.Generic.List[object[]]]::new()
                    if ($srvD.ImapSettings) {
                        $im = $srvD.ImapSettings
                        $protoSrvRows.Add(@((L 'IMAP4 — Externer Namespace' 'IMAP4 — External namespace'),      (SafeVal (($im.ExternalConnectionSettings | ForEach-Object { $_.ToString() }) -join '; ') (L '(nicht gesetzt — bitte manuell ergänzen)' '(not set — please fill in manually)'))))
                        $protoSrvRows.Add(@((L 'IMAP4 — Interner Namespace' 'IMAP4 — Internal namespace'),      (SafeVal (($im.InternalConnectionSettings | ForEach-Object { $_.ToString() }) -join '; ') (L '(nicht gesetzt)' '(not set)'))))
                        $protoSrvRows.Add(@((L 'IMAP4 — X.509-Zertifikatname' 'IMAP4 — X.509 certificate name'), (SafeVal $im.X509CertificateName (L '(nicht gesetzt)' '(not set)'))))
                        $protoSrvRows.Add(@((L 'IMAP4 — Anmeldetyp' 'IMAP4 — Login type'),                       (SafeVal $im.LoginType)))
                    }
                    if ($srvD.PopSettings) {
                        $pop = $srvD.PopSettings
                        $protoSrvRows.Add(@((L 'POP3 — Externer Namespace' 'POP3 — External namespace'),         (SafeVal (($pop.ExternalConnectionSettings | ForEach-Object { $_.ToString() }) -join '; ') (L '(nicht gesetzt — bitte manuell ergänzen)' '(not set — please fill in manually)'))))
                        $protoSrvRows.Add(@((L 'POP3 — Interner Namespace' 'POP3 — Internal namespace'),         (SafeVal (($pop.InternalConnectionSettings | ForEach-Object { $_.ToString() }) -join '; ') (L '(nicht gesetzt)' '(not set)'))))
                        $protoSrvRows.Add(@((L 'POP3 — X.509-Zertifikatname' 'POP3 — X.509 certificate name'),  (SafeVal $pop.X509CertificateName (L '(nicht gesetzt)' '(not set)'))))
                    }
                    $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $protoSrvRows.ToArray()))
                }

                # 5.x.7 Zertifikate
                $null = $parts.Add((New-WdHeading (L 'Zertifikate' 'Certificates') 3))
                $certRows2 = [System.Collections.Generic.List[object[]]]::new()
                foreach ($cert2 in $srvD.Certificates) {
                    # phantom certs already filtered in Get-ServerReportData; belt-and-suspenders guard
                    if (-not $cert2.Thumbprint -or $cert2.NotAfter -le [datetime]'1970-01-01') { continue }
                    $expiry2   = $cert2.NotAfter.ToString('yyyy-MM-dd')
                    $daysLeft2 = [int][Math]::Floor(($cert2.NotAfter - (Get-Date)).TotalDays)
                    $tp2 = if ($cust) { ('{0}...' -f $cert2.Thumbprint.Substring(0, [Math]::Min(8, $cert2.Thumbprint.Length))) } else { $cert2.Thumbprint }
                    $certRows2.Add(@($cert2.Subject, $expiry2, ('{0}d' -f $daysLeft2), $cert2.Services, $tp2))
                }
                if ($certRows2.Count -eq 0) { $certRows2.Add(@((L '(keine)' '(none)'), '', '', '', '')) }
                $null = $parts.Add((New-WdTable -Compact -Headers @('Subject', (L 'Ablauf' 'Expiry'), (L 'Verbleibend' 'Remaining'), (L 'Dienste' 'Services'), (L 'Fingerabdruck' 'Thumbprint')) -Rows $certRows2.ToArray()))

                # Transport Agents are documented in section 9 (Anti-Spam) — not repeated here.
            }
        }

        # ── 6. Netzwerk & DNS (lokal) ─────────────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '6. Netzwerk und DNS (lokaler Server)' '6. Network and DNS (local server)') 1))
        $null = $parts.Add((New-WdParagraph (L 'Die folgende Tabelle zeigt die Netzwerkkonfiguration des lokalen Exchange-Servers. Für Exchange Server ist eine korrekte DNS-Auflösung (Forward und Reverse) eine grundlegende Betriebsvoraussetzung. Als DNS-Server müssen ausschließlich Active-Directory-integrierte DNS-Server der eigenen Domäne eingetragen sein — kein öffentlicher DNS (z. B. 8.8.8.8), da Exchange für Autodiscover, SCP-Lookups und interne Namensauflösung auf AD-DNS angewiesen ist.' 'The table below shows the network configuration of the local Exchange server. Correct DNS resolution (forward and reverse) is a fundamental operational requirement for Exchange Server. Only Active Directory-integrated DNS servers of the own domain must be configured — no public DNS (e.g. 8.8.8.8), as Exchange relies on AD DNS for Autodiscover, SCP lookups and internal name resolution.')))
        $netRows = [System.Collections.Generic.List[object[]]]::new()
        try {
            $nicIPs = @{}; $nicDNS = @{}
            Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.InterfaceAlias -notlike '*Loopback*' } | ForEach-Object { $nicIPs[$_.InterfaceAlias] = ('{0}/{1}' -f $_.IPAddress, $_.PrefixLength) }
            Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.InterfaceAlias -notlike '*Loopback*' -and $_.ServerAddresses } | ForEach-Object { $nicDNS[$_.InterfaceAlias] = ($_.ServerAddresses -join ', ') }
            foreach ($nic in ($nicIPs.Keys | Sort-Object)) {
                $ip2  = Mask-Ip $nicIPs[$nic]
                $dns2 = if ($nicDNS[$nic]) { Mask-Ip $nicDNS[$nic] } else { (L '(nicht gesetzt)' '(not set)') }
                $netRows.Add(@(('NIC: {0}' -f $nic), ('{0} — DNS: {1}' -f $ip2, $dns2)))
            }
        } catch { Write-MyVerbose ('NIC / DNS query failed: {0}' -f $_) }
        if ($netRows.Count -eq 0) { $netRows.Add(@((L '(keine NIC-Daten abrufbar)' '(no NIC data available)'), '')) }
        $null = $parts.Add((New-WdTable -Headers @((L 'NIC / Eigenschaft' 'NIC / Property'), (L 'Wert' 'Value')) -Rows $netRows.ToArray()))

        # 6.1 DNS-Einträge (relevant für Exchange-Dienste)
        $null = $parts.Add((New-WdHeading (L '6.1 DNS-Einträge (Exchange-Dienste)' '6.1 DNS Records (Exchange services)') 2))
        $null = $parts.Add((New-WdParagraph (L 'Für einen Exchange-Server müssen grundsätzlich die folgenden öffentlichen DNS-Einträge je SMTP-Domäne auf dem für diese Domäne zuständigen DNS eingetragen sein: autodiscover.<domain> (A oder CNAME auf den externen Namespace), MX (zeigt auf den eingehenden Mailflow — direkt auf den Exchange-Namespace, einen Smarthost oder einen eingehenden Cloud-Filter), sowie die Authentifizierungseinträge SPF (TXT), DKIM (TXT via Selektor) und DMARC (_dmarc.<domain> TXT). Bei Hybrid-Szenarien kommt ein CNAME auf onmicrosoft.com hinzu, außerdem ggf. _autodiscover._tcp SRV.' 'For an Exchange server the following public DNS records must exist per SMTP domain on the DNS authoritative for that domain: autodiscover.<domain> (A or CNAME pointing to the external namespace), MX (controls incoming mail flow — directly to the Exchange namespace, to a smart host, or to an inbound cloud filter), and the authentication records SPF (TXT), DKIM (TXT via selector) and DMARC (_dmarc.<domain> TXT). In hybrid scenarios a CNAME to onmicrosoft.com is added, plus optionally _autodiscover._tcp SRV.')))
        $null = $parts.Add((New-WdParagraph (L 'Hinweis: In Split-DNS-Szenarien (AD-Domäne entspricht einer gerouteten SMTP-Domäne) existieren diese Einträge zusätzlich auf dem internen AD-DNS; für rein interne AD-Domänen (z. B. .local/.lan) sind MX/SPF/DKIM/DMARC nicht relevant. Eine automatische Auflösung aller Einträge aus dem Server heraus ist nicht aussagekräftig, da die Antworten je nach DNS-View (intern/extern) abweichen und sich externe Einträge typischerweise erst nach Umzug der primären Maildomäne bzw. mit weiteren akzeptierten Domänen ergänzen.' 'Note: In split-DNS scenarios (AD domain identical to a routed SMTP domain) these records also exist on the internal AD DNS; for purely internal AD domains (e.g. .local/.lan) MX/SPF/DKIM/DMARC are not relevant. Automatic resolution of all records from the server itself is not conclusive, since answers differ depending on the DNS view (internal/external), and external records are typically added only after cut-over of the primary mail domain or when additional accepted domains are configured.')))

        # Autodiscover SCP (internal clients) — always sensible to document for a fresh server
        $scpRows = [System.Collections.Generic.List[object[]]]::new()
        try {
            $casList = Get-ClientAccessService -ErrorAction SilentlyContinue
            foreach ($cas in $casList) {
                $scpRows.Add(@($cas.Name, (SafeVal ([string]$cas.AutoDiscoverServiceInternalUri))))
            }
        } catch { Write-MyVerbose ('Get-ClientAccessService (SCP) failed: {0}' -f $_) }
        if ($scpRows.Count -gt 0) {
            $null = $parts.Add((New-WdParagraph (L 'Autodiscover Service Connection Point (SCP) — für domänenmitgliedschaftsfähige Clients im internen Netzwerk maßgeblich. Wird im AD (CN=Configuration) gespeichert und von Outlook bevorzugt vor DNS-basiertem Autodiscover verwendet.' 'Autodiscover Service Connection Point (SCP) — decisive for domain-joined clients on the internal network. Stored in AD (CN=Configuration) and preferred by Outlook over DNS-based autodiscover.')))
            $null = $parts.Add((New-WdTable -Headers @((L 'Client Access Server' 'Client Access server'), 'AutoDiscoverServiceInternalUri') -Rows $scpRows.ToArray()))
        }

        # DNS record template — Autodiscover pre-filled with configured namespace; MX/SPF/DKIM/DMARC
        # resolved via Resolve-DnsName (best-effort; internal DNS view may differ from external).
        # Domain list: prefer the mail domain (namespace parent or MailDomain) over the AD domain,
        # and skip non-routable suffixes (.local/.lan/.internal/.corp) that never need public DNS records.
        $dnsTemplateRows = [System.Collections.Generic.List[object[]]]::new()
        $nsForDns = if ($State['Namespace']) { $State['Namespace'] } else { $null }
        # Compute mail domain (same logic as Enable-AccessNamespaceMailConfig)
        $mailDomainForDns = if ($State['MailDomain']) {
            $State['MailDomain']
        } elseif ($nsForDns) {
            $nsPart = ($nsForDns -split '\.', 2)[1]
            if ($nsPart -match '\.') { $nsPart } else { $nsForDns }
        } else { $null }
        $authDomainNames = [System.Collections.Generic.List[string]]::new()
        if ($mailDomainForDns) { $authDomainNames.Add($mailDomainForDns) }
        if ($rd.Org -and $rd.Org.AcceptedDomains) {
            $nonRoutable = @('\.local$','\.lan$','\.internal$','\.corp$','\.home$','\.intranet$')
            $rd.Org.AcceptedDomains | Where-Object { $_.DomainType -eq 'Authoritative' } | Select-Object -ExpandProperty DomainName | ForEach-Object {
                $dn = [string]$_
                $isNonRoutable = $nonRoutable | Where-Object { $dn -match $_ }
                if (-not $isNonRoutable -and $dn -ne $mailDomainForDns -and $authDomainNames.Count -lt 5) {
                    $authDomainNames.Add($dn)
                }
            }
        }
        if ($authDomainNames.Count -eq 0) { $authDomainNames.Add('<domain>') }
        $dnsManual  = L '(bitte manuell ergänzen)' '(please fill in manually)'
        $dnsInvalid = L '(nicht auflösbar — bitte manuell ergänzen)' '(not resolvable — please fill in manually)'
        foreach ($d in $authDomainNames) {
            $d = [string]$d
            # Autodiscover A/CNAME: points to the configured namespace
            $autodisVal = if ($nsForDns) { $nsForDns } else { $dnsManual }
            $dnsTemplateRows.Add(@('A / CNAME', "autodiscover.$d", $autodisVal))
            # MX
            $mxVal = try {
                $r = @(Resolve-DnsName -Name $d -Type MX -ErrorAction Stop | Where-Object { $_.Type -eq 'MX' } | Sort-Object Preference)
                if ($r) { ($r | ForEach-Object { "$($_.Preference) $($_.NameExchange)" }) -join '; ' } else { '' }
            } catch { '' }
            $mxRow = if ($mxVal) { $mxVal } else { $dnsManual }
            $dnsTemplateRows.Add(@('MX', $d, $mxRow))
            # SPF
            $spfVal = try {
                $r = @(Resolve-DnsName -Name $d -Type TXT -ErrorAction Stop | Where-Object { $_.Type -eq 'TXT' -and ($_.Strings -join '') -match 'v=spf' })
                if ($r) { ($r | ForEach-Object { $_.Strings -join '' }) -join '; ' } else { '' }
            } catch { '' }
            $spfRow = if ($spfVal) { $spfVal } else { $dnsManual }
            $dnsTemplateRows.Add(@('TXT (SPF)', $d, $spfRow))
            # DKIM
            $dkimName = "selector1._domainkey.$d"
            $dkimVal = try {
                $r = @(Resolve-DnsName -Name $dkimName -Type TXT -ErrorAction Stop | Where-Object { $_.Type -eq 'TXT' })
                if ($r) { ($r | ForEach-Object { $_.Strings -join '' }) -join '; ' } else { '' }
            } catch { '' }
            $dkimRow = if ($dkimVal) { $dkimVal } else { $dnsManual }
            $dnsTemplateRows.Add(@('TXT (DKIM)', $dkimName, $dkimRow))
            # DMARC
            $dmarcName = "_dmarc.$d"
            $dmarcVal = try {
                $r = @(Resolve-DnsName -Name $dmarcName -Type TXT -ErrorAction Stop | Where-Object { $_.Type -eq 'TXT' -and ($_.Strings -join '') -match 'v=DMARC' })
                if ($r) { ($r | ForEach-Object { $_.Strings -join '' }) -join '; ' } else { '' }
            } catch { '' }
            $dmarcRow = if ($dmarcVal) { $dmarcVal } else { $dnsManual }
            $dnsTemplateRows.Add(@('TXT (DMARC)', $dmarcName, $dmarcRow))
        }
        $null = $parts.Add((New-WdParagraph (L 'Externe DNS-Einträge sind nach Go-Live über den autoritativen öffentlichen DNS zu prüfen (z. B. mxtoolbox.com, dig, nslookup). Die folgende Tabelle zeigt die typischerweise erforderlichen Einträge — bitte nach Einrichtung manuell ergänzen.' 'External DNS records must be verified after go-live via the authoritative public DNS (e.g. mxtoolbox.com, dig, nslookup). The table below lists the typically required records — please fill in after setup.')))
        $null = $parts.Add((New-WdTable -Headers @('Type', (L 'Name' 'Name'), (L 'Wert / Antwort' 'Value / Answer')) -Rows $dnsTemplateRows.ToArray()))

        # 6.2 Erforderliche Ports und Firewall-Regeln
        $null = $parts.Add((New-WdHeading (L '6.2 Erforderliche Ports und Firewall-Regeln' '6.2 Required Ports and Firewall Rules') 2))
        $null = $parts.Add((New-WdParagraph (L 'Die folgende Tabelle listet die für den Exchange Server-Betrieb erforderlichen TCP-Ports auf. Externe Ports müssen durch eine Firewall oder einen Reverse-Proxy abgesichert werden — Exchange Server sollte niemals direkt aus dem Internet erreichbar sein.' 'The table below lists the TCP ports required for Exchange Server operation. External ports must be secured by a firewall or reverse proxy — Exchange Server should never be directly reachable from the internet.')))
        $null = $parts.Add((New-WdTable -Headers @('Port', (L 'Protokoll' 'Protocol'), (L 'Dienst / Verwendung' 'Service / Purpose'), (L 'Sichtbarkeit' 'Visibility')) -Rows @(
            ,@('25',    'TCP', (L 'SMTP eingehend (extern + intern)' 'SMTP inbound (external + internal)'),                                               (L 'extern + intern' 'external + internal'))
            ,@('587',   'TCP', (L 'SMTP Submission / AUTH (Client-Einlieferung)' 'SMTP Submission / AUTH (client submission)'),                             (L 'intern / auth. Clients' 'internal / auth. clients'))
            ,@('443',   'TCP', (L 'HTTPS: OWA, ECP, EWS, Autodiscover, MAPI/HTTP, ActiveSync, OAB' 'HTTPS: OWA, ECP, EWS, Autodiscover, MAPI/HTTP, ActiveSync, OAB'), (L 'extern + intern' 'external + internal'))
            ,@('80',    'TCP', (L 'HTTP — Redirect auf HTTPS (am Reverse-Proxy)' 'HTTP — redirect to HTTPS (at reverse proxy)'),                           (L 'extern (Redirect)' 'external (redirect)'))
            ,@('993',   'TCP', (L 'IMAP4S (wenn aktiviert)' 'IMAP4S (if enabled)'),                                                                        (L 'intern / optional' 'internal / optional'))
            ,@('995',   'TCP', (L 'POP3S (wenn aktiviert)' 'POP3S (if enabled)'),                                                                          (L 'intern / optional' 'internal / optional'))
            ,@('135',   'TCP', (L 'RPC Endpoint Mapper (MAPI/RPC Legacy)' 'RPC Endpoint Mapper (MAPI/RPC legacy)'),                                        (L 'intern' 'internal'))
            ,@('445',   'TCP', (L 'SMB — DAG-Cluster, File Share Witness' 'SMB — DAG cluster, File Share Witness'),                                        (L 'intern (DAG)' 'internal (DAG)'))
            ,@('3268',  'TCP', (L 'Global Catalog LDAP' 'Global Catalog LDAP'),                                                                            (L 'intern (AD)' 'internal (AD)'))
            ,@('3269',  'TCP', (L 'Global Catalog LDAPS' 'Global Catalog LDAPS'),                                                                          (L 'intern (AD)' 'internal (AD)'))
            ,@('5985',  'TCP', (L 'WinRM HTTP (EMS, EXpress)' 'WinRM HTTP (EMS, EXpress)'),                                                               (L 'intern' 'internal'))
            ,@('5986',  'TCP', (L 'WinRM HTTPS (EMS, EXpress)' 'WinRM HTTPS (EMS, EXpress)'),                                                             (L 'intern' 'internal'))
            ,@('64327', 'TCP', (L 'DAG-Replikation (Mailbox Replication Service)' 'DAG Replication (Mailbox Replication Service)'),                        (L 'intern (DAG)' 'internal (DAG)'))
        ) -Compact))
        $null = $parts.Add((New-WdParagraph (L 'Hinweis: IMAP4 und POP3 sind auf Exchange Server standardmäßig deaktiviert und sollten nur bei explizitem Bedarf aktiviert werden. Port 80 (HTTP) sollte am Reverse-Proxy ausschließlich auf HTTPS (443) umgeleitet werden — Exchange-Dienste dürfen nicht unverschlüsselt exponiert sein.' 'Note: IMAP4 and POP3 are disabled by default on Exchange Server and should only be enabled when explicitly required. Port 80 (HTTP) should be redirected to HTTPS (443) at the reverse proxy — Exchange services must not be exposed unencrypted.')))

        # ── 7. Exchange-Installation (lokal, nur wenn kein Ad-hoc) ────────────────
        if (-not $isAdHoc) {
            $null = $parts.Add((New-WdHeading (L '7. Exchange-Installation (lokal)' '7. Exchange Installation (local)') 1))
            $null = $parts.Add((New-WdParagraph (L 'Die Exchange Server-Installation wurde mit EXpress vollautomatisch (Autopilot) bzw. interaktiv (Copilot) durchgeführt. EXpress übernimmt alle Installationsphasen 0–6 inklusive Windows-Features, .NET, VC++, URL Rewrite, UCMA, Active-Directory-Vorbereitung (PrepareSchema/PrepareAD), Exchange-Setup, Sicherheitshärtung und Post-Konfiguration. Die folgende Tabelle dokumentiert die installierte Exchange-Instanz auf diesem Server.' 'The Exchange Server installation was performed fully automated (Autopilot) or interactively (Copilot) using EXpress. EXpress handles all installation phases 0–6 including Windows features, .NET, VC++, URL Rewrite, UCMA, Active Directory preparation (PrepareSchema/PrepareAD), Exchange setup, security hardening and post-configuration. The table below documents the installed Exchange instance on this server.')))
            $exInstRows2 = [System.Collections.Generic.List[object[]]]::new()
            try {
                $exSrvLocal = Get-ExchangeServer $env:COMPUTERNAME -ErrorAction Stop
                $exReadableInst = try { Get-SetupTextVersion $State['SetupVersion'] } catch { '' }
                $exVerStrInst = if ($exReadableInst) { '{0} ({1})' -f $exSrvLocal.AdminDisplayVersion.ToString(), $exReadableInst } else { $exSrvLocal.AdminDisplayVersion.ToString() }
                $exInstRows2.Add(@((L 'Exchange-Version' 'Exchange version'), $exVerStrInst))
                $exInstRows2.Add(@((L 'Serverrolle' 'Server role'), ($exSrvLocal.ServerRole -join ', ')))
                $editionInst = $exSrvLocal.Edition.ToString()
                if ($editionInst -like '*Evaluation*' -or $editionInst -like '*Trial*') { $editionInst += (L ' — ACHTUNG: Testlizenz' ' — WARNING: evaluation licence') }
                $exInstRows2.Add(@((L 'Edition' 'Edition'), $editionInst))
                $exInstRows2.Add(@((L 'AD-Standort' 'AD site'), $exSrvLocal.Site.ToString()))
            } catch { Write-MyVerbose ('Get-ExchangeServer (local install info) failed: {0}' -f $_) }
            $exInstRows2.Add(@((L 'Installationspfad' 'Install path'), (SafeVal $State['InstallPath'])))
            $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $exInstRows2.ToArray()))

            # 7.1 Geplante Tasks (MEAC + Log Cleanup) — operative Exchange-Aufgaben, keine OS-Härtungen
            if ($rd.Org -and $rd.Org.ScheduledTasks -and $rd.Org.ScheduledTasks.Count -gt 0) {
                $null = $parts.Add((New-WdHeading (L '7.1 Geplante Tasks' '7.1 Scheduled Tasks') 2))
                $null = $parts.Add((New-WdParagraph (L 'EXpress registriert zwei operative geplante Aufgaben für den Exchange-Betrieb: MEAC (MonitorExchangeAuthCertificate.ps1) überwacht täglich das Exchange Auth-Zertifikat und erneuert es automatisch 60 Tage vor Ablauf — damit werden OAuth-/Hybrid-Ausfälle zuverlässig verhindert. EXpress Log-Cleanup bereinigt Exchange-Log-Verzeichnisse (Transport-Logs, IIS-Logs, HttpProxy-Logs, ETL/ETW) entsprechend der konfigurierten Aufbewahrungsfrist und verhindert ein Volllaufen des Log-Volumes.' 'EXpress registers two operational scheduled tasks for Exchange operations: MEAC (MonitorExchangeAuthCertificate.ps1) monitors the Exchange Auth certificate daily and automatically renews it 60 days before expiry — reliably preventing OAuth/Hybrid outages. EXpress Log-Cleanup purges Exchange log directories (transport logs, IIS logs, HttpProxy logs, ETL/ETW) according to the configured retention period, preventing the log volume from filling up.')))
                $stRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($t in $rd.Org.ScheduledTasks) {
                    $last = if ($t.LastRun)  { $t.LastRun.ToString('yyyy-MM-dd HH:mm')  } else { '—' }
                    $next = if ($t.NextRun)  { $t.NextRun.ToString('yyyy-MM-dd HH:mm')  } else { '—' }
                    $res  = if ($null -ne $t.LastResult) { ('0x{0:X}' -f $t.LastResult) } else { '—' }
                    $purpose =
                        if     ($t.Name -match 'Daily Auth Certificate|MonitorExchangeAuthCertificate|Monitor Exchange Auth') { (L 'Auto-Erneuerung Exchange Auth-Zertifikat (OAuth/Hybrid) — MEAC/CSS-Exchange' 'Auto-renewal of Exchange Auth certificate (OAuth/Hybrid) — MEAC/CSS-Exchange') }
                        elseif ($t.Name -match 'Log.?Cleanup|EXpressLogCleanup')                                            { (L 'Bereinigung Exchange-Log-Verzeichnisse' 'Cleanup of Exchange log directories') }
                        else                                                                           { '' }
                    $stRows.Add(@($t.Name, (SafeVal $t.Path), (SafeVal $t.State), $last, $next, $res, $purpose))
                }
                $null = $parts.Add((New-WdTable -Headers @((L 'Aufgabe' 'Task'), (L 'Pfad' 'Path'), (L 'Status' 'State'), (L 'Letzter Lauf' 'Last run'), (L 'Nächster Lauf' 'Next run'), (L 'Ergebnis' 'Result'), (L 'Zweck' 'Purpose')) -Rows $stRows.ToArray()))
            }

            # 7.2 Sicherheitsupdate-Stand
            $null = $parts.Add((New-WdHeading (L '7.2 Sicherheitsupdate-Stand' '7.2 Security Update Status') 2))
            $null = $parts.Add((New-WdParagraph (L 'Für Auditierbarkeit und Compliance ist der Patch-Stand des Exchange-Servers zu dokumentieren. Exchange Security Updates (SU) beheben kritische Sicherheitslücken (CVE) und müssen innerhalb der internen Patch-Window-Frist eingespielt werden. Neue SUs erscheinen monatlich (Patch Tuesday) oder außerplanmäßig bei kritischen Lücken. Der aktuelle Patch-Stand lässt sich über HealthChecker und Get-ExchangeDiagnosticInfo überprüfen.' 'For auditability and compliance, the patch status of the Exchange server must be documented. Exchange Security Updates (SU) fix critical vulnerabilities (CVE) and must be applied within the internal patch window. New SUs are released monthly (Patch Tuesday) or out-of-band for critical issues. The current patch status can be verified via HealthChecker and Get-ExchangeDiagnosticInfo.')))
            $suRows = [System.Collections.Generic.List[object[]]]::new()
            try {
                $exSrvSU = Get-ExchangeServer $env:COMPUTERNAME -ErrorAction Stop
                $suRows.Add(@((L 'Exchange-Version (Build)' 'Exchange version (build)'), $exSrvSU.AdminDisplayVersion.ToString()))
            } catch { Write-MyVerbose ('Get-ExchangeServer (SU version) failed: {0}' -f $_) }
            try {
                $osVer = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue)
                if ($osVer) {
                    $suRows.Add(@((L 'Windows-Version' 'Windows version'), ('{0} (Build {1})' -f $osVer.Caption, $osVer.BuildNumber)))
                    $suRows.Add(@((L 'Letzter Systemstart' 'Last system boot'), $osVer.LastBootUpTime.ToString('yyyy-MM-dd HH:mm:ss')))
                }
            } catch { Write-MyVerbose ('Get-CimInstance Win32_OperatingSystem (SU section) failed: {0}' -f $_) }
            if ($State['ExchangeSUVersion']) { $suRows.Add(@((L 'Exchange SU (dieser Lauf)' 'Exchange SU (this run)'), (SafeVal $State['ExchangeSUVersion']))) }
            $suRows.Add(@((L 'Empfehlung' 'Recommendation'), (L 'HealthChecker nach jedem SU ausführen — https://aka.ms/ExchangeHealthChecker' 'Run HealthChecker after every SU — https://aka.ms/ExchangeHealthChecker')))
            $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $suRows.ToArray()))
        }

        # ── 8. Optimierungen und Härtungen (lokal) ────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '8. Optimierungen und Härtungen (lokaler Server)' '8. Optimisations and Hardening (local server)') 1))
        $null = $parts.Add((New-WdParagraph (L 'Im Rahmen der Installation wurden auf diesem Server gezielte Sicherheitshärtungen und Leistungsoptimierungen angewendet. Die Maßnahmen orientieren sich an den Empfehlungen des Microsoft Exchange-Teams, dem CIS Benchmark sowie Best Practices für Exchange Server in Unternehmensumgebungen. Die folgende Tabelle dokumentiert den aktuellen Konfigurationsstatus der wichtigsten Härtungsmaßnahmen.' 'As part of the installation, targeted security hardening measures and performance optimisations were applied to this server. The measures are based on the recommendations of the Microsoft Exchange team, the CIS Benchmark, and best practices for Exchange Server in enterprise environments. The following table documents the current configuration status of the most important hardening measures.')))

        $null = $parts.Add((New-WdHeading (L '8.1 TLS und Kryptografie' '8.1 TLS and Cryptography') 2))
        $null = $parts.Add((New-WdParagraph (L 'Exchange Server kommuniziert intern (MAPI, EWS, Autodiscover) und extern (SMTP, OWA, ActiveSync) ausschließlich über TLS-verschlüsselte Verbindungen. TLS 1.0 und 1.1 gelten als unsicher (POODLE, BEAST) und wurden deaktiviert. TLS 1.2 ist das Mindestprotokoll; TLS 1.3 wird auf Windows Server 2022/2025 zusätzlich aktiviert. Die .NET Strong Crypto-Einstellung stellt sicher, dass auch alle .NET-Anwendungen auf diesem Server ausschließlich sichere Cipher Suites verwenden.' 'Exchange Server communicates internally (MAPI, EWS, Autodiscover) and externally (SMTP, OWA, ActiveSync) exclusively over TLS-encrypted connections. TLS 1.0 and 1.1 are considered insecure (POODLE, BEAST) and have been disabled. TLS 1.2 is the minimum protocol; TLS 1.3 is additionally enabled on Windows Server 2022/2025. The .NET Strong Crypto setting ensures that all .NET applications on this server also use only secure cipher suites.')))
        $tlsRows = [System.Collections.Generic.List[object[]]]::new()
        # Helper: derive a semantic protocol state from Enabled + DisabledByDefault registry values.
        # Raw "Enabled=0" / "Disabled=1" values are ambiguous at a glance ("Disabled=0 means active?"),
        # so translate into plain text: Enabled / Disabled / OS-Default.
        function Get-TlsProtocolState([string]$proto, [bool]$shouldBeEnabled) {
            $base = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\{0}\Server' -f $proto
            $en  = Get-SecReg $base 'Enabled'
            $dbd = Get-SecReg $base 'DisabledByDefault'
            $effEnabled = $null
            if ($null -ne $en)        { $effEnabled = ([int]$en -ne 0) }
            elseif ($null -ne $dbd)   { $effEnabled = ([int]$dbd -eq 0) }
            if ($null -eq $effEnabled) {
                # Not present in registry → OS default. WS2022+/SE: TLS 1.0/1.1 disabled by default.
                $osDefEnabled = ($proto -in 'TLS 1.2','TLS 1.3')
                return if ($osDefEnabled) { (L 'aktiviert (OS-Standard)' 'enabled (OS default)') } else { (L 'deaktiviert (OS-Standard)' 'disabled (OS default)') }
            }
            $stateText = if ($effEnabled) { (L 'aktiviert' 'enabled') } else { (L 'deaktiviert' 'disabled') }
            $warn      = ''
            if ($shouldBeEnabled -and -not $effEnabled)    { $warn = (L ' — ACHTUNG: sollte aktiviert sein'   ' — WARNING: should be enabled') }
            if (-not $shouldBeEnabled -and $effEnabled)    { $warn = (L ' — ACHTUNG: sollte deaktiviert sein' ' — WARNING: should be disabled') }
            if ($warn) { '{0}{1}' -f $stateText, $warn } else { $stateText }
        }
        $tlsRows.Add(@('TLS 1.0 Server', (Get-TlsProtocolState 'TLS 1.0' $false)))
        $tlsRows.Add(@('TLS 1.1 Server', (Get-TlsProtocolState 'TLS 1.1' $false)))
        $tlsRows.Add(@('TLS 1.2 Server', (Get-TlsProtocolState 'TLS 1.2' $true)))
        $tlsRows.Add(@('TLS 1.3 Server', (Get-TlsProtocolState 'TLS 1.3' $true)))
        $tlsRows.Add(@('.NET Strong Crypto (v4)', (Format-RegBool (Get-SecReg 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319' 'SchUseStrongCrypto'))))
        $tlsRows.Add(@('.NET Strong Crypto (v2)', (Format-RegBool (Get-SecReg 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727' 'SchUseStrongCrypto'))))
        $null = $parts.Add((New-WdTable -Headers @((L 'Maßnahme' 'Measure'), (L 'Registrierungswert / Status' 'Registry value / status')) -Rows $tlsRows.ToArray()))

        # TLS Cipher Suite inventory
        $null = $parts.Add((New-WdParagraph (L 'TLS Cipher Suites (aktiv auf diesem Server):' 'TLS Cipher Suites (active on this server):')))
        $cipherRows = [System.Collections.Generic.List[object[]]]::new()
        # Where-Object { $_ } filters out $null pipeline values that Select-Object would otherwise
        # silently promote to a single all-null PSCustomObject, bypassing the Count -eq 0 fallback.
        $cipherSuites = @(Get-TlsCipherSuite -ErrorAction SilentlyContinue | Where-Object { $_ } | Select-Object Name, Exchange, Hash, KeyExchange)
        foreach ($cs2 in $cipherSuites) {
            $cipherRows.Add(@($cs2.Name, (SafeVal $cs2.Exchange '—'), (SafeVal $cs2.Hash '—'), (SafeVal $cs2.KeyExchange '—')))
        }
        if ($cipherRows.Count -eq 0) { $cipherRows.Add(@((L '(nicht abrufbar)' '(not available)'), '', '', '')) }
        $null = $parts.Add((New-WdTable -Compact -Headers @((L 'Suite' 'Suite'), (L 'Schlüsseltausch' 'Key exchange'), (L 'Hash' 'Hash'), (L 'Algorithmus' 'Algorithm')) -Rows $cipherRows.ToArray()))

        $null = $parts.Add((New-WdHeading (L '8.2 Authentifizierung und Credential-Schutz' '8.2 Authentication and Credential Protection') 2))
        $null = $parts.Add((New-WdParagraph (L 'WDigest-Authentifizierung speichert Anmeldeinformationen im Klartextformat im LSASS-Speicher und ist für Pass-the-Hash- und Credential-Dumping-Angriffe (Mimikatz) anfällig. Sie wurde deaktiviert. LSA-Schutz (RunAsPPL) verhindert das Injizieren von unsigniertem Code in den LSASS-Prozess — ein zentraler Schutz gegen moderne Angriffswerkzeuge. Der LM-Kompatibilitätslevel bestimmt, welche Authentifizierungsprotokolle zugelassen werden; Level 5 (nur NTLMv2/Kerberos) entspricht dem aktuellen Sicherheitsstandard. Credential Guard (VBS) isoliert Credential-Hashes in einer virtualisierten Umgebung und ist auf Exchange-Servern zu deaktivieren, da Exchange interne Dienst-Konten mit NTLM-Authentifizierung nutzt.' 'WDigest authentication stores credentials in cleartext in LSASS memory and is vulnerable to pass-the-hash and credential dumping attacks (Mimikatz). It has been disabled. LSA protection (RunAsPPL) prevents injection of unsigned code into the LSASS process — a central protection against modern attack tools. The LM compatibility level determines which authentication protocols are permitted; level 5 (NTLMv2/Kerberos only) meets the current security standard. Credential Guard (VBS) isolates credential hashes in a virtualised environment and must be disabled on Exchange servers, as Exchange uses internal service accounts with NTLM authentication.')))
        $authRows = [System.Collections.Generic.List[object[]]]::new()
        $authRows.Add(@('WDigest UseLogonCredential', (Format-RegBool (Get-SecReg 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest' 'UseLogonCredential'))))
        $authRows.Add(@('LSA RunAsPPL',               (Format-RegBool (Get-SecReg 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' 'RunAsPPL'))))
        $lmLevel = Get-SecReg 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' 'LmCompatibilityLevel'
        $lmText  = if ($null -eq $lmLevel) { (L 'nicht gesetzt (Standard: 3)' 'not set (default: 3)') } else { 'Level {0}' -f $lmLevel }
        $authRows.Add(@('LM Compatibility Level', $lmText))
        $authRows.Add(@('Credential Guard (VBS)',  (Format-RegBool (Get-SecReg 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard' 'EnableVirtualizationBasedSecurity'))))
        $null = $parts.Add((New-WdTable -Headers @((L 'Maßnahme' 'Measure'), (L 'Registrierungswert / Status' 'Registry value / status')) -Rows $authRows.ToArray()))

        $null = $parts.Add((New-WdHeading (L '8.3 Netzwerkprotokolle' '8.3 Network Protocols') 2))
        $null = $parts.Add((New-WdParagraph (L 'SMBv1 ist ein veraltetes Dateifreigabeprotokoll ohne Verschlüsselung, das für WannaCry, NotPetya und ähnliche Ransomware-Angriffe genutzt wurde. Es wurde vollständig deaktiviert. HTTP/2 für Exchange-Webdienste wird deaktiviert, da es mit bestimmten Load-Balancer-Konfigurationen und dem Exchange Extended Protection-Mechanismus interferiert. SSL-Offloading (Beendigung der TLS-Verbindung am Load Balancer) ist deaktiviert, da Extended Protection eine End-to-End-TLS-Bindung erfordert.' 'SMBv1 is an outdated file-sharing protocol without encryption that was exploited by WannaCry, NotPetya and similar ransomware attacks. It has been completely disabled. HTTP/2 for Exchange web services is disabled as it interferes with certain load balancer configurations and the Exchange Extended Protection mechanism. SSL offloading (terminating the TLS connection at the load balancer) is disabled because Extended Protection requires end-to-end TLS binding.')))
        $protoRows = [System.Collections.Generic.List[object[]]]::new()
        try { $smb1 = (Get-SmbServerConfiguration -ErrorAction Stop).EnableSMB1Protocol; $protoRows.Add(@('SMBv1', (Format-RegBool $smb1))) } catch { Write-MyVerbose ('Get-SmbServerConfiguration failed: {0}' -f $_) }
        $protoRows.Add(@('HTTP/2 Cleartext (Exchange FE)', (Format-RegBool (Get-SecReg 'HKLM:\SYSTEM\CurrentControlSet\Services\HTTP\Parameters' 'EnableHttp2Cleartext'))))
        $null = $parts.Add((New-WdTable -Headers @((L 'Maßnahme' 'Measure'), (L 'Registrierungswert / Status' 'Registry value / status')) -Rows $protoRows.ToArray()))

        $null = $parts.Add((New-WdHeading (L '8.4 Exchange-spezifische Härtung' '8.4 Exchange-specific Hardening') 2))
        $null = $parts.Add((New-WdParagraph (L 'Extended Protection (EPA) verhindert Man-in-the-Middle-Angriffe auf HTTP-Verbindungen, indem die TLS-Channel-Binding-Information in die Authentifizierung einbezogen wird. Serialized Data Signing (SDS) schützt vor Deserialisierungsangriffen auf Exchange-interne Kommunikation. AMSI-Body-Scanning prüft HTTP-Anfragen (OWA, ECP, EWS, PowerShell) auf bekannte Angriffsmuster durch die Windows Defender AMSI-Schnittstelle. Die MAPI-Verschlüsselung stellt sicher, dass Outlook-MAPI-Verbindungen ausschließlich verschlüsselt erfolgen. Strict Mode für Powershell-Remoting und die Deaktivierung der PowerShell Autodiscover-App-Pools senken die Angriffsfläche der Exchange-Management-Schnittstellen weiter.' 'Extended Protection (EPA) prevents man-in-the-middle attacks on HTTP connections by incorporating TLS channel binding information into authentication. Serialized Data Signing (SDS) protects against deserialization attacks on Exchange internal communication. AMSI body scanning checks HTTP requests (OWA, ECP, EWS, PowerShell) for known attack patterns via the Windows Defender AMSI interface. MAPI encryption ensures that Outlook MAPI connections are exclusively encrypted. Strict mode for PowerShell remoting and disabling the PowerShell Autodiscover app pools further reduce the attack surface of the Exchange management interfaces.')))
        $exHardRows = [System.Collections.Generic.List[object[]]]::new()
        # Pull authoritative values from Exchange where available; fall back to registry-only hints otherwise.
        $epaState = '(unknown)'
        try {
            $epAuthDirs = @(Get-ExchangeServer $env:COMPUTERNAME -ErrorAction Stop | Out-Null)  # ensure EMS available
            $vdAuth = @()
            try { Get-OwaVirtualDirectory -Server $env:COMPUTERNAME -ErrorAction Stop       | ForEach-Object { $vdAuth += ('OWA={0}'        -f $_.ExtendedProtectionTokenChecking) } } catch { Write-MyVerbose ('Get-OwaVirtualDirectory (EPA) failed: {0}' -f $_) }
            try { Get-EcpVirtualDirectory -Server $env:COMPUTERNAME -ErrorAction Stop       | ForEach-Object { $vdAuth += ('ECP={0}'        -f $_.ExtendedProtectionTokenChecking) } } catch { Write-MyVerbose ('Get-EcpVirtualDirectory (EPA) failed: {0}' -f $_) }
            try { Get-WebServicesVirtualDirectory -Server $env:COMPUTERNAME -ErrorAction Stop | ForEach-Object { $vdAuth += ('EWS={0}'       -f $_.ExtendedProtectionTokenChecking) } } catch { Write-MyVerbose ('Get-WebServicesVirtualDirectory (EPA) failed: {0}' -f $_) }
            try { Get-OabVirtualDirectory -Server $env:COMPUTERNAME -ErrorAction Stop       | ForEach-Object { $vdAuth += ('OAB={0}'        -f $_.ExtendedProtectionTokenChecking) } } catch { Write-MyVerbose ('Get-OabVirtualDirectory (EPA) failed: {0}' -f $_) }
            try { Get-ActiveSyncVirtualDirectory -Server $env:COMPUTERNAME -ErrorAction Stop | ForEach-Object { $vdAuth += ('EAS={0}'       -f $_.ExtendedProtectionTokenChecking) } } catch { Write-MyVerbose ('Get-ActiveSyncVirtualDirectory (EPA) failed: {0}' -f $_) }
            try { Get-MapiVirtualDirectory -Server $env:COMPUTERNAME -ErrorAction Stop      | ForEach-Object { $vdAuth += ('MAPI={0}'       -f $_.ExtendedProtectionTokenChecking) } } catch { Write-MyVerbose ('Get-MapiVirtualDirectory (EPA) failed: {0}' -f $_) }
            try { Get-AutodiscoverVirtualDirectory -Server $env:COMPUTERNAME -ErrorAction Stop | ForEach-Object { $vdAuth += ('Autodiscover={0}' -f $_.ExtendedProtectionTokenChecking) } } catch { Write-MyVerbose ('Get-AutodiscoverVirtualDirectory (EPA) failed: {0}' -f $_) }
            if ($vdAuth.Count -gt 0) { $epaState = ($vdAuth -join ', ') }
        } catch { Write-MyVerbose ('EPA vdir query — EMS not available: {0}' -f $_) }
        $exHardRows.Add(@('Extended Protection (EPA)', $epaState, (L 'Channel-Binding-Schutz gegen MITM auf IIS-VDirs' 'Channel-binding protection against MITM on IIS VDirs')))
        # Registry value name is EnableSerializationDataSigning (Microsoft's actual spelling), not EnableSerializedDataSigning.
        $exHardRows.Add(@('Serialized Data Signing', (Format-RegBool (Get-SecReg 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Diagnostics' 'EnableSerializationDataSigning')), (L 'Schutz gegen Deserialisierungs-Angriffe (ab März 2024 verpflichtend)' 'Protection against deserialization attacks (required since March 2024)')))
        $amsiVal  = Get-SecReg 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Diagnostics' 'DisableAMSIScanning'
        $amsiText = if ($null -eq $amsiVal) { (L 'aktiviert (Standard)' 'enabled (default)') } elseif (([int]"$amsiVal") -eq 0) { (L 'aktiviert' 'enabled') } else { (L 'deaktiviert' 'disabled') }
        $exHardRows.Add(@('AMSI Body Scanning', $amsiText, (L 'HTTP-Request-Scan über Windows Defender AMSI' 'HTTP request scan via Windows Defender AMSI')))
        $mapiEnc = try { (Get-RpcClientAccess -Server $env:COMPUTERNAME -ErrorAction Stop | Select-Object -First 1).EncryptionRequired.ToString() } catch { '(unknown)' }
        $exHardRows.Add(@('MAPI Encryption Required', (SafeVal $mapiEnc), (L 'Outlook-/MAPI-Verbindungen nur verschlüsselt' 'Outlook/MAPI connections encrypted only')))
        # Throttling / rate limiting for Exchange Web Services (mitigates abuse / DoS on EWS endpoint)
        $throt = try {
            $tp = Get-ThrottlingPolicy -ErrorAction Stop | Where-Object { $_.IsDefault } | Select-Object -First 1
            if ($tp -and $null -ne $tp.EwsMaxConcurrency) { $tp.EwsMaxConcurrency.ToString() }
            else { (L '(nicht gesetzt — Standard: 27)' '(not set — default: 27)') }
        } catch { (L '(nicht abrufbar)' '(not available)') }
        $exHardRows.Add(@('EWS Max Concurrency (default policy)', $throt, (L 'Throttling-Policy gegen EWS-Überlastung' 'Throttling policy against EWS overload')))
        # Authentication flags on OWA/ECP
        $owaBasic = try { (Get-OwaVirtualDirectory -Server $env:COMPUTERNAME -ErrorAction Stop | Select-Object -First 1).BasicAuthentication.ToString() } catch { '(unknown)' }
        $exHardRows.Add(@('OWA Basic Authentication', (SafeVal $owaBasic), (L 'Basic-Auth auf OWA ist gegen Credential-Stuffing anfällig' 'Basic auth on OWA is vulnerable to credential stuffing')))
        # ECC certificate support (cipher modernization)
        $exHardRows.Add(@('ECC Certificate Support', (Format-RegBool (Get-SecReg 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Diagnostics' 'EnableEccCertificateSupport')), (L 'Moderne ECC-Zertifikate in Exchange zugelassen' 'Modern ECC certificates permitted in Exchange')))
        # Setup-Override files (SettingOverride framework) — CVE-bezogene Kill-Switches
        if (Get-Command Get-ExchangeSettingOverride -ErrorAction SilentlyContinue) {
            try {
                $overrides = @(Get-ExchangeSettingOverride -ErrorAction Stop)
                if ($overrides) {
                    $ovList = ($overrides | ForEach-Object { '{0}:{1}={2}' -f $_.ComponentName, $_.SectionName, ($_.Parameters -join ',') }) -join '; '
                    $exHardRows.Add(@('Exchange SettingOverrides', (SafeVal $ovList), (L 'Aktive Konfigurations-Overrides (CVE-Mitigationen, Features)' 'Active configuration overrides (CVE mitigations, features)')))
                }
            } catch { Write-MyDebug ('Get-ExchangeSettingOverride failed: {0}' -f $_) }
        }
        # HSTS — HTTP Strict Transport Security
        $hstsVal = try {
            Import-Module WebAdministration -ErrorAction SilentlyContinue
            $hh = Get-WebConfigurationProperty -Filter "system.webServer/httpProtocol/customHeaders/add[@name='Strict-Transport-Security']" -PSPath 'IIS:\Sites\Default Web Site' -Name 'value' -ErrorAction Stop
            if ($hh) { 'max-age={0}' -f $hh } else { (L 'nicht gesetzt' 'not set') }
        } catch { (L 'nicht gesetzt / nicht abrufbar' 'not set / not available') }
        $exHardRows.Add(@('HSTS (Strict-Transport-Security)', $hstsVal, (L 'Verhindert HTTP-Downgrade-Angriffe (RFC 6797)' 'Prevents HTTP downgrade attacks (RFC 6797)')))
        # CVE-2021-1730 Download Domains
        $dlEnabled = try {
            $dlCfg = Get-OrganizationConfig -ErrorAction Stop
            if ($null -ne $dlCfg.EnableDownloadDomains) { $dlCfg.EnableDownloadDomains.ToString() } else { (L 'nicht gesetzt' 'not set') }
        } catch { '(unknown)' }
        $dlDomain = if ($State['DownloadDomain']) { $State['DownloadDomain'] } else { '—' }
        $exHardRows.Add(@('EnableDownloadDomains (CVE-2021-1730)', ('{0} — Domain: {1}' -f $dlEnabled, $dlDomain), (L 'Isoliert OWA-Anhänge auf Subdomain (verhindert CSRF/Cookie-Hijacking)' 'Isolates OWA attachments on subdomain (prevents CSRF/cookie hijacking)')))
        # SMTP Protocol Logging
        $smtpLogVal = try {
            $logConns = @(Get-ReceiveConnector -Server $env:COMPUTERNAME -ErrorAction Stop |
                Where-Object { $_.Name -match '^Default Frontend |^Anonymous (Internal|External) Relay' })
            if ($logConns.Count -gt 0) {
                $verboseCount = ($logConns | Where-Object { $_.ProtocolLoggingLevel -eq 'Verbose' }).Count
                '{0}/{1} Verbose' -f $verboseCount, $logConns.Count
            } else { (L '(kein passender Connector)' '(no matching connector)') }
        } catch { (L '(nicht abrufbar)' '(not available)') }
        $exHardRows.Add(@('SMTP Protocol Logging', $smtpLogVal, (L 'Verbose-Logging auf Default Frontend + Relay-Connectors (BSI APP.5.3)' 'Verbose logging on Default Frontend + relay connectors (BSI APP.5.3)')))
        $null = $parts.Add((New-WdTable -Headers @((L 'Härtungsmaßnahme' 'Hardening measure'), (L 'Status / Wert' 'Status / value'), (L 'Zweck' 'Purpose')) -Rows $exHardRows.ToArray()))

        # 8.5 Windows Defender Exclusions
        $localSrvData = @($rd.Servers | Where-Object { $_.IsLocalServer }) | Select-Object -First 1
        if ($localSrvData -and $localSrvData.DefenderExclusions) {
            $null = $parts.Add((New-WdHeading (L '8.5 Windows Defender — Ausnahmen' '8.5 Windows Defender — Exclusions') 2))
            $null = $parts.Add((New-WdParagraph (L 'Microsoft dokumentiert umfangreiche Pfad-, Prozess- und Dateityp-Ausnahmen für Exchange Server, ohne die Antivirus-Software Datenbank-Dateien, Transport-Warteschlangen oder Logs blockiert und Leistung wie Stabilität schwer beeinträchtigt. EXpress trägt diese Ausnahmen automatisch in Windows Defender ein. Bei Drittanbieter-Antivirus müssen dieselben Pfade manuell in das entsprechende Produkt übernommen werden. Weitere Informationen: Microsoft Docs "Exchange antivirus software".' 'Microsoft documents extensive path, process and filetype exclusions for Exchange Server without which antivirus software would block database files, transport queues or logs and severely impact performance and stability. EXpress automatically registers these exclusions with Windows Defender. For third-party antivirus, the same paths must be manually configured in the corresponding product. Further information: Microsoft Docs "Exchange antivirus software".')))
            $exr = $localSrvData.DefenderExclusions
            $defRows = [System.Collections.Generic.List[object[]]]::new()
            $dvModeStr = if ($exr.AMRunningMode) { ' ({0})' -f $exr.AMRunningMode } else { '' }
            $defRows.Add(@((L 'Echtzeit-Überwachung' 'Real-time monitoring'), ((Lc $exr.RealTimeEnabled (L 'aktiv' 'enabled') (L 'inaktiv' 'disabled')) + $dvModeStr)))
            $defRows.Add(@((L 'Pfad-Ausnahmen' 'Path exclusions'), (SafeVal (($exr.ExclusionPath | Sort-Object) -join "`n") (L '(keine)' '(none)'))))
            $defRows.Add(@((L 'Prozess-Ausnahmen' 'Process exclusions'), (SafeVal (($exr.ExclusionProcess | Sort-Object) -join "`n") (L '(keine)' '(none)'))))
            $defRows.Add(@((L 'Dateityp-Ausnahmen' 'Extension exclusions'), (SafeVal (($exr.ExclusionExtension | Sort-Object) -join "`n") (L '(keine)' '(none)'))))
            # ColWidths: narrow label 2500, wide value (paths up to 121 chars) 6760 — total 9260 twips
            $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $defRows.ToArray() -ColWidths @(2500, 6760)))
        }

        # 8.6 IIS- und Exchange-Logs
        $null = $parts.Add((New-WdHeading (L '8.6 Protokollierung — IIS und Exchange' '8.6 Logging — IIS and Exchange') 2))
        $null = $parts.Add((New-WdParagraph (L 'Exchange Server schreibt umfangreiche Betriebsprotokolle in den Logging-Pfad unter dem Exchange-Installationsverzeichnis (Transport, Managed Availability, HttpProxy, CAS). IIS protokolliert Zugriffe auf OWA, ECP, EWS, ActiveSync, MAPI, OAB. Ohne automatische Bereinigung füllen diese Logs innerhalb weniger Wochen das Log-Volume vollständig auf. EXpress registriert hierfür einen geplanten Task (siehe 7.1), der Logs älter als der konfigurierten Aufbewahrungsfrist (Standard: 30 Tage) automatisch entfernt. Die tatsächlichen IIS-Log-Pfade zeigt die folgende Tabelle.' 'Exchange Server writes extensive operational logs to the logging path below the Exchange installation directory (Transport, Managed Availability, HttpProxy, CAS). IIS logs access to OWA, ECP, EWS, ActiveSync, MAPI, OAB. Without automatic cleanup these logs fill the log volume completely within a few weeks. EXpress registers a scheduled task for this purpose (see 7.1) which automatically removes logs older than the configured retention (default: 30 days). Actual IIS log paths are shown in the table below.')))
        $null = $parts.Add((New-WdParagraph (L 'Hinweis zu Forensik und Compliance: Die regelmäßige lokale Bereinigung dient ausschließlich dazu, ein Vollaufen des Log-Volumes (und damit den Ausfall von Transport, IIS und Managed Availability) zu verhindern — sie ist kein Ersatz für eine revisionssichere Langzeit-Aufbewahrung. Für forensische Auswertung sicherheitsrelevanter Vorfälle (Authentifizierungs-Anomalien, EWS-/MAPI-Zugriffsmuster, Transport-Spuren bei Datenabfluss) und zur Erfüllung gesetzlicher Aufbewahrungspflichten (BSI APP.5.2, DSGVO Rechenschaftspflicht, GoBD) sind IIS-, HttpProxy-, MessageTracking-, Transport- und Windows-Security-Eventlogs idealerweise per Log-Forwarder (z. B. NXLog, WEF/WEC, Filebeat, Azure Monitor Agent) an ein zentrales SIEM (z. B. Splunk, Elastic Security, Microsoft Sentinel, Wazuh, IBM QRadar) auszuleiten. Die Aufbewahrungsdauer im SIEM sollte sich an der internen Sicherheitsleitlinie und branchenspezifischen Vorgaben orientieren (typisch 12 Monate Hot-Storage, 7 Jahre Archiv). Erst diese Kombination — kurze Aufbewahrung am Server, lange Aufbewahrung im SIEM — erfüllt sowohl operative Stabilitätsanforderungen als auch forensische und Compliance-Anforderungen.' 'Note on forensics and compliance: Periodic local cleanup is intended solely to prevent the log volume from filling up (which would take down Transport, IIS and Managed Availability) — it is not a substitute for tamper-evident long-term retention. For forensic investigation of security-relevant incidents (authentication anomalies, EWS/MAPI access patterns, transport traces during data exfiltration) and to meet legal retention obligations (BSI APP.5.2, GDPR accountability, GoBD), IIS, HttpProxy, MessageTracking, Transport and Windows Security event logs should ideally be forwarded via a log shipper (e.g. NXLog, WEF/WEC, Filebeat, Azure Monitor Agent) to a central SIEM (e.g. Splunk, Elastic Security, Microsoft Sentinel, Wazuh, IBM QRadar). Retention in the SIEM should follow the organisation''s security policy and industry-specific requirements (typically 12 months hot storage, 7 years archive). Only this combination — short retention on the server, long retention in the SIEM — satisfies both operational stability and forensic/compliance requirements.')))
        $logRows = [System.Collections.Generic.List[object[]]]::new()
        $logRows.Add(@((L 'Exchange Logging-Pfad' 'Exchange logging path'), (SafeVal (Join-Path (Split-Path $env:ExchangeInstallPath -Parent) 'Logging'))))
        $logRows.Add(@((L 'ETL/Diagnostic-Pfad' 'ETL/Diagnostic path'), (SafeVal (Join-Path (Split-Path $env:ExchangeInstallPath -Parent) 'Bin\Search\Ceres\HostController\Data\Events'))))
        $retDays = if ($State['LogRetentionDays']) { $State['LogRetentionDays'] } else { 30 }
        $logRows.Add(@((L 'Aufbewahrung (EXpress Log Cleanup)' 'Retention (EXpress log cleanup)'), ('{0} {1}' -f $retDays, (L 'Tage' 'days'))))
        if ($localSrvData -and $localSrvData.IISLogs) {
            foreach ($site in $localSrvData.IISLogs.Sites) {
                $logRows.Add(@(('IIS: {0}' -f $site.Name), ('{0} — Format: {1} — Period: {2}' -f (SafeVal $site.LogDir), (SafeVal $site.LogFormat), (SafeVal $site.Period))))
            }
        }
        $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $logRows.ToArray()))

        # 8.7 Kerberos Load Balancing
        $null = $parts.Add((New-WdHeading (L '8.7 Kerberos Load Balancing' '8.7 Kerberos Load Balancing') 2))
        $null = $parts.Add((New-WdParagraph (L 'In Umgebungen mit Hardware- oder Software-Load-Balancern (NLB, F5, Kemp, HAProxy u. a.) ist Kerberos-Authentifizierung für Exchange-Dienste ohne Session Affinity möglich, sofern ein dedizierter Kerberos-Service-Account (KSA) konfiguriert wird. Ohne KSA fällt Kerberos auf NTLM zurück, wenn der Client an einen anderen Server weitergeleitet wird als den, den er ursprünglich kontaktiert hat — NTLM erzeugt höhere Latenz und kann in großen Umgebungen zu NtLM-Stau führen. Der KSA erhält einen Service Principal Name (SPN) für jeden HTTPS-Dienst (OWA, EWS, Autodiscover, MAPI, ECP, ActiveSync, OAB) und wird in AD als Konto mit gesetztem "Kerberos-Einschränkungen zulassen" hinterlegt. Ab Exchange 2016 mit CAS-Array-Entfall ist Kerberos-LB eine optionale, aber empfehlenswerte Konfiguration für Umgebungen mit mehreren Exchange-Servern hinter einem LB.' 'In environments with hardware or software load balancers (NLB, F5, Kemp, HAProxy etc.), Kerberos authentication for Exchange services without session affinity is possible provided a dedicated Kerberos Service Account (KSA) is configured. Without a KSA, Kerberos falls back to NTLM when a client is redirected to a different server than the one it originally contacted — NTLM causes higher latency and can lead to NTLM saturation in large environments. The KSA is assigned a Service Principal Name (SPN) for each HTTPS service (OWA, EWS, Autodiscover, MAPI, ECP, ActiveSync, OAB) and registered in AD as an account with "Constrain Kerberos delegation" set. Since Exchange 2016 with the removal of the CAS array, Kerberos LB is an optional but recommended configuration for environments with multiple Exchange servers behind a load balancer.')))
        $krbRows = [System.Collections.Generic.List[object[]]]::new()
        try {
            $cas = @(Get-ClientAccessService -ErrorAction Stop)
            foreach ($c in $cas) {
                $ksa = try { $c.AlternateServiceAccountCredential | Select-Object -First 1 } catch { $null }
                $ksaName = if ($ksa -and $ksa.Credential) { $ksa.Credential.UserName } elseif ($c.AlternateServiceAccountCredential) { (SafeVal ($c.AlternateServiceAccountCredential -join ', ')) } else { (L '(kein KSA konfiguriert)' '(no KSA configured)') }
                $krbRows.Add(@($c.Name, $ksaName, (SafeVal $c.AutoDiscoverServiceInternalUri)))
            }
        } catch { Write-MyVerbose ('Get-ClientAccessService (Kerberos LB) failed: {0}' -f $_) }
        if ($krbRows.Count -eq 0) { $krbRows.Add(@((L '(Get-ClientAccessService nicht verfügbar)' '(Get-ClientAccessService not available)'), '', '')) }
        $null = $parts.Add((New-WdTable -Headers @((L 'CAS-Server' 'CAS server'), (L 'Kerberos-Service-Account' 'Kerberos service account'), 'Autodiscover URI') -Rows $krbRows.ToArray()))
        $null = $parts.Add((New-WdParagraph (L 'Konfigurationsreferenz: Set-ClientAccessService -AlternateServiceAccountCredential. Weitere Details und SPN-Registrierung: Microsoft Docs "Configure Kerberos authentication for load-balanced Exchange servers".' 'Configuration reference: Set-ClientAccessService -AlternateServiceAccountCredential. Further details and SPN registration: Microsoft Docs "Configure Kerberos authentication for load-balanced Exchange servers".')))

        # 8.8 Compliance-Mapping CIS / BSI IT-Grundschutz
        $null = $parts.Add((New-WdHeading (L '8.8 Compliance-Mapping (CIS / BSI IT-Grundschutz)' '8.8 Compliance Mapping (CIS / BSI)') 2))
        $null = $parts.Add((New-WdParagraph (L 'Die folgende Tabelle ordnet die von EXpress angewendeten Härtungsmaßnahmen den relevanten Kontrollen aus dem CIS Benchmark for Microsoft Windows Server und dem BSI IT-Grundschutz-Kompendium zu. Sie dient als Nachweis für Audits und interne Compliance-Prüfungen.' 'The table below maps the hardening measures applied by EXpress to the relevant controls from the CIS Benchmark for Microsoft Windows Server and the BSI IT-Grundschutz Compendium. It serves as evidence for audits and internal compliance reviews.')))
        $null = $parts.Add((New-WdParagraph (L 'Wichtiger Hinweis zur Protokoll-Auswertung: Mehrere der nachfolgenden Kontrollen — insbesondere Admin Audit Log, Mailbox Audit Log, Windows Security Eventlog und IIS-Zugriffsprotokolle — entfalten ihren vollen Compliance- und forensischen Nutzen erst, wenn die erzeugten Ereignisse zentral zusammengeführt, korreliert und revisionssicher aufbewahrt werden. EXpress aktiviert und konfiguriert die Protokollquellen auf dem Server, sieht jedoch ausdrücklich keine SIEM-Anbindung vor — diese ist organisationsweit zu planen und liegt außerhalb des Scopes einer Server-Installation. Für die Erfüllung von BSI APP.5.2 A13 (Protokollierung), BSI OPS.1.1.5 (Protokollierung), CIS Control 8 (Audit Log Management) sowie der DSGVO-Rechenschaftspflicht (Art. 5 Abs. 2) ist die Anbindung an ein SIEM (Security Information and Event Management) dringend empfohlen. Ein SIEM ermöglicht: (1) zentrale Korrelation über mehrere Exchange-Server, Domain Controller und Edge-Komponenten hinweg; (2) Alarmierung bei Anomalien (Brute-Force-Versuche, ungewöhnliche EWS-/PowerShell-Zugriffe, Mass-Mail-Abfluss); (3) revisionssichere Langzeit-Aufbewahrung über die lokale Bereinigungsfrist hinaus; (4) Nachweisführung gegenüber Auditoren ohne Eingriff am Produktivsystem. Empfohlene Quellen für die Auslieferung: Windows Security/System/Application-Eventlog, IIS-W3C-Logs, Exchange MessageTracking, HttpProxy, Managed Availability, sowie das Admin- und Mailbox-Audit-Log via Search-AdminAuditLog / Search-MailboxAuditLog oder New-MailboxAuditLogSearch.' 'Important note on log evaluation: Several of the controls below — in particular Admin Audit Log, Mailbox Audit Log, Windows Security event log and IIS access logs — only deliver their full compliance and forensic value when the generated events are centrally aggregated, correlated and retained tamper-evidently. EXpress enables and configures the log sources on the server, but explicitly does not provide SIEM integration — this must be planned organisation-wide and is out of scope for a server installation. To meet BSI APP.5.2 A13 (logging), BSI OPS.1.1.5 (logging), CIS Control 8 (Audit Log Management) and the GDPR accountability obligation (Art. 5(2)), integration with a SIEM (Security Information and Event Management) is strongly recommended. A SIEM enables: (1) central correlation across multiple Exchange servers, domain controllers and edge components; (2) alerting on anomalies (brute-force attempts, unusual EWS/PowerShell access, mass mail exfiltration); (3) tamper-evident long-term retention beyond the local cleanup period; (4) audit evidence without touching the production system. Recommended sources for forwarding: Windows Security/System/Application event log, IIS W3C logs, Exchange MessageTracking, HttpProxy, Managed Availability, plus the Admin and Mailbox Audit Log via Search-AdminAuditLog / Search-MailboxAuditLog or New-MailboxAuditLogSearch.')))
        $smtpLogStatus = if (Test-Feature 'SMTPConnectorLogging') { (L 'Umgesetzt' 'Implemented') } else { (L 'Deaktiviert' 'Disabled') }
        $null = $parts.Add((New-WdTable -Headers @((L 'Maßnahme' 'Measure'), (L 'CIS-Kontrolle' 'CIS Control'), (L 'BSI-Grundschutz' 'BSI Control'), (L 'Status' 'Status')) -Rows @(
            ,@((L 'TLS 1.0 / 1.1 deaktiviert' 'TLS 1.0 / 1.1 disabled'),                          'CIS WS2022 18.4.x',   'BSI SYS.1.2 A5',  (L 'Umgesetzt' 'Implemented'))
            ,@((L 'TLS 1.2 erzwungen + .NET Strong Crypto' 'TLS 1.2 enforced + .NET Strong Crypto'), 'CIS WS2022 18.4.x',   'BSI SYS.1.2 A5',  (L 'Umgesetzt' 'Implemented'))
            ,@('RC4 / 3DES / NULL Ciphers deaktiviert',                                              'CIS WS2022 2.3.11.x', 'BSI SYS.1.2 A6',  (L 'Umgesetzt' 'Implemented'))
            ,@((L 'SMBv1 deaktiviert' 'SMBv1 disabled'),                                             'CIS WS2022 18.3.4',   'BSI NET.3.4 A2',  (L 'Umgesetzt' 'Implemented'))
            ,@('NTLMv2 (LmCompatibilityLevel = 5)',                                                   'CIS WS2022 2.3.11.8', 'BSI SYS.1.2 A7',  (L 'Umgesetzt' 'Implemented'))
            ,@((L 'WDigest deaktiviert' 'WDigest disabled'),                                         'CIS WS2022 18.3.7',   'BSI SYS.1.6 A3',  (L 'Umgesetzt' 'Implemented'))
            ,@((L 'LSA-Schutz aktiviert' 'LSA Protection enabled'),                                  'CIS WS2022 18.4.5',   'BSI SYS.1.6 A5',  (L 'Umgesetzt' 'Implemented'))
            ,@('Extended Protection for Authentication (EPA)',                                         'CIS WS2022 18.4.x',   'BSI APP.5.2 A10', (L 'Umgesetzt' 'Implemented'))
            ,@('Serialized Data Signing',                                                              'MS Exchange SE Baseline', 'BSI APP.5.2 A10', (L 'Umgesetzt' 'Implemented'))
            ,@((L 'Defender Ausnahmen (Exchange-VSS, Transport, IIS)' 'Defender exclusions (Exchange VSS, Transport, IIS)'), 'MS Exchange Best Practice', 'BSI APP.5.2 A4', (L 'Umgesetzt' 'Implemented'))
            ,@('LLMNR / mDNS deaktiviert',                                                            'CIS WS2022 18.5.4.2', 'BSI NET.3.1 A10', (L 'Umgesetzt' 'Implemented'))
            ,@((L 'Dienste minimiert (Browser/Fax/Xcopy u. a.)' 'Services minimised (Browser/Fax/Xcopy etc.)'), 'CIS WS2022 5.x', 'BSI SYS.1.2 A3', (L 'Umgesetzt' 'Implemented'))
            ,@((L 'SMTP Protocol Logging (Default Frontend + Relay)' 'SMTP Protocol Logging (Default Frontend + relay)'), 'MS Exchange Best Practice', 'BSI APP.5.3', $smtpLogStatus)
            ,@((L 'Admin Audit Log aktiviert' 'Admin Audit Log enabled'),                             'CIS EX2019 1.1',      'BSI APP.5.2 A13', (L 'Umgesetzt' 'Implemented'))
            ,@((L 'SIEM-Anbindung / zentrale Log-Auswertung' 'SIEM integration / central log evaluation'), 'CIS Control 8',     'BSI OPS.1.1.5 / APP.5.2 A13', (L 'Out of Scope — organisationsweit zu planen' 'Out of scope — to be planned organisation-wide'))
            ,@((L 'Log-Bereinigung am Server (Volume-Schutz)' 'Local log cleanup (volume protection)'), 'MS Best Practice',     'BSI APP.5.2 A4',  (L 'Umgesetzt — geplante Aufgabe (siehe 7.1)' 'Implemented — scheduled task (see 7.1)'))
            ,@((L 'Defender Echtzeit deaktiviert (Exchange-Konflikt mit AWL)' 'Defender realtime disabled (Exchange AWL conflict)'), 'CIS WS2022 n/a', 'BSI SYS.1.2 A4', (L 'Ausnahme — Exchange-AWL-Konflikt; AV-Ausnahmen gesetzt' 'Exception — Exchange AWL conflict; AV exclusions applied'))
        ) -Compact))

        # 8.9 Datenschutz und DSGVO-Relevanz
        $null = $parts.Add((New-WdHeading (L '8.9 Datenschutz und DSGVO-Relevanz' '8.9 Data Protection and GDPR Relevance') 2))
        $null = $parts.Add((New-WdParagraph (L 'Exchange Server verarbeitet personenbezogene Daten (E-Mail-Inhalte, Adressdaten, Kalendereinträge, Postfachberechtigungen) und ist daher für Organisationen in der EU als Datenverarbeitungssystem im Sinne der DSGVO (Art. 4 Nr. 2) einzustufen. Die folgende Checkliste fasst die datenschutzrelevanten Aspekte zusammen.' 'Exchange Server processes personal data (email content, address data, calendar entries, mailbox permissions) and must therefore be classified as a data processing system under the GDPR (Art. 4 No. 2) for organisations in the EU. The checklist below summarises the data protection-relevant aspects.')))
        $null = $parts.Add((New-WdTable -Headers @((L 'Datenschutzaspekt' 'Data protection aspect'), (L 'Status / Hinweis' 'Status / Note')) -Rows @(
            ,@((L 'Transportverschlüsselung (TLS 1.2+)' 'Transport encryption (TLS 1.2+)'),                          (L 'Umgesetzt — TLS 1.2 auf allen Verbindungspunkten erzwungen' 'Implemented — TLS 1.2 enforced on all connection points'))
            ,@((L 'Ruheverschlüsselung (Encryption at rest)' 'Encryption at rest'),                                   (L 'BitLocker (OS-Ebene) empfohlen; Exchange-native DB-Verschlüsselung nicht verfügbar' 'BitLocker (OS level) recommended; Exchange-native DB encryption not available'))
            ,@((L 'Admin-Auditprotokoll' 'Admin Audit Log'),                                                           (L 'Umgesetzt — administrative Cmdlet-Ausführungen werden protokolliert' 'Implemented — administrative cmdlet executions are logged'))
            ,@((L 'Postfach-Zugriffsprotokoll (Mailbox Audit Logging)' 'Mailbox Audit Logging'),                      (L 'Ab Exchange 2019: standardmäßig aktiviert (Default Audit Logging)' 'From Exchange 2019: enabled by default (Default Audit Logging)'))
            ,@((L 'Aufbewahrungsrichtlinien / Löschfristen' 'Retention policies / deletion periods'),                  (L 'Über Compliance-Tags und Retention Policies im Compliance Center konfigurieren' 'Configure via Compliance Tags and Retention Policies in the Compliance Center'))
            ,@((L 'Verarbeitungsverzeichnis (Art. 30 DSGVO)' 'Records of processing activities (Art. 30 GDPR)'),      (L 'Exchange Server ist im Verarbeitungsverzeichnis zu führen' 'Exchange Server must be included in the records of processing activities'))
            ,@((L 'DSFA / DPIA (Art. 35 DSGVO)' 'DPIA (Art. 35 GDPR)'),                                              (L 'Bei umfangreicher Verarbeitung sensibler Daten ggf. erforderlich' 'May be required for extensive processing of sensitive data'))
            ,@((L 'Auftragsverarbeitung (AV-Vertrag / DPA)' 'Data Processing Agreement (DPA)'),                       (L 'Mit M365/EOP/AIP-Diensten ist ein AV-Vertrag (Microsoft DPA) abzuschließen' 'A DPA (Microsoft DPA) must be concluded for M365/EOP/AIP services'))
        )))

        # ── 9. Anti-Spam / Agents (lokal) ─────────────────────────────────────────
        Invoke-DocSection-AntiSpam -Parts $parts -OrgData $orgD -DE $DE -Cust $cust

        # ── 10. Backup- & DR-Readiness (lokal) ────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '10. Backup- und DR-Readiness' '10. Backup and DR Readiness') 1))
        $null = $parts.Add((New-WdParagraph (L 'Exchange Server unterstützt datenbankebene Sicherungen über die Volume Shadow Copy Service (VSS)-Schnittstelle. Eine ordnungsgemäß funktionierende VSS-Integration ist Voraussetzung für konsistente Exchange-Backups durch Backup-Software (Veeam, Windows Server Backup, Commvault u. a.). Nach einem Backup werden die Transaktionsprotokolle automatisch abgeschnitten (Log Truncation) — vorausgesetzt, Circular Logging ist deaktiviert. Für die Disaster-Recovery-Fähigkeit sind funktionierende VSS Writer, korrekte Exchange-Defender-Ausnahmen und ein regelmäßig getestetes Restore-Verfahren entscheidend.' 'Exchange Server supports database-level backups via the Volume Shadow Copy Service (VSS) interface. Correctly functioning VSS integration is a prerequisite for consistent Exchange backups by backup software (Veeam, Windows Server Backup, Commvault, etc.). After a backup, transaction logs are automatically truncated — provided Circular Logging is disabled. For disaster recovery capability, functioning VSS writers, correct Exchange Defender exclusions and a regularly tested restore procedure are essential.')))
        $null = $parts.Add((New-WdHeading (L '10.1 VSS Writer Status' '10.1 VSS Writer Status') 2))
        $null = $parts.Add((New-WdParagraph (L 'Alle Exchange-relevanten VSS Writer müssen im Zustand "Stabil" sein. Fehlerhafte Writer führen zu inkonsistenten oder fehlschlagenden Backups. Bei dauerhaft fehlerhaften Writern ist ein Neustart des betroffenen Dienstes (Microsoft Exchange Writer → MSExchangeIS) oder ein Server-Neustart erforderlich.' 'All Exchange-relevant VSS writers must be in a "Stable" state. Faulty writers lead to inconsistent or failed backups. For persistently faulty writers, a restart of the affected service (Microsoft Exchange Writer → MSExchangeIS) or a server restart is required.')))
        $vssRows = [System.Collections.Generic.List[object[]]]::new()
        try {
            $vssOut = (vssadmin list writers 2>&1) -join "`n"
            $curWriter = ''
            foreach ($line in ($vssOut -split "`n")) {
                if ($line -match "Writer name:\s+'(.+)'") { $curWriter = $Matches[1] }
                elseif ($line -match 'State:\s*\[\d+\]\s+(.+)') { $vssRows.Add(@($curWriter, $Matches[1].Trim())) }
            }
        } catch { $vssRows.Add(@((L 'VSS-Abfrage fehlgeschlagen' 'VSS query failed'), '')) }
        if ($vssRows.Count -eq 0) { $vssRows.Add(@((L '(keine VSS Writer gefunden)' '(no VSS writers found)'), '')) }
        $null = $parts.Add((New-WdTable -Headers @((L 'VSS Writer' 'VSS Writer'), (L 'Zustand' 'State')) -Rows $vssRows.ToArray()))
        $null = $parts.Add((New-WdHeading (L '10.2 Empfehlungen Backup-Strategie' '10.2 Backup Strategy Recommendations') 2))
        $null = $parts.Add((New-WdParagraph (L 'Für Exchange Server werden folgende Backup-Praktiken empfohlen:' 'The following backup practices are recommended for Exchange Server:')))
        $null = $parts.Add((New-WdBullet (L 'Tägliche VSS-Vollsicherung der Exchange-Datenbanken über eine Exchange-aware Backup-Lösung (kein File-Level-Backup laufender EDB-Dateien)' 'Daily VSS full backup of Exchange databases via an Exchange-aware backup solution (no file-level backup of running EDB files)')))
        $null = $parts.Add((New-WdBullet (L 'Transaktionsprotokolle werden nach erfolgreichem Backup automatisch abgeschnitten — Circular Logging sollte deaktiviert bleiben' 'Transaction logs are automatically truncated after a successful backup — Circular Logging should remain disabled')))
        $null = $parts.Add((New-WdBullet (L 'Restore-Test mindestens einmal jährlich in einer Testumgebung (Recovery Database, RDB) durchführen' 'Perform restore test at least once annually in a test environment (Recovery Database, RDB)')))
        $null = $parts.Add((New-WdBullet (L 'Backup der Active-Directory-Domänencontroller separat sicherstellen (Exchange ist AD-abhängig)' 'Ensure separate backup of Active Directory domain controllers (Exchange is AD-dependent)')))
        $null = $parts.Add((New-WdHeading (L '10.3 Disaster-Recovery-Szenarien' '10.3 Disaster Recovery Scenarios') 2))
        $null = $parts.Add((New-WdParagraph (L 'Die folgende Tabelle gibt einen Überblick über typische DR-Szenarien und die empfohlene Vorgehensweise.' 'The table below provides an overview of typical DR scenarios and the recommended approach.')))
        $drRows = @(
            ,@((L 'Datenbankausfall (keine DAG)' 'Database failure (no DAG)'), (L 'Restore aus Backup in Recovery Database (RDB), Mailbox-Merge in Produktionsdatenbank' 'Restore from backup into Recovery Database (RDB), mailbox merge into production database'))
            ,@((L 'Datenbankausfall (DAG vorhanden)' 'Database failure (DAG present)'), (L 'Automatischer/manueller Failover auf Datenbankkopie; fehlerhafte Kopie per Update-MailboxDatabaseCopy reseed' 'Automatic/manual failover to database copy; reseed faulty copy via Update-MailboxDatabaseCopy'))
            ,@((L 'Server-Totalausfall' 'Complete server failure'), (L 'setup.exe /m:RecoverServer auf ersetztem Server; danach Datenbanken mounten bzw. DAG-Kopien reseed' 'setup.exe /m:RecoverServer on replacement server; then mount databases or reseed DAG copies'))
            ,@((L 'Verlust des File Share Witness (FSW)' 'Loss of File Share Witness (FSW)'), (L 'DAG kann noch lesen; Alternate FSW übernimmt automatisch (wenn konfiguriert). Manuell: Set-DatabaseAvailabilityGroup -AlternateWitnessServer' 'DAG can still read; Alternate FSW takes over automatically (if configured). Manually: Set-DatabaseAvailabilityGroup -AlternateWitnessServer'))
            ,@((L 'Active-Directory-Ausfall' 'Active Directory failure'), (L 'Exchange kann ohne AD nicht starten (Ausnahme: Edge Transport). AD-Wiederherstellung hat Vorrang.' 'Exchange cannot start without AD (exception: Edge Transport). AD recovery takes priority.'))
        )
        # ColWidths: short scenario name 2500, long procedure text (up to 139 chars) 6760 — total 9260 twips
        $null = $parts.Add((New-WdTable -Headers @((L 'Szenario' 'Scenario'), (L 'Vorgehensweise' 'Procedure')) -Rows $drRows -ColWidths @(2500, 6760)))

        # 10.4 Backup-Nachweis und Testzyklus
        $null = $parts.Add((New-WdHeading (L '10.4 Backup-Nachweis und Testzyklus' '10.4 Backup Evidence and Test Cycle') 2))
        $null = $parts.Add((New-WdParagraph (L 'Für Auditierbarkeit muss dokumentiert sein, dass Backups regelmäßig durchgeführt und getestet werden. Bitte nach Abschluss der ersten Produktionsbackups und nach jedem Restore-Test ausfüllen.' 'For auditability it must be documented that backups are performed and tested regularly. Please complete after the first production backups and after each restore test.')))
        $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert / Datum' 'Value / Date')) -Rows @(
            ,@((L 'Backup-Lösung (Produkt)' 'Backup solution (product)'), '')
            ,@((L 'Erstes erfolgreiches Backup' 'First successful backup'), '')
            ,@((L 'Backup-Frequenz' 'Backup frequency'), '')
            ,@((L 'Aufbewahrungsdauer Backups' 'Backup retention period'), '')
            ,@((L 'Letzter Restore-Test (Datum)' 'Last restore test (date)'), '')
            ,@((L 'Restore-Test durchgeführt von' 'Restore test performed by'), '')
            ,@((L 'Restore-Ergebnis' 'Restore result'), '')
            ,@((L 'Nächster Restore-Test geplant' 'Next restore test planned'), '')
        )))

        # ── 11. HealthChecker ──────────────────────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '11. HealthChecker' '11. HealthChecker') 1))
        $null = $parts.Add((New-WdParagraph (L 'Der Microsoft CSS Exchange HealthChecker ist ein offizielles Diagnoseskript des Microsoft Exchange-Teams (https://aka.ms/ExchangeHealthChecker). Er prüft den Exchange-Server auf bekannte Konfigurationsprobleme, fehlende Sicherheitsupdates, falsche Registry-Einstellungen, TLS-Konfiguration, Zertifikatsprobleme, OS-Konfiguration und Performance-Indikatoren. Der HealthChecker wird am Ende jeder EXpress-Installation automatisch ausgeführt. Das Ergebnis sollte nach der Installation gesichtet und offene Findings abgearbeitet werden.' 'The Microsoft CSS Exchange HealthChecker is an official diagnostic script from the Microsoft Exchange team (https://aka.ms/ExchangeHealthChecker). It checks the Exchange server for known configuration issues, missing security updates, incorrect registry settings, TLS configuration, certificate issues, OS configuration and performance indicators. HealthChecker is automatically executed at the end of every EXpress installation. The result should be reviewed after installation and any open findings addressed.')))
        $hcPath = SafeVal $State['HCReportPath']
        if ($hcPath) {
            $null = $parts.Add((New-WdParagraph ((L 'HealthChecker HTML-Report (generiert während der Installation): ' 'HealthChecker HTML report (generated during installation): ') + $hcPath)))
        } else {
            $null = $parts.Add((New-WdParagraph (L 'HealthChecker wurde nicht ausgeführt oder der Report-Pfad ist nicht verfügbar. Bitte manuell ausführen: https://aka.ms/ExchangeHealthChecker' 'HealthChecker was not run or the report path is not available. Please run manually: https://aka.ms/ExchangeHealthChecker')))
        }

        # ── 12. Monitoring-Readiness ───────────────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '12. Monitoring-Readiness' '12. Monitoring Readiness') 1))
        $null = $parts.Add((New-WdParagraph (L 'Exchange Server enthält mit Managed Availability ein eingebautes Selbstheilungssystem, das Komponenten überwacht und bei Fehler automatisch Recover-Aktionen auslöst (Dienst-Neustart, IIS-Reset, Server-Failover). Managed Availability ersetzt jedoch kein aktives externes Monitoring. Für den produktiven Betrieb wird ein dediziertes Monitoring-System empfohlen, das Exchange-spezifische Metriken, Event-Log-Einträge und Service-Zustände überwacht.' 'Exchange Server includes Managed Availability, a built-in self-healing system that monitors components and automatically triggers recovery actions on failure (service restart, IIS reset, server failover). However, Managed Availability does not replace active external monitoring. A dedicated monitoring system is recommended for production operation, monitoring Exchange-specific metrics, event log entries and service states.')))
        $monRows = [System.Collections.Generic.List[object[]]]::new()
        try { $svc2 = Get-Service MSExchangeMitigation -ErrorAction SilentlyContinue; if ($svc2) { $monRows.Add(@('EEMS (MSExchangeMitigation)', $svc2.Status.ToString())) } } catch { Write-MyVerbose ('Get-Service MSExchangeMitigation failed: {0}' -f $_) }
        try {
            $evtLogs2 = @('Application','System','MSExchange Management') | ForEach-Object {
                try { '{0}: MaxSize={1}MB' -f $_, [math]::Round((Get-WinEvent -ListLog $_ -ErrorAction Stop).MaximumSizeInBytes / 1MB, 0) } catch { Write-MyVerbose ('Get-WinEvent -ListLog {0} failed: {1}' -f $_, $_) }
            } | Where-Object { $_ }
            if ($evtLogs2) { $monRows.Add(@((L 'Event-Log-Größen' 'Event log sizes'), ($evtLogs2 -join '; '))) }
        } catch { Write-MyVerbose ('Event log size query failed: {0}' -f $_) }
        if ($monRows.Count -eq 0) { $monRows.Add(@((L '(keine Daten abrufbar)' '(no data available)'), '')) }
        $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert / Status' 'Value / status')) -Rows $monRows.ToArray()))
        $null = $parts.Add((New-WdParagraph (L 'Empfehlungen für das Monitoring nach Go-Live:' 'Recommendations for monitoring after go-live:')))
        $null = $parts.Add((New-WdBullet (L 'Perfmon-Baseline innerhalb von 4 Wochen nach Go-Live aufzeichnen (MSExchangeIS, RPC-Latenz, Disk-Queue, CPU)' 'Record Perfmon baseline within 4 weeks of go-live (MSExchangeIS, RPC latency, disk queue, CPU)')))
        $null = $parts.Add((New-WdBullet (L 'Event-IDs überwachen: 1009 (MSExchangeIS), 2142/2144 (RPC-Latenz), 4999 (Watson), 1022 (Datenbankfehler)' 'Monitor event IDs: 1009 (MSExchangeIS), 2142/2144 (RPC latency), 4999 (Watson), 1022 (database errors)')))
        $null = $parts.Add((New-WdBullet (L 'Exchange-Zertifikatsablauf überwachen — Auth-Zertifikat (2 Jahre) und IIS/SMTP-Zertifikat (kundenabhängig). MEAC-Scheduled-Task übernimmt Auth-Cert-Erneuerung automatisch.' 'Monitor Exchange certificate expiry — Auth certificate (2 years) and IIS/SMTP certificate (customer-dependent). MEAC scheduled task handles Auth Cert renewal automatically.')))
        $null = $parts.Add((New-WdBullet (L 'Datenbankkopienstatus (DAG): Get-MailboxDatabaseCopyStatus täglich prüfen oder per Monitoring automatisieren' 'Database copy status (DAG): Check Get-MailboxDatabaseCopyStatus daily or automate via monitoring')))

        # 12.1 Exchange Crimson Event Log Channels
        $null = $parts.Add((New-WdHeading (L '12.1 Exchange Crimson Event Log Kanäle' '12.1 Exchange Crimson Event Log Channels') 2))
        $null = $parts.Add((New-WdParagraph (L 'Exchange schreibt strukturierte Ereignisdaten in dedizierte Windows-Ereigniskanäle ("Crimson Channels") unterhalb von Microsoft-Exchange-*. Diese Kanäle sind feingranularer als das Application-Protokoll und ermöglichen gezieltes Monitoring einzelner Exchange-Subsysteme. Die folgende Tabelle zeigt alle aktivierten oder bereits beschriebenen Exchange-Ereigniskanäle auf diesem Server.' 'Exchange writes structured event data to dedicated Windows event channels ("Crimson channels") under Microsoft-Exchange-*. These channels are more granular than the Application log and allow targeted monitoring of individual Exchange subsystems. The table below shows all enabled or already written Exchange event channels on this server.')))
        $crimsonRows = [System.Collections.Generic.List[object[]]]::new()
        try {
            $exchLogs = @(Get-WinEvent -ListLog 'Microsoft-Exchange*' -ErrorAction SilentlyContinue |
                Where-Object { ($_.IsEnabled -or $_.RecordCount -gt 0) -and $_.LogName -match '/Operational$|/Admin$' } |
                Sort-Object LogName)
            foreach ($log in $exchLogs) {
                $sizeMB   = if ($log.MaximumSizeInBytes -gt 0) { '{0} MB' -f [math]::Round($log.MaximumSizeInBytes / 1MB, 0) } else { '—' }
                $records  = if ($log.RecordCount -gt 0) { $log.RecordCount.ToString() } else { '0' }
                # NOTE: $logState not $state/$State — PowerShell is case-insensitive; $state would shadow the outer $State hashtable.
                $logState = if ($log.IsEnabled) { (L 'aktiv' 'enabled') } else { (L 'inaktiv' 'disabled') }
                $crimsonRows.Add(@($log.LogName, $logState, $sizeMB, $records))
            }
        } catch { Write-MyVerbose ('Get-WinEvent -ListLog Microsoft-Exchange* failed: {0}' -f $_) }
        if ($crimsonRows.Count -eq 0) { $crimsonRows.Add(@((L '(keine Kanäle gefunden oder WinEvent nicht verfügbar)' '(no channels found or WinEvent not available)'), '', '', '')) }
        $null = $parts.Add((New-WdTable -Headers @((L 'Kanal' 'Channel'), (L 'Status' 'State'), (L 'Max. Größe' 'Max size'), (L 'Einträge' 'Records')) -Rows $crimsonRows.ToArray()))
        $null = $parts.Add((New-WdParagraph (L 'Wichtige Kanäle für das Exchange-Monitoring: Microsoft-Exchange-HighAvailability/Operational (DAG-Failover), Microsoft-Exchange-ManagedAvailability/Monitoring (Selbstheilung), Microsoft-Exchange-Store Driver/Operational (Mailbox-Speicher), Microsoft-Exchange-Transport/Operational (Mailflow). Für historische Fehlersuche: Get-WinEvent -LogName "Microsoft-Exchange-*" -MaxEvents 1000.' 'Key channels for Exchange monitoring: Microsoft-Exchange-HighAvailability/Operational (DAG failover), Microsoft-Exchange-ManagedAvailability/Monitoring (self-healing), Microsoft-Exchange-Store Driver/Operational (mailbox store), Microsoft-Exchange-Transport/Operational (mail flow). For historical troubleshooting: Get-WinEvent -LogName "Microsoft-Exchange-*" -MaxEvents 1000.')))

        # 12.2 Monitoring-Checkliste Go-Live
        $null = $parts.Add((New-WdHeading (L '12.2 Monitoring-Checkliste Go-Live' '12.2 Monitoring Checklist Go-Live') 2))
        $null = $parts.Add((New-WdParagraph (L 'Die folgende Checkliste dokumentiert den Aufbau des produktiven Monitorings nach Go-Live. Bitte nach Einrichtung jedes Monitoring-Elements ausfüllen.' 'The checklist below documents the setup of production monitoring after go-live. Please complete after each monitoring element is configured.')))
        $null = $parts.Add((New-WdTable -Headers @((L 'Monitoring-Element' 'Monitoring element'), (L 'Tool / System' 'Tool / system'), (L 'Eingerichtet (Datum)' 'Configured (date)'), (L 'Verantwortlich' 'Owner')) -Rows @(
            ,@((L 'Exchange-Dienst-Überwachung (MSExchange*)' 'Exchange service monitoring (MSExchange*)'), '', '', '')
            ,@((L 'Zertifikatsablauf-Überwachung (IIS/SMTP)' 'Certificate expiry monitoring (IIS/SMTP)'), '', '', '')
            ,@((L 'Postfachvolumen / Datenbankgröße' 'Mailbox volume / database size'), '', '', '')
            ,@((L 'Datenbankkopien-Status (DAG)' 'Database copy status (DAG)'), '', '', '')
            ,@((L 'Mailflow-Test (eingehend + ausgehend)' 'Mail flow test (inbound + outbound)'), '', '', '')
            ,@((L 'Log-Volume-Auslastung' 'Log volume utilisation'), '', '', '')
            ,@((L 'Event-ID-Alerting (1009, 4999, 1022)' 'Event ID alerting (1009, 4999, 1022)'), '', '', '')
            ,@((L 'Perfmon-Baseline aufgezeichnet' 'Perfmon baseline recorded'), '', '', '')
            ,@('HealthChecker (nach jedem SU/CU)', '', '', '')
        )))

        # ── 13. Public Folders ─────────────────────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '13. Öffentliche Ordner' '13. Public Folders') 1))
        $null = $parts.Add((New-WdParagraph (L 'Öffentliche Ordner (Public Folders) sind eine Legacy-Kollaborationsfunktion in Exchange, die seit Exchange 2013 auf Postfach-Infrastruktur (Public Folder Mailboxes) umgestellt wurde ("Modern Public Folders"). Sie ermöglichen gemeinsamen Zugriff auf E-Mail, Kalender, Kontakte und Dateien in einer Ordnerhierarchie, die allen Benutzern oder ausgewählten Gruppen zugänglich ist. In modernen Umgebungen werden Public Folders zunehmend durch Shared Mailboxes (gemeinsamer Posteingang, geteilter Kalender) und Microsoft Teams/SharePoint (Dokumentenablage, Teamzusammenarbeit) abgelöst. Microsoft hat mehrfach die Abkündigung von Public Folders angekündigt und empfiehlt für neue Implementierungen ausschließlich die modernen Alternativen.' 'Public Folders are a legacy collaboration feature in Exchange that has been migrated to mailbox infrastructure (Public Folder Mailboxes) since Exchange 2013 ("Modern Public Folders"). They allow shared access to email, calendars, contacts and files in a folder hierarchy accessible to all users or selected groups. In modern environments, Public Folders are increasingly replaced by Shared Mailboxes (shared inbox, shared calendar) and Microsoft Teams/SharePoint (document storage, team collaboration). Microsoft has announced the deprecation of Public Folders multiple times and recommends only the modern alternatives for new implementations.')))
        $null = $parts.Add((New-WdParagraph (L 'Hinweis zur Migration: Öffentliche Ordner können nach Exchange Online migriert werden (Migration zu EXO Modern Public Folders). Alternativ können Inhalte in Shared Mailboxes oder SharePoint-Dokumentbibliotheken überführt werden. Für die Migration zu EXO ist das Skript-Paket unter https://aka.ms/publicfoldermigration verfügbar.' 'Migration note: Public Folders can be migrated to Exchange Online (migration to EXO Modern Public Folders). Alternatively, contents can be transferred to Shared Mailboxes or SharePoint document libraries. The script package for migration to EXO is available at https://aka.ms/publicfoldermigration.')))
        try {
            $pfMailboxes = @(Get-Mailbox -PublicFolder -ErrorAction SilentlyContinue)
            if ($pfMailboxes -and $pfMailboxes.Count -gt 0) {
                $null = $parts.Add((New-WdParagraph (L 'Folgende Public-Folder-Postfächer sind in der Organisation konfiguriert:' 'The following Public Folder mailboxes are configured in the organisation:')))
                $pfRows = $pfMailboxes | ForEach-Object { @($_.Name, (SafeVal $_.PrimarySmtpAddress), (SafeVal $_.Database), (SafeVal $_.IsRootPublicFolderMailbox)) }
                $null = $parts.Add((New-WdTable -Headers @((L 'Name' 'Name'), 'SMTP', (L 'Datenbank' 'Database'), (L 'Root-PF-Postfach' 'Root PF mailbox')) -Rows $pfRows))
                try {
                    $pfStats = Get-PublicFolderStatistics -ErrorAction SilentlyContinue | Measure-Object -Property ItemCount, TotalItemSize -Sum
                    if ($pfStats) {
                        $pfCountRow = [System.Collections.Generic.List[object[]]]::new()
                        $pfCountRow.Add(@((L 'Anzahl Öffentliche Ordner (gesamt)' 'Total public folder count'), (SafeVal ($pfStats | Where-Object { $_.Property -eq 'ItemCount' } | Select-Object -ExpandProperty Sum))))
                        $null = $parts.Add((New-WdTable -Headers @((L 'Statistik' 'Statistic'), (L 'Wert' 'Value')) -Rows $pfCountRow.ToArray()))
                    }
                } catch { Write-MyVerbose ('Get-PublicFolderStatistics failed: {0}' -f $_) }
            } else {
                $null = $parts.Add((New-WdParagraph (L 'Öffentliche Ordner sind in dieser Organisation nicht konfiguriert. Es sind keine Public-Folder-Postfächer vorhanden.' 'Public Folders are not configured in this organisation. No Public Folder mailboxes exist.')))
            }
        } catch {
            $null = $parts.Add((New-WdParagraph (L 'Abfrage nicht möglich (Edge/Management-Tools-Modus oder keine Exchange-Session).' 'Query not possible (Edge/Management Tools mode or no Exchange session).')))
        }

        # ── 14. Ausgeführte Konfigurationsbefehle (nur bei tatsächlichem Setup-Lauf) ──
        # Chronological list of the config-level cmdlets the script actually ran
        # during this installation (recorded via Register-ExecutedCommand). Covers
        # Virtual Directory URLs, antispam config, relay connectors, certificate
        # import/enable, DAG join, send-connector source updates, and scheduled
        # tasks. Low-level hardening (registry/Schannel/services) is described in
        # the preceding chapters and is not repeated here to keep the list readable.
        if (-not $isAdHoc) {
            $null = $parts.Add((New-WdHeading (L '14. Ausgeführte Konfigurationsbefehle' '14. Executed configuration commands') 1))
            $execCmds = @()
            if ($State.ContainsKey('ExecutedCommands') -and $State['ExecutedCommands']) {
                $execCmds = @($State['ExecutedCommands'])
            }
            if ($execCmds.Count -eq 0) {
                $null = $parts.Add((New-WdParagraph (L 'Während dieses Laufs wurden keine Konfigurationsbefehle aufgezeichnet (z. B. reiner Tools-Modus oder Lauf ohne Namespace/Zertifikat/DAG).' 'No configuration commands were recorded during this run (e.g. tools-only mode or run without namespace/certificate/DAG).')))
            }
            else {
                $null = $parts.Add((New-WdParagraph (L 'Die folgenden Befehle wurden in chronologischer Reihenfolge mit der angegebenen Syntax ausgeführt. Passwörter und Secure-Strings sind durch Platzhalter ersetzt.' 'The following commands were executed in chronological order with the shown syntax. Passwords and secure strings are replaced by placeholders.')))
                # Render as a table grouped by category: #, Category, Command
                # A table with text-wrap wraps long cmdlets correctly; code blocks overflow the page.
                $cmdRows = [System.Collections.Generic.List[object[]]]::new()
                $cmdNum  = 0
                $byCat   = $execCmds | Group-Object -Property Category | Sort-Object Name
                foreach ($g in $byCat) {
                    $catLabel = if ($g.Name) { $g.Name } else { (L 'Sonstige' 'Other') }
                    foreach ($e in $g.Group) {
                        foreach ($cmd in ($e.Command -split '; ')) {
                            $cmdNum++
                            $cmdRows.Add(@([string]$cmdNum, $catLabel, $cmd.Trim()))
                        }
                    }
                }
                # ColWidths: # 360, Category 1500, Command 7140 — total 9000 twips (fits A4 1-inch margins)
                # -Compact: 8pt font gives ~40% more characters per line for long cmdlet strings.
                $null = $parts.Add((New-WdTable -Compact -Headers @('#', (L 'Kategorie' 'Category'), (L 'Befehl' 'Command')) -Rows $cmdRows.ToArray() -ColWidths @(360, 1500, 7140)))
            }
            $null = $parts.Add((New-WdParagraph (L 'Die vollständige Installationsausgabe (inkl. Statusmeldungen, Versionsprüfungen und Fehlern) steht in der EXpress-Logdatei (siehe Kapitel 1 "Dokumenteigenschaften" → "Logdatei").' 'The complete installation output (including status messages, version checks, and errors) is available in the EXpress log file (see chapter 1 "Document Properties" → "Log file").' )))
        }

        # ── 15. Exchange Online und Microsoft 365 (promoted from former §4.17) ─────
        # Placed here, directly before the runbooks, so hybrid/EXO considerations are
        # read together with day-2 operations rather than buried inside §4 org-config.
        $null = $parts.Add((New-WdHeading (L '15. Exchange Online und Microsoft 365' '15. Exchange Online and Microsoft 365') 1))
        $null = $parts.Add((New-WdParagraph (L 'Exchange Online (EXO) ist die cloud-gehostete E-Mail-Plattform in Microsoft 365. In Hybrid-Szenarien koexistieren Exchange Server on-premises und Exchange Online — Postfächer können auf beiden Plattformen liegen, E-Mails werden plattformübergreifend weitergeleitet (Shared Namespace), und Benutzer erfahren keine funktionalen Unterschiede. Der Hybrid Configuration Wizard (HCW) richtet die notwendigen Konnektoren, Zertifikate und OAuth-Vertrauensbeziehungen ein.' 'Exchange Online (EXO) is the cloud-hosted email platform in Microsoft 365. In hybrid scenarios, Exchange Server on-premises and Exchange Online coexist — mailboxes can reside on either platform, emails are routed across platforms (Shared Namespace), and users experience no functional differences. The Hybrid Configuration Wizard (HCW) sets up the necessary connectors, certificates and OAuth trust relationships.')))
        $null = $parts.Add((New-WdParagraph (L 'Folgende Aspekte sind in Hybrid-Umgebungen besonders zu beachten:' 'The following aspects are particularly important in hybrid environments:')))
        $null = $parts.Add((New-WdBullet (L 'Mailflow-Routing: In Centralised Mail Transport (CMT) läuft alle E-Mail über den on-premises-Server — ideal für Compliance/Archivierung. In dezentralem Routing sendet EXO direkt. CMT verursacht höhere Latenz und Abhängigkeit vom on-premises-System.' 'Mail flow routing: In Centralised Mail Transport (CMT) all email passes through the on-premises server — ideal for compliance/archiving. In decentralised routing EXO sends directly. CMT causes higher latency and dependency on the on-premises system.')))
        $null = $parts.Add((New-WdBullet (L 'Free/Busy-Integration: Verfügbarkeitsanzeige zwischen on-premises- und EXO-Postfächern erfordert funktionierende OAuth/Federation-Vertrauensbeziehung (Get-FederationTrust, Get-IntraOrganizationConnector). Bei Fehler sehen Benutzer "Keine Informationen" für cloud-Kalender.' 'Free/Busy integration: Availability display between on-premises and EXO mailboxes requires a functioning OAuth/Federation trust (Get-FederationTrust, Get-IntraOrganizationConnector). On failure, users see "No information" for cloud calendars.')))
        $null = $parts.Add((New-WdBullet (L 'Postfach-Migration (Move Request): Postfächer werden über New-MoveRequest zwischen on-premises und EXO bewegt. MRSProxy-Endpunkt muss auf dem on-premises-CAS extern erreichbar sein (TCP 443, mrsProxy.svc).' 'Mailbox migration (Move Request): Mailboxes are moved between on-premises and EXO via New-MoveRequest. MRSProxy endpoint must be externally reachable on the on-premises CAS (TCP 443, mrsProxy.svc).')))
        $null = $parts.Add((New-WdBullet (L 'Exchange Online Protection (EOP) / Defender for Office 365: In Hybrid-Szenarien ist EOP für eingehende E-Mails aus dem Internet der primäre Schutz. On-premises Anti-Spam-Filter (Content Filter, Sender Filter) werden typischerweise deaktiviert, da EOP/MDO die Filterung bereits vollständig übernimmt.' 'Exchange Online Protection (EOP) / Defender for Office 365: In hybrid scenarios, EOP is the primary protection for inbound email from the internet. On-premises anti-spam filters (Content Filter, Sender Filter) are typically disabled as EOP/MDO already performs complete filtering.')))
        $null = $parts.Add((New-WdBullet (L 'Namespace-Planung: Alle HTTPS-Dienste (OWA, EWS, Autodiscover, MAPI) sollten über einen einzigen externen FQDN erreichbar sein, der auf den on-premises-Exchange oder einen vorgelagerten Reverse-Proxy zeigt. EXO-Benutzer nutzen denselben Autodiscover-FQDN; der SCP-Record im AD ist für interne Clients maßgebend.' 'Namespace planning: All HTTPS services (OWA, EWS, Autodiscover, MAPI) should be reachable via a single external FQDN pointing to the on-premises Exchange or a reverse proxy. EXO users use the same Autodiscover FQDN; the SCP record in AD is authoritative for internal clients.')))
        $null = $parts.Add((New-WdBullet (L 'Lizenzierung: Exchange Online-Postfächer benötigen eine M365-Lizenz mit Exchange Online-Plan (F1, E1, E3, E5). On-premises-Postfächer benötigen Exchange Server-CALs (Standard/Enterprise). In Hybrid-Szenarien dürfen keine EXO-Lizenzen für on-premises-Postfächer zugewiesen werden.' 'Licensing: Exchange Online mailboxes require an M365 licence with an Exchange Online plan (F1, E1, E3, E5). On-premises mailboxes require Exchange Server CALs (Standard/Enterprise). In hybrid scenarios, EXO licences must not be assigned to on-premises mailboxes.')))
        if ($scope -in 'All','Org','Local' -and $orgD -and $orgD.HybridConfig) {
            $hyb3 = $orgD.HybridConfig
            $eo365Rows = [System.Collections.Generic.List[object[]]]::new()
            $eo365Rows.Add(@((L 'Hybrid-Konfiguration' 'Hybrid configuration'), (L 'Aktiv — Hybrid Configuration Wizard wurde ausgeführt' 'Active — Hybrid Configuration Wizard has been run')))
            if ($hyb3.OnPremisesSMTPDomains) { $eo365Rows.Add(@((L 'Freigegebene SMTP-Domänen' 'Shared SMTP domains'), ($hyb3.OnPremisesSMTPDomains -join ', '))) }
            if ($hyb3.Features) { $eo365Rows.Add(@((L 'HCW-Features' 'HCW features'), ($hyb3.Features -join ', '))) }
            $null = $parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $eo365Rows.ToArray()))
        } else {
            $null = $parts.Add((New-WdParagraph (L 'Hybrid Configuration Wizard wurde (noch) nicht ausgeführt — diese Exchange-Umgebung ist rein on-premises. Für eine spätere Migration zu Exchange Online ist der HCW der empfohlene Einstiegspunkt: https://aka.ms/HybridWizard' 'Hybrid Configuration Wizard has not (yet) been run — this Exchange environment is purely on-premises. For a later migration to Exchange Online, HCW is the recommended entry point: https://aka.ms/HybridWizard')))
        }

        # ── 16. Abnahmetest / Funktionsnachweis ───────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '16. Abnahmetest und Funktionsnachweis' '16. Acceptance Testing and Functional Verification') 1))
        $null = $parts.Add((New-WdParagraph (L 'Nach Abschluss der Installation sind die folgenden Funktions- und Akzeptanztests durchzuführen und zu dokumentieren. Die Testergebnisse dienen als Nachweis für die formale Abnahme des Systems (vgl. Kapitel 1.1 Freigabe und Change-Management). Bitte Ergebnis und Datum nach jedem Test eintragen.' 'After completing the installation, the following functional and acceptance tests must be performed and documented. The test results serve as evidence for the formal acceptance of the system (cf. chapter 1.1 Sign-off and Change Management). Please enter result and date after each test.')))
        # Build OWA / ECP / EWS / Autodiscover URLs from namespace if available
        $nsBase = if ($State['Namespace']) { 'https://' + $State['Namespace'] } else { 'https://<Namespace>' }
        $null = $parts.Add((New-WdTable -Headers @((L 'Testfall' 'Test case'), (L 'Prüfpunkt' 'Check'), (L 'Ergebnis' 'Result'), (L 'Datum / Tester' 'Date / Tester')) -Rows @(
            ,@('OWA',         ('{0}/owa — Login mit Testpostfach / Login with test mailbox' -f $nsBase),                              '', '')
            ,@('ECP',         ('{0}/ecp — Admin-Login, Postfach erstellen / Admin login, create mailbox' -f $nsBase),                '', '')
            ,@('EWS',         ('{0}/ews/exchange.asmx — HTTP 200 / 401' -f $nsBase),                                                '', '')
            ,@('Autodiscover', ('{0}/autodiscover/autodiscover.xml — HTTP 200 / 401' -f $nsBase),                                   '', '')
            ,@((L 'SMTP eingehend' 'SMTP inbound'),  (L 'Testmail an internes Postfach senden (extern → Exchange)' 'Send test mail to internal mailbox (external → Exchange)'),   '', '')
            ,@((L 'SMTP ausgehend' 'SMTP outbound'), (L 'Testmail vom Exchange nach extern senden' 'Send test mail from Exchange to external'),           '', '')
            ,@('MAPI/HTTP',     (L 'Outlook-Client verbinden (Autodiscover, kein TCP 135 erforderlich)' 'Connect Outlook client (Autodiscover, no TCP 135 required)'), '', '')
            ,@('ActiveSync',    (L 'Mobiles Gerät verbinden (EAS, HTTPS 443)' 'Connect mobile device (EAS, HTTPS 443)'),             '', '')
            ,@((L 'Zertifikat' 'Certificate'), (L 'TLS-Zertifikat gültig, kein Browser-Warning' 'TLS certificate valid, no browser warning'), '', '')
            ,@('DAG',           (L 'DAG-Datenbankkopien-Status: alle Healthy / Mounted' 'DAG database copy status: all Healthy / Mounted'), '', '')
            ,@('Backup',        (L 'Erstes VSS-Backup erfolgreich, Logs abgeschnitten' 'First VSS backup successful, logs truncated'), '', '')
            ,@('HealthChecker',  (L 'Keine kritischen Findings (Reds)' 'No critical findings (Reds)'),                               '', '')
        )))

        # ── 17. Operative Runbooks ─────────────────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '17. Operative Runbooks' '17. Operational Runbooks') 1))
        $null = $parts.Add((New-WdParagraph (L 'Dieses Kapitel enthält vorgefertigte Befehlssequenzen für die häufigsten operativen Aufgaben auf Exchange Server. Die Befehle sind in der Exchange Management Shell (EMS) auszuführen, sofern nicht anders angegeben. Platzhalter (<Server>, <DB> etc.) sind vor der Ausführung durch die tatsächlichen Werte zu ersetzen.' 'This chapter contains pre-built command sequences for the most common operational tasks on Exchange Server. Commands are to be executed in the Exchange Management Shell (EMS) unless otherwise stated. Placeholders (<Server>, <DB>, etc.) must be replaced with actual values before execution.')))
        $null = $parts.Add((New-WdHeading (L '17.1 DAG-Wartungsmodus' '17.1 DAG Maintenance Mode') 2))
        $null = $parts.Add((New-WdParagraph (L 'Vor Wartungsarbeiten (Patches, Hardwarearbeiten) an einem DAG-Mitglied muss der Server in den Wartungsmodus versetzt werden. Dies löst einen kontrollierten Failover aller aktiven Datenbanken auf andere DAG-Mitglieder aus und verhindert, dass während der Wartung neue Datenbanken aktiviert werden.' 'Before maintenance work (patches, hardware work) on a DAG member, the server must be placed in maintenance mode. This triggers a controlled failover of all active databases to other DAG members and prevents new databases from being activated during maintenance.')))
        $null = $parts.Add((New-WdCode 'Set-ServerComponentState <Server> -Component ServerWideOffline -State Inactive -Requester Maintenance'))
        $null = $parts.Add((New-WdCode 'Suspend-MailboxDatabaseCopy <DB>\<Server> -SuspendComment "Wartung"'))
        $null = $parts.Add((New-WdCode '# Wartungsarbeiten durchführen'))
        $null = $parts.Add((New-WdCode 'Resume-MailboxDatabaseCopy <DB>\<Server>'))
        $null = $parts.Add((New-WdCode 'Set-ServerComponentState <Server> -Component ServerWideOffline -State Active -Requester Maintenance'))
        $null = $parts.Add((New-WdHeading (L '17.2 Cumulative Update / Security Update installieren' '17.2 Install Cumulative Update / Security Update') 2))
        $null = $parts.Add((New-WdParagraph (L 'Exchange-Updates (CU und SU) müssen als lokaler Administrator oder als SYSTEM-Konto ausgeführt werden. Empfohlen wird die Ausführung über einen geplanten Task als SYSTEM (PSExec oder Task Scheduler). Vor dem Update: DAG-Wartungsmodus aktivieren, Backup erstellen, Health-Checker-Baseline sichern. Nach dem Update: Health-Checker erneut ausführen.' 'Exchange updates (CU and SU) must be executed as local administrator or SYSTEM account. Execution via a scheduled task as SYSTEM (PSExec or Task Scheduler) is recommended. Before the update: enable DAG maintenance mode, create backup, save HealthChecker baseline. After the update: run HealthChecker again.')))
        $null = $parts.Add((New-WdCode '# Als SYSTEM (PSExec): psexec -s setup.exe ...'))
        $null = $parts.Add((New-WdCode 'setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF /PrepareAllDomains'))
        $null = $parts.Add((New-WdCode 'setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF /Mode:Upgrade'))
        $null = $parts.Add((New-WdHeading (L '17.3 Zertifikatstausch' '17.3 Certificate Replacement') 2))
        $null = $parts.Add((New-WdParagraph (L 'Exchange-Zertifikate (IIS, SMTP) laufen typischerweise nach 1–3 Jahren ab. Der Tausch muss auf allen Exchange-Servern der Organisation durchgeführt werden. Das Auth-Zertifikat (OAuth) wird durch den MEAC-Scheduled-Task automatisch 60 Tage vor Ablauf erneuert und erfordert keinen manuellen Eingriff.' 'Exchange certificates (IIS, SMTP) typically expire after 1–3 years. The replacement must be performed on all Exchange servers in the organisation. The Auth certificate (OAuth) is automatically renewed 60 days before expiry by the MEAC scheduled task and does not require manual intervention.')))
        $null = $parts.Add((New-WdCode 'Import-ExchangeCertificate -FileName <pfx> -Password (ConvertTo-SecureString <pwd> -AsPlainText -Force) -Server <srv>'))
        $null = $parts.Add((New-WdCode 'Enable-ExchangeCertificate -Thumbprint <tp> -Services IIS,SMTP -Server <srv> -Confirm:$false'))
        $null = $parts.Add((New-WdHeading (L '17.4 Aktive Datenbank verschieben (Failover)' '17.4 Move Active Database (Failover)') 2))
        $null = $parts.Add((New-WdParagraph (L 'Manueller Failover einer aktiven Datenbankkopie auf ein anderes DAG-Mitglied — z. B. vor Wartungsarbeiten oder zur Lastverteilung.' 'Manual failover of an active database copy to another DAG member — e.g. before maintenance or for load balancing.')))
        $null = $parts.Add((New-WdCode 'Move-ActiveMailboxDatabase <DB> -ActivateOnServer <TargetServer> -Confirm:$false'))
        $null = $parts.Add((New-WdCode 'Get-MailboxDatabaseCopyStatus <DB>\* | Select Name, Status, CopyQueueLength, ReplayQueueLength'))
        $null = $parts.Add((New-WdHeading (L '17.5 Datenbankkopie neu erstellen (Reseed)' '17.5 Reseed Database Copy') 2))
        $null = $parts.Add((New-WdParagraph (L 'Wenn eine passive Datenbankkopie in einem DAG stark in Verzug geraten ist oder beschädigt wurde, kann sie neu erstellt (reseeded) werden. Der Reseed kopiert die aktive Datenbank vollständig auf das Ziel-DAG-Mitglied.' 'If a passive database copy in a DAG has fallen significantly behind or been corrupted, it can be reseeded. The reseed fully copies the active database to the target DAG member.')))
        $null = $parts.Add((New-WdCode 'Update-MailboxDatabaseCopy <DB>\<Server> -DeleteExistingFiles'))
        $null = $parts.Add((New-WdCode 'Get-MailboxDatabaseCopyStatus <DB>\<Server>  # Status verfolgen / monitor status'))
        $null = $parts.Add((New-WdHeading (L '17.6 Server wiederherstellen (RecoverServer)' '17.6 Recover Server') 2))
        $null = $parts.Add((New-WdParagraph (L 'Bei einem vollständigen Serverausfall ohne DAG-Redundanz kann Exchange auf einem neuen Server mit denselben Eigenschaften (Name, IP) wiederhergestellt werden. Voraussetzung: AD-Computerkonto noch vorhanden, Exchange-Datenbanken aus Backup verfügbar.' 'In case of a complete server failure without DAG redundancy, Exchange can be restored on a new server with the same properties (name, IP). Prerequisite: AD computer account still exists, Exchange databases available from backup.')))
        $null = $parts.Add((New-WdCode 'setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF /m:RecoverServer'))

        # ── 18. Offene Punkte ──────────────────────────────────────────────────────
        $null = $parts.Add((New-WdHeading (L '18. Offene Punkte' '18. Open Items') 1))
        # Comma operator prefix prevents PS 5.1 from flattening the jagged array when
        # binding to [object[][]]; without it Rows becomes a flat 15-element array.
        $null = $parts.Add((New-WdTable -Headers @('Nr.', (L 'Offener Punkt' 'Open item'), (L 'Verantwortlich' 'Owner'), (L 'Fällig am' 'Due date'), (L 'Status' 'Status')) -Rows @(
            ,@('1', '', '', '', '')
            ,@('2', '', '', '', '')
            ,@('3', '', '', '', '')
        )))

        # Write document
        $headerLabel = if ($DE) { 'EXCHANGE SERVER INSTALLATIONSDOKUMENTATION' } else { 'EXCHANGE SERVER INSTALLATION DOCUMENTATION' }
        if ($useTpl) {
            # F24: inject chapter body into customer template and fill cover page tokens.
            $tplTokens = @{
                document_body  = ($parts -join '')
                Organization   = (SafeVal $State['OrganizationName'] '')
                ServerName     = $env:COMPUTERNAME
                Scenario       = $scenario
                InstallMode    = $instMode
                Version        = ((Get-Date -Format 'yyyy-MM-dd') + ' / EXpress v' + $ScriptVersion)
                DateLong       = (Get-Date -Format 'dd.MM.yyyy')
                Author         = $author
                Company        = $company
                Classification = $classification
                HeaderLabel    = $headerLabel
                DocTitle       = $docTitle
                CoverSub       = $coverSub
            }
            Write-WdFromTemplate -TemplatePath $tplPath -OutputPath $docPath -Tokens $tplTokens
        } else {
            New-WdFile -OutputPath $docPath -BodyParts $parts.ToArray() -DocTitle $docTitle -HeaderLabel $headerLabel -LogoPath $logoFile
        }
        $State['WordDocPath'] = $docPath
        Write-MyStep -Label 'Word Document' -Value $docPath -Status OK
    }

