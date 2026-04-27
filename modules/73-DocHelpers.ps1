    function Invoke-DocSection-Organisation {
        param(
            [System.Collections.Generic.List[string]]$Parts,
            [object]$ReportData,
            [bool]$DE,
            [bool]$Cust
        )
        # Rebind language/format helpers using function parameters
        function L([string]$d, [string]$e) { if ($DE) { $d } else { $e } }
        function Lc([bool]$c, [string]$a, [string]$b) { if ($c) { $a } else { $b } }
        function SafeVal([object]$v, [string]$fallback = '') { if ($null -eq $v -or "$v" -eq '') { $fallback } else { "$v" } }
        function Format-RegBool($v) {
            if ($null -eq $v -or "$v" -eq '') { return (L '(nicht gesetzt)' '(not set)') }
            if ([bool]$v) { return (L 'aktiviert' 'enabled') }
            return (L 'deaktiviert' 'disabled')
        }
        function Mask-Ip([string]$text) {
            if (-not $Cust) { return $text }
            $text -replace '\b(10|172\.(1[6-9]|2[0-9]|3[01])|192\.168)\.\d{1,3}\.\d{1,3}\b', 'x.x.x.x'
        }
        function Mask-Val([string]$text) { if ($Cust -and $text) { '[redacted]' } else { $text } }

        $orgD = $ReportData.Org

            $null = $Parts.Add((New-WdHeading (L '4. Organisation — übergreifende Konfiguration' '4. Organisation — Global Configuration') 1))
            $null = $Parts.Add((New-WdParagraph (L 'Die Exchange-Organisation umfasst alle Exchange-Server in der AD-Gesamtstruktur. Die folgenden Abschnitte dokumentieren die organisationsweiten Einstellungen, die auf alle Server und Postfächer in der Organisation wirken.' 'The Exchange organisation encompasses all Exchange servers in the AD forest. The following sections document the organisation-wide settings that apply to all servers and mailboxes in the organisation.')))

            # 4.1 Org-Übersicht
            $null = $Parts.Add((New-WdHeading (L '4.1 Org-Übersicht' '4.1 Organisation Overview') 2))
            $orgRows = [System.Collections.Generic.List[object[]]]::new()
            if ($orgD -and $orgD.OrgConfig) {
                $oc = $orgD.OrgConfig
                $orgRows.Add(@((L 'Name' 'Name'), (SafeVal $oc.Name)))
                $orgRows.Add(@((L 'Exchange-Version' 'Exchange version'), (SafeVal $oc.AdminDisplayVersion)))
                $orgRows.Add(@((L 'MAPI/HTTP' 'MAPI/HTTP'), (SafeVal $oc.MapiHttpEnabled)))
                $orgRows.Add(@((L 'Modern Auth (OAuth2)' 'Modern Auth (OAuth2)'), (SafeVal $oc.OAuth2ClientProfileEnabled)))
                $orgRows.Add(@((L 'CEIP deaktiviert' 'CEIP disabled'), (SafeVal (-not $oc.CustomerFeedbackEnabled))))
                if ($null -ne $oc.DefaultPublicFolderMailbox) { $orgRows.Add(@((L 'Standard-PF-Postfach' 'Default PF mailbox'), (SafeVal $oc.DefaultPublicFolderMailbox))) }
            }
            $null = $Parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $orgRows.ToArray()))

            # 4.2 Accepted Domains
            $null = $Parts.Add((New-WdHeading (L '4.2 Accepted Domains' '4.2 Accepted Domains') 2))
            $adDomRows = [System.Collections.Generic.List[object[]]]::new()
            foreach ($dom in $orgD.AcceptedDomains) { $adDomRows.Add(@($dom.DomainName, $dom.DomainType, (Lc $dom.Default (L 'Standard' 'Default') ''))) }
            $null = $Parts.Add((New-WdTable -Headers @((L 'Domäne' 'Domain'), (L 'Typ' 'Type'), (L 'Standard' 'Default')) -Rows $adDomRows.ToArray()))

            # 4.3 Remote Domains
            $null = $Parts.Add((New-WdHeading (L '4.3 Remote Domains' '4.3 Remote Domains') 2))
            $rdRows = [System.Collections.Generic.List[object[]]]::new()
            foreach ($rd2 in $orgD.RemoteDomains) { $rdRows.Add(@($rd2.DomainName, (SafeVal $rd2.ContentType), (Lc $rd2.AutoReplyEnabled (L 'Auto-Reply aktiv' 'Auto-reply active') ''))) }
            $null = $Parts.Add((New-WdTable -Headers @((L 'Domäne' 'Domain'), (L 'Content-Typ' 'Content type'), (L 'Hinweis' 'Note')) -Rows $rdRows.ToArray()))

            # 4.4 E-Mail-Adressrichtlinien
            $null = $Parts.Add((New-WdHeading (L '4.4 E-Mail-Adressrichtlinien' '4.4 Email Address Policies') 2))
            $eapRows = [System.Collections.Generic.List[object[]]]::new()
            foreach ($pol in $orgD.EmailAddressPolicies) { $eapRows.Add(@($pol.Name, (SafeVal $pol.RecipientFilter), (SafeVal ($pol.EnabledEmailAddressTemplates -join ', ')))) }
            $null = $Parts.Add((New-WdTable -Headers @((L 'Name' 'Name'), (L 'Empfängerfilter' 'Recipient filter'), (L 'Adressvorlagen' 'Address templates')) -Rows $eapRows.ToArray()))

            # 4.5 Transport Rules
            $null = $Parts.Add((New-WdHeading (L '4.5 Transportregeln' '4.5 Transport Rules') 2))
            $trRows = [System.Collections.Generic.List[object[]]]::new()
            foreach ($tr in $orgD.TransportRules) { $trRows.Add(@($tr.Name, $tr.State, $tr.Priority, (SafeVal $tr.Comments))) }
            if ($trRows.Count -eq 0) {
                $trRows.Add(@((L '(keine Regeln konfiguriert)' '(no rules configured)'), '', '', ''))
                $null = $Parts.Add((New-WdTable -Headers @((L 'Name' 'Name'), (L 'Status' 'State'), (L 'Priorität' 'Priority'), (L 'Kommentar' 'Comment')) -Rows $trRows.ToArray()))
                $null = $Parts.Add((New-WdParagraph (L '📋 Hinweis: Auf einer neu installierten Exchange-Organisation sind Transportregeln typischerweise noch nicht konfiguriert. Dieser Abschnitt ist nach Einrichtung der Kundenumgebung zu ergänzen (z. B. Disclaimer, Routing-Regeln, Compliance-Regeln).' '📋 Note: Transport rules are typically not yet configured on a freshly installed Exchange organisation. This section should be completed after the customer environment has been set up (e.g. disclaimers, routing rules, compliance rules).')))
            } else {
                $null = $Parts.Add((New-WdTable -Headers @((L 'Name' 'Name'), (L 'Status' 'State'), (L 'Priorität' 'Priority'), (L 'Kommentar' 'Comment')) -Rows $trRows.ToArray()))
            }

            # 4.6 Transport-Konfiguration (Org)
            $null = $Parts.Add((New-WdHeading (L '4.6 Transport-Konfiguration' '4.6 Transport Configuration') 2))
            $tcRows = [System.Collections.Generic.List[object[]]]::new()
            if ($orgD.TransportConfig) {
                $tc2 = $orgD.TransportConfig
                # MaxSendSize / MaxReceiveSize may be Unlimited ($null .Value) on a fresh org
                # or when the Exchange snap-in is unavailable. Format explicitly with a null-guard.
                $fmtSize = {
                    param($sz)
                    if ($null -eq $sz) { return (L 'nicht gesetzt' 'not set') }
                    if ($null -eq $sz.Value) { return (L 'unbegrenzt' 'Unlimited') }
                    '{0} MB' -f [math]::Round($sz.Value.ToBytes() / 1MB, 0)
                }
                $tcRows.Add(@((L 'Max. Sendegröße' 'Max send size'),    (& $fmtSize $tc2.MaxSendSize)))
                $tcRows.Add(@((L 'Max. Empfangsgröße' 'Max receive size'), (& $fmtSize $tc2.MaxReceiveSize)))
                $tcRows.Add(@('Safety Net Hold Time', (SafeVal $tc2.SafetyNetHoldTime)))
                $tcRows.Add(@((L 'HTML-NDRs (intern/extern)' 'HTML NDRs (internal/external)'), ('{0} / {1}' -f $tc2.InternalDsnSendHtml, $tc2.ExternalDsnSendHtml)))
            }
            $null = $Parts.Add((New-WdTable -Headers @((L 'Einstellung' 'Setting'), (L 'Wert' 'Value')) -Rows $tcRows.ToArray()))

            # 4.7 Journal / DLP / Retention
            $null = $Parts.Add((New-WdHeading (L '4.7 Journal-, DLP- und Aufbewahrungsrichtlinien' '4.7 Journal, DLP and Retention Policies') 2))
            $null = $Parts.Add((New-WdParagraph (L 'Journaling erfasst eine Kopie aller oder ausgewählter E-Mails an eine Compliance-Postfachadresse — häufig gesetzlich vorgeschrieben (GoBD, MiFID II, SOX). Aufbewahrungsrichtlinien (Retention Policies) steuern die automatische Verschiebung oder Löschung von E-Mails nach definierten Zeiträumen (Messaging Records Management, MRM). DLP-Richtlinien (Data Loss Prevention) erkennen sensible Inhalte (z. B. Kreditkartennummern, Ausweisdaten) in E-Mails und können diese blockieren, umleiten oder markieren. In rein on-premises-Umgebungen ohne Exchange Online ist DLP nur mit eigenem Regelwerk verfügbar; die vordefinierten Microsoft 365-Vorlagen sind auf EXO beschränkt.' 'Journaling captures a copy of all or selected emails to a compliance mailbox address — often legally required (GoBD, MiFID II, SOX). Retention policies control automatic moving or deletion of emails after defined periods (Messaging Records Management, MRM). DLP policies (Data Loss Prevention) detect sensitive content (e.g. credit card numbers, ID data) in emails and can block, redirect or tag them. In purely on-premises environments without Exchange Online, DLP is only available with a custom rule set; the predefined Microsoft 365 templates are restricted to EXO.')))
            if ($orgD.JournalRules.Count -gt 0) {
                $jRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($jr in $orgD.JournalRules) { $jRows.Add(@($jr.Name, (SafeVal $jr.JournalEmailAddress), $jr.Scope, (Lc $jr.Enabled (L 'Aktiv' 'Enabled') (L 'Inaktiv' 'Disabled')))) }
                $null = $Parts.Add((New-WdTable -Headers @((L 'Journal-Regel' 'Journal rule'), (L 'Empfänger' 'Recipient'), 'Scope', (L 'Status' 'Status')) -Rows $jRows.ToArray()))
            }
            if ($orgD.RetentionPolicies.Count -gt 0) {
                $rpRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($rp in $orgD.RetentionPolicies) { $rpRows.Add(@($rp.Name, (SafeVal ($rp.RetentionPolicyTagLinks -join ', ')))) }
                $null = $Parts.Add((New-WdTable -Headers @((L 'Aufbewahrungsrichtlinie' 'Retention policy'), (L 'Verknüpfte Tags' 'Linked tags')) -Rows $rpRows.ToArray()))
            }
            if ($orgD.RetentionPolicyTags -and $orgD.RetentionPolicyTags.Count -gt 0) {
                $null = $Parts.Add((New-WdParagraph (L 'Konfigurierte Aufbewahrungs-Tags (Retention Tags) — definieren je Postfachordner oder benutzergewählt, nach welcher Frist welche Aktion (Verschieben ins Archiv, Löschen mit/ohne Wiederherstellung, MarkAsPastRetentionLimit) ausgeführt wird:' 'Configured retention tags — define per mailbox folder or user-selectable which action (move to archive, delete with/without recovery, MarkAsPastRetentionLimit) is executed after which retention period:')))
                $rtRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($rt in ($orgD.RetentionPolicyTags | Sort-Object Type, Name)) {
                    # AgeLimitForRetention is a deserialized EnhancedTimeSpan over implicit remoting —
                    # CLR-inherited .Days returns $null on deserialized objects; parse from ToString() instead.
                    $age = if ($null -ne $rt.AgeLimitForRetention) {
                        $ageTs = try { [TimeSpan]::Parse($rt.AgeLimitForRetention.ToString()) } catch { $null }
                        if ($null -ne $ageTs) { '{0} {1}' -f [int]$ageTs.TotalDays, (L 'Tage' 'days') } else { $rt.AgeLimitForRetention.ToString() }
                    } else { (L '(unbegrenzt)' '(unlimited)') }
                    $rtRows.Add(@(
                        $rt.Name,
                        (SafeVal $rt.Type),
                        $age,
                        (SafeVal $rt.RetentionAction),
                        (Lc $rt.RetentionEnabled (L 'Aktiv' 'Enabled') (L 'Inaktiv' 'Disabled'))
                    ))
                }
                # ColWidths: Tag name 3500 (long names), Type 1200, Retention 1200, Action 1800, Status 1560 — total 9260 twips
                $null = $Parts.Add((New-WdTable -Headers @((L 'Tag-Name' 'Tag name'), (L 'Typ' 'Type'), (L 'Aufbewahrung' 'Retention'), (L 'Aktion' 'Action'), (L 'Status' 'Status')) -Rows $rtRows.ToArray() -Compact -ColWidths @(3500, 1200, 1200, 1800, 1560)))
            }
            if ($orgD.DlpPolicies.Count -gt 0) {
                $dlpRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($dp in $orgD.DlpPolicies) { $dlpRows.Add(@($dp.Name, $dp.Mode, (Lc $dp.Activated (L 'Aktiv' 'Active') (L 'Inaktiv' 'Inactive')))) }
                $null = $Parts.Add((New-WdTable -Headers @('DLP', 'Mode', (L 'Status' 'Status')) -Rows $dlpRows.ToArray()))
            }
            if ($orgD.JournalRules.Count -eq 0 -and $orgD.RetentionPolicies.Count -eq 0 -and $orgD.DlpPolicies.Count -eq 0) {
                $null = $Parts.Add((New-WdParagraph (L '(Keine Journal-, DLP- oder Aufbewahrungsregeln konfiguriert)' '(No journal, DLP or retention policies configured)')))
            }

            # 4.8 Mobile / OWA Policies
            $null = $Parts.Add((New-WdHeading (L '4.8 Mobile- und OWA-Richtlinien' '4.8 Mobile and OWA Policies') 2))
            $null = $Parts.Add((New-WdParagraph (L 'Mobile Device Mailbox Policies steuern, welche Anforderungen mobile Geräte (ActiveSync, Exchange Active Sync/EAS) für die Verbindung mit Exchange erfüllen müssen: PIN-Schutz, Geräteverschlüsselung, Passwort-Komplexität, Fernlöschung (Remote Wipe). In Hybrid-Umgebungen übernehmen Intune-MDM-Richtlinien zunehmend diese Funktion; Exchange ActiveSync bleibt für on-premises-verwaltete Geräte relevant. OWA-Richtlinien kontrollieren den Funktionsumfang in Outlook Web App: Dateianhänge, S/MIME, OneNote-Integration, Skype for Business, SharePoint-Zugriff. In Hybrid-Szenarien ist die OWA-Policy-Zuweisung zwischen on-premises und EXO-Postfächern zu synchronisieren.' 'Mobile Device Mailbox Policies control which requirements mobile devices (ActiveSync, Exchange Active Sync/EAS) must meet to connect to Exchange: PIN protection, device encryption, password complexity, remote wipe. In hybrid environments, Intune MDM policies are increasingly taking over this function; Exchange ActiveSync remains relevant for on-premises-managed devices. OWA policies control the feature scope in Outlook Web App: file attachments, S/MIME, OneNote integration, Skype for Business, SharePoint access. In hybrid scenarios, OWA policy assignment between on-premises and EXO mailboxes needs to be synchronised.')))
            if ($orgD.MobileDevicePolicies.Count -gt 0) {
                $mobRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($mp in $orgD.MobileDevicePolicies) { $mobRows.Add(@($mp.Name, (Lc $mp.IsDefault (L 'Standard' 'Default') ''), (SafeVal $mp.DevicePasswordEnabled), (SafeVal $mp.DeviceEncryptionEnabled))) }
                $null = $Parts.Add((New-WdTable -Headers @((L 'Richtlinie' 'Policy'), (L 'Standard' 'Default'), (L 'PIN erforderlich' 'PIN required'), (L 'Verschlüsselung' 'Encryption')) -Rows $mobRows.ToArray()))
            }
            if ($orgD.OwaPolicies.Count -gt 0) {
                $owaPolRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($op in $orgD.OwaPolicies) { $owaPolRows.Add(@($op.Name, (Lc $op.IsDefault (L 'Standard' 'Default') ''), (SafeVal $op.LogonFormat))) }
                $null = $Parts.Add((New-WdTable -Headers @((L 'OWA-Richtlinie' 'OWA policy'), (L 'Standard' 'Default'), (L 'Anmeldung' 'Logon format')) -Rows $owaPolRows.ToArray()))
            }

            # 4.9 DAGs (alle)
            $null = $Parts.Add((New-WdHeading (L '4.9 Database Availability Groups' '4.9 Database Availability Groups') 2))
            if ($orgD.DAGs -and $orgD.DAGs.Count -gt 0) {
                foreach ($dagEntry in $orgD.DAGs) {
                    $dag2 = $dagEntry.DAG
                    $null = $Parts.Add((New-WdHeading $dag2.Name 3))
                    $dagInfoRows = @(
                        @((L 'Mitglieder' 'Members'), ($dag2.Servers -join ', '))
                        @('FSW', (Mask-Ip (SafeVal $dag2.WitnessServer)))
                        @('Alternate FSW', (Mask-Ip (SafeVal $dag2.AlternateWitnessServer)))
                        @('DAC Mode', (SafeVal $dag2.DatacenterActivationMode))
                        @((L 'Replikationsnetz' 'Replication networks'), (SafeVal ($dag2.ReplicationDagNetwork -join ', ')))
                    )
                    $null = $Parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $dagInfoRows))
                }
                # Database copy status across all servers (consolidated view)
                $anyCopies49 = @($ReportData.Servers | ForEach-Object { $_.DatabaseCopies } | Where-Object { $_ })
                if ($anyCopies49.Count -gt 0) {
                    $null = $Parts.Add((New-WdHeading (L 'Datenbank-Kopien-Status' 'Database Copy Status') 3))
                    $null = $Parts.Add((New-WdParagraph (L 'CopyQueueLength = noch nicht replizierte Logs, ReplayQueueLength = noch nicht eingespielt. Im Normalbetrieb sollten beide Werte einstellig bleiben. ContentIndexState = "Healthy" ist Voraussetzung für die Postfachsuche.' 'CopyQueueLength = logs not yet replicated, ReplayQueueLength = logs not yet replayed. Both values should stay single-digit in normal operation. ContentIndexState = "Healthy" is required for mailbox search.')))
                    $dcRows49 = [System.Collections.Generic.List[object[]]]::new()
                    foreach ($srv2 in $ReportData.Servers) {
                        foreach ($dc in $srv2.DatabaseCopies) {
                            $dcRows49.Add(@($dc.DatabaseName, $dc.MailboxServer, $dc.Status, $dc.CopyQueueLength, $dc.ReplayQueueLength, (SafeVal $dc.ContentIndexState), (SafeVal $dc.ActivationPreference)))
                        }
                    }
                    $null = $Parts.Add((New-WdTable -Headers @((L 'Datenbank' 'Database'), (L 'Server' 'Server'), (L 'Status' 'Status'), 'Copy-Q', 'Replay-Q', (L 'Suchindex' 'Content index'), (L 'AktPref' 'ActPref')) -Rows $dcRows49.ToArray()))
                }
            } else {
                $null = $Parts.Add((New-WdParagraph (L '(Keine DAG konfiguriert — Standalone-Umgebung)' '(No DAG configured — standalone environment)')))
                $null = $Parts.Add((New-WdParagraph (L '📋 Hinweis: Bei einer neuen Installation ohne DAG ist dieser Abschnitt leer. Nach Einrichtung einer Database Availability Group ist dieser Abschnitt zu aktualisieren.' '📋 Note: This section is empty on a new installation without DAG. Update this section after a Database Availability Group has been configured.')))
            }

            # 4.10 Send Connectors
            $null = $Parts.Add((New-WdHeading (L '4.10 Send Connectors' '4.10 Send Connectors') 2))
            $scRows = [System.Collections.Generic.List[object[]]]::new()
            foreach ($sc in $orgD.SendConnectors) {
                $enabledSc  = if ($sc.Enabled) { (L 'aktiviert' 'enabled') } else { (L 'deaktiviert' 'disabled') }
                $reqTlsSc   = Lc ([bool]$sc.RequireTLS) (L 'ja' 'yes') (L 'nein' 'no')
                $maxMsgSc   = if ($sc.MaxMessageSize) { $sc.MaxMessageSize.ToString() } else { '—' }
                $scRows.Add(@($sc.Name, ($sc.AddressSpaces -join ', '), (Mask-Ip (SafeVal ($sc.SmartHosts -join ', '))), (Mask-Ip ($sc.SourceTransportServers -join ', ')), (SafeVal $sc.Fqdn '—'), $reqTlsSc, $maxMsgSc, $enabledSc))
            }
            $scEmpty = $scRows.Count -eq 0
            if ($scEmpty) { $scRows.Add(@((L '(keine konfiguriert)' '(none configured)'), '', '', '', '', '', '', '')) }
            $null = $Parts.Add((New-WdTable -Headers @((L 'Name' 'Name'), (L 'Adressraum' 'Address space'), 'Smarthost', (L 'Quell-Server' 'Source servers'), 'FQDN', 'TLS', (L 'Max. Größe' 'Max size'), (L 'Status' 'Status')) -Rows $scRows.ToArray()))
            if ($scEmpty) {
                $null = $Parts.Add((New-WdParagraph (L '📋 Hinweis: Auf einer neuen Exchange-Installation ist noch kein Send Connector konfiguriert. EXpress legt auf Wunsch einen Standard-Send-Connector an (siehe Installationsparameter). Bitte nach Abschluss der Konfiguration ergänzen.' '📋 Note: No send connector is configured on a new Exchange installation. EXpress can create a default send connector on request (see installation parameters). Please update this section after completing the configuration.')))
            }

            # 4.11 Federation / Hybrid / OAuth
            $null = $Parts.Add((New-WdHeading (L '4.11 Federation, Hybrid und OAuth' '4.11 Federation, Hybrid and OAuth') 2))
            $null = $Parts.Add((New-WdParagraph (L 'Federation und Hybrid-Konfiguration verbinden die on-premises Exchange-Organisation mit Exchange Online (Microsoft 365) bzw. anderen Exchange-Organisationen. Eine Hybrid-Konfiguration ist Voraussetzung für eine schrittweise Migration in die Cloud, für Cross-Premises-Postfachbewegungen (New-MoveRequest), für geteilte Kalenderfreigaben (Free/Busy), Nachrichtenverfolgung und für die gemeinsame Nutzung der gleichen SMTP-Domäne zwischen on-premises und Cloud. OAuth ermöglicht serverseitige Authentifizierung zwischen Exchange Server und anderen Workloads (EXO, SharePoint, Skype for Business).' 'Federation and hybrid configuration connect the on-premises Exchange organisation with Exchange Online (Microsoft 365) or other Exchange organisations. A hybrid configuration is a prerequisite for a staged cloud migration, for cross-premises mailbox moves (New-MoveRequest), for shared calendar/free-busy, message tracing, and for sharing a single SMTP namespace between on-premises and the cloud. OAuth enables server-to-server authentication between Exchange Server and other workloads (EXO, SharePoint, Skype for Business).')))
            if ($orgD.FederationTrust -and $orgD.FederationTrust.Count -gt 0) {
                $fedRows = $orgD.FederationTrust | ForEach-Object { @($_.Name, (SafeVal $_.ApplicationUri), (SafeVal $_.TokenIssuerUri)) }
                $null = $Parts.Add((New-WdTable -Headers @((L 'Federation Trust' 'Federation trust'), 'Application URI', 'Token Issuer') -Rows $fedRows))
            }
            if ($orgD.HybridConfig) {
                $hyb2 = $orgD.HybridConfig
                $hybRows2 = @(
                    @((L 'Hybrid-Features' 'Hybrid features'), (SafeVal ($hyb2.Features -join ', ')))
                    @((L 'On-Premises SMTP-Domänen' 'On-premises SMTP domains'), (SafeVal ($hyb2.OnPremisesSMTPDomains -join ', ')))
                    @((L 'Edge-Transport-Server' 'Edge Transport servers'), (SafeVal ($hyb2.EdgeTransportServers -join ', ')))
                    @((L 'Client Access Server' 'Client Access servers'), (SafeVal ($hyb2.ClientAccessServers -join ', ')))
                    @((L 'Empfangs-Connector' 'Receive connector'), (SafeVal ($hyb2.ReceivingTransportServers -join ', ')))
                    @((L 'Sende-Connector' 'Send connector'), (SafeVal ($hyb2.SendingTransportServers -join ', ')))
                    @((L 'Externe SMTP-Domänen' 'External SMTP domains'), (SafeVal ($hyb2.ExternalIPAddresses -join ', ')))
                    @((L 'TLS-Zertifikatsname' 'TLS certificate name'), (SafeVal $hyb2.TlsCertificateName))
                )
                $null = $Parts.Add((New-WdTable -Headers @((L 'Hybrid-Eigenschaft' 'Hybrid property'), (L 'Wert' 'Value')) -Rows $hybRows2))
                $null = $Parts.Add((New-WdParagraph (L 'Hinweis: Hybrid Configuration Wizard (HCW) prüft und aktualisiert diese Einstellungen automatisch. Änderungen sollten stets über den HCW oder Set-HybridConfiguration erfolgen, nicht über manuelle ADSIEdit- oder Registry-Eingriffe.' 'Note: Hybrid Configuration Wizard (HCW) validates and updates these settings automatically. Changes should always be made via HCW or Set-HybridConfiguration, never via manual ADSIEdit or registry edits.')))
            }
            if ($orgD.IntraOrgConnectors -and $orgD.IntraOrgConnectors.Count -gt 0) {
                $iocRows = $orgD.IntraOrgConnectors | ForEach-Object { @($_.Name, (SafeVal $_.TargetAddressDomains), (SafeVal $_.DiscoveryEndpoint), (Lc $_.Enabled (L 'Aktiv' 'Active') (L 'Inaktiv' 'Inactive'))) }
                $null = $Parts.Add((New-WdTable -Headers @('IntraOrg Connector', (L 'Zieldomänen' 'Target domains'), 'Discovery', (L 'Status' 'Status')) -Rows $iocRows))
            }
            if (-not $orgD.FederationTrust -and -not $orgD.HybridConfig -and -not ($orgD.IntraOrgConnectors | Where-Object { $_ })) {
                $null = $Parts.Add((New-WdParagraph (L '(Keine Federation/Hybrid-Konfiguration vorhanden — reine on-premises Umgebung)' '(No federation/hybrid configuration present — on-premises only environment)')))
                $null = $Parts.Add((New-WdParagraph (L '📋 Hinweis: Federation Trust, Hybrid Configuration und IntraOrg-Connectoren werden erst nach Ausführung des Hybrid Configuration Wizard (HCW) oder einer manuellen Hybrid-Konfiguration angelegt. Dieser Abschnitt ist nach einer Cloud-Migration oder Hybrid-Einrichtung zu aktualisieren.' '📋 Note: Federation trust, hybrid configuration and intra-org connectors are only created after running the Hybrid Configuration Wizard (HCW) or a manual hybrid setup. Update this section after a cloud migration or hybrid setup.')))
            }

            # 4.12 AuthConfig + Auth-Zertifikat
            $null = $Parts.Add((New-WdHeading (L '4.12 Auth-Zertifikat und OAuth-Konfiguration' '4.12 Auth Certificate and OAuth Configuration') 2))
            $null = $Parts.Add((New-WdParagraph (L 'Das Auth-Zertifikat ist das zentrale Sicherheitsobjekt für die server-interne OAuth-Kommunikation (OAuth 2.0). Es signiert die Token, die Exchange-Dienste untereinander und gegenüber Exchange Online austauschen. Die Lebensdauer beträgt standardmäßig 5 Jahre; läuft das Zertifikat ab, schlägt OAuth fehl (Hybrid-Szenarien, Exchange Online Federation, OWA/ECP-Rückfragen auf andere Server). MEAC (MonitorExchangeAuthCertificate.ps1) übernimmt die automatische Erneuerung 60 Tage vor Ablauf durch einen geplanten Task (siehe Kapitel 7).' 'The Auth Certificate is the central security artifact for server-internal OAuth communication (OAuth 2.0). It signs the tokens that Exchange services exchange among themselves and with Exchange Online. Default lifetime is 5 years; once it expires OAuth fails (hybrid scenarios, Exchange Online federation, OWA/ECP cross-server calls). MEAC (MonitorExchangeAuthCertificate.ps1) handles automatic renewal 60 days before expiry via a scheduled task (see chapter 7).')))
            if ($orgD.AuthConfig) {
                $ac = $orgD.AuthConfig
                $fmtTp = {
                    param($thumb)
                    if (-not $thumb) { return (L '(nicht gesetzt)' '(not set)') }
                    if ($Cust)       { return ('{0}...' -f $thumb.Substring(0, [Math]::Min(8, $thumb.Length))) }
                    [string]$thumb
                }
                $tp     = & $fmtTp $ac.CurrentCertificateThumbprint
                $tpNext = & $fmtTp $ac.NextCertificateThumbprint
                $tpPrev = & $fmtTp $ac.PreviousCertificateThumbprint
                # Auth cert validity: AuthConfig does not expose NotAfter directly — look up the cert
                # by thumbprint from the local server's Exchange cert store.
                $validUntil = (L '(unbekannt)' '(unknown)')
                $daysLeft   = $null
                if ($ac.CurrentCertificateThumbprint) {
                    try {
                        $authCert = Get-ExchangeCertificate -Thumbprint $ac.CurrentCertificateThumbprint -Server $env:COMPUTERNAME -ErrorAction Stop
                        if ($authCert -and $authCert.NotAfter) {
                            $validUntil = $authCert.NotAfter.ToString('yyyy-MM-dd')
                            $daysLeft = [int]([Math]::Floor(($authCert.NotAfter - (Get-Date)).TotalDays))
                        }
                    } catch {
                        try {
                            $certStore = Get-ChildItem -Path 'Cert:\LocalMachine\My' -ErrorAction Stop | Where-Object { $_.Thumbprint -eq $ac.CurrentCertificateThumbprint } | Select-Object -First 1
                            if ($certStore) {
                                $validUntil = $certStore.NotAfter.ToString('yyyy-MM-dd')
                                $daysLeft = [int]([Math]::Floor(($certStore.NotAfter - (Get-Date)).TotalDays))
                            }
                        } catch { Write-MyVerbose ('Auth certificate store lookup failed: {0}' -f $_) }
                    }
                }
                $validUntilCell = if ($null -ne $daysLeft) { ('{0} ({1} Tage verbleibend / {1} days remaining)' -f $validUntil, $daysLeft) } else { $validUntil }
                $authRows = [System.Collections.Generic.List[object[]]]::new()
                $authRows.Add(@((L 'Aktuelles Auth-Zertifikat (Fingerabdruck)' 'Current Auth cert thumbprint'), $tp))
                $authRows.Add(@((L 'Gültig bis' 'Valid until'), $validUntilCell))
                $authRows.Add(@((L 'Nächstes Auth-Zertifikat' 'Next Auth certificate'), $tpNext))
                $authRows.Add(@((L 'Vorheriges Auth-Zertifikat' 'Previous Auth certificate'), $tpPrev))
                $authRows.Add(@((L 'Realm' 'Realm'), (SafeVal $ac.Realm (L '(leer — Default)' '(empty — default)'))))
                $authRows.Add(@((L 'Service Name' 'Service name'), (SafeVal $ac.ServiceName (L '(nicht gesetzt)' '(not set)'))))
                $null = $Parts.Add((New-WdTable -Compact -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $authRows.ToArray()))
            } else {
                $null = $Parts.Add((New-WdParagraph (L '(AuthConfig nicht abrufbar)' '(AuthConfig not available)')))
            }
            $null = $Parts.Add((New-WdParagraph (L 'Wichtig: Eine manuelle Rotation des Auth-Zertifikats wird ausschließlich im Notfall empfohlen. Reguläre Rotation erfolgt über den MEAC-Task oder per Set-AuthConfig -PublishCertificate nach vorheriger Erzeugung eines "Next"-Zertifikats. Nach einer Rotation ist IISRESET auf allen Exchange-Servern erforderlich.' 'Important: Manual rotation of the Auth Certificate is only recommended as an emergency procedure. Regular rotation is handled by the MEAC task or via Set-AuthConfig -PublishCertificate after creating a "Next" certificate. After any rotation an IISRESET is required on all Exchange servers.')))

            # 4.13 Namensräume-Übersicht
            $null = $Parts.Add((New-WdHeading (L '4.13 Namensräume — konsolidierte Übersicht' '4.13 Namespaces — Consolidated Overview') 2))
            $null = $Parts.Add((New-WdParagraph (L 'Diese Tabelle aggregiert die Internal- und External-URLs aller Client-zugewandten Dienste über alle Exchange-Server hinweg. Identische URLs über alle Server sind Voraussetzung für Load Balancing ohne Session Affinity (ab Exchange 2016). Abweichende URLs innerhalb eines Dienstes deuten auf inkonsistente Namespace-Konfiguration hin und sollten korrigiert werden.' 'This table aggregates internal and external URLs for all client-facing services across all Exchange servers. Identical URLs across all servers are a prerequisite for load balancing without session affinity (since Exchange 2016). Diverging URLs within one service indicate inconsistent namespace configuration and should be corrected.')))
            $nsRows = [System.Collections.Generic.List[object[]]]::new()
            $vdirServices = @(
                @{ Name='OWA'        ; Prop='VDirOWA'  }
                @{ Name='ECP'        ; Prop='VDirECP'  }
                @{ Name='EWS'        ; Prop='VDirEWS'  }
                @{ Name='OAB'        ; Prop='VDirOAB'  }
                @{ Name='ActiveSync' ; Prop='VDirAS'   }
                @{ Name='MAPI'       ; Prop='VDirMAPI' }
                @{ Name='PowerShell' ; Prop='VDirPW'   }
            )
            foreach ($svc in $vdirServices) {
                $intUrls = @(); $extUrls = @()
                foreach ($srv2 in $ReportData.Servers) {
                    $vd = $srv2.($svc.Prop) | Select-Object -First 1
                    if ($vd) {
                        if ($vd.InternalUrl) { $intUrls += $vd.InternalUrl.AbsoluteUri }
                        if ($vd.ExternalUrl) { $extUrls += $vd.ExternalUrl.AbsoluteUri }
                    }
                }
                $intU = if ($intUrls) { ($intUrls | Select-Object -Unique) -join ', ' } else { (L '(nicht gesetzt)' '(not set)') }
                $extU = if ($extUrls) { ($extUrls | Select-Object -Unique) -join ', ' } else { (L '(nicht gesetzt)' '(not set)') }
                $consistency = if (($intUrls | Select-Object -Unique).Count -le 1 -and ($extUrls | Select-Object -Unique).Count -le 1) { (L 'konsistent' 'consistent') } else { (L 'ABWEICHUNG' 'DIVERGENT') }
                $nsRows.Add(@($svc.Name, (Mask-Ip $intU), (Mask-Ip $extU), $consistency))
            }
            $autodiscoverUrls = @($ReportData.Servers | ForEach-Object { if ($_.AutodiscoverSCP -and $_.AutodiscoverSCP.AutoDiscoverServiceInternalUri) { $_.AutodiscoverSCP.AutoDiscoverServiceInternalUri.ToString() } } | Where-Object { $_ })
            if ($autodiscoverUrls) {
                $adIn = ($autodiscoverUrls | Select-Object -Unique) -join ', '
                $adC  = if (($autodiscoverUrls | Select-Object -Unique).Count -le 1) { (L 'konsistent' 'consistent') } else { (L 'ABWEICHUNG' 'DIVERGENT') }
                $nsRows.Add(@('Autodiscover SCP', (Mask-Ip $adIn), '—', $adC))
            }
            $null = $Parts.Add((New-WdTable -Headers @((L 'Dienst' 'Service'), (L 'Interne URL' 'Internal URL'), (L 'Externe URL' 'External URL'), (L 'Konsistenz' 'Consistency')) -Rows $nsRows.ToArray()))

            # 4.14 RBAC — Rollengruppen (was 4.15 — 4.14 Database Copy Status merged into 4.9)
            if ($orgD.RoleGroups -and $orgD.RoleGroups.Count -gt 0) {
                $null = $Parts.Add((New-WdHeading (L '4.14 RBAC — Rollengruppen' '4.14 RBAC — Role Groups') 2))
                $null = $Parts.Add((New-WdParagraph (L 'Role-Based Access Control (RBAC) steuert, welche Exchange-Cmdlets und -Parameter ein Benutzer ausführen darf. Built-in-Rollengruppen wie "Organization Management", "Recipient Management" oder "View-Only Organization Management" werden von Exchange bereitgestellt. Benutzerdefinierte Rollengruppen erlauben feingranulare Delegation (z. B. Helpdesk ohne Zugriff auf Transport oder Hybrid). Diese Tabelle zeigt alle Rollengruppen mit ihren Mitgliedern — eine Dokumentation ist wichtig für Audits und Zugriffskontrollen.' 'Role-Based Access Control (RBAC) governs which Exchange cmdlets and parameters a user may run. Built-in role groups such as "Organization Management", "Recipient Management" or "View-Only Organization Management" are provided by Exchange. Custom role groups allow fine-grained delegation (e.g. helpdesk without access to transport or hybrid). This table lists all role groups with their members — documentation matters for audits and access reviews.')))
                $rgRows = [System.Collections.Generic.List[object[]]]::new()
                foreach ($rg in $orgD.RoleGroups) {
                    $memStr = if ($rg.Members -and $rg.Members.Count -gt 0) {
                        ($rg.Members | ForEach-Object { if ($Cust) { ('{0} ({1})' -f (Mask-Val $_.Name), $_.Type) } else { ('{0} ({1})' -f $_.Name, $_.Type) } }) -join '; '
                    } else { (L '(keine Mitglieder)' '(no members)') }
                    $rgRows.Add(@($rg.Name, (SafeVal $rg.Description), $memStr))
                }
                # ColWidths: Role group 2200, Description 4200 (long), Members 2860 — total 9260 twips (A4/Letter)
                $null = $Parts.Add((New-WdTable -Headers @((L 'Rollengruppe' 'Role group'), (L 'Beschreibung' 'Description'), (L 'Mitglieder' 'Members')) -Rows $rgRows.ToArray() -ColWidths @(2200, 4200, 2860)))
                $null = $Parts.Add((New-WdParagraph (L 'Hinweis: Eine detaillierte RBAC-Aufstellung mit verwalteten Rollen liefert der Befehl Get-RoleGroup | Format-List und Get-ManagementRoleAssignment. EXpress legt optional einen separaten RBAC-Report (.txt) im Reports-Verzeichnis ab.' 'Note: A detailed RBAC listing with managed roles is available via Get-RoleGroup | Format-List and Get-ManagementRoleAssignment. EXpress optionally writes a separate RBAC report (.txt) to the reports directory.')))
            }

            # 4.15 Audit-Konfiguration (was 4.16)
            $null = $Parts.Add((New-WdHeading (L '4.15 Audit-Konfiguration' '4.15 Audit Configuration') 2))
            $null = $Parts.Add((New-WdParagraph (L 'Das Admin-Auditprotokoll zeichnet alle Exchange-Verwaltungscmdlets auf, die von Administratoren ausgeführt werden (wer hat wann was geändert). Es ist Grundlage für Compliance-Anforderungen wie ISO 27001, BSI-Grundschutz und DSGVO-Rechenschaftspflicht. Das Protokoll wird in einem dedizierten verborgenen Postfach in der Exchange-Organisation gespeichert und kann per Search-AdminAuditLog abgefragt werden. Die Aufbewahrungsfrist (AdminAuditLogAgeLimit) bestimmt, wie lange Einträge erhalten bleiben (Standard: 90 Tage).' 'The admin audit log records all Exchange management cmdlets executed by administrators (who changed what and when). It is the basis for compliance requirements such as ISO 27001, BSI baseline protection and GDPR accountability. The log is stored in a dedicated hidden mailbox in the Exchange organisation and can be queried via Search-AdminAuditLog. The retention period (AdminAuditLogAgeLimit) determines how long entries are kept (default: 90 days).')))
            if ($orgD.AdminAuditLog) {
                $aal = $orgD.AdminAuditLog
                $aalRows = [System.Collections.Generic.List[object[]]]::new()
                $aalRows.Add(@((L 'Admin-Auditprotokoll aktiviert' 'Admin audit log enabled'),  (Format-RegBool $aal.AdminAuditLogEnabled)))
                $aalRows.Add(@((L 'Aufbewahrungsfrist' 'Log age limit'),                         (SafeVal $aal.AdminAuditLogAgeLimit (L '(Standard: 90 Tage)' '(default: 90 days)'))))
                $aalRows.Add(@((L 'Log-Postfach' 'Log mailbox'),                                 (SafeVal $aal.AdminAuditLogMailbox   (L '(Standard — automatisch)' '(default — automatic)'))))
                $aalCmdlets    = if ($aal.AdminAuditLogCmdlets)    { $aal.AdminAuditLogCmdlets -join ', '    } else { $null }
                $aalExclusions = if ($aal.AdminAuditLogExclusions) { $aal.AdminAuditLogExclusions -join ', ' } else { $null }
                $aalRows.Add(@((L 'Aufgezeichnete Cmdlets' 'Logged cmdlets'),  (SafeVal $aalCmdlets    (L '(alle — Standard)' '(all — default)'))))
                $aalRows.Add(@((L 'Ausgeschlossene Cmdlets' 'Excluded cmdlets'), (SafeVal $aalExclusions (L '(keine)' '(none)'))))
                $aalRows.Add(@((L 'Test-Cmdlet-Protokollierung' 'Test cmdlet logging'),          (Format-RegBool $aal.TestCmdletLoggingEnabled)))
                $aalRows.Add(@((L 'Log-Level' 'Log level'),                                      (SafeVal $aal.LogLevel)))
                $null = $Parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $aalRows.ToArray()))
            } else {
                $null = $Parts.Add((New-WdParagraph (L '(Admin-Auditprotokoll-Konfiguration nicht abrufbar)' '(Admin audit log configuration not available)')))
            }

            # 4.16 Exchange AD-Sicherheitsgruppen und Dienstkonten (was 4.17)
            $null = $Parts.Add((New-WdHeading (L '4.16 Exchange AD-Sicherheitsgruppen und Dienstkonten' '4.16 Exchange AD Security Groups and Service Accounts') 2))
            $null = $Parts.Add((New-WdParagraph (L 'Exchange erstellt bei der Installation und PrepareAD automatisch universelle Sicherheitsgruppen im Active Directory. Diese Gruppen steuern die Exchange-internen AD-Berechtigungen und sollten niemals manuell verändert werden. Dedizierte Dienstkonten für externe Integrationen (Backup, Monitoring, Archivierung, Relay) sind nach dem Principle of Least Privilege zu erstellen und mit minimalen Exchange-Rechten auszustatten.' 'Exchange creates universal security groups in Active Directory automatically during installation and PrepareAD. These groups control Exchange-internal AD permissions and must never be modified manually. Dedicated service accounts for external integrations (backup, monitoring, archiving, relay) must be created following the Principle of Least Privilege and assigned minimum Exchange rights.')))
            $adSecGrpRows = [System.Collections.Generic.List[object[]]]::new()
            $adSecGrpRows.Add(@('Exchange Trusted Subsystem',        (L 'Exchange-Servercomputer — vollständige AD-Kontrolle über Exchange-Objekte' 'Exchange server computers — full AD control over Exchange objects')))
            $adSecGrpRows.Add(@('Exchange Windows Permissions',       (L 'Lese-/Schreibzugriff auf AD-Benutzerobjekte für Transport und EWS' 'Read/write access to AD user objects for transport and EWS')))
            $adSecGrpRows.Add(@('Exchange Servers',                   (L 'Alle Exchange-Servercomputer in der Organisation' 'All Exchange server computers in the organisation')))
            $adSecGrpRows.Add(@('Managed Availability Servers',       (L 'Managed-Availability-Überwachung (Health Monitoring)' 'Managed Availability monitoring (Health Monitoring)')))
            $adSecGrpRows.Add(@('ExchangeLegacyInterop',              (L 'Interoperabilität mit Exchange 2010 und älter (kann leer sein)' 'Interoperability with Exchange 2010 and older (may be empty)')))
            $null = $Parts.Add((New-WdTable -Headers @((L 'Gruppe' 'Group'), (L 'Zweck' 'Purpose')) -Rows $adSecGrpRows.ToArray() -ColWidths @(2800, 6460)))
            $null = $Parts.Add((New-WdParagraph (L 'Dienstkonten — bitte nach der Installation ergänzen:' 'Service accounts — please fill in after installation:')))
            $svcAcctRows = [System.Collections.Generic.List[object[]]]::new()
            $svcAcctRows.Add(@((L 'Backup (z. B. Veeam, BackupExec)' 'Backup (e.g. Veeam, BackupExec)'), '— ', (L 'bitte eintragen' 'please fill in')))
            $svcAcctRows.Add(@((L 'Monitoring (z. B. SCOM, Zabbix)' 'Monitoring (e.g. SCOM, Zabbix)'), '—', (L 'bitte eintragen' 'please fill in')))
            $svcAcctRows.Add(@((L 'SMTP-Relay (Application Mail)' 'SMTP relay (application mail)'), '—', (L 'bitte eintragen' 'please fill in')))
            $svcAcctRows.Add(@((L 'Archivierung / Journaling' 'Archiving / journaling'), '—', (L 'bitte eintragen' 'please fill in')))
            $svcAcctRows.Add(@((L 'Sonstige Integrationen' 'Other integrations'), '—', (L 'bitte eintragen' 'please fill in')))
            $null = $Parts.Add((New-WdTable -Headers @((L 'Integration' 'Integration'), (L 'Konto (sAMAccountName)' 'Account (sAMAccountName)'), (L 'Berechtigungen / Rolle' 'Permissions / Role')) -Rows $svcAcctRows.ToArray() -ColWidths @(3000, 2500, 3760)))

            # Exchange Online / Microsoft 365 was formerly 4.17 inside "Organisation".
            # Moved to its own top-level section 15 (before Operative Runbooks) — belongs
            # to customer-ops context, not org-config telemetry. See below, just before
            # "Operative Runbooks".
    }

    function Invoke-DocSection-AntiSpam {
        param(
            [System.Collections.Generic.List[string]]$Parts,
            [object]$OrgData,
            [bool]$DE,
            [bool]$Cust
        )
        function L([string]$d, [string]$e) { if ($DE) { $d } else { $e } }
        function Lc([bool]$c, [string]$a, [string]$b) { if ($c) { $a } else { $b } }
        function SafeVal([object]$v, [string]$fallback = '') { if ($null -eq $v -or "$v" -eq '') { $fallback } else { "$v" } }
        function Format-RegBool($v) {
            if ($null -eq $v -or "$v" -eq '') { return (L '(nicht gesetzt)' '(not set)') }
            if ([bool]$v) { return (L 'aktiviert' 'enabled') }
            return (L 'deaktiviert' 'disabled')
        }

        $orgD = $OrgData

        $null = $Parts.Add((New-WdHeading (L '9. Transport-Agents und Anti-Spam (lokaler Server)' '9. Transport Agents and Anti-Spam (local server)') 1))
        $null = $Parts.Add((New-WdParagraph (L 'Exchange Server enthält integrierte Anti-Spam-Agents, die auf Mailbox-Servern standardmäßig nicht aktiviert sind. EXpress aktiviert die Anti-Spam-Agents und konfiguriert sie so, dass ausschließlich der Recipient Filter Agent aktiv bleibt — dieser prüft, ob Empfänger im Active Directory existieren, und lehnt E-Mails an nicht vorhandene Empfänger bereits auf SMTP-Ebene ab (Directory Harvest Attack Protection). Content Filter, Sender Filter und andere Agents werden deaktiviert, da diese Aufgaben in Unternehmensumgebungen typischerweise durch dedizierte Gateway-Lösungen (z. B. Hornetsecurity, Proofpoint, Mimeacst) oder Exchange Online Protection (EOP) übernommen werden.' 'Exchange Server includes built-in anti-spam agents that are not enabled by default on Mailbox servers. EXpress enables the anti-spam agents and configures them so that only the Recipient Filter Agent remains active — this checks whether recipients exist in Active Directory and rejects emails to non-existent recipients at the SMTP level (Directory Harvest Attack Protection). Content Filter, Sender Filter and other agents are disabled, as these tasks are typically handled by dedicated gateway solutions (e.g. Hornetsecurity, Proofpoint, Mimeacst) or Exchange Online Protection (EOP) in enterprise environments.')))
        $agentRows2 = [System.Collections.Generic.List[object[]]]::new()
        try {
            # Collect agents from all transport scopes (HubTransport is the default; on Mailbox
            # servers the FrontendTransport and MailboxSubmission/Delivery scopes each expose a
            # separate agent list). Deduplicate by Identity to keep the table compact.
            $seenAg = @{}
            $scopes = @('TransportService','FrontendTransport','MailboxSubmission','MailboxDelivery')
            $collected = @()
            foreach ($sc in $scopes) {
                try { $collected += @(Get-TransportAgent -TransportService $sc -ErrorAction SilentlyContinue) } catch { Write-MyVerbose ('Get-TransportAgent -TransportService {0} failed: {1}' -f $sc, $_) }
            }
            if (-not $collected -or $collected.Count -eq 0) {
                $collected = @(Get-TransportAgent -ErrorAction SilentlyContinue)
            }
            # Lookup used by section 9.1 to cross-reference org-wide *FilterConfig.Enabled
            # with the actual TransportAgent.Enabled state. Without this cross-reference the
            # doc shows "Enabled=True" for Content/Sender/Sender-ID even after the installer
            # has disabled the corresponding agents, because *FilterConfig.Enabled is just the
            # org-level feature switch, not the effective pipeline state.
            $script:__agentByKind = @{}
            foreach ($ag in $collected) {
                if (-not $ag) { continue }
                $agName = if ($ag.Name) { [string]$ag.Name } elseif ($ag.Identity) { [string]$ag.Identity } else { '(unbenannt)' }
                $kind = switch -Regex ($agName) {
                    'Content Filter'         { 'Content'; break }
                    'Sender Filter'          { 'Sender'; break }
                    'Recipient Filter'       { 'Recipient'; break }
                    'Sender ?Id|Sender Id'   { 'SenderId'; break }
                    'Connection Filter(ing)?'{ 'Connection'; break }
                    'Protocol Analysis'      { 'ProtocolAnalysis'; break }
                    default                  { $null }
                }
                if ($kind -and -not $script:__agentByKind.ContainsKey($kind)) {
                    $script:__agentByKind[$kind] = $ag
                }
                if ($seenAg.ContainsKey($agName)) { continue }
                $seenAg[$agName] = $true
                $agentState2 = if ($ag.Enabled) { (L 'Aktiv' 'Enabled') } else { (L 'Inaktiv' 'Disabled') }
                $agentRows2.Add(@($agName, $agentState2, $ag.Priority))
            }
        } catch { Write-MyVerbose ('Transport agent enumeration failed: {0}' -f $_) }

        # Helper — renders the effective pipeline state for a filter's underlying TransportAgent.
        # Distinguishes three cases so a reader can tell the difference between "org switch on, agent off"
        # (EXpress default: org config says Enabled, agent is disabled → filter inert) and the other two.
        function Get-EffectiveAgentState {
            param([string]$Kind)
            $ag = $script:__agentByKind[$Kind]
            if (-not $ag) { return (L 'Nicht installiert' 'Not installed') }
            if ($ag.Enabled) { return (L 'Aktiv — Agent läuft im Transport-Pipeline' 'Enabled — agent runs in transport pipeline') }
            return (L 'Inaktiv — Transport-Agent ist deaktiviert, Filter greift nicht (Org-Schalter ist nur ein Feature-Flag)' 'Inactive — transport agent is disabled, filter does not fire (org switch is only a feature flag)')
        }
        if ($agentRows2.Count -eq 0) { $agentRows2.Add(@((L '(keine konfiguriert)' '(none configured)'), '', '')) }
        $null = $Parts.Add((New-WdTable -Headers @('Agent', (L 'Status' 'Status'), (L 'Priorität' 'Priority')) -Rows $agentRows2.ToArray()))

        # 9.1 Anti-Spam-Filter-Konfiguration (org-weite Filtereinstellungen)
        $hasAnyFilter = $orgD.ContentFilterConfig -or $orgD.SenderFilterConfig -or $orgD.RecipientFilterConfig -or $orgD.SenderIdConfig
        if ($hasAnyFilter) {
            $null = $Parts.Add((New-WdHeading (L '9.1 Anti-Spam-Filter-Konfiguration' '9.1 Anti-Spam Filter Configuration') 2))
            $null = $Parts.Add((New-WdParagraph (L 'Die folgenden Tabellen zeigen die organisationsweite Konfiguration der installierten Anti-Spam-Filter-Agents. In reinen on-premises-Umgebungen ohne vorgelagerten Cloud-Filter (EOP/Hornetsecurity/Proofpoint) sind diese Einstellungen aktiv wirksam. In Hybrid-Umgebungen oder mit vorgelagerten Gateways werden Content- und Sender-Filter typischerweise deaktiviert (Recipient Filter bleibt für Directory Harvest Attack Protection aktiv).' 'The following tables show the organisation-wide configuration of the installed anti-spam filter agents. In pure on-premises environments without an upstream cloud filter (EOP/Hornetsecurity/Proofpoint), these settings are actively effective. In hybrid environments or with upstream gateways, Content and Sender Filters are typically disabled (Recipient Filter remains active for Directory Harvest Attack Protection).')))
            $null = $Parts.Add((New-WdParagraph (L 'Hinweis zur Unterscheidung: "Effektiver Status (Transport-Agent)" zeigt, ob der Agent tatsächlich in der Transport-Pipeline läuft (Get-TransportAgent). "Org-Konfig Enabled" ist nur der organisationsweite Feature-Schalter (Get-*FilterConfig) und sagt nichts darüber aus, ob der Filter wirklich greift. EXpress deaktiviert standardmäßig alle Transport-Agents außer dem Recipient Filter — "Org-Konfig Enabled = True" bei deaktiviertem Transport-Agent bedeutet daher: Filter greift nicht.' 'Note on interpretation: "Effective status (transport agent)" shows whether the agent actually runs in the transport pipeline (Get-TransportAgent). "Org config Enabled" is only the organisation-wide feature switch (Get-*FilterConfig) and says nothing about whether the filter actually fires. EXpress disables all transport agents by default except Recipient Filter — "Org config Enabled = True" with a disabled transport agent therefore means: filter does not fire.')))
            if ($orgD.ContentFilterConfig) {
                $cf = $orgD.ContentFilterConfig
                $cfRows = [System.Collections.Generic.List[object[]]]::new()
                $cfRows.Add(@((L 'Effektiver Status (Transport-Agent)' 'Effective status (transport agent)'), (Get-EffectiveAgentState 'Content')))
                $cfRows.Add(@((L 'Org-Konfig Enabled (Feature-Flag)' 'Org config Enabled (feature flag)'),   (Format-RegBool $cf.Enabled)))
                $cfRows.Add(@((L 'Aktion (SCL ≥ 6)' 'Action (SCL ≥ 6)'),  (SafeVal $cf.SCLRejectEnabled (L '(nicht gesetzt)' '(not set)'))))
                $cfRows.Add(@((L 'SCL-Ablehneschwellenwert' 'SCL reject threshold'), (SafeVal $cf.SCLRejectThreshold)))
                $cfRows.Add(@((L 'SCL-Löschschwellenwert' 'SCL delete threshold'),  (SafeVal $cf.SCLDeleteThreshold)))
                $cfRows.Add(@((L 'SCL-Quarantäneschwellenwert' 'SCL quarantine threshold'), (SafeVal $cf.SCLQuarantineThreshold)))
                $cfRows.Add(@((L 'Quarantäne-Postfach' 'Quarantine mailbox'),       (SafeVal $cf.QuarantineMailbox (L '(nicht gesetzt)' '(not set)'))))
                $null = $Parts.Add((New-WdHeading (L 'Content Filter' 'Content Filter') 3))
                $null = $Parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $cfRows.ToArray()))
            }
            if ($orgD.SenderFilterConfig) {
                $sf = $orgD.SenderFilterConfig
                $sfRows = [System.Collections.Generic.List[object[]]]::new()
                $sfRows.Add(@((L 'Effektiver Status (Transport-Agent)' 'Effective status (transport agent)'), (Get-EffectiveAgentState 'Sender')))
                $sfRows.Add(@((L 'Org-Konfig Enabled (Feature-Flag)' 'Org config Enabled (feature flag)'),   (Format-RegBool $sf.Enabled)))
                $sfRows.Add(@((L 'Leere Absender blockieren' 'Block blank senders'), (Format-RegBool $sf.BlankSenderBlockingEnabled)))
                $sfBlockedSenders = if ($sf.BlockedSenders) { $sf.BlockedSenders -join '; ' } else { $null }
                $sfBlockedDomains = if ($sf.BlockedDomains) { $sf.BlockedDomains -join '; ' } else { $null }
                $sfRows.Add(@((L 'Blockliste (Absender)' 'Block list (senders)'), (SafeVal $sfBlockedSenders (L '(leer)' '(empty)'))))
                $sfRows.Add(@((L 'Blockliste (Domänen)' 'Block list (domains)'),  (SafeVal $sfBlockedDomains (L '(leer)' '(empty)'))))
                $null = $Parts.Add((New-WdHeading (L 'Sender Filter' 'Sender Filter') 3))
                $null = $Parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $sfRows.ToArray()))
            }
            if ($orgD.RecipientFilterConfig) {
                $rf = $orgD.RecipientFilterConfig
                $rfRows = [System.Collections.Generic.List[object[]]]::new()
                $rfRows.Add(@((L 'Effektiver Status (Transport-Agent)' 'Effective status (transport agent)'), (Get-EffectiveAgentState 'Recipient')))
                $rfRows.Add(@((L 'Org-Konfig Enabled (Feature-Flag)' 'Org config Enabled (feature flag)'),   (Format-RegBool $rf.Enabled)))
                $rfBlockedRecipients = if ($rf.BlockedRecipients) { $rf.BlockedRecipients -join '; ' } else { $null }
                $rfRows.Add(@((L 'Blockliste (Empfänger)' 'Block list (recipients)'), (SafeVal $rfBlockedRecipients (L '(leer)' '(empty)'))))
                $rfRows.Add(@((L 'Empfänger-Lookup aktiviert' 'Recipient lookup enabled'), (Format-RegBool $rf.RecipientValidationEnabled)))
                $null = $Parts.Add((New-WdHeading (L 'Recipient Filter' 'Recipient Filter') 3))
                $null = $Parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $rfRows.ToArray()))
            }
            if ($orgD.SenderIdConfig) {
                $si = $orgD.SenderIdConfig
                $siRows = [System.Collections.Generic.List[object[]]]::new()
                $siRows.Add(@((L 'Effektiver Status (Transport-Agent)' 'Effective status (transport agent)'), (Get-EffectiveAgentState 'SenderId')))
                $siRows.Add(@((L 'Org-Konfig Enabled (Feature-Flag)' 'Org config Enabled (feature flag)'),   (Format-RegBool $si.Enabled)))
                $siRows.Add(@((L 'Aktion (Spoofed)' 'Action (spoofed)'),             (SafeVal $si.SpoofedDomainAction)))
                $siRows.Add(@((L 'Aktion (Temporary Error)' 'Action (temp error)'),  (SafeVal $si.TempErrorAction)))
                $null = $Parts.Add((New-WdHeading 'Sender ID' 3))
                $null = $Parts.Add((New-WdTable -Headers @((L 'Eigenschaft' 'Property'), (L 'Wert' 'Value')) -Rows $siRows.ToArray()))
            }
        }
    }
