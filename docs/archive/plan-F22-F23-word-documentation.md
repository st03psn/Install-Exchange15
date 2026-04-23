# Word-Dokumentation für EXpress (F22 + F23)

## Kontext

EXpress erzeugt heute HTML-Reports (`New-PreflightReport`, `New-InstallationReport`).
Benötigt werden zwei Word-Artefakte für Exchange-Einführungsprojekte:

1. **F23 — Konzept-/Freigabedokument** (*vor* der Installation) — Planung, Stand
   der Technik, Lizenzbedingungen, Fragenkatalog zur Klärung/Freigabe mit
   Kunde/Management. **Nicht Teil des Tools**, sondern eigenständiges Beiwerk
   im Repo, manuell genutzt.
2. **F22 — Installations-/Konfigurationsdokumentation** (*nach* der
   Installation) — Word-Pendant zum HTML-`InstallationReport`. **Wird vom Tool
   erzeugt.**

Struktur-Referenz: Kundendoku `Exchange Installation&Konfiguration.docx`
(Headings extrahiert, kundenspezifische Inhalte verworfen, Lücken bewusst
ergänzt um typische Projekt-Themen).

Engine: **Pure-PowerShell OpenXML-ZIP** (`System.IO.Compression`) — keine
externen Abhängigkeiten, läuft auf Server Core, kompatibel mit PS2Exe-Build.

---

## Entschieden — Design-Optionen

| Punkt | Entscheidung |
|---|---|
| **Sprache** | Zwei separate Dokument-Varianten via `-Language DE` / `-Language EN` (Default `DE`) |
| **Branding** | Neutrale Farben, `[LOGO]`-Platzhalter im Header |
| **Seitenlayout** | A4 hoch, Kopfzeile (Logo + Titel), Fußzeile (Seitenzahl + Klassifizierung `INTERN`) |
| **Referenz** | `Exchange Installation&Konfiguration.docx` (EVK Bergisch Gladbach) — Struktur extrahiert, Kundenspezifika verworfen |
| **F23 Zielgruppe** | IT + Kunde (Freigabedokument) |
| **F23 Freigabe** | 4-spaltige Freigabetabelle (Ersteller / Prüfer / Freigeber / Kunde) |
| **F23 Compliance** | CIS + BSI + ISO 27001 + DSGVO |
| **F23 Fragenkatalog** | Tabelle mit Content Controls (SDT), auch handschriftlich ausfüllbar |
| **F23 Lizenzen** | Nur Verweis auf MS-Preisleitfaden |
| **F22 CMDLets** | Zwei Unterabschnitte: (a) vom Skript aufgerufen (b) alle Exchange-Cmdlets aus Transcript (max 200) |
| **F22 Härtung** | Tabelle (Maßnahme / Status-Icon / CIS-Ref) + Fließtext je Gruppe |
| **F22 HealthChecker** | Nur Pfad-Verweis auf HTML-Report |
| **F22 Varianten** | `-CustomerDocument` maskiert Passwörter + interne IPs |

---

## Teil A (F23) — Mitgeliefertes Konzeptdokument

### Ablage
Neues Verzeichnis `templates/` im Repo-Root mit:

- `templates/Exchange-concept-template-DE.docx`
- `templates/Exchange-concept-template-EN.docx`
- `templates/README.md` — wofür, wie zu verwenden, Änderungshinweise

### Einmalige Erzeugung
Hilfsskript `tools/Build-KonzeptTemplate.ps1` (nicht im Installations-Flow) baut
beide `.docx` aus OpenXML-ZIP. Läuft beim Maintainer, Output wird committed.

### Gliederung F23 (16 Kapitel, DE + EN)

1. **Projektrahmen** — Auftrag, Scope, Beteiligte, Zeitplan
2. **Lizenzbedingungen & Stand der Technik**
   - Nur **Exchange Server Subscription Edition (SE)** — 2016/2019 out-of-support
     (14.10.2025), explizit als nicht mehr einsetzbar gekennzeichnet
   - WS 2022 / WS 2025 Kompatibilität
   - **Subscription-Modell** (Server-Lizenz + CAL-Subscription / M365-Inclusion),
     keine Perpetual-Lizenzen
   - CU/SU-Kadenz (SE RTM, CU1, …)
   - Migrationspfad-Hinweis (2016/2019 → Legacy-Koexistenz → SE)
3. **IST-Aufnahme**
   - 3.1 Active Directory (Forest/Domain Functional Level, FSMO, Schema, Sites)
   - 3.2 Exchange-Bestandsumgebung (falls vorhanden — Legacy 2016/2019)
   - 3.3 Berechtigungen & Delegation
   - 3.4 Netzwerk & DNS
4. **Sizing & Kapazitätsplanung**
   - 4.1 Mailbox-Anzahl, Größenprofile, Wachstum
   - 4.2 CPU/RAM/Disk-Dimensionierung (Calculator-Referenz)
   - 4.3 Storage-Layout (DB/Log-Trennung, 64 KB NTFS, ReFS-Option)
5. **SOLL-Architektur**
   - 5.1 Namensräume (intern/extern, Split-DNS, SAN-Zertifikat)
   - 5.2 Server-Topologie (Mailbox, Edge, DAG-Mitglieder)
   - 5.3 **DAG-Design** — FSW + Alternate FSW, DAC-Mode, Replication/MAPI-Netze,
         Activation Preference, Lag Copies
   - 5.4 **Netzwerk & Load Balancer** — Persistence, Health Probes, SNAT, HA-LB
   - 5.5 **Firewall-Matrix** — intra-Exchange, Exchange↔AD, Exchange↔Internet,
         Edge↔Mailbox
   - 5.6 Datenbanken & Disk-Layout
   - 5.7 Konnektoren (Receive / Send / Relay)
   - 5.8 **Zertifikatskonzept** — SAN/UCC, **Auth-Cert separat**, SMTP vs. IIS,
         Rotationsstrategie, Internal CA vs. Public CA
6. **Sicherheits- & Härtungskonzept** — TLS 1.2/1.3, SCHANNEL, AMSI, Extended
   Protection, SDS, LSA, SMBv1, WDigest, HTTP/2, Download Domains,
   Defender-Exclusions — mit CIS + BSI + ISO 27001 + DSGVO-Referenzen
7. **Message Hygiene** — Anti-Spam Agents vs. Edge Transport vs. 3rd Party
   (Hornetsecurity / Proofpoint / Mimeacst), Entscheidungsmatrix
8. **Backup, Recovery & Disaster Recovery** *(zusammengefasst)*
   - 8.1 VSS-Integration (Writer-Liste, Snapshot-Policy) + Backup-Strategie
   - 8.2 Circular Logging ja/nein, Truncation, RDB-Strategie, Restore-Test-Kadenz
   - 8.3 DR-Szenarien: FSW-Loss / Split-Brain, Server-Loss (`setup.exe /m:RecoverServer`),
         Namespace-Failover
9. **Monitoring-Konzept** — Managed Availability, PRTG/Checkmk/SCOM-Optionen,
   Event-IDs-Katalog, Perfmon-Baseline
10. **Migration / Koexistenz** *(konditional)*
    - 10.1 Legacy 2016/2019 → SE
    - 10.2 Cutover vs. schrittweise Koexistenz
    - 10.3 Public-Folder-Migration (Legacy → Modern)
    - 10.4 Namespace-Migration
11. **Hybrid / M365-Integration** — HCW, OAuth vs. Classic, Mailflow zentral/direkt,
    Free/Busy
12. **Public Folders & moderne Alternativen** — Platzhalter „Einsatz: ☐ Ja / ☐ Nein"
    - Bei „Ja": PF-Design (Modern Mailbox-basiert, Hierarchie)
    - Bei „Nein": Standard-Hinweistext **„Public Folders werden nicht eingesetzt.
      Moderne Ablösung über Shared Mailboxes (gemeinsamer Posteingang, Kalender)
      und Microsoft Teams (Dokumentenablage, Zusammenarbeit). Begründung:
      geringerer Administrationsaufwand, Cloud-native, bessere
      Mobile-/Web-Experience, keine DAG-Replikations-Abhängigkeit."**
13. **Compliance / eDiscovery / Journaling** — Litigation Hold, In-Place Archive,
    Retention Policies, DLP, Journal Rules
14. **Mobile & ActiveSync** — Policies, Intune-Integration, Quarantäne
15. **Fragenkatalog** — Tabelle mit Content Controls (SDT), handschriftlich
    ausfüllbar (alle offenen Parameter aus Kap. 3–14)
16. **Freigabeseite** — 4-spaltige Tabelle (Ersteller / Prüfer / Freigeber / Kunde)

### Verweise in bestehender Doku
- `README.md` um Abschnitt *„Planning Template"* ergänzen
- `docs/index.html` kurz verlinken (optional)

---

## Teil B (F22) — Generierte Installations-/Umgebungsdokumentation

### Scope-Erweiterung (v5.83+)

Das Dokument beschreibt **die gesamte Exchange-Organisation**, nicht nur den
lokal installierten Server. Drei Einsatzszenarien sind abgedeckt:

1. **Neue Umgebung** — nur ein Server (= lokaler), Org frisch angelegt
2. **Server-Ergänzung** — neuer Server zu bestehender Org hinzugefügt; alle
   vorhandenen Server + der neue werden dokumentiert
3. **Ad-hoc-Inventar** — `-StandaloneDocument` auf beliebigem Exchange-Server
   einer bestehenden Umgebung, ohne vorangegangenes EXpress-Setup

Dafür zwei Ebenen:

- **Org-weite Konfiguration** (einmal): Org-Config, Accepted/Remote Domains,
  Address Policies, Transport Rules, Journal/DLP/Retention, DAGs, Send
  Connectors, Federation/Hybrid/OAuth, AuthConfig.
- **Pro-Server-Konfiguration** (Schleife über `Get-ExchangeServer`): Identität,
  optional Hardware/Pagefile/Volumes via WinRM/CIM, Datenbanken, VDirs, Receive
  Connectors, Zertifikate, Transport Agents. Lokaler Server mit Marker
  „← Neu installiert durch diesen Lauf" wenn `$env:ComputerName` in Phase-6-Aufruf.

### Neue Funktion
`New-InstallationDocument` (Position: unmittelbar nach `New-InstallationReport`,
ca. Zeile 4500+).

**Wiederverwendung:** Datenabfragen aus `New-InstallationReport`
(`EXpress.ps1:4379`) extrahieren in interne Helper:

- `Get-OrganizationReportData` → Hashtable mit org-weiten Settings
- `Get-ServerReportData -Server <Name>` → Hashtable pro Server
  (Exchange-Cmdlets immer; CIM-Daten optional via `Get-RemoteServerData`)
- `Get-InstallationReportData` → aggregiert Org + ForEach(Server)

Beide Report-Funktionen (HTML + Word) konsumieren dasselbe Datenobjekt — keine
doppelten AD/Exchange-Queries.

### Remote-Query-Standard (neu)

**Transport:** CIM über WSMan (WinRM TCP 5985/5986, Kerberos), **nicht**
WMI/DCOM. Begründung: WinRM ist für Exchange EMS ohnehin Pflicht, firewall-
freundlich (ein Port), keine dynamischen RPC-Ports.

**Helper:** `Get-RemoteServerData -ComputerName <x>` in EXpress.ps1 —
einheitliches Rückgabeschema (`@{Reachable; OS; CPU; Memory; PageFile; Volumes;
NICs; Error}`), 30 s Timeout, try/finally mit `Remove-CimSession`.

**Abfrageklassen:** `Win32_OperatingSystem`, `Win32_Processor`,
`Win32_ComputerSystem`, `Win32_PageFileSetting`, `Win32_Volume` (Filter
`DriveType=3`), `Win32_NetworkAdapterConfiguration` (Filter `IPEnabled=TRUE`).

**Pre-Requisites (zwei Wege):**

1. Script `tools/Enable-EXpressRemoteQuery.ps1` — `Enable-PSRemoting`, Firewall,
   optional HTTPS-Listener. Aufruf lokal auf jedem Ziel-Server oder via
   `Invoke-Command` zentral (wenn WinRM initial läuft).
2. GPO `docs/remote-query-setup.md` — „Remoteserver­verwaltung über WinRM
   zulassen" + Firewall-Regel + WinRM-Dienst = Automatisch.

**Härtung:**
- kein `TrustedHosts *` — nur Kerberos in Domäne
- optional HTTPS-Listener (5986) mit Exchange-Auth-Cert
- WinRM-ACL auf dedizierte AD-Gruppe `EXpress-DocReader` (read-only)
- `New-InstallationDocument` nutzt ausschließlich `Get-*`-Cmdlets

**Degradation:** Fehlschlag pro Server → Hinweistext im Dokument
(„System-Details nicht abrufbar — WinRM nicht erreichbar. Abhilfe: siehe
`tools/Enable-EXpressRemoteQuery.ps1` oder GPO-Anleitung"), Lauf fährt fort.

**Interaktive Nachfrage (Copilot-Modus):** Schlägt `Get-RemoteServerData` für
einen Server fehl, zeigt EXpress einen Dialog:

```
[!] Remote-Abfrage für EX02.contoso.local fehlgeschlagen
    Fehler: WinRM cannot complete the operation

    Abhilfe: Script tools/Enable-EXpressRemoteQuery.ps1 auf dem Zielserver
             ausführen, oder GPO "Exchange — EXpress Remote Query" anwenden
             (Details: docs/remote-query-setup.md).

    [R] Erneut versuchen   [S] Überspringen   (Auto-Skip in 10:00)
```

- Countdown 600 s auf `Write-Progress -Id 2` (gleiches Schema wie WU/Reboot-Prompts)
- `[R]` → wiederholt den Aufruf einmal; bei erneutem Fehlschlag gleicher Dialog
- `[S]` oder Timeout → markiert den Server als `Reachable=$false`, Hinweistext
  im Dokument, Lauf fährt fort
- **Autopilot-Modus:** Prompt wird übersprungen, Verhalten = `[S]` (lautloser Skip)
- **`-StandaloneDocument` ohne interaktive Konsole** (z. B. via Scheduled Task):
  ebenfalls automatischer Skip

**Helper:** `Invoke-RemoteQueryWithPrompt -ComputerName <x>` wrappt
`Get-RemoteServerData` mit dieser Logik. Read-MenuInput mit Timeout-Parameter
als Basis (ggf. neu einführen, falls nicht vorhanden).

### OpenXML-Helper (privater Funktionsblock nach `New-InstallationReport`)

- `New-WordDocument` — Hauptfunktion, schreibt `.docx` via
  `[System.IO.Compression.ZipArchive]` mit Minimal-Parts:
  - `[Content_Types].xml`
  - `_rels/.rels`
  - `word/_rels/document.xml.rels`
  - `word/document.xml`
  - `word/styles.xml`
  - `word/numbering.xml`
  - `word/header1.xml` + `word/footer1.xml`
- Fragment-Generatoren: `Add-WordHeading`, `Add-WordParagraph`, `Add-WordTable`,
  `Add-WordCodeBlock`, `Add-WordBullet`, `Add-WordContentControl`
- UTF-8 ohne BOM; `xml:space="preserve"` bei Leerzeichen
- Styles: `Heading1/2/3/4`, `Normal`, `Code` (Consolas), `TableGrid`
- Deutsche Überschriften + Umlaute

### Gliederung F22 (16 Kapitel, DE + EN, `-CustomerDocument` maskiert)

| # | Heading | Quelle |
|---|---|---|
| 1 | Titelblatt | Lokaler Server, Org, Datum, EXpress-Version, Szenario (neu/ergänzt/ad-hoc) |
| 2 | Installationsparameter | `$State`-Auszug (gefiltert, bei `-CustomerDocument` maskiert); bei Ad-hoc entfällt Kapitel |
| 3 | IST-Aufnahme AD | `Get-ForestFunctionalLevel`, FSMO, Schema, Sites |
| 4 | **Organisation — übergreifende Konfiguration** | |
| 4.1 | Org-Übersicht | `Get-OrganizationConfig` (Name, Version, MAPI/HTTP, Modern Auth, CEIP, OAuth2) |
| 4.2 | Accepted Domains | `Get-AcceptedDomain` |
| 4.3 | Remote Domains | `Get-RemoteDomain` |
| 4.4 | E-Mail-Adressrichtlinien | `Get-EmailAddressPolicy` |
| 4.5 | Transport Rules | `Get-TransportRule` (Name, State, Priority, Comment) |
| 4.6 | Transport-Konfiguration (Org) | `Get-TransportConfig` (Message Size Limits, Safety Net) |
| 4.7 | Journal/DLP/Retention | `Get-JournalRule`, `Get-DlpPolicy`, `Get-RetentionPolicy` |
| 4.8 | Mobile/OWA-Policies | `Get-MobileDeviceMailboxPolicy`, `Get-OwaMailboxPolicy` |
| 4.9 | DAGs (alle) | pro DAG: Mitglieder, FSW, Alternate FSW, DAC, Netze, Copy-Layout |
| 4.10 | Send Connectors (Org-Scope) | Adressräume, Smarthosts, Source-Server-Liste |
| 4.11 | Federation / Hybrid / OAuth | `Get-FederationTrust`, `Get-HybridConfiguration`, OAuth-Konfiguration |
| 4.12 | AuthConfig | `Get-AuthConfig` (Auth-Zertifikat — separat von Server-Zertifikaten) |
| 5 | **Server in der Organisation** | Schleife über `Get-ExchangeServer` |
| 5.x.1 | Identität | Name, FQDN, AD-Site, Edition, Version, Rolle, Install-Datum; Marker „← Neu installiert" bei lokalem Server in Phase-6-Aufruf |
| 5.x.2 | Systemdetails | `Get-RemoteServerData` via CIM/WSMan (OS, CPU, RAM, Pagefile, Volumes, NICs) oder Hinweistext bei Fehlschlag |
| 5.x.3 | Datenbanken | `Get-MailboxDatabase -Server`, `-Status` |
| 5.x.4 | Virtuelle Verzeichnisse + SCP | `Get-*VirtualDirectory -Server`, `Get-ClientAccessService` |
| 5.x.5 | Receive Connectors | `Get-ReceiveConnector -Server` |
| 5.x.6 | Zertifikate | `Get-ExchangeCertificate -Server` |
| 5.x.7 | Transport Agents | `Get-TransportAgent -Server` (nur Transport-Rolle) |
| 6 | Netzwerk & DNS (lokal) | NIC-Bindings, DNS-Server, Autodiscover-DNS-Records |
| 7 | Installation Exchange (lokal) | Setup-Version, Pfade, Phasen 0–6 + Timings; bei Ad-hoc entfällt |
| 8 | Optimierungen und Härtungen (lokal) | Tabelle (Maßnahme / Status-Icon / CIS-Ref) + Fließtext je Gruppe |
| 9 | Anti-Spam / Agents (lokal) | `Get-TransportAgent`, Content/Sender/Recipient Filter Status |
| 10 | Backup- & DR-Readiness (lokal) | VSS-Writer (`vssadmin list writers`), Defender-Exclusions, DR-Hinweise |
| 11 | HealthChecker | Pfad-Verweis auf HTML-Report (`$State['HCReportPath']`) |
| 12 | Monitoring-Readiness | Managed Availability Status, Event-Log-Retention, Perfmon-Baseline-Capture |
| 13 | Public Folders | Auto-Detektion via `Get-Mailbox -PublicFolder` — Liste + Statistik ODER Hinweistext „nicht im Einsatz, moderne Ablösung über Shared Mailboxes / Teams" |
| 14 | Ausgeführte CMDLets | 14.1 Vom Skript aufgerufen · 14.2 Exchange-Cmdlets aus Transcript (max 200); bei Ad-hoc entfällt |
| 15 | Operative Runbooks | DAG-Wartungsmodus, CU/SU-Update, Cert-Tausch, DB-Move, DB-Reseed, Failover-Test |
| 16 | Offene Punkte | Platzhalter-Tabelle mit Content Controls |

### Aufruf
Direkt nach `New-InstallationReport` in Phase 6 (`EXpress.ps1:8467`):

```powershell
if (-not $State['NoWordDoc']) {
    try { New-InstallationDocument } catch { Write-MyWarning "Word doc: $_" }
}
```

### Neue Parameter (`param()`-Block, `EXpress.ps1:949`)
- `[switch]$NoWordDoc`
- `[switch]$StandaloneDocument` — lädt bestehendes `$State` + Exchange-Session,
  ruft nur `New-InstallationDocument` (analog `-StandaloneOptimize`). Ohne
  State-Datei arbeitet die Funktion als reines Ad-hoc-Inventar.
- `[switch]$CustomerDocument` — maskiert Passwörter + interne IPs
- `[ValidateSet('DE','EN')][string]$Language = 'DE'`
- `[ValidateSet('All','Org','Local')][string]$DocumentScope = 'All'` — bei
  großen Farms einschränkbar
- `[string[]]$IncludeServers` — gezielte Filterung für große Umgebungen

Ebenso in State-Persistierung und Config-File-Parser.

### Menü
`Show-InstallationMenu` neuer Eintrag **„Generate Installation Document (Word)"**
→ ruft `-StandaloneDocument`-Pfad mit Language-Prompt.

### Output-Pfad
`$State['ReportsPath']\{ComputerName}_InstallationDocument_{Language}_{yyyyMMddHHmmss}.docx`

---

## Implementierungsreihenfolge

1. **OpenXML-Engine** in `tools/Build-KonzeptTemplate.ps1` entwickeln + validieren
   (`New-WordDocument`, `Add-WordHeading`, `Add-WordParagraph`, `Add-WordTable`,
   `Add-WordBullet`, `Add-WordContentControl`, `Add-WordCodeBlock`)
2. **F23 DE + EN** als statische Templates erzeugen + committen
   (`templates/Exchange-concept-template-DE.docx` + `-EN.docx` + `templates/README.md`)
3. **OpenXML-Engine** nach `EXpress.ps1` portieren
4. **`Get-InstallationReportData`** aus `New-InstallationReport` herausfaktorisieren
5. **`New-InstallationDocument`** implementieren + Params
   (`-NoWordDoc`, `-StandaloneDocument`, `-CustomerDocument`, `-Language`)
6. **Phase-6-Aufruf + Menüeintrag**
7. **Doku-Updates** (RELEASE-NOTES, CLAUDE.md, README), Tests, Commit

---

## Kritische Dateien / Funktionen

| Datei | Abschnitt | Aktion |
|---|---|---|
| `tools/Build-KonzeptTemplate.ps1` | (neu) | OpenXML-Engine + F23-Generator für DE/EN |
| `templates/Exchange-concept-template-DE.docx` | (neu) | Statisches Konzeptdokument DE |
| `templates/Exchange-concept-template-EN.docx` | (neu) | Statisches Konzeptdokument EN |
| `templates/README.md` | (neu) | Nutzungshinweis |
| `EXpress.ps1:949` | `param()` | `-NoWordDoc`, `-StandaloneDocument`, `-CustomerDocument`, `-Language` |
| `EXpress.ps1:4379` | `New-InstallationReport` | Datensammlung in `Get-InstallationReportData` extrahieren |
| `EXpress.ps1:~4500` | (neu) | OpenXML-Helper + `New-InstallationDocument` |
| `EXpress.ps1:8467` | Phase 6 | Aufruf hinter `New-InstallationReport` |
| `EXpress.ps1` `Show-InstallationMenu` | Menü | Neuer Eintrag |
| `RELEASE-NOTES.md` | neue Version | Feature-Eintrag F22 + F23 |
| `README.md` | What's New + Abschnitt „Planning Template" | |
| `CLAUDE.md` | Function Overview + Known Pitfalls | `New-InstallationDocument`, OpenXML-Gotchas |
| `deploy-example.psd1` | Kommentare | Neue Params dokumentieren |

---

## Bekannte Fallstricke (für `Known Pitfalls` in CLAUDE.md)

- **ZIP-Encoding** — `.docx` verlangt **UTF-8 ohne BOM**. PowerShell 5.1
  `Set-Content -Encoding UTF8` schreibt BOM → `[System.Text.UTF8Encoding]::new($false)`
- **XML-Escaping** — alle Benutzer-Strings (`$env:COMPUTERNAME`, Org-Namen mit `&`,
  Pfade) durch `[Security.SecurityElement]::Escape()` schleifen
- **Smart Quotes** — XML-Entities (`&#x201C;` / `&#x201D;`) für deutsche Typografie
- **Große Transcript-Extrakte** — CMDLet-Block auf 200 begrenzen
- **Content_Types Extension-Duplikate** — nur ein `<Default>` pro Extension
- **PS2Exe-Kompatibilität** — kein inline C#, `System.IO.Compression` ist
  .NET-Standard und PS2Exe-sicher
- **Content Controls (SDT)** — `<w:sdt>`-Block mit `<w:sdtPr>` + `<w:sdtContent>`;
  im SDT-Content muss mindestens ein `<w:p>` oder `<w:r>` stehen, sonst Repair-Dialog
- **Header/Footer-Relationships** — separate `.rels` für Header/Footer nötig
  (`word/_rels/header1.xml.rels` etc.), `sectPr` in `document.xml` verweist via
  `r:id` auf Header/Footer

---

## Verifikation

1. **Post-Install-Doku (F22)**
   ```powershell
   .\EXpress.ps1 -StandaloneDocument -Language DE -InstallPath C:\Install
   .\EXpress.ps1 -StandaloneDocument -Language EN -CustomerDocument -InstallPath C:\Install
   ```
   Erwartung: zwei `.docx` im Reports-Pfad, öffnen in Word 2016+/LibreOffice
   ohne Reparaturdialog. Kapitel 1–15 vorhanden. `-CustomerDocument` maskiert
   Passwörter + interne IPs.
2. **Struktur-Validierung**
   ```bash
   python scripts/office/validate.py <doc>.docx
   ```
3. **Volllauf** — Lab-Installation; Phase 6 erzeugt HTML + beide Word-Varianten
   ohne zusätzliche Laufzeit > 5 s.
4. **Konzept-Template (F23)** — `tools/Build-KonzeptTemplate.ps1` ausführen,
   beide Outputs in Word/LibreOffice öffnen, Headings-Struktur + Content
   Controls im Fragenkatalog prüfen, committen.
5. **PS2Exe-Build** — `Build.ps1` ausführen; resultierende EXpress.exe in Lab-VM
   testen (keine fehlenden Assemblies zur Laufzeit).
6. **Unicode** — Umlaute in Headings prüfen („Härtungsmaßnahmen",
   „Virtuelle Verzeichnisse", „Öffentliche Ordner").
7. **Public-Folder-Konditional (F22)** — mit und ohne `Get-Mailbox -PublicFolder`
   testen; beide Pfade zeigen korrekte Kapitel 13-Ausgabe.
8. **Hybrid-Konditional (F22)** — mit und ohne `Get-HybridConfiguration` testen;
   Kapitel 12 erscheint nur bei vorhandener Hybrid-Konfiguration.
