# Install-Exchange15.ps1

PowerShell-Skript zur vollautomatischen Installation von Microsoft Exchange Server 2016, 2019 und Exchange SE — inklusive Voraussetzungen, Active-Directory-Vorbereitung und Post-Konfiguration.

**Autor:** Michel de Rooij (michel@eightwone.com) · [eightwone.com](http://eightwone.com)  
**Version:** 4.22 (Dezember 2025)  
**Lizenz:** As-Is, ohne Gewährleistung

---

## Unterstützte Versionen

| Exchange | Windows Server |
|---|---|
| Exchange 2016 CU23 | Windows Server 2016 |
| Exchange 2019 CU10–CU14 | Windows Server 2019, 2022 |
| Exchange 2019 CU15+ / Exchange SE | Windows Server 2022, 2025 |

---

## Voraussetzungen

- PowerShell 5.1 oder höher
- Ausführung als lokaler Administrator
- Domänenmitgliedschaft (außer Edge-Rolle)
- Schema Admin + Enterprise Admin Rechte für AD-Vorbereitung
- Statische IP-Adresse (oder Azure-Guest-Agent erkannt)
- Exchange-Setup-Dateien (ISO oder entpackt) erreichbar

---

## Verwendung

```powershell
# Mailbox-Rolle installieren (interaktiv)
.\Install-Exchange15.ps1 -InstallMailbox -SourcePath D:\Exchange

# Vollautomatisch mit AutoPilot (rebootet automatisch durch alle Phasen)
.\Install-Exchange15.ps1 -InstallMailbox -SourcePath D:\Exchange -AutoPilot -Credentials (Get-Credential)

# Nur Voraussetzungen installieren, kein Exchange-Setup
.\Install-Exchange15.ps1 -NoSetup

# Edge-Transport-Rolle
.\Install-Exchange15.ps1 -InstallEdge -SourcePath D:\Exchange

# Server wiederherstellen
.\Install-Exchange15.ps1 -Recover -SourcePath D:\Exchange
```

### Wichtige Parameter

| Parameter | Beschreibung |
|---|---|
| `-InstallMailbox` | Installiert die Mailbox-Rolle |
| `-InstallEdge` | Installiert die Edge-Transport-Rolle |
| `-SourcePath` | Pfad zur Exchange-Setup-Datei / ISO |
| `-TargetPath` | Zielordner für Exchange (Standard: `C:\Program Files\Microsoft\Exchange Server\V15`) |
| `-AutoPilot` | Vollautomatischer Modus mit automatischen Reboots |
| `-Credentials` | Anmeldedaten für AutoPilot |
| `-OrganizationName` | Name der Exchange-Organisation (bei Neuinstallation) |
| `-InstallMDBName` | Name der ersten Mailbox-Datenbank |
| `-InstallMDBDBPath` | Pfad für Datenbankdateien (.edb) |
| `-InstallMDBLogPath` | Pfad für Transaktionsprotokolle |
| `-IncludeFixes` | Installiert empfohlene Sicherheitsupdates nach Setup |
| `-DisableSSL3` | Deaktiviert SSL 3.0 (POODLE) |
| `-DisableRC4` | Deaktiviert RC4-Verschlüsselung |
| `-EnableTLS12` | Aktiviert TLS 1.2 explizit |
| `-EnableTLS13` | Aktiviert TLS 1.3 (WS2022+, Exchange 2019 CU15+) |
| `-EnableECC` | Aktiviert ECC-Zertifikate |
| `-EnableAMSI` | Aktiviert AMSI-Body-Scanning |
| `-NoSetup` | Installiert nur Prereqs, überspringt Exchange-Setup |
| `-Phase` | Startet direkt in einer bestimmten Phase (0–6) |

---

## Ablauf

Das Skript durchläuft 7 Phasen (0–6) und speichert den Zustand in einer XML-Datei,
um nach Reboots automatisch fortzufahren:

```
Phase 0 → Preflight-Checks, AD-Vorbereitung
Phase 1 → Windows-Features, .NET Framework
Phase 2 → Visual C++ Redistributables, URL Rewrite, weitere Prereqs
Phase 3 → Hotfixes, zusätzliche Pakete
Phase 4 → Exchange Setup ausführen
Phase 5 → Post-Konfiguration (Defender, TLS, Power Plan, Pagefile, TCP)
Phase 6 → Dienste hochfahren, IIS-Healthcheck, Cleanup
```

---

## Änderungen gegenüber Original (v4.22 — Optimierungsrunde März 2025)

### Bugfixes

- **`$WS2025_PREFULL`** korrigiert: `10.0.26100` (war fälschlicherweise `10.0.20348` = WS2022)
- **`Get-WindowsFeature`-Prüfung** korrigiert: verwendet jetzt `.Installed`-Property statt implizitem Boolean-Cast
- **Fehlermeldung** in `Remove-NETFrameworkInstallBlock` korrigiert: „Unable to remove" statt „Unable to set"
- **Streunende Konsolenausgabe** in `Enable-WindowsDefenderExclusions` entfernt
- **Endlosschleifen** in Autodiscover-SCP-Background-Jobs: Retry-Limit von 30 × 10 Sek eingebaut
- **`$Error[0].ExceptionMessage`** in allen `catch`-Blöcken durch `$_.Exception.Message` ersetzt
- **Typo** `'Wil run Setup'` → `'Will run Setup'`

### API-Modernisierung

| Alt | Neu |
|---|---|
| `Get-WmiObject` (alle 9 Stellen) | `Get-CimInstance` |
| `$obj.psbase.Put()` | `Set-CimInstance -InputObject $obj -Property @{...}` |
| `New-Object Net.WebClient` + `ServerCertificateValidationCallback` | `Invoke-WebRequest -SkipCertificateCheck -UseBasicParsing` |
| `New-Object -com shell.application` (ZIP-Extraktion) | `Expand-Archive` |
| `$PSHome\powershell.exe` (RunOnce) | `(Get-Process -Id $PID).Path` (PS 7-kompatibel) |
| `mkdir` | `New-Item -ItemType Directory` |

### Refactoring

- **Logging:** Neue interne Hilfsfunktion `Write-ToTranscript` — eliminiert 4× duplizierte `Test-Path`/`Out-File`-Logik in `Write-My*`
- **TLS:** Neue Hilfsfunktionen `Set-SchannelProtocol` und `Set-NetFrameworkStrongCrypto` — reduziert `Set-TLSSettings` von ~90 auf ~35 Zeilen
- **LDAP-Filter:** Konstante `$AUTODISCOVER_SCP_FILTER` eingeführt (war 4× identisch hardcodiert)
- **Funktionsname:** `get-FullDomainAccount` → `Get-FullDomainAccount` (PS-Konvention)
- **`Test-RebootPending`:** Dritter Registry-Check für Windows Update hinzugefügt

### Sicherheit

- **`Enable-AutoLogon`:** Kommentar zum Klartext-Passwort-Risiko in der Registry hinzugefügt

---

## Hinweise

- Das Skript legt eine Zustandsdatei unter `%TEMP%\<Computername>_Install-Exchange15_state.xml` an
- Log-Datei: `%TEMP%\<Computername>_Install-Exchange15.log`
- Bei `-AutoPilot`: UAC wird temporär deaktiviert und nach Abschluss wieder aktiviert
- Das Passwort für AutoLogon wird nach dem nächsten Login automatisch aus der Registry entfernt
- `-SkipCertificateCheck` bei `Invoke-WebRequest` erfordert PowerShell 6+; bei reinem PS 5.1 ggf. Fallback nötig

---

## Quellen & Dokumentation

- [Exchange Server Build Numbers](https://docs.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates)
- [Exchange 2019 Prerequisites](https://docs.microsoft.com/en-us/exchange/plan-and-deploy/prerequisites)
- [eightwone.com Blog](http://eightwone.com)
