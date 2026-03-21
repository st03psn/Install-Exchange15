# Install-Exchange15.ps1 — Projekt-Kontext für Claude Code

---

## Was ist dieses Skript?

`Install-Exchange15.ps1` ist ein PowerShell-Automatisierungsskript zur vollständigen
Installation von Microsoft Exchange Server 2016, 2019 und Exchange SE (Subscription Edition)
inkl. aller Voraussetzungen, AD-Vorbereitung und Post-Konfiguration.

**Autor:** Michel de Rooij (michel@eightwone.com)  
**Aktuelle Version:** 4.22 (Dezember 2025)  
**PowerShell-Anforderung:** `#Requires -Version 5.1`  
**Ausführung:** Muss als Administrator gestartet werden  

---

## Unterstützte Umgebungen

| Exchange-Version | Windows Server |
|---|---|
| Exchange 2016 CU23 | Windows Server 2016 |
| Exchange 2019 CU10–CU14 | Windows Server 2019, 2022 |
| Exchange 2019 CU15+ | Windows Server 2025 |
| Exchange SE RTM | Windows Server 2022, 2025 |

---

## Architektur

### Phasen (InstallPhase 0–6)

Das Skript arbeitet phasenbasiert. Der Zustand wird in einer XML-Datei persistiert,
sodass nach jedem Reboot automatisch in der richtigen Phase fortgefahren wird.

| Phase | Inhalt |
|---|---|
| 0 | Initialisierung, Preflight-Checks, AD-Vorbereitung |
| 1 | Windows-Features installieren, .NET Framework |
| 2 | Reboots abwarten, Prereqs installieren (VC++, URL Rewrite, etc.) |
| 3 | Weitere Prereqs und Hotfixes |
| 4 | Exchange Setup ausführen, Transport-Dienste auf Manual setzen |
| 5 | Post-Konfiguration (Defender-Ausschlüsse, TLS, SSL3, AMSI, ECC, CBC) |
| 6 | Dienste wiederherstellen, IIS-Healthcheck, Aufräumen |

### State-Management

```powershell
$StateFile = "$InstallPath\${env:computerName}_Install-Exchange15_state.xml"
Save-State $State      # Export-Clixml
Restore-State          # Import-Clixml (gibt leere Hashtable wenn nicht vorhanden)
```

Der State-Hashtable enthält alle übergebenen Parameter plus Laufzeitvariablen
(Phase, Versionen, Pfade, Flags).

### AutoPilot-Modus

Mit `-AutoPilot` fährt das Skript nach jeder Phase automatisch den Server neu hoch
und setzt sich selbst per `RunOnce`-Registry-Eintrag fort. Credentials werden
verschlüsselt im State gespeichert.

```powershell
# RunOnce-Eintrag (seit Optimierung: dynamischer PS-Interpreter-Pfad)
$PSExe = (Get-Process -Id $PID).Path   # powershell.exe oder pwsh.exe
```

---

## Wichtige Konstanten

```powershell
# OS-Versionen (Build-Präfix)
$WS2016_MAJOR   = '10.0'
$WS2019_PREFULL = '10.0.17709'
$WS2022_PREFULL = '10.0.20348'
$WS2025_PREFULL = '10.0.26100'   # WICHTIG: war fälschlicherweise '10.0.20348'

# Exchange Setup-Versionen (ExSetup.exe)
$EX2016SETUPEXE_CU23    = '15.01.2507.006'
$EX2019SETUPEXE_CU10–15 = '15.02.xxxx.xxx'
$EXSESETUPEXE_RTM       = '15.02.2562.017'

# .NET Framework
$NETVERSION_48  = 528040
$NETVERSION_481 = 533320

# Autodiscover SCP LDAP-Filter (zentrale Konstante, 4× verwendet)
$AUTODISCOVER_SCP_FILTER    = '(&(cn={0})(objectClass=serviceConnectionPoint)...)'
$AUTODISCOVER_SCP_MAX_RETRIES = 30   # 30 × 10 Sek = 5 Min Timeout
```

---

## Funktions-Übersicht

### Logging

```powershell
Write-ToTranscript $Level $Text   # Interne Hilfsfunktion (neu seit Refactoring)
Write-MyOutput  $Text             # Write-Output + Transcript [INFO]
Write-MyWarning $Text             # Write-Warning + Transcript [WARNING]
Write-MyError   $Text             # Write-Error + Transcript [ERROR]
Write-MyVerbose $Text             # Write-Verbose + Transcript [VERBOSE]
```

### Preflight-Checks (`Test-Preflight`)

Prüft: Admin-Rechte, Domänenmitgliedschaft, OS-Version, Exchange-Version,
AD-Forest/-Domain-Level, statische IP, Rollen, Organisations-Name, Setup-Pfad.

### Paket-Installation

```powershell
Get-MyPackage   $Package $URL $FileName $InstallPath   # Download via BITS
Install-MyPackage $PackageID $Package $FileName $URL $Arguments
Test-MyPackage  $PackageID                              # Registry + WMI-Prüfung
Invoke-Process  $FilePath $FileName $ArgumentList       # MSU/MSI/MSP/EXE
Invoke-Extract  $FilePath $FileName                     # ZIP via Expand-Archive
```

### TLS/Kryptografie

```powershell
Set-SchannelProtocol -Protocol 'TLS 1.2' -Enable $true/$false   # Hilfsfunktion
Set-NetFrameworkStrongCrypto                                      # Hilfsfunktion
Set-TLSSettings -TLS12 -TLS13                                    # Hauptfunktion
Disable-SSL3
Disable-RC4
Enable-ECC
Enable-CBC
Enable-AMSI
```

### AD / Exchange

```powershell
Get-ForestRootNC / Get-RootNC / Get-ForestConfigurationNC
Get-ForestFunctionalLevel / Get-ExchangeForestLevel / Get-ExchangeDomainLevel
Get-ExchangeOrganization / Test-ExchangeOrganization
Test-ExistingExchangeServer $Name
Clear-AutodiscoverServiceConnectionPoint $Name [-Wait]
Set-AutodiscoverServiceConnectionPoint $Name $ServiceBinding [-Wait]
Initialize-Exchange          # PrepareAD / PrepareSchema
```

### Post-Konfiguration

```powershell
Enable-WindowsDefenderExclusions   # Ordner- und Prozess-Ausschlüsse
Enable-HighPerformancePowerPlan
Disable-NICPowerManagement
Set-Pagefile
Set-TCPSettings                    # RPC Timeout, Keep-Alive
```

---

## Bekannte Fallstricke & Designentscheidungen

### 1. CIM statt WMI (vollständig migriert)
Alle `Get-WmiObject`-Aufrufe wurden auf `Get-CimInstance` umgestellt.
Bei Schreibzugriffen: `Set-CimInstance -InputObject $obj -Property @{...}`
statt `$obj.Eigenschaft = ...; $obj.psbase.Put()`.

### 2. Get-WindowsFeature prüft immer `.Installed`
```powershell
# FALSCH - gibt immer ein Objekt zurück, auch wenn nicht installiert
if (Get-WindowsFeature 'Web-Server') { ... }

# RICHTIG
if ((Get-WindowsFeature -Name 'Web-Server').Installed) { ... }
```

### 3. Autodiscover-SCP Background-Jobs
`Clear-` und `Set-AutodiscoverServiceConnectionPoint` starten Jobs mit `do..while($true)`.
Der `$AUTODISCOVER_SCP_MAX_RETRIES`-Counter verhindert Endlosschlaufen.
Filter-Template wird als Parameter übergeben, da Skript-Scope in Jobs nicht verfügbar ist.

### 4. AutoLogon schreibt Klartext-Passwort
`Enable-AutoLogon` schreibt das Passwort nach `HKLM:\...\Winlogon\DefaultPassword`.
`Disable-AutoLogon` entfernt es beim nächsten Login. Intentional by Design.

### 5. `Invoke-WebRequest -SkipCertificateCheck`
Nur ab PowerShell 6+ verfügbar. Bei PS 5.1 Einsatz ggf. Fallback einbauen.

### 6. `$AUTODISCOVER_SCP_FILTER` als Template
Der Filter enthält `{0}` als Platzhalter für den Servernamen:
```powershell
$LDAPSearch.Filter = $AUTODISCOVER_SCP_FILTER -f $Name
```

### 7. Error-Handling in catch-Blöcken
Immer `$_.Exception.Message` verwenden, nicht `$Error[0].ExceptionMessage`
(kann durch parallele Fehler überschrieben werden).

---

## Optimierungs-Historie (2025-03-21)

### Runde 1 — Kritische Fixes
| # | Was | Zeile(n) vorher |
|---|---|---|
| Bug | `$WS2025_PREFULL` = `10.0.26100` (war `10.0.20348` = WS2022) | 645 |
| Refactor | `Write-ToTranscript` Hilfsfunktion, alle 4 `Write-My*` vereinfacht | 709–739 |
| API | `Get-WmiObject` (MSExchangeServiceHost) → `Get-CimInstance` | 2797 |
| API | `WebClient`/`ServerCertificateValidationCallback` → `Invoke-WebRequest -SkipCertificateCheck` | 2838–2851 |
| Feature | `Test-RebootPending` um Windows-Update-Key ergänzt | 805–814 |

### Runde 2 — Sicherheit & Codequalität
| # | Was |
|---|---|
| Konstanten | `$AUTODISCOVER_SCP_FILTER` + `$AUTODISCOVER_SCP_MAX_RETRIES` eingeführt |
| Sicherheit | Kommentar in `Enable-AutoLogon` zu Klartext-Risiko |
| API | `Enable-RunOnce`: `$PSHome\powershell.exe` → `(Get-Process -Id $PID).Path` |
| API | `Invoke-Extract`: COM `shell.application` → `Expand-Archive` |
| Bug | Endlosschleifen in SCP-Background-Jobs: Retry-Limit + Timeout |
| API | `Get-WmiObject win32_quickfixengineering` → `Get-CimInstance Win32_QuickFixEngineering` |
| Konvention | `get-FullDomainAccount` → `Get-FullDomainAccount` |
| Typo | `'Wil run Setup'` → `'Will run Setup'` |
| Exception | `$Error[0].ExceptionMessage` → `$_.Exception.Message` in allen catch-Blöcken |

### Runde 3 — Weitere WMI-Migration & Bugs
| # | Was |
|---|---|
| API | `mkdir` → `New-Item -ItemType Directory` |
| Bug | `Remove-NETFrameworkInstallBlock`: Fehlermeldung "set" → "remove" |
| Bug | Streunende `$Location`-Ausgabe in `Enable-WindowsDefenderExclusions` entfernt |
| API | `$CS = Get-WmiObject Win32_ComputerSystem` + `.Put()` → CIM + `Set-CimInstance` |
| API | `Get-WmiObject Win32_NetworkAdapter` + `MSPower_DeviceEnable` + `psbase.Put()` → CIM |
| API | `Get-WmiObject Win32_ComputerSystem/Win32_NetworkAdapterConfiguration` → CIM |
| API | `Get-WmiObject Win32_PowerPlan` → CIM |
| Bug | `Get-WindowsFeature` Prüfung: `if (Get-WindowsFeature $x)` → `.Installed` |
| Refactor | `Set-TLSSettings`: 50 Zeilen Duplikatcode → `Set-SchannelProtocol` + `Set-NetFrameworkStrongCrypto` |
| Kosmetik | `$Env:SystemRoot` ohne unnötige String-Interpolation |

---

## Offene Punkte / Mögliche nächste Schritte

- [ ] `Invoke-WebRequest -SkipCertificateCheck` PS 5.1-Fallback einbauen
- [ ] `$Error[0]` Vorkommen vollständig auditieren (außerhalb von catch-Blöcken)
- [ ] Pester-Tests für die wichtigsten Hilfsfunktionen
- [ ] Parameter-Block-Redundanz reduzieren (viele Parameter mit 4× identischen `[parameter()]`-Attributen)
- [ ] `Get-SetupTextVersion` effizienter gestalten (direkter Hashtable-Lookup)
