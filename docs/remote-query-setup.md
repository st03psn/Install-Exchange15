# Remote-Query-Setup für EXpress Installationsdokumentation

`New-InstallationDocument` erfasst Hardware-, Pagefile-, Volume- und NIC-Details
aller Exchange-Server einer Organisation. Dafür wird **CIM über WSMan (WinRM)**
verwendet — derselbe Transport, den die Exchange Management Shell ohnehin nutzt.

Dieses Dokument beschreibt den verbindlichen, minimalen Standard, der auf jedem
zu dokumentierenden Server bereitgestellt werden muss.

---

## Transport und Grenzen

| | |
|---|---|
| Transport | CIM über WSMan (WinRM), **nicht** WMI/DCOM |
| Ports | TCP 5985 (HTTP, Default), TCP 5986 (HTTPS, optional) |
| Auth | Kerberos (Domäne) — **kein** `TrustedHosts` |
| Timeout | 30 s pro Server |
| Rechte | Read-only — ausschließlich `Get-CimInstance`-Aufrufe |
| Degradation | Bei Fehlschlag: Hinweistext im Dokument, Lauf fährt fort |

---

## Weg A: Script (pro Server lokal ausführen)

```powershell
\\filer\tools\Enable-EXpressRemoteQuery.ps1                   # HTTP-Listener reicht im LAN
\\filer\tools\Enable-EXpressRemoteQuery.ps1 -EnableHttps      # zusätzlich HTTPS (5986)
\\filer\tools\Enable-EXpressRemoteQuery.ps1 -RestrictToGroup 'EXpress-DocReader'
```

Idempotent — mehrfache Aufrufe sind unschädlich.

---

## Weg B: Gruppenrichtlinie (empfohlen für Domänen)

Neue GPO `Exchange — EXpress Remote Query`, verlinkt an die OU mit den
Exchange-Servern.

### 1. WinRM-Dienst

**Computerkonfiguration → Richtlinien → Windows-Einstellungen → Sicherheitseinstellungen → Systemdienste**

- `Windows Remote Management (WS-Management)` → **Automatisch**

### 2. WinRM-Service-Konfiguration

**Computerkonfiguration → Richtlinien → Administrative Vorlagen → Windows-Komponenten → Windows-Remoteverwaltung (WinRM) → WinRM-Dienst**

| Einstellung | Wert |
|---|---|
| Remoteserververwaltung über WinRM zulassen | **Aktiviert**, IPv4-Filter `*`, IPv6-Filter `*` |
| Kerberos-Authentifizierung zulassen | **Aktiviert** |
| Nicht verschlüsselten Datenverkehr zulassen | **Nicht konfiguriert** (bleibt verschlüsselt) |

### 3. Firewall

**Computerkonfiguration → Richtlinien → Windows-Einstellungen → Sicherheitseinstellungen → Windows Defender Firewall mit erweiterter Sicherheit → Eingehende Regeln**

Neue Regel oder bestehende aktivieren:

| | |
|---|---|
| Name | `Windows Remote Management (HTTP-In)` |
| Protokoll | TCP |
| Lokaler Port | 5985 |
| Profil | Domäne, Privat |
| Aktion | Zulassen |
| Remote-IP | Management-Subnetz bzw. Exchange-Server-Subnetz |

Für HTTPS zusätzlich:

| | |
|---|---|
| Name | `Windows Remote Management (HTTPS-In)` |
| Protokoll | TCP |
| Lokaler Port | 5986 |

### 4. (Optional) Zugriff auf Gruppe einschränken

PSSessionConfiguration-ACL via GPO-Preferences-Script oder DSC setzen auf
`BUILTIN\Administrators` + `DOMAIN\EXpress-DocReader`. Alternativ über das Script
`Enable-EXpressRemoteQuery.ps1 -RestrictToGroup`.

---

## Verifikation

Vom Management-Host aus:

```powershell
Test-WSMan -ComputerName ex01.contoso.local
Get-CimInstance Win32_OperatingSystem -CimSession (New-CimSession -ComputerName ex01.contoso.local)
```

Erwartung: WSMan-Identity + OS-Objekt mit Caption/Version. Fehlerfall siehe
unten.

---

## Fehlerbilder

| Fehler | Ursache | Abhilfe |
|---|---|---|
| `WinRM cannot complete the operation` | Dienst gestoppt oder Firewall blockt | Script oder GPO anwenden |
| `Access is denied` | Kein lokaler Admin / nicht in `Remote Management Users` | Konto in Zielgruppe aufnehmen oder mit Admin-Account abfragen |
| `The WinRM client cannot process the request. Kerberos authentication failed` | Ziel nicht in Domäne oder SPN fehlt | Server domänenjoinen oder HTTPS-Listener mit Hostnamen nutzen |
| `The connection to the specified remote host was refused` | Kein Listener konfiguriert | `winrm quickconfig` bzw. Script |

---

## Härtungsempfehlung

- **Kein `TrustedHosts *`** — Kerberos innerhalb der Domäne ist hinreichend
- HTTPS-Listener (5986) mit Exchange-Auth-Zertifikat, falls Management-Netz nicht
  vertrauenswürdig
- Zugriff per AD-Gruppe `EXpress-DocReader` einschränken; EXpress benötigt
  ausschließlich Leserechte
