# EXpress — Pre-staged Sources

This folder is the **downloads cache** for EXpress. Any file placed here is used directly
during installation without downloading it again — enabling fully air-gapped or
proxy-restricted deployments.

Run `tools\Get-EXpressDownloads.ps1` on an internet-connected machine to populate this
folder automatically, then copy the entire `sources\` directory to the target server.

---

## Prerequisites

| File | Description | Size | Download |
|---|---|---|---|
| `vcredist_x64_2012.exe` | Visual C++ 2012 Redistributable (x64) | ~6 MB | [Microsoft Download Center](https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe) |
| `vcredist_x64_2013.exe` | Visual C++ 2013 Redistributable (x64, ≥ 12.0.40664) | ~7 MB | [aka.ms/highdpimfc2013x64enu](https://aka.ms/highdpimfc2013x64enu) |
| `rewrite_amd64_en-US.msi` | IIS URL Rewrite Module 2.1 | ~1.5 MB | [Microsoft Download Center](https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_en-US.msi) |
| `UcmaRuntimeSetup.exe` | Unified Communications Managed API 4.0 Runtime | ~240 MB | [Microsoft Download Center](https://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe) |
| `NDP48-x86-x64-AllOS-ENU.exe` | .NET Framework 4.8 | ~116 MB | [Microsoft Download Center](https://go.microsoft.com/fwlink/?linkid=2088631) |
| `NDP481-x86-x64-AllOS-ENU.exe` | .NET Framework 4.8.1 | ~100 MB | [Microsoft Download Center](https://download.microsoft.com/download/4/b/2/cd00d4ed-ebdd-49ee-8a33-eabc3d1030e3/NDP481-x86-x64-AllOS-ENU.exe) |

> **.NET 4.8.1 is built into Windows Server 2025.** Use `-SkipDotNet` with `Get-EXpressDownloads.ps1`
> when targeting WS2025-only environments to skip the large .NET downloads.

---

## CSS-Exchange Tools

Downloaded automatically from the [CSS-Exchange latest release](https://github.com/microsoft/CSS-Exchange/releases/latest).
Pre-stage by placing the files here before running EXpress.

| File | Purpose |
|---|---|
| `HealthChecker.ps1` | Post-install health analysis — run at end of Phase 6 |
| `SetupAssist.ps1` | Diagnoses Exchange setup failures in Phase 4 |
| `SetupLogReviewer.ps1` | Companion to SetupAssist — reviews setup log |
| `ExchangeExtendedProtectionManagement.ps1` | Extended Protection configuration (pre-CU14) |
| `MonitorExchangeAuthCertificate.ps1` | Auth Certificate auto-renewal scheduled task (MEAC) |
| `EOMT.ps1` | Emergency Mitigation Tool — only needed when `RunEOMT` is enabled |

---

## Exchange Security Updates

Downloaded automatically when `IncludeFixes = $true`. Pre-stage by placing the installer
here with the exact filename shown below.

| File | Exchange version | KB | Download |
|---|---|---|---|
| `ExchangeSubscriptionEdition-KB5074992-x64-en.exe` | Exchange SE RTM (15.02.2562.017) | KB5074992 | [support.microsoft.com/help/5074992](https://support.microsoft.com/help/5074992) |
| `Exchange2019-KB5049233-x64-en.exe` | Exchange 2019 CU13–CU15 | KB5049233 | [support.microsoft.com/help/5049233](https://support.microsoft.com/help/5049233) |
| `Exchange2016-KB5049233-x64-en.exe` | Exchange 2016 CU23 | KB5049233 | [support.microsoft.com/help/5049233](https://support.microsoft.com/help/5049233) |

> Exchange SU files are version-specific. Verify the exact KB number against your
> Exchange build before downloading. Check [Exchange build numbers](https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates)
> for the current SU for your CU.

---

## Optional

| File | Purpose |
|---|---|
| `logo.png` | Company logo embedded in HTML and Word reports. EXpress also checks `assets\logo.png` next to the script. |

> **Note:** `UcmaRedist\Setup.exe` (UCMA Server Core offline installer) must be copied
> from the Exchange installation media — it is not available as a standalone download.
> Extract it from the Exchange ISO at `\UCMARedist\Setup.exe`.
