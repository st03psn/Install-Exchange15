# Remote Query Setup for EXpress Installation Documentation

`New-InstallationDocument` gathers hardware, pagefile, volume, and NIC
details from every Exchange server in the organisation. It uses **CIM over
WSMan (WinRM)** — the same transport the Exchange Management Shell already
relies on.

This document describes the minimum standard that must be provisioned on
every server to be documented.

---

## Transport and boundaries

| | |
|---|---|
| Transport | CIM over WSMan (WinRM), **not** WMI/DCOM |
| Ports | TCP 5985 (HTTP, default), TCP 5986 (HTTPS, optional) |
| Authentication | Kerberos (domain) — **no** `TrustedHosts` |
| Timeout | 30 s per server |
| Rights | Read-only — exclusively `Get-CimInstance` calls |
| Degradation | On failure: inline note in the document, the run continues |

---

## Option A: Script (run locally on each server)

```powershell
\\filer\tools\Enable-EXpressRemoteQuery.ps1                   # HTTP listener is enough on LAN
\\filer\tools\Enable-EXpressRemoteQuery.ps1 -EnableHttps      # add HTTPS listener (5986)
\\filer\tools\Enable-EXpressRemoteQuery.ps1 -RestrictToGroup 'EXpress-DocReader'
```

Idempotent — repeated invocations are harmless.

---

## Option B: Group Policy (recommended for domains)

Create GPO `Exchange — EXpress Remote Query`, linked to the OU containing
the Exchange servers.

### 1. WinRM service

**Computer Configuration → Policies → Windows Settings → Security Settings → System Services**

- `Windows Remote Management (WS-Management)` → **Automatic**

### 2. WinRM service configuration

**Computer Configuration → Policies → Administrative Templates → Windows Components → Windows Remote Management (WinRM) → WinRM Service**

| Setting | Value |
|---|---|
| Allow remote server management through WinRM | **Enabled**, IPv4 filter `*`, IPv6 filter `*` |
| Allow Kerberos authentication | **Enabled** |
| Allow unencrypted traffic | **Not configured** (traffic stays encrypted) |

### 3. Firewall

**Computer Configuration → Policies → Windows Settings → Security Settings → Windows Defender Firewall with Advanced Security → Inbound Rules**

Create a new rule or enable the existing one:

| | |
|---|---|
| Name | `Windows Remote Management (HTTP-In)` |
| Protocol | TCP |
| Local port | 5985 |
| Profile | Domain, Private |
| Action | Allow |
| Remote IP | Management subnet or Exchange server subnet |

For HTTPS, additionally:

| | |
|---|---|
| Name | `Windows Remote Management (HTTPS-In)` |
| Protocol | TCP |
| Local port | 5986 |

### 4. (Optional) Restrict access to a group

Apply PSSessionConfiguration ACL via GPO preferences script or DSC,
granting `BUILTIN\Administrators` + `DOMAIN\EXpress-DocReader`.
Alternatively use `Enable-EXpressRemoteQuery.ps1 -RestrictToGroup`.

---

## Verification

From the management host:

```powershell
Test-WSMan -ComputerName ex01.contoso.local
Get-CimInstance Win32_OperatingSystem -CimSession (New-CimSession -ComputerName ex01.contoso.local)
```

Expected: a WSMan identity response plus an OS object with Caption/Version.
See below for failure cases.

---

## Failure modes

| Error | Cause | Remedy |
|---|---|---|
| `WinRM cannot complete the operation` | Service stopped or firewall blocks | Apply the script or the GPO |
| `Access is denied` | Caller is not a local admin / not in `Remote Management Users` | Add the account to the target group, or query with an admin account |
| `The WinRM client cannot process the request. Kerberos authentication failed` | Target is not domain-joined, or SPN missing | Domain-join the server, or use an HTTPS listener with hostname |
| `The connection to the specified remote host was refused` | No listener configured | `winrm quickconfig`, or run the script |

---

## Hardening recommendations

- **Never set `TrustedHosts *`** — Kerberos inside the domain is sufficient
- Use the HTTPS listener (5986) with the Exchange auth certificate if the
  management network is not trusted
- Restrict access via an AD group such as `EXpress-DocReader`; EXpress
  requires read-only rights
