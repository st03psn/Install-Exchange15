# Plan: Advanced Configuration Menu (F25)

Target: next feature release (post-5.93 or bundled with v1.0).

WIP scaffolding lives in the `feature/advanced-menu` stash (commit `787b221`, 319-line diff against `Install-Exchange15.ps1`). That branch/stash exists only as a reference — master is authoritative. When F25 resumes, cherry-pick the four new functions out of the stash and continue from the call-site migration checklist below.

---

## Goal

Replace today's flat main menu (every toggle shown at once) with a two-level UX:

1. **Main menu** — installation-flow toggles only (A Autopilot, B Install SU, R Windows Updates, U Generate Doc, N Preflight-only).
2. **Advanced Configuration menu** — paginated by category, offered after the main menu with a 60-second auto-skip prompt. Default = skip (keep current hardening defaults). Covers ~55 mitigations/tweaks.

Non-goal: changing *what* gets applied by default. The default state of every toggle matches current v5.x behaviour. The menu only exposes opt-outs and non-default knobs.

## Always-on (never exposed as toggles)

Non-negotiable for Exchange to work correctly — hardcoded, no menu entry:

- Defender Exclusions + transient realtime / Tamper-Protection disable during install
- Page File fixed size
- Virtual Directory URLs (when namespace is configured)
- Register Exchange Log Cleanup scheduled task
- High-Performance Power Plan

## Design

### Four new functions

1. **`Get-AdvancedFeatureCatalog`** — returns `[ordered]@{}` hashtable. Each entry:
   ```powershell
   Name = @{ Category; Label; Description; Default; Condition (optional) }
   ```
   `Condition` is a scriptblock; entries whose condition returns `$false` are hidden (e.g. TLS 1.3 pre-WS2022, Shadow Redundancy without DAG).

2. **`Show-AdvancedMenu`** — paginated-by-category interactive menu.
   - One page per category. Two-column layout, auto-split at ceil(count/2).
   - Letter keys A..Z toggle entry on current page. N=Next, P=Prev, A=Apply (last page), S=Skip-all (use defaults), ESC=Cancel.
   - Description panel updates as the user presses letters; highlighted row = last-touched entry.
   - Falls back to `Read-Host` when `$host.UI.RawUI.KeyAvailable` is unavailable (PS2Exe/redirected hosts).
   - Returns `@{Name=$bool; ...}` (all visible entries) or `$null` on cancel.

3. **`Invoke-AdvancedConfigurationPrompt`** — 60-second countdown prompt *before* the menu. Default answer = skip. Skips silently in Autopilot or when `$State['SuppressAdvancedPrompt']` is set. On confirm, calls `Show-AdvancedMenu`, stores result in `$State['AdvancedFeatures']`, persists via `Save-State`. Uses `Write-Progress -Id 2` (countdown bar).

4. **`Test-Feature -Name <string>`** — single gate used at every call site. Precedence: `$State['AdvancedFeatures'][Name]` > catalog default > `$false` (fail closed) + verbose warning for unknown names.

### Wiring into call sites

Every existing hardening/tuning action gains a `Test-Feature` gate. Example:

```powershell
# before
Disable-SMBv1
# after
if (Test-Feature 'SMBv1Disable') { Disable-SMBv1 }
```

~40 call sites total, one per catalog entry. Mechanical change, one commit per category to keep diffs reviewable.

### Config-file parity

`deploy-example.psd1` grows a nested block (Option A — grouped, inline short descriptions):

```powershell
AdvancedFeatures = @{
    # Security / TLS
    DisableSSL3   = $true   # Disable legacy SSL 3.0 (POODLE)
    DisableRC4    = $true   # Disable RC4 cipher
    EnableECC     = $true   # Prefer ECC key exchange
    # Security / Hardening
    SMBv1Disable  = $true
    LSAProtection = $true
    # ... etc.
}
```

Backwards-compatible: existing top-level keys (`-DisableSSL3`, `-EnableTLS12`, etc.) still work; when both are present, the nested block wins (explicit > implicit).

---

## Feature Catalog

Six categories, ~55 entries. Defaults match current v5.x behaviour unless flagged otherwise.

### Security / TLS
| Name | Label | Default | Description |
|---|---|---|---|
| DisableSSL3 | Disable SSL 3.0 | ✓ | POODLE, CVE-2014-3566 |
| DisableRC4 | Disable RC4 cipher | ✓ | Deprecated stream cipher |
| EnableECC | Prefer ECC key exchange | ✓ | ECC suites + preference over RSA |
| NoCBC | Disable CBC ciphers | ✗ | Breaks several clients — not recommended |
| EnableAMSI | Enable AMSI | ✓ | AMSI for transport and OWA |
| EnableTLS12 | Enforce TLS 1.2 | ✓ | Disable TLS 1.0/1.1 + .NET StrongCrypto |
| EnableTLS13 | Enable TLS 1.3 | ✓ | WS2022+; hidden on older OS |
| DoNotEnableEP | Opt-out: Extended Protection | ✗ | Required for Hybrid/Modern Hybrid Topology |

### Security / Hardening
| Name | Label | Default | Description |
|---|---|---|---|
| SMBv1Disable | Disable SMBv1 | ✓ | WannaCry mitigation, MS17-010 |
| NetBIOSDisable | Disable NetBIOS/TCP | ✓ | Reduce attack surface on all NICs |
| LLMNRDisable | Disable LLMNR | ✓ | CIS L1 §18.5.4.2 |
| MDNSDisable | Disable mDNS | ✓ | WS2022+ multicast DNS responder |
| WDigestDisable | Disable WDigest caching | ✓ | Prevent plaintext creds in LSASS |
| LSAProtection | Enable LSA Protection | ✓ | RunAsPPL for LSASS |
| LmCompat5 | LmCompatibilityLevel=5 | ✓ | Enforce NTLMv2, refuse LM/NTLMv1 |
| SerializedDataSig | SerializedDataSigning | ✓ | MS-mandatory post CVE-2023-21529 |
| ShutdownTrackerOff | Disable Shutdown Tracker | ✓ | Suppress shutdown-reason dialog |
| HTTP2Disable | Disable HTTP/2 | ✓ | Exchange HTTP/2 compat workaround |
| CredentialGuardOff | Disable Credential Guard | ✓ | Incompatible with Exchange |
| UnnecessaryServices | Disable unneeded services | ✓ | Spooler, Xbox, Geolocation, … |
| WindowsSearchOff | Disable Windows Search | ✓ | Unused; saves CPU/IO |
| CRLTimeout | CRL Check Timeout | ✓ | Avoid slow startup on unreachable OCSP/CRL |
| RootCAAutoUpdate | Root CA Auto-Update | ✓ | Required for Modern Auth / O365 Hybrid |
| SMTPBannerHarden | Harden SMTP banner | ✓ | Replace version banner with `220 Mail Service` |

### Performance / Tuning
| Name | Label | Default | Description |
|---|---|---|---|
| MaxConcurrentAPI | MaxConcurrentAPI | ✓ | MS KB 2688798 — NTLM auth bottleneck |
| DiskAllocHint | Disk allocation hint | ✓ | Warn if DB/log volumes not 64K NTFS |
| CtsProcAffinity | Content conv. affinity | ✓ | Stabilise CPU load |
| NodeRunnerMemLimit | NodeRunner RAM cap | ✓ | Prevent runaway allocations |
| MapiFeGC | MAPI FE Server GC | ✓ | Server GC mode for MAPI FE AppPool |
| NICPowerMgmtOff | NIC Power Management | ✓ | Disable "allow turn off this device" |
| RSSEnable | Receive Side Scaling | ✓ | Multi-core packet processing |
| TCPTuning | TCP tuning | ✓ | Autotuning + Chimney + stack tweaks |
| TCPOffloadOff | Disable TCP offload | ✓ | Avoid driver bugs on Exchange |
| IPv4OverIPv6Off | Disable IPv4-over-IPv6 | ✓ | DisabledComponents=0x20; avoid DNS delay on IPv6-only hosts |

### Exchange Org Policy
| Name | Label | Default | Description |
|---|---|---|---|
| ModernAuth | Modern Auth (OAuth2) | ✓ | Org-wide; required for Outlook 2016+, Teams, mobile |
| OWASessionTimeout6h | OWA Session Timeout 6h | ✓ | Activity-based OWA/ECP timeout |
| DisableTelemetry | Disable CEIP telemetry | ✓ | `-CustomerFeedbackEnabled $false` (GDPR) |
| MapiHttp | MAPI over HTTP | ✓ | Replaces legacy RPC/HTTP |
| MaxMessageSize150MB | Max message size 150MB | ✓ | Org + FE receive connector |
| MessageExpiration7d | Expiration 7 days | ✓ | Hidden when CopyServerConfig |
| HtmlNDR | HTML NDR formatting | ✓ | `-InternalDsnSendHtml / -ExternalDsnSendHtml` |
| ShadowRedundancy | Shadow Redundancy | ✗ | DAG-only; hidden without DAG |
| SafetyNet2d | Safety Net 2d hold | ✓ | Safety Net hold time |

### Post-Config / Integration
| Name | Label | Default | Description |
|---|---|---|---|
| MECA | MECA Auth Cert Renewal | ✓ | CSS-Exchange scheduled task |
| AntispamAgents | Install Antispam Agents | ✓ | Mailbox role only |
| SSLOffloading | SSL Offloading tuning | ✓ | IIS/OWA for load-balanced deployments |
| MRSProxy | Enable MRS Proxy | ✓ | Cross-forest/cross-org mailbox moves |
| IANATimezone | IANA timezone mapping | ✓ | iCal interop |
| AnonymousRelay | Anonymous relay connector | ✓ | Hidden without `RelaySubnets` |
| SkipHealthCheck | Opt-out: HealthChecker | ✗ | Skip HC run at end of Phase 6 |
| RBACReport | RBAC Report | ✓ | HTML role-group / role-assignment report |
| RunEOMT | Run EOMT | ✗ | Legacy CUs; no-op on current |

### Install-Flow / Debug
| Name | Label | Default | Description |
|---|---|---|---|
| DiagnosticData | Send diagnostic data | ✗ | `/IAcceptExchangeServerLicenseTerms_DiagnosticDataON` |
| Lock | Lock screen during run | ✗ | Autopilot only |
| SkipRolesCheck | Skip AD roles check | ✗ | Bypass Schema/Enterprise/Domain Admin check |
| NoCheckpoint | Skip System Restore | ✗ | Skip pre-install checkpoints |
| NoNet481 | Skip .NET 4.8.1 | ✗ | Debug only — may break setup |
| WaitForADSync | Wait for AD replication | ✗ | Up to 6 min post-PrepareAD |

---

## Implementation Checklist

Scaffolding (already WIP on feature branch):

- [x] `Get-AdvancedFeatureCatalog` — ordered hashtable
- [x] `Show-AdvancedMenu` — paginated interactive menu
- [x] `Invoke-AdvancedConfigurationPrompt` — 60s auto-skip prompt
- [x] `Test-Feature` — gate with precedence rules

Open work:

- [ ] Shrink main menu to installation-flow toggles (A / B / R / U / N)
- [ ] Call-site migration: wrap ~40 hardening/tuning invocations with `Test-Feature 'Name'` (one commit per catalog category)
- [ ] `deploy-example.psd1`: add `AdvancedFeatures = @{...}` nested block; preserve backwards-compat with top-level keys
- [ ] Config loader: merge top-level keys into `$State['AdvancedFeatures']` for backwards-compat (top-level wins unless both present, then nested wins)
- [ ] Autopilot path: wire `Invoke-AdvancedConfigurationPrompt` into the main-menu flow (skip silently when `$State['Autopilot']`)
- [ ] Restore WIP scaffolding: `git stash pop` from commit `787b221`, or cherry-pick `feature/advanced-menu` WIP
- [ ] README + RELEASE-NOTES: document Advanced menu, new config block, migration note for scripted installs
- [ ] Test matrix: fresh install (defaults), fresh install (all opt-outs), Autopilot with `AdvancedFeatures` block, Autopilot without (→ defaults), interactive skip, interactive cancel
- [ ] Deprecation note in CLAUDE.md once the old top-level keys are gone (v1.x milestone after F25 lands)

---

## Risks / Notes

- **Call-site drift** — every new hardening added between now and F25 landing must also get a catalog entry + `Test-Feature` gate. Add to PR checklist once F25 merges.
- **PS2Exe raw-key path** — `Show-AdvancedMenu` already falls back to `Read-Host` when `RawUI.KeyAvailable` is unavailable. Test explicitly against the PS2Exe build before release.
- **Config-file backwards compat** — nested-block + top-level-key precedence must be covered by a dedicated unit test in `Install-Exchange15.Tests.ps1`; getting the precedence wrong silently changes behaviour for existing `.psd1` consumers.
- **60-second auto-skip** — default is skip, so a user who hits Enter by accident keeps current v5.x behaviour. Don't invert the default without a release-notes deprecation cycle.
