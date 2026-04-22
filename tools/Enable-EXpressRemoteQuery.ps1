<#
.SYNOPSIS
    Enables WinRM / CIM over WSMan for EXpress remote documentation queries.

.DESCRIPTION
    EXpress's New-InstallationDocument collects hardware, pagefile, volume and
    NIC information from every Exchange server in the organization via CIM over
    WSMan (WinRM). On Exchange servers WinRM is already enabled for the
    Management Shell, but it may be disabled on freshly deployed or hardened
    hosts.

    This script is the official, supported pre-requisite for the remote query
    path used by EXpress. It is read-only from the target's perspective:
    EXpress only issues Get-CimInstance calls.

    Run locally on every server that should be documented, or deploy the
    equivalent settings via GPO (see docs/remote-query-setup.md).

.PARAMETER EnableHttps
    Adds an HTTPS listener on TCP 5986 using the Exchange auth certificate (or
    the first server-auth certificate found). Recommended for cross-segment or
    untrusted networks.

.PARAMETER RestrictToGroup
    AD group name (e.g. 'EXpress-DocReader'). If supplied, the script adjusts
    the PSSessionConfiguration ACL so only members of that group can connect.
    The group must already exist in AD.

.EXAMPLE
    .\Enable-EXpressRemoteQuery.ps1

.EXAMPLE
    .\Enable-EXpressRemoteQuery.ps1 -EnableHttps -RestrictToGroup 'EXpress-DocReader'

.NOTES
    Requires: Administrator, PS 5.1+, domain-joined Windows Server 2016+.
    Ports:    TCP 5985 (HTTP, default), TCP 5986 (HTTPS, optional).
    Auth:     Kerberos (domain default). TrustedHosts is never modified.
#>
#Requires -Version 5.1
#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [switch]$EnableHttps,
    [string]$RestrictToGroup
)

$ErrorActionPreference = 'Stop'

function Write-Step { param([string]$Msg) Write-Host "[*] $Msg" -ForegroundColor Cyan }
function Write-Ok   { param([string]$Msg) Write-Host "[+] $Msg" -ForegroundColor Green }
function Write-Warn { param([string]$Msg) Write-Host "[!] $Msg" -ForegroundColor Yellow }

Write-Step 'Enabling PowerShell Remoting (WinRM)'
Enable-PSRemoting -Force -SkipNetworkProfileCheck | Out-Null
Write-Ok 'PSRemoting enabled'

Write-Step 'Setting WinRM service to Automatic'
Set-Service -Name WinRM -StartupType Automatic
if ((Get-Service WinRM).Status -ne 'Running') { Start-Service WinRM }
Write-Ok 'WinRM running + Automatic'

Write-Step 'Ensuring firewall rule for WinRM HTTP (TCP 5985)'
$rule = Get-NetFirewallRule -Name 'WINRM-HTTP-In-TCP' -ErrorAction SilentlyContinue
if ($rule) {
    Set-NetFirewallRule -Name 'WINRM-HTTP-In-TCP' -Enabled True -Profile Domain,Private
    Write-Ok 'Firewall rule WINRM-HTTP-In-TCP enabled (Domain, Private)'
} else {
    Write-Warn 'WINRM-HTTP-In-TCP rule not found — Enable-PSRemoting should have created it'
}

if ($EnableHttps) {
    Write-Step 'Configuring HTTPS listener (TCP 5986)'
    $existing = Get-ChildItem WSMan:\localhost\Listener | Where-Object { $_.Keys -match 'Transport=HTTPS' }
    if ($existing) {
        Write-Ok 'HTTPS listener already present — skipping'
    } else {
        $cert = Get-ChildItem Cert:\LocalMachine\My |
            Where-Object { $_.EnhancedKeyUsageList.FriendlyName -contains 'Server Authentication' -and $_.NotAfter -gt (Get-Date) } |
            Sort-Object NotAfter -Descending | Select-Object -First 1
        if (-not $cert) {
            Write-Warn 'No Server Authentication certificate found in LocalMachine\My — HTTPS listener skipped'
        } else {
            $fqdn = "$env:COMPUTERNAME.$((Get-CimInstance Win32_ComputerSystem).Domain)"
            New-Item -Path WSMan:\localhost\Listener -Transport HTTPS -Address * -CertificateThumbPrint $cert.Thumbprint -HostName $fqdn -Force | Out-Null
            Write-Ok "HTTPS listener created (cert $($cert.Thumbprint.Substring(0,8))..., host $fqdn)"

            $httpsRule = Get-NetFirewallRule -DisplayName 'Windows Remote Management (HTTPS-In)' -ErrorAction SilentlyContinue
            if (-not $httpsRule) {
                New-NetFirewallRule -DisplayName 'Windows Remote Management (HTTPS-In)' `
                    -Direction Inbound -Protocol TCP -LocalPort 5986 -Action Allow -Profile Domain,Private | Out-Null
                Write-Ok 'Firewall rule for HTTPS (5986) created'
            }
        }
    }
}

if ($RestrictToGroup) {
    Write-Step "Restricting PSSessionConfiguration to group '$RestrictToGroup'"
    try {
        $grp = ([ADSISearcher]"(&(objectCategory=group)(sAMAccountName=$RestrictToGroup))").FindOne()
        if (-not $grp) { throw "AD group '$RestrictToGroup' not found" }

        $sid = (New-Object System.Security.Principal.NTAccount($RestrictToGroup)).Translate([System.Security.Principal.SecurityIdentifier]).Value
        $sddl = "O:NSG:BAD:P(A;;GA;;;BA)(A;;GA;;;$sid)S:P(AU;FA;GA;;;WD)(AU;SA;GXGW;;;WD)"

        Set-PSSessionConfiguration -Name Microsoft.PowerShell -SecurityDescriptorSddl $sddl -Force -WarningAction SilentlyContinue | Out-Null
        Write-Ok "PSSessionConfiguration restricted to BUILTIN\Administrators + $RestrictToGroup"
    } catch {
        Write-Warn "Could not apply group restriction: $_"
    }
}

Write-Host ''
Write-Ok 'Server is ready for EXpress remote documentation queries.'
Write-Host '    Test from the management host:'
Write-Host "    Test-WSMan -ComputerName $env:COMPUTERNAME"
Write-Host "    Get-CimInstance Win32_OperatingSystem -CimSession (New-CimSession -ComputerName $env:COMPUTERNAME)"
