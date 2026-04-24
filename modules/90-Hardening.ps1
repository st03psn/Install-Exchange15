    function Cleanup {
        Write-MyOutput "Cleaning up .."

        if ( (Get-WindowsFeature -Name 'Bits').Installed) {
            Write-MyOutput "Removing BITS feature"
            Remove-WindowsFeature Bits
        }
        Write-MyVerbose "Removing state file $Statefile"
        Remove-Item $Statefile
    }

    function Write-PhaseProgress {
        # Lightweight wrapper: Write-Progress for phase-level and step-level feedback.
        # Id 0 = overall install progress (Phase X of 6)
        # Id 1 = current-phase step progress (used in Phase 5 only)
        # PS2Exe does not render Write-Progress visually — fall back to Write-MyOutput for
        # meaningful milestones so progress is still visible in the console window.
        param(
            [int]$Id = 0,
            [string]$Activity,
            [string]$Status,
            [int]$PercentComplete = -1,
            [switch]$Completed
        )
        if ($Completed) {
            Write-Progress -Id $Id -Activity $Activity -Completed
        }
        elseif ($PercentComplete -ge 0) {
            Write-Progress -Id $Id -Activity $Activity -Status $Status -PercentComplete $PercentComplete
        }
        else {
            Write-Progress -Id $Id -Activity $Activity -Status $Status
        }

        # PS2Exe fallback: emit status as plain output so progress is not lost
        if ($IsPS2Exe -and -not $Completed -and $Status) {
            if ($Id -eq 0) {
                # Phase-level: only log when status changes (major transitions)
                Write-MyOutput ('[{0}] {1}' -f $Activity, $Status)
            }
            elseif ($Id -eq 1) {
                # Step-level (Phase 5): log every step
                Write-MyOutput ('  -> {0}' -f $Status)
            }
        }
    }

    function LockScreen {
        Write-MyVerbose 'Locking system'
        rundll32.exe user32.dll, LockWorkStation
    }

    function Clear-DesktopBackground {
        # Remove the desktop wallpaper during install — reduces visual distraction and
        # avoids Windows trying to render/cache wallpaper images while setup runs.
        # No restore needed: the server reboots multiple times during installation.
        # Uses registry + RUNDLL32 to avoid slow Add-Type/C# compilation on each phase start.
        Write-MyVerbose 'Clearing desktop background'
        Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name Wallpaper -Value '' -ErrorAction SilentlyContinue
        Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name WallpaperStyle -Value '0' -ErrorAction SilentlyContinue
        $p = Start-Process -FilePath 'RUNDLL32.EXE' -ArgumentList 'user32.dll, UpdatePerUserSystemParameters' -NoNewWindow -Wait -PassThru -ErrorAction SilentlyContinue
        if ($p -and $p.ExitCode -ne 0) {
            Write-MyWarning "RUNDLL32 UpdatePerUserSystemParameters exited with code $($p.ExitCode)"
        }
    }

    function Enable-HighPerformancePowerPlan {
        Write-MyVerbose 'Configuring Power Plan'
        $CurrentPlan = Get-CimInstance -Namespace root/cimv2/power -ClassName Win32_PowerPlan | Where-Object { $_.IsActive }
        if ($CurrentPlan.InstanceID -match $POWERPLAN_HIGH_PERFORMANCE) {
            Write-MyVerbose 'High Performance power plan already active'
        }
        else {
            $p = Start-Process -FilePath 'powercfg.exe' -ArgumentList ('/setactive', $POWERPLAN_HIGH_PERFORMANCE) -NoNewWindow -PassThru -Wait
            if ($p.ExitCode -ne 0) {
                Write-MyWarning "powercfg /setactive exited with code $($p.ExitCode)"
            }
            $CurrentPlan = Get-CimInstance -Namespace root/cimv2/power -ClassName Win32_PowerPlan | Where-Object { $_.IsActive }
            Write-MyOutput "Power Plan active: $($CurrentPlan.ElementName)"
        }
    }

    function Disable-NICPowerManagement {
        # http://support.microsoft.com/kb/2740020
        Write-MyVerbose 'Disabling Power Management on Network Adapters'
        # Find physical adapters that are OK and are not disabled
        $NICs = Get-CimInstance -ClassName Win32_NetworkAdapter | Where-Object { $_.AdapterTypeId -eq 0 -and $_.PhysicalAdapter -and $_.ConfigManagerErrorCode -eq 0 -and $_.ConfigManagerErrorCode -ne 22 }
        foreach ( $NIC in $NICs) {
            $PNPDeviceID = ($NIC.PNPDeviceID).ToUpper()
            $NICPowerMgt = Get-CimInstance -ClassName MSPower_DeviceEnable -Namespace root/wmi | Where-Object { $_.instancename -match [regex]::escape( $PNPDeviceID) }
            if ($NICPowerMgt.Enable) {
                Set-CimInstance -InputObject $NICPowerMgt -Property @{ Enable = $false }
                $NICPowerMgt = Get-CimInstance -ClassName MSPower_DeviceEnable -Namespace root/wmi | Where-Object { $_.instancename -match [regex]::escape( $PNPDeviceID) }
                if ($NICPowerMgt.Enable) {
                    Write-MyError "Problem disabling power management on $($NIC.Name) ($PNPDeviceID)"
                }
                else {
                    Write-MyOutput "Disabled power management on $($NIC.Name) ($PNPDeviceID)"
                }
            }
            else {
                Write-MyVerbose "Power management already disabled on $($NIC.Name) ($PNPDeviceID)"
            }
        }
    }

    function Set-Pagefile {
        Write-MyVerbose 'Checking Pagefile Configuration'
        $CS = Get-CimInstance -ClassName Win32_ComputerSystem
        if ($CS.AutomaticManagedPagefile) {
            Write-MyVerbose 'System configured to use Automatic Managed Pagefile, reconfiguring'
            try {
                $InstalledMem = $CS.TotalPhysicalMemory
                if ( $State["MajorSetupVersion"] -ge $EX2019_MAJOR) {
                    # 25% of RAM
                    $DesiredSize = [int]($InstalledMem / 4 / 1MB)
                    Write-MyVerbose ('Configuring PageFile to 25% of Total Memory: {0}MB' -f $DesiredSize)
                }
                else {
                    # RAM + 10 MB, with maximum of 32GB + 10MB
                    $DesiredSize = (($InstalledMem + 10MB), (32GB + 10MB) | Measure-Object -Minimum).Minimum / 1MB
                    Write-MyVerbose ('Configuring PageFile Total Memory+10MB with maximum of 32GB+10MB: {0}MB' -f $DesiredSize)
                }
                Set-CimInstance -InputObject $CS -Property @{ AutomaticManagedPagefile = $false }
                $CPF = Get-CimInstance -ClassName Win32_PageFileSetting
                Set-CimInstance -InputObject $CPF -Property @{ InitialSize = [int]$DesiredSize; MaximumSize = [int]$DesiredSize }
                Register-ExecutedCommand -Category 'Hardening' -Command 'Set-CimInstance -ClassName Win32_ComputerSystem -Property @{AutomaticManagedPagefile=$false}'
                Register-ExecutedCommand -Category 'Hardening' -Command ("Set-CimInstance -ClassName Win32_PageFileSetting -Property @{{InitialSize={0};MaximumSize={0}}}  # {0} MB" -f [int]$DesiredSize)
            }
            catch {
                Write-MyError "Problem reconfiguring pagefile: $($_.Exception.Message)"
            }
            $CPF = Get-CimInstance -ClassName Win32_PageFileSetting
            Write-MyOutput "Pagefile set to manual, initial/maximum size: $($CPF.InitialSize)MB / $($CPF.MaximumSize)MB"
        }
        else {
            Write-MyVerbose 'Manually configured page file, skipping configuration'
        }
    }

    function Set-TCPSettings {
        $currentRPC = (Get-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows NT\RPC' -Name 'MinimumConnectionTimeout' -ErrorAction SilentlyContinue).MinimumConnectionTimeout
        if ($currentRPC -eq 120) {
            Write-MyVerbose 'RPC Timeout already set to 120 seconds'
        }
        else {
            Write-MyOutput 'Setting RPC Timeout to 120 seconds'
            Set-RegistryValue -Path 'HKLM:\Software\Policies\Microsoft\Windows NT\RPC' -Name 'MinimumConnectionTimeout' -Value 120
        }
        $currentKA = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters' -Name 'KeepAliveTime' -ErrorAction SilentlyContinue).KeepAliveTime
        if ($currentKA -eq 900000) {
            Write-MyVerbose 'Keep-Alive Timeout already set to 15 minutes'
        }
        else {
            Write-MyOutput 'Setting Keep-Alive Timeout to 15 minutes'
            Set-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters' -Name 'KeepAliveTime' -Value 900000
        }
    }

    function Disable-SMBv1 {
        Write-MyOutput 'Disabling SMBv1 protocol (security best practice)'
        try {
            $feature = Get-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -ErrorAction SilentlyContinue
            if ($feature -and $feature.State -eq 'Enabled') {
                Disable-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -NoRestart -ErrorAction Stop | Out-Null
                Write-MyVerbose 'SMBv1 Windows feature disabled'
            }
            else {
                Write-MyVerbose 'SMBv1 Windows feature already disabled or not present'
            }
        }
        catch {
            Write-MyWarning ('Problem disabling SMBv1 feature: {0}' -f $_.Exception.Message)
        }
        try {
            Set-SmbServerConfiguration -EnableSMB1Protocol $false -Force -ErrorAction Stop
            Write-MyVerbose 'SMBv1 server protocol disabled'
        }
        catch {
            Write-MyWarning ('Problem disabling SMBv1 server config: {0}' -f $_.Exception.Message)
        }
    }

    function Disable-WindowsSearchService {
        Write-MyOutput 'Disabling Windows Search service (Exchange uses own content indexing)'
        $svc = Get-Service WSearch -ErrorAction SilentlyContinue
        if ($svc) {
            if ($svc.Status -eq 'Running') {
                Stop-Service WSearch -Force -ErrorAction SilentlyContinue
            }
            Set-Service WSearch -StartupType Disabled -ErrorAction SilentlyContinue
            Write-MyVerbose 'Windows Search service disabled'
        }
        else {
            Write-MyVerbose 'Windows Search service not found'
        }
    }

    function Disable-UnnecessaryServices {
        Write-MyOutput 'Disabling unnecessary Windows services (security hardening)'
        $services = @(
            @{ Name = 'Spooler';  Desc = 'Print Spooler (PrintNightmare attack surface, CVE-2021-34527)' }
            @{ Name = 'Fax';      Desc = 'Fax service (not required on Exchange)' }
            @{ Name = 'seclogon'; Desc = 'Secondary Logon (pass-the-hash / privilege escalation vector)' }
            @{ Name = 'SCardSvr'; Desc = 'Smart Card (not required on Exchange)' }
        )
        foreach ($svc in $services) {
            $s = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
            if ($s) {
                if ($s.Status -eq 'Running') {
                    Stop-Service -Name $svc.Name -Force -ErrorAction SilentlyContinue
                    Register-ExecutedCommand -Category 'Hardening' -Command ('Stop-Service -Name {0} -Force' -f $svc.Name)
                }
                Set-Service -Name $svc.Name -StartupType Disabled -ErrorAction SilentlyContinue
                Register-ExecutedCommand -Category 'Hardening' -Command ('Set-Service -Name {0} -StartupType Disabled  # {1}' -f $svc.Name, $svc.Desc)
                Write-MyVerbose ('Disabled: {0}' -f $svc.Desc)
            }
            else {
                Write-MyVerbose ('Service not found, skipping: {0}' -f $svc.Name)
            }
        }
    }

    function Disable-ShutdownEventTracker {
        # Redundant with Event IDs 1074/6006/6008; dialog blocks unattended Autopilot reboots
        Write-MyOutput 'Disabling Shutdown Event Tracker (redundant with event log; blocks unattended reboots)'
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Reliability' -Name 'ShutdownReasonOn' -Value 0
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Reliability' -Name 'ShutdownReasonUI' -Value 0
    }

    function Disable-WDigestCredentialCaching {
        # Prevents cleartext credential storage in LSASS memory (Mimikatz mitigation)
        Write-MyOutput 'Disabling WDigest credential caching (security hardening)'
        Set-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest' -Name 'UseLogonCredential' -Value 0
    }

    function Disable-HTTP2 {
        # HTTP/2 causes known issues with Exchange MAPI/RPC connections
        Write-MyOutput 'Disabling HTTP/2 protocol (Exchange compatibility)'
        Set-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\HTTP\Parameters' -Name 'EnableHttp2Tls' -Value 0
        Set-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\HTTP\Parameters' -Name 'EnableHttp2Cleartext' -Value 0
    }

    function Disable-TCPOffload {
        # Microsoft recommends disabling TCP offload features on Exchange servers.
        # chimney=disabled was removed from netsh in WS2019 — only apply on WS2016 (build < 17763).
        Write-MyOutput 'Disabling TCP Task Offload and autotuning settings'
        try {
            $osBuild = [int](Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name CurrentBuildNumber -ErrorAction SilentlyContinue).CurrentBuildNumber
            if ($osBuild -gt 0 -and $osBuild -lt 17763) {
                Invoke-NativeCommand -FilePath 'netsh.exe' -Arguments @('int','tcp','set','global','chimney=disabled') -Tag 'netsh chimney' | Out-Null
                if ($LASTEXITCODE -ne 0) { Write-MyWarning ('netsh chimney=disabled exited with code {0}' -f $LASTEXITCODE) }
            }
            Invoke-NativeCommand -FilePath 'netsh.exe' -Arguments @('int','tcp','set','global','autotuninglevel=restricted') -Tag 'netsh autotuninglevel' | Out-Null
            if ($LASTEXITCODE -ne 0) { Write-MyWarning ('netsh autotuninglevel=restricted exited with code {0}' -f $LASTEXITCODE) }
            Set-NetOffloadGlobalSetting -TaskOffload Disabled -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            Write-MyVerbose 'TCP offload settings configured'
        }
        catch {
            Write-MyWarning ('Problem configuring TCP offload: {0}' -f $_.Exception.Message)
        }
    }

    function Test-DiskAllocationUnitSize {
        # Exchange best practice: database and log volumes should use 64KB allocation units
        Write-MyOutput 'Checking disk allocation unit sizes (64KB recommended for Exchange volumes)'
        Get-Volume | Where-Object { $_.DriveLetter -and $_.FileSystem -eq 'NTFS' } | ForEach-Object {
            $letter = $_.DriveLetter
            $auSize = $_.AllocationUnitSize
            if ($auSize -and $auSize -ne 65536) {
                Write-MyWarning ('Drive {0}: uses {1} byte allocation units (64KB/65536 recommended for Exchange database/log volumes)' -f $letter, $auSize)
            }
            else {
                Write-MyVerbose ('Drive {0}: allocation unit size OK ({1})' -f $letter, $auSize)
            }
        }
    }

    function Disable-UnnecessaryScheduledTasks {
        Write-MyOutput 'Disabling unnecessary scheduled tasks (performance optimization)'
        $tasksToDisable = @(
            '\Microsoft\Windows\Defrag\ScheduledDefrag'
        )
        foreach ($taskName in $tasksToDisable) {
            try {
                $task = Get-ScheduledTask -TaskName (Split-Path $taskName -Leaf) -TaskPath ((Split-Path $taskName -Parent) + '\') -ErrorAction SilentlyContinue
                if ($task -and $task.State -ne 'Disabled') {
                    $task | Disable-ScheduledTask | Out-Null
                    Write-MyVerbose ('Disabled scheduled task: {0}' -f $taskName)
                }
                else {
                    Write-MyVerbose ('Scheduled task already disabled or not found: {0}' -f $taskName)
                }
            }
            catch {
                Write-MyWarning ('Problem disabling scheduled task {0}: {1}' -f $taskName, $_.Exception.Message)
            }
        }
    }

    function Disable-ServerManagerAtLogon {
        # Disable Server Manager at logon for ALL users (machine-wide).
        # Three layers are used for complete coverage:
        #   1. Machine-wide Group Policy key — overrides per-user HKCU settings
        #   2. Default user hive — applies to new user profiles created after this point
        #   3. Scheduled task — belt-and-suspenders, prevents task-triggered launch
        # Idempotent: silent if all three layers are already configured.
        $policyPath    = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Server\ServerManager'
        $alreadyPolicy = (Get-ItemProperty -Path $policyPath -Name 'DoNotOpenAtLogon' -ErrorAction SilentlyContinue).DoNotOpenAtLogon -eq 1
        $smTask        = Get-ScheduledTask -TaskName 'ServerManager' -TaskPath '\Microsoft\Windows\Server Manager\' -ErrorAction SilentlyContinue
        $alreadyTask   = -not $smTask -or $smTask.State -eq 'Disabled'
        if ($alreadyPolicy -and $alreadyTask) {
            Write-MyVerbose 'Server Manager at logon already disabled — skipping'
            return
        }
        Write-MyOutput 'Disabling Server Manager at logon for all users'

        # Layer 1: Machine-wide policy (overrides HKCU for all users)
        $policyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Server\ServerManager'
        if (-not (Test-Path $policyPath -ErrorAction SilentlyContinue)) {
            New-Item -Path $policyPath -Force -ErrorAction SilentlyContinue | Out-Null
        }
        Set-RegistryValue -Path $policyPath -Name 'DoNotOpenAtLogon' -Value 1 -PropertyType DWord

        # Layer 2: Default user profile hive (new users created after this point)
        $defaultHive    = 'C:\Users\Default\NTUSER.DAT'
        $defaultHiveKey = 'HKU\ExchangeInstallDefault'
        if (Test-Path $defaultHive) {
            Invoke-NativeCommand -FilePath 'reg.exe' -Arguments @('load', $defaultHiveKey, $defaultHive) -Tag 'reg load default hive' | Out-Null
            if (Test-Path "Registry::$defaultHiveKey\Software\Microsoft\ServerManager") {
                Set-ItemProperty -Path "Registry::$defaultHiveKey\Software\Microsoft\ServerManager" -Name 'DoNotOpenServerManagerAtLogon' -Value 1 -Type DWord -ErrorAction SilentlyContinue
            }
            else {
                New-Item -Path "Registry::$defaultHiveKey\Software\Microsoft\ServerManager" -Force -ErrorAction SilentlyContinue | Out-Null
                New-ItemProperty -Path "Registry::$defaultHiveKey\Software\Microsoft\ServerManager" -Name 'DoNotOpenServerManagerAtLogon' -Value 1 -PropertyType DWord -Force -ErrorAction SilentlyContinue | Out-Null
            }
            Invoke-NativeCommand -FilePath 'reg.exe' -Arguments @('unload', $defaultHiveKey) -Tag 'reg unload default hive' | Out-Null
        }

        # Layer 3: Disable the ServerManager scheduled task (machine-wide)
        $smTask = Get-ScheduledTask -TaskName 'ServerManager' -TaskPath '\Microsoft\Windows\Server Manager\' -ErrorAction SilentlyContinue
        if ($smTask -and $smTask.State -ne 'Disabled') {
            $smTask | Disable-ScheduledTask | Out-Null
            Write-MyVerbose 'Disabled scheduled task: \Microsoft\Windows\Server Manager\ServerManager'
        }
    }

    function Set-CRLCheckTimeout {
        # Prevents Exchange startup delays when CRL endpoints are unreachable
        Write-MyOutput 'Configuring Certificate Revocation List check timeout (15 seconds)'
        $regPath = 'HKLM:\SOFTWARE\Microsoft\Cryptography\OID\EncodingType 0\CertDllCreateCertificateChainEngine\Config'
        if (-not (Test-Path $regPath -ErrorAction SilentlyContinue)) {
            New-Item -Path $regPath -Force -ErrorAction SilentlyContinue | Out-Null
        }
        Set-RegistryValue -Path $regPath -Name 'ChainUrlRetrievalTimeoutMilliseconds' -Value 15000
    }

    function Disable-CredentialGuard {
        # HealthChecker flags Credential Guard as causing performance issues on Exchange servers.
        # On Windows Server 2025 it is enabled by default and must be explicitly disabled.
        Write-MyOutput 'Disabling Credential Guard (Exchange performance best practice)'
        Set-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA' -Name 'LsaCfgFlags' -Value 0
        Set-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard' -Name 'EnableVirtualizationBasedSecurity' -Value 0
    }

    function Set-LmCompatibilityLevel {
        # HealthChecker recommends level 5: send NTLMv2 only, refuse LM and NTLM
        Write-MyOutput 'Setting LAN Manager compatibility level to 5 (NTLMv2 only)'
        Set-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' -Name 'LmCompatibilityLevel' -Value 5
    }

    function Enable-RSSOnAllNICs {
        # HealthChecker warns if RSS is disabled or queue count does not match physical core count
        Write-MyOutput 'Enabling Receive Side Scaling (RSS) on all supported NICs'
        $physicalCores = (Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue |
            Measure-Object -Property NumberOfCores -Sum).Sum
        if (-not $physicalCores -or $physicalCores -lt 1) { $physicalCores = 1 }
        Write-MyVerbose ('Physical core count: {0} — setting RSS queue count to match' -f $physicalCores)
        Register-ExecutedCommand -Category 'Hardening' -Command 'Enable-NetAdapterRss -Name *'
        Register-ExecutedCommand -Category 'Hardening' -Command ("Set-NetAdapterRss -Name * -NumberOfReceiveQueues $physicalCores  # = physical core count")
        try {
            Get-NetAdapterRss -ErrorAction SilentlyContinue | ForEach-Object {
                if (-not $_.Enabled) {
                    Write-MyVerbose ('Enabling RSS on adapter: {0}' -f $_.Name)
                    Enable-NetAdapterRss -Name $_.Name -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                }
                if ($_.NumberOfReceiveQueues -ne $physicalCores) {
                    Set-NetAdapterRss -Name $_.Name -NumberOfReceiveQueues $physicalCores -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                    Write-MyVerbose ('Set RSS queues to {0} on adapter: {1}' -f $physicalCores, $_.Name)
                }
                else {
                    Write-MyVerbose ('RSS queues already at {0} on adapter: {1}' -f $physicalCores, $_.Name)
                }
            }
        }
        catch {
            Write-MyWarning ('Problem configuring RSS: {0}' -f $_.Exception.Message)
        }
    }

    function Set-IPv4OverIPv6Preference {
        # Microsoft recommendation for Exchange: prefer IPv4 over IPv6 (DisabledComponents = 0x20).
        # Disables IPv6 on all non-loopback interfaces but keeps the IPv6 loopback intact,
        # which Exchange internal components rely on. Full IPv6 disable (0xFF) is not recommended.
        $regPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters'
        $current = (Get-ItemProperty -Path $regPath -Name DisabledComponents -ErrorAction SilentlyContinue).DisabledComponents
        if ($current -eq 0x20) {
            Write-MyVerbose 'IPv4 over IPv6 preference already set (DisabledComponents = 0x20) (OK)'
        } else {
            Set-RegistryValue -Path $regPath -Name 'DisabledComponents' -Value 0x20 -PropertyType DWord
            Write-MyOutput 'IPv4 over IPv6 preference set (DisabledComponents = 0x20) — effective after next reboot'
            # Do not flag RebootRequired: the value is re-read at boot and the install's
            # end-of-Phase-6 reboot (or the next natural reboot) activates it. Forcing a
            # mid-install reboot here would trigger the Phase 5→6 skip-logic unnecessarily.
        }
    }

    function Disable-NetBIOSOnAllNICs {
        # Disables NetBIOS over TCP/IP on all NICs. Exchange does not require NetBIOS;
        # disabling it reduces attack surface (LLMNR/NBT-NS poisoning, credential capture).
        # SetTcpipNetbios(2) = Disable NetBIOS over TCP/IP
        Write-MyOutput 'Disabling NetBIOS over TCP/IP on all NICs'
        try {
            $nics = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True' -ErrorAction Stop
            $changed = 0
            foreach ($nic in $nics) {
                $result = ($nic | Invoke-CimMethod -MethodName SetTcpipNetbios -Arguments @{ TcpipNetbiosOptions = [uint32]2 } -ErrorAction SilentlyContinue).ReturnValue
                if ($result -eq 0) {
                    Write-MyVerbose ('NetBIOS disabled on: {0}' -f $nic.Description)
                    $changed++
                } elseif ($result -eq 1) {
                    Write-MyVerbose ('NetBIOS disable on {0}: takes effect after next reboot' -f $nic.Description)
                    $changed++
                    # Do not flag RebootRequired: the setting activates on the next boot
                    # anyway (end-of-Phase-6 reboot covers it). Forcing a Phase 5→6 reboot
                    # for a NIC flag is unnecessary.
                } else {
                    Write-MyWarning ('NetBIOS disable on {0} returned code {1}' -f $nic.Description, $result)
                }
            }
            Write-MyVerbose ('NetBIOS disabled on {0} NIC(s)' -f $changed)
        } catch {
            Write-MyWarning ('Failed to disable NetBIOS: {0}' -f $_.Exception.Message)
        }
    }

    function Disable-LLMNR {
        # CIS L1 18.5.4.2: Disable Link-Local Multicast Name Resolution.
        # LLMNR broadcasts unresolved names to the local subnet; Responder-class tools
        # answer with spoofed records and capture NTLM hashes. Exchange relies on DNS.
        Write-MyOutput 'Disabling LLMNR (Link-Local Multicast Name Resolution)'
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient' -Name 'EnableMulticast' -Value 0
    }

    function Disable-MDNS {
        # WS2022+ enables mDNS by default (port 5353 UDP). Same poisoning vector as LLMNR.
        # Registry value EnableMDNS under Dnscache\Parameters disables it globally.
        Write-MyOutput 'Disabling mDNS (Multicast DNS)'
        Set-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\Dnscache\Parameters' -Name 'EnableMDNS' -Value 0
    }

    function Enable-LSAProtection {
        # Enables LSA Protection (RunAsPPL) to prevent credential theft from LSASS memory.
        # Exchange 2019 CU12+ and Exchange SE are compatible with LSA Protection.
        # Earlier Exchange versions (2016, pre-CU12 2019) may conflict with legacy auth providers.
        # The setting takes effect after the next reboot.
        $regPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'
        $current = (Get-ItemProperty -Path $regPath -Name RunAsPPL -ErrorAction SilentlyContinue).RunAsPPL
        if ($current -eq 1) {
            Write-MyVerbose 'LSA Protection (RunAsPPL) already enabled'
            return
        }
        Write-MyOutput 'Enabling LSA Protection (RunAsPPL) — effective after next reboot'
        Set-RegistryValue -Path $regPath -Name 'RunAsPPL' -Value 1 -PropertyType DWord
        # Audit mode first (2) is not used here as Exchange servers are domain-joined production systems
        # and Exchange 2019 CU12+/SE are fully compatible with RunAsPPL = 1.
    }

    function Set-MaxConcurrentAPI {
        # Netlogon MaxConcurrentApi limits simultaneous Kerberos/NTLM authentication requests
        # against domain controllers. Exchange generates heavy auth load; the default (10) can
        # cause 0xC000005E (No logon servers) errors under load on busy servers.
        # Microsoft recommendation for Exchange: raise to match logical processor count (min 10).
        # Edge Transport is not domain-joined — Netlogon optimization does not apply.
        if ($State['InstallEdge']) { Write-MyVerbose 'Set-MaxConcurrentAPI: skipped (Edge Transport)'; return }
        Write-MyOutput 'Setting Netlogon MaxConcurrentApi for Kerberos authentication optimization'
        $logicalProcs = (Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue).NumberOfLogicalProcessors
        if (-not $logicalProcs -or $logicalProcs -lt 10) { $logicalProcs = 10 }
        $regPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters'
        Register-ExecutedCommand -Category 'Hardening' -Command ("Set-ItemProperty '$regPath' MaxConcurrentApi $logicalProcs  # = logical processor count, min 10")
        Set-RegistryValue -Path $regPath -Name 'MaxConcurrentApi' -Value $logicalProcs -PropertyType DWord
        Write-MyVerbose ('MaxConcurrentApi set to {0}' -f $logicalProcs)
    }

    function Set-CtsProcessorAffinityPercentage {
        # HealthChecker flags any non-zero value as harmful to Exchange Search performance
        # Edge Transport uses a different search stack — this registry path does not exist there.
        if ($State['InstallEdge']) { Write-MyVerbose 'Set-CtsProcessorAffinityPercentage: skipped (Edge Transport)'; return }
        Write-MyOutput 'Setting CtsProcessorAffinityPercentage to 0 (Exchange Search best practice)'
        $regPath = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Search\SystemParameters'
        if (-not (Test-Path $regPath -ErrorAction SilentlyContinue)) {
            New-Item -Path $regPath -Force -ErrorAction SilentlyContinue | Out-Null
        }
        Set-RegistryValue -Path $regPath -Name 'CtsProcessorAffinityPercentage' -Value 0
    }

    function Enable-SerializedDataSigning {
        # HealthChecker validates this security feature (mitigates PowerShell serialization attacks)
        Write-MyOutput 'Enabling Serialized Data Signing (security hardening)'
        $regPath = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Diagnostics'
        if (-not (Test-Path $regPath -ErrorAction SilentlyContinue)) {
            New-Item -Path $regPath -Force -ErrorAction SilentlyContinue | Out-Null
        }
        Set-RegistryValue -Path $regPath -Name 'EnableSerializationDataSigning' -Value 1
    }

    function Set-NodeRunnerMemoryLimit {
        # HealthChecker flags any non-zero memoryLimitMegabytes as a Search performance limiter
        # Edge Transport does not run Exchange Search / NodeRunner.
        if ($State['InstallEdge']) { Write-MyVerbose 'Set-NodeRunnerMemoryLimit: skipped (Edge Transport)'; return }
        Write-MyOutput 'Setting NodeRunner memory limit to 0 (unlimited, Exchange Search best practice)'
        $exchangeInstallPath = (Get-ItemProperty -Path $EXCHANGEINSTALLKEY -Name MsiInstallPath -ErrorAction SilentlyContinue).MsiInstallPath
        if ($exchangeInstallPath) {
            $configFile = Join-Path $exchangeInstallPath 'Bin\Search\Ceres\Runtime\1.0\noderunner.exe.config'
            if (Test-Path $configFile) {
                try {
                    $xml = [XML](Get-Content $configFile)
                    $node = $xml.SelectSingleNode('//nodeRunnerSettings')
                    if ($node -and $node.memoryLimitMegabytes -ne '0') {
                        $node.memoryLimitMegabytes = '0'
                        $xml.Save($configFile)
                        Write-MyVerbose 'NodeRunner memoryLimitMegabytes set to 0'
                    }
                    else {
                        Write-MyVerbose 'NodeRunner memoryLimitMegabytes already 0 or node not found'
                    }
                }
                catch {
                    Write-MyWarning ('Problem configuring NodeRunner: {0}' -f $_.Exception.Message)
                }
            }
            else {
                Write-MyVerbose 'NodeRunner config file not found (may not be installed yet)'
            }
        }
    }

    function Enable-MAPIFrontEndServerGC {
        # HealthChecker recommends Server GC for MAPI Front End App Pool on systems with 20+ GB RAM
        Write-MyOutput 'Checking MAPI Front End App Pool GC mode'
        $installedMem = (Get-CimInstance -ClassName Win32_ComputerSystem).TotalPhysicalMemory
        if ($installedMem -ge 20GB) {
            $exchangeInstallPath = (Get-ItemProperty -Path $EXCHANGEINSTALLKEY -Name MsiInstallPath -ErrorAction SilentlyContinue).MsiInstallPath
            if ($exchangeInstallPath) {
                $configFile = Join-Path $exchangeInstallPath 'bin\MSExchangeMapiFrontEndAppPool_CLRConfig.config'
                if (Test-Path $configFile) {
                    try {
                        $xml = [XML](Get-Content $configFile)
                        $gcNode = $xml.SelectSingleNode('//gcServer')
                        if ($gcNode -and $gcNode.enabled -ne 'true') {
                            $gcNode.enabled = 'true'
                            $xml.Save($configFile)
                            Write-MyOutput 'Enabled Server GC for MAPI Front End App Pool (20+ GB RAM detected)'
                        }
                        else {
                            Write-MyVerbose 'Server GC already enabled or node not found'
                        }
                    }
                    catch {
                        Write-MyWarning ('Problem configuring MAPI FE GC: {0}' -f $_.Exception.Message)
                    }
                }
                else {
                    Write-MyVerbose 'MAPI FE config file not found (may not be installed yet)'
                }
            }
        }
        else {
            Write-MyVerbose 'Less than 20 GB RAM, skipping Server GC configuration'
        }
    }

    function Disable-SSL3 {
        # SSL3 disabling/Poodle, https://support.microsoft.com/en-us/kb/187498
        Write-MyVerbose 'Disabling SSL3 protocol for services'
        Set-RegistryValue -Path 'HKLM:\System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server' -Name 'Enabled' -Value 0
    }

    function Disable-RC4 {
        # https://support.microsoft.com/en-us/kb/2868725
        # Note: Can't use regular New-Item as registry path contains '/' (always interpreted as path splitter)
        Write-MyVerbose 'Disabling RC4 protocol for services'
        $RC4Keys = @('RC4 128/128', 'RC4 40/128', 'RC4 56/128')
        $RegKey = 'SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers'
        $RegName = "Enabled"
        foreach ( $RC4Key in $RC4Keys) {
            if ( -not( Get-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue)) {
                if ( -not (Test-Path $RegKey -ErrorAction SilentlyContinue)) {
                    $RegHandle = (Get-Item 'HKLM:\').OpenSubKey( $RegKey, $true)
                    $RegHandle.CreateSubKey( $RC4Key) | Out-Null
                    $RegHandle.Close()
                }
            }
            Write-MyVerbose "Setting registry $RegKey\$RegName\RC4Key to 0"
            New-ItemProperty -Path (Join-Path (Join-Path 'HKLM:\' $RegKey) $RC4Key) -Name $RegName -Value 0 -Force -ErrorAction SilentlyContinue | Out-Null
            Register-ExecutedCommand -Category 'Hardening' -Command ("New-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\{0}' -Name Enabled -Value 0 -Force" -f $RC4Key)
        }
    }

    function Enable-ECC {
        # https://learn.microsoft.com/en-us/exchange/architecture/client-access/certificates?view=exchserver-2019#elliptic-curve-cryptography-certificates-support-in-exchange-server
        Write-MyOutput 'Enabling Elliptic Curve Cryptography (ECC) certificate support'

        $RegKey = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Diagnostics'
        $RegName = 'EnableEccCertificateSupport'

        if ( -not( Get-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue)) {
            Write-MyVerbose ('Setting {0}\{1} to 1' -f $RegKey, $RegName)
            New-ItemProperty -Path $RegKey -Name $RegName -Value 1 -Type String -Force -ErrorAction SilentlyContinue | Out-Null
        }

        # If overrides were configured, disable these (obsolete and not fully supporting ECC)
        $Override = Get-SettingOverride | Where-Object { ($_.SectionName -eq "ECCCertificateSupport") -and ($_.Parameters -eq "Enabled=true") }
        if ( $Override) {
            Write-MyVerbose ('Override for ECC found, removing (obsolete)')
            $Override | Remove-SettingOverride
            Get-ExchangeDiagnosticInfo -Process Microsoft.Exchange.Directory.TopologyService -Component VariantConfiguration -Argument Refresh | Out-Null
            $script:p5NeedsIisRestart = $true
        }
        else {
            Write-MyVerbose ('No override configuration for ECC found')
        }
    }

    function Enable-CBC {
        # https://support.microsoft.com/en-us/topic/enable-support-for-aes256-cbc-encrypted-content-in-exchange-server-august-2023-su-add63652-ee17-4428-8928-ddc45339f99e
        Write-MyOutput 'Enabling AES256-CBC encryption mode support'

        $Override = Get-SettingOverride | Where-Object { ($_.SectionName -eq "EnableEncryptionAlgorithmCBC") -and ($_.Parameters -eq "Enabled=True") }
        if ( $Override) {
            Write-MyVerbose ('Configuration for CBC already configured')
        }
        else {
            New-SettingOverride -Name "EnableEncryptionAlgorithmCBC" -Parameters @("Enabled=True") -Component Encryption -Reason "Enable CBC encryption" -Section EnableEncryptionAlgorithmCBC | Out-Null
            Get-ExchangeDiagnosticInfo -Process Microsoft.Exchange.Directory.TopologyService -Component VariantConfiguration -Argument Refresh | Out-Null
            $script:p5NeedsIisRestart = $true
        }
    }

    function Enable-AMSI {
        param(
            [string[]]$ConfigParam = @("EnabledEcp=True", "EnabledEws=True", "EnabledOwa=True", "EnabledPowerShell=True")
        )
        # https://learn.microsoft.com/en-us/exchange/antispam-and-antimalware/amsi-integration-with-exchange?view=exchserver-2019#enable-exchange-server-amsi-body-scanning
        # Edge Transport is not domain-joined and has no org connection; New-SettingOverride would fail.
        if ($State['InstallEdge']) { Write-MyVerbose 'Enable-AMSI: skipped (Edge Transport — no org connection)'; return }
        Write-MyOutput 'Enabling AMSI body scanning for OWA, ECP, EWS and PowerShell'

        $amsiOverride = Get-SettingOverride | Where-Object { $_.SectionName -eq 'AmsiRequestBodyScanning' }
        if ($amsiOverride) {
            Write-MyVerbose 'AMSI body scanning override already configured'
        }
        else {
            New-SettingOverride -Name "EnableAMSIBodyScan" -Component Cafe -Section AmsiRequestBodyScanning -Parameters $ConfigParam -Reason "Enabling AMSI body Scan" | Out-Null
            Get-ExchangeDiagnosticInfo -Process Microsoft.Exchange.Directory.TopologyService -Component VariantConfiguration -Argument Refresh | Out-Null
            $script:p5NeedsIisRestart = $true
        }
    }

    function Enable-IanaTimeZoneMappings {
        # Exchange 2019 CU14+ ships IanaTimeZoneMappings.xml in the bin folder.
        # HealthChecker flags its absence as a calendar timezone issue.
        $setupKey = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup' -ErrorAction SilentlyContinue
        if (-not $setupKey) { Write-MyVerbose 'Enable-IanaTimeZoneMappings: Exchange setup registry key not found'; return }
        $exBin = Join-Path $setupKey.MsiInstallPath 'Bin'
        $mappingFile = Join-Path $exBin 'IanaTimeZoneMappings.xml'
        if (Test-Path $mappingFile) {
            Write-MyVerbose ('IANA timezone mappings file present: {0}' -f $mappingFile)
        }
        else {
            Write-MyWarning ('IANA timezone mappings file not found ({0}). Calendar timezone issues may occur. Update Exchange to a newer CU to resolve.' -f $mappingFile)
        }

        # Exchange 2019 CU14+ / SE: enable IANA timezone IDs for calendar items if supported
        try {
            $orgConfig = Get-OrganizationConfig -ErrorAction Stop
            if ($orgConfig.PSObject.Properties['UseIanaTimeZoneId']) {
                if (-not $orgConfig.UseIanaTimeZoneId) {
                    Set-OrganizationConfig -UseIanaTimeZoneId $true -ErrorAction Stop
                    Write-MyOutput 'IANA timezone IDs enabled for calendar items (UseIanaTimeZoneId)'
                }
                else {
                    Write-MyVerbose 'IANA timezone IDs already enabled (UseIanaTimeZoneId)'
                }
            }
            else {
                Write-MyVerbose 'UseIanaTimeZoneId not available on this Exchange version — skipping'
            }
        }
        catch {
            Write-MyVerbose ('Enable-IanaTimeZoneMappings: {0}' -f $_.Exception.Message)
        }
    }

    function Disable-SSLOffloading {
        # F13: SSL offloading at a reverse proxy prevents Extended Protection channel-binding from working.
        # Always set to $false — Exchange should terminate TLS itself, not receive plaintext from a proxy.
        if ($State['InstallEdge']) { return }
        Write-MyOutput 'Configuring Outlook Anywhere SSL offloading (required for Extended Protection)'
        try {
            $oa = Get-OutlookAnywhere -Server $env:computername -ErrorAction SilentlyContinue
            if ($oa) {
                if ($oa.SSLOffloading) {
                    Set-OutlookAnywhere -Identity $oa.Identity -SSLOffloading $false -Confirm:$false -ErrorAction Stop
                    Register-ExecutedCommand -Category 'ExchangeTuning' -Command ("Set-OutlookAnywhere -Identity '{0}' -SSLOffloading `$false" -f $oa.Identity)
                    Write-MyVerbose 'Outlook Anywhere SSL offloading disabled'
                }
                else {
                    Write-MyVerbose 'Outlook Anywhere SSL offloading already disabled (OK)'
                }
            }
            else {
                Write-MyVerbose 'No Outlook Anywhere virtual directory found on this server'
            }
        }
        catch {
            Write-MyWarning ('Could not configure Outlook Anywhere SSL offloading: {0}' -f $_.Exception.Message)
        }
    }

    function Enable-ExtendedProtection {
        # F6: Windows Extended Protection (channel binding) mitigates NTLM relay / pass-the-hash attacks on IIS.
        # Prerequisite: SSL offloading must be disabled (F13), TLS 1.2 must be enforced.
        # Exchange 2019 CU14+ / SE: EP is enabled by setup — this function validates the configuration.
        # Exchange 2016 / 2019 pre-CU14: downloads and runs ExchangeExtendedProtectionManagement.ps1 from CSS-Exchange.
        if ($State['DoNotEnableEP']) { Write-MyVerbose 'DoNotEnableEP set — skipping Extended Protection'; return }
        if ($State['InstallEdge'])   { Write-MyVerbose 'Edge Transport — Extended Protection not applicable'; return }

        $exSetupVer    = [System.Version]$State['ExSetupVersion']
        $isCU14OrNewer = $exSetupVer -ge [System.Version]$EX2019SETUPEXE_CU14

        if ($isCU14OrNewer) {
            Write-MyOutput 'Exchange 2019 CU14+ / SE — Extended Protection enabled by setup; validating OWA'
            try {
                $owa = Get-OwaVirtualDirectory -Server $env:computername -ErrorAction SilentlyContinue
                if ($owa) {
                    $ep = $owa.ExtendedProtectionTokenChecking
                    if ($ep -eq 'None') {
                        Write-MyWarning ('OWA Extended Protection is None (expected Require/Allow for Exchange {0}). Review ExtendedProtectionTokenChecking on all virtual directories.' -f $State['ExSetupVersion'])
                    }
                    else {
                        Write-MyVerbose ('OWA ExtendedProtectionTokenChecking: {0} (OK)' -f $ep)
                    }
                }
            }
            catch { Write-MyVerbose ('Extended Protection validation: {0}' -f $_.Exception.Message) }
            return
        }

        # Exchange 2016 / 2019 pre-CU14: configure via CSS-Exchange ExchangeExtendedProtectionManagement.ps1
        Write-MyOutput 'Enabling Extended Protection via CSS-Exchange ExchangeExtendedProtectionManagement.ps1'
        $epPath = Join-Path $State['SourcesPath'] 'ExchangeExtendedProtectionManagement.ps1'
        $epUrl  = 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/ExchangeExtendedProtectionManagement.ps1'
        # Note: previously named ExchangeExtendedProtection.ps1 — renamed in CSS-Exchange 2024 releases

        if (-not (Test-Path $epPath)) {
            try {
                Invoke-WebDownload -Uri $epUrl -OutFile $epPath
                Write-MyVerbose ('ExchangeExtendedProtectionManagement.ps1 downloaded, SHA256: {0}' -f (Get-FileHash $epPath -Algorithm SHA256).Hash)
            }
            catch {
                Write-MyWarning ('Could not download ExchangeExtendedProtectionManagement.ps1: {0}' -f $_.Exception.Message)
                return
            }
        }

        try {
            $epArgs    = @('-ExchangeServerNames', $env:computername)
            $epSkipEWS = if ($State['DoNotEnableEP_FEEWS']) { ' -SkipEWS' } else { '' }
            if ($epSkipEWS) { $epArgs += '-SkipEWS' }
            $epCmd = '& ExchangeExtendedProtectionManagement.ps1 -ExchangeServerNames {0}{1}' -f $env:computername, $epSkipEWS
            Register-ExecutedCommand -Category 'ExchangeTuning' -Command $epCmd
            & $epPath @epArgs *>&1 | ForEach-Object { Write-ToTranscript ([string]$_) }
        }
        catch {
            Write-MyWarning ('ExchangeExtendedProtectionManagement.ps1 failed: {0}' -f $_.Exception.Message)
        }
    }

    function Enable-RootCertificateAutoUpdate {
        # F17: Prevents Exchange Online connectivity failures caused by stale/missing root CA certificates.
        # Group Policy or hardening baselines sometimes disable Windows automatic root certificate updates,
        # which breaks connectivity to Exchange Online, Microsoft 365, and any modern PKI-dependent service.
        Write-MyOutput 'Verifying automatic root certificate update (AuthRoot policy)'
        $regPath = 'HKLM:\SOFTWARE\Policies\Microsoft\SystemCertificates\AuthRoot'
        try {
            $val = (Get-ItemProperty -Path $regPath -Name DisableRootAutoUpdate -ErrorAction SilentlyContinue).DisableRootAutoUpdate
            if ($val -eq 1) {
                Set-RegistryValue -Path $regPath -Name 'DisableRootAutoUpdate' -Value 0 -PropertyType DWord
                Write-MyOutput 'Root certificate auto-update re-enabled (was disabled by policy — required for Exchange Online / M365 connectivity)'
            }
            else {
                Write-MyVerbose 'Root certificate auto-update: not disabled by policy (OK)'
            }
        }
        catch {
            Write-MyVerbose ('Root certificate auto-update check: {0}' -f $_.Exception.Message)
        }
    }

    function Disable-MRSProxy {
        # F18: MRS Proxy enables cross-forest / cross-org mailbox moves. Disable when not in use —
        # HealthChecker flags an enabled MRS Proxy endpoint as unnecessary attack surface.
        # Re-enable with: Set-WebServicesVirtualDirectory -MRSProxyEnabled $true -Confirm:$false
        if (-not $State['InstallMailbox']) { return }
        Write-MyOutput 'Disabling MRS Proxy on EWS virtual directory (enable manually for cross-forest migrations)'
        try {
            Get-WebServicesVirtualDirectory -Server $env:computername -ErrorAction Stop |
                Set-WebServicesVirtualDirectory -MRSProxyEnabled $false -Confirm:$false -ErrorAction Stop
            Write-MyVerbose 'MRS Proxy disabled (Set-WebServicesVirtualDirectory -MRSProxyEnabled $false)'
        }
        catch {
            Write-MyWarning ('Could not disable MRS Proxy: {0}' -f $_.Exception.Message)
        }
    }

    function Set-MAPIEncryptionRequired {
        # F19: Requires MAPI encryption on all MAPI-over-RPC Outlook connections.
        # Prevents signing-only or cleartext MAPI sessions. ExchangeDsc / HealthChecker recommendation.
        if (-not $State['InstallMailbox']) { return }
        Write-MyOutput 'Setting MAPI encryption as required on mailbox server'
        try {
            Set-MailboxServer -Identity $env:computername -MAPIEncryptionRequired $true -Confirm:$false -ErrorAction Stop
            Write-MyVerbose 'MAPI encryption required (Set-MailboxServer -MAPIEncryptionRequired $true)'
        }
        catch {
            Write-MyWarning ('Could not set MAPI encryption required: {0}' -f $_.Exception.Message)
        }
    }

    function Set-SchannelProtocol {
        param( [string]$Protocol, [bool]$Enable )
        $base = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols"
        $enabled       = if ($Enable) { 1 } else { 0 }
        $disabledByDef = if ($Enable) { 0 } else { 1 }
        New-Item -Path $base -Name $Protocol -ErrorAction SilentlyContinue | Out-Null
        foreach ( $role in 'Client', 'Server') {
            New-Item -Path "$base\$Protocol" -Name $role -ErrorAction SilentlyContinue | Out-Null
            Set-ItemProperty -Path "$base\$Protocol\$role" -Name 'DisabledByDefault' -Value $disabledByDef -Type DWord
            Set-ItemProperty -Path "$base\$Protocol\$role" -Name 'Enabled'           -Value $enabled       -Type DWord
            Register-ExecutedCommand -Category 'Hardening' -Command ("Set-ItemProperty '{0}\{1}\{2}' -Name Enabled -Value {3}; Set-ItemProperty '{0}\{1}\{2}' -Name DisabledByDefault -Value {4}" -f $base, $Protocol, $role, $enabled, $disabledByDef)
        }
    }

    function Set-NetFrameworkStrongCrypto {
        # HealthChecker requires all 4 paths (v4.0 + v2.0, 64-bit + 32-bit)
        foreach ( $path in 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319',
                            'HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319',
                            'HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727',
                            'HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v2.0.50727') {
            if (Test-Path $path) {
                Set-ItemProperty -Path $path -Name 'SystemDefaultTlsVersions' -Value 1 -Type DWord
                Set-ItemProperty -Path $path -Name 'SchUseStrongCrypto'        -Value 1 -Type DWord
                Register-ExecutedCommand -Category 'Hardening' -Command ("Set-ItemProperty '{0}' -Name SystemDefaultTlsVersions -Value 1; Set-ItemProperty '{0}' -Name SchUseStrongCrypto -Value 1" -f $path)
            }
        }
    }

    function Set-TLSSettings {

        param(
            [switch]$TLS12,
            [switch]$TLS13
        )

        if ( $TLS12) {
            Write-MyVerbose 'Enabling TLS 1.2 and configuring .NET Framework strong crypto'
            Set-NetFrameworkStrongCrypto
            Set-SchannelProtocol -Protocol 'TLS 1.2' -Enable $true
        }
        else {
            Write-MyVerbose 'Disabling TLS 1.2'
            Set-SchannelProtocol -Protocol 'TLS 1.2' -Enable $false
        }

        if ( [System.Version]$FullOSVersion -ge [System.Version]$WS2022_PREFULL -and [System.Version]$State['ExSetupVersion'] -ge [System.Version]$EX2019SETUPEXE_CU15) {
            if ( $TLS13) {
                Write-MyVerbose 'Enabling TLS 1.3 and configuring .NET Framework strong crypto'
                Set-NetFrameworkStrongCrypto
                Set-SchannelProtocol -Protocol 'TLS 1.3' -Enable $true
                # Configure the TLS 1.3 cipher suites (cmdlet requires WS2022+)
                if (Get-Command Enable-TlsCipherSuite -ErrorAction SilentlyContinue) {
                    Enable-TlsCipherSuite -Name TLS_AES_256_GCM_SHA384 -Position 0
                    Enable-TlsCipherSuite -Name TLS_AES_128_GCM_SHA256 -Position 1
                }
                else {
                    Write-MyWarning 'Enable-TlsCipherSuite cmdlet not available on this OS, skipping TLS 1.3 cipher suite configuration'
                }
            }
            else {
                Write-MyVerbose 'Disabling TLS 1.3'
                Set-SchannelProtocol -Protocol 'TLS 1.3' -Enable $false
                Disable-TlsCipherSuite -Name TLS_AES_256_GCM_SHA384 -ErrorAction SilentlyContinue
                Disable-TlsCipherSuite -Name TLS_AES_128_GCM_SHA256 -ErrorAction SilentlyContinue
            }
        }
        else {
            Write-MyWarning 'TLS13 configuration not supported for this OS or Exchange version'
        }

    }

    function Enable-WindowsDefenderExclusions {

        if ( Get-Command -Name Add-MpPreference -ErrorAction SilentlyContinue) {
            $SystemRoot = $Env:SystemRoot
            $SystemDrive = $Env:SystemDrive

            Write-MyOutput 'Configuring Windows Defender folder exclusions'
            if ( $State['TargetPath']) {
                $InstallFolder = $State['TargetPath']
            }
            else {
                # TargetPath not specified, using default location
                $InstallFolder = 'C:\Program Files\Microsoft\Exchange Server\V15'
            }

            $Locations = @(
                "$SystemRoot|Cluster",
                "$InstallFolder|ClientAccess\OAB,FIP-FS,GroupMetrics,Logging,Mailbox",
                "$InstallFolder\TransportRoles\Data|IpFilter,Queue,SenderReputation,Temp",
                "$InstallFolder\TransportRoles|Logs,Pickup,Replay",
                "$InstallFolder\UnifiedMessaging|Grammars,Prompts,Temp,VoiceMail",
                "$InstallFolder|Working\OleConverter",
                "$SystemDrive\InetPub\Temp|IIS Temporary Compressed Files",
                "$SystemDrive|Temp\OICE_*"
            )

            foreach ( $Location in $Locations) {
                $Parts = $Location -split '\|'
                $Items = $Parts[1] -split ','
                foreach ( $Item in $Items) {
                    $ExcludeLocation = Join-Path -Path $Parts[0] -ChildPath $Item
                    Write-MyVerbose "WindowsDefender: Excluding location $ExcludeLocation"
                    try {
                        Add-MpPreference -ExclusionPath $ExcludeLocation -ErrorAction SilentlyContinue
                    }
                    catch {
                        Write-MyWarning $_.Exception.Message
                    }
                }
            }

            Write-MyOutput 'Configuring Windows Defender exclusions: NodeRunner process'
            $Processes = @(
                "$InstallFolder\Bin|ComplianceAuditService.exe,Microsoft.Exchange.Directory.TopologyService.exe,Microsoft.Exchange.EdgeSyncSvc.exe,Microsoft.Exchange.Notifications.Broker.exe,Microsoft.Exchange.ProtectedServiceHost.exe,Microsoft.Exchange.RPCClientAccess.Service.exe,Microsoft.Exchange.Search.Service.exe,Microsoft.Exchange.Store.Service.exe,Microsoft.Exchange.Store.Worker.exe,MSExchangeCompliance.exe,MSExchangeDagMgmt.exe,MSExchangeDelivery.exe,MSExchangeFrontendTransport.exe,MSExchangeMailboxAssistants.exe,MSExchangeMailboxReplication.exe,MSExchangeRepl.exe,MSExchangeSubmission.exe,MSExchangeThrottling.exe,OleConverter.exe,UmService.exe,UmWorkerProcess.exe,wsbexchange.exe,EdgeTransport.exe,Microsoft.Exchange.AntispamUpdateSvc.exe,Microsoft.Exchange.Diagnostics.Service.exe,Microsoft.Exchange.Servicehost.exe,MSExchangeHMHost.exe,MSExchangeHMWorker.exe,MSExchangeTransport.exe,MSExchangeTransportLogSearch.exe",
                "$InstallFolder\FIP-FS\Bin|fms.exe,ScanEngineTest.exe,ScanningProcess.exe,UpdateService.exe",
                "$InstallFolder|Bin\Search\Ceres|HostController\HostControllerService.exe,Runtime\1.0\Noderunner.exe,ParserServer\ParserServer.exe",
                "$InstallFolder|FrontEnd\PopImap|Microsoft.Exchange.Imap4.exe,Microsoft.Exchange.Pop3.exe",
                "$InstallFolder|ClientAccess\PopImap\Microsoft.Exchange.Imap4service.exe,Microsoft.Exchange.Pop3service.exe",
                "$InstallFolder|FrontEnd\CallRouter|Microsoft.Exchange.UM.CallRouter.exe",
                "$InstallFolder|TransportRoles\agents\Hygiene\Microsoft.Exchange.ContentFilter.Wrapper.exe"
            )

            foreach ( $Process in $Processes) {
                $Parts = $Process -split '\|'
                $Items = $Parts[1] -split ','
                foreach ( $Item in $Items) {
                    $ExcludeProcess = Join-Path -Path $Parts[0] -ChildPath $Item
                    Write-MyVerbose "WindowsDefender: Excluding process $ExcludeProcess"
                    try {
                        Add-MpPreference -ExclusionProcess $ExcludeProcess -ErrorAction SilentlyContinue
                    }
                    catch {
                        Write-MyWarning $_.Exception.Message
                    }
                }
            }

            $Extensions = 'dsc', 'txt', 'cfg', 'grxml', 'lzx', 'config', 'chk', 'edb', 'jfm', 'jrs', 'log', 'que'
            foreach ( $Extension in $Extensions) {
                $ExcludeExtension = '.{0}' -f $Extension
                Write-MyVerbose "WindowsDefender: Excluding extension $ExcludeExtension"
                try {
                    Add-MpPreference -ExclusionExtension $ExcludeExtension -ErrorAction SilentlyContinue
                }
                catch {
                    Write-MyWarning $_.Exception.Message
                }
            }
            Register-ExecutedCommand -Category 'Hardening' -Command ("Add-MpPreference -ExclusionPath '{0}\Mailbox','{0}\Logging','{0}\FIP-FS',...  # see chapter 8.5 for complete path list" -f $InstallFolder)
            Register-ExecutedCommand -Category 'Hardening' -Command ("Add-MpPreference -ExclusionProcess '{0}\Bin\MSExchangeDelivery.exe','{0}\Bin\MSExchangeTransport.exe',...  # see chapter 8.5 for complete process list" -f $InstallFolder)
            Register-ExecutedCommand -Category 'Hardening' -Command 'Add-MpPreference -ExclusionExtension .edb,.jrs,.jfm,.chk,.log,.que,.cfg,.grxml,.lzx,.config,.dsc,.txt'
        }
        else {
            Write-MyVerbose 'Windows Defender not installed'
        }
    }

    function Disable-DefenderTamperProtection {
        # Tamper Protection blocks Set-MpPreference from taking effect. It cannot be disabled
        # via PowerShell/registry once MDE/Intune enforces it — those must be set via the
        # Security Center / Intune policy. On unmanaged devices we can flip the registry flag
        # as best-effort. Re-enabled in Enable-DefenderTamperProtection.
        if (-not (Get-Command -Name Get-MpComputerStatus -ErrorAction SilentlyContinue)) { return }
        try {
            $status = Get-MpComputerStatus -ErrorAction Stop
            if (-not $status.IsTamperProtected) {
                Write-MyVerbose 'Defender Tamper Protection already off — nothing to do'
                return
            }
            Write-MyOutput 'Attempting to disable Defender Tamper Protection (best-effort, registry)'
            $tpPath = 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Features'
            # Capture current value so we can restore it, even if not present
            $prev   = (Get-ItemProperty -Path $tpPath -Name 'TamperProtection' -ErrorAction SilentlyContinue).TamperProtection
            if ($null -eq $prev) { $State['DefenderTPPrev'] = '__absent__' } else { $State['DefenderTPPrev'] = [int]$prev }
            Set-RegistryValue -Path $tpPath -Name 'TamperProtection' -Value 0
            Start-Sleep -Seconds 2
            $post = Get-MpComputerStatus -ErrorAction SilentlyContinue
            if ($post -and $post.IsTamperProtected) {
                Write-MyWarning 'Tamper Protection still active — likely enforced by Intune/MDE. Realtime disable may be ignored.'
                Write-MyWarning '  Disable Tamper Protection manually in Windows Security / Intune before install, or accept that setup runs with AV active.'
            }
            else {
                Write-MyVerbose 'Tamper Protection flag cleared successfully'
            }
            $State['DefenderTPDisabledByEXpress'] = $true
            Save-State $State
        }
        catch {
            Write-MyWarning ('Could not inspect/disable Tamper Protection: {0}' -f $_.Exception.Message)
        }
    }

    function Enable-DefenderTamperProtection {
        # Restore the Tamper Protection registry value we captured before flipping it.
        if (-not $State['DefenderTPDisabledByEXpress']) { return }
        try {
            $tpPath = 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Features'
            $prev   = $State['DefenderTPPrev']
            if ($prev -eq '__absent__') {
                Remove-ItemProperty -Path $tpPath -Name 'TamperProtection' -ErrorAction SilentlyContinue
                Write-MyOutput 'Tamper Protection registry value removed (original state)'
            }
            elseif ($null -ne $prev) {
                Set-RegistryValue -Path $tpPath -Name 'TamperProtection' -Value ([int]$prev)
                Write-MyOutput ('Tamper Protection registry value restored to {0}' -f $prev)
            }
            $State.Remove('DefenderTPDisabledByEXpress') | Out-Null
            $State.Remove('DefenderTPPrev') | Out-Null
            Save-State $State
        }
        catch {
            Write-MyWarning ('Could not restore Tamper Protection: {0}' -f $_.Exception.Message)
        }
    }

    function Disable-DefenderRealtimeMonitoring {
        # Temporarily disable Defender real-time scanning during Exchange install/hardening.
        # Setup and SU runs generate massive file I/O (ECP/OWA .config unpacking, assembly
        # ngen, transport agents) that Defender scans inline, causing setup to stall or fail
        # with random file-lock errors. Re-enabled at the start of Phase 6.
        # Accepted risk: GPO/Intune may re-enable during the window. Flag is idempotent.
        if (-not (Get-Command -Name Set-MpPreference -ErrorAction SilentlyContinue)) {
            Write-MyVerbose 'Windows Defender not installed — skipping realtime disable'
            return
        }
        # Tamper Protection must be cleared first, otherwise Set-MpPreference is silently ignored.
        Disable-DefenderTamperProtection
        try {
            $pref = Get-MpPreference -ErrorAction Stop
            if ($pref.DisableRealtimeMonitoring) {
                Write-MyVerbose 'Defender realtime monitoring already disabled — leaving as-is'
                return
            }
            Write-MyOutput 'Disabling Windows Defender realtime monitoring during Exchange install'
            Set-MpPreference -DisableRealtimeMonitoring $true -ErrorAction Stop
            Start-Sleep -Seconds 1
            $post = Get-MpPreference -ErrorAction SilentlyContinue
            if ($post -and -not $post.DisableRealtimeMonitoring) {
                Write-MyWarning 'Realtime monitoring did not stay disabled — Tamper Protection or policy override active. Continuing with AV on.'
                return
            }
            $State['DefenderRealtimeDisabledByEXpress'] = $true
            Save-State $State
        }
        catch {
            Write-MyWarning ('Could not disable Defender realtime monitoring: {0}' -f $_.Exception.Message)
        }
    }

    function Enable-DefenderRealtimeMonitoring {
        # Re-enable Defender realtime scanning. With -Force the function always attempts
        # to turn realtime on, regardless of whether EXpress was the one to disable it —
        # this is used right before the Word report generates so the report reflects an
        # active protection state after installation. Without -Force, it only reverses
        # an EXpress-initiated disable (flag set in Disable-DefenderRealtimeMonitoring).
        param([switch]$Force)
        if (-not (Get-Command -Name Set-MpPreference -ErrorAction SilentlyContinue)) { return }
        $shouldAct = $Force -or $State['DefenderRealtimeDisabledByEXpress']
        if ($shouldAct) {
            try {
                $pref = Get-MpPreference -ErrorAction Stop
                if (-not $pref.DisableRealtimeMonitoring) {
                    Write-MyVerbose 'Defender realtime monitoring already enabled'
                }
                else {
                    if ($Force) { Write-MyOutput 'Ensuring Windows Defender realtime monitoring is enabled (pre-report)' }
                    else        { Write-MyOutput 'Re-enabling Windows Defender realtime monitoring' }
                    Set-MpPreference -DisableRealtimeMonitoring $false -ErrorAction Stop
                    Start-Sleep -Seconds 1
                    $post = Get-MpPreference -ErrorAction SilentlyContinue
                    if ($post -and $post.DisableRealtimeMonitoring) {
                        Write-MyWarning 'Realtime monitoring still disabled after set — Tamper Protection or policy override active.'
                    }
                }
                if ($State['DefenderRealtimeDisabledByEXpress']) {
                    $State.Remove('DefenderRealtimeDisabledByEXpress') | Out-Null
                    Save-State $State
                }
            }
            catch {
                Write-MyWarning ('Could not re-enable Defender realtime monitoring: {0}' -f $_.Exception.Message)
            }
        }
        else {
            Write-MyVerbose 'Defender realtime monitoring was not disabled by EXpress — skipping re-enable'
        }
        # Restore Tamper Protection regardless — it may have been flipped without realtime change.
        Enable-DefenderTamperProtection
    }

    # Return location of mounted drive if ISO specified
    function Resolve-SourcePath {
        param (
            [String]$SourceImage
        )
        $disk = Get-DiskImage -ImagePath $SourceImage -ErrorAction SilentlyContinue
        if ( $disk) {
            if ( $disk.Attached) {
                $vol = $disk | Get-Volume -ErrorAction SilentlyContinue
                if ( $vol) {
                    $Drive = $vol.DriveLetter
                }
                else {
                    Write-Verbose ('{0} already attached but no drive letter - will mount again' -f $SourceImage)
                    $Drive = (Mount-DiskImage -ImagePath $SourceImage -PassThru | Get-Volume).DriveLetter
                }
            }
            else {
                $Drive = (Mount-DiskImage -ImagePath $SourceImage -PassThru | Get-Volume).DriveLetter
            }
            $SourcePath = '{0}:\' -f $Drive
            Write-Verbose ('Mounted {0} on drive {1}' -f $SourceImage, $SourcePath)
            return $SourcePath
        }
        else {
            return $null
        }
    }

    function Get-VCRuntime {
        param (
            [String]$version,
            [String]$MinBuild = ''
        )
        Write-MyVerbose ('Looking for presence of Visual C++ v{0} Runtime' -f $version)
        $presence = $false
        $build = $null

        # Primary check: VisualStudio registry paths (used by VC++ 2015+ / VS 14.x bundles,
        # and some variants of 2012/2013 installers).
        $RegPaths = @(
            'HKLM:\Software\WOW6432Node\Microsoft\VisualStudio\{0}\VC\Runtimes\x64',
            'HKLM:\Software\Microsoft\VisualStudio\{0}\VC\Runtimes\x64',
            'HKLM:\Software\WOW6432Node\Microsoft\VisualStudio\{0}\VC\VCRedist\x64',
            'HKLM:\Software\Microsoft\VisualStudio\{0}\VC\VCRedist\x64')
        foreach ( $RegPath in $RegPaths) {
            $Key = (Get-ItemProperty -Path ($RegPath -f $version) -Name Installed -ErrorAction SilentlyContinue).Installed
            if ( $Key -eq 1) {
                $build = (Get-ItemProperty -Path ($RegPath -f $version) -Name Version -ErrorAction SilentlyContinue).Version
                $presence = $true
                break
            }
        }

        # Fallback 1: scan Add/Remove Programs for matching display name.
        # VC++ 2013 (12.0) and older standalone redistributables do not write to
        # the VisualStudio\{ver}\VC\Runtimes path — they only register here.
        if (-not $presence) {
            $yearMap = @{ '10.0' = '2010'; '11.0' = '2012'; '12.0' = '2013'; '14.0' = '2015' }
            $yearStr  = $yearMap[$version]
            if ($yearStr) {
                foreach ($hive in @('HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
                                    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')) {
                    # Match display name without requiring "(x64)" — the format varies by installer
                    # version (e.g. "...Redistributable (x64)..." vs "...x64 Minimum Runtime...").
                    $entry = Get-ChildItem $hive -ErrorAction SilentlyContinue |
                             Get-ItemProperty -ErrorAction SilentlyContinue |
                             Where-Object { $_.DisplayName -like "Microsoft Visual C++ $yearStr*" } |
                             Sort-Object DisplayVersion -Descending |
                             Select-Object -First 1
                    if ($entry) {
                        $build    = $entry.DisplayVersion
                        $presence = $true
                        Write-MyVerbose ('Found Visual C++ v{0} in Add/Remove Programs: {1}' -f $version, $entry.DisplayName)
                        break
                    }
                }
            }
        }

        # Fallback 2: check the runtime DLL in System32 — the same check Exchange Setup uses.
        # msvcr110.dll = VC++ 2012 (11.0), msvcr120.dll = VC++ 2013 (12.0)
        if (-not $presence) {
            $dllMap = @{ '11.0' = 'msvcr110.dll'; '12.0' = 'msvcr120.dll' }
            $dll    = $dllMap[$version]
            if ($dll) {
                $dllPath = Join-Path $env:SystemRoot "System32\$dll"
                if (Test-Path $dllPath) {
                    $build    = (Get-Item $dllPath -ErrorAction SilentlyContinue).VersionInfo.ProductVersion
                    $presence = $true
                    Write-MyVerbose ('Found Visual C++ v{0} via {1}, version {2}' -f $version, $dll, $build)
                }
            }
        }

        if ($presence) {
            Write-MyVerbose ('Found Visual C++ Runtime v{0}, build {1}' -f $version, $build)
            if ($MinBuild -and $build -and ([System.Version]$build -lt [System.Version]$MinBuild)) {
                Write-MyVerbose ('Visual C++ v{0} build {1} is older than required minimum {2} — will update' -f $version, $build, $MinBuild)
                return $false
            }
        }
        else {
            Write-MyVerbose ('Could not find Visual C++ v{0} Runtime installed' -f $version)
        }
        return $presence
    }

    function Start-DisableMSExchangeAutodiscoverAppPoolJob {

        $ScriptBlock = {
            do {
                # Use Test-Path instead of Get-WebAppPoolState: the latter internally calls
                # Get-WebItemState which throws PathNotFound and is NOT suppressed by -ErrorAction SilentlyContinue.
                if (Test-Path 'IIS:\AppPools\MSExchangeAutodiscoverAppPool') {

                    Write-Verbose 'Stopping and blocking startup of MSExchangeAutodiscoverAppPool'
                    if ( (Get-WebAppPoolState -Name 'MSExchangeAutodiscoverAppPool').Value -ine 'Stopped') {
                        try {
                            Stop-WebAppPool -Name 'MSExchangeAutodiscoverAppPool' -ErrorAction Stop
                        }
                        catch {
                            Write-Error ('Failed to stop app pool: {0}' -f $_.Exception.Message)
                        }
                    }
                    try {
                        Set-ItemProperty "IIS:\AppPools\MSExchangeAutodiscoverAppPool" -Name "autoStart" -Value $false -ErrorAction Stop
                        Set-ItemProperty "IIS:\AppPools\MSExchangeAutodiscoverAppPool" -Name "startMode" -Value "OnDemand" -ErrorAction Stop
                    }
                    catch {
                        Write-Error ('Failed to update app pool properties: {0}' -f $_.Exception.Message)
                    }
                    return $true
                }
                else {
                    Write-Verbose ('MSExchangeAutodiscoverAppPool not found, waiting a bit ..')
                    Start-Sleep -Seconds 10
                }
            } while ($true)
        }

        $Job = Start-Job -ScriptBlock $ScriptBlock -Name ('DisableMSExchangeAutodiscoverAppPoolJob-{0}' -f $env:COMPUTERNAME)
        Add-BackgroundJob $Job

        Write-MyOutput ('Started background job to disable MSExchangeAutodiscoverAppPool (Job ID: {0})' -f $Job.Id)
        return $Job
    }

    function Enable-MSExchangeAutodiscoverAppPool {
        # Use Test-Path instead of Get-WebAppPoolState: the latter internally calls
        # Get-WebItemState which throws a provider PathNotFound error that is NOT
        # suppressed by -ErrorAction SilentlyContinue.
        if (-not (Test-Path 'IIS:\AppPools\MSExchangeAutodiscoverAppPool' -ErrorAction SilentlyContinue)) {
            Write-MyVerbose 'MSExchangeAutodiscoverAppPool not found'
            return $false
        }

        Write-MyOutput 'Starting and enabling startup of MSExchangeAutodiscoverAppPool'
        try {
            Start-WebAppPool -Name 'MSExchangeAutodiscoverAppPool' -ErrorAction Stop
        }
        catch {
            Write-MyWarning ('Failed to start app pool: {0}' -f $_.Exception.Message)
        }
        try {
            Set-ItemProperty 'IIS:\AppPools\MSExchangeAutodiscoverAppPool' -Name 'autoStart' -Value $true  -ErrorAction Stop
            Set-ItemProperty 'IIS:\AppPools\MSExchangeAutodiscoverAppPool' -Name 'startMode' -Value 'OnDemand' -ErrorAction Stop
        }
        catch {
            Write-MyWarning ('Failed to update app pool properties: {0}' -f $_.Exception.Message)
        }
        return $true
    }

    function Stop-BackgroundJobs {
        if ($Global:BackgroundJobs -and $Global:BackgroundJobs.Count -gt 0) {
            Write-MyVerbose "Cleaning up $($Global:BackgroundJobs.Count) background job(s)..."
            foreach ($Job in $Global:BackgroundJobs) {
                if ($Job.State -eq 'Running') {
                    # Wait up to 30 seconds for job to finish gracefully
                    $null = $Job | Wait-Job -Timeout 30 -ErrorAction SilentlyContinue
                    if ($Job.State -eq 'Running') {
                        Write-MyWarning ('Background job {0} (ID {1}) did not finish within 30 seconds, forcing stop' -f $Job.Name, $Job.Id)
                        Stop-Job -Job $Job -ErrorAction SilentlyContinue
                    }
                }
                $JobOutput = Receive-Job -Job $Job
                Write-MyVerbose ('Cleanup background job: {0} (ID {1}), Output {2}' -f $Job.Name, $Job.Id, $JobOutput)
                Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue
            }
            $Global:BackgroundJobs = @()
            Write-MyVerbose "Background job cleanup completed."
        }
    }

