    function Invoke-NativeCommand {
        param(
            [Parameter(Mandatory)][string]$FilePath,
            [string[]]$Arguments = @(),
            [string]$Tag = ''
        )
        $label = if ($Tag) { $Tag } else { Split-Path $FilePath -Leaf }
        $out = & $FilePath @Arguments 2>&1
        $rc  = $LASTEXITCODE
        if ($State['LogDebug']) {
            Write-ToTranscript 'EXE' ('{0} exit={1} args=[{2}]' -f $label, $rc, ($Arguments -join ' '))
            foreach ($line in $out) {
                if ($null -eq $line) { continue }
                $text = if ($line -is [System.Management.Automation.ErrorRecord]) { 'stderr: ' + $line.Exception.Message } else { [string]$line }
                if ($text) { Write-ToTranscript 'EXE' ('  {0}' -f $text) }
            }
        }
        return $rc
    }

    function Set-RegistryValue {
        param( [string]$Path, [string]$Name, $Value, [string]$PropertyType = 'DWord')
        if ( -not (Test-Path $Path -ErrorAction SilentlyContinue)) {
            New-Item -Path $Path -Force -ErrorAction SilentlyContinue | Out-Null
        }
        else {
            $existing = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
            if ($null -ne $existing -and $existing.$Name -eq $Value) {
                Write-MyVerbose ('Registry value already set: {0}\{1} = {2}' -f $Path, $Name, $Value)
                return
            }
        }
        New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $PropertyType -Force -ErrorAction SilentlyContinue | Out-Null
    }

    function Get-PSExecutionPolicy {
        $PSPolicyKey = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell' -Name ExecutionPolicy -ErrorAction SilentlyContinue
        if ( $PSPolicyKey) {
            Write-MyWarning "PowerShell Execution Policy is set to $($PSPolicyKey.ExecutionPolicy) through GPO"
        }
        else {
            Write-MyVerbose 'PowerShell Execution Policy not configured through GPO'
        }
        return $PSPolicyKey
    }

    function Invoke-WebDownload {
        # PS 5.1-compatible web download. Uses -SkipCertificateCheck on PS 6+,
        # falls back to WebClient with TLS 1.2 and cert bypass on PS 5.1.
        param([string]$Uri, [string]$OutFile)
        if ($PSVersionTable.PSVersion.Major -ge 6) {
            Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing -SkipCertificateCheck -ErrorAction Stop
        }
        else {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $prevCallback = [Net.ServicePointManager]::ServerCertificateValidationCallback
            [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
            try {
                $wc = New-Object System.Net.WebClient
                $wc.DownloadFile($Uri, $OutFile)
            }
            finally {
                [Net.ServicePointManager]::ServerCertificateValidationCallback = $prevCallback
            }
        }
    }

    function Get-MyPackage () {
        param ( [String]$Package, [String]$URL, [String]$FileName, [String]$InstallPath)
        $res = $true
        if ( !( Test-Path $(Join-Path $InstallPath $Filename))) {
            if ( $URL) {
                Write-MyOutput "Package $Package not found, downloading to $FileName"
                Write-MyVerbose "Source: $URL"
                $destPath = Join-Path $InstallPath $Filename
                $downloaded = $false
                $savedPP = $ProgressPreference
                $ProgressPreference = 'SilentlyContinue'
                for ($attempt = 1; $attempt -le 3; $attempt++) {
                    try {
                        Start-BitsTransfer -Source $URL -Destination $destPath -ErrorAction Stop
                        $downloaded = $true
                        break
                    }
                    catch {
                        Get-BitsTransfer -ErrorAction SilentlyContinue | Where-Object { $_.JobState -notin 'Transferred','Acknowledged' } | Remove-BitsTransfer -ErrorAction SilentlyContinue
                        Remove-Item -Path $destPath -ErrorAction SilentlyContinue
                        # 0x800704DD = ERROR_NOT_LOGGED_ON: BITS has no network logon session
                        # (common in Autopilot RunOnce context after reboot). Fall back to
                        # WebClient immediately — no point retrying BITS in this scenario.
                        $isBitsLogonError = $_.Exception.Message -match '0x800704DD|not logged on to the network'
                        if ($attempt -lt 3 -and -not $isBitsLogonError) {
                            Write-MyWarning ('Download attempt {0}/3 failed, retrying in {1} seconds: {2}' -f $attempt, ($attempt * 5), $_.Exception.Message)
                            Start-Sleep -Seconds ($attempt * 5)
                        }
                        else {
                            # Final attempt or BITS network-logon error: try web download as fallback
                            try {
                                if ($isBitsLogonError) {
                                    Write-MyVerbose 'BITS unavailable (no network logon session) — using web download'
                                } else {
                                    Write-MyVerbose 'BITS failed after 3 attempts, trying web download as fallback'
                                }
                                Invoke-WebDownload -Uri $URL -OutFile $destPath
                                $downloaded = $true
                                break
                            }
                            catch {
                                Write-MyWarning ('Problem downloading package from URL: {0}' -f $_.Exception.Message)
                                Remove-Item -Path $destPath -ErrorAction SilentlyContinue
                            }
                        }
                    }
                }
                $ProgressPreference = $savedPP
                if (-not $downloaded) {
                    $res = $false
                    Write-MyWarning ('Could not download {0}. For offline or proxy-restricted deployments:' -f $FileName)
                    Write-MyOutput  ('  1. Run  .\tools\Get-EXpressDownloads.ps1  on an internet-connected machine.')
                    Write-MyOutput  ('  2. Copy the sources\ folder to {0}' -f $InstallPath)
                }
            }
            else {
                Write-MyWarning "$FileName not present, skipping"
                $res = $false
            }
        }
        else {
            Write-MyVerbose "Located $Package ($InstallPath\$FileName)"
        }
        return $res
    }

    function Get-CurrentUserName {
        return [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }

    function Test-Admin {
        $currentPrincipal = New-Object System.Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )
        return $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )
    }

    function Test-RebootPending {
        # Returns $true if Windows signals a pending reboot. Used to decide whether
        # a phase boundary really needs to reboot, or if we can continue in-process.
        $reasons = @()
        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') {
            $reasons += 'CBS RebootPending'
        }
        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') {
            $reasons += 'WindowsUpdate RebootRequired'
        }
        $pfro = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
        if ($pfro -and $pfro.PendingFileRenameOperations) {
            $reasons += 'PendingFileRenameOperations'
        }
        $cn = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName' -Name 'ComputerName' -ErrorAction SilentlyContinue
        $pcn = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName' -Name 'ComputerName' -ErrorAction SilentlyContinue
        if ($cn -and $pcn -and $cn.ComputerName -ne $pcn.ComputerName) {
            $reasons += 'Pending computer rename'
        }
        try {
            $ccm = Invoke-CimMethod -Namespace 'ROOT\ccm\ClientSDK' -ClassName 'CCM_ClientUtilities' -MethodName 'DetermineIfRebootPending' -ErrorAction Stop
            if ($ccm -and ($ccm.RebootPending -or $ccm.IsHardRebootPending)) {
                $reasons += 'CCM ClientSDK'
            }
        } catch { }
        if ($reasons.Count -gt 0) {
            Write-MyVerbose ('Reboot pending: {0}' -f ($reasons -join ', '))
            return $true
        }
        return $false
    }

    function Test-ADGroupMember ([int]$RelativeId) {
        try {
            $FRNC = Get-ForestRootNC
            $ADRootSID = ([ADSI]"LDAP://$FRNC").ObjectSID[0]
            if ($null -eq $ADRootSID) {
                Write-MyWarning 'Could not retrieve forest root SID — AD may be unreachable'
                return $false
            }
            $SID = (New-Object System.Security.Principal.SecurityIdentifier ($ADRootSID, 0)).Value.toString()
            return [Security.Principal.WindowsIdentity]::GetCurrent().Groups | Where-Object { $_.Value -eq "$SID-$RelativeId" }
        }
        catch {
            Write-MyWarning ('Test-ADGroupMember failed: {0}' -f $_.Exception.Message)
            return $false
        }
    }

    function Test-SchemaAdmin     { Test-ADGroupMember 518 }
    function Test-EnterpriseAdmin { Test-ADGroupMember 519 }

    function Test-ServerCore {
        (Get-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows NT\CurrentVersion' -Name 'InstallationType' -ErrorAction SilentlyContinue).InstallationType -eq 'Server Core'
    }

    function Test-RebootPending {
        $Pending = $False
        if ( Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue) {
            $Pending = $True
        }
        if ( Test-Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending' -ErrorAction SilentlyContinue) {
            $Pending = $True
        }
        if ( Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired' -ErrorAction SilentlyContinue) {
            $Pending = $True
        }
        return $Pending
    }

    function Enable-RunOnce {
        Write-MyOutput 'Set script to run once after reboot'
        # When compiled with PS2Exe the script runs as a standalone .exe — invoke it directly.
        # Otherwise use the current PowerShell interpreter (powershell.exe or pwsh.exe).
        $isExe = $ScriptFullName -imatch '\.exe$'
        $logFlags = ''
        if ($State['LogVerbose']) { $logFlags += ' -Verbose' }
        if ($State['LogDebug'])   { $logFlags += ' -Debug' }
        if ($isExe) {
            $RunOnce = "`"$ScriptFullName`" -InstallPath `"$InstallPath`"$logFlags"
        }
        else {
            $PSExe = (Get-Process -Id $PID).Path
            $RunOnce = "`"$PSExe`" -NoProfile -ExecutionPolicy Unrestricted -Command `"& `'$ScriptFullName`' -InstallPath `'$InstallPath`'$logFlags`""
        }
        Write-MyVerbose "RunOnce: $RunOnce"
        Set-RegistryValue -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce' -Name $ScriptName -Value $RunOnce -PropertyType String
    }

    function Disable-UAC {
        Write-MyVerbose 'Disabling User Account Control'
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 0
    }

    function Enable-UAC {
        Write-MyVerbose 'Enabling User Account Control'
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 1
    }

    function Disable-IEESC {
        $AdminKey = 'HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}'
        $UserKey  = 'HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}'
        $alreadyOff = ((Get-ItemProperty -Path $AdminKey -Name IsInstalled -ErrorAction SilentlyContinue).IsInstalled -eq 0) -and
                      ((Get-ItemProperty -Path $UserKey  -Name IsInstalled -ErrorAction SilentlyContinue).IsInstalled -eq 0)
        if ($alreadyOff) { Write-MyVerbose 'IE Enhanced Security Configuration already disabled'; return }
        Write-MyOutput 'Disabling IE Enhanced Security Configuration'
        New-Item -Path (Split-Path $AdminKey -Parent) -Name (Split-Path $AdminKey -Leaf) -ErrorAction SilentlyContinue | Out-Null
        Set-ItemProperty -Path $AdminKey -Name 'IsInstalled' -Value 0 -Force | Out-Null
        New-Item -Path (Split-Path $UserKey -Parent) -Name (Split-Path $UserKey -Leaf) -ErrorAction SilentlyContinue | Out-Null
        Set-ItemProperty -Path $UserKey  -Name 'IsInstalled' -Value 0 -Force | Out-Null
        if ( Get-Process -Name explorer.exe -ErrorAction SilentlyContinue) {
            Stop-Process -Name Explorer
        }
    }

    function Enable-IEESC {
        Write-MyVerbose 'Enabling IE Enhanced Security Configuration'
        $AdminKey = 'HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}'
        $UserKey = 'HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}'
        New-Item -Path (Split-Path $AdminKey -Parent) -Name (Split-Path $AdminKey -Leaf) -ErrorAction SilentlyContinue | Out-Null
        Set-ItemProperty -Path $AdminKey -Name 'IsInstalled' -Value 1 -Force | Out-Null
        New-Item -Path (Split-Path $UserKey -Parent) -Name (Split-Path $UserKey -Leaf) -ErrorAction SilentlyContinue | Out-Null
        Set-ItemProperty -Path $UserKey -Name 'IsInstalled' -Value 1 -Force | Out-Null
        if ( Get-Process -Name explorer.exe -ErrorAction SilentlyContinue) {
            Stop-Process -Name Explorer
        }
    }

    function Get-FullDomainAccount {
        $PlainTextAccount = $State['AdminAccount']
        if ( $PlainTextAccount.indexOf('\') -gt 0) {
            $Parts = $PlainTextAccount.split('\')
            $Domain = $Parts[0]
            $UserName = $Parts[1]
            return "$Domain\$UserName"
        }
        else {
            if ( $PlainTextAccount.indexOf('@') -gt 0) {
                return $PlainTextAccount
            }
            else {
                $Domain = $env:USERDOMAIN
                $UserName = $PlainTextAccount
                return "$Domain\$UserName"
            }
        }
    }

    function Test-LocalCredential {
        [CmdletBinding()]
        param
        (
            [string]$UserName,
            [string]$ComputerName = $env:COMPUTERNAME,
            [string]$Password
        )
        if (!($UserName) -or !($Password)) {
            Write-Warning 'Test-LocalCredential: Please specify both user name and password'
        }
        else {
            Add-Type -AssemblyName System.DirectoryServices.AccountManagement
            $DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('machine', $ComputerName)
            $DS.ValidateCredentials($UserName, $Password )
        }
    }

    function Test-Credentials {
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR((ConvertTo-SecureString $State['AdminPassword']))
        $PlainTextPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        $FullPlainTextAccount = Get-FullDomainAccount
        try {
            if ( $State['InstallEdge']) {
                $Username = $FullPlainTextAccount.split("\")[-1]
                return $( Test-LocalCredential -UserName $Username -Password $PlainTextPassword)
            }
            else {
                $dc = New-Object DirectoryServices.DirectoryEntry( $Null, $FullPlainTextAccount, $PlainTextPassword)
                if ($dc.Name) {
                    return $true
                }
                else {
                    return $false
                }
            }

        }
        catch {
            return $false
        }
        return $false
    }

    function Get-ValidatedCredentials {
        # Interactive credential prompt with validation retry loop (max 3 attempts).
        # Returns $true when valid credentials are stored in State, $false if all attempts fail.
        # Only call this when [Environment]::UserInteractive is $true.
        #
        # GUI detection: Get-Credential shows a Win32 dialog only when all three hold:
        #   1. ConsoleHost (not ISE, not PS2Exe, not a remote host)
        #   2. UserInteractive (not a service / scheduled-task session)
        #   3. A real desktop session (SESSIONNAME = Console or RDP-*; empty = Session 0 / no window station)
        # When any condition is false we go straight to Read-Host to avoid the silent-$null fallback.
        $sessionName = [string]$env:SESSIONNAME
        $useGui = (-not $IsPS2Exe) -and
                  ($Host.Name -eq 'ConsoleHost') -and
                  [Environment]::UserInteractive -and
                  ($sessionName -match '^(Console|RDP)')

        $maxAttempts = 3
        for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
            try {
                $defaultUser = if ($State['AdminAccount']) { $State['AdminAccount'] } else { [System.Security.Principal.WindowsIdentity]::GetCurrent().Name }
                $Script:Credentials = $null
                if ($useGui) {
                    $rawCred = Get-Credential -UserName $defaultUser -Message ('Enter credentials for Autopilot (attempt {0}/{1})' -f $attempt, $maxAttempts)
                    # Get-Credential can return a PSObject wrapper in some terminal environments; unwrap before assigning to typed variable.
                    $Script:Credentials = if ($rawCred -is [pscredential]) { $rawCred }
                                          elseif ($rawCred -and $rawCred.PSObject.BaseObject -is [pscredential]) { $rawCred.PSObject.BaseObject }
                                          else { $null }
                }
                if (-not $Script:Credentials) {
                    Write-MyOutput ('Enter credentials for Autopilot (attempt {0}/{1})' -f $attempt, $maxAttempts)
                    $fbUser = Read-Host -Prompt ('Username [{0}]' -f $defaultUser)
                    if ([string]::IsNullOrWhiteSpace($fbUser)) { $fbUser = $defaultUser }
                    $fbPass = Read-Host -Prompt 'Password' -AsSecureString
                    if ($fbPass -and $fbPass.Length -gt 0) {
                        $Script:Credentials = New-Object System.Management.Automation.PSCredential($fbUser, $fbPass)
                    }
                }
                if (-not $Script:Credentials) {
                    Write-MyWarning 'No credentials entered'
                }
                else {
                    $State['AdminAccount'] = $Script:Credentials.UserName
                    # ConvertFrom-SecureString without -Key uses DPAPI (user+machine bound).
                    # Autopilot always resumes as the same user on the same machine, so this is safe.
                    $State['AdminPassword'] = ($Script:Credentials.Password | ConvertFrom-SecureString)
                    Write-MyOutput ('Checking credentials (attempt {0}/{1})' -f $attempt, $maxAttempts)
                    if (Test-Credentials) {
                        Write-MyOutput 'Credentials valid'
                        return $true
                    }
                    else {
                        Write-MyWarning ("Credentials for '{0}' are invalid" -f $State['AdminAccount'])
                    }
                }
            }
            catch {
                Write-MyWarning ('Credential prompt cancelled or failed: {0}' -f $_.Exception.Message)
            }
            if ($attempt -lt $maxAttempts) {
                $choice = $Host.UI.PromptForChoice('Invalid credentials', 'Retry or quit?', @('&Retry', '&Quit'), 0)
                if ($choice -ne 0) {
                    Write-MyError 'Credential entry aborted by user'
                    return $false
                }
            }
        }
        Write-MyError ('Credential validation failed after {0} attempts' -f $maxAttempts)
        return $false
    }

    function Enable-AutoLogon {
        Write-MyVerbose 'Enabling Automatic Logon'
        # SECURITY NOTE: This writes the password in plaintext to the registry.
        # Disable-AutoLogon is called after the next login to remove these values immediately.
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR((ConvertTo-SecureString $State['AdminPassword']))
        $PlainTextPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        $PlainTextAccount = $State['AdminAccount']
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name AutoAdminLogon -Value 1 -PropertyType String
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultUserName -Value $PlainTextAccount -PropertyType String
        Set-RegistryValue -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultPassword -Value $PlainTextPassword -PropertyType String
    }

    function Disable-AutoLogon {
        Write-MyVerbose 'Disabling Automatic Logon'
        Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name AutoAdminLogon -ErrorAction SilentlyContinue | Out-Null
        Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultUserName -ErrorAction SilentlyContinue | Out-Null
        Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultPassword -ErrorAction SilentlyContinue | Out-Null
    }

    function Disable-OpenFileSecurityWarning {
        Write-MyVerbose 'Disabling File Security Warning dialog'
        Set-RegistryValue -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations' -Name 'LowRiskFileTypes' -Value '.exe;.msp;.msu;.msi' -PropertyType String
        Set-RegistryValue -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments' -Name 'SaveZoneInformation' -Value 1
        Remove-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations' -Name 'LowRiskFileTypes' -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments' -Name 'SaveZoneInformation' -ErrorAction SilentlyContinue
    }

    function Enable-OpenFileSecurityWarning {
        Write-MyVerbose 'Enabling File Security Warning dialog'
        Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations' -Name 'LowRiskFileTypes' -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments' -Name 'SaveZoneInformation' -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations' -Name 'LowRiskFileTypes' -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments' -Name 'SaveZoneInformation' -ErrorAction SilentlyContinue
    }

    function Invoke-Extract ( $FilePath, $FileName) {
        Write-MyVerbose "Extracting $FilePath\$FileName to $FilePath"
        $FullPath = Join-Path $FilePath $FileName
        if ( Test-Path $FullPath) {
            $TempNam = "$FullPath.zip"
            try {
                Copy-Item $FullPath $TempNam -Force -ErrorAction Stop
                Expand-Archive -Path $TempNam -DestinationPath $FilePath -Force -ErrorAction Stop
            }
            catch {
                Write-MyError ('Failed to extract {0}: {1}' -f $FullPath, $_.Exception.Message)
            }
            finally {
                Remove-Item $TempNam -ErrorAction SilentlyContinue
            }
        }
        else {
            Write-MyWarning "$FilePath\$FileName not found"
        }
    }

    function Invoke-Process ( $FilePath, $FileName, $ArgumentList) {
        $rval = 0
        $mspTempDir = $null
        $FullName = Join-Path $FilePath $FileName
        if ( Test-Path $FullName) {
            switch ( ([io.fileinfo]$Filename).extension.ToUpper()) {
                '.MSU' {
                    $ArgumentList += @( $FullName)
                    $ArgumentList += @( '/f')
                    $Cmd = "$env:SystemRoot\System32\WUSA.EXE"
                }
                '.MSI' {
                    $ArgumentList += @( '/i')
                    $ArgumentList += @( $FullName)
                    $Cmd = "MSIEXEC.EXE"
                }
                '.MSP' {
                    $ArgumentList += @( '/update')
                    $ArgumentList += @( $FullName)
                    $Cmd = 'MSIEXEC.EXE'
                }
                '.CAB' {
                    $mspTempDir = Join-Path $env:TEMP ('ExSU_' + [IO.Path]::GetFileNameWithoutExtension($FileName))
                    New-Item -ItemType Directory -Path $mspTempDir -Force | Out-Null
                    $expandOut = & "$env:SystemRoot\System32\expand.exe" -F:* $FullName $mspTempDir 2>&1
                    Write-MyVerbose ('expand.exe output: {0}' -f ($expandOut -join ' | '))
                    # Exchange SU CABs are often multi-level: expand any nested CABs into the same temp dir
                    $nestedCabs = Get-ChildItem -Path $mspTempDir -Filter '*.cab' -File -ErrorAction SilentlyContinue
                    foreach ($nestedCab in $nestedCabs) {
                        $nestedOut = & "$env:SystemRoot\System32\expand.exe" -F:* $nestedCab.FullName $mspTempDir 2>&1
                        Write-MyVerbose ('Nested CAB {0}: {1}' -f $nestedCab.Name, ($nestedOut -join ' | '))
                    }
                    $extractedFiles = Get-ChildItem -Path $mspTempDir -Recurse -File -ErrorAction SilentlyContinue
                    if ($extractedFiles) {
                        Write-MyVerbose ('CAB contents: {0}' -f ($extractedFiles.Name -join ', '))
                    } else {
                        Write-MyVerbose 'CAB extraction produced no files'
                    }
                    $mspFile = $extractedFiles | Where-Object { $_.Extension -eq '.msp' } | Select-Object -First 1
                    $exeFile = $extractedFiles | Where-Object { $_.Extension -eq '.exe' -and $_.Name -notlike '*.cab' } | Select-Object -First 1
                    if ($mspFile) {
                        $ArgumentList += @('/update')
                        $ArgumentList += @($mspFile.FullName)
                        $Cmd = 'MSIEXEC.EXE'
                    }
                    elseif ($exeFile) {
                        $Cmd = $exeFile.FullName
                    }
                    else {
                        # No MSP/EXE found — WU-style CABs (Exchange SE SU) carry a compressed payload
                        # that expand.exe cannot unpack as MSP. Install directly via DISM /Add-Package.
                        Write-MyVerbose ('No MSP or EXE found in {0} — falling back to DISM /Add-Package' -f $FileName)
                        Remove-Item -Path $mspTempDir -Recurse -Force -ErrorAction SilentlyContinue
                        $mspTempDir = $null
                        $Cmd = "$env:SystemRoot\System32\dism.exe"
                        $ArgumentList = @('/Online', '/Add-Package', "/PackagePath:$FullName", '/Quiet', '/NoRestart')
                    }
                }
                default {
                    $Cmd = $FullName
                }
            }
            Write-MyVerbose "Executing $Cmd $($ArgumentList -Join ' ')"
            $rval = ( Start-Process -FilePath $Cmd -ArgumentList $ArgumentList -NoNewWindow -PassThru -Wait).Exitcode
            Write-MyVerbose "Process exited with code $rval"
            if ($mspTempDir) { Remove-Item -Path $mspTempDir -Recurse -Force -ErrorAction SilentlyContinue }
        }
        else {
            Write-MyWarning "$FullName not found"
            $rval = -1
        }
        return $rval
    }
