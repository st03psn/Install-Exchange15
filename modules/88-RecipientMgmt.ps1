    function Test-IsClientOS {
        # Returns $true when running on a client SKU (Windows 10/11), $false on Server
        $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
        # ProductType: 1 = Workstation/Client, 2 = Domain Controller, 3 = Server
        return ($osInfo.ProductType -eq 1)
    }

    function Install-RecipientManagementPrereqs {
        # Phase 1 of Recipient Management install: OS detection and prerequisite installation
        if (Test-IsClientOS) {
            Write-MyVerbose 'Client OS detected, installing RSAT Active Directory tools via Add-WindowsCapability'
            try {
                $cap = Get-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools*' -ErrorAction Stop
                if ($cap.State -ne 'Installed') {
                    Add-WindowsCapability -Online -Name $cap.Name -ErrorAction Stop | Out-Null
                    Write-MyStep -Label 'RSAT ADDS tools' -Value 'installed' -Status OK
                }
                else {
                    Write-MyStep -Label 'RSAT ADDS tools' -Value 'already installed' -Status OK
                }
            }
            catch {
                Write-MyError ('Failed to install RSAT ADDS tools: {0}' -f $_.Exception.Message)
                exit $ERR_PROBLEMADDINGFEATURE
            }
        }
        else {
            Write-MyVerbose 'Server OS detected, installing RSAT-ADDS via Install-WindowsFeature'
            try {
                if (-not (Get-WindowsFeature -Name 'RSAT-ADDS').Installed) {
                    Install-WindowsFeature -Name 'RSAT-ADDS' -ErrorAction Stop | Out-Null
                    Write-MyStep -Label 'RSAT-ADDS' -Value 'installed' -Status OK
                }
                else {
                    Write-MyStep -Label 'RSAT-ADDS' -Value 'already installed' -Status OK
                }
            }
            catch {
                Write-MyError ('Failed to install RSAT-ADDS feature: {0}' -f $_.Exception.Message)
                exit $ERR_PROBLEMADDINGFEATURE
            }
        }
    }

    function Install-RecipientManagement {
        # Phase 2 of Recipient Management install: run setup.exe /roles:ManagementTools + EMT permission script
        Write-MyVerbose 'Validating Exchange organization is reachable'
        if (-not (Test-ExchangeOrganization)) {
            Write-MyWarning 'Exchange organization not detected in Active Directory - installation may fail if AD was not prepared'
        }

        $setupExe = Join-Path $State['SourcePath'] 'setup.exe'
        if (-not (Test-Path $setupExe)) {
            Write-MyError ('Exchange setup.exe not found at {0}' -f $setupExe)
            exit $ERR_UNEXPTECTEDPHASE
        }

        Write-MyStep -Label 'Exchange Mgmt Tools setup' -Value '/roles:ManagementTools' -Status Run
        $rc = Invoke-Process -FilePath $State['SourcePath'] -FileName 'setup.exe' -ArgumentList '/mode:install /roles:ManagementTools /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF'
        if ($rc -ne 0) {
            Write-MyError ('Exchange setup returned exit code {0}' -f $rc)
            exit $ERR_UNEXPTECTEDPHASE
        }
        Write-MyStep -Label 'Exchange Mgmt Tools setup' -Value 'completed' -Status OK

        # Run CSS-Exchange Add-PermissionForEMT.ps1 if available (pre-stage in sources\).
        # This script was removed from CSS-Exchange releases; only runs if the file is pre-staged.
        $emtScript = Join-Path $State['SourcesPath'] 'Add-PermissionForEMT.ps1'
        $emtUrl = $null   # no longer available from CSS-Exchange releases
        if (Test-Path $emtScript) {
            try {
                Write-MyStep -Label 'EMT permissions' -Value 'running Add-PermissionForEMT.ps1' -Status Run
                & $emtScript
            }
            catch {
                Write-MyWarning ('Add-PermissionForEMT.ps1 execution failed: {0}' -f $_.Exception.Message)
            }
        }
    }

    function New-RecipientManagementShortcut {
        # Phase 3 of Recipient Management install: create desktop shortcut loading the RecipientManagement snapin
        try {
            $desktop = [Environment]::GetFolderPath('CommonDesktopDirectory')
            $shortcutPath = Join-Path $desktop 'Exchange Recipient Management.lnk'
            $shell = New-Object -ComObject WScript.Shell
            $shortcut = $shell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = (Get-Command powershell.exe).Source
            $shortcut.Arguments = '-NoExit -Command "Add-PSSnapin *RecipientManagement; Write-Host ''Recipient Management snap-in loaded'' -ForegroundColor Green"'
            $shortcut.IconLocation = '%SystemRoot%\System32\dsa.msc, 0'
            $shortcut.Description = 'Exchange Recipient Management PowerShell'
            $shortcut.Save()
            Write-MyStep -Label 'Desktop shortcut' -Value $shortcutPath -Status OK
        }
        catch {
            Write-MyWarning ('Could not create desktop shortcut: {0}' -f $_.Exception.Message)
        }
    }

    function Invoke-RecipientManagementADCleanup {
        # Optional AD cleanup after Recipient Management upgrade install
        Write-MyStep -Label 'RecipientMgmtCleanup' -Value 'reviewing legacy Exchange permissions' -Status Run
        Write-MyWarning 'AD cleanup is a manual safety gate. Review the following and run required Set-ADPermission commands manually if desired.'
        Write-MyVerbose 'Reference: https://learn.microsoft.com/en-us/exchange/plan-and-deploy/post-installation-tasks/post-installation-tasks'
    }

    function Install-ManagementToolsPrereqs {
        # Phase 1 of Management Tools install: Windows prerequisites
        Write-MyStep -Label 'Windows prerequisites' -Value 'installing for Exchange Mgmt Tools' -Status Run
        if (Test-IsClientOS) {
            Write-MyError 'Exchange Management Tools setup requires a Windows Server OS. Use -InstallRecipientManagement for client OS installs.'
            exit $ERR_UNEXPECTEDOS
        }
        $features = @('RSAT-ADDS', 'NET-Framework-45-Features')
        foreach ($f in $features) {
            if (-not (Get-WindowsFeature -Name $f -ErrorAction SilentlyContinue).Installed) {
                try {
                    Install-WindowsFeature -Name $f -ErrorAction Stop | Out-Null
                    Write-MyStep -Label 'Windows feature' -Value ('{0} installed' -f $f) -Status OK
                }
                catch {
                    Write-MyWarning ('Could not install {0}: {1}' -f $f, $_.Exception.Message)
                }
            }
        }
    }

    function Install-ManagementToolsRuntimePrereqs {
        # Phase 2 of Management Tools install: runtime prerequisites (VC++, URL Rewrite)
        Write-MyStep -Label 'Runtime prerequisites' -Value 'installing for Exchange Mgmt Tools' -Status Run
        # Management Tools only needs the baseline runtimes, not the full Exchange server stack.
        # Reuse existing VC++ helper functions where applicable (Install-MyPackage with the same IDs).
        Write-MyVerbose 'VC++ and URL Rewrite prerequisites are pulled in by setup.exe /roles:ManagementTools on demand'
    }

    function Install-ManagementToolsOnly {
        # Phase 3 of Management Tools install: run setup /roles:ManagementTools
        $setupExe = Join-Path $State['SourcePath'] 'setup.exe'
        if (-not (Test-Path $setupExe)) {
            Write-MyError ('Exchange setup.exe not found at {0}' -f $setupExe)
            exit $ERR_UNEXPTECTEDPHASE
        }
        Write-MyStep -Label 'Exchange Mgmt Tools setup' -Value '/roles:ManagementTools' -Status Run
        $rc = Invoke-Process -FilePath $State['SourcePath'] -FileName 'setup.exe' -ArgumentList '/mode:install /roles:ManagementTools /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF'
        if ($rc -ne 0) {
            Write-MyError ('Exchange setup returned exit code {0}' -f $rc)
            exit $ERR_UNEXPTECTEDPHASE
        }
        Write-MyStep -Label 'Exchange Mgmt Tools' -Value 'installed successfully' -Status OK
    }

