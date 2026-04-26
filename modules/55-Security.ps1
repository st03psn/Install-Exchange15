    function Invoke-EOMT {
        if (-not $State['RunEOMT']) {
            Write-MyVerbose 'RunEOMT not specified, skipping EOMT'
            return
        }
        Write-MyStep -Label 'EOMT' -Value 'running (CSS-Exchange Emergency Mitigation)' -Status Run
        $eomtPath = Join-Path $State['SourcesPath'] 'EOMT.ps1'
        $eomtUrl  = 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/EOMT.ps1'

        if (-not (Test-Path $eomtPath)) {
            $downloaded = $false
            $savedPP = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                try {
                    Write-MyVerbose ('Downloading EOMT from {0} (attempt {1}/3)' -f $eomtUrl, $attempt)
                    Start-BitsTransfer -Source $eomtUrl -Destination $eomtPath -ErrorAction Stop
                    $downloaded = $true
                    break
                }
                catch {
                    Get-BitsTransfer -ErrorAction SilentlyContinue | Where-Object { $_.JobState -notin 'Transferred','Acknowledged' } | Remove-BitsTransfer -ErrorAction SilentlyContinue
                    Remove-Item -Path $eomtPath -ErrorAction SilentlyContinue
                    if ($attempt -eq 3) {
                        try {
                            Invoke-WebDownload -Uri $eomtUrl -OutFile $eomtPath
                            $downloaded = $true
                        }
                        catch {
                            Write-MyWarning ('Could not download EOMT after 3 attempts: {0}' -f $_.Exception.Message)
                        }
                    }
                    else {
                        Start-Sleep -Seconds ($attempt * 5)
                    }
                }
            }
            $ProgressPreference = $savedPP
            if (-not $downloaded) { return }
        }

        if (Test-Path $eomtPath) {
            try {
                Write-MyVerbose ('EOMT SHA256: {0}' -f (Get-FileHash -Path $eomtPath -Algorithm SHA256).Hash)
                & $eomtPath
                Write-MyStep -Label 'EOMT' -Value 'completed successfully' -Status OK
            }
            catch {
                Write-MyWarning ('EOMT execution failed: {0}' -f $_.Exception.Message)
            }
        }
    }

    function Set-HSTSHeader {
        if ($State['InstallEdge']) {
            Write-MyVerbose 'Edge role has no OWA/ECP — skipping HSTS configuration'
            return
        }
        Write-MyStep -Label 'HSTS' -Value 'configuring on OWA + ECP' -Status Run
        try {
            Import-Module WebAdministration -ErrorAction Stop
            $site = 'IIS:\Sites\Default Web Site'
            foreach ($vDir in @('owa', 'ecp')) {
                $path = '{0}\{1}' -f $site, $vDir
                if (-not (Test-Path $path)) {
                    Write-MyVerbose ('Virtual directory /{0} not found in IIS, skipping HSTS' -f $vDir)
                    continue
                }
                $filter   = 'system.webServer/httpProtocol/customHeaders/add[@name="Strict-Transport-Security"]'
                $existing = Get-WebConfigurationProperty -PSPath $path -Filter $filter -Name '.' -ErrorAction SilentlyContinue
                if ($existing) {
                    Write-MyVerbose ('HSTS header already present on /{0}' -f $vDir)
                }
                else {
                    Add-WebConfigurationProperty -PSPath $path -Filter 'system.webServer/httpProtocol/customHeaders' -Name '.' -Value @{ name = 'Strict-Transport-Security'; value = 'max-age=31536000' }
                    Write-MyStep -Label ('HSTS /{0}' -f $vDir) -Value 'configured (max-age=31536000)' -Status OK
                }
            }
        }
        catch {
            Write-MyWarning ('Failed to configure HSTS: {0}' -f $_.Exception.Message)
        }
    }

    function Import-ExchangeCertificateFromPFX {
        if (-not $State['CertificatePath'] -or -not $State['CertificatePassword']) {
            Write-MyVerbose 'No certificate import requested'
            return
        }

        $pfxPath = $State['CertificatePath']
        if (-not (Test-Path $pfxPath)) {
            Write-MyError ('PFX file not found: {0}' -f $pfxPath)
            return
        }

        Write-MyStep -Label 'PFX certificate' -Value ('importing from {0}' -f $pfxPath) -Status Run
        try {
            $secPwd = ConvertTo-SecureString $State['CertificatePassword']
            Register-ExecutedCommand -Category 'Certificate' -Command ("Import-ExchangeCertificate -FileData ([IO.File]::ReadAllBytes('$pfxPath')) -Password <SecureString> -PrivateKeyExportable `$true")
            $cert = Import-ExchangeCertificate -FileData ([System.IO.File]::ReadAllBytes($pfxPath)) -Password $secPwd -PrivateKeyExportable $true -ErrorAction Stop
            Write-MyStep -Label 'Certificate' -Value ('imported: {0}' -f $cert.Subject) -Status OK

            # Detect wildcard certificate (CN=* or SAN with *.domain)
            $isWildcard = ($cert.Subject -match 'CN=\*') -or ($cert.SubjectAlternativeNames -match '^\*\.')
            if ($isWildcard) {
                # Wildcard: enable for IIS and SMTP only (IMAP/POP use specific SANs)
                Register-ExecutedCommand -Category 'Certificate' -Command ("Enable-ExchangeCertificate -Thumbprint '$($cert.Thumbprint)' -Services IIS,SMTP -Force")
                Enable-ExchangeCertificate -Thumbprint $cert.Thumbprint -Services IIS,SMTP -Force -ErrorAction Stop
                Write-MyStep -Label 'Certificate (wildcard)' -Value 'enabled (IIS, SMTP)' -Status OK
            }
            else {
                # Named certificate: also enable for IMAP and POP
                Register-ExecutedCommand -Category 'Certificate' -Command ("Enable-ExchangeCertificate -Thumbprint '$($cert.Thumbprint)' -Services IIS,SMTP,IMAP,POP -Force")
                Enable-ExchangeCertificate -Thumbprint $cert.Thumbprint -Services IIS,SMTP,IMAP,POP -Force -ErrorAction Stop
                Write-MyStep -Label 'Certificate' -Value 'enabled (IIS, SMTP, IMAP, POP)' -Status OK
            }
        }
        catch {
            Write-MyError ('Failed to import/enable certificate: {0}' -f $_.Exception.Message)
        }
    }

