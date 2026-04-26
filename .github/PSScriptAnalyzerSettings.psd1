@{
    # EXpress PSScriptAnalyzer configuration
    # Suppressed rules: deliberately chosen, not silenced to hide real issues.

    ExcludeRules = @(
        # Write-Host is intentional: Server Core has no WinForms; PS2Exe/.exe builds redirect
        # output; RDP/Hyper-V consoles and transcripts all work with Write-Host. Every other
        # output goes through Write-MyOutput / Write-MyWarning / Write-MyError.
        'PSAvoidUsingWriteHost',

        # Function names follow Exchange cmdlet conventions (Get-ExchangeServerObjects,
        # Install-AntispamAgents, etc.). Renaming would break callsites and operator familiarity.
        'PSUseSingularNouns',

        # ShouldProcess would add noise with no benefit: the script is a single-instance
        # installer that always runs under explicit admin consent; Exchange cmdlets themselves
        # do not surface -WhatIf/-Confirm in this pipeline.
        'PSUseShouldProcessForStateChangingFunctions'
    )

    Severity = @('Error', 'Warning')
}
