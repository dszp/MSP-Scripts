<# Disable-PowerShell-V2.ps1
Check if PowerShell v2 is enabled disable it if it is.

Security report with example disable command: https://www.stigviewer.com/stig/windows_10/2017-04-28/finding/V-70637

Version 1.0.0 - 2023-10-18 - Initial Version

USAGE: Run script to report and disable PowerShell v2. No change will be made if it's already disabled.
In order to provide an undo option, pass the argument -Enable to instead enable PowerShell v2, if disabled.
Returns 0 if the requested action was successful or no action was taken, or 1 if a requested change fails to 
succeed after testing. Will report if a reboot is required to complete the change (it usually isn't).
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)][switch] $Enable
)
if($env:enable -eq $true) {
    $Enable = $true
}

function Disable-Or-Enable-PowerShellV2 {
    param(
        [Parameter(Mandatory=$false)][bool] $Enable
    )
    $feature = $false
    [bool]$exists = $false
    $features = Get-WindowsOptionalFeature -Online
    foreach($feat in $features) {
        if($feat.FeatureName -eq "MicrosoftWindowsPowerShellV2Root") {
            $exists = $true
            $feature = Get-WindowsOptionalFeature -Online -FeatureName MicrosoftWindowsPowerShellV2Root
        }
    }
    if(!$exists) {
        Write-Host "The Windows feature 'MicrosoftWindowsPowerShellV2Root' does not exist. This OS does not support PowerShell V2. Quitting."
        return 0
    }

    if($feature.State -ne "Disabled" -and !$Enable) {
        Write-Host "PowerShell v2 is currently ENABLED. Disabling..."
        $change = Disable-WindowsOptionalFeature -Online -FeatureName MicrosoftWindowsPowerShellV2Root
        if($change.RestartNeeded -eq $true) {
            Write-Host "PowerShell v2 requires a restart to complete disabling. Please restart the system."
        }
        $feature = Get-WindowsOptionalFeature -Online -FeatureName MicrosoftWindowsPowerShellV2Root
        if($feature.State -ne "Disabled") {
            Write-Host "PowerShell v2 was NOT DISABLED even after attempting the change. Something went wrong."
            return 1
        } else {
            Write-Host "PowerShell v2 was DISABLED successfully."
            return 0
        }
    } elseif ($feature.State -ne "Enabled" -and $Enable) {
        Write-Host "PowerShell v2 is currently DISABLED. Enabling..."
        $change = Enable-WindowsOptionalFeature -Online -FeatureName MicrosoftWindowsPowerShellV2Root -NoRestart
        if($change.RestartNeeded -eq $true) {
            Write-Host "PowerShell v2 requires a restart to complete enabling. Please restart the system."
        }
        $feature = Get-WindowsOptionalFeature -Online -FeatureName MicrosoftWindowsPowerShellV2Root
        if($feature.State -ne "Enabled") {
            Write-Host "PowerShell v2 was NOT ENABLED even after attempting the change. Something went wrong."
            return 1
        } else {
            Write-Host "PowerShell v2 was ENABLED successfully."
            return 0
        }
    } elseif ($feature.State -eq "Enabled" -and $Enable) {
        Write-Host "PowerShell v2 is already ENABLED."
        return 0
    } else {
        Write-Host "PowerShell v2 is already DISABLED."
        return 0
    }
}


exit (Disable-Or-Enable-PowerShellV2 $Enable)
