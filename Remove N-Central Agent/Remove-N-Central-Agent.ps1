<# Remove-N-Central-Agent.ps1

.SYNOPSIS
Remove N-able N-Central Windows Agent and related services, both via uninstallation of the agent and related apps (not the Probe) and also via manual cleanup (optional with the -Clean parameter) if the uninstall fails.

.DESCRIPTION
Uninstalls and optionally cleans/removes the N-able N-Central Windows Agent and related services and folders. It will attempt to run uninstallers for all known agent entries from Add/Remove Programs.

If the -Clean parameter is specified, it will also attempt (after uninstallation) to clean up agent remnants in addition to attempting uninstallation--this includes stopping and disabling related services, removing both the Uninstall and regular application registry entries and disabling the services, attempting to delete the services, and then removing the installation folders and the related folders in C:\ProgramData.

The script should NOT remove N-able Cove backup installations or files/services.

The script should uninstall (but does not clean up remants of) the Windows Probe service or agent, if it exists.

The script does NOT remove the registry setting showing the N-central device ID that would be used if the agent were ever reinstalled to map to the same device in the future. This ID should be a harmless artifact in most cases especially if the agent is no longer running on the system, but it could be updated to remove it if desired.

The script may leave some top-level folders under Program Files, that are empty, or may leave some subfolders that are in use and cannot be removed due to permissions, but it will attempt to remove some of these after a reboot if possible, and any remaining remnants afterwards should not allow the agent to run, even if they aren't completely cleaned up.

The script WILL delete agent logs and history in ProgramData subfolders.

The script WILL attempt to remove the Take Control (BeAnywhere) services and folders if they exist and -Clean is specified, but does not attempt any separate uninstallation.

The script should be run with admin privileges, ideally as SYSTEM, and will quit if it is not.

Paths and services to clean up are hardcoded into the script under CONFIG AND SETUP, and will use the correct system drive for the system but the rest of the paths are hardcoded. The installation folders that any existing services refer to will be added to the cleanup list, if they are different and exist during the run (if -Clean is run as part of the initial pass).

While it can be run manually, it is recommended that the script be run via a different RMM tool, and supports but does not require NinjaRMM Script Variables with the parameter names (as checkboxes) for configuration.

Service deletion issue with version 0.0.1 and 0.0.2 has been resolved, the script properly deletes services during cleanup, in addition to stopping and disabling.

.PARAMETER Clean
Clean up agent remnants in addition to attempting uninstallation.

.PARAMETER TestOnly
Test removal of agent and services without actually removing them--will output test info to console instead of making changes. Kind of like a custom -WhatIf dry run without being official.

.EXAMPLE
Remove-N-Central-Agent.ps1

.EXAMPLE
Remove-N-Central-Agent.ps1 -Clean

.NOTES
Version 0.0.3 - 2024-03-26 by David Szpunar - Resolution of service deletion bug in cleanup
Version 0.0.2 - 2024-03-26 by David Szpunar - Update service deletion options
Version 0.0.1 - 2024-03-25 by David Szpunar - Initial release
#>
[CmdletBinding()]
param(
    [switch] $Clean,
    [switch] $TestOnly
)

### PROCESS NINJRAMM SCRIPT VARIABLES AND ASSIGN TO NAMED SWITCH PARAMETERS
# Get all named parameters and overwrite with any matching Script Variables with value of 'true' from environment variables
# Otherwise, if not a checkbox ('true' string), assign any other Script Variables provided to matching named parameters
$switchParameters = (Get-Command -Name $MyInvocation.InvocationName).Parameters
foreach ($param in $switchParameters.keys) {
    $var = Get-Variable -Name $param -ErrorAction SilentlyContinue
    if ($var) {
        $envVarName = $var.Name.ToLower()
        $envVarValue = [System.Environment]::GetEnvironmentVariable("$envVarName")
        if (![string]::IsNullOrWhiteSpace($envVarValue) -and ![string]::IsNullOrEmpty($envVarValue) -and $envVarValue.ToLower() -eq 'true') {
            # Checkbox variables
            $PSBoundParameters[$envVarName] = $true
            Set-Variable -Name "$envVarName" -Value $true -Scope Script
        }
        elseif (![string]::IsNullOrWhiteSpace($envVarValue) -and ![string]::IsNullOrEmpty($envVarValue) -and $envVarValue -ne 'false') {
            # non-Checkbox string variables
            $PSBoundParameters[$envVarName] = $envVarValue
            Set-Variable -Name "$envVarName" -Value $envVarValue -Scope Script
        }
    }
}
### END PROCESS SCRIPT VARIABLES

##### CONFIG AND SETUP #####
# These itemss should generally be set via parameters or environment/script variables, but can be manually overridden for testing:
# $TestOnly = $false
# $Clean = $true
# $Verbose = $true

<# Some of this information was used for interactive troubleshooting and script design but is not a part of the final script, left for reference:

# $Application = "N-able"
# # $AgentInstall = ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall") | ForEach-Object { Get-ChildItem -Path $_ | ForEach-Object { Get-ItemProperty $_.PSPath } | Where-Object { $_.DisplayName -match "$Application" } }
# $AgentInstall = ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall") | ForEach-Object { Get-ChildItem -Path $_ | ForEach-Object { Get-ItemProperty $_.PSPath } | Where-Object { $_.Publisher -match "$Application" } }
# $AgentVersion = $AgentInstall.DisplayVersion
# $AgentGUID = $AgentInstall.PSChildName
# $AgentInstall | Format-List
#>

<#
PREPARE THE LIST OF APPS TO UNINSTALL AND SERVICES TO REMOVE
#>
$AppList = @('Ecosystem Agent', 'Patch Management Service Controller', 'Request Handler Agent', 'File Cache Service Agent')
$MSIList = @('Windows Agent', 'Windows Probe')

$ServiceList = @('AutomationManagerAgent', 'EcosystemAgent', 'EcosystemAgentMaintenance', 'Windows Agent Service', 'Windows Agent Maintenance Service', 'PME.Agent.PmeService', 'BASupportExpressSrvcUpdater_N_Central', 'BASupportExpressStandaloneService_N_Central')

# Resolve-Path "$($env:systemdrive)\Program Files*\MspPlatform"
$FolderPathsList = @("$($env:systemdrive)\Program Files*\MspPlatform", 
    "$($env:systemdrive)\Program Files*\SolarWinds MSP", 
    "$($env:systemdrive)\Program Files*\MSPEcosystem", 
    "$($env:systemdrive)\Program Files*\BeAnywhere Support Express",
    "$($env:systemdrive)\ProgramData\MspPlatform", 
    "$($env:systemdrive)\ProgramData\N-able Technologies\AutomationManager",
    "$($env:systemdrive)\ProgramData\N-able Technologies\AVDefender",
    "$($env:systemdrive)\ProgramData\N-able Technologies\Windows Agent",
    "$($env:systemdrive)\ProgramData\N-able Technologies\N-able\AutomationManager",
    "$($env:systemdrive)\ProgramData\N-able\AutomationManager",
    "$($env:systemdrive)\ProgramData\Solarwinds MSP",
    "$($env:systemdrive)\ProgramData\GetSupportService",
    "$($env:systemdrive)\ProgramData\GetSupportService_Common",
    "$($env:systemdrive)\ProgramData\GetSupportService_Common_N-central",
    "$($env:systemdrive)\ProgramData\GetSupportService_N-central",
    "$($env:systemdrive)\ProgramData\N-able Technologies"
    )

<#
PREPARE THE EMPTY LIST OF FILE PATHS TO LATER REMOVE
#>
$InstallPaths = New-Object System.Collections.Generic.List[System.Object]

<#
PREPARE THE EMPTY LIST OF FILE PATHS TO LATER REMOVE
#>
$ServicePaths = New-Object System.Collections.Generic.List[System.Object]

<#
PREPARE THE EMPTY LIST OF REGISTRY PATHS TO LATER REMOVE
#>
$RegistryPaths = New-Object System.Collections.Generic.List[System.Object]

# $RegistryPaths.Add('HKLM:\SOFTWARE\N-able Technologies')
# $RegistryPaths.Add('HKLM:\SOFTWARE\WOW6432Node\N-able')
$RegistryPaths.Add('HKLM:\SOFTWARE\WOW6432Node\N-able\AM')
$RegistryPaths.Add('HKLM:\SOFTWARE\N-able\AM')
# $RegistryPaths.Add('HKLM:\SOFTWARE\WOW6432Node\N-able Technologies')
$RegistryPaths.Add('HKLM:\SOFTWARE\WOW6432Node\N-able Technologies\Windows Agent')
$RegistryPaths.Add('HKLM:\SOFTWARE\N-able Technologies\Windows Agent')

###### FUNCTIONS ######
function Test-IsElevated {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object System.Security.Principal.WindowsPrincipal($id)
    $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

Function Remove-ItemOnReboot {
    # SOURCE: https://gist.github.com/rob89m/6bbea14651396f5870b23f1b2b8e4d0d
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)][string]$Item
    )
    END {
        # Read current items from PendingFileRenameOperations in Registry
        $PendingFileRenameOperations = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations).PendingFileRenameOperations
     
        # Append new item to be deleted to variable
        $NewPendingFileRenameOperations = $PendingFileRenameOperations + "\??\$Item"
     
        # Reload PendingFileRenameOperations with existing values plus newly defined item to delete on reboot
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name "PendingFileRenameOperations" -Value $NewPendingFileRenameOperations
    }
}

function Uninstall-GUID ([string]$GUID) {
    # $AgentInstall = ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall") | ForEach-Object { Get-ChildItem -Path $_ | ForEach-Object { Get-ItemProperty $_.PSPath } | Where-Object { $_.DisplayName -match "$Application" } }
    # $AgentVersion = $AgentInstall.DisplayVersion
    # $AgentGUID = $AgentInstall.PSChildName
    # 
    # $UninstallString = "$($AgentInstall.UninstallString) /quiet /norestart"
    $UninstallString = "MsiExec.exe /x$AgentGUID /quiet /norestart"

    if ($GUID) {
        Write-Host "Uninstalling now via GUID: $GUID"
        Write-Verbose "Uninstall String: $UninstallString"
        if (!$TestOnly) {
            $Process = Start-Process cmd.exe -ArgumentList "/c $UninstallString" -Wait -PassThru
            if ($Process.ExitCode -eq 1603) {
                Write-Host "Uninstallation attempt failed with error code: $($Process.ExitCode). Please review manually."
                Write-Host "Hint: This exit code likely requires the system to reboot prior to installation."
            }
            elseif ($Process.ExitCode -ne 0) {
                Write-Host "Uninstallation attempt failed with error code: $($Process.ExitCode). Please review manually."
            }
            else {
                Write-Host "Uninstallation attempt completed."
            }
            return $($Process.ExitCode)
        }
        else {
            Write-Host "TEST ONLY: No uninstallation attempt was made."
            return 0
        }
    }
    else {
        Write-Host "Pass a GUID to the function."
        return $false
    }
}

function Uninstall-App ($Agent) {
    $QuietUninstall = $Agent.QuietUninstallString
    if ([string]::IsNullOrWhiteSpace($QuietUninstall)) {
        Write-Host "No QuietUninstall string, skipping silent uninstall for" $Agent.DisplayName
        return $false
    }
    $UninstallString = $Agent.UninstallString + " /SILENT /VERYSILENT /SUPPRESSMSGBOXES"

    if ($Agent) {
        Write-Host "Uninstalling now via Inno Setup Silent Removal:" $Agent.DisplayName
        Write-Host "Uninstall String: $UninstallString"
        if (!$TestOnly) {
            $Process = Start-Process cmd.exe -ArgumentList "/c $UninstallString" -Wait -PassThru
            if ($Process.ExitCode -ne 0) {
                Write-Host "Uninstallation attempt failed with error code: $($Process.ExitCode). Please review manually."
            }
            else {
                Write-Host "Uninstallation attempt completed for" $Agent.DisplayName
                return 0
            }
            return $($Process.ExitCode)
        }
        else {
            Write-Host "TEST ONLY: No uninstallation attempt was made."
            return 0
        }
    }
    else {
        Write-Host "Pass an uninstall registry object to the function."
        return 1
    }
}



Function Get-ServiceStatus ([string]$Name) {
    (Get-Service -Name $Name -ErrorAction SilentlyContinue).Status
}

Function Stop-RunningService ($svc) {
    Write-Verbose "Checking if $($svc.Name) service is running to STOP"
    # If ( $(Get-ServiceStatus -Name $Name) -eq "Running" ) {
    If ( $svc.Status -eq "Running" ) {
        Write-Host "Stopping : $($svc.Name) service"
        if (!$TestOnly) {
            # Stop-Service -Name $Svc -Force
            $svc | Stop-Service -Force
        }
        else {
            Write-Host "TEST ONLY: Not stopping $Name service"
        }
    }
    else {
        Write-Verbose "The $($svc.Name) service is not running, not stopping it!"
    }
}

Function Disable-Service ($svc) {
    If ( $svc ) {
        Write-Host "Disabling : $($svc.Name) service"
        if (!$TestOnly) {
            # Set-Service $Svc -StartupType Disabled
            $svc | Set-Service -StartupType Disabled
        }
        else {
            Write-Host "TEST ONLY: Not disabling $($svc.Name) service"
        }
    }
    else {
        Write-Verbose "The $($svc.Name) service doesn't exist, not disabling it!"
    }
}

Function Remove-StoppedService ($svc) {
    If ( $svc ) {
        If ( $svc.Status -eq "Stopped" ) {
            Write-Host "Deleting : $($svc.Name) service"
            if (!$TestOnly) {
                Stop-Process -Name $($svc.Name) -Force -ErrorAction SilentlyContinue
                sc.exe delete $($svc.Name)
                Remove-Item "HKLM:\SYSTEM\CurrentControlSet\Services\$($svc.Name)" -Force -Recurse -ErrorAction SilentlyContinue
            }
            else {
                Write-Host "TEST ONLY: Not deleting $Name service"
            }
        }
        else {
            Write-Host "The $($svc.Name) service is not stopped, not deleting it!"
        }
    }
    Else {
        Write-Verbose "Not Found to Remove: $($svc.Name) service"
    }
}

Function Remove-File-Path ([string]$Path) {
    Write-Host "Deleting folder if it exists: $Path"
    $FolderPath = Resolve-Path -Path $Path.Trim('"') -ErrorAction SilentlyContinue
    if (![string]::IsNullOrEmpty($FolderPath) -and (Test-Path $FolderPath)) {
        Write-Host "Removing folder: $FolderPath"
        if (!$TestOnly) {
            try {
                Remove-Item -Path $FolderPath -Recurse -Force -ErrorAction Stop
            }
            catch {
                Write-host "Error deleting folder '$FolderPath\', adding to delete on reboot list."
                Remove-ItemOnReboot -Item "$FolderPath\"
            }
        }
        else {
            Write-Host "TEST ONLY: Not removing $FolderPath"
        }
    }
    else {
        Write-Verbose "Not found and thus not deleting: $Path"
    }
}

Function Remove-Registry-Path ([string]$Path) {
    Write-Verbose "Deleting registry path: $Path"
    $KeyPath = Resolve-Path $Path -ErrorAction SilentlyContinue
    if (![string]::IsNullOrEmpty($KeyPath) -and (Test-Path $KeyPath)) {
        Write-Host "Removing key $KeyPath"
        if (!$TestOnly) {
            Remove-Item -Path $KeyPath -Recurse -Force
        }
        else {
            Write-Host "TEST ONLY: Not removing $KeyPath"
        }
    }
    else {
        Write-Verbose "Not found and thus not deleting: ${Path}"
    }
}

function Remove-Agent {
    foreach ($service in $ServiceList) {
        Write-Host "`nGetting Service $service"
        $ServiceObj = Get-Service $service -ErrorAction SilentlyContinue
        if (($ServiceObj)) {
            $SvcInfo = Get-WmiObject win32_service | Where-Object { $_.Name -eq "$service" } | Select-Object Name, DisplayName, State, StartMode, PathName
            Write-Host "STATE: $($SvcInfo.State) MODE: $($SvcInfo.StartMode) `tSERVICE: $($SvcInfo.DisplayName) '$($SvcInfo.Name)'"
            Write-Host "PATH: $($SvcInfo.PathName)"

            $SvcPath = Split-Path -Path $($SvcInfo.PathName).Trim('"') -Parent
            $ServicePaths.Add($SvcPath)

            Stop-RunningService $ServiceObj
            Disable-Service $ServiceObj
            $ServiceObj = Get-Service $service -ErrorAction SilentlyContinue
            Remove-StoppedService $ServiceObj
        }
        else {
            Write-Host "Service $service not found."
        }
    }

    Write-Host

    foreach ($Folder in $FolderPathsList) {
        Remove-File-Path $Folder
    }

    foreach ($Folder in $InstallPaths) {
        Remove-File-Path $Folder
    }

    foreach ($Folder in $ServicePaths) {
        Remove-File-Path $Folder
    }

    foreach ($Key in $RegistryPaths) {
        Remove-Registry-Path $Key
    }
}

##### BEGIN SCRIPT #####

# If not elevated error out. Admin priveledges are required to uninstall software
if (-not (Test-IsElevated)) {
    Write-Error -Message "Access Denied. Please run with Administrator privileges."
    exit 1
}

foreach ($app in $MSIList) {
    $AgentInstall = ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall") | ForEach-Object { Get-ChildItem -Path $_ | ForEach-Object { Get-ItemProperty $_.PSPath } | Where-Object { $_.DisplayName -match "$app" } }
    $AgentGUID = $AgentInstall.PSChildName

    if ($AgentInstall) {
        $InstallPaths.Add($AgentInstall.InstallLocation)
        $RegistryPaths.Add($AgentInstall.PSPath)
        Write-Host "`nUninstalling app '$app' using GUID: " $AgentGUID
        if ((Uninstall-GUID $AgentGUID) -eq 0) {
            Write-Host "Successfully uninstalled '$app' via MSI command using GUID $AgentGUID"
        }
    }
    else {
        Write-Verbose "No installation entry found to uninstall '$app' via MSI command."
    }
}

foreach ($app in $AppList) {
    $AgentInstall = ("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall") | ForEach-Object { Get-ChildItem -Path $_ | ForEach-Object { Get-ItemProperty $_.PSPath } | Where-Object { $_.DisplayName -match "$app" } }
    $AgentGUID = $AgentInstall.PSChildName

    if ($AgentInstall) {
        $InstallPaths.Add($AgentInstall.InstallLocation)
        $RegistryPaths.Add($AgentInstall.PSPath)
        Write-Host "`nUninstalling app '$app' using Inno Setup Quiet Removal."
        if ((Uninstall-App $AgentInstall) -eq 0) {
            Write-Host "Successfully uninstalled '$app' silently via Inno Setup Quiet Removal."
        }
        else {
            Write-Host "Unable to uninstall '$app' silently via Inno Setup Quiet Removal."
        }
    }
    else {
        Write-Verbose "No installation entry found to uninstall '$app' silently via Inno Setup command."
    }
}

Write-Host "Folder Paths:"
$InstallPaths
Write-Host "Registry Paths:"
$RegistryPaths


if ($Clean) {
    Write-Host "`nAttempting to clean up agent remnants..."
    Remove-Agent $AgentInstall

    Write-Host "`nService Folder Paths after Agent Removal:"
    $ServicePaths
}
