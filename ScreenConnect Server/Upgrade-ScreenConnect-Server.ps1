<# Upgrade-ScreenConnect-Server.ps1
Before running, confirm the $Downloads folder path exists and contains the latest ScreenConnect MSI installer.

Downloads are located at https://screenconnect.connectwise.com/download online. NOTE: ScreenConnect Client is NOT installed as part of this script!! This is for the server itself!

Run this PowerShell script from an elevated PowerShell prompt to upgrade your ScreenConnect server to the latest version. Will prompt you before proceeding, but otherwise will run silently.

Existing ScreenConnect server must already be installed. Version numbers will be verified and the upgrade will only be performed if the installer is for a newer version than the installation.

Install the PowerShell module "MSI" using "Install-Module MSI" if it is not already installed (you will a runtime error if it's not installed).

Use the switch -Force to install without prompting, assuming newer installer than installed version.

Version 0.0.1 - 2024-02-24 - Initial version by David Szpunar
#>
param(
    [switch] $Force
)
#Requires -Modules MSI

### CONFIG
# Define where the ScreenConnect MSI files are located, defaults to current user's Downloads folder.
# Only the most recently modified MSI file with name starting with ScreenConnect_ will be used.
$Downloads = "~\Downloads"
### END CONFIG

$InstallerPath = Resolve-Path -Path $Downloads
if(!$InstallerPath) {
    Write-Host "Unable to locate the $InstallerPath folder to look for installers. Quitting"
    exit 1
} else {
    Write-Host "Looking for installers in the $InstallerPath folder..."
}
$installer = (Get-ChildItem "~\Downloads\ScreenConnect_*.msi" -Attributes !Directory | Sort-Object -Descending -Property LastWriteTime | select -First 1)
if(!($installer)) {
    Write-Host "Unable to locate a ScreenConnect MSI installer to install."
    exit 1
} else {
    $InstallerFile = $installer.FullName
    Write-Host "Located the installer $InstallerFile, analyzing..."
}

$InstallerProps = $installer | get-msiproperty -passthru | select Name, MSIProperties
$InstallerVersion = $InstallerProps.productversion
$InstallerName = $InstallerProps.productname

if($InstallerName -ne 'ScreenConnect') {
    Write-Host "The $InstallerFile file is not a ScreenConnect installer! Quitting."
    exit 1
} else {
    Write-Host
}

Write-Host "Installer File:`t`t" $InstallerFile
Write-Host "File write time:`t" $installer.LastWriteTime
Write-Host "Installer Version:`t" $InstallerVersion -ForegroundColor Yellow
Write-Host

$svcPath = [System.IO.Path]::Combine(${env:ProgramFiles(x86)}, "ScreenConnect\Bin", "ScreenConnect.Service.exe")
if(!(Test-Path $svcPath)) {
    Write-Host "Unable to locate ScreenConnect Service executable file. Must already be installed. Quitting." -ForegroundColor Red
    exit 1
}
$svcVersion = (Get-Command $svcPath).FileVersionInfo.FileVersion
Write-Host "Installed Path:`t`t" $svcPath
Write-Host "Installed Version:`t" $svcVersion -ForegroundColor Yellow
Write-Host

$installedVer = [version]$svcVersion
$installerVer = [version]$InstallerVersion
if($installerVer -eq $installedVer) {
    Write-Host "The installer $($installer.Name) version is the SAME as current install! Quitting." -ForegroundColor Red
    exit 1
} elseif ($installerVer -lt $installedVer) {
    Write-Host "The installer $($installer.Name) is OLDER than current install! Quitting." -ForegroundColor Red
    exit 1
} elseif ($installerVer -gt $installedVer) {
    Write-Host "The installer $($installer.Name) is newer than current install. Proceeding..." -ForegroundColor Green
}

If ($Force -or (Read-Host "Silently install this MSI file? (Yes/No)") -Like "y*") {
    $InstallerLogFile = [io.path]::ChangeExtension([io.path]::GetTempFileName(), ".log")
    Write-Host "InstallerLogFile:`t" $InstallerLogFile
    $Arguments = " /c msiexec /i `"$InstallerFile`" /qn /l*v `"$InstallerLogFile`""
    $Process = Start-Process -Wait cmd -ArgumentList $Arguments -PassThru
    if ($Process.ExitCode -ne 0) {
        Get-Content $InstallerLogFile -ErrorAction SilentlyContinue | Select-Object -Last 200
        Write-Host "Upgrade failed, please troubleshoot manually. Log file: $InstallerLogFile" -ForegroundColor Red
    } else {
        Write-Host "Upgrade successfully completed. Please upgrade endpoint agents and adjust filter list version in application." -ForegroundColor Green
    }
} else {
    Write-Host "Cancelled, no action taken." -ForegroundColor Red
}
