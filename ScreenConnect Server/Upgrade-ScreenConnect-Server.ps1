<# Upgrade-ScreenConnect-Server.ps1
Run this PowerShell script from an elevated PowerShell prompt to upgrade your ScreenConnect server to the latest version.

Before running, confirm the $Downloads folder path exists (configured below) and optionally contains the latest ScreenConnect installer.

The script will first check the online download page for ScreenConnect and attempt to verify the newest verison. If there is a newer version available 
to download, it will prompt you to download the latest version. If you choose not to, it will continue with the latest installer in the $Downloads folder, 
if one already exists.

Note that the download page parsing may or may not be accurate and has only been tested with one version avialble, for Windows. If parsing fails, 
please report the issue to the author. You should be able to install using the latest installer in the $Downloads folder regardless, if one exists.

Use the -SkiDownload parameter to skip the download prompt or cloud version check and only use locally downloaded installers, if any.

Downloads are located at https://www.screenconnect.com/Download online. NOTE: ScreenConnect Client is NOT installed as part of this script!! This is for the server itself!

Run this PowerShell script from an elevated PowerShell prompt to upgrade your ScreenConnect server to the latest version. Will prompt you before proceeding, but otherwise will run silently.

Existing ScreenConnect server must already be installed. Version numbers will be verified and the upgrade will only be performed if the installer is for a newer version than the installation.

Install the PowerShell module "MSI" using "Install-Module MSI" if it is not already installed (you will a runtime error if it's not installed).

Use the switch -Force to download (if necessary) and install without prompting, assuming newer installer than installed version.

Version 0.0.1 - 2024-02-24 - Initial version by David Szpunar
Version 0.0.2 - 2025-05-01 - Updated to incorporate version checking and download functionality directly in this script.
#>
param(
    [switch] $Force,
    [switch] $SkipDownload
)
#Requires -Modules MSI

### CONFIG
# Define where the ScreenConnect MSI files are located, defaults to current user's Downloads folder.
# Only the most recently modified MSI file with name starting with ScreenConnect_ will be used.
$Downloads = "~\Downloads"
### END CONFIG

#region Functions
function Get-LatestScreenConnectVersion {
    [CmdletBinding()]
    param()

    try {
        # Download the ScreenConnect download page
        $url = "https://www.screenconnect.com/Download"
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing
        $content = $response.Content

        # Look for the download URL of the stable release version
        $downloadUrlPattern = 'https://[^"]+ScreenConnect_[\d\.]+_Release\.msi'
        if ($content -match $downloadUrlPattern) {
            $downloadUrl = $matches[0]
            $fileName = $downloadUrl.Split('/')[-1]

            # Extract version from filename
            $version = "Unknown"
            if ($fileName -match 'ScreenConnect_(\d+\.\d+\.\d+\.\d+)') {
                $version = $matches[1]
            }

            # Extract information directly using the HTML structure provided by the user
            # Look for the row containing this URL
            $rowPattern = '<a href="' + [regex]::Escape($downloadUrl) + '"[^>]*>(.*?)</a>'
            if ($content -match $rowPattern) {
                $rowContent = $matches[1]
                
                # Extract the column values using the known structure
                $sizePattern = '<div class="\s*column index-1\s*">\s*(.*?)\s*</div>'
                $releaseDatePattern = '<div class="\s*column index-2\s*">\s*(.*?)\s*</div>'
                $fileNamePattern = '<div class="\s*column index-3\s*">\s*(.*?)\s*</div>'
                $platformsPattern = '<div class="\s*column index-4\s*">\s*(.*?)\s*</div>'
                
                $size = if ($rowContent -match $sizePattern) { $matches[1].Trim() } else { "Unknown" }
                $releaseDate = if ($rowContent -match $releaseDatePattern) { $matches[1].Trim() } else { "Unknown" }
                $fileNameFromHtml = if ($rowContent -match $fileNamePattern) { $matches[1].Trim() } else { "Unknown" }
                $platforms = if ($rowContent -match $platformsPattern) { $matches[1].Trim() } else { "Unknown" }
            }
            else {
                # If we can't find the row, try a different approach
                # Use the exact HTML structure from the user's example
                $tableRowPattern = '<div class="row-item">\s*<div class="\s*column index-1\s*">\s*(.*?)\s*</div>\s*<div class="\s*column index-2\s*">\s*(.*?)\s*</div>\s*<div class="\s*column index-3\s*">\s*(.*?)\s*</div>\s*<div class="\s*column index-4\s*">\s*(.*?)\s*</div>'
                
                if ($content -match $tableRowPattern) {
                    $size = $matches[1].Trim()
                    $releaseDate = $matches[2].Trim()
                    $fileNameFromHtml = $matches[3].Trim()
                    $platforms = $matches[4].Trim()
                }
                else {
                    # Last resort - try to find values directly
                    $size = "Unknown"
                    $releaseDate = "Unknown"
                    $platforms = "Unknown"
                    
                    # Look for a pattern like "123 MB" near the file name
                    if ($content -match '(\d+)\s*MB[^<]*' + [regex]::Escape($fileName)) {
                        $size = $matches[1] + " MB"
                    }
                    
                    # Look for a date pattern near the file name
                    if ($content -match '(\d{1,2}/\d{1,2}/\d{4})[^<]*' + [regex]::Escape($fileName)) {
                        $releaseDate = $matches[1]
                    }
                    
                    # Look for platform information near the file name
                    if ($content -match [regex]::Escape($fileName) + '[^<]*(Windows|Linux|Mac)') {
                        $platforms = $matches[1]
                    }
                }
            }

            # Create and return a custom object with the extracted information
            [PSCustomObject]@{
                Size        = $size
                ReleaseDate = $releaseDate
                FileName    = $fileName
                Platforms   = $platforms
                Version     = $version
                DownloadUrl = $downloadUrl
            }
        }
        else {
            Write-Error "Could not find stable release download link on the page."
            return $null
        }
    }
    catch {
        Write-Error "An error occurred while retrieving ScreenConnect version information: $_"
        return $null
    }
}

function Compare-ScreenConnectVersions {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$CurrentVersion,
        
        [Parameter(Mandatory = $true)]
        [string]$LatestVersion
    )

    try {
        # Parse versions into version objects for proper comparison
        $currentVersionObj = [System.Version]::new($CurrentVersion)
        $latestVersionObj = [System.Version]::new($LatestVersion)

        # Compare versions
        $comparisonResult = $currentVersionObj.CompareTo($latestVersionObj)

        # Return comparison result
        [PSCustomObject]@{
            CurrentVersion        = $CurrentVersion
            LatestVersion         = $LatestVersion
            IsUpdateAvailable     = $comparisonResult -lt 0
            IsCurrentVersionNewer = $comparisonResult -gt 0
            AreVersionsEqual      = $comparisonResult -eq 0
        }
    }
    catch {
        Write-Error "An error occurred while comparing versions: $_"
        return $null
    }
}

function Download-ScreenConnectFile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Url,
        
        [Parameter(Mandatory = $true)]
        [string]$DestinationFolder,
        
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )
    
    try {
        # Ensure the destination folder exists
        if (-not (Test-Path -Path $DestinationFolder)) {
            New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null
            Write-Host "Created download folder: $DestinationFolder" -ForegroundColor Cyan
        }
        
        $destinationPath = Join-Path -Path $DestinationFolder -ChildPath $FileName
        
        # Download the file
        Write-Host "Downloading file from $Url" -ForegroundColor Cyan
        Write-Host "Saving to $destinationPath" -ForegroundColor Cyan
        
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($Url, $destinationPath)
        
        Write-Host "Download completed successfully." -ForegroundColor Green
        return $destinationPath
    }
    catch {
        Write-Error "Failed to download file: $_"
        return $null
    }
}
#endregion Functions

# Get the path to the currently installed ScreenConnect service
$svcPath = [System.IO.Path]::Combine(${env:ProgramFiles(x86)}, "ScreenConnect\Bin", "ScreenConnect.Service.exe")
if (!(Test-Path $svcPath)) {
    Write-Host "Unable to locate ScreenConnect Service executable file. Must already be installed. Quitting." -ForegroundColor Red
    exit 1
}

# Get the currently installed version
$svcVersion = (Get-Command $svcPath).FileVersionInfo.FileVersion
Write-Host "Installed Path:`t`t" $svcPath
Write-Host "Installed Version:`t" $svcVersion -ForegroundColor Yellow
Write-Host

$latestVersionInfo = ''
if ($SkipDownload) {
    Write-Host "Skipping download check." -ForegroundColor Yellow
    $downloadedNewVersion = $false
}
else {
    Write-Host "Checking for new version..." -ForegroundColor Cyan
    # Check for the latest version available online
    Write-Host "Checking for latest ScreenConnect version online..." -ForegroundColor Cyan
    $latestVersionInfo = Get-LatestScreenConnectVersion
}

if ('' -eq $latestVersionInfo -or $null -eq $latestVersionInfo) {
    Write-Host "Failed to retrieve (or skipped) latest version information. Will check for local installers." -ForegroundColor Yellow
}
else {
    Write-Host "Latest version information:" -ForegroundColor Cyan
    Write-Host "  Version: $($latestVersionInfo.Version)" -ForegroundColor White
    Write-Host "  File: $($latestVersionInfo.FileName)" -ForegroundColor White
    Write-Host "  Size: $($latestVersionInfo.Size)" -ForegroundColor White
    Write-Host "  Released: $($latestVersionInfo.ReleaseDate)" -ForegroundColor White
    Write-Host "  Platforms: $($latestVersionInfo.Platforms)" -ForegroundColor White
    Write-Host "  Download URL: $($latestVersionInfo.DownloadUrl)" -ForegroundColor White
    
    # Compare versions
    $comparisonResult = Compare-ScreenConnectVersions -CurrentVersion $svcVersion -LatestVersion $latestVersionInfo.Version
    
    if ($null -eq $comparisonResult) {
        Write-Host "Failed to compare versions. Will check for local installers." -ForegroundColor Yellow
    }
    elseif ($comparisonResult.IsUpdateAvailable) {
        Write-Host "`nAn update is available!" -ForegroundColor Yellow
        Write-Host "Current version: $($comparisonResult.CurrentVersion)" -ForegroundColor White
        Write-Host "Latest version: $($comparisonResult.LatestVersion)" -ForegroundColor White
        
        # Prompt user to download the update
        if ($Force -or (Read-Host -Prompt "Do you want to download the update? (Y/N)") -like "y*") {
            # Ensure download folder exists and is resolved
            $DownloadFolder = Resolve-Path -Path $Downloads -ErrorAction SilentlyContinue
            if (-not $DownloadFolder) {
                $DownloadFolder = (New-Item -Path $Downloads -ItemType Directory -Force).FullName
                Write-Host "Created download folder: $DownloadFolder" -ForegroundColor Cyan
            }
            
            # Download the file
            $downloadPath = Download-ScreenConnectFile -Url $latestVersionInfo.DownloadUrl -DestinationFolder $DownloadFolder -FileName $latestVersionInfo.FileName
            
            if ($null -ne $downloadPath -and (Test-Path -Path $downloadPath)) {
                Write-Host "`nDownload successful!" -ForegroundColor Green
                Write-Host "File saved to: $downloadPath" -ForegroundColor White
                $InstallerFile = $downloadPath
                $InstallerVersion = $latestVersionInfo.Version
                $installerVer = [version]$InstallerVersion
                $installedVer = [version]$svcVersion
                $downloadedNewVersion = $true
            }
            else {
                Write-Host "`nDownload failed. Will check for existing installers." -ForegroundColor Red
                $downloadedNewVersion = $false
            }
        }
        else {
            Write-Host "`nDownload canceled. Will check for existing installers." -ForegroundColor Yellow
            $downloadedNewVersion = $false
        }
    }
    elseif ($comparisonResult.AreVersionsEqual) {
        Write-Host "`nYou have the latest version installed." -ForegroundColor Green
        
        # Check if the installer for the current version exists in the Downloads folder
        $DownloadFolder = Resolve-Path -Path $Downloads -ErrorAction SilentlyContinue
        $existingInstaller = $null
        if ($DownloadFolder) {
            $existingInstaller = Get-ChildItem -Path "$DownloadFolder\ScreenConnect_$($latestVersionInfo.Version)*.msi" -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        
        if ($null -eq $existingInstaller) {
            # Installer doesn't exist, prompt to download it
            if ($Force -or (Read-Host -Prompt "Installer for current version not found. Download current installer anyway? (Y/N)") -like "y*") {
                # Ensure download folder exists and is resolved
                if (-not $DownloadFolder) {
                    $DownloadFolder = (New-Item -Path $Downloads -ItemType Directory -Force).FullName
                    Write-Host "Created download folder: $DownloadFolder" -ForegroundColor Cyan
                }
                
                # Download the file
                $downloadPath = Download-ScreenConnectFile -Url $latestVersionInfo.DownloadUrl -DestinationFolder $DownloadFolder -FileName $latestVersionInfo.FileName
                
                if ($null -ne $downloadPath -and (Test-Path -Path $downloadPath)) {
                    Write-Host "`nDownload successful!" -ForegroundColor Green
                    Write-Host "File saved to: $downloadPath" -ForegroundColor White
                    $InstallerFile = $downloadPath
                    $InstallerVersion = $latestVersionInfo.Version
                    $installerVer = [version]$InstallerVersion
                    $installedVer = [version]$svcVersion
                    $downloadedNewVersion = $true
                }
                else {
                    Write-Host "`nDownload failed. Will check for existing installers." -ForegroundColor Red
                    $downloadedNewVersion = $false
                }
            }
            else {
                Write-Host "`nDownload canceled. Will check for existing installers." -ForegroundColor Yellow
                $downloadedNewVersion = $false
            }
        }
        else {
            Write-Host "Installer for current version already exists: $($existingInstaller.FullName)" -ForegroundColor Cyan
            $InstallerFile = $existingInstaller.FullName
            $InstallerVersion = $latestVersionInfo.Version
            $installerVer = [version]$InstallerVersion
            $installedVer = [version]$svcVersion
            $downloadedNewVersion = $true
        }
    }
    elseif ($comparisonResult.IsCurrentVersionNewer) {
        Write-Host "`nYour current version is newer than the latest stable release." -ForegroundColor Magenta
        Write-Host "You might be using a pre-release or custom build. Will check for existing installers anyway." -ForegroundColor White
        $downloadedNewVersion = $false
    }
}

# If we didn't download a new version or if download failed, check for existing installers
if (-not (Get-Variable -Name downloadedNewVersion -ErrorAction SilentlyContinue) -or -not $downloadedNewVersion) {
    $InstallerPath = Resolve-Path -Path $Downloads -ErrorAction SilentlyContinue
    if (!$InstallerPath) {
        Write-Host "Unable to locate the $Downloads folder to look for installers. Quitting" -ForegroundColor Red
        exit 1
    }
    else {
        Write-Host "Looking for installers in the $InstallerPath folder..." -ForegroundColor Cyan
    }
    
    $installer = (Get-ChildItem "$InstallerPath\ScreenConnect_*.msi" -Attributes !Directory | Sort-Object -Descending -Property LastWriteTime | Select-Object -First 1)
    if (!($installer)) {
        Write-Host "Unable to locate a ScreenConnect MSI installer to install." -ForegroundColor Red
        exit 1
    }
    else {
        $InstallerFile = $installer.FullName
        Write-Host "Located the installer $InstallerFile, analyzing..." -ForegroundColor Cyan
    }

    $InstallerProps = $installer | Get-MsiProperty -PassThru | Select-Object Name, MSIProperties
    $InstallerVersion = $InstallerProps.productversion
    $InstallerName = $InstallerProps.productname

    if ($InstallerName -ne 'ScreenConnect') {
        Write-Host "The $InstallerFile file is not a ScreenConnect installer! Quitting." -ForegroundColor Red
        exit 1
    }
    else {
        Write-Host
    }

    Write-Host "Installer File:`t`t" $InstallerFile
    Write-Host "File write time:`t" $installer.LastWriteTime
    Write-Host "Installer Version:`t" $InstallerVersion -ForegroundColor Yellow
    Write-Host
    
    $installedVer = [version]$svcVersion
    $installerVer = [version]$InstallerVersion
}

# Compare the installer version to the installed version
if ($installerVer -eq $installedVer) {
    Write-Host "The installer version is the SAME as current install! Quitting." -ForegroundColor Red
    exit 1
}
elseif ($installerVer -lt $installedVer) {
    Write-Host "The installer is OLDER than current install! Quitting." -ForegroundColor Red
    exit 1
}
elseif ($installerVer -gt $installedVer) {
    Write-Host "The installer is newer than current install. Proceeding..." -ForegroundColor Green
}

If ($Force -or (Read-Host "Silently install this MSI file? (Yes/No)") -Like "y*") {
    $InstallerLogFile = [io.path]::ChangeExtension([io.path]::GetTempFileName(), ".log")
    Write-Host "InstallerLogFile:`t" $InstallerLogFile
    $Arguments = " /c msiexec /i `"$InstallerFile`" /qn /l*v `"$InstallerLogFile`""
    $Process = Start-Process -Wait cmd -ArgumentList $Arguments -PassThru
    if ($Process.ExitCode -ne 0) {
        Get-Content $InstallerLogFile -ErrorAction SilentlyContinue | Select-Object -Last 200
        Write-Host "Upgrade failed, please troubleshoot manually. Log file: $InstallerLogFile" -ForegroundColor Red
    }
    else {
        Write-Host "Uprade successfully completed. Please upgrade endpoint agents and adjust filter list version in application." -ForegroundColor Green
    }
}
else {
    Write-Host "Cancelled, no action taken." -ForegroundColor Red
}
