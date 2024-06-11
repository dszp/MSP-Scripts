#If existing install, require to stop the current service from running before being able to install an update

# Get-Service -Name "DefensX Agent" -ErrorAction SilentlyContinue | Stop-Service -Force -ErrorAction SilentlyContinue

$Arguments = @"
/i "$InstallerFile" /quiet /norestart /l*v "$InstallerLogFile" KEY=$LicenseValue BRAVE_DISABLE_PRIVATE_WINDOW=$BRAVE_DISABLE_PRIVATE_WINDOW CHROME_DISABLE_PRIVATE_WINDOW=$CHROME_DISABLE_PRIVATE_WINDOW CHROMIUM_DISABLE_PRIVATE_WINDOW=$CHROMIUM_DISABLE_PRIVATE_WINDOW EDGE_DISABLE_PRIVATE_WINDOW=$EDGE_DISABLE_PRIVATE_WINDOW FIREFOX_DISABLE_PRIVATE_WINDOW=$FIREFOX_DISABLE_PRIVATE_WINDOW VIVALDI_DISABLE_PRIVATE_WINDOW=$VIVALDI_DISABLE_PRIVATE_WINDOW DISABLE_UNINSTALL=$DISABLE_UNINSTALL ENABLE_BYPASS_MODE=$ENABLE_BYPASS_MODE ENABLE_DRIVER_MODE=$ENABLE_DRIVER_MODE ENABLE_IAM_USER=$ENABLE_IAM_USER ENABLE_LOGON_USER=$ENABLE_LOGON_USER SYSTEM_COMPONENT=$SYSTEM_COMPONENT
"@

Write-Host "Arguments: $Arguments"
Write-Host "InstallerLogFile: $InstallerLogFile"        
$Process = Start-ProcessWithLogTail msiexec -ArgumentList $Arguments -LogFilePath $InstallerLogFile
Write-Host "Exit Code: $($Process.ExitCode)";