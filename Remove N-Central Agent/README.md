# [Remove-N-Central-Agent.ps1](./Remove-N-Central-Agent.ps1)

## SYNOPSIS
Remove N-able N-Central Windows Agent and related services, both via uninstallation of the agent and related apps (not the Probe) and also via manual cleanup (optional with the `-Clean` parameter) if the uninstall fails.

## DESCRIPTION
Uninstalls and optionally cleans/removes the N-able N-Central Windows Agent and related services and folders. It will attempt to run uninstallers for all known agent entries from Add/Remove Programs.

If the `-Clean` parameter is specified, it will also attempt (after uninstallation) to clean up agent remnants in addition to attempting uninstallation--this includes stopping and disabling related services, removing both the Uninstall and regular application registry entries and disabling the services, attempting to delete the services, and then removing the installation folders and the related folders in `C:\ProgramData`.

The script _should NOT_ remove N-able Cove backup installations or files/services.

The script _should NOT_ uninstall or remove the Windows Probe service or agent, if it exists.

The script _does NOT_ remove the registry setting showing the N-central device ID that would be used if the agent were ever reinstalled to map to the same device in the future. This ID should be a harmless artifact in most cases especially if the agent is no longer running on the system, but it could be updated to remove it if desired.

The script may leave some top-level folders under Program Files, that are empty, or may leave some subfolders that are in use and cannot be removed due to permissions, but it will attempt to remove some of these after a reboot if possible, and any remaining remnants afterwards should not allow the agent to run, even if they aren't completely cleaned up.

The script _WILL delete_ agent logs and history in ProgramData subfolders.

The script _WILL attempt to remove_ the Take Control (BeAnywhere) services and folders if they exist and -Clean is specified, but does not attempt any separate uninstallation.

The script should be run with admin privileges, ideally as SYSTEM, and will quit if it is not.

Paths and services to clean up are hardcoded into the script under **CONFIG AND SETUP**, and will use the correct system drive for the system but the rest of the paths are hardcoded. The installation folders that any existing services refer to will be added to the cleanup list, if they are different and exist during the run (if `-Clean` is run as part of the initial pass).

While it can be run manually, it is recommended that the script be run via a different RMM tool, and supports but does not require NinjaRMM Script Variables with the parameter names (as checkboxes) for configuration.

## KNOWN ISSUES
*Known issues with version 0.0.1:* The services, while they are deleted, are not always fully removed from the system when cleaned if the uninstallations fail. This is a bug that has not been diagnosed/fixed yet, but the services should still be left in the Stopped and Disabled state.

## PARAMETER Clean
Clean up agent remnants in addition to attempting uninstallation.

## PARAMETER TestOnly
Test removal of agent and services without actually removing them--will output test info to console instead of making changes. Kind of like a custom -WhatIf dry run without being official.

## EXAMPLE
```
Remove-N-Central-Agent.ps1
```

## EXAMPLE
```
Remove-N-Central-Agent.ps1 -Clean
```

## NOTES
**Version 0.0.1** - 2024-03-25 by David Szpunar - Initial release
