# Windows Security Hardening
These scripts help to harden various Windows security settings.

## [Harden-Security-Windows-Registry.ps1](Harden-Security-Windows-Registry.ps1)
The [Harden-Security-Windows-Registry.ps1](Harden-Security-Windows-Registry.ps1 "Harden-Security-Windows-Registry.ps1") script configures several common registry settings that lock down Windows client (and sometimes server) security.

Usually this means disabling old protocols or requiring SMB signing, etc., and on modern networks most of these already should not be used, but also in theory it could break things, so only run this on systems where you know it won't have a negative effect, or are prepared to diagnose issues/test things after adjusting.

There is currently no automatic undo for these settings, but given that these are all basic registry settings, usually deleting the key or changing the value will reverse the change. Many times, rebooting is required before a change will take effect. These settings may be overridden by local or group policy settings, or Intune, if those are configured.

Each setting in the script is accompanied by a comment with one or more links to the reference source showing where the solution was obtained.

This script could cause important network functions to stop working for networks that rely on one or more of the disabled protocols, even if they are legacy protocols that generally should no longer be used if possible. Don’t roll out without testing first in each environment.

## [Disable-PowerShell-V2.ps1](Disable-PowerShell-V2.ps1)
The [Disable-PowerShell-V2.ps1](Disable-PowerShell-V2.ps1) script checks to see if PowerShell v2 is enabled and disables it if it is.

Run the script to report and disable PowerShell v2. No change will be made if it's already disabled.

Leaving PowerShell V2 enabled on a system that supports it could allow for a security bypass due to being able to use it to skip some logging and auditing, hence [the recommendation to disable at StigViewer.com](https://www.stigviewer.com/stig/windows_10/2017-04-28/finding/V-70637).

In order to provide an undo option, pass the argument `-Enable` to instead enable PowerShell v2, if disabled.

Returns 0 if the requested action was successful or no action was taken, or 1 if a requested change fails to succeed after testing.

Will report if a reboot is required to complete the change (it usually isn't).

Checks Windows Optional Feature List to ensure PowerShell V2 is a possible feature first, and quits if not (for servers or OS versions where it's not supported at all), and succeeds but reports the finding. (Newer versions of Windows Server, including 2022, do not include PowerShell 2 at all, so the feature doesn’t exist to be disabled.)