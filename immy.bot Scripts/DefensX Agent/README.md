# Deploying DefensX Agent with immy.bot

## Setting up immy.bot Software Deployment

Create a new Software Library entry in immy.bot named DefensX agent. I use the following Notes field:

> This is the Windows agent for the DefensX.com web filtering and browser control/management software. The online console to create and manage tenants and get deployment keys is https://cloud.defensx.com

The Software Info field looks like this for me:  
![immy.bot Software Info example screenshot.](<assets/screenshot_Deploying Defen_Image-4.png>)

The Icon logo is in this repository as [DefensXShield.png](DefensXShield.png). (Note that this is for internal use only and is trademarked by DefensX, and I will remove it should DefensX request removal.)

For licensing, change to Required of type Key. You'll create a License for each tenant with the short key from DefensX (the long key should work) to select for each deployment:  
![immy.bot Licensing Configuration screenshot for DefensX.](<assets/screenshot_Deploying Defen_Image-1.png>)

The description I used for the license key field is:

> Short activation key from cloud.defensx.com in each customer tenant's Policies->Policy Groups->Deployments->RMM button->Enable "Use Short Deployment Key"->Copy value after KEY= in URL and use as license key value for tenant. Default long Deployment Key from the same area may work as well.

But use whatever you want for these fields, they are for reference only.

Choose Upgrade Code as the Version Detection method. You won't see the existing code on a fresh deployment, so ignore the rest beyond choosing "Upgrade Code" as the Detection Method:  
![Version Detection portion of immy.bot deployment.](<assets/screenshot_Deploying Defen_Image-2.png>)

Set up the scripts as shown below, create New scripts for the two that are included in this repository and select them for the relevant areas:  
![DefensX delpoyment scripts selection in immy.bot.](<assets/screenshot_Deploying Defen_Image-3.png>)

The [DefensXSilentInstall.ps1](DefensXSilentInstall.ps1) and [DefensXDynamicVersions.ps1](DefensXDynamicVersions.ps1) scripts are included in this repository and are the only ones needed that are not already included with immy.bot as Global scripts that don't require modification.

## Configuration Task
The Configuration Task needs to be created to have all of the flags that can be passed to the installer to configure settings in the agent at deployment. The flags are listed below, you should create a New Configuration Task with the following parameter names, all of type "Number", with no Requires User Input options and with a default value of `0` or `1` as desired. The description field is optional.

| Parameter Name | Default Value | Description |
| -------------- | ------------- | ----------- |
| CHROME_DISABLE_PRIVATE_WINDOW | 0 | When set to 1, private browser functionality will NOT be disabled for Chrome. It will be disabled with value of 0 (default). |
| EDGE_DISABLE_PRIVATE_WINDOW | 0 | When set to 1, private browser functionality will NOT be disabled for Edge. It will be disabled with value of 0 (default). |
| FIREFOX_DISABLE_PRIVATE_WINDOW | 0 | When set to 1, private browser functionality will NOT be disabled for Firefox. It will be disabled with value of 0 (default). |
| CHROMIUM_DISABLE_PRIVATE_WINDOW | 0 | When set to 1, private browser functionality will NOT be disabled for Chromium. It will be disabled with value of 0 (default). |
| BRAVE_DISABLE_PRIVATE_WINDOW | 0 | When set to 1, private browser functionality will NOT be disabled for Brave. It will be disabled with value of 0 (default). |
| VIVALDI_DISABLE_PRIVATE_WINDOW | 0 | When set to 1, private browser functionality will NOT be disabled for Vivaldi. It will be disabled with value of 0 (default). |
| DISABLE_UNINSTALL | 0 | If set to 1, users cannot uninstall the agent from Add/Remove Programs, but it can still be seen in the list. |
| SYSTEM_COMPONENT | 0 | If set to 1, users cannot see the agent in Add/Remove Programs. |
| ENABLE_IAM_USER | 0 | When enabled (set to 1), DefensX Agent allows users to sign in using configured methods (Google, Octa, Local Users with or without MFA) within the DefensX interface. Please note that this method requires users to sign in interactively. |
| ENABLE_LOGON_USER | 1 | When enabled with value 1 (or not set), DefensX Agent automatically creates new users on the DefensX backend for logged-in Windows users based on their Windows Logon usernames. Set to 0 to disable. |
| ENABLE_DRIVER_MODE | 1 | When kernel driver is enabled (default, set to 1), DefensX Agent will try to load kernel drivers to redirect DNS requests itself without changing system wide DNS settings. |
| ENABLE_BYPASS_MODE | 0 | When Bypass mode is set to Enabled (value of 1), DefensX Agent will allow users to temporarily remove protection. It can be useful in some Captive-Portal authentication and troubleshooting. |

Save/update the Software and create a deployment using the new software library item. The Configuration Task area should provide a customizable list of parameters for you to edit if desired.

Although the default here is 0 for the options to disable private browser windows, we generally deploy DefensX with the `CHROME_DISABLE_PRIVATE_WINDOW`, `EDGE_DISABLE_PRIVATE_WINDOW`, and `FIREFOX_DISABLE_PRIVATE_WINDOW` flags set to 1 in order to allow Private Window to bypass filtering for Chrome, Edge, and Firefox, allowing for troubleshooting and testing, although this is a less-secure overall configuration so consider that in your deployment plans.

## Parameter Block for Reference
The parameter block that immy.bot defines for these parameters is equal to the following, but I've added the parameters directly in the UI so this is not currently being used and is added for reference only:

```powershell
param(
[Parameter(Position=0,Mandatory=$False,HelpMessage=@'
When set to 1, private browser functionality will NOT be disabled for Chrome. It will be disabled with value of 0 (default).
'@)]
[Int32]$CHROME_DISABLE_PRIVATE_WINDOW=0,
[Parameter(Position=1,Mandatory=$False,HelpMessage=@'
When set to 1, private browser functionality will NOT be disabled for Edge. It will be disabled with value of 0 (default).
'@)]
[Int32]$EDGE_DISABLE_PRIVATE_WINDOW=0,
[Parameter(Position=2,Mandatory=$False,HelpMessage=@'
When set to 1, private browser functionality will NOT be disabled for Firefox. It will be disabled with value of 0 (default).
'@)]
[Int32]$FIREFOX_DISABLE_PRIVATE_WINDOW=0,
[Parameter(Position=3,Mandatory=$False,HelpMessage=@'
When set to 1, private browser functionality will NOT be disabled for Chromium. It will be disabled with value of 0 (default).
'@)]
[Int32]$CHROMIUM_DISABLE_PRIVATE_WINDOW=0,
[Parameter(Position=4,Mandatory=$False,HelpMessage=@'
When set to 1, private browser functionality will NOT be disabled for Brave. It will be disabled with value of 0 (default).
'@)]
[Int32]$BRAVE_DISABLE_PRIVATE_WINDOW=0,
[Parameter(Position=5,Mandatory=$False,HelpMessage=@'
When set to 1, private browser functionality will NOT be disabled for Vivaldi. It will be disabled with value of 0 (default).
'@)]
[Int32]$VIVALDI_DISABLE_PRIVATE_WINDOW=0,
[Parameter(Position=6,Mandatory=$False,HelpMessage=@'
If set to 1, users cannot uninstall the agent from Add/Remove Programs, but it can still be seen in the list.
'@)]
[Int32]$DISABLE_UNINSTALL=0,
[Parameter(Position=7,Mandatory=$False,HelpMessage=@'
If set to 1, users cannot see the agent in Add/Remove Programs.
'@)]
[Int32]$SYSTEM_COMPONENT=0,
[Parameter(Position=8,Mandatory=$False,HelpMessage=@'
When enabled (set to 1), DefensX Agent allows users to sign in using configured methods (Google, Octa, Local Users with or without MFA) within the DefensX interface. Please note that this method requires users to sign in interactively.
'@)]
[Int32]$ENABLE_IAM_USER=0,
[Parameter(Position=9,Mandatory=$False,HelpMessage=@'
When enabled with value 1 (or not set), DefensX Agent automatically creates new users on the DefensX backend for logged-in Windows users based on their Windows Logon usernames. Set to 0 to disable.

This feature simplifies deployments in both device-joined Azure AD and Active Directory environments by eliminating the need for user interaction.
'@)]
[Int32]$ENABLE_LOGON_USER=1,
[Parameter(Position=10,Mandatory=$False,HelpMessage=@'
When kernel driver is enabled (default, set to 1), DefensX Agent will try to load kernel drivers to redirect DNS requests itself without changing system wide DNS settings.

It is safe to use kernel driver, if the driver not available for the installed platform, agent automatically fallback to standard mode.
'@)]
[Int32]$ENABLE_DRIVER_MODE=1,
[Parameter(Position=11,Mandatory=$False,HelpMessage=@'
When Bypass mode is set to Enabled (value of 1), DefensX Agent will allow users to temporarily remove protection. It can be useful in some Captive-Portal authentication and troubleshooting.
'@)]
[Int32]$ENABLE_BYPASS_MODE=0
)
```
