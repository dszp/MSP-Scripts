# # --- Dynamic Versions Script ---
# # The script MUST return an array of applicable version objects.
# # Each object is required to provide a "Version" and "Url" property.


# # use the New-DynamicVersion function to build out valid versions and assign them to $Response.Versions

# # you must explicitly return the data in the following structure
# $Response = New-Object PSObject -Property @{
#     Versions = @()
# }

# return $Response;
Get-DynamicVersionFromInstallerURL "https://cloud.defensx.com/defensx-installer/latest.msi"