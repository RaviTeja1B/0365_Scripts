# Import the Active Directory module
Import-Module ActiveDirectory

# Define the output file paths for enabled and disabled users
$enabledUsersFilePath = "C:\Users\ravi.morampudi\Downloads\AD_Groups_Users_Enable_Disable\EnabledUsers.csv"
$disabledUsersFilePath = "C:\Users\ravi.morampudi\Downloads\AD_Groups_Users_Enable_Disable\DisabledUsers.csv"

# Get all user accounts in the AD domain
$users = Get-ADUser -Filter *

# Separate enabled and disabled users
$enabledUsers = $users | Where-Object { $_.Enabled -eq $true }
$disabledUsers = $users | Where-Object { $_.Enabled -eq $false }

# Export enabled users to CSV
$enabledUsers | Select-Object SamAccountName, Enabled | Export-Csv -Path $enabledUsersFilePath -NoTypeInformation

# Export disabled users to CSV
$disabledUsers | Select-Object SamAccountName, Enabled | Export-Csv -Path $disabledUsersFilePath -NoTypeInformation
