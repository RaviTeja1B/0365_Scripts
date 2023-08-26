# Get all groups in AD
$groups = Get-ADGroup -Filter *

# Iterate through each group and retrieve its members
foreach ($group in $groups) {
    Write-Host "Group: $($group.Name)"
    Write-Host "Members:"
    Get-ADGroupMember -Identity $group | Select-Object Name, SamAccountName, objectClass
    Write-Host "------------------------"
}
