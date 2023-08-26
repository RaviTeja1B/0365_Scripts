# Import the Active Directory module
Import-Module ActiveDirectory

# Get all groups in AD
$groups = Get-ADGroup -Filter *

# Iterate through each group and retrieve its members
foreach ($group in $groups) {
    $groupData = Get-ADGroupMember -Identity $group | Select-Object Name, SamAccountName, objectClass

    # Create an array to store the group and member data
    $data = @()
    
    foreach ($member in $groupData) {
        $data += [PSCustomObject]@{
            'Group Name'      = $group.Name
            'Member Name'     = $member.Name
            'Member Account'  = $member.SamAccountName
            'Member Type'     = $member.objectClass
        }
    }

    # Export the data to a CSV file
    $csvPath = "C:\Users\ravi.morampudi\Downloads\AD_Group_Users\$($group.Name)_users.csv"
    $data | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "Data exported for $($group.Name) to $csvPath"
}
