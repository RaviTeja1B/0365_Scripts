# Import the Active Directory module
Import-Module ActiveDirectory -Force

# Get all groups in AD
$groups = Get-ADGroup -Filter *

# Create an array to store the group and member data
$data = @()

# Iterate through each group and retrieve its members
foreach ($group in $groups) {
    $groupData = Get-ADGroupMember -Identity $group | Select-Object Name, SamAccountName, objectClass
    foreach ($member in $groupData) {
        $data += [PSCustomObject]@{
            'Group Name'      = $group.Name
            'Member Name'     = $member.Name
            'Member Account'  = $member.SamAccountName
            'Member Type'     = $member.objectClass
        }
    }
}

# Export the data to a CSV file
$csvPath = "C:\Users\ravi.morampudi\Downloads\AD_Groups_Output\groups_and_members.csv"
$data | Export-Csv -Path $csvPath -NoTypeInformation

Write-Host "Data exported to $csvPath"
