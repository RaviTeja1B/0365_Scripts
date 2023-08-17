# Parameters
$folderName = "$(Get-Date -format MMM)$(Get-Date -format yyyy)"
$folderPath = "C:\O365-Audit\Results\$folderName"

# Check if folder exists, create it if it doesn't
if (!(Test-Path $folderPath)) {
    New-Item -ItemType Directory -Path $folderPath
}

$OutputCsv = "$folderPath\DL-Members$((Get-Date -format MMM-dd` hh-mm` tt).ToString()).csv"

# Retrieve stored credentials
$Credential = Get-StoredCredential -Target EPSoft

Try {
    # Connect to Exchange Online using the retrieved credentials
    Connect-ExchangeOnline -Credential $Credential -ShowBanner:$False

    # Get all Distribution Lists
    $Result = @()
    $DistributionGroups = Get-DistributionGroup -ResultSize Unlimited
    $GroupsCount = $DistributionGroups.Count
    $Counter = 1

    $DistributionGroups | ForEach-Object {
        $Group = $_
        Write-Progress -Activity "Processing Distribution List: $($Group.DisplayName)" -Status "$Counter out of $GroupsCount completed" -PercentComplete (($Counter / $GroupsCount) * 100)

        Get-DistributionGroupMember -Identity $Group.Name -ResultSize Unlimited | ForEach-Object {
            $Member = $_
            $Result += New-Object PSObject -property @{
                GroupName     = $Group.Name
                GroupEmail    = $Group.PrimarySmtpAddress
                Member        = $Member.Name
                EmailAddress  = $Member.PrimarySMTPAddress
                RecipientType = $Member.RecipientType
            }
        }

        $Counter++
    }

    # Export the result to CSV
    $Result | Export-CSV $OutputCsv -NoTypeInformation -Encoding UTF8
}
Catch {
    Write-Host -ForegroundColor Red "Error: $($_.Exception.Message)"
}
