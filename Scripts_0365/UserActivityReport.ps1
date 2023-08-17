param (
    [Parameter(Mandatory = $false)]
    [switch]$MFA,
    [switch]$Default,
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$UserID,
    [string]$AdminName,
    [string]$Password
)

# Check for EXO v2 module installation
$Module = Get-Module ExchangeOnlineManagement -ListAvailable
if ($Module.Count -eq 0) {
    Write-Host "Exchange Online PowerShell V2 module is not available" -ForegroundColor Yellow
    $Confirm = Read-Host "Are you sure you want to install the module? [Y] Yes [N] No"
    if ($Confirm -match "[yY]") {
        Write-Host "Installing Exchange Online PowerShell module"
        Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
    }
    else {
        Write-Host "EXO V2 module is required to connect to Exchange Online. Please install the module using Install-Module ExchangeOnlineManagement cmdlet."
        Exit
    }
}

# Check for MSOnline module
$Module = Get-Module -Name MSOnline -ListAvailable
if ($Module.Count -eq 0) {
    Write-Host "MSOnline module is not available" -ForegroundColor Yellow
    $Confirm = Read-Host "Are you sure you want to install the module? [Y] Yes [N] No"
    if ($Confirm -match "[yY]") {
        Install-Module MSOnline
        Import-Module MSOnline
    }
    else {
        Write-Host "MSOnline module is required to connect to Azure AD. Please install the module using Install-Module MSOnline cmdlet."
        Exit
    }
}

# Connect to Exchange Online
if ($MFA.IsPresent) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline
    Write-Host "Connecting to MSOnline module..."
    Connect-MsolService
}
else {
    Write-Host "Connecting to Exchange Online..."
    $Credential = Get-StoredCredential -Target EPSoft
    Connect-ExchangeOnline -Credential $Credential
}

# Calculate the StartDate and EndDate if not provided
if (-not $StartDate -and -not $EndDate) {
    $EndDate = Get-Date
    $StartDate = $EndDate.AddMonths(-1)
}

# Format the folder path and create it if it doesn't exist
$folderName = "$(Get-Date -format MMM)$(Get-Date -format yyyy)"
$folderPath = "C:\O365-Audit\Results\$folderName"
if (!(Test-Path $folderPath)) {
    New-Item -ItemType Directory -Path $folderPath | Out-Null
}

# Generate the output CSV file path
$OutputCSV = Join-Path $folderPath "UserActivityReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"

# Get user activity for the specified date range
Write-Host "Retrieving user activity log from $StartDate to $EndDate..."
$Results = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000

# Export the results to CSV
$Results | Export-Csv -Path $OutputCSV -NoTypeInformation

# Display the output file path
Write-Host "`nThe output file is available at: $OutputCSV" -ForegroundColor Green

# Disconnect Exchange
