Param
(
    [Parameter(Mandatory = $false)]
    [switch]$MFA,
    [switch]$Default,
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$UserID,
    [string]$AdminName,
    [string]$Password
)

#Check for EXO v2 module inatallation
$Module = Get-Module ExchangeOnlineManagement -ListAvailable
if($Module.count -eq 0) 
{ 
 Write-Host Exchange Online PowerShell V2 module is not available  -ForegroundColor yellow  
 $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
 if($Confirm -match "[yY]") 
 { 
  Write-host "Installing Exchange Online PowerShell module"
  Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
 } 
 else 
 { 
  Write-Host EXO V2 module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
  Exit
 }
} 

#Check for MSOnline module 
$Module=Get-Module -Name MSOnline -ListAvailable  
if($Module.count -eq 0) 
{ 
 Write-Host MSOnline module is not available  -ForegroundColor yellow  
 $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
 if($Confirm -match "[yY]") 
 { 
  Install-Module MSOnline 
  Import-Module MSOnline
 } 
 else 
 { 
  Write-Host MSOnline module is required to connect AzureAD.Please install module using Install-Module MSOnline cmdlet. 
  Exit
 }
} 

#Connect Exchange Online with MFA
 if($MFA.IsPresent)
 {
  Write-Host Connecting to Exchange Online...
  Connect-ExchangeOnline
  Write-Host Connecting to MSOnline module...
  Connect-MsolService
 }

#Storing credential in script for scheduling purpose/ Passing credential as parameter - Authentication using non-MFA account
Connecting to Exchange Online...
    Write-Host "Connecting to ExchangeOnline..." -ForegroundColor Cyan

 $Credential  = Get-StoredCredential -Target EPSoft
 Connect-ExchangeOnline -Credential $Credential


#Getting user activity for past 90 days
if(($StartDate -eq $null) -and ($EndDate -eq $null))
{
 $EndDate=(Get-Date).Date
 $StartDate=$MaxStartDate
}
 
#Getting start date for Audit log  
While($true)
{
 if ($StartDate -eq $null)
 {
  $StartDate=Read-Host Enter start time for audit collection '(Eg:11/20/2019)'
 }
 Try
 {
  $Date=[DateTime]$StartDate
  if($Date -ge $MaxStartDate)
  { 
   break
  }
  else
  {
   Write-Host `nAudit log can be retrieved only for past 90 days. Please select a date after (Get-Date).AddDays(-90) -ForegroundColor Red
   return
  }
 }
 Catch
 {
  Write-Host `nNot a valid date -ForegroundColor Red
 }
}


#Getting end date for Audit log
While($true)
{
 if ($EndDate -eq $null)
 {
  $EndDate=Read-Host Enter End time for audit collecton '(Eg: 11/20/2019)'
 }
 Try
 {
  $Date=[DateTime]$EndDate
  if($EndDate -lt ($StartDate))
  {
   Write-Host End time should be later than start time -ForegroundColor Red
   return
  }
  break
 }
 Catch
 {
  Write-Host `nNot a valid date -ForegroundColor Red
 }
}

$folderName = "$(Get-Date -format MMM)$(Get-Date -format yyyy)"
$folderPath = "C:\O365-Audit\Results\$folderName"
# Check if folder exists, create it if it doesn't
if (!(Test-Path $folderPath)) {
    New-Item -ItemType Directory -Path $folderPath
}

$OutputCSV="C:\O365-Audit\Results\$folderName\$UserActivityReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$IntervalTimeInMinutes=1440    #$IntervalTimeInMinutes=Read-Host Enter interval time period '(in minutes)'
$CurrentStart=$StartDate
$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)

#Check whether CurrentEnd exceeds EndDate
if($CurrentEnd -gt $EndDate)
{
 $CurrentEnd=$EndDate
}

if($CurrentStart -eq $CurrentEnd)
{
 Write-Host Start and end time are same.Please enter different time range -ForegroundColor Red
 Exit
}

Connect_EXO
$AggregateResults = @()
$CurrentResult= @()
$CurrentResultCount=0
$AggregateResultCount=0
Write-Host `nRetrieving user activity log from $StartDate to $EndDate... -ForegroundColor Yellow
$i=0
$ExportResult=""   
$ExportResults=@()  
while($true)
{ 
 #Write-Host Retrieving user activity log between StartDate $CurrentStart to EndDate $CurrentEnd ******* IntervalTime $IntervalTimeInMinutes minutes
 if($CurrentStart -eq $CurrentEnd)
 {
  Write-Host Start and end time are same.Please enter different time range -ForegroundColor Red
  Exit
 }
 #Write-Host !!!!!!!!!!!!!!
 #Getting audit log for given time range
 $Results=Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -UserIds $UserID -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000
 $ResultCount=($Results | Measure-Object).count
 $AllAuditData=@()
 foreach($Result in $Results)
 {
  $i++
  $MoreInfo=$Result.auditdata
  $AuditData=$Result.auditdata | ConvertFrom-Json
  $ActivityTime=Get-Date($AuditData.CreationTime) -format g
  $UserID=$AuditData.userId
  $Operation=$AuditData.Operation
  $ResultStatus=$AuditData.ResultStatus
  $Workload=$AuditData.Workload

  #Export result to csv
  $ExportResult=@{'Activity Time'=$ActivityTime;'User Name'=$UserID;'Operation'=$Operation;'Result'=$ResultStatus;'Workload'=$Workload;'More Info'=$MoreInfo}
  $ExportResults= New-Object PSObject -Property $ExportResult  
  $ExportResults | Select-Object 'Activity Time','User Name','Operation','Result','Workload','More Info' | Export-Csv -Path $OutputCSV -Notype -Append 
 }
 Write-Progress -Activity "`n     Retrieving audit log from $StartDate to $EndDate.."`n" Processed audit record count: $i"
 $currentResultCount=$CurrentResultCount+$ResultCount
 if($CurrentResultCount -eq 50000)
 {
  Write-Host Retrieved max record for current range.Proceeding further may cause data loss or rerun the script with reduced time interval. -ForegroundColor Red
  $Confirm=Read-Host `nAre you sure you want to continue? [Y] Yes [N] No
  if($Confirm -match "[Y]")
  {
   Write-Host Agg $AggregateResultCount CurrentResu $CurrentResultCount
   $AggregateResultCount +=$CurrentResultCount
   Write-Host Proceeding audit log collection with data loss
   [DateTime]$CurrentStart=$CurrentEnd
   [DateTime]$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)
   $CurrentResultCount=0
   $CurrentResult = @()
   if($CurrentEnd -gt $EndDate)
   {
    $CurrentEnd=$EndDate
   }
  }
  else
  {
   Write-Host Please rerun the script with reduced time interval -ForegroundColor Red
   Exit
  }
 }

 
 if($Results.count -lt 5000)
 {
  #$AggregateResults +=$CurrentResult
  $AggregateResultCount +=$CurrentResultCount
  if($CurrentEnd -eq $EndDate)
  {
   break
  }
  $CurrentStart=$CurrentEnd 
  if($CurrentStart -gt (Get-Date))
  {
   break
  }
  $CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)
  $CurrentResultCount=0
  $CurrentResult = @()
  if($CurrentEnd -gt $EndDate)
  {
   $CurrentEnd=$EndDate
  }
 }
}

If($AggregateResultCount -eq 0)
{
 Write-Host No records found
}
else
{
 Write-Host `nThe output file contains $AggregateResultCount audit records
 if((Test-Path -Path $OutputCSV) -eq "True") 
 {
  Write-Host `nThe Output file availble in $OutputCSV -ForegroundColor Green
  #$Prompt = New-Object -ComObject wscript.shell   
  #$UserInput = $Prompt.popup("Do you want to open output file?",`   
 #0,"Open Output File",4)   
 # If ($UserInput -eq 6)   
 # {   
  # Invoke-Item "$OutputCSV"   
  #} 
 }
}

#Disconnect Exchange Online session
 Disconnect-ExchangeOnline -Confirm:$false | Out-Null
