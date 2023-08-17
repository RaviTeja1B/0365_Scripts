
$Credential = Get-StoredCredential -Target EPSoft

 

$ReportPath = 'C:\O365-Audit\Results\'

$ZipPath = 'C:\O365-Audit\Results\AuditReports.zip'

 

# Get the latest file in the directory based on creation date

$LatestFile = Get-ChildItem -Path $ReportPath | Sort-Object -Property CreationTime -Descending | Select-Object -First 1

 

# If you want to use the latest file based on modification date instead, replace "CreationTime" with "LastWriteTime" in the above line.

 

# Delete the previous zip file, if it exists

if (Test-Path $ZipPath) {

    Remove-Item $ZipPath

}

 

# Create a new zip file containing only the latest file

Compress-Archive -Path $LatestFile.FullName -DestinationPath $ZipPath

 

$date = Get-Date -UFormat %B-%Y

$Subject = "O365 Audit Reports - $date"

$Body = "PFA O365 Audit Reports for the month of $date"
$Recipient = "venkat.koneru@epsoftinc.com", "srikanth.varry@epsoftinc.com", "ravi.morampudi@epsoftinc.com".
 

Send-MailMessage -SmtpServer smtp.office365.com -Port 587 -UseSsl -From itadmin@epsoftinc.com -To $Recipient -Subject $Subject -Body $Body -Credential $Credential -Attachments $ZipPath
