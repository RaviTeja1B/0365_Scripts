Install-Module SharePointPnPPowerShellOnline -Force

$siteurl = "https://epsoftwareinc.sharepoint.com/sites/IT-Audits"
$username = "audit.user@epsoftinc.com"
$password = "Pup90286"

$securePassword = ConvertTo-SecureString -String $password -AsPlainText -Force
$credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $securePassword

$credentialName = "$siteurl"
Add-PnPStoredCredential -Name $credentialName -Username $credentials.UserName -Password $credentials.Password

Connect-PnPOnline -Url $siteurl -WarningAction Ignore