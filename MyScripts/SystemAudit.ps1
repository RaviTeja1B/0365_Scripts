$computer = hostname
$exportLocation = 'C:\Users\Public\Documents\AuditData.csv'
$siteurl="https://epsoftwareinc.sharepoint.com/sites/IT-Audits"

$Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$Employee_Name = ( ( Get-WMIObject -class Win32_ComputerSystem | Select-Object -ExpandProperty username ) -split '\\' )[1]
$str1="$Employee_Name"
$str2="@epsoftinc.com"  
$Employee_Mail= $str1 + $str2
$Domain = (gwmi win32_ComputerSystem).Domain
$Admin_Access = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
$OS = Get-WmiObject Win32_OperatingSystem -Computername $computer
$Last_Boot = $OS.ConvertToDateTime($OS.LastBootUpTime)
$Disk_Space = Get-WmiObject win32_volume -computername $computer -Filter 'drivetype = 3' |
Select-Object PScomputerName, driveletter, label, @{LABEL='GBfreespace';EXPRESSION={'{0:N2}' -f($_.freespace/1GB)} } |
Where-Object { $_.driveletter -match 'C:' }
$Hardware = Get-WmiObject Win32_computerSystem -Computername $computer
$Total_Memory = [math]::round($Hardware.TotalPhysicalMemory/1024/1024/1024, 2)
#$Networks = (Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $computer | Where-Object {$_.IPEnabled -EQ $true}).Macaddress | select-object -first 1}
$MACAddress  = (Get-WmiObject Win32_NetworkAdapterConfiguration | where {$_.ipenabled -EQ $true}).Macaddress | select-object -first 1
$Cpu = Get-WmiObject Win32_Processor  -computername $computer
$User_List = (get-wmiobject -Class Win32_UserAccount -filter "localaccount=true" | where {$_.disabled -eq $False} | select object -ExpandProperty Name) -join " , "
$ProgramsInstalled = (Get-Package | Where-Object {$_.ProviderName -in @('Programs','msi','Chocolatey')} | Select-Object -ExpandProperty Name) -join [environment]::NewLine
$str1= (Get-WmiObject Win32_computerSystem -Computername $computer).PrimaryOwnerName
$str2= (Get-WmiObject Win32_computerSystem -Computername $computer).Model
$Model= $str1 + $str2
$Bios = Get-WmiObject win32_bios -Computername $computer
$systemBios = $Bios.serialnumber
$Sysbuild = Get-WmiObject Win32_WmiSetting -Computername $computer
$OS = Get-WmiObject Win32_OperatingSystem -Computername $computer
$Admin_Passwd_Set = ((net localgroup administrators) |where {$_ -AND $_ -notmatch "command completed successfully"} | select -skip 4)  -join ','
Get-ItemProperty 'HKLM:\SOFTWARE\WOW6432Node\Sophos\Health\Status' -Name service.* | Select * -Exclude PS* | out-string | ForEach-Object { $_.Trim() } | Out-File -FilePath 'C:\Users\Public\Documents\sophos.txt'
$Sophos1=((Get-Content 'C:\Users\Public\Documents\sophos.txt' ) -Replace '0', 'Running' -Replace '1', 'Stopped') | where { $_.Contains("Stopped") }
$Sophos2=((Get-Content 'C:\Users\Public\Documents\sophos.txt' ) -Replace '0', 'Running' -Replace '1', 'Stopped') | where { $_.Contains("Running") }
$Sophos=$Sophos1 + $Sophos2 -join [environment]::NewLine
       
$OutputObj  = New-Object -Type PSObject
$OutputObj | Add-Member -MemberType NoteProperty -Name Audit_Date -Value $Date
$OutputObj | Add-Member -MemberType NoteProperty -Name Employee_Name -Value $Employee_Name
$OutputObj | Add-Member -MemberType NoteProperty -Name Computer_Name -Value $computer.ToUpper()
$OutputObj | Add-Member -MemberType NoteProperty -Name Employee_Mail -Value $Employee_Mail
$OutputObj | Add-Member -MemberType NoteProperty -Name Domain -Value $Domain
$OutputObj | Add-Member -MemberType NoteProperty -Name Admin_Access -Value $Admin_Access
$OutputObj | Add-Member -MemberType NoteProperty -Name Last_ReBoot -Value $Last_Boot
$OutputObj | Add-Member -MemberType NoteProperty -Name C:_FreeSpace_GB -Value $Disk_Space.GBfreespace
$OutputObj | Add-Member -MemberType NoteProperty -Name Total_Memory_GB -Value $Total_Memory
$OutputObj | Add-Member -MemberType NoteProperty -Name MAC_Address -Value $MACAddress
$OutputObj | Add-Member -MemberType NoteProperty -Name Processor_Type -Value $Cpu.Name
$OutputObj | Add-Member -MemberType NoteProperty -Name User_List -Value $User_list
$OutputObj | Add-Member -MemberType NoteProperty -Name Installed_Apps -Value $ProgramsInstalled
$OutputObj | Add-Member -MemberType NoteProperty -Name Model -Value $Model
$OutputObj | Add-Member -MemberType NoteProperty -Name System_Type -Value $Hardware.SystemType
$OutputObj | Add-Member -MemberType NoteProperty -Name Operating_System -Value $OS.Caption
$OutputObj | Add-Member -MemberType NoteProperty -Name Operating_System_Version -Value $OS.version
$OutputObj | Add-Member -MemberType NoteProperty -Name Operating_System_BuildVersion -Value $SysBuild.BuildVersion
$OutputObj | Add-Member -MemberType NoteProperty -Name Serial_Number -Value $systemBios
$OutputObj | Add-Member -MemberType NoteProperty -Name Local_Admin_Accounts -Value $Admin_Passwd_Set
$OutputObj | Add-Member -MemberType NoteProperty -Name Sophos -Value $Sophos
$OutputObj | Export-Csv -Path $exportLocation -NoTypeInformation -Force

$Services=Get-Service -DisplayName *sophos* | select -expand Name
    if ($Services -eq $null)
	{
          $OutputObj | Add-Member -MemberType NoteProperty -Name Sophos_Service 'Sophos Endpoint is not installed'
        }else
	{
         $OutputObj | Add-Member -MemberType NoteProperty -Name Sophos_Service 'Sophos Endpoint is installed'
} 
$OutputObj | Export-Csv -Path $exportLocation -NoTypeInformation -Force

Import-Module  -Name SharePointPnPPowerShellOnline  -DisableNameChecking
Connect-PnPOnline -Url $siteurl -WarningAction Ignore
$AuditData = Import-CSV $exportLocation
foreach ($Record in $AuditData){
Add-PnPListItem -List "WorkStation-Audit" -Values @{
"Audit_Date"= $Record.'Audit_Date';
"Employee_Name"= $Record.'Employee_Name';
"Computer_Name"= $Record.'Computer_Name';
"Employee_Mail"= $Record. 'Employee_Mail';
"Domain"= $Record.'Domain';
"Admin_Access"= $Record.'Admin_Access';
"Last_ReBoot"= $Record.'Last_ReBoot';
"C_x003a_FreeSpace_GB"= $Record.'C:_FreeSpace_GB';
"Total_Memory_GB"= $Record.'Total_Memory_GB';
"MAC_Address"= $Record.'MAC_Address';
"Processor_Type"= $Record.'Processor_Type';
"User_List"= $Record.'User_List';
"Installed_Apps"= $Record.'Installed_Apps';
"Model"= $Record.'Model';
"System_Type"= $Record.'System_Type';
"Operating_System"= $Record.'Operating_System';
"Operating_System_Version"= $Record.'Operating_System_Version';
"Operating_System_BuildVersion"= $Record.'Operating_System_BuildVersion';
"Serial_Number"= $Record.'Serial_Number';
"Sophos"= $Record.'Sophos';
"Admin_Passwd_Set"= $Record.'Local_Admin_Accounts';
}
}
