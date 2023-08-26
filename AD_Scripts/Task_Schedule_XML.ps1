#Importing tasks in task scheduler

$XmlFilePath = "C:\Users\ravi.morampudi\OneDrive - EPSoft®\Desktop\SystemAudit.xml"
$TaskName = "SystemAudit"

 

# Import the task using schtasks.exe
Start-Process schtasks.exe -ArgumentList "/Create /XML `"$XmlFilePath`" /TN `"$TaskName`"" -Wait -NoNewWindow