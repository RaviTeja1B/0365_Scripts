$taskAction = New-ScheduledTaskAction -Execute "C:\MyScripts\domain_join.bat"
$taskTrigger = New-ScheduledTaskTrigger -AtLogOn
$taskTrigger.EndBoundary = (Get-Date).AddMinutes(5).ToString("yyyy-MM-ddTHH:mm:ss")
$taskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopOnIdleEnd
$task = Register-ScheduledTask -TaskName "DomainJoinTask" -Action $taskAction -Trigger $taskTrigger -Settings $taskSettings
Set-ScheduledTask -InputObject $task