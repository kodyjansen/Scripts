& "\\ww-file16-01\ProfileRedirect\kjansen\Documents\computername.ps1"
Stop-Service "TimeBrokerSvc"
Sleep 5
Get-Service "Time Broker" | Out-file "\\ww-file16-01\ProfileRedirect\kjansen\Documents\logs.txt"