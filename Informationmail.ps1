
Write-Host -foregroundcolor red "COMPUTER INFORMATION:" 
Get-ComputerInfo | Select-Object "WindowsCurrentVersion", "WindowsProductName", "OSName", "OSDisplayVersion", "OSSerialNumber", "CsCaption", "CsDomain", "CsModel", "CsNetworkAdapters", "CsProcessors", "CsSystemFamily", "Cstotalphysicalmemory"
Write-Host -foregroundcolor red "SOFTWARE INSTALLED:
"
Get-AppxPackage –AllUsers | Select-Object -property "Name"
Write-Host -foregroundcolor red "FIREWALL STATUS:
"
get-netfirewallprofile
Write-Host -foregroundcolor red "USAGE:
"
Get-Counter
Write-Host -foregroundcolor red "EVENT LOG (LATEST 20):
"
Get-eventlog -LogName System -Newest 20 | Select-Object "Index", "Time", "EntryType", "Source", "InstanceID", "Message"
Write-Host -ForegroundColor red "USERS
"
get-localuser
get-localgroup | select-object "Name", "Description"
Write-Host -ForegroundColor red "SHARES
"
Get-SmbShare
Write-Host -ForegroundColor red "PRINTERS
"
get-printer | Where-object -property "Type"
Write-host -ForegroundColor red "NETWORK INFORMATION 
" 
Get-NetIPConfiguration | Out-file -FilePath "\\ww-file16-01\ProfileRedirect\kjansen\Documents\green.txt"



