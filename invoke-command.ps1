
$credentials = Get-credential
Invoke-Command -Computername "WW-SEANWE-PC" -Scriptblock {ipconfig /renew} -credential $credentials


























