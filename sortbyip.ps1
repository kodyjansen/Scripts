$adComputers = Get-ADComputer -Filter * -SearchBase 'OU=WW_IA_DesktopsOU,OU=WW_IA_OU,DC=Worldwide,DC=Local' | Where-Object {$_.DistinguishedName -notmatch 'OU=MD_DesktopsOU|OU=Laptops|OU=DeletedComputers|OU=Converence Rooms|OU=Training Computers'}

$hostname = $adcomputers.name

$arrayList = New-Object -Typename "System.Collections.ArrayList"

Foreach ($object in $hostname){

$null = $arrayList.Add([System.Net.Dns]::GetHostAddresses($object).ipaddresstostring)
}

$sorted = $arrayList | Sort-Object

Foreach($item in $sorted)
{
    Write-Host $item ([System.Net.DNS]::GetHostByAddress($item).hostname)

    
}


