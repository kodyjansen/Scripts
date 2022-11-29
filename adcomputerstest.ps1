$desktopsearchbase = 'OU=WW_IA_DesktopsOU,OU=WW_IA_OU,DC=Worldwide,DC=Local'
$adComputers = Get-ADComputer -Filter * -SearchBase $desktopsearchbase | Where-Object {$_.DistinguishedName -notmatch 'OU=MD_DesktopsOU|OU=Laptops|OU=DeletedComputers|OU=Converence Rooms|OU=Training Computers'}

Foreach($computer in $adcomputers)
{
  $computer.name
}









