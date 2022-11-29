Function getRam
{
    param($entry)

#Edits Entry to Match Hostname Format (HARDCODED)
    $entrylength = $entry.indexof(" ") + 3
    $newentry = $entry.substring(0,$entrylength) 
    $newentry = $newentry.replace(" ", "")
    $hostentry = "WW-"+ $newentry + "-PC"
    $entry

#Gets all Hostnames in ADComputers
$desktopsearchbase = 'OU=WW_IA_DesktopsOU,OU=WW_IA_OU,DC=Worldwide,DC=Local'
$adComputers = (Get-ADComputer -Filter * -SearchBase $desktopsearchbase | Where-Object {$_.DistinguishedName -notmatch 'OU=MD_DesktopsOU|OU=Laptops|OU=DeletedComputers|OU=Converence Rooms|OU=Training Computers'}).name

$credential = "adminisaac"

#Checks if the Parameter Matches the Hostname
 Foreach($computer in $adcomputers)
        {
            $hostname = $computer
            If($hostentry.toupper() -like $hostname)
            {
                $cim = New-CimSession -ComputerName $hostname -Credential $credential
                Get-CimInstance -CimSession $cim -Class Win32_ComputerSystem | Select-Object -property TotalPhysicalMemory, Model
            }
            
        }
}

getRam "Isaac Estell"