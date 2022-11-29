Function getRam
{
   param($entry)

#Gets all of the users in ADUsers but filters for Full Name
$searchbase='OU=WW_IA_UsersOU,OU=WW_IA_OU,DC=Worldwide,DC=Local'
$adUsersFullNames = (Get-ADUser -Filter * -SearchBase $searchbase | Where-Object {$_.DistinguishedName -notmatch 'OU=MD_DesktopsOU|OU=Laptops|OU=DeletedComputers|OU=Converence Rooms|OU=Training Computers'}).name

#Loops through SearchBase
   Foreach($fullname in $adUsersFullNames)
    {
#Checks if the parameter is found within the SearchBase
     If($fullname -eq $entry)
       {
          Write-host ""
       }
    } 
#Gets all Hostnames in ADComputers
    $desktopsearchbase = 'OU=WW_IA_DesktopsOU,OU=WW_IA_OU,DC=Worldwide,DC=Local'
    $adComputers = (Get-ADComputer -Filter * -SearchBase $desktopsearchbase | Where-Object {$_.DistinguishedName -notmatch 'OU=MD_DesktopsOU|OU=Laptops|OU=DeletedComputers|OU=Converence Rooms|OU=Training Computers'}).name

#Changes Parameter to Match Hostname
    $entrylength = $entry.indexof(" ") + 3
    $newentry = $entry.substring(0,$entrylength) 
    $newentry = $newentry.replace(" ", "")
    


#Creates Edited Hostname
    Foreach($computer in $adcomputers)
        {
            $hostnamelength = $computer.lastindexof("-") - 3
            $hostname = $computer.substring(3, $hostnamelength)
            
        
 
#Checks if the Parameter contains the Hostname
        If($newentry.toupper() -like $hostname)
            {
                Write-Host "True"
                Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $hostname
            }

        }
   


}
getRam "Isaac Estell"  




