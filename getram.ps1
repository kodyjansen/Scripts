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
          Write-host "Found"
       }
    } 
#Pulls hostname (First 7 Last 2) from ADComputers
$desktopsearchbase = 'OU=WW_IA_DesktopsOU,OU=WW_IA_OU,DC=Worldwide,DC=Local'
$adComputers = Get-ADComputer -Filter * -SearchBase $desktopsearchbase | Where-Object {$_.DistinguishedName -notmatch 'OU=MD_DesktopsOU|OU=Laptops|OU=DeletedComputers|OU=Converence Rooms|OU=Training Computers'}

#Turns Parameter into Only First Name (This is to match it with the HostName Entry)
$entryindexof = $entry.indexof(" ")
$newentry = $entry.substring(0, $entryindexof)

    
#Creates Only First Name of the User from the HostName(HARDCODED)
 Try{   
    foreach($computer in $adcomputers)
        {
            $computername = $computer.name
            $indexof = $computername.lastindexof("-") - 5
            $firstname = $computername.substring(3, $indexof)

#Checks if the Parameter contains the Firstname
        If($newentry -like $firstname)
            {
                Write-Host "True"
            }

        }
            
    }
Catch{}

  



}
getRam "Isaac Estell"  
