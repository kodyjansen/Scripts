#Active Directory Folders
$searchbase = 'OU=Laptops,OU=WW_IA_DesktopsOU,OU=WW_IA_OU,DC=Worldwide,DC=Local'
$adComputers = Get-ADComputer -Filter * -SearchBase $searchbase | Where-Object {$_.DistinguishedName -notmatch 'OU=MD_DesktopsOU|OU=DeletedComputers|OU=Conference Rooms|OU=Training Computers'}

#Creates an array
$laptops = @()
#Adds each laptop found in ADLaptops as a new item in the array
foreach($computer in $adcomputers)
{
     $laptops = $laptops + $computer.name
    
}
#Calculates the amount of items in the array
$laptopcount = "Total Laptop Count: " + $laptops.length 
#Writes the total to a file
$laptopcount | Out-file -filepath "\\Profile\Users\kjansen\Desktop\Laptops.txt"


