$searchbase = 'OU=WW_IA_DesktopsOU,OU=WW_IA_OU,DC=Worldwide,DC=Local'
$targetOU = 'OU=Laptops,OU=WW_IA_DesktopsOU,OU=WW_IA_OU,DC=Worldwide,DC=Local'
$adComputers = Get-ADComputer -Filter * -SearchBase $searchbase | Where-Object {$_.DistinguishedName -notmatch 'OU=TestUpdateGroup|OU=CarefulUpdateGroup|OU=MD_DesktopsOU|OU=Laptops|OU=DeletedComputers|OU=Conference Rooms|OU=Training Computers|OU=Test Computer Group'}
<#
$User = "adminkody"
$File = "C:\Users\kjansen\Scripts\standalone.ps1"
$Credential=New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, (Get-Content $File | ConvertTo-SecureString)
#>






Foreach($computer in $adComputers)
    {
        $computername=$computer.name | Sort-Object




       $lastindexof = $computername.lastindexof('-')
    Try{
        $laptopidentifier = $computer.name.substring($lastindexof, $computername.length - $lastindexof)
       }
    Catch{}
 
  If($laptopidentifier -eq "-LT")
        {
            $laptopname = $computername
            $laptopname
            Get-adcomputer $laptopname | Move-ADObject -targetpath $targetOU -credential $credential 
        } 
    
    }