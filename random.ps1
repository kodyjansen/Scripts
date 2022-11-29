function renamePCs{
    #$adComputers = Get-ADComputer -Filter * -SearchBase 'OU=WW_IA_DesktopsOU,OU=WW_IA_OU,DC=Worldwide,DC=Local' | Where-Object {$_.DistinguishedName -notmatch 'OU=MD_DesktopsOU|OU=Laptops|OU=DeletedComputers|OU=Converence Rooms|OU=Training Computers'}
    #$adUsers = Get-ADUser -Filter * -SearchBase 'OU=WW_IA_OU,DC=Worldwide,DC=Local' | Where-Object {$_.DistinguishedName -notmatch 'OU=WW_DeletedUsersOU|OU=TMC_UsersOU|OU=WW_ServiceUserOU'}
    $adComputers = Get-ADComputer -Filter * -SearchBase 'OU=WW_IA_DesktopsOU,OU=WW_IA_OU,DC=Worldwide,DC=Local' | Where-Object {$_.DistinguishedName -match 'OU=TestUpdateGroup'}
    $adUsers = Get-ADUser -Filter * -SearchBase 'OU=WW_IA_OU,DC=Worldwide,DC=Local' | Where-Object {$_.DistinguishedName -match 'OU=WW_DeletedUsersOU|OU=TMC_UsersOU|OU=WW_ServiceUserOU'}
    $errorUsers = [System.Collections.ArrayList]::new()
    #$adminCreds = Get-Credential
    
    #Logic Variables
    $isFound = $false

    foreach($user in $adUsers){

        #Gets the first and last name of the user
        $firstName = $user.GivenName
        $lastName = $user.Surname
        
        #Trys to build the host name from the given first and last name
        try{
            $oldHostName = ("WW-" + $firstName.Substring(0,7) + $lastName.Substring(0,2) + "-PC").ToUpper()
        }
        catch{
            try{
                 $oldHostName = ("WW-" + $firstName + $lastName.Substring(0,2) + "-PC").ToUpper()
            }
            catch{
                $errorUsers.Add($user.Name)
                $oldHostName = $null
            }
        }

        #If the script was able to create a hostname, it checks to see if that hostname exists
        #If the hostname exists in ADComputers then it renames it based on the new naming protocol (first initial 
        if($oldHostName -ne $null){
            $newHostName = ($firstName.Substring(0,1) + $lastName + "-PC").ToUpper()
            $username = $newHostName.Substring(0, $newHostName.Length - 3)

            foreach($computer in $adComputers){
                if($computer.Name -eq $oldHostName){

                    Write-Host ("User: " + $user.Name + " NewHostName: " + $newHostName + " OldHostName: " + $oldHostName)
                    #Rename-Computer -ComputerName $oldHostName -DomainCredential $adminCreds -NewName $newHostName
                    #createRPD -hostname $newHostName -username $username
                    $isFound = $true
                    break
                }
            }

            if($isFound -eq $false){
                $null = $errorUsers.Add($user.Name)
            }
            $isFound = $false
        }
        
    }
    $errorUsers
}

renamePCs