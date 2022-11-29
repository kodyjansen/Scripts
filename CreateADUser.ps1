Function CreateADUser($firstName, $lastName, $jobTitle, $director, $Manager, $phone, $fax, $userPassword, $Extension, $UserName){
        $userPassword = $userPassword | ConvertTo-SecureString -AsPlainText -Force
        $fullName = $firstName + " " + $lastName
        $email = $UserName + "@worldwide-logistics.com"
        $proxyAddress = "SMTP:" + $email
        $directorUserName = (Get-ADUser -Filter 'Name -like $director').samAccountName
        $phoneNumber = "(" + $phone.Substring(0,3) + ")" + $phone.Substring(3,3) + "-" + $phone.Substring(6,4)
        #
        #
        #In Active Directory Create New User, Set Profile Path, Private Path and Proxy Address
        $profilePath    = "\\ww-filesvr01\Profiles2\$UserName"
        $homeDirectory  = "\\ww-filesvr01\Private\$UserName"
        New-ADUser -Name $fullName -GivenName $firstName -Surname $lastName -SamAccountName $UserName -UserPrincipalName $email -AccountPassword $userPassword -EmailAddress $email -DisplayName $fullName -title $jobTitle -OfficePhone $phoneNumber -fax $fax -ProfilePath $profilePath -Manager $directorUserName -Company $Manager -Path "OU=WW_IA_UsersOU,OU=WW_IA_OU,DC=Worldwide,DC=local"
        Set-ADUser -Identity $UserName -Add @{proxyAddresses=$proxyAddress}
        Set-ADUser $UserName -Add @{othertelephone=$Extension}
        Enable-ADAccount -Identity $UserName

        Start-Sleep -Seconds 10
        #
        #
        #Create Profile folder with full permissions
        $profileFolder = $UserName + ".V6"
        $Path = "\\ww-filesvr01\Profiles2\" + $profileFolder
        $inherit = [system.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
        $propagation = [system.security.accesscontrol.PropagationFlags]"None"
        $ar = New-Object system.security.accesscontrol.filesystemaccessrule($UserName,"FullControl", $inherit, $propagation, "Allow")
        New-Item -Path "\\ww-filesvr01\Profiles2\" -Name $profileFolder -ItemType Directory
        $acl = Get-Acl $Path
        $acl.SetAccessRule($ar)
        Set-Acl -Path $Path -AclObject $acl
        Set-Owner -Path $Path -Recurse -Account "WorldWide\$UserName"

        #Create Private Drive Folder and Set it as the Employee's Home Directory
        Set-ADUser -Identity $UserName -HomeDrive "F:" -HomeDirectory $homeDirectory
        #Sync AD to Office 365 Exchange
        $AADComputer	= "WW-ADDS03.worldwide.local"
        $session		= New-PSSession -ComputerName $AADComputer
        Invoke-Command -Session $session -ScriptBlock {Import-Module -Name 'ADSync'}
        Invoke-Command -Session $session -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}
        Remove-PSSession $session
}