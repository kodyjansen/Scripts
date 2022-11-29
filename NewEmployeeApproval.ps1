$RootScript = "\\ww-file16-01\IT\NewEmployeeScript"

. $RootScript\FunctionLibrary\FunctionLibrary.ps1

#$date = Get-Date -Format yyyy-MM-dd

function Send-Email($sendBody, $sendSubject, $sendTo){
    #Email Info
    $BodyMessage    = $sendBody
    $From 			= ""
    $To 			= $sendTo
    $Subject 		= $sendSubject
    $SMTPServer 	= "smtp.office365.com"
    $EmailUsername 	= ''
    $Password 		= Get-Content "" | ConvertFrom-SecureString
    $Creds 			= New-Object System.Management.Automation.PSCredential -ArgumentList $EmailUsername, $Password
    send-mailmessage -from $From -to $To -Subject $Subject -SmtpServer $SMTPServer -Body $BodyMessage -UseSsl -Credential $Creds
}

#Query the Sql table for Results
$queryResultsStart = SQLQuery -type "Select" -parameters "*" -where "Status = 'P'"
#$queryResultsEnd = SQLQuery -type "Select" -parameters "*" -where "(EndDate = '$date' AND Status <> 'R') or Status = 'D'"

if($queryResultsStart){
    foreach($queryResult in $queryResultsStart){
        $mainDirectors = @("Nic Marzen", "Sharla Stephenson", "Shelley Hill", "Brandon Renshaw", "Jeff Slump", "Peter Dugan", "Tim Annett", "Marsha Smothers")
        $geminiDirectors = @("Dan Fowler", "Jami Kitchel")
        ##############################################################
        if($queryResult.Director -in $mainDirectors){
            $phoneNumber = Get-Content \\ww-file16-01\IT\NewEmployeeScript\Data\MainPhoneNumbers.txt -first 1
            $phoneTextFile = Get-Content \\ww-file16-01\IT\NewEmployeeScript\Data\MainPhoneNumbers.txt | where-Object {$_ -notlike "$($phoneNumber)"}
            Set-Content \\ww-file16-01\IT\NewEmployeeScript\Data\MainPhoneNumbers.txt -Value $phoneTextFile
            $phoneBuilding = "Main Building Phone Numbers are Needed"
        }elseif($queryResult.Director -in $geminiDirectors){
            $phoneNumber = Get-Content \\ww-file16-01\IT\NewEmployeeScript\Data\GeminiPhoneNumbers.txt -first 1
            $phoneTextFile = Get-Content \\ww-file16-01\IT\NewEmployeeScript\Data\GeminiPhoneNumbers.txt | where-Object {$_ -notlike "$($phoneNumber)"}
            Set-Content \\ww-file16-01\IT\NewEmployeeScript\Data\GeminiPhoneNumbers.txt -Value $phoneTextFile
            $phoneBuilding = "Gemini Building Phone Numbers are Needed"
        }else{
            $phoneNumber = Get-Content \\ww-file16-01\IT\NewEmployeeScript\Data\DouglasPhoneNumbers.txt -first 1
            $phoneTextFile = Get-Content \\ww-file16-01\IT\NewEmployeeScript\Data\DouglasPhoneNumbers.txt | where-Object {$_ -notlike "$($phoneNumber)"}
            Set-Content \\ww-file16-01\IT\NewEmployeeScript\Data\DouglasPhoneNumbers.txt -Value $phoneTextFile
            $phoneBuilding = "Douglas Building Phone Numbers are Needed" 
        }
        if($phoneNumber){
            $Extension = $phoneNumber.Substring(6,4)
            $userName = $queryResult.Username
            $email = $userName + "@worldwide-logistics.com"
            $pass = New-SimpleRandomPassword -WordListFilePath "\\ww-file16-01\IT\EmployeeApprovalScript\Data\Passwordlist.txt" -WordCount 2
            $firstName = $queryResult.First
            $lastName = $queryResult.Last
            $jobTitle = $queryResult.Title
            $director = $queryResult.Director
            $manager = $queryResult.Manager
            $fax = $queryResult.Fax
            $fullName = $firstName + " " + $lastName
            #
            #
            #Connect to Office 365 -logistics and add license to user
            $O365Username 	= ''
            $O365Password 	= Get-Content "" | ConvertTo-SecureString
            $O365Creds		= New-Object System.Management.Automation.PSCredential -ArgumentList $O365Username, $O365Password

            Connect-MsolService -Credential $O365Creds

            $StandardLicense = ((get-msolaccountsku | Where-Object {$_.AccountSkuID -eq "reseller-account:o365_Business_Premium"}).ActiveUnits) - ((get-msolaccountsku | Where-Object {$_.AccountSkuID -eq "reseller-account:o365_Business_Premium"}).ConsumedUnits)
            $SPBLicense = ((get-msolaccountsku | Where-Object {$_.AccountSkuID -eq "reseller-account:SPB"}).ActiveUnits) - ((get-msolaccountsku | Where-Object {$_.AccountSkuID -eq "reseller-account:SPB"}).ConsumedUnits)
            $MeetingLicense = ((get-msolaccountsku | Where-Object {$_.AccountSkuID -eq "reseller-account:MCOMEETADV"}).ActiveUnits) - ((get-msolaccountsku | Where-Object {$_.AccountSkuID -eq "reseller-account:MCOMEETADV"}).ConsumedUnits)

            if($StandardLicense -lt 10 -and $SPBLicense -lt 10){
                $BodyMessage = "Running low on Office 365 Licenses"
                $Subject = "Running low on Office 365 Licenses"
                $To = "itsystems@worldwide-logistics.com"

                Send-Email -sendBody $BodyMessage -sendSubject $Subject -sendTo $To
            }elseif($MeetingLicense -lt 10){
                $BodyMessage = "Running low on Office 365 Teams Meetings Licenses"
                $Subject = "Running low on Office 365 Teams Meetings Licenses"
                $To = "itsystems@worldwide-logistics.com"
                
                Send-Email -sendBody $BodyMessage -sendSubject $Subject -sendTo $To
            }

            if($StandardLicense -lt 1 -and $SPBLicense -lt 1){
                $BodyMessage = "Out of Office 365 Licenses"
                $Subject = "Out of Office 365 Licenses"
                $To = "itsystems@worldwide-logistics.com"

                Send-Email -sendBody $BodyMessage -sendSubject $Subject -sendTo $To
                exit
            }elseif($MeetingLicense -lt 1){
                $BodyMessage = "Out of Office 365 Teams Meetings Licenses"
                $Subject = "Out of Office 365 Teams Meetings Licenses"
                $To = "itsystems@worldwide-logistics.com"
                
                Send-Email -sendBody $BodyMessage -sendSubject $Subject -sendTo $To
                exit
            }

            SQLQuery -type "Update" -parameters "Status" -values "A" -where "UserName = '$userName' AND Status = 'P'"

            #SQLQuery -table "" -type "Insert" -parameters 

            Start-Sleep -Seconds 60

            $phone = "(" + $phoneNumber.Substring(0,3) + ")" + $phoneNumber.Substring(3,3) + "-" + $phoneNumber.Substring(6,4)

            #Send Email with Credentials for new user
            #$Ptr 			= [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($pass)
            #$result 		= [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
            $BodyMessage    = @"
            Username: $userName
            Password: $pass
    
            Name: $fullName
            Title: $jobTitle
            Director: $director
            Manager: $Manager
            Phone Number: $phone
"@
            $To 			= "newempcred@worldwide-logistics.com"
            $Subject 		= "New Employee Credentials - " + $FullName
            
            Send-Email -sendBody $BodyMessage -sendSubject $Subject -sendTo $To

            Start-Sleep -Seconds 5

            AddO365License -email $email

            Start-Sleep -Seconds 60

            AddGroups -UserName $UserName -jobTitle $jobTitle
        }else{
            #Email Info
            $BodyMessage    = "$phoneBuilding"
            $Subject 		= "$phoneBuilding"
            $To 			= "itsystems@worldwide-logistics.com"
            
            Send-Email -sendBody $BodyMessage -sendSubject $Subject -sendTo $To
            exit
        }

    }
}