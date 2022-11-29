#Create Root Script Path
$RootScript = "\\ww-file16-01\IT\HREmployeeMod"

. $RootScript\FunctionLibrary\FunctionLibrary.ps1

#Get XML and Convert it
$inputXML = Get-Content "$RootScript\XAML\HREmpModXml.txt"
$gui = ConvertXaml($inputXML)

#Set Variables from Pulled Files
#Title
function GetJobTitles{
    #$jobTitles = Get-Content "$RootScript\Data\JobTitles.txt"
    #$jobTitles = $jobTitles | Sort-Object
    #$gui.cmbJobTitle.Items.Clear()
    #$jobTitles | Foreach-Object {
    #$gui.cmbJobTitle.Items.Add($_)
    #$gui.cmbJobTitleChange.Items.Add($_)
    #}
    $jobTitles = DevSQLQuery -Type "Select" -parameters "Title" -table "UserTitles" -where "Deleted = 0"
    $jobTitles = $jobTitles.Title | Sort-Object
    $gui.cmbJobTitle.Items.Clear()
    $jobTitles | Foreach-Object {
    $gui.cmbJobTitle.Items.Add($_)
    }
}

GetJobTitles

#Director
function GetDirectors{
    $Directors = Get-Content "$RootScript\Data\Directors.txt"
    $Directors = $Directors | Sort-Object
    $gui.cmbDirector.Items.Clear()
    $Directors | Foreach-Object {
    $gui.cmbDirector.Items.Add($_)
    }
}

GetDirectors

#Direct Reports
function GetDirectReports{
    $DirectReports = Get-Content "$RootScript\Data\DirectReports.txt"
    $DirectReports = $DirectReports | Sort-Object
    $gui.cmbDirectReport.Items.Clear()
    $DirectReports | Foreach-Object {
    $gui.cmbDirectReport.Items.Add($_)
    }
}

GetDirectReports

function GetManagers {
    if ($gui.cmbDirector.SelectedIndex -ne -1){
        $directorRegex = Get-Content \\ww-file16-01\IT\HREmployeeMod\Data\DirectorRegex.txt
        $regex = [regex]::new('/' + $directorRegex + '/g')
        if ($regex -match $gui.cmbDirector.SelectedItem) {
            $gui.cmbManager.Visibility = "Visible"
            $gui.btnRemovemanager.Visibility = "Visible"
            $Managers = Get-Content "$RootScript\Data\$($gui.cmbDirector.SelectedItem)Managers.txt"
        } else {
            $gui.cmbManager.Visibility = "Collapsed"
            $gui.lblManager.Visibility = "Collapsed"
            $gui.btnAddManager.Visibility = "Collapsed"
            $gui.btnRemovemanager.Visibility = "Collapsed"
        }
        $gui.lblManager.Visibility = "Visible"
        $gui.btnAddManager.Visibility = "Visible"
        $Managers = $Managers | Sort-Object
        $gui.cmbManager.Items.Clear()
        $Managers | Foreach-Object {
        $gui.cmbManager.Items.Add($_)
        }
    }
}

#Manager
$gui.cmbDirector.Add_SelectionChanged({
    GetManagers
})

#Add Button Click Event
$gui.btnCreate.Add_Click( {
    if($gui.txtName.text -eq "" -or $gui.cmbJobTitle.SelectedIndex -eq -1 -or $gui.cmbDirector.SelectedIndex -eq -1 -or $gui.cmbDirectReport.SelectedIndex -eq -1 -or $gui.dpStart.SelectedDate -eq $null){
        $gui.lblMessage.text = "All Fields Must Be Populated"
    }
    else{
        $Name = $gui.txtFirstName.Text + " " + $gui.txtLastName.Text
        $FirstName = $gui.txtFirstName.Text
        $FirstName = $FirstName.TrimStart()
        $FirstName = $FirstName.TrimEnd()
        $LastName = $gui.txtLastName.Text
        $LastName = $LastName.TrimStart()
        $LastName = $LastName.TrimEnd()
        $FullName = $FirstName + " " + $LastName
        $jobTitle = $gui.cmbJobTitle.Text
        $director = $gui.cmbDirector.Text
        $Manager = $gui.cmbManager.Text
        $DirectReport = $gui.cmbDirectReport.Text
        $Date = Get-Date 
        $CurrentUser = $env:UserName
        $Rehire = $gui.chkboxRihire.IsChecked

        #Create Variables from Input Credentials (Username, Prinicpal Name, Full Name, email, and Proxy Address)
        $newUserName = $FirstName.Substring(0,1) + $LastName
        $newUserName = $newUserName.ToLower()
        #Check if username exists in active directory, if it does add another letter.
        $firstNameCharacters = 1
        $test = ([adsisearcher]"(&(ObjectCategory=Person)(ObjectClass=User)(samaccountname=$newUserName))").FindAll()
        while ($test -ne $null) {
            $newUserName = $FirstName.Substring(0,$firstNameCharacters) + $LastName
            $newUserName = $newUserName.ToLower()
            $test = ([adsisearcher]"(&(ObjectCategory=Person)(ObjectClass=User)(samaccountname=$newUserName))").FindAll()
            if ($firstNameCharacters -lt $FirstName.Length){
                $firstnameCharacters++
            }
        }

        $Email = $newUserName + "@worldwide-logistics.com"

        if($Manager -eq ""){
            $Manager = $director
        }
        $startDate = '{0:MM-dd-yyyy}' -f $gui.dpStart.SelectedDate

        $From           = "hralerts@worldwide-logistics.com"
        $To             = "newemployees@worldwide-logistics.com"
        $SMTPServer     = "smtp.office365.com"
        $Username       = 'hralerts@worldwide-logistics.com'
        $AESKey         =  Get-Content \\ww-file16-01\it\HREmployeeMod\HREmployeeMod.ps1 -Stream "AES"
        $Password       = Get-Content \\ww-file16-01\it\HREmployeeMod\HREmployeeMod.ps1 -Stream "SEP" | ConvertTo-SecureString -key $AESKey
        $Creds          = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $Password
        $Subject        = $Name
        $BodyMessage    = @"
        Name: $Name
        Start Date: $startDate
        Title: $jobTitle
        Director: $director
        Manager: $Manager
        Direct Report: $DirectReport
        Username: $newUserName
"@
        send-mailmessage -from $From -to $To -Subject $Subject -SmtpServer $SMTPServer -Body $BodyMessage -UseSsl -Credential $Creds

        if($director -eq "Sharla Stephenson"){
            $Fax = "(515)251-6635"
        }else{
            $Fax = "(515)223-6455"
        }

        SQLQuery -table "EmployeeApproval" -type "Insert" -parameters "First", "Last", "Title", "Director", "Manager", "Fax", "StartDate", "Username", "Email", "Status" -values $FirstName, $LastName, $jobTitle, $director, $Manager, $Fax, $startDate, $newUserName, $Email, "P"

        if($Rehire -eq $true){
            DevSQLQuery -table "um_NewHires" -type "Insert" -parameters "FirstName", "LastName", "FullName", "Username", "Status", "Title", "TeamLead", "DirectReport", "HireDate", "Rehired", "CreatedDate", "CreatedBy", "LastModifiedDate", "LastModifiedBy" -values $FirstName, $LastName, $FullName, $newUserName, 1, $jobTitle, $director, $DirectReport, $startDate, 1, $Date, $CurrentUser, $Date, $CurrentUser
        }else{
            DevSQLQuery -table "um_NewHires" -type "Insert" -parameters "FirstName", "LastName", "FullName", "Username", "Status", "Title", "TeamLead", "DirectReport", "HireDate", "Rehired", "CreatedDate", "CreatedBy", "LastModifiedDate", "LastModifiedBy" -values $FirstName, $LastName, $FullName, $newUserName, 0, $jobTitle, $director, $DirectReport, $startDate, 0, $Date, $CurrentUser, $Date, $CurrentUser
        }
        
        if($gui.chkboxTeams.IsChecked){
            $user = $env:UserName
            $NewEmpCheckFolder = "C:\Users\$user\OneDrive - Worldwide Logistics\General\"
            $NewEmpCheckFile = (get-childitem $NewEmpCheckFolder | Where-Object {$_.Name -like '*'+$StartDate+'.xlsx'}).Name        
            if($NewEmpCheckFile){
                $NewEmpCheckPath = $NewEmpCheckFolder + $NewEmpCheckFile
            }else{
                $NewEmpCheckPath = $NewEmpCheckFolder + "New Employee Checklist.xlsx"
            }
            $Excel = New-Object -ComObject Excel.Application
            $ExcelWorkbook = $Excel.Workbooks.Open($NewEmpCheckPath)
    
            $ExcelWorkSheet = $ExcelWorkBook.Sheets.Item('Main')
    
            $NextRow = $ExcelWorkSheet.UsedRange.Rows.Count + 1
    
            $ExcelWorkSheet.Cells.Item($NextRow,1) = "$Name"
            $ExcelWorkSheet.Cells.Item($NextRow,2) = "$jobTitle"
            $ExcelWorkSheet.Cells.Item($NextRow,3) = "$Director"
            $ExcelWorkSheet.Cells.Item($NextRow,4) = "$Manager"
    
            $ExcelWorkSheet2 = $ExcelWorkBook.Sheets.Item('Desk Set Up')
    
            $NextRow2 = $ExcelWorkSheet2.UsedRange.Rows.Count + 1
    
            $ExcelWorkSheet2.Cells.Item($NextRow2,1) = "$Name"
    
            if($NewEmpCheckFile){
                $ExcelWorkBook.Save()
            }else{
                $ExcelWorkBook.SaveAs($NewEmpCheckFolder + "New Employee Checklist $StartDate.xlsx")
            }
    
            $Excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
        }
        
        $gui.txtFirstName.Text = ""
        $gui.txtLastName.Text = ""
        $gui.cmbJobTitle.SelectedIndex = -1
        $gui.cmbManager.SelectedIndex = -1
        $gui.cmbDirector.SelectedIndex = -1
        $gui.cmbDirectReport.SelectedIndex = -1

        $gui.lblMessage.Text = $Name + " Has Been Created"
    }
})

#Add Users into List Box for Deleted Users Tab
$directoryUsers = ([adsisearcher]"(&(ObjectCategory=Person)(ObjectClass=User))").FindAll()
$directoryUsers = $directoryUsers | Where-Object {$_.Properties.Item("distinguishedName") -like "*WW_IA_UsersOU*" -and $_.Properties.Item("distinguishedName") -notlike "*WW_SharedMailboxUsersOU*" -and $_.Properties.Item("distinguishedName") -notlike "*TMC_UsersOU*"}
$allUsers = @()
Foreach($directoryUser in $directoryUsers){
    $allUsers += $directoryUser.properties["Name"]
}
$allUsers = $allUsers | Sort-Object
$gui.listBoxDirectory.Items.Clear()
Foreach($ouUser in $allUsers){
    [void]$gui.listBoxDirectory.items.add($($ouUser))
}

#Search Event for Deleted User Tab
$gui.txtBoxSearch.Add_TextChanged({
    $searchEmployees = @()
    $gui.listBoxDirectory.Items.Clear()
    if ($gui."txtbox$($fieldName)Search".Text -eq "") {
        $searchEmployees = $allUsers
    } else {
        $searchText = $gui.txtBoxSearch.Text
        $searchEmployees = $allUsers | Where-Object {$_ -like "*$searchText*"}
    }
    $searchEmployees = $searchEmployees | Sort-Object
    Foreach($searchEmployee in $searchEmployees){
        [void]$gui."listBox$($fieldName)Directory".Items.Add($searchEmployee)
    }
})

#Delete Selected Person Event
$gui.listBoxDirectory.Add_SelectionChanged({
    if($gui.listBoxDirectory.SelectedIndex -ne -1){
        $DeletedUserDetail = $gui.listBoxDirectory.SelectedItem.ToString()
        $DeletedUserDetails = ([adsisearcher]"(&(ObjectCategory=Person)(ObjectClass=User)(CN=$DeletedUserDetail))").FindAll()
        $gui.txtBoxName.Text = $DeletedUserDetails.properties["Name"]
        $gui.txtBoxPhone.Text = $DeletedUserDetails.properties["TelephoneNumber"]
        $gui.txtBoxEmail.Text = $DeletedUserDetails.properties["Mail"]
        $gui.txtBoxTitle.Text = $DeletedUserDetails.properties["Title"]
        $DeletedDirector = $DeletedUserDetails.Properties['Manager']
        $gui.txtBoxDirector.Text = ([adsi]"LDAP://$DeletedDirector").Name
        $gui.txtBoxManager.Text = $DeletedUserDetails.properties["Company"]
    }else{
        $gui.txtBoxName.Text = ""
        $gui.txtBoxPhone.Text = ""
        $gui.txtBoxEmail.Text = ""
        $gui.txtBoxTitle.Text = ""
        $gui.txtBoxDirector.Text = ""
        $gui.txtBoxManager.Text = ""
    }
})

#Delete Button Click Event
$gui.btnDelete.Add_Click({
    $DeletedName    = $gui.listBoxDirectory.SelectedItem.tostring()
    $From           = "hralerts@worldwide-logistics.com"
    $To             = "employeeterm@worldwide-logistics.com"
    $SMTPServer     = "smtp.office365.com"
    $Username       = 'hralerts@worldwide-logistics.com'
    $AESKey         =  Get-Content \\ww-file16-01\it\HREmployeeMod\HREmployeeMod.ps1 -Stream "AES"
    $Password       = Get-Content \\ww-file16-01\it\HREmployeeMod\HREmployeeMod.ps1 -Stream "SEP" | ConvertTo-SecureString -key $AESKey
    $Creds          = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $Password
    $Subject        = "$DeletedName Term - $date"
    $BodyMessage    = @"
    Hey All,

    Today was $DeletedName's last day here at WorldWide.

    Padraic, please let Night Ops know.

    Thanks,
    HR
"@
    send-mailmessage -from $From -to $To -Subject $Subject -SmtpServer $SMTPServer -Body $BodyMessage -UseSsl -Credential $Creds

    $gui.lblMessage1.Text = $DeletedName + " Has Been Deleted"
})
<#
$script:bAddTitleOpen = $false
#Add Title Button Event
$gui.btnAddTitle.Add_Click({
    if($script:bAddTitleOpen){
        $gui.txtTitle.Visibility = "Collapsed"
        $gui.btnAcceptTitle.Visibility = "Collapsed"
        $script:bAddTitleOpen = $false
    }else{
        $gui.txtTitle.Visibility = "Visible"
        $gui.btnAcceptTitle.Visibility = "Visible"
        $script:bRemoveTitleOpen = $false
        $script:bAddTitleOpen = $true
    }
})

$script:bRemoveTitleOpen = $false
#Remove Title Button Event
$gui.btnRemoveTitle.Add_Click({
    if($script:bRemoveTitleOpen){
        $gui.btnAcceptTitle.Visibility = "Collapsed"
        $script:bRemoveTitleOpen = $false
    }else{
        $gui.txtTitle.Visibility = "Collapsed"
        $gui.btnAcceptTitle.Visibility = "Visible"
        $script:bAddTitleOpen = $false
        $script:bRemoveTitleOpen = $true
    }
})

#Accept Title Button Event
$gui.btnAcceptTitle.Add_Click({
    if($script:bAddTitleOpen){
        if($gui.txtTitle.text -ne ""){
            Add-Content \\ww-file16-01\IT\HREmployeeMod\Data\JobTitles.txt -Value $gui.txtTitle.text
            GetJobTitles
            $gui.txtTitle.Visibility = "Collapsed"
            $gui.btnAcceptTitle.Visibility = "Collapsed"
            $script:bAddTitleOpen = $false
            $gui.txtTitle.text = ""
        }
    }elseif($script:bRemoveTitleOpen){
        if($gui.cmbJobTitle.SelectedIndex -ne -1){
            $JobTextFile = Get-Content \\ww-file16-01\IT\HREmployeeMod\Data\JobTitles.txt | where-Object {$_ -notlike "$($gui.cmbJobTitle.text)"}
            Set-Content \\ww-file16-01\IT\HREmployeeMod\Data\JobTitles.txt -Value $JobTextFile 
            GetJobTitles
            $gui.btnAcceptTitle.Visibility = "Collapsed"
            $script:bRemoveTitleOpen = $false
        }
    }
})
#>
$script:bAddDirectorOpen = $false
#Add Director Button Event
$gui.btnAddDirector.Add_Click({
    if($script:bAddDirectorOpen){
        $gui.txtDirector.Visibility = "Collapsed"
        $gui.btnAcceptDirector.Visibility = "Collapsed"
        $script:bAddDirectorOpen = $false
    }else{
        $gui.txtDirector.Visibility = "Visible"
        $gui.btnAcceptDirector.Visibility = "Visible"
        $script:bRemoveDirectorOpen = $false
        $script:bAddDirectorOpen = $true
    }
})

$script:bRemoveDirectorOpen = $false
#Remove Director Button Event
$gui.btnRemoveDirector.Add_Click({
    if($script:bRemoveDirectorOpen){
        $gui.btnAcceptDirector.Visibility = "Collapsed"
        $script:bRemoveDirectorOpen = $false
    }else{
        $gui.txtDirector.Visibility = "Collapsed"
        $gui.btnAcceptDirector.Visibility = "Visible"
        $script:bAddDirectorOpen = $false
        $script:bRemoveDirectorOpen = $true
    }
})

#Accept Director Button Event
$gui.btnAcceptDirector.Add_Click({
    if($script:bAddDirectorOpen){
        if($gui.txtDirector.text -ne ""){
            Add-Content \\ww-file16-01\IT\HREmployeeMod\Data\Directors.txt -Value $gui.txtDirector.text
            GetDirectors
            $gui.txtDirector.Visibility = "Collapsed"
            $gui.btnAcceptDirector.Visibility = "Collapsed"
            $script:bAddDirectorOpen = $false
            $gui.txtDirector.text = ""
        }
    }elseif($script:bRemoveDirectorOpen){
        if($gui.cmbDirector.SelectedIndex -ne -1){
            $DirectorTextFile = Get-Content \\ww-file16-01\IT\HREmployeeMod\Data\Directors.txt | where-Object {$_ -notlike "$($gui.cmbDirector.text)"}
            Set-Content \\ww-file16-01\IT\HREmployeeMod\Data\Directors.txt -Value $DirectorTextFile 
            GetDirectors
            $gui.btnAcceptDirector.Visibility = "Collapsed"
            $script:bRemoveDirectorOpen = $false
        }
    }
})

$script:bAddDirectReportOpen = $false
#Add Direct Report Button Event
$gui.btnAddDirectReport.Add_Click({
    if($script:bAddDirectReportOpen){
        $gui.txtDirectReport.Visibility = "Collapsed"
        $gui.btnAcceptDirectReport.Visibility = "Collapsed"
        $script:bAddDirectReportOpen = $false
    }else{
        $gui.txtDirectReport.Visibility = "Visible"
        $gui.btnAcceptDirectReport.Visibility = "Visible"
        $script:bRemoveDirectReportOpen = $false
        $script:bAddDirectReportOpen = $true
    }
})

$script:bRemoveDirectReportOpen = $false
#Remove Direct Report Button Event
$gui.btnRemoveDirectReport.Add_Click({
    if($script:bRemoveDirectReportOpen){
        $gui.btnAcceptDirectReport.Visibility = "Collapsed"
        $script:bRemoveDirectReportOpen = $false
    }else{
        $gui.txtDirectReport.Visibility = "Collapsed"
        $gui.btnAcceptDirectReport.Visibility = "Visible"
        $script:bAddDirectReportOpen = $false
        $script:bRemoveDirectReportOpen = $true
    }
})


#Accept Direct Report Button Event
$gui.btnAcceptDirectReport.Add_Click({
    if($script:bAddDirectReportOpen){
        if($gui.txtDirectReport.text -ne ""){
            Add-Content \\ww-file16-01\IT\HREmployeeMod\Data\DirectReports.txt -Value $gui.txtDirectReport.text
            GetDirectReports
            $gui.txtDirectReport.Visibility = "Collapsed"
            $gui.btnAcceptDirectReport.Visibility = "Collapsed"
            $script:bAddDirectReportOpen = $false
            $gui.txtDirectReport.text = ""
        }
    }elseif($script:bRemoveDirectReportOpen){
        if($gui.cmbDirectReport.SelectedIndex -ne -1){
            $DirectReportTextFile = Get-Content \\ww-file16-01\IT\HREmployeeMod\Data\DirectReports.txt | where-Object {$_ -notlike "$($gui.cmbDirectReport.text)"}
            Set-Content \\ww-file16-01\IT\HREmployeeMod\Data\DirectReports.txt -Value $DirectReportTextFile 
            GetDirectReports
            $gui.btnAcceptDirectReport.Visibility = "Collapsed"
            $script:bRemoveDirectReportOpen = $false
        }
    }
})

$script:bAddManagerOpen = $false
#Add Manager Button Event
$gui.btnAddManager.Add_Click({
    if($script:bAddManagerOpen){
        $gui.txtManager.Visibility = "Collapsed"
        $gui.btnAcceptManager.Visibility = "Collapsed"
        $script:bAddManagerOpen = $false
    }else{
        $gui.txtManager.Visibility = "Visible"
        $gui.btnAcceptManager.Visibility = "Visible"
        $script:bRemoveManagerOpen = $false
        $script:bAddManagerOpen = $true
    }
})

$script:bRemoveManagerOpen = $false
#Remove Manager Button Event
$gui.btnRemoveManager.Add_Click({
    if($script:bRemoveManagerOpen){
        $gui.btnAcceptManager.Visibility = "Collapsed"
        $script:bRemoveManagerOpen = $false
    }else{
        $gui.txtManager.Visibility = "Collapsed"
        $gui.btnAcceptManager.Visibility = "Visible"
        $script:bAddManagerOpen = $false
        $script:bRemoveManagerOpen = $true
    }
})

#Accept Manager Button Event
$gui.btnAcceptManager.Add_Click({
    if($script:bAddManagerOpen){
        if($gui.txtManager.text -ne ""){
            Add-Content \\ww-file16-01\IT\HREmployeeMod\Data\$($gui.cmbDirector.text)Managers.txt -Value $gui.txtManager.text
            $directorRegex = Get-Content \\ww-file16-01\IT\HREmployeeMod\Data\DirectorRegex.txt
            $regex = [regex]::new('/' + $directorRegex + '/g')
            if ($regex -notmatch $gui.cmbDirector.text){
                Add-Content \\ww-file16-01\IT\HREmployeeMod\Data\DirectorRegex.txt -Value "|$($gui.cmbDirector.text)"
            }
            GetManagers
            $gui.txtManager.Visibility = "Collapsed"
            $gui.btnAcceptManager.Visibility = "Collapsed"
            $script:bAddManagerOpen = $false
            $gui.txtManager.text = ""
        }
    }elseif($script:bRemoveManagerOpen){
        if($gui.cmbDirector.SelectedIndex -ne -1){
            $ManagerTextFile = Get-Content \\ww-file16-01\IT\HREmployeeMod\Data\$($gui.cmbDirector.text)Managers.txt | where-Object {$_ -notlike "$($gui.cmbManager.text)"}
            Set-Content \\ww-file16-01\IT\HREmployeeMod\Data\$($gui.cmbDirector.text)Managers.txt -Value $ManagerTextFile
            GetManagers
            $gui.btnAcceptManager.Visibility = "Collapsed"
            $script:bRemoveManagerOpen = $false
        }
    }
})

$date = Get-Date
$gui.dpStart.DisplayDateStart = $date

$gui.form.ShowDialog() | out-null