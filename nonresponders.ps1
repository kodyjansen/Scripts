$desktopsearchbase = 'OU=WW_IA_UsersOU,OU=WW_IA_OU,DC=Worldwide,DC=Local'
$adUser = Get-ADUser -Filter * -SearchBase $desktopsearchbase | Where-Object {$_.DistinguishedName -notmatch 'OU=MD_UsersOU|OU=TMC_UsersOU|OU=WW_DeletedUsersOU|OU=WW_ServiceUserOU|OU=WW_SharedMailboxUsersOU'}


 $path = Get-ChildItem -path "C:\Users\kjansen\Downloads\Wifi_(1-333).xlsx"


    $objExcel = New-Object -ComObject Excel.Application  
    $WorkBook = $objExcel.Workbooks.Open($path) 
    $WorkSheet = $WorkBook.Sheets.Item(1)

    $username = $Worksheet.columns.item(5).value2
    $username = $username | Where-Object {$_ -ne "Name"}



Foreach($user in $adUser)
{

$writeto = @($user.name | ?{$username -notcontains $_})



$writeto #| Out-file -filepath "\\Profile\Users\kjansen\Desktop\NUDGE.txt"
}


