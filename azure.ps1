
        #Reftab Report paths
    $path = Get-ChildItem -path "C:\Reftab Reports\report.xlsx"
    $newmacfilepath = "C:\Reftab Reports\macstorage.txt"
    $querylog = "C:\Reftab Reports\querylog.txt"

        #Clear Query Log
    Clear-Content $querylog

        #Open excel report and connect to proper sheet
    $objExcel = New-Object -ComObject Excel.Application  
    $WorkBook = $objExcel.Workbooks.Open($path) 
    $WorkSheet = $WorkBook.Sheets.Item(2)

        #Gathers computername needed for SQL query
    $computername = $Worksheet.columns.item(2).value2
    $computername = $computername | Where-Object {$_ -ne "Title"}

        #Gathers name needed for SQL query
    $name = $Worksheet.columns.item(4).value2
    $name = $name | Where-Object {$_ -ne "Loanee" -and $_ -ne ""}

        #Formula for indexing through excel sheet
    $Range = $WorkSheet.UsedRange
    $Rows = $Range.rows | Select-Object -property Row
    $Rows = $Rows.count - 1 


        #Create SQL Connection
    $con = new-object System.data.sqlclient.SQLconnection

        #Set Connection String
    $con.ConnectionString ="Data Source=WW-SEAT-SVR;Initial Catalog=WW-SEAT-DB;User ID=Seat;Password=Chartblock38;Connect Timeout=30"
    $con.open()

        #Connect to SQL
    $sqlcmd = new-object System.data.sqlclient.sqlcommand
    $sqlcmd.connection = $con
    $sqlcmd.CommandTimeout = 600000

 

    #Begin Looping
For($i=0;$i -lt $Rows; $i++)
{
       #Query Variable Declaration
    $property = 'ComputerName'
    $propertyvalue = $computername[$i]
    $whereproperty = 'Name'
    $wherevalue = $name[$i]

$hostname = $propertyvalue
Try{
    $cim = New-CimSession -ComputerName $hostname 
    $ciminstance = Get-CimInstance -CimSession $cim -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" 
    $ciminstance | Select-Object -Property MACAddress | Out-file -filepath $newmacfilepath 
    Get-Content -path $newmacfilepath | Tee-object -variable newMAC | Out-Null

    $newMAC = "'$newMAC'"
    $newMAC = $newMAC.replace(":", "-")
    $newMAC = $newMAC.replace("MACAddress", "") 
    $newMAC = $newMAC.replace(' ', '')
    $newMAC = $newMAC.Substring($newMAC.Length - 18, 17)
}
Catch{Write-host "Could not create WinRM session for $hostname"
    continue}



 #Check to see if data has already been updated
    $Sqlcmd.commandtext = "SELECT MAC FROM dbo.SeatTable WHERE Name = '$wherevalue'" 
    $macCheck = $SqlCmd.ExecuteScalar()
    

if($macCheck = $newMAC)
{
    Write-Host "Duplicate MAC: $macCheck, $newMAC, $wherevalue"
}    


}