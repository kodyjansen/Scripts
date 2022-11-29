Function getMAC()
{
    Param($Computername)

$cim = New-CimSession -ComputerName "BSTANTON-LT" 
Get-CimInstance -CimSession $cim -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" | Select-Object -Property MACAddress

}

getMAC -Computername 'DPATTSCHULL-LT'








$cim = New-CimSession -ComputerName KJANSEN-PC
    $ciminstance = Get-CimInstance -CimSession $cim -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" 
    $ciminstance | Select-Object -Property MACAddress | Tee-object -Variable MAC
    $MAC | Export-Excel -path "C:\Reftab Reports\macstorage.xlsx"
    $macpath = "C:\Reftab Reports\macstorage.xlsx"
    #Open excel report and connect to proper sheet
    $objExcel = New-Object -ComObject Excel.Application  
    $WorkBook = $objExcel.Workbooks.Open($macpath) 
    $WorkSheet = $WorkBook.Sheets.Item(1)
    $macAddress = $Worksheet.Range("A1").EntireColumn.value2
    $macAddress = $macAddress | Where-Object {$_ -ne "MACAddress"}
$newmacAddress = $macAddress.replace(':','-')
$newmacAddress