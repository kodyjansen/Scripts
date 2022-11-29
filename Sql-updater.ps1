function ReportData
{
        #Reftab Report paths
    $path = Get-ChildItem -path "C:\Reftab Reports\report.xlsx"
    $newmacfilepath = "C:\Reftab Reports\macstorage.txt"
    $querylog = "C:\Reftab Reports\querylog.txt"

        #Clear Query Log
    Clear-Content $querylog

        #Open excel report and connect to proper sheet
    $objExcel = New-Object -ComObject Excel.Application  
    $WorkBook = $objExcel.Workbooks.Open($path) 
    $WorkSheet = $WorkBook.Sheets.Item(1)

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


        #TEST CASES/FILTER NAMES
if($wherevalue -eq $null)
    {"Missing name for $propertyvalue`n"  >> $querylog
        continue
    }
elseif($propertyvalue -eq $null)
    {"Missing computer name for $wherevalue`n"  >> $querylog
        continue
    }
elseif($wherevalue -eq 'Jeff Slump')
    {"Did not update data for $wherevalue`n"  >> $querylog
        continue
    }
elseif($wherevalue -eq 'Peter Dugan')
    {"Did not update data for $wherevalue`n"  >> $querylog
        continue
    }
elseif($wherevalue -eq 'Mitch Annett')
    {"Did not update data for $wherevalue`n"  >> $querylog
        continue
    }


        #SQL Query to change ComputerName
    $sqlcmd.CommandText = "UPDATE dbo.SeatTable SET $property = '$propertyvalue' WHERE $whereproperty = '$wherevalue'" 
    $sqlcmd.ExecuteNonQuery() | Out-Null

         #Writes Output to Query Log
     "$wherevalue's Computername has been changed to $propertyvalue in SeatTable"  >> $querylog

        #GET MAC
Try{
    $hostname = $prop\ertyvalue
    $cim = New-CimSession -ComputerName $hostname -Credential Get-Credential 
    $ciminstance = Get-CimInstance -CimSession $cim -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" 
    $ciminstance | Select-Object -Property MACAddress | Out-file -filepath $newmacfilepath 
    Get-Content -path $newmacfilepath | Tee-object -variable newMAC | Out-Null
   }
Catch{"Error Finding MAC for $hostname`n" >> $querylog
       continue 
     }   

Try{
    $newMAC = "'$newMAC'"
    $newMAC = $newMAC.replace(":", "-")
    $newMAC = $newMAC.replace("MACAddress", "") 
    $newMAC = $newMAC.replace(' ', '')
    $newMAC = $newMAC.Substring($newMAC.Length - 18, 17)
   } 
Catch{"Current MAC address is $newMAC" >> $querylog}

        #SQL Query to change MAC
    $sqlcmd.CommandText = "UPDATE dbo.SeatTable SET MAC = '$newMAC' WHERE $whereproperty = '$wherevalue'" 
    $sqlcmd.ExecuteNonQuery() | Out-Null

        #Writes Output to Query Log
     "$wherevalue's MAC has been changed to $newMAC in SeatTable`n"  >> $querylog

  }
    
}
ReportData 