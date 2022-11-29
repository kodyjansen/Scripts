#Establish connection to SQL server
function ConnectSQL ($DB) {
    switch($DB){
        "Seat"{
            #SQL Creds
            $sqlAesKey = Get-Content \\ww-file16-01\it\forwardscript\FunctionLibrary\SQLFunctions.ps1 -Stream "SqlAES"
            $sqlPassword = Get-Content \\ww-file16-01\it\forwardscript\FunctionLibrary\SQLFunctions.ps1 -Stream "SqlAEP" | convertto-securestring -key $sqlAesKey
            $creds = New-Object System.Management.Automation.PSCredential -ArgumentList "nobodyCares", $sqlPassword
        
            #$ConnectionString = "Data Source=WW-SEAT-SVR;Initial Catalog=WW-SEAT-DB;User ID=Seat;Password=$($creds.GetNetworkCredential().Password);Connect Timeout=30"
            $ConnectionString = "Data Source=WW-SEAT-SVR;Initial Catalog=WW-SEAT-DB;User ID=Seat;Password=Chartblock38;Connect Timeout=30"
            $conn = new-object System.Data.SqlClient.SQLConnection
            $conn.ConnectionString = $ConnectionString
            $conn.Open()
            Return $conn
            } 
    }
}

#Create Query
function SQLQuery($type, $parameters, $values, $table = "SeatTable", $where){
    $conn = ConnectSQL("Seat")
    {
        $type = 'Update'
        $parameters = '[Phone]'
        $values = '515-644-4453'
        $where = "[Name] = 'Kody Jansen'"

        {
            $query = "$type $table SET $parameters = '$values' Where $where"   
        }
       
    
}
    $cmd = new-object system.Data.SqlClient.SqlCommand($query,$conn)
    if($type -eq "Insert"){
        for($index = 0;$index -lt $count;$index++){
            if($parameters[$index] -like "*Date"){
                $cmd.Parameters.Add("@$($parameters[$index])", [System.Data.SqlDbType]::DateTime).Value = $Values[$index]
            }else{
                $cmd.Parameters.Add("@$($parameters[$index])", [System.Data.SqlDbType]::Char).Value = $Values[$index]
            }
        }
    }
    ExecNonQuery -cmd $cmd
    $Conn.Close()
}


#Execute NonQuery
function ExecNonQuery($cmd){
    $cmd.ExecuteNonQuery() | out-null
    $cmd.Dispose()
}

#Execute Query 
function ExecQuery($query,$conn){
    $cmd=new-object system.Data.SqlClient.SqlCommand($Query,$conn)
    $ds=New-Object system.Data.DataSet
    $da=New-Object system.Data.SqlClient.SqlDataAdapter($cmd)
    [void]$da.fill($ds)
    $cmd.dispose()
    return $ds.Tables[0]
}