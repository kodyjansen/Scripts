#GetMAC Function
Import-Module NetTCPIP
function getMac(){       
    Param(
        $ComputerName 
    )
    $NetIpConfig = ((Get-NetIPConfiguration -ComputerName $ComputerName ))#| Where-Object {$null -ne $_.IPv4DefaultGateway -and $_.NetAdapter.Status -ne "Disconnected"}))
    $targetIp = $NetIpConfig.IPv4Address.IPAddress
    $InterfaceAlias = $NetIpConfig[[array]::indexof($NetIpConfig, $targetIp)].IPv4Address.InterfaceAlias
    $InterfaceAlias
    $script:MAC = get-netadapter -CimSession $ComputerName #-Name $InterfaceAlias #| Format-List -Property MacAddress
    $script:MAC

}

#MAIN VARIABLE FOR DATA UPDATE
$Computername = 'KJANSEN-LT'

#Variable Declaration
$ConnectionList = getMac -Computername $computerName

#Searches MAC's to find the Connected/Up Connection
Foreach($connection in $ConnectionList)
{
  If($connection.status -Contains 'Up')
    {
        $connectedMAC = $connection.MacAddress
        $connectedMAC
    }
}

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
        "Dev"{
            $ConnectionString = "Data Source=WW-IA-SQLDB06;Initial Catalog=threepl;User ID=Svrc_newhire;Password=W@terBl@5tDunk!;Connect Timeout=30"
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
    switch ($type){
        "Select"{
            if($where){
                $query = "$type $parameters From $table Where $where"
            }else{
                $query = "$type $parameters From $table"
            }
            $result = ExecQuery -query $query -conn $conn
            $Conn.Close()
            return $result
        }