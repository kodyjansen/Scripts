Import-Module NetTCPIP
function getMac(){       
    Param(
        $ComputerName 
    )
    #Gets the active IP address then finds the interface alias associated with it
    $NetIpConfig = ((Get-NetIPConfiguration -ComputerName $ComputerName ))#| Where-Object {$null -ne $_.IPv4DefaultGateway -and $_.NetAdapter.Status -ne "Disconnected"}))
    $targetIp = $NetIpConfig.IPv4Address.IPAddress
    $InterfaceAlias = $NetIpConfig[[array]::indexof($NetIpConfig, $targetIp)].IPv4Address.InterfaceAlias
    $InterfaceAlias
    #GetMac and trim down to only the MAC
    $script:MAC = get-netadapter -CimSession $ComputerName #-Name $InterfaceAlias #| Format-List -Property MacAddress
    $script:MAC
    #$script:MAC = out-string -InputObject $script:MAC
    #$script:MAC = $script:MAC.Trim()
    #Write-Host ($ComputerName + " " + $script:MAC)
    #$script:MAC = $script:MAC.Split(" ")[2]
    #$script:MAC | Out-File -FilePath 'C:\VisualStudio\SeatPortMap Admin Tools\SeatPortMap Admin Tools\MAC.txt'
}
getMac -ComputerName 'MSMOTHERS-LT'