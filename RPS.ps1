Function RPS($inputTest)
{
    param($inputtest)
        $choice = "Rock", "Paper", "Scissors"
       
        $choice = Get-random -InputObject $choice
        $choice
        
        
        If ($inputTest -eq "Rock" -and $choice -eq "Scissors")
        {
            Write-host "You Win!"
        }
        Elseif ($inputTest -eq "Paper" -and $choice -eq "Rock")
        {
            Write-Host "You Win!"
        }
        Elseif($inputTest -eq "Scissors" -and ($choice -eq "Paper"))
        {
            Write-Host "You Win!"
        }
        Elseif($choice -eq $inputTest)
        {
            Write-Host "Tie"
        }
        Else {Write-Host "You Lose!"}
    
    
       
}