$destinationfilecreate = New-Item -path $destinationpath "Monday.txt" -itemtype "file"
$contents | Out-file -filepath $destinationfilecreate 

    If(Test-Path $destinationfilecreate -name "Monday.txt" -eq false)
    {
        $destinationfilecreate = New-Item -path $destinationpath -name "Monday.txt" -itemtype "file"     
    }
    Else
    {
        $contents | Out-file -filepath $destinationfilecreate
    }


        