function swap{
    Param(
        $file1, $file2
    )

    $filePath1 = $file1.substring(0, $file1.LastIndexOf('\') + 1)
    $filePath2 = $file2.substring(0, $file2.LastIndexOf('\') + 1)

    Write-host $filePath1
    Write-Host $filePath2

    Move-Item $file1 $filePath2
    Move-Item $file2 $filePath1

    Write-Host "Program Completed."
}

function swapFromFile{
    Param(
        $swapInfoDoc
    )
    
    $swapText = Get-Content $swapInfoDoc
    $file1 = $swapText[0]
    $file2 = $swapText[1]

    swap -file1 $file1 -file2 $file2

}