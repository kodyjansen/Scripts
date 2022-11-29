function createFile{
    Param(
        $filePath, 
        $fileName 
    )
   
    $fileText = "Test Text"
    $fullFilePath = $filePath + $fileName
    $fileText | Out-File -FilePath $fullFilePath
    Write-Host $fullFilePath
  
    
}
