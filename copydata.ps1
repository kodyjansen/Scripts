#Path Variables
    $sourcepath = "\\ww-file16-01\ProfileRedirect\kjansen\Documents\Test Docs\Copydata\" 
    $destinationpath = "\\ww-file16-01\ProfileRedirect\kjansen\Documents\Test Docs\Destination"
    $contentdestinationpath = "\\ww-file16-01\ProfileRedirect\kjansen\Documents\Test Docs\Destination\Day 1.txt"
#Gets contents of each file in source folder
    $subcontents = Get-Childitem $sourcepath
    $contents = $subcontents | Get-Content
#Gets contents of Destination Folder
    $newfilename = Get-Childitem $destinationpath

#Checks if file exists, if not, creates it and copies data
    If(-not(Test-path -path $contentdestinationpath))
    { 
        $newfile = New-Item -path $destinationpath -name "Day 1.txt"
        Write-host "File" $newfile.name "has been created"
        $contents | out-file -filepath $newfile     
    }
#If file already exists, copies data to file
    Else
    {
        Write-Host "File" $newfilename "already exists, File has been Updated"
        $contents | out-file -filepath $contentdestinationpath
    }
    

   

           
        

       





#Get-childitem $sourcepath | measure-object   