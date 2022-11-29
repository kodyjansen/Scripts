Function createFiles
{
	Param ($numOfFiles)


	New-Item -Path "\\ww-file16-01\ProfileRedirect\kjansen\desktop\CreateFiles" -ItemType "directory" #Folder Created
        
            If((Test-Path "\\ww-file16-01\ProfileRedirect\kjansen\desktop\CreateFiles\") -eq $true)
            {}
           
            For($num = 1; $num -le $numOfFiles; $num++)
            {  

               
                If((Test-Path "\\ww-file16-01\ProfileRedirect\kjansen\desktop\CreateFiles\file" + $num + ".txt") -eq $false)
                {
                   New-Item -Path ("\\ww-file16-01\ProfileRedirect\kjansen\desktop\CreateFiles\file" + $num + ".txt")
                }
            }
}
      
