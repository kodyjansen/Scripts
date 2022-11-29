#Path Variables
$sourcefolderpath = "\\Profile\Users\kjansen\Documents\Source Folder"
$destinationfolderpath = "\\Profile\Users\kjansen\Documents\Event Log"
$sourcefolderpathitems = $sourcefolderpath | Get-Childitem
$destionationfolderpathitems = $destinationfolderpath | Get-Childitem

#Create source values for original text
Function CreateSourceFiles
{
    For($i = 1; $i -lt 8; $i++)
    {
        $newfiles = New-Item -path ($sourcefolderpath + "Source File " + $i + ".txt") 
    }
}

#Create destination values for text to be copied
Function CreateDestinationFiles
{
    Foreach($file in $sourcefolderpathitems)
    {  
        $newentrylog = New-Item -path $destinationfolderpath -name $file
    }
}   

#Check if destination values already exist, if false, create
Try{
    If(-not(Test-Path -path $destionationfolderpathitems))
        {CreateDestinationFiles}
   }
Catch 
{
    CreateDestinationFiles
}

#Copy text from source values to destination values
Function CopyValue
{
for($x = 0; $x -lt 8; $x++)
{
    $sourcetext = Get-Childitem $sourcefolderpath | Select-object -index($x) | Get-Content
    $newfilepath = "\\Profile\Users\kjansen\Documents\Event Log\Source File " + $x + ".txt"
    $sourcetext | Out-file -filepath $newfilepath


$displayed = Get-Childitem $newfilepath | Get-content
$displayed

}
}

Function Register-Watcher{
    param ($folder)
    $folder = $sourcefolderpath
    $filter = "*.*" #all files
    $watcher = New-Object IO.FileSystemWatcher $folder, $filter -Property @{ 
        IncludeSubdirectories = $false
        EnableRaisingEvents = $true
    }
}

     $changeAction = [CopyValue]::Create(
     $path = $Event.SourceEventArgs.FullPath
        $name = $Event.SourceEventArgs.Name
        $changeType = $Event.SourceEventArgs.ChangeType
        $timeStamp = $Event.TimeGenerated
        Write-Host("The file " + $name + "was" + $changeType + "at" + $timeStamp)
    )

 Register-ObjectEvent $watcher -EventName "Changed" -Action $changeAction


 Register-Watcher $folder

