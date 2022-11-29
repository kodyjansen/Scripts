#if()


# Add-Type -AssemblyName presentationCore
# $mediaPlayer = New-Object system.windows.media.mediaplayer
# $mediaPlayer.open('C:\mario death.mp3')
# $mediaPlayer.Play()


$Outlook = New-Object -comobject Outlook.Application
$Namespace = $Outlook.GetNameSpace("MAPI")
$EmailFolderName = "HR Alerts"
$EmailFolder = $Namespace.Folders.Items(0).Folders.Item($EmailFolderName)

