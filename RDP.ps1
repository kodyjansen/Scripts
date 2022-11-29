#Assemblys
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

#Creates a new RDP file from a template based off the $hostname and 
#$username passed.
function createRPD(){
    Param(
        $hostname,
        $username
    )

    #Gets the RDP template file
    $rdpTemplate = Get-Content -Path "\\ww-file16-01\IT\RDPLinkCreation\ww-rdpTemplate.rdp"

    #Sets the full address(hostname) property
    $rdpTemplate = $rdpTemplate.Replace("WW--PC", $hostname)

    #Sets the username property
    $rdpTemplate = $rdpTemplate.Replace("worldwide\", ("worldwide\" + $username))

    #Saves rdp file
    $rdpFilePath = ("\\ww-file16-01\IT\RDP_Temp_Location\" + $username + ".RDP")
    $rdpTemplate | Out-File -FilePath $rdpFilePath

    #Creates email service account creds
    $serviceEmail = "itnotifications@worldwide-logistics.com"
    $servicePass = ConvertTo-SecureString "compcyHawk228" -AsPlainText -Force
    $serviceEmailCreds = New-Object System.Management.Automation.PSCredential ($serviceEmail, $servicePass)

    #Emails the target user the RDP file
    $toEmailAddress = $username + "@worldwide-logistics.com"
    $subject = "RDP Link"
    $body = "Windows Users: Download the attached RDP link to your desktop. Then double click to run. `n`nMac Users: Download 'Microsoft Remote Desktop' then download the RDP link to your desktop and double click to run. `n`n'Microsoft Remote Desktop' Download Link: https://apps.apple.com/us/app/microsoft-remote-desktop-10/id1295203466?mt=12"
    Send-MailMessage -Attachments $rdpFilePath -To $toEmailAddress -From $serviceEmail -Subject $subject -Body $body -Credential $serviceEmailCreds -SmtpServer "smtp-mail.outlook.com" -Port 587 -UseSSL
}

#GUI
Function Set-OGVWindow {
    [OutputType('System.Automation.WindowInfo')]
    [cmdletbinding()]
    Param (
        [parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$True)]
        $ProcessName,
        [parameter(Mandatory=$true)]
        [string]$WindowTitle,
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height
    )
    Begin {
        Try {
            [void][Window]
        } Catch {
        Add-Type @"
              using System;
              using System.Runtime.InteropServices;
              public class Window {
                [DllImport("user32.dll")]
                [return: MarshalAs(UnmanagedType.Bool)]
                public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
                [DllImport("user32.dll")]
                public extern static bool MoveWindow(IntPtr handle, int x, int y, int width, int height, bool redraw);
              }
              public struct RECT
              {
                public int Left;        // x position of upper-left corner
                public int Top;         // y position of upper-left corner
                public int Right;       // x position of lower-right corner
                public int Bottom;      // y position of lower-right corner
              }
"@
        }
    }
    Process {
        $Rectangle = New-Object RECT
        $Handle = (Get-Process -Name $ProcessName | Where-Object MainWindowTitle -eq $WindowTitle).MainWindowHandle
        if($Handle) {
            $Return = [Window]::GetWindowRect($Handle,[ref]$Rectangle)
            If ($Return) {
                $Return = [Window]::MoveWindow($Handle, $x, $y, $Width, $Height,$True)
            }
        }
    }
}
function Get-Variables {
    ##################
    #Script Variables#
    ##################
    $X = [System.Windows.Forms.Cursor]::Position.X
    $Y = [System.Windows.Forms.Cursor]::Position.Y
    $computerNames  = (get-adcomputer -filter * -searchbase 'OU=WW_IA_DesktopsOU, OU=WW_IA_OU, DC=worldwide, DC=local').Name
    $computerNames  = $computerNames | Sort-Object
    $job = Start-Job -scriptblock {
        Param($func,$X,$Y)
        $function = [scriptblock]::Create($func)
        For ($i = 0; $i -le 10; $i++) {
            & $function -ProcessName powershell -WindowTitle "RDP: Select a User" -Width 250 -Height 500 -x $X -y $Y
        }
    } -ArgumentList ${function:Set-OGVWindow},$X,$Y
    $DesktoptoRun = $computerNames | Out-GridView -Title "RDP: Select a PC" -PassThru
    $userNames = (Get-ADUser -Filter * -SearchBase "OU=WW_IA_usersOU,OU=WW_IA_OU,DC=WorldWide,DC=local" | Where-Object {($_.DistinguishedName -notlike "*OU=MD_UsersOU*" -and $_.DistinguishedName -notlike "*OU=TMC_UsersOU*" -and $_.DistinguishedName -notlike "*OU=WW_DeletedUsersOU*" -and $_.DistinguishedName -notlike "*OU=WW_SharedMailboxUsersOU*")}).Name
    $userNames = $userNames | Sort-Object
    $UserName = $userNames | Out-GridView -Title "RDP: Select a User" -PassThru
    $UsertoRun = (Get-ADUser -Filter {displayName -like $UserName}).SamAccountName
    Remove-Job $job -Force -Confirm:$false
    return ($DesktoptoRun, $UsertoRun)
}

#Run
$hostname, $username = Get-Variables
createRPD -username $username -hostname $hostname 