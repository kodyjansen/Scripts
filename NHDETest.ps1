#Connect to SQL Server
$con = new-object System.data.sqlclient.SQLconnection
$con.ConnectionString ="Data Source=WW-SEAT-SVR;Initial Catalog=WW-SEAT-DB;User ID=Seat;Password=Chartblock38;Connect Timeout=30"
$con.open()
$sqlcmd = new-object System.data.sqlclient.sqlcommand
$sqlcmd.connection = $con
$sqlcmd.CommandTimeout = 600000
$sqlcmd.CommandText = "SELECT * FROM dbo.EmployeeApprovalTest WHERE [Status] = 'P'"
$Reader = $sqlcmd.ExecuteReader()

#Read Cell Data
    while ($reader.read())
    {
        For($i = 0;$i -lt 10;$i++)
        {
            $null = $reader.item($i)
            
            if($i -eq 0)
            {
                $firstName = $reader.item($i)
            }

            if($i -eq 1)
            {
                $lastName = $reader.item($i)
                      
            }
            if($i -eq 2)
            {
                $title = $reader.item($i)
                
            }
            if($i -eq 4)
            {
                $manager = $reader.item($i)
                
            }
            if($i -eq 8)
            {
                $startDate = $reader.item($i)
                $startDate = $startDate.toshortdatestring()            
            }       
        }

#A/AN Logic
$titlechar = $title[0]
switch($titlechar)
{
    A{$an = "an"}
    E{$an = "an"}
    I{$an = "an"}
    O{$an = "an"}
    U{$an = "an"}
    Default{$an = "a"}
}

#Construct Manager email
$indexof = $manager.lastindexof(' ') + 1
$managerLastName = $manager.substring($indexof)
$manageremail = $manager.substring(0,1) + $managerLastName 
$manageremail = $manageremail.toLower() + "@worldwide-logistics.com"

#Construct Outgoing Email
function Send-Email($sendBody, $sendSubject, $sendTo)
    {
    #Email Info
    $From         = "scriptad@wwidelogistics.onmicrosoft.com"
    $SMTPServer   = "smtp.office365.com"
    $Password     = Convertto-SecureString "Temperedcontrols31" -AsPlainText -Force
    $Creds        = New-Object System.Management.Automation.PSCredential -ArgumentList $From, $Password

    $BodyMessage    = @"
    You have a new employee, $firstName $lastName, starting as $an $title on $startDate. Please respond to this email detailing where they will be sitting by the Thursday before start date.

    Thank you!
    IT   
"@
    $To           = "kjansen@worldwide-logistics.com"
    $Subject = "New Employee: $firstName $lastName" 

#######SEND EMAIL#######
    send-mailmessage -from $From -to $To -Subject $Subject -SmtpServer $SMTPServer -Body $BodyMessage -UseSsl -Credential $Creds

    }
    
 Send-Email -sendBody $BodyMessage -sendSubject $Subject -sendTo $To
   
    }




