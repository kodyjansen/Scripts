
$User = "worldwide\adminkody"
$password = 01000000d08c9ddf0115d1118c7a00c04fc297eb0100000024341d9c92f10b4aa65c4585f53c8f320000000002000000000003660000c00000001000000054c7704f9113be62bba49d8144c07c900000000004800000a000000010000000ed5e840988c775ee8b2237c97b908c1d20000000a5bbd9530a6be1f7a00f8750857df3c70e199d50217b211d180c95f5f0566db114000000cbef8145e9e9f7073b4303ef9ecce8a53d6ebe79 | ConvertTo-SecureString
$File = "C:\Users\kjansen\Scripts\standalone.ps1"
$Credential=Get-credential 

Start-Process "C:\Users\kjansen\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Postman\Postman.lnk" -runas "Adminkody"