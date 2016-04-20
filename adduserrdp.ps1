$user = Read-Host "enter domain user id" 
$pc = Read-Host "enter pc number" 
$objUser = [ADSI]("WinNT://DOMAIN/$user") 
$objGroup = [ADSI]("WinNT://$pc/Remote Desktop Users") 
$objGroup.PSBase.Invoke("Add",$objUser.PSBase.Path)