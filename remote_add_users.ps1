$group = Read-Host "Enter the group you want a user to add in"
$user = Read-Host "enter domain user id"
$pc = Read-Host "enter pc number"
$objUser = [ADSI]("WinNT://DOMAIN/$user")
$objGroup = [ADSI]("WinNT://$pc/$group")
$objGroup.PSBase.Invoke("Add",$objUser.PSBase.Path)

