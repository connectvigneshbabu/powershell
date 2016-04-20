


$PATH = "D:\Powershell\test\" 

$InputFile = "$PATH\ipaddress.txt" 

$OUTFile = "$PATH\out.csv" 
$UserName = Read-Host "Enter User Name "
$Password = Read-Host -AsSecureString "Enter Your Password "

$ComputerName = Get-Content -Path $InputFile 

$Obj=@()

Foreach($Computer in $ComputerName)
{
Write-host "Working on $Computer"
	
		$AllLocalAccounts = Get-WmiObject -Class Win32_UserAccount -Namespace "root\cimv2" `
		-Filter "LocalAccount='$True'" -ComputerName $Computer -ErrorAction Stop
	
	
	Foreach($LocalAccount in $AllLocalAccounts)
	{
		$Object = New-Object -TypeName PSObject
		$Object|Add-Member -MemberType NoteProperty -Name "IP Address" -Value $Computer
		$Object|Add-Member -MemberType NoteProperty -Name "Name" -Value $LocalAccount.Name
		$Object|Add-Member -MemberType NoteProperty -Name "Full Name" -Value $LocalAccount.FullName
		$Object|Add-Member -MemberType NoteProperty -Name "Caption" -Value $LocalAccount.Caption
      	$Object|Add-Member -MemberType NoteProperty -Name "Disabled" -Value $LocalAccount.Disabled
      	$Object|Add-Member -MemberType NoteProperty -Name "Status" -Value $LocalAccount.Status
      	$Object|Add-Member -MemberType NoteProperty -Name "LockOut" -Value $LocalAccount.LockOut
		$Object|Add-Member -MemberType NoteProperty -Name "Password Changeable" -Value $LocalAccount.PasswordChangeable
		$Object|Add-Member -MemberType NoteProperty -Name "Password Expires" -Value $LocalAccount.PasswordExpires
		$Object|Add-Member -MemberType NoteProperty -Name "Password Required" -Value $LocalAccount.PasswordRequired
		$Object|Add-Member -MemberType NoteProperty -Name "SID" -Value $LocalAccount.SID
		$Object|Add-Member -MemberType NoteProperty -Name "SID Type" -Value $LocalAccount.SIDType
		$Object|Add-Member -MemberType NoteProperty -Name "Account Type" -Value $LocalAccount.AccountType
		$Object|Add-Member -MemberType NoteProperty -Name "Domain" -Value $LocalAccount.Domain
		$Object|Add-Member -MemberType NoteProperty -Name "Description" -Value $LocalAccount.Description
		
		$Obj+=$Object
	}
	
	If($AccountName)
	{
		Foreach($Account in $AccountName)
		{
			$Obj|Where-Object{$_.Name -like "$Account"}
		}
	}
	else
	{
		$Obj
	}
}