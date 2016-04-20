[CmdletBinding()]
Param(
	[Parameter(	ValueFromPipeline=$true,
				ValueFromPipelineByPropertyName=$true
				)]
	[string[]]
	$ComputerName = $env:ComputerName,
	[Parameter()]
	[string]
	$LocalGroupName = "Administrators",
	[Parameter()]
	[string]
	$OutputDir = "D:\Powershell\"
)
Begin {
	$OutputFile = Join-Path $OutputDir "LocalGroupMembers.csv"
	Write-Verbose "Script will write the output to $OutputFile folder"
	Add-Content -Path $OutPutFile -Value "ComputerName, LocalGroupName, Status, MemberType, MemberDomain, MemberName"
}
Process {
	ForEach($Computer in $ComputerName) {
		Write-host "Working on $Computer"
		If(!(Test-Connection -ComputerName $Computer -Count 1 -Quiet)) {
			Write-Verbose "$Computer is offline. Proceeding with next computer"
			Add-Content -Path $OutputFile -Value "$Computer,$LocalGroupName,Offline"
			Continue
		} else {
			Write-Verbose "Working on $computer"
			try {
				$group = [ADSI]"WinNT://$Computer/$LocalGroupName"
				$members = @($group.Invoke("Members"))
				Write-Verbose "Successfully queries the members of $computer"
				if(!$members) {
					Add-Content -Path $OutputFile -Value "$Computer,$LocalGroupName,NoMembersFound"
					Write-Verbose "No members found in the group"
					continue
				}
			}
			catch {
				Write-Verbose "Failed to query the members of $computer"
				Add-Content -Path $OutputFile -Value "$Computer,,FailedToQuery"
				Continue
			}
			foreach($member in $members) {
				try {
					$MemberName = $member.GetType().Invokemember("Name","GetProperty",$null,$member,$null)
					$MemberType = $member.GetType().Invokemember("Class","GetProperty",$null,$member,$null)
					$MemberPath = $member.GetType().Invokemember("ADSPath","GetProperty",$null,$member,$null)
					$MemberDomain = $null
					if($MemberPath -match "^Winnt\:\/\/(?<domainName>\S+)\/(?<CompName>\S+)\/") {
						if($MemberType -eq "User") {
							$MemberType = "LocalUser"
						} elseif($MemberType -eq "Group"){
							$MemberType = "LocalGroup"
						}
						$MemberDomain = $matches["CompName"]
					} elseif($MemberPath -match "^WinNT\:\/\/(?<domainname>\S+)/") {
						if($MemberType -eq "User") {
							$MemberType = "DomainUser"
						} elseif($MemberType -eq "Group"){
							$MemberType = "DomainGroup"
						}
						$MemberDomain = $matches["domainname"]
					} else {
						$MemberType = "Unknown"
						$MemberDomain = "Unknown"
					}
				Add-Content -Path $OutPutFile -Value "$Computer, $LocalGroupName, SUCCESS, $MemberType, $MemberDomain, $MemberName"
				} catch {
					Write-Verbose "failed to query details of a member. Details $_"
					Add-Content -Path $OutputFile -Value "$Computer,,FailedQueryMember"
				}
			}
		}
	}
}
End {} 