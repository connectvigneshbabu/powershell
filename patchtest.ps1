$computer = Get-Content d:\powershell\ipaddress.txt
get-wmiobject -ComputerName $computer -Namespace root\ccm\ClientSDK -Class CCM_SoftwareUpdate | `
Select-Object PSComputername,ArticleID,BulletinID,Name,EvaluationState | Export-Csv -NoType d:\powershell\output.csv