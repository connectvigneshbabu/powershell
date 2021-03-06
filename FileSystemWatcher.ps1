<#   
Name: FileSystemwatcher.ps1
Author: Vigneshbabu 
DateCreated: 14-04-2016
Project : Corporate Data center
Version: 1.1

To stop the monitoring, run the following commands:
 Unregister-Event FileDeleted
 Unregister-Event FileCreated
 Unregister-Event FileChanged    
#>

<#
function module Get-Name for getting the hostname and date of the file
#>
Function Get-Name {[system.environment]::MachineName}

$hostname = Get-Name
$datetime = Get-Date -Format “ddMMyyyy”

$outfile= "$hostname($datetime)"

$folder = 'C:\scripts' # Enter the root path you want to monitor.
$filter = '*.*'  # You can enter a wildcard filter here.

# In the following line, you can change 'IncludeSubdirectories to $true if required.                          
$fsw = New-Object IO.FileSystemWatcher $folder, $filter -Property @{IncludeSubdirectories = $true;NotifyFilter = [IO.NotifyFilters]'FileName, LastWrite'}


Register-ObjectEvent $fsw Created -SourceIdentifier FileCreated -Action {
$name = $Event.SourceEventArgs.Name
$changeType = $Event.SourceEventArgs.ChangeType
$timeStamp = $Event.TimeGenerated
Write-Host "The file '$name' was $changeType at $timeStamp" -fore green
Out-File -FilePath C:\out\$outfile.csv -Append -InputObject "In the server $hostname file '$name' was $changeType at $timeStamp"}

Register-ObjectEvent $fsw Deleted -SourceIdentifier FileDeleted -Action {
$name = $Event.SourceEventArgs.Name
$changeType = $Event.SourceEventArgs.ChangeType
$timeStamp = $Event.TimeGenerated
Write-Host "The file '$name' was $changeType at $timeStamp" -fore red
Out-File -FilePath C:\out\$outfile.csv -Append -InputObject "In the server $hostname file '$name' was $changeType at $timeStamp"}

Register-ObjectEvent $fsw Changed -SourceIdentifier FileChanged -Action {
$name = $Event.SourceEventArgs.Name
$changeType = $Event.SourceEventArgs.ChangeType
$timeStamp = $Event.TimeGenerated
Write-Host "The file '$name' was $changeType at $timeStamp" -fore white
Out-File -FilePath C:\out\$outfile.csv -Append -InputObject "In the server $hostname file '$name' was $changeType at $timeStamp"}
