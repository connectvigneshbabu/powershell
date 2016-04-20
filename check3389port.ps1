#check telnet ports of various hosts
$hosts=get-content -Path "<Enter the path where the list of files >"
foreach ($h in $hosts)
{
    #Write-Host $h
    Write-Host "Trying to connect on $h on port 3389"
    try{
        if($sock=New-Object System.Net.Sockets.TcpClient($h,3389)){
            Write-Host "Connection to $h on 3389 succedded" -ForegroundColor Green
        }
    }
    catch [Exception]{
        #Write-Host $_.Exception.GetType().FullName -ForegroundColor red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
} 