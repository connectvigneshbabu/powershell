On Error Resume Next 
Const ForAppending = 8 
Const ForReading = 1
Dim objUser, dtmLastLogin, strLogonInfo, strGroup
Dim objNetwowrk
Dim objShell
 
'Declaring the variables 
Set objNetwowrk=CreateObject("wscript.Network")
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set SrvList = objFSO.OpenTextFile("ipaddress.txt", ForReading) 
Set ReportFile = objFSO.OpenTextFile ("Account_status.html", ForAppending, True) 
i = 0 
 
'Initializing the HTML Tags for better formatting 
ReportFile.writeline("<html>") 
ReportFile.writeline("<head>") 
ReportFile.writeline("<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>") 
ReportFile.writeline("<title>" & "Exchage  Servers Disk Space Report</title>") 
ReportFile.writeline("<style type='text/css'>") 
ReportFile.writeline("<!--") 
ReportFile.writeline("td {") 
ReportFile.writeline("font-family: Tahoma;") 
ReportFile.writeline("font-size: 11px;") 
ReportFile.writeline("border-top: 1px solid #999999;") 
ReportFile.writeline("border-right: 1px solid #999999;") 
ReportFile.writeline("border-bottom: 1px solid #999999;") 
ReportFile.writeline("border-left: 1px solid #999999;") 
ReportFile.writeline("padding-top: 0px;") 
ReportFile.writeline("padding-right: 0px;") 
ReportFile.writeline("padding-bottom: 0px;") 
ReportFile.writeline("padding-left: 0px;") 
ReportFile.writeline("}") 
ReportFile.writeline("body {") 
ReportFile.writeline("margin-left: 5px;") 
ReportFile.writeline("margin-top: 5px;") 
ReportFile.writeline("margin-right: 0px;") 
ReportFile.writeline("margin-bottom: 10px;") 
ReportFile.writeline("") 
ReportFile.writeline("table {") 
ReportFile.writeline("border: thin solid #000000;") 
ReportFile.writeline("}") 
ReportFile.writeline("-->") 
ReportFile.writeline("</style>") 
ReportFile.writeline("</head>") 
ReportFile.writeline("<body>") 
ReportFile.writeline("<table width='80%'>") 
ReportFile.writeline("<tr bgcolor='#CCCCCC'>") 
ReportFile.writeline("<td colspan='7' height='25' align='center'>") 
ReportFile.writeline("<font face='tahoma' color='#003399' size='2'><strong>Local Server Account Report</strong></font>")
ReportFile.writeline("<td colspan='6' height='25' align='Center'>")
ReportFile.writeline("<font face='tahoma' color='#003399' size='2'><strong>Date: </strong></font>")
ReportFile.writeline(Date) 
ReportFile.writeline("</td>") 
ReportFile.writeline("</tr>") 
ReportFile.writeline("</table>")
ReportFile.writeline("<td width='50%' colSpan=10 bgcolor='#0000FF'>&nbsp;</td>")

Set objNTInfo = CreateObject("WinNTSystemInfo")
GetComputerName = Ucase(objNTInfo.ComputerName)

 ReportFile.writeline("<table width='80%'>")
 ReportFile.writeline("<tr bgcolor='#CCCCCC'>")
 'ReportFile.WriteLine("<td width='05%' align='Left'>&nbsp;&nbsp COMPUTER NAME </font></td>")
 
  ' Code For Computer Host Name.
  
  'ComputerName=objNTInfo.ComputerName
'ReportFile.writeline("<td width='05%' align='Center'>" & objNTInfo.ComputerName &  "</font></td>")


'ReportFile.WriteLine("<p>" "<'img src="C:\Scripts\w3schoolslogoNEW310113.gif alt=Pulpit rock" width="304" height="228'>" "</p>")
'ReportFile.WriteLine("[<img src="C:\Scripts\w3schoolslogoNEW310113.gif" align=\"middle\">Local Accounts</img>]")

'ReportFile.WriteLine("<a> <img src=C:\Scripts\w3schoolslogoNEW310113.gif)></a>")

'objShape.AddPicture("C:\Scripts\w3schoolslogoNEW310113.gif") 

'Declaring the Server Name for report generation 
Do Until SrvList.AtEndOfStream 
    StrComputer = SrvList.Readline 
 
    ReportFile.writeline("<table width='80%'><tbody>") 
    ReportFile.writeline("<tr bgcolor='#CCCCCC'>") 
    ReportFile.writeline("<td width='50%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong>" & StrComputer & "</strong></font></td>")
    'ReportFile.WriteLine("<td width='05%' align='Left' colSpan=6><font face='tahoma' color='#003399' size='2'><strong>&nbsp;&nbsp COMPUTER NAME </font></td>")
    'ReportFile.writeline("<td width='50%' align='Center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong>" & objNTInfo.ComputerName & "</strong></font></td>")
    ReportFile.writeline("</tr>") 
 
 
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2") 
    Set colItems = objWMIService.ExecQuery("Select * from Win32_UserAccount",,48)
    Set colItems = objWMIService.ExecQuery("Select * from Win32_UserAccount Where Domain = '" & strComputer & "'")

 	
 	ReportFile.writeline("<table width='80%'>")
    ReportFile.writeline("<tr bgcolor=#CCCCCC>") 
        ReportFile.writeline("<td width='05%' height='20' align='center'>Description</td>") 
        ReportFile.writeline("<td width='05%' align='center'>Name</td>") 
        ReportFile.writeline("<td width='05%' align='center'>Lockout</td>") 
        ReportFile.writeline("<td width='05%' align='center'>Password Expires</td>") 
        ReportFile.writeline("<td width='05%' align='center'>Account Disabled</td>")
        ReportFile.writeline("<td width='05%' align='center'>Last Logon</td>")
        ReportFile.writeline("<td width='05%' align='center'>Password Changeable</td>")
        ReportFile.writeline("<td width='05%' align='center'>When Password Change</td>")
        ReportFile.writeline("<td width='05%' align='center'>Password Age</td>")
        'ReportFile.writeline("<td width='05%' align='center'>Member Of</td>")
    ReportFile.writeline("</tr>") 
 
        'Starting the loop to gather values from all Hard Drives 
        For Each objItem in colItems
        dtmLastLogin = ""
        dtmChangeDate = ""
        Set objUser = GetObject("WinNT://" & strComputer _
    	& "/" & objItem.Name & ",user")
    	'Set objUser=GetObject("WinNT://" & strComputer & "/"&objNetwowrk.Username &",user")
    	
    	'Set objUser = GetObject("WinNT://" & strComputer _
    	'& "/" & objItem.Name & ",Group")
    	
    	
    	intPasswordAge = objUser.PasswordAge
    	intPasswordAge = intPasswordAge * -1 
		dtmChangeDate = DateAdd("s", intPasswordAge, Now)
		dtmChangeDate = objUser.dtmChangeDate
		dtmLastLogin = objUser.lastLogin
		PassAge=INT(objUser.PasswordAge/86400)

		'strGroup=objItem.memberOf
 
            'Delcaring the Variables 
            
            AccountType=objItem.Description
            Name=objItem.Name
            PasswordChangeble=objItem.Lockout
            PasswordExpires=objItem.PasswordExpires
            Account Disabled=objItem.Disabled
            Last Logon=dtmLastLogin
            Password Changeable=objItem.PasswordChangeable         
            Password Change=dtmChangeDate
            Password Age=PassAge
            'Member Of=strGroup
            
    ReportFile.Writeline("<table width='80%'><tr><td width='05%' height='25' align=Left>" & objItem.Description & "</td>")
    ReportFile.Writeline("<td width='05%' align=center>" & objItem.Name & "</td>")
    ReportFile.Writeline("<td width='05%' align=center>" & objItem.Lockout & "</td>")
    ReportFile.Writeline("<td width='05%' align=center>" & objItem.PasswordExpires & "</td>")
    ReportFile.Writeline("<td width='05%' align=center>" & objItem.Disabled & "</td>")
    ReportFile.Writeline("<td width='05%' align=center>" & dtmLastLogin & "</td>")
    ReportFile.Writeline("<td width='05%' align=center>" & objItem.PasswordChangeable & "</td>")
    ReportFile.Writeline("<td width='05%' align=center>" & dtmChangeDate & "</td>")
    If PassAge > 90 Then
    	ReportFile.Writeline("<td width='05%' align=center bgcolor='#FFFF00' align=center>" & PassAge & "</td>")
    ElseIf PassAge < 90 Then
 		ReportFile.Writeline("<td width='05%' align=center bgcolor='#00FF00' align=center>" & PassAge & "</td>")   
    	
    'ReportFile.Writeline("<td width='05%' align=center>" & strGroup & "</td>")
    End If          
        Next 
 
    ReportFile.writeline("<tr>") 
    ReportFile.writeline("<td width='50%' colSpan=10 bgcolor='#806D7E'>&nbsp;</td>") 
    ReportFile.writeline("</tr>") 
 
    ReportFile.writeline("</tbody></table>") 
Loop 
ReportFile.WriteLine("<center> script executed by vigneshbabu(875328).  All Rights Reserved </center>")
ReportFile.WriteLine "</body></html>" 
