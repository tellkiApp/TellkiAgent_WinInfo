'###################################################################################################################################
'## This script was developed by Guberni and is part of Tellki's Monitoring Solution								              ##
'##																													              ##
'## December, 2014																									              ##
'##																													              ##
'## Version 1.0																										              ##
'##																													              ##
'## DESCRIPTION: Collect server information (operating system, partitions, RAM, services, ...)						              ##
'##																													              ##
'## SYNTAX: cscript "//Nologo" "//E:vbscript" "//T:90" "Server.vbs" <HOST> <USERNAME> <PASSWORD> <DOMAIN>             			  ##
'##																													              ##
'## EXAMPLE: cscript "//Nologo" "//E:vbscript" "//T:90" "Server.vbs" "10.10.10.1" "user" "pwd" "domain"	              			  ##
'##																													              ##
'## README:	<USERNAME>, <PASSWORD> and <DOMAIN> are only required if you want to monitor a remote server. If you want to use this ##
'##			script to monitor the local server where agent is installed, leave this parameters empty ("") but you still need to   ##
'##			pass them to the script.																						      ##
'## 																												              ##
'###################################################################################################################################

'Start Execution
Option Explicit
'Enable error handling
On Error Resume Next
If WScript.Arguments.Count <> 4 Then
	CALL ShowError(3, 0)
End If
'Set Culture - en-us
SetLocale(1033)

'INPUTS
Dim Host, Username, Password, Domain
Host = WScript.Arguments(0)
Username = WScript.Arguments(1)
Password = WScript.Arguments(2)
Domain = WScript.Arguments(3)


Dim infoData


Dim objSWbemLocator, objSWbemServices, colItems
Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")

Dim objItem, FullUserName, Counter, Version

	If Domain <> "" Then
		FullUserName = Domain & "\" & Username
	Else
		FullUserName = Username
	End If
	Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", FullUserName, Password)
	If Err.Number = -2147217308 Then
		Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", "", "")
		Err.Clear
	End If
	if Err.Number = -2147023174 Then
		CALL ShowError(4, Host)
		WScript.Quit (222)
	End If
	if Err.Number = -2147024891 Then
		CALL ShowError(2, Host)
	End If
	If Err Then CALL ShowError(1, Host)
	if Err.Number = 0 Then
		objSWbemServices.Security_.ImpersonationLevel = 3
		'Query 0
		Set colItems = objSWbemServices.ExecQuery( _
			"select BuildVersion from Win32_WMISetting")
		For Each objItem in colItems
			Version = CInt(objItem.BuildVersion)
		Next
		infoData = Version
		'Query 1
		Set colItems = objSWbemServices.ExecQuery( _
			"select * from Win32_ComputerSystem",,16) 
		For Each objItem in colItems 
			infoData = infoData & "||" & objItem.Name
			if Version >= 3000 Then
				infoData = infoData & "||" & objItem.DNSHostName
			else
				infoData = infoData & "||" & objItem.Name
			end if
			infoData = infoData & "||" & objItem.DomainRole
			infoData = infoData & "||" & objItem.Domain
			infoData = infoData & "||" & FormatNumber(objItem.TotalPhysicalMemory/1048576)
			infoData = infoData & "||" & objItem.Manufacturer
			infoData = infoData & "||" & objItem.Model
		Next
		'Query 2 
		Set colItems = objSWbemServices.ExecQuery( _
			"select * from Win32_OperatingSystem") 
		For Each objItem in colItems 
			infoData = infoData & "||" & objItem.Caption
			infoData = infoData & "||" & objItem.version
			infoData = infoData & "||" & objItem.CSDVersion
			if Version >= 3000 Then
				infoData = infoData & "||" & objItem.SystemDrive
			else
				infoData = infoData & "||" & "N/A"
			end if
			infoData = infoData & "||" & FormatNumber(objItem.TotalVisibleMemorySize/1024)
		Next
		'Query 3 - Array Partitions
		Set colItems = objSWbemServices.ExecQuery( _
			"select Name,DriveType,Description,FileSystem,Size from Win32_LogicalDisk") 
		infoData = infoData + "||"
		If IsEmpty(colItems) = False Then
			Counter = 0
			For Each objItem in colItems
				If Counter > 0 Then infoData = infoData & ","
				infoData = infoData & "{" & _
					objItem.Name & ";" & _
					objItem.DriveType & ";" & _
					objItem.Description & ";" & _
					objItem.FileSystem & ";" & _
					objItem.Size & "}"
				Counter = Counter + 1
			Next
		End If
		'Query 4 - CPU
		'http://en.wikipedia.org/wiki/Windows_NT
		Set colItems = objSWbemServices.ExecQuery( _
			"Select * from Win32_Processor") 
		infoData = infoData + "||"
		If IsEmpty(colItems) = False Then
			Counter = 0
			For Each objItem in colItems
				If Counter > 0 Then infoData = infoData & ","
				if Version >= 6000 Then
					infoData = infoData & "{" & _
						objItem.MaxClockSpeed & ";" & _
						objItem.NumberOfCores & "}"
					Counter = Counter + 1
				else
					Counter = Counter + 1
					infoData = infoData & "{" & _
						objItem.MaxClockSpeed & ";" & Counter & "}"
				end if
			Next
		End If
		'Query 5 - Windows Services
		Set colItems = objSWbemServices.ExecQuery( _
		"select Caption, Name from Win32_Service")
		infoData = infoData + "||"
		If IsEmpty(colItems) = False Then
			Counter = 0
			For Each objItem in colItems
				If Counter > 0 Then infoData = infoData & ","
				infoData = infoData & "{" & _
					objItem.Caption & ";" & _
					objItem.Name & "}"
				Counter = Counter + 1
			Next
		End If
		'Query 6 - Network Info
		Set colItems = objSWbemServices.ExecQuery("SELECT Caption, IPAddress, MACAddress FROM Win32_NetworkAdapterConfiguration where MACAddress is not null",,16)
		infoData = infoData + "||"
		Dim ip
		Dim singleIP
		Counter = 0
		For Each objItem in colItems
			If Counter > 0 Then infoData = infoData & ","
			CALL appendCollection(ip, objItem.IPAddress)
			infoData = infoData & "{" & _
						objItem.Caption & ";" & _
						objItem.MACAddress & ";" & _
						ip & "}"
			Counter = Counter + 1
			if ip <> "" Then
				singleIP = singleIP + ip
			end if
			ip = ""
		Next
		'Get IPv4 for the Server
		Dim arIP, IPs, i
		Counter = 0
		arIP = Split(singleIP,"^")
		for each i in arIP
			if len(i)<16 and len(i) >0 then
				If Counter > 0 Then IPs = IPs & ","
				IPs = IPs + i
				Counter = Counter + 1
			end if
		Next
		infoData = infoData + "||" + IPs
		CALL Output("2",infoData)
	End If
	Err.Clear


If Err Then 
	CALL ShowError(1, 0)
Else
	WScript.Quit(0)
End If

Sub appendCollection(msg, colctn)
Dim t
	 if IsArray(colctn) = true Then
		 for each t in colctn
			 msg = msg & t & "^"
		 next
	 End if 
End sub

Sub ShowError(ErrorCode, Param)
	Dim Msg
	Msg = "(" & Err.Number & ") " & Err.Description
	If ErrorCode=2 Then Msg = "Access is denied"
	If ErrorCode=3 Then Msg = "Wrong number of parameters on execution"
	If ErrorCode=4 Then Msg = "The specified target cannot be accessed"
	If ErrorCode=5 Then Msg = "There is no response in WMI or returned query is empty"
	WScript.Echo Msg
	WScript.Quit(ErrorCode)
End Sub

Sub Output(InfoType, Values)
	If Values <> "" Then
		WScript.Echo ToUTC() & "||" & InfoType & "||" & Values
	Else
		CALL ShowError(5, Host) 
	End If

End Sub

Function Debug(scriptFile, text)
'Open up the path to save the information into a text file
Dim Stuff, myFSO, WriteStuff, dateStamp

Set myFSO = CreateObject("Scripting.FileSystemObject")
Set WriteStuff = myFSO.OpenTextFile("C:\TELLAI\temp\" & scriptFile, 8, True)
WriteStuff.WriteLine(text)
WriteStuff.WriteLine("----------------------")
WriteStuff.Close
SET WriteStuff = NOTHING
SET myFSO = NOTHING
End Function

Function ToUTC()
	Dim dtmDateValue, dtmAdjusted
	Dim objShell, lngBiasKey, lngBias, k, UTC
	dtmDateValue = Now()
	'Obtain local Time Zone bias from machine registry.
	Set objShell = CreateObject("Wscript.Shell")
	lngBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")
	If (UCase(TypeName(lngBiasKey)) = "LONG") Then
		lngBias = lngBiasKey
		ElseIf (UCase(TypeName(lngBiasKey)) = "VARIANT()") Then
			lngBias = 0
		For k = 0 To UBound(lngBiasKey)
			lngBias = lngBias + (lngBiasKey(k) * 256^k)
		Next
	End If
	'Convert datetime value to UTC.
	UTC = DateAdd("n", lngBias, dtmDateValue)
	ToUTC =  FormatDateTime(UTC,2) & " " & FormatDateTime(UTC,3)
End Function
