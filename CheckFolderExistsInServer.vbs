On Error Resume Next
Const For_READING = 1
Const FOR_WRITING = 2
dim counter
dim minicounter
Dim objShell
Set objShell = Wscript.CreateObject("WScript.Shell")
Set net = CreateObject("WScript.Network")
Set FSO = CreateObject("Scripting.FileSystemObject")
counter = 0

'Warning Window
' result=Msgbox("Script will change all instances of '=DEBUG' to '=INFO' for all log files of the servers in ServerList.txt. These changes are not reversible."&vbCrLf&vbCrLf&"Do you want to Continue? ",vbYesNo, "Warning")
' If result = 7 Then
    ' Wscript.Quit
' End If

' StartTime = Timer()

'I/O files
strReadFile = ".\servers.txt"
strWriteFile = ".\DebugInfo.csv"

'Credentials
strUser = "infra\joseiby.hernandez"
strPassword = "Portami.0" 'insert infra credentials

'Set up Input file
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objServerList = objFS.OpenTextFile(strReadFile, For_READING)

'Set up Output File'
Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objOutFile = objFSO.CreateTextFile(strWriteFile,True)
' objOutFile.WriteLine "HostName" & "," & "IP" & "," & "Connection Status" & "," & "Files Found" & "," & "Instances Found" & ","  & "Files" & "," & "Time Elapsed (s)"

'Main Loop - Iterates each server on the list
Do Until objServerList.AtEndOfStream
	
	'Set ups directory to search based on IP
	parsedIP = Split(objServerList.ReadLine," ")
	strComputer = parsedIP(0)
	
	remotePath = "\\" & strComputer & "\c$\ProgramData"

	'Pings Server
	If fPingTest(strComputer) Then

		'Map Network Drive
		Wscript.Echo now & ": Connecting to " & strComputer
		drive = "S:"
        net.MapNetworkDrive drive, remotePath, False, strUser, strPassword 
        'strResult = strHostName & ","& strComputer & "," & "connected"
        WScript.Echo now & ": Connection Succesful"
		
		WScript.Echo "Press [ENTER] to continue..."
		' Read dummy input. This call will not return until [ENTER] is pressed.
		WScript.StdIn.ReadLine
		
		
        'Executes Script on S: Drive
        path="S:\HP\BPM\workspace\sandbox" 'path to folder    
		exists = fso.FolderExists(path)

		If (exists) then 
			Wscript.Echo strComputer & " OK"
		Else
			Wscript.Echo strComputer & " Not found"
		end if

		'Removes S: network drive
        net.RemoveNetworkDrive drive, True
        WScript.Echo now & ": Connection Closed"

        Set outFso  = CreateObject("Scripting.FileSystemObject")
		Set outFile = outFso.OpenTextFile(".\output.txt", 1)
		
		' 'Reads Output - Logs relevant info (see findDebugInServer for more info)
		' If outFso.GetFile(".\output.txt").size <> 0 Then
			' strOutputFile = outFile.ReadAll
			' If VarType(strOutputFile) = 8 Then
				' If strOutputFile <> "" Then 

					' 'String is Not Null And Not Empty
					' strResult = strResult & "," & strOutputFile
					
				' End If
			' End If
			' Wscript.Echo now & ": " & strResult
			' objOutFile.WriteLine strResult
		' Else
		' strResult = strResult & "," & "Output File Empty"	

		' End If

		outFile.Close

	Else
		WScript.Echo now & "Server " & strComputer & " is unreachable"
		objOutFile.WriteLine strHostName & ","& strComputer & ","& "unreachable"
	End If

Loop	

'Deletes temporary output file
objOutFile.close
objFSO.DeleteFile("output.txt")

'Script Summary
EndTime = Timer()
Wscript.echo
Wscript.echo "Elapsed Time: " &FormatNumber(EndTime - StartTime, 2) & "seconds"


'Ping to server to avoid WMI timeout for unreachable or misspelled servers
Function fPingTest(strComputer) 
	Set objshell = CreateObject("WScript.shell")
	Set objPing = objShell.Exec ("ping " & strComputer & " -n 2 -w 20")
	strPingOut = objPing.StdOut.ReadAll
	if instr(Lcase(strPingOut), "reply") then
		fPingTest = TRUE
	Else
		fPingTest = FALSE
	End If
End Function

Function HostName(IPAddress)
 Dim objWMI, objItem, colItems

 On Error Resume Next
 ' Get local WMI CIMv2 object
 Set objWMIService = GetObject("winmgmts:\\" & IPAddress & "\root\cimv2")
 If Err.Number <> 0 Then
  HostName = "Error"
  Err.Clear
   On Error Goto 0
  Exit Function
 End If
 On Error Goto 0

 Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

 For Each objItem In colItems
  HostName = objItem.Name
 Next

 Set colItems = Nothing
 Set objWMI = Nothing
End Function


Function ResolveIP(ComputerName)
  Dim objShell, objExec, StrOutput, RegEx 
  Set objShell = Sys.OleObject("WScript.Shell")
  Set objExec = objShell.Exec("ping " & ComputerName & " -n 1")
  StrOutput = objExec.StdOut.ReadAll
  Set RegEx = HISUtils.RegExpr
  RegEx.Expression = "(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})"
  RegEx.Exec(StrOutput)
  If RegEx.MatchPos(0) Then
    ResolveIP = RegEx.Substitute("$&")
  Else
    ResolveIP = ""
  End If
End Function
