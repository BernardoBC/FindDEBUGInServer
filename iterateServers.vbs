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
result=Msgbox("Script will change all instances of '=DEBUG' to '=INFO' for all log files of the servers in ServerList.txt. These changes are not reversible."&vbCrLf&vbCrLf&"Do you want to Continue? ",vbYesNo, "Warning")
If result = 7 Then
    Wscript.Quit
End If

StartTime = Timer()

'I/O files
strReadFile = ".\ServerList.txt"
strWriteFile = ".\DebugInfo.csv"

'Credentials
strUser = "infra\bernardo.bonilla"
strPassword = "" 'insert credentials

'Set up Input file
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objServerList = objFS.OpenTextFile(strReadFile, For_READING)

'Set up Output File'
Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objOutFile = objFSO.CreateTextFile(strWriteFile,True)
objOutFile.WriteLine "IP" & "," & "Connection Status" & "," & "Files Found" & "," & "Instances Found" & "," & "Time Elapsed (s)"

'Main Loop - Iterates each server on the list
Do Until objServerList.AtEndOfStream
	
	'Set ups directory to search based on IP
	strComputer = objServerList.ReadLine
	remotePath = "\\" & strComputer & "\C$\HPBSM\conf\core\Tools\log4j"

	'Pings Server
	If fPingTest(strComputer) Then

		'Map Network Drive
		Wscript.Echo now & ": Connecting to " & strComputer
		drive = "s:"
        net.MapNetworkDrive drive, remotePath, False, strUser, strPassword 
        strResult = strComputer& "," & "connected"
        WScript.Echo now & ": Connection Succesful"

        'Executes Script on S: Drive
        WScript.Echo now & ": Executing Search"
		objShell.Run "cscript.exe findDebugInServer.vbs", 2, true 
		Set objShell = Nothing

		'Removes S: network drive
        net.RemoveNetworkDrive drive, True
        WScript.Echo now & ": Connection Closed"

        Set outFso  = CreateObject("Scripting.FileSystemObject")
		Set outFile = outFso.OpenTextFile(".\output.txt", 1)
		
		'Reads Output - Logs relevant info (see findDebugInServer for more info)
		If outFso.GetFile(".\output.txt").size <> 0 Then
			strOutputFile = outFile.ReadAll
			If VarType(strOutputFile) = 8 Then
				If strOutputFile <> "" Then 

					'String is Not Null And Not Empty
					strResult = strResult & "," & strOutputFile
					
				End If
			End If
			Wscript.Echo now & ": " & strResult
			objOutFile.WriteLine strResult
		Else
		strResult = strResult & "," & "Output File Empty"	

		End If

		outFile.Close

	Else
		WScript.Echo now & "Server " & strComputer & " is unreachable"
		objOutFile.WriteLine strComputer & ","& "unreachable"
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