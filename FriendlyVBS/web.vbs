Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oLogFile = oFSO.OpenTextFile("C:\IResult.txt", 2, True)
Set oShell = WScript.CreateObject("WScript.Shell")
strHost = "google.com"
strPingCommand = "ping -n 1 -w 300 " & strHost
ReturnCode = oShell.Run(strPingCommand, 0 , True)
If ReturnCode = 0 Then
	oLogFile.WriteLine "Successful ping."
Else
	oLogFile.WriteLine "Unsuccessful ping."
End If
oLogFile.Close