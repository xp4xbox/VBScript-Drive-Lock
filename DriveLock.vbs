Option Explicit
Dim objFso, objWshShl
Dim colDrives, objDrive, strDriveList, strChoice, strDrives
Dim intDriveNumber, intSplash
Set objFso = CreateObject("Scripting.FileSystemObject")
Set objWshShl = CreateObject("WScript.Shell")
Const cTitleBarMsg = "Drive Locker"
RunAsAdmin()
Function DisplayPrompt()
	intSplash = MsgBox("What would you like to do?" & vbCrLf & vbCrLf  _
 	& "[Click on YES to lock a drive] " & vbCrLf _
  	 & "[Click on NO to unlock drive(s)]",35, cTitleBarMsg)
	If intSplash = 2 Then
		DisplaySplashScreen()
	ElseIf intSplash = 7 Then
		On Error Resume Next
		objWshShl.RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoViewOnDrive"
		If Err.Number <> 0 Then
			MsgBox "Drives are already unlocked.",16,cTitleBarMsg
			DisplayPrompt()
		End If
		On Error Goto 0
		objWshShl.Run "Taskkill /f /im explorer.exe",0
		WScript.Sleep 300
		objWshShl.Run "cmd /c explorer.exe",0
		MsgBox "Drive unlocked was succesfull!",64,cTitleBarMsg
		DisplayPrompt()
	End If
End Function
DisplayPrompt()
Set colDrives = objFSO.Drives
For Each objDrive in colDrives
    strDriveList = strDriveList & objDrive.DriveLetter & Space(10)
Next
strDrives = LCase(Replace(strDriveList," ","",1,-1))
Set colDrives = objFSO.Drives
strDriveList = ""
For Each objDrive in colDrives
    strDriveList = strDriveList & objDrive.DriveLetter & ":\" & Space(5)
Next
InputMenu()
Sub InputMenu
	strChoice = InputBox("Enter letter of the drive you wish to lock." & _
	 " Or type ALL to lock all drives." & _
 	  vbcrlf & vbcrlf & "Available drives" & Space(3) & _
  	   ":" & vbCrLf & vbCrLf & strDriveList, cTitleBarMsg)
	If IsEmpty(strChoice) Then
		DisplaySplashScreen()
	ElseIf strChoice = "" Then
		MsgBox "Do not leave this blank.",16, cTitleBarMsg
		InputMenu()
	ElseIf LCase(strChoice) = "all" Then
		'Do Nothing
	ElseIf Len(strChoice) <> 1 Then 
		MsgBox "You must enter the letter ONLY.",16, cTitleBarMsg
		InputMenu()
	ElseIf Not InStr(1,strDrives,LCase(strChoice),1) <> 0 Then 
		MsgBox "Invalid choice, please try again.",16, cTitleBarMsg
		InputMenu()
	End If
End Sub
Sets()
WriteToRegistry()
objWshShl.Run "Taskkill /f /im explorer.exe",0
WScript.Sleep 300
objWshShl.Run "cmd /c explorer.exe",0
MsgBox "Drive locked was succesfull!",64,cTitleBarMsg
objWshShl.Run """"&WScript.ScriptFullName&""""
WScript.Quit()
Sub Sets
	strChoice = LCase(strChoice)
	If strChoice = "all" Then
		intDriveNumber = 67108863
	ElseIf strChoice = "a" Then
		intDriveNumber = 1
	ElseIf strChoice = "b" Then
		intDriveNumber = 2
	ElseIf strChoice = "c" Then
		intDriveNumber = 4
	ElseIf strChoice = "d" Then
		intDriveNumber = 8
	ElseIf strChoice = "e" Then
		intDriveNumber = 16
	ElseIf strChoice = "f" Then
		intDriveNumber = 32
	ElseIf strChoice = "g" Then
		intDriveNumber = 64
	ElseIf strChoice = "h" Then
		intDriveNumber = 128
	ElseIf strChoice = "i" Then
		intDriveNumber = 256
	ElseIf strChoice = "j" Then
		intDriveNumber = 512
	ElseIf strChoice = "k" Then
		intDriveNumber = 1024
	ElseIf strChoice = "l" Then
		intDriveNumber = 2048
	ElseIf strChoice = "m" Then
		intDriveNumber = 4096
	ElseIf strChoice = "n" Then
		intDriveNumber = 8192
	ElseIf strChoice = "o" Then
		intDriveNumber = 16384
	ElseIf strChoice = "p" Then
		intDriveNumber = 32768
	ElseIf strChoice = "q" Then
		intDriveNumber = 65536
	ElseIf strChoice = "r" Then
		intDriveNumber = 131072
	ElseIf strChoice = "s" Then
		intDriveNumber = 262144
	ElseIf strChoice = "t" Then
		intDriveNumber = 524288
	ElseIf strChoice = "u" Then
		intDriveNumber = 1048576
	ElseIf strChoice = "v" Then
		intDriveNumber = 2097152
	ElseIf strChoice = "w" Then
		intDriveNumber = 4194304
	ElseIf strChoice = "x" Then
		intDriveNumber = 8388608
	ElseIf strChoice = "y" Then
		intDriveNumber = 16777216
	ElseIf strChoice = "z" Then
		intDriveNumber = 33554432
	End If
End Sub
Function WriteToRegistry()
	objWshShl.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoViewOnDrive", _
	 intDriveNumber, "REG_DWORD"
End Function
Function RunAsAdmin()
	If WScript.Arguments.length = 0 Then
		CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & _
		WScript.ScriptFullName & """" & " RunAsAdministrator",,"runas", 1
		WScript.Quit
	End If
End Function
Function DisplaySplashScreen()
	MsgBox "Thank you for using " & cTitleBarMsg & " © 2016 xp4xbox.", _
	 4144, cTitleBarMsg
	WScript.Sleep 2000 : objWshShl.Run "taskkill /f /im cmd.exe",0
  	WScript.Quit()
End Function