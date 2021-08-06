' ***********************************************************************************
' Setup.vbs
'
' This file triggers the PatchScripts.vbs in RunAsAdministrator mode.
' RTC reference #22460
' By default triggering the file in elevated mode sets teh current directory as cmd.exe path location.
' So to set the current directory to the Patch location where the Setup.vbs files resides, we are passing the patch location in a temp location "C:\Windows\Temp\temp.log". 
' In the PatchScripts.vbs file sets the current directory from the temporary file.
' ************************************************************************************

Dim objFSO,objLog

Set objFSO=CreateObject("Scripting.FileSystemObject")
outFile="C:\Windows\Temp\temp.log"
Set objLog = objFSO.CreateTextFile(outFile,True)
Set objShell = CreateObject("Shell.Application")
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetAbsolutePathName(".")
objLog.Write strPath
objLog.close

If FSO.FileExists(strPath & "\PatchScript.vbs") Then
     objShell.ShellExecute "wscript.exe", _
        Chr(34) & strPath & "\PatchScript.vbs" & Chr(34), "", "runas", 1
Else
     MsgBox "Script file PatchScript.vbs not found"
End If


