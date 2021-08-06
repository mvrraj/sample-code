'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PatchScript.vbs
' As per the reference to RTC # 22460 : this file should be triggerred via RunAs Administrator mode.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' --- some handy global variables ---
Public oPD,objFSO,objLog' console object
Dim scount, acount, xcount ' src, als and xml counter
Dim srcfiles( 1999 ), dstfiles( 1999 ), alsfiles( 2000 ), xmlfiles( 1999 ) ' src, dst and als file list
Dim readme, dirlet, ceDir, installLoc, installLocCE, prodNum, datadir ' input parameters
Dim WhatToDo, CRLF, patchDir, productName, version, IsConsoleOpen, patch_no,FSO,Source,Destination ' my own stuff
Const ForWriting = 2
Const ForReading = 1
Set ObjFSO = CreateObject("Scripting.FileSystemObject")
readme = "patchinfo.txt"

' --- product name for branding ---
productName = "FlashScan_1.0.034H_01"

CRLF = Chr( 13 ) & Chr( 10 )
IsConsoleOpen = False
patchDir = "Patch"

Sub InstallPatch()
	' --- read and validate user inputs ---
	objLog.WriteLine vbCrlf & "Patch Installation Process Started. "
	Call GetUserInput()

	
	' --- open console window ---
	On Error Resume Next
   	Call OpenConsole( "Installing patch ..." )
   	BugAssert ( Err.Number = 0 ), "OpenConsole Failed."
	On Error goto 0
	
		
	Call PreValidation()
	
	' --- parse and validate readme file ---
	On Error Resume Next
   	Call ParseReadmeFile()
       	BugAssert ( Err.Number = 0 ), "ParseReadmeFile Failed."
	On Error goto 0
	
	
	 		
	' --- apply patch ---
	On Error Resume Next
	Call ApplyPatch()
	BugAssert (Err.Number = 0), "ApplyPatch Failed."
	On Error goto 0

	' --- create patch uninstaller ---
	On Error Resume Next
	Call CreatePatchUninstaller()
	BugAssert (Err.Number = 0), "CreatePatchUninstaller Failed."
	On Error goto 0

	' On Error Resume Next
	' Call RemovePythonScripts()
	' BugAssert (Err.Number = 0), "RemovePythonScripts Failed."
	' On Error goto 0

	' --- prepare importer ---
	On Error Resume Next
	Call PrepareImporter()
	BugAssert (Err.Number = 0), "PrepareImporter Failed."
	On Error goto 0

	Call StartService()
	
	
	' --- close console window, importer won't use it ---
	Call CloseConsole()

	' --- launch config editor ---
	' On Error Resume Next
	' Call ConfigEditor()
	' BugAssert (Err.Number = 0), "ConfigEditor Failed."
	' On Error goto 0

	' --- launch importer ---
	On Error Resume Next
	Call LaunchImporter()
	BugAssert (Err.Number = 0), "LaunchImporter Failed."
	On Error goto 0
	objLog.WriteLine(Now &" Patch Installation Completed Sucessfully...")
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Source=( "C:\Windows\Temp\" & patchDir & "_Install_" & patch_no & "_" & dt & ".log " )
	Destination= dirlet &"\Install_Logs\"& prodNum &"\"
	FSO.CopyFile Source,Destination
End Sub

Sub GetUserInput()
	Dim fso, objArgs, objFSO, objFolder, strDir1, strDirectory, strDir2, strDir3, strDir4, strDir5

	' --- get readme filename and install root location ---
	Set objArgs = WScript.Arguments
	If ( objArgs.Count < 2 ) Then
		' --- do not prompt for readme file since it's now default ---
		' readme = InputBox( "Readme File Name: ", productName & " Patch Installation", "patchinfo.txt" )
		readme = "patchinfo.txt"
		' --- obtain install location from EBI Product Directory
		Call GetInstallLocation()
		dirlet = InputBox( "Installation Directory: ", productName & " Patch Installation", installLoc )
		objLog.WriteLine(Now &" Patch Directory: " & installLoc )
	Else
		readme = objArgs( 0 )
		dirlet = objArgs( 1 )
	End If

	' --- if filename or location is invalid, report and exit ---
	If ( readme = "" Or dirlet = "" ) Then
		objLog.WriteLine vbCrlf & "Error: Both readme file and drive letter are required."		
		BugAssert 0, "Both readme file and drive letter are required."
	End If

	' --- if file cannot be read, report and exit ---
	Set fso = CreateObject( "Scripting.FileSystemObject" )
	If Not fso.FileExists( readme ) Then
		objLog.WriteLine vbCrlf & "Error: " & readme & " file doesn't exist."
		BugAssert 0, readme & " file doesn't exist."
	End If

	' --- if install location cannot be read, report and exit ---

	If Not fso.FolderExists( dirlet ) Then
		objLog.WriteLine vbCrlf & "Error: " & dirlet & " location doesn't exist."
		BugAssert 0, dirlet & " location doesn't exist."
	End If
End Sub

Sub PreValidation()
	Dim uninstDir, uninstallReadmePath, FSO, objFile, strSearchString, colMatches, ObjFSO, oMatch, lowerVersion , strText , strNewTxt ,t2 ,objtemp
	Dim oRE, oMatches, IsMsgAlert, temp, matchTemp, IsSameVersionInstalled, isHigherVersionInstalled, installingVersionNoStr, installingVersionNo, patchversionNoStr, patchVersionNo
	Const ForReading = 1
	Const ForWriting = 2
	Set oRE = CreateObject("VBScript.RegExp")
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	IsMsgAlert = False
	IsSameVersionInstalled = False
	isHigherVersionInstalled = False
	oRE.Pattern = "((Patch)(\s+)(\d{1}\.)(\d{1}\.)(\d{3}[A-Z])(\_\d{2}))"
	oRE.Global = True
	uninstDir = dirlet & Chr( 92 ) & prodNum & Chr( 92 ) & "_patch"
	If Not FSO.FolderExists(uninstDir) Then
		Exit Sub
	End IF
	
	uninstallReadmePath = FSO.BuildPath(uninstDir, "patchinfo.txt")

	If Not FSO.FileExists(uninstallReadmePath) Then
		objLog.WriteLine(vbCrlf & Now & " Installed Patch Number	: No patch Installed Earlier")
		Exit Sub
	End IF

	
	Set objFile = ObjFSO.OpenTextFile(uninstallReadmePath, ForReading)
    strSearchString = objFile.readAll
	
	
    Set colMatches = oRE.Execute(strSearchString)
	temp = Replace(productName, "FlashScan_", "")
	patchversionNoStr = Mid(temp, Len(temp) - 1, Len(temp))
	patchversionNo = CInt(patchversionNoStr)
    objLog.WriteLine(vbCrlf & Now & " Installed Patch Number	:"+temp )
	If colMatches.Count <> 0 Then
		For Each oMatch In colMatches
			matchTemp = Replace(oMatch.Value, "Patch ", "")
			lowerVersion = Replace(OMatch.Value, "Patch ", "FlashScan_")
			installingVersionNoStr = Mid(MatchTemp, Len(MatchTemp) - 1, Len(MatchTemp))
			installingVersionNo = CInt(installingVersionNoStr)
			
			If matchTemp = temp Then
				IsSameVersionInstalled = True
				Exit For
			ElseIf matchTemp <> temp Then
					If installingVersionNo > patchVersionNo Then
						isHigherVersionInstalled = True
						Exit For
					End If
			End If
		Next
		If (IsSameVersionInstalled) Then
			objLog.WriteLine vbCrlf &  "Error: Patch " & productName & " is installed already." 
			MsgBox "Patch "&productName&" is installed already.", 0, productName & " Patch Installation"
			If (IsConsoleOpen) Then
					CloseConsole()
				End If
			WScript.Quit
		ElseIf (isHigherVersionInstalled) Then
			objLog.WriteLine vbCrlf & "Error: Already installed "&lowerVersion &", failed to install lower version of patch "&productName&". Please uninstall the existing patch and re-try." , 0, productName & " Patch Installation"
			MsgBox "Already installed "&lowerVersion&", failed to install lower version of patch "&productName&". Please uninstall the existing patch and re-try." , 0, productName & " Patch Installation"
			If (IsConsoleOpen) Then
					CloseConsole()
				End If
			WScript.Quit
		End If
    End If
	' objFile.Close
End Sub

Sub GetUserInputCE()
	Dim fso, objArgs
	
	installLocCE = "D:\KTInsptr\1.2.037\MainUI\Bin\Config.xml"
	cedir = InputBox( "Please enter eS31 Config.xml Location for MDS Migration " & CRLF & CRLF & "Default Location is shown Below ", productName & " Patch Installation", installLocCE )


	' --- if filename or location is invalid, report and exit ---
	If ( dirlet = "" ) Then
		objLog.WriteLine vbCrlf & "Error: Drive letter is required."
		BugAssert 0, "Drive letter is required."
	End If

	Set fso = CreateObject( "Scripting.FileSystemObject" )

	' --- if install location cannot be read, report and exit ---
	If Not fso.FileExists( cedir ) Then
		objLog.WriteLine vbCrlf & "Error: " & dirlet & " location doesn't exist."
		BugAssert 0, dirlet & " location doesn't exist."
	End If

End Sub

Sub ParseReadmeFile()
	Dim fso, ts, tl, temp, plist, datadir, xmlDoc ,objNodeList ,plot, x

	objLog.WriteLine vbCrlf & "Processing readme file, please wait..."
	oPD.SetLine2 = "Processing readme file, please wait..."

	' --- some initial values for parsing plist file ---
	acount = 0
	scount = 0
	xcount = 0
	Const ForReading = 1
	prodNum = "1.1.111"

	plist = patchDir & Chr( 92 ) & "plist"
	
	Set fso = CreateObject( "Scripting.FileSystemObject" )
	' --- obtain product number from readme file ---
	Set ts = fso.OpenTextFile( readme, ForReading )
	Do While ts.AtEndOfStream <> True
		tl = ts.ReadLine
		tl = Trim( tl )
		If InStr( 1, tl, "Release:", 1) Then
			temp = Split( tl, ": ", -1, 1 )
			prodNum = temp( 1 )
		End If
		If InStr( 1, tl, "\CougarToPHXAdapter", 1) Then
			objLog.WriteLine vbCrlf & "Cougar Adapter patching is included in this patch , please wait CougarToPHXAdapter services are  stopped..."
			call StopService()
		End If
	Loop
	ts.Close
	
	' --- Obtain the data dir location from <prodNum>\MainUI\bin\config.xml
	Set xmlDoc = CreateObject("Msxml2.DOMDocument")
	xmlDoc.load(dirlet & Chr( 92 ) & prodNum & "\MainUI\bin\config.xml")
	Set objNodeList = xmlDoc.getElementsByTagName("SYSTEM_ROOT")
	If objNodeList.length > 0 then
			' search for SYSTEM_ROOT line & pick data dir from config file
		For each x in objNodeList
			plot=x.Text
			temp = Split( plot, "\KTInsptr", -1, 1 )
			datadir = temp( 0 )
			msgbox   chr(34) & " Data folder Location: "& chr(34) & datadir  
		Next	
	Else
		msgbox chr(34) & "SYSTEM_ROOT" & chr(34) & " not found."
	End If


	' --- if no folder exist for product number, report and exit ---
	If Not fso.FolderExists( dirlet & Chr( 92 ) & prodNum ) Then
		objLog.WriteLine vbCrlf & "Error: No folder "  & dirlet & Chr( 92 ) & prodNum & " found for product number " & prodNum & " ."
		BugAssert 0, "No folder " & dirlet & Chr( 92 ) & prodNum & " found for product number " & prodNum & " ."
	End If
	If Not fso.FolderExists( datadir & Chr( 92 ) & " " ) Then
		BugAssert 0, "No folder " & datadir & Chr( 92 ) & " found for product number " & prodNum & " ."
	End If
	' --- open and start parsing plist file ---
	If Not fso.FileExists( plist ) Then
		objLog.WriteLine vbCrlf & "Error:" & plist & " file doesn't exist."
		BugAssert 0, plist & " file doesn't exist."
	End If

	Set ts = fso.OpenTextFile( plist, ForReading )

	Do While ts.AtEndOfStream <> True
		tl = ts.ReadLine
		tl = Trim( tl )
		If Not ( tl = "" ) Then
			temp = Split( tl, ";", -1, 1 )
			alsfiles( acount ) = temp( 0 )

			If InStr( 1, temp( 2 ), Chr( 92 ) & prodNum & Chr( 92 ) & "BOBsToPatch", 1 ) Then
				xmlfiles( xcount ) = temp( 1 )
				objLog.WriteLine(Now & "value in count : " &xmlfiles( xcount ))
				' --- if xml file cannot be read, report and exit ---
				If Not fso.FileExists( patchDir & Chr( 92 ) & alsfiles( acount ) ) Then
					objLog.WriteLine vbCrlf & "Error:Patch database " & xmlfiles( xcount ) & " file doesn't exist."
					BugAssert 0, "Patch database " & xmlfiles( xcount ) & " file doesn't exist."
				End If
				xcount = xcount + 1
			Else
				srcfiles( scount ) = temp( 1 )
				dstfiles( scount ) = temp( 2 )
				objLog.WriteLine("----------------------------------------------------------------")
				objLog.WriteLine(Now & " Patch source : " &srcfiles( scount ))
				objLog.WriteLine(Now & " Patch destination : " &dirlet &dstfiles( scount ))
				' --- if source file cannot be read, report and exit ---
				If Not fso.FileExists( patchDir & Chr( 92 ) & alsfiles( acount ) ) Then
					objLog.WriteLine "Error: Patch source " & srcfiles( scount ) & " file doesn't exist."
					BugAssert 0, "Patch source " & srcfiles( scount ) & " file doesn't exist."
				End If

				
				If NOT ((InStr(dstfiles( scount ),prodNum)) > 0) Then
					If NOT ((InStr(dstfiles( scount ),":\")) > 0) Then		
						 If Not fso.FolderExists( datadir & dstfiles( scount ) & Chr( 92 ) ) Then
							BugAssert 0, "Patch destination " & datadir & dstfiles( scount ) & Chr( 92 ) & " location doesn't exist."
						End If
					Else
						If Not fso.FolderExists( dstfiles( scount ) & Chr( 92 ) ) Then
						objLog.WriteLine "Error: Patch destination " & dstfiles( scount ) & Chr( 92 ) & " location doesn't exist."
							BugAssert 0, "Patch destination " & dstfiles( scount ) & Chr( 92 ) & " location doesn't exist."
						End If
					End IF	
				Else
					If Not fso.FolderExists( dirlet & dstfiles( scount ) & Chr( 92 ) ) Then
						objLog.WriteLine "Error: Patch destination " & dirlet & dstfiles( scount ) & Chr( 92 ) & " location doesn't exist."
						BugAssert 0, "Patch destination " & dirlet & dstfiles( scount ) & Chr( 92 ) & " location doesn't exist."
					End If
				
				END IF

				scount = scount + 1
			End If
			acount = acount + 1
		End If
	Loop
	ts.Close

	oPD.SetLine2 = " "
	objLog.WriteLine("----------------------------------------------------------------")
	objLog.WriteLine(Now & " File Count : " &acount)
	objLog.WriteLine(Now & " Readme File Validation Completed ")
End Sub

Sub ApplyPatch()
	Dim s, fso, patch, backup, I, junk
	objLog.WriteLine vbCrlf & "******Applying patch with below file list.********"

	Set fso = CreateObject( "Scripting.FileSystemObject" )

        ' walk through patch list ---
	For I = 0 To scount-1
		Set patch = fso.GetFile( patchDir & Chr( 92 ) & alsfiles( I ) )

		' --- make backup prior to patch ---
		'If to check if its not installation directory [check via patch number]
		If NOT ((InStr(dstfiles( I ),prodNum)) > 0) Then 
			If NOT ((InStr(dstfiles( I ),":\")) > 0) Then
				If ( fso.FileExists( datadir & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) ) ) Then
					Set backup = fso.GetFile( datadir & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) )
					' --- skip if original exists, trust keep file and never overwrite it ---
					If Not (fso.FileExists( datadir & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) & ".keep" )) Then 
						If Not (fso.FileExists( datadir & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) & ".patch" )) Then
							backup.Move( datadir & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) & ".keep" )
						End If
					Else
						backup.Delete True
					End If
				Else 
					patch.Copy( datadir & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) )
					Set backup = fso.GetFile( datadir & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) )
					backup.Move( datadir & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) & ".patch" )
				End If
			Else
				If ( fso.FileExists( dstfiles( I ) & Chr( 92 ) & srcfiles( I ) ) ) Then
					Set backup = fso.GetFile( dstfiles( I ) & Chr( 92 ) & srcfiles( I ) )
					' --- skip if original exists, trust keep file and never overwrite it ---
					If Not (fso.FileExists( dstfiles( I ) & Chr( 92 ) & srcfiles( I ) & ".keep" )) Then 
						If Not (fso.FileExists( dstfiles( I ) & Chr( 92 ) & srcfiles( I ) & ".patch" )) Then
							backup.Move( dstfiles( I ) & Chr( 92 ) & srcfiles( I ) & ".keep" )
						objLog.WriteLine("----------------------------------------------------------------")
						objLog.WriteLine(Now &" Keep File created : " &dirlet & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) & ".keep" )
						End If
					Else
						backup.Delete True
						objLog.WriteLine("----------------------------------------------------------------")
						objLog.WriteLine(Now &" Keep File Exist : " &dirlet & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) & ".keep" )	
				
					End If
				Else 
					patch.Copy(dstfiles( I ) & Chr( 92 ) & srcfiles( I ))
					Set backup = fso.GetFile( dstfiles( I ) & Chr( 92 ) & srcfiles( I ) )
					backup.Move( dstfiles( I ) & Chr( 92 ) & srcfiles( I ) & ".patch" )
				End If
			
			End IF
		Else
			If ( fso.FileExists( dirlet & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) ) ) Then
				Set backup = fso.GetFile( dirlet & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) )
			' --- skip if original exists, trust keep file and never overwrite it ---
				If Not (fso.FileExists( dirlet & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) & ".keep" )) Then 
					If NOT (fso.FileExists( dirlet & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) & ".patch" )) Then
						backup.Move( dirlet & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) & ".keep" )
						objLog.WriteLine("----------------------------------------------------------------")
						objLog.WriteLine(Now &" Keep File created : " &dirlet & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) & ".keep" )
					End If
				Else
					backup.Delete True
					objLog.WriteLine("----------------------------------------------------------------")
					objLog.WriteLine(Now &" Keep File Exist : " &dirlet & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) & ".keep" )
				End If
			Else
				patch.Copy(dirlet & dstfiles( I ) & Chr( 92 ) & srcfiles( I ))
				Set backup = fso.GetFile( dirlet & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) )
				backup.Move( dirlet & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) & ".patch" )
				objLog.WriteLine("----------------------------------------------------------------")
				objLog.WriteLine(Now &" patch File created : " &dirlet & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) & ".patch" )
			End If
			
		End IF
				' --- apply patch through blind copy ---
		oPD.SetLine2 = " " & srcfiles( I ) & " ..."
		If NOT ((InStr(dstfiles( I ),prodNum)) > 0) Then
			If NOT ((InStr(dstfiles( I ),":\")) > 0) Then
				patch.Copy( datadir & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) )
			Else
				patch.Copy( dstfiles( I ) & Chr( 92 ) & srcfiles( I ) )
				objLog.WriteLine("----------------------------------------------------------------")
				objLog.WriteLine(Now &" File copied : "  & dstfiles( I ) & Chr( 92 ) & srcfiles( I ))	
			End if
		Else
			patch.Copy( dirlet & dstfiles( I ) & Chr( 92 ) & srcfiles( I ) )
			objLog.WriteLine("----------------------------------------------------------------")
			objLog.WriteLine(Now &" File copied : " &dirlet & dstfiles( I ) & Chr( 92 ) & srcfiles( I ))
		End IF

		oPD.PctComplete = Int( ( 100/scount ) * ( I+1 ) )
	        WScript.Sleep 200
	Next

	oPD.SetLine2 = " "
	oPD.SetLine2 = CStr( scount ) & " files were patched."
	WScript.Sleep 2000 ' one second
	objLog.WriteLine("----------------------------------------------------------------")
	objLog.WriteLine(Now &" Applying Patch Completed")
End Sub

Sub RemovePythonScripts()
	Dim names, names1, names2, fso, baseDir, baseDir1, baseDir2, pyFilename, name

	' --- remove some python scripts ---
	names = Array("ACBtoIIB_CRC_Diagonal_BCD_Ramp_100MPPS_test.py", "ACBtoIIB_CRC_Diagonal_BCD_Ramp_200MPPS_test.py", "ACBtoIIB_CRC_Diagonal_BCD_Ramp_400MPPS_test.py", "ACBtoIIB_CRC_Diagonal_BCD_Ramp_800MPPS_test.py", "ACBtoIIB_CRC_Diagonal_Binary_Ramp_100MPPS_test.py", "ACBtoIIB_CRC_Diagonal_Binary_Ramp_200MPPS_test.py", "ACBtoIIB_CRC_Diagonal_Binary_Ramp_400MPPS_test.py", "ACBtoIIB_CRC_Diagonal_Binary_Ramp_800MPPS_test.py", "ACBtoIIB_CRC_Horizontal_BCD_Ramp_100MPPS_test.py", "ACBtoIIB_CRC_Horizontal_BCD_Ramp_200MPPS_test.py", "ACBtoIIB_CRC_Horizontal_BCD_Ramp_400MPPS_test.py", "ACBtoIIB_CRC_Horizontal_BCD_Ramp_800MPPS_test.py", "ACBtoIIB_CRC_Horizontal_Binary_Ramp_100MPPS_test.py", "ACBtoIIB_CRC_Horizontal_Binary_Ramp_200MPPS_test.py", "ACBtoIIB_CRC_Horizontal_Binary_Ramp_400MPPS_test.py", "ACBtoIIB_CRC_Horizontal_Binary_Ramp_800MPPS_test.py", "ACBtoIIB_CRC_Vertical_Binary_Ramp_100MPPS_test.py", "ACBtoIIB_CRC_Vertical_Binary_Ramp_200MPPS_test.py", "ACBtoIIB_CRC_Vertical_Binary_Ramp_400MPPS_test.py", "ACBtoIIB_CRC_Vertical_Binary_Ramp_800MPPS_test.py")

	names1 = Array("ACBToIIB_CRC_Diagonal_BCD_Ramp_100MPPS_test0.txt", "ACBToIIB_CRC_Diagonal_BCD_Ramp_200MPPS_test0.txt", "ACBToIIB_CRC_Diagonal_BCD_Ramp_400MPPS_test0.txt", "ACBToIIB_CRC_Diagonal_BCD_Ramp_800MPPS_test0.txt", "ACBToIIB_CRC_Diagonal_Binary_Ramp_100MPPS_test0.txt", "ACBToIIB_CRC_Diagonal_Binary_Ramp_200MPPS_test0.txt", "ACBToIIB_CRC_Diagonal_Binary_Ramp_400MPPS_test0.txt", "ACBToIIB_CRC_Diagonal_Binary_Ramp_800MPPS_test0.txt", "ACBToIIB_CRC_Horizontal_BCD_Ramp_100MPPS_test0.txt", "ACBToIIB_CRC_Horizontal_BCD_Ramp_200MPPS_test0.txt", "ACBToIIB_CRC_Horizontal_BCD_Ramp_400MPPS_test0.txt", "ACBToIIB_CRC_Horizontal_BCD_Ramp_800MPPS_test0.txt", "ACBToIIB_CRC_Horizontal_Binary_Ramp_100MPPS_test0.txt", "ACBToIIB_CRC_Horizontal_Binary_Ramp_200MPPS_test0.txt", "ACBToIIB_CRC_Horizontal_Binary_Ramp_400MPPS_test0.txt", "ACBToIIB_CRC_Horizontal_Binary_Ramp_800MPPS_test0.txt", "ACBToIIB_CRC_Vertical_Binary_Ramp_100MPPS_test0.txt", "ACBToIIB_CRC_Vertical_Binary_Ramp_200MPPS_test0.txt", "ACBToIIB_CRC_Vertical_Binary_Ramp_400MPPS_test0.txt", "ACBToIIB_CRC_Vertical_Binary_Ramp_800MPPS_test0.txt")

	names2 = Array("DsmTest_0.py", "Test1.py", "Test2.py", "Test3.py")

	Set fso = CreateObject( "Scripting.FileSystemObject" )
	baseDir = dirlet & Chr( 92 ) & prodNum & Chr( 92 ) & "Cougar\Es35ImageComputerTests\InterfaceTests\ACB-IIB\"

	baseDir1 = dirlet & Chr( 92 ) & prodNum & Chr( 92 ) & "Cougar\Es35ImageComputerTests\GoldenFiles\ACB-IIB\"

	baseDir2 = dirlet & Chr( 92 ) & prodNum & Chr( 92 ) & "Cougar\ImageComputerTests\BoardDiagnosticsTests\EverestPythonDiag\"

	For Each name In names
		pyFilename = baseDir & name
		If fso.FileExists( pyFilename ) Then
			fso.DeleteFile pyFilename, True
		End If
	Next

	For Each name In names1
		pyFilename = baseDir1 & name
		If fso.FileExists( pyFilename ) Then
			fso.DeleteFile pyFilename, True
		End If
	Next

	For Each name In names2
		pyFilename = baseDir2 & name
		If fso.FileExists( pyFilename ) Then
			fso.DeleteFile pyFilename, True
		End If
	Next
End Sub


Sub CreatePatchUninstaller()
	Dim fso, fpl, p, r, s, c, I, objFolder, colFiles, objFile
	Dim uninstDir, patchDir, consoleFile, uninstallFile
	Const ReadOnly = 1
	oPD.SetLine2 = "Creating uninstaller, please wait..."

	' --- some initial values for uninstaller ---
	uninstDir = dirlet & Chr( 92 ) & prodNum & Chr( 92 ) & "_patch"
	patchDir = uninstDir & Chr( 92 ) & "patch"
	consoleFile = "PatchConsole.wsc"
	uninstallFile = "UninstallPatch.txt"

	Set fso = CreateObject( "Scripting.FileSystemObject" )
	' --- cleanup patch folder, if exist ---
	If fso.FolderExists( uninstDir ) Then
	Set objFolder = fso.GetFolder(uninstDir)
	
	Set colFiles = objFolder.Files

	For Each objFile in colFiles
		If objFile.Attributes AND ReadOnly Then
			objFile.Attributes = objFile.Attributes XOR ReadOnly
		End If
	Next
	
	If objFolder.Attributes AND ReadOnly Then
	objFolder.Attributes = objFolder.Attributes XOR ReadOnly
	End If
 
		fso.DeleteFolder uninstDir, True
		WScript.Sleep 1000
		On Error Resume Next
	End If

	' --- create uninstall and patch folders ---
	fso.CreateFolder( uninstDir )
	WScript.Sleep 1000
	On Error Resume Next

	fso.CreateFolder( patchDir )
	WScript.Sleep 1000
	' --- copy readme file ---
	If Not fso.FileExists( readme ) Then
		objLog.WriteLine "Error:" & readme & " file doesn't exist."
		BugAssert 0, readme & " file doesn't exist."
	End If
	Set r = fso.GetFile( readme )
	r.Copy( uninstDir & Chr( 92 ) & readme )

	' --- copy console script ---
	If Not fso.FileExists( consoleFile ) Then
		objLog.WriteLine "Error:" & consoleFile & " file doesn't exist."
		BugAssert 0, consoleFile & " file doesn't exist."
	End If
	Set c = fso.GetFile( consoleFile )
	c.Copy( uninstDir & Chr( 92 ) & consoleFile )

	' --- copy uninstall script ---
	If Not fso.FileExists ( uninstallFile ) Then
		objLog.WriteLine "Error:" & uninstallFile & " file doesn't exist."
		BugAssert 0, uninstallFile & " file doesn't exist."
	End If
	Set s = fso.GetFile( uninstallFile )
	s.Copy( uninstDir & Chr( 92 ) & "UninstallPatch.vbs" )

	' --- write plist file ---
	Set fpl = fso.CreateTextFile( patchDir & Chr( 92 ) & "plist" )
	For I = 0 To scount-1
		If NOT ((InStr(dstfiles( I ),prodNum)) > 0) Then
			If NOT ((InStr(dstfiles( I ),":\")) > 0) Then
				fpl.WriteLine( srcfiles( I ) & ";" & datadir & dstfiles( I ) )
			Else
				fpl.WriteLine( srcfiles( I ) & ";"  & dstfiles( I ) )
			End IF
		Else
			fpl.WriteLine( srcfiles( I ) & ";" & dirlet  & dstfiles( I ) )
		End IF
	Next
	fpl.Close

	oPD.SetLine2 = " "
	objLog.WriteLine(Now &" Patch uninstaller Created")
End Sub

Sub PrepareImporter()
	If (xcount = 0) Then Exit Sub

	oPD.SetLine2 = "Creating importer, please wait..."

	Dim fso, fpl, stageDir, bobsfile, xml, I

	' --- some initial values for importer ---
	stageDir = "C:\Temp\BOBsToPatch"
	bobsFile = stageDir & Chr( 92 ) & "BOBsToPatch.txt"

	Set fso = CreateObject( "Scripting.FileSystemObject" )

	' --- destroy stage if exists ---
	If fso.FolderExists( stageDir ) Then
		fso.DeleteFolder stageDir, True
	End If
	fso.CreateFolder( stageDir )

	' --- copy xml files to stage while creating bobsfile ---
	Set fpl = fso.CreateTextFile( bobsfile )
	For I = 0 To xcount-1
		Set xml = fso.GetFile( patchDir & Chr( 92 ) & alsfiles( scount + I ) )
		xml.Copy( stageDir & Chr( 92 ) & xmlfiles( I ) )
		fpl.WriteLine( stageDir & Chr( 92 ) & xmlfiles( I ) )
	Next
	fpl.Close

	oPD.SetLine2 = " "
End Sub

Sub LaunchImporter()
	If (xcount = 0) Then
		MsgBox CStr( scount ) & " files were patched. ", 64, productName & " Patch Installation"
		Exit Sub
	End If

	' --- prompt user about xml upload ---
	WhatToDo = MsgBox( CStr( scount ) & " files were patched. Found " & CStr( xcount ) &_
                " database file(s) to import! " & CRLF & CRLF & "Do you want to import data" &_
                " file(s) now? ", 68, productName & " Patch Installation" )

	' --- if user response is no, bail out ---
	Select Case WhatToDo
		Case "7"
			Exit Sub
	End Select

	Dim sh, wkDir, hdr, cmd, cmdStr, stageDir, bobsfile

	' --- some initial values for importer ---
	stageDir = "C:\Temp\BOBsToPatch"
	bobsFile = stageDir & Chr( 92 ) & "BOBsToPatch.txt"

	wkDir = dirlet & Chr( 92 ) & prodNum & Chr( 92 ) & "CommandScripts"
        hdr = "TITLE " & productName & " Patch Installation"
	cmd = "ImportFileListP.cmd ""Startup IDS"" " & """" & bobsfile & """"
	cmdStr = "CMD /C CD /D " & wkDir & " & " & hdr & " & " & cmd & " & " & " pause"

	' --- launch importer ---
	Set sh = CreateObject( "WScript.Shell" )
	sh.Run cmdStr, 1, true
End Sub


Sub ConfigEditor()
	Dim sh, wkDir, hdr, cmd, cmdStr

	' --- prompt user about xml upload ---
	WhatToDo = MsgBox( "All Files were patched." & " Do you want to do MDS Migration from eS31 now? ", 68, productName & " Patch Installation" )

	' --- if user response is no, bail out ---
	Select Case WhatToDo
		Case "7"
			Exit Sub
	End Select

	Call GetUserInputCE()
	wkDir = dirlet & Chr( 92 ) & prodNum & Chr( 92 ) & "CommandScripts"
        hdr = "TITLE " & productName & " Patch Installation"
	cmd = "MeteorConfigEditor.cmd " & """" & cedir & """"	
	cmdStr = "CMD /C CD /D " & wkDir & " & " & cmd

	Set sh = CreateObject( "WScript.Shell" )
	sh.Run cmdStr, 1, true

End Sub


Sub OpenConsole( strTask )
	Dim WshShell
	On Error Resume Next
	
	Set WshShell = WScript.CreateObject("WScript.Shell")
	WshShell.run "regsvr32 /u /s PatchConsole.wsc", 1, true
	WScript.Sleep 1000
	WshShell.run "regsvr32 /s PatchConsole.wsc", 1, true
		
		' --- initiate console window ---
	Set oPD = WScript.CreateObject( "PatchConsole.Scriptlet", "oPD_" )
	If ( IsObject( oPD ) = FALSE ) Then
		BugAssert 0, "Instantiate PatchConsole Failed."
	End If
	' --- create the form, and add the controls ---
	oPD.Open "FlashScan Patch Installation", 60, 100, 400, 200
	oPD.Show TRUE
	oPD.SetLine1 = strTask
	oPD.SetLine2 = " "
  	' --- finished with creating the form ---
	IsConsoleOpen = True
End Sub

Sub CloseConsole()
	Dim WshShell,strNewTxt
	objLog.writeline (vbCrlf & strNewTxt)
	objLog.WriteLine vbCrlf & "****** Applying patch completed Successfully.********"
	' --- destroy the form ---
	oPD.Close  ' clean up
	Set oPD = Nothing
	IsConsoleOpen = False
	' --- clean up console window ---
	Set WshShell = WScript.CreateObject( "WScript.Shell" )
	WshShell.Run "regsvr32 /u /s PatchConsole.wsc", 1, true
End Sub

Sub oPD_Cancel()
	Call CloseConsole()
	BugAssert 0, "Patch Installation CANCELLED."
End Sub

Sub oPD_OnUnLoad()
	Set oPD = Nothing
	' --- clean up console window ---
	Set WshShell = WScript.CreateObject( "WScript.Shell" )
	WshShell.Run "regsvr32 /u /s PatchConsole.wsc", 1, true
End Sub

Sub GetInstallLocation()
	Dim fso, ts, tl, regfile

	regfile = "C:\WINNT\epd.properties"

	' --- set default installation location ---
	installLoc = "D:\KTInsptr"

	' --- open EBI product directory ---
	Set fso = CreateObject( "Scripting.FileSystemObject" )
	If fso.FileExists( regfile ) Then
		Set ts = fso.OpenTextFile( regfile, 1 )
		Do While ts.AtEndOfStream <> True
			tl = ts.ReadLine
			tl = Trim( tl )
			If ( tl <> "" ) Then
				If fso.FolderExists( tl ) Then
					' --- reset default installation location ---
					installLoc = tl
				End If
			End IF
		Loop
		ts.Close
	End If

End Sub

Sub BugAssert( bTest, sErrMsg )
	If bTest Then Exit Sub

	If ( Err.Number = 0 ) Then
		MsgBox "Error: " & sErrMsg & CRLF & CRLF, vbCritical, productName & " Patch Installation"
		objLog.WriteLine(Now & " Error: " & sErrMsg & CRLF & CRLF & vbCritical & productName & " Patch Installation")
	Else
		MsgBox "Error: " & sErrMsg & CRLF & CRLF & "Error#: " & CStr( Err.Number ) & ", " & Err.Description & CRLF & CRLF,_
			vbCritical, productName & " Patch Installation"
		objLog.WriteLine(Now & " Error: " & sErrMsg & CRLF & CRLF & " Error#: " & CStr( Err.Number ) & ", " & Err.Description & CRLF & CRLF &_
			vbCritical & productName & " Patch Installation")
	End If

	If ( IsConsoleOpen ) Then
		Call CloseConsole()
	End If

  	WScript.Quit
End Sub

Sub Readdata()
	Dim ts,tl,temp
	Set ts = objFSO.OpenTextFile( readme, ForReading )
			Do While ts.AtEndOfStream <> True
				tl = ts.ReadLine
				tl = Trim( tl )
				If InStr( 1, tl, "Release:", 1) Then
				temp = Split( tl, ": ", -1, 1 )
				prodNum = temp( 1 )

				End If
				If InStr( 1, tl, "Patch No.:", 1) Then
					temp = Split( tl, ": ", -1, 1 )
					patch_no = temp( 1 )
				End If
			Loop
	ts.Close
End Sub

sub StopService()
	Dim objWMIService, objService,sTargetSvc,strService,strComputer,colListOfServices
	strService = Array("Cougar to Phoenix Adapter Service","Cougar Adapter ActiveMQ")
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	For Each sTargetSvc In strService
		objLog.WriteLine(Now & "Current State of :: "& sTargetSvc & " :: to Stopped"  )	
		Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service Where DisplayName='"& sTargetSvc & "'")
		For Each objService in colListOfServices
			If LCase(objService.Name) = LCase(sTargetSvc) Then
				If objService.State <> "Stopped" Then
					' WScript.Echo "Current State of :: "& sTargetSvc & " :: "& objService.State  
					objService.StopService()
				End If 
			End If  
		Next
	Next
End Sub

sub StartService()
	Dim objWMIService, objService,sTargetSvc,strService,strComputer,colListOfServices
	strService = Array("Cougar to Phoenix Adapter Service","Cougar Adapter ActiveMQ")
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	For Each sTargetSvc In strService
		objLog.WriteLine(Now & "Current State of :: "& sTargetSvc & " :: to Running"  )	
		Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service Where DisplayName='"& sTargetSvc & "'")
		For Each objService in colListOfServices
			If LCase(objService.Name) = LCase(sTargetSvc) Then
				If objService.State <> "Running" Then
					' WScript.Echo "Current State of :: "& sTargetSvc & " :: "& objService.State  
					objService.StartService()
				End If 
			End If  
		Next
	Next
End sub

	
' --- main program ---
Dim fso2,file,FileContent,oShell

' Reads the temporary file to set the current directory
Set fso2 = CreateObject( "Scripting.FileSystemObject" )
Set file = fso2.OpenTextFile ("C:\Windows\Temp\temp.log", 1)
FileContent = file.ReadAll
Set oShell = CreateObject("WScript.Shell")
oShell.CurrentDirectory = fso2.GetAbsolutePathName(FileContent)

WhatToDo = MsgBox( "Welcome to " & productName & " Patch Installer! Do you want to Install this Patch?", 68, productName & " Patch Installation" )
call Readdata()

Dim dt
dt=Replace(Now,"/","_")
dt=Replace(dt,":","_")
dt=Replace(dt," ","_")

' Create the patch log file
Set objLog = objFSO.CreateTextFile( "C:\Windows\Temp\" & patchDir & "_Install_" & patch_no & "_" & dt & ".log " )
objLog.WriteLine "****************    PATCH INSTALLATION LOG   ******************"
objLog.WriteLine(Now & " Patch process started.")
objLog.WriteLine(Now & " Build Number 		: "& prodNum)
objLog.WriteLine(Now & " Current Patch Number 	: "& patch_no)
If WhatToDo = VbNo then
	objLog.WriteLine(Now & " Patch process Aborted by User.") 
 End If

' call StopService()
 
Select Case WhatToDo
	Case "6"
		Call InstallPatch()
	Case "7"
		WScript.Quit
End Select