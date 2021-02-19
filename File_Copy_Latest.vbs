'****************************************************************************************
'	Copy Latest File Script								*
'											*
'	Created: 15/08/2013 by Kish Jogia						*
'											*
'****************************************************************************************
'     Date	|     Name	|   Descritption				|Version*
'---------------|---------------|-------------------------------------------------------*
'  15/08/2012   |  Kish Jogia   | Intial creation                               | 1.0.0 *
'  03/11/2014   |  Kish Jogia   | Recursively moves through folders		| 1.1.0 *
'               |               |                                               |       *
'****************************************************************************************

Option Explicit

'Const TestMode = True		'Enable test mode, by setting true, this will NOT run any programs
Const TestMode = False		'Enable test mode, by setting true, this will NOT run any programs

'// SETTINGS
Const ForReading = 1, ForWriting = 2, ForAppending = 8, OverwriteExisting = True


'// DECLARATIONS
Dim fso
Dim srcPath
Dim tgtPath
Dim date

Set fso = WScript.CreateObject("Scripting.FilesystemObject")

srcPath = "C:\Users\user\Documents\Test\"
tgtPath = "C:\Users\user\Documents\Test1\"

date = WScript.Arguments(0)

ReplaceIfNewer srcPath, tgtPath, date


'*******************************
'	CLEANUP	
'*******************************

Set fso = Nothing		'Close object

Wscript.Echo "Script completed!"

'*******************************
'	Extra Function
'*******************************

Sub ReplaceIfNewer(strSourcePath, strTargetPath, orgDate)
	Const OVERWRITE_EXISTING = True
	Dim objFso, objFolder, objSubFolder
	Dim objTargetFile, objTargetFolder, newTargetFolder
	Dim dtmTargetDate
	Dim objSourceFile, objSourceFolder
	Dim dtmSourceDate
	Dim colFiles, fldName, fldCnt, fld

'	wscript.echo orgDate

	Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
	Set objSourceFolder = objFso.GetFolder(strSourcePath)
	Set colFiles = objSourceFolder.Files

	If Not objfso.FolderExists(strTargetPath) Then
		objFSO.CreateFolder strTargetPath
	End If

	Set objTargetFolder = objFso.GetFolder(strTargetPath)

	For Each objSourceFile in colFiles

		objTargetFile = objTargetFolder & "\" & objFso.GetFileName(objSourceFile)

		If Not objfso.FileExists(objTargetFile) Then
			objfso.CopyFile objSourceFile, objTargetFile, True
		Else
			set objTargetFile = objFso.GetFile(objTargetFile)

			'Get Target file date modified
			dtmTargetDate = objTargetFile.DateLastModified

			'Get Source file data modified
			dtmSourceDate = objSourceFile.DateLastModified

			If (dtmTargetDate < orgDate) Then
				objFso.CopyFile objSourceFile.Path, objTargetFile.Path, OVERWRITE_EXISTING
			End If
		End If
	Next

	For Each objSubFolder in objSourceFolder.SubFolders

		Set objFolder = objFso.GetFolder(objSubFolder.Path)
		Set colFiles = objFolder.Files

		For Each objSourceFile in colFiles

			fldName = Split(objFolder, "\")
			fldCnt = 0
			For Each fld in fldName
				fldCnt = fldCnt + 1
			Next

			newTargetFolder = strTargetPath & "\" & fldName(fldCnt-1)

			ReplaceIfNewer objSubFolder, newTargetFolder , orgDate

		Next
	Next

	Set objFso = Nothing
End Sub 

'Error Handling
If Err.Number <> 0 Then
	Wscript.Echo "Error!" & vbCrLf & Err.Description
End If
