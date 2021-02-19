'****************************************************************************************
'	Update Log Audit Folder Structure						*
'											*
'	Created: 13/01/2017 by Kish Jogia						*
'											*
'****************************************************************************************
'     Date	|     Name	|   Descritption				|Version*
'---------------|---------------|-------------------------------------------------------*
'  13/01/2017   |  Kish Jogia   | Intial creation                               | 1.0.0 *
'****************************************************************************************

Option Explicit

'Const TestMode = True		'Enable test mode, by setting true, this will NOT run any programs
Const TestMode = False		'Enable test mode, by setting true, this will NOT run any programs

'// SETTINGS
Const ForReading = 1, ForWriting = 2, ForAppending = 8, OverwriteExisting = True
Const AuditPath = "c:\path"				'Path to store the log file
Dim keepFolders
'keepFolders = 7		'number of historical folders to keep

'// DECLARATIONS
Dim objFSO, objFile, objFolder, objShell, objRegistry, logFSO, logFile, fileSize
Dim sText, strDate, timeStamp, strValue, subFolder, Folders
Dim foldersExist, filesExist, stringsExist, stringCount
Dim wshNetwork, strHostName, workingFolder, i, cmdReturn, dbString

'Intialisation
Set objFSO = CreateObject("Scripting.FileSystemObject")		'Create Object

'Ouput test messages to a file
If TestMode = True Then
	Set logFSO = CreateObject("Scripting.FileSystemObject")		'Create Object
	Set logFile = logFSO.OpenTextFile("Backup_Script_Log.txt", ForAppending, True)
	timestamp = Now()
	logFile.WriteLine (timeStamp & " Log Started")
End If

' Check if Folders Exist
If objFSO.FolderExists(AuditPath & "\" & year(Now)) Then
	If TestMode = True Then
		timestamp = Now()
		logFile.WriteLine (timeStamp & "Folder Exists " & AuditPath & "\" & year(Now))
	End If

	If objFSO.FolderExists(AuditPath & "\" & year(Now) & "\" & month(Now)) Then
		If TestMode = True Then
			timestamp = Now()
			logFile.WriteLine (timeStamp & "Folder Exists " & AuditPath & "\" & year(Now) & "\" & month(Now))
		End If

	Else
		' Create new months Folder
		objFSO.CreateFolder(AuditPath & "\" & year(Now) & "\" & month(Now))

	End If
Else
	' Create new months Folder
	objFSO.CreateFolder(AuditPath & "\" & year(Now))

End If
