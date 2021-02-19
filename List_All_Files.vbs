'****************************************************************************************
'	Created: 09/03/2016 by Kish Jogia						*
'											*
'****************************************************************************************
'     Date	|     Name	|   Descritption				|Version*
'---------------|---------------|-----------------------------------------------|-------*
'  09/03/2016   |  Kish Jogia   | Intial creation                               | 1.0.0	*
'               |               |                                               |	*
'****************************************************************************************

Option Explicit

Const TestMode = False		'Enable test mode, by setting true, this will NOT run any programs

'// SETTINGS
Const ForReading = 1, ForWriting = 2, ForAppending = 8, OverwriteExisting = True

Const systemFolder = "C:\Users"

'// DECLARATIONS
Dim objFSO, objFile, objFolder, objSubFolder, objShell, logFSO, logFile, outputFile
Dim sText, strDate, timeStamp, file(), path(), foundArray(), version(), line
Dim i, j, foundCnt
Dim wshNetwork, strHostName, workingFolder, cmdReturn, ipaddress

If TestMode <> True Then
'	On Error Resume Next				'Enable error handling
End If

'Intialisation
Set objFSO = CreateObject("Scripting.FileSystemObject")		'Create Object

'Ouput test messages to a file
If TestMode = True Then
	Set logFSO = CreateObject("Scripting.FileSystemObject")		'Create Object
	Set logFile = logFSO.OpenTextFile("File_List_Script_Log.txt", ForAppending, True)
	timestamp = Now()
	logFile.WriteLine (timeStamp & " Log Started")
End If

'Get Hostname
Set wshNetwork = WScript.CreateObject("WScript.Network")
strHostName = wshNetwork.ComputerName

'Convert hostname to Uppercase
strHostName = UCase(strHostName)

'Get the date
strDate = month(Now) & "-" & day(Now) & "-" & year(Now)

Set outputFile = objFSO.OpenTextFile("File_List_" & strHostName & "_" & strDate & ".csv", ForWriting, True)	'Open File

'Get file version
Set objFolder = objFSO.GetFolder(systemFolder)

ReDim file(100)
ReDim path(100)
ReDim version(100)
ReDim foundArray(100)
foundCnt = 0

'Get all of the file versions and compare them
ListFiles objFolder

'CLEANUP
If TestMode Then
	logFile.Close
	Set logFile = Nothing
	Set logFSO = Nothing
End If
Set objFile = Nothing		'Close file
Set objFSO = Nothing		'Close object
Set objShell = Nothing		'Close object

MsgBox "Script completed!"

'Extra function
Sub ListFiles (Folders)
	For Each objFile in Folders.Subfolders

		ReDim preserve file(foundCnt)
		ReDim preserve path(foundCnt)
		ReDim preserve version(foundCnt)
		ReDim preserve foundArray(foundCnt)

		path(foundCnt) = objFSO.GetParentFolderName(objFile)
		file(foundCnt) = objFile.Name
		version(foundCnt) = objFSO.GetFileVersion(objFile)

		outputFile.WriteLine path(foundCnt) & "\" & file(foundCnt)

		If TestMode = True Then
			timestamp = Now()
			logFile.WriteLine (timeStamp & " file = " & file(foundCnt - 1))
		End If
	Next

'	For Each objSubFolder in Folders.Subfolders
'		ListFiles objSubFolder
'	Next
End Sub

'Error Handling
If Err.Number <> 0 Then
	MsgBox "Error!" & vbCrLf & Err.Description
End If
