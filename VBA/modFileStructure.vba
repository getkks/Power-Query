'
'
'	Copyright 2021 - 2021 Sutherland Global Services
'
'	Project			: xvba-app		@ => d:\Karthik K Selvan\Testing\Checker
'
'	File			: modFileStructure.vba		@ => d:\Karthik K Selvan\Testing\Checker\Code\VBA\modFileStructure.vba
'	Created			: Thursday, 21st October 2021 5:25:37 pm
'
'	Author			: IN086482 (karthikk.selvan@sutherlandglobal.com)
'
'	Change History:
'
'
Option Explicit

Public Sub CreateDir(strPath As String)
	Dim elm As Variant
	Dim strCheckPath As String
	On Error Resume Next
	strCheckPath = ""
	For Each elm In Split(strPath, "\")
		strCheckPath = strCheckPath & elm & "\"
		If Len(Dir(strCheckPath, vbDirectory)) = 0 Then MkDir strCheckPath
	Next
	If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub UpdateFolders()
	If ThisWorkbook.Name = "PQ Template.xltm" Then Exit Sub
	Dim FSO As New FileSystemObject
	If Not FSO.FolderExists(Parameter("Data Path").Value) Then CreateDir Parameter("Data Path").Value
	If Not FSO.FolderExists(Parameter("Output Path").Value) Then CreateDir Parameter("Output Path").Value
	If Not FSO.FolderExists(Parameter("File Support Path").Value) Then CreateDir Parameter("File Support Path").Value
	If Not FSO.FolderExists(Parameter("Power Query").Value) Then CreateDir Parameter("Power Query").Value
End Sub
