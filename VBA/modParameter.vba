'
'
'	Copyright 2021 - 2021 Sutherland Global Services
'
'	Project			: xvba-app		@ => d:\Karthik K Selvan\Testing\Checker
'
'	File			: modParameter.vba		@ => d:\Karthik K Selvan\Testing\Checker\Code\VBA\modParameter.vba
'	Created			: Friday, 30th April 2021 3:20:06 pm
'
'	Author			: IN086482 (karthikk.selvan@sutherlandglobal.com)
'
'	Change History:
'
'
Option Explicit

Dim FSO As New Scripting.FileSystemObject

Public Function Parameter(prm As String, Optional tblParam As ListObject = Nothing, Optional prmCol As String = "Parameter", Optional prmValue As String = "Value") As Range
	Dim colParam As ListColumn, colValue As ListColumn
	If tblParam Is Nothing Then
		Set tblParam = shLocation.ListObjects(1)
	End If
	Set colParam = tblParam.ListColumns(prmCol)
	Set colValue = tblParam.ListColumns(prmValue)
	Set Parameter = Application.WorksheetFunction.XLookup(prm, colParam.DataBodyRange, colValue.DataBodyRange, "", 0)
End Function

Public Sub UpdateQueries()
	ThisWorkbook.Queries.FastCombine = TRUE
	Dim Query As QueryTable, Sheet As Worksheet
	For Each Sheet In ThisWorkbook.Worksheets
		For Each Query In Sheet.QueryTables
			Query.BackgroundQuery = FALSE
		Next
	Next
End Sub
