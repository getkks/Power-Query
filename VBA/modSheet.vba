'
'
'	Copyright 2021 - 2021 Sutherland Global Services
'
'	Project			: xvba-app		@ => d:\Karthik K Selvan\Testing\Checker
'
'	File			: modSheet.vba		@ => d:\Karthik K Selvan\Testing\Checker\Code\VBA\modSheet.vba
'	Created			: Thursday, 21st October 2021 5:43:12 pm
'
'	Author			: IN086482 (karthikk.selvan@sutherlandglobal.com)
'
'	Change History:
'
'
Option Explicit

Public Function TableUnion(Sheet As Worksheet) As Range
	Dim Table As ListObject
	For Each Table In Sheet.ListObjects
		Set TableUnion = Union(Table.Range, IIf(TableUnion Is Nothing, Table.Range, TableUnion))
	Next
End Function

Public Sub RefreshSheets(SheetsWithData)
	Dim Sh, Tbl As ListObject
	TrackProgress UBound(SheetsWithData) - LBound(SheetsWithData) + 1
	For Each Sh In SheetsWithData
		For Each Tbl In Sh.ListObjects
			Debug.Print "Table Refresh: " & Tbl.Name
			Debug.Assert Tbl.Name <> "AribaInvoices"
			Tbl.Refresh
			With Tbl
				.TableStyle = "TableStyleLight9"
				.ShowTableStyleFirstColumn = False
				.ShowTableStyleColumnStripes = False
				.ShowTableStyleLastColumn = False
				.ShowTableStyleRowStripes = False
				.ShowAutoFilter = False
			End With
		Next
		UpdateProgress
	Next
End Sub

Public Sub ToggleSheetVisibility(SheetsWithData, Visible As XlSheetVisibility)
	Dim Sh
	TrackProgress UBound(SheetsWithData) - LBound(SheetsWithData) + 1
On Error Resume Next
	For Each Sh In SheetsWithData
		Sh.Visible = Visible
		UpdateProgress
	Next
On Error GoTo 0
End Sub

public Function CopySheets(SheetsToCopy, Optional ToBook As Workbook) As Workbook
	TrackProgress Sheets.Count + 1
	Application.ScreenUpdating = False
	Dim wn As String, Nw As String
	wn = ThisWorkbook.Name
	If ToBook Is Nothing Then
		Workbooks.Add
		Nw = ActiveWorkbook.Name
	Else
		Nw = ToBook.Name
	End If
	Set CopySheets = Workbooks(Nw)
	Workbooks(wn).Activate
	UpdateProgress
	Dim I As Integer
	For I = 1 To Sheets.Count
		Dim Sh
		For Each Sh In SheetsToCopy
			With Workbooks(wn).Sheets(I)
				If Sh.Name = .Name Then
					Dim T As XlSheetVisibility
					T = .Visible
					.Visible = xlSheetVisible
					.Copy After:=CopySheets.Sheets(CopySheets.Sheets.Count)
					.Visible = T
					CleanWorkbook CopySheets
				End If
			End With
		Next
		UpdateProgress
	Next
	If Not CopySheets.Sheets("Sheet1") Is Nothing Then CopySheets.Sheets("Sheet1").Delete
	CopySheets.Sheets(1).Activate
	Application.ScreenUpdating = True
End Function

Public Sub DeleteTableData(SheetsWithData, optional EntireRow as Boolean = false)
	Dim Sh, Tbl As ListObject
	TrackProgress UBound(SheetsWithData) - LBound(SheetsWithData) + 1
On Error Resume Next
	For Each Sh In SheetsWithData
		For Each Tbl In Sh.ListObjects
			if EntireRow Then
				Tbl.DataBodyRange.EntireRow.Delete xlUp
			Else
				Tbl.DataBodyRange.Delete xlUp
		Next
		UpdateProgress
	Next
On Error GoTo 0
End Sub
