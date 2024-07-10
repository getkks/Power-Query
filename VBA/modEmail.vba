'
'
'	Copyright 2021 - 2021 Sutherland Global Services
'
'	Project			: xvba-app		@ => d:\Karthik K Selvan\Testing\Checker
'
'	File			: modEmail.vba		@ => d:\Karthik K Selvan\Testing\Checker\Code\VBA\modEmail.vba
'	Created			: Thursday, 21st October 2021 5:40:43 pm
'
'	Author			: IN086482 (karthikk.selvan@sutherlandglobal.com)
'
'	Change History:
'
'
Option Explicit

Const ERR_APP_NOTRUNNING As Long = 429

Const Accent1 = -738131969
Const Accent2 = -721354753
Const Accent3 = -704577537
Const Accent4 = -687800321
Const Accent5 = -671023105
Const Accent6 = -654245889

Public Sub Mail(Attach)
	Dim Outlook As Outlook.Application
	TrackProgress 1
	On Error Resume Next
	Set Outlook = GetObject(, "Outlook.Application")
	If Err = ERR_APP_NOTRUNNING Then
		'Set Outlook = CreateObject("Outlook.Application")
		MsgBox "Please (re)open Outlook To generate email.", vbCritical + vbOKOnly, ThisWorkbook.Name
	End If

	On Error GoTo 0

	Dim NewEmail As Outlook.MailItem
	Set NewEmail = Outlook.CreateItem(olMailItem)

	Dim Inspector As Inspector, Editor As Word.Document
	Set Editor = NewEmail.GetInspector.WordEditor

	NewEmail.Subject = Parameter("Subject")
	Editor.Activate

	With NewEmail
		Dim A
		TrackProgress LBound(Attach)
		For Each A In Attach
			.Attachments.Add A
			UpdateProgress
		Next
		.To = Parameter("To")
		.CC = Parameter("CC")
		.Recipients.ResolveAll
		.BCC = Parameter("BCC")
		.Display
	End With

	With Editor.Application.Selection

		.InsertAfter Parameter("Body")
		'PasteTable Editor, shData.ListObjects(1)
		.HomeKey wdStory
	End With

	Set Outlook = Nothing
	Set NewEmail = Nothing
	Set Editor = Nothing
	UpdateProgress

End Sub

Public Sub PasteTable(Editor As Word.Document, Table As ListObject)
	Table.Range.Copy
	With Editor.Application.Selection
		.MoveDown wdLine, 2, wdMove
		.PasteExcelTable False, False, FALSE
		AutoFitTables Editor
	End With
End Sub

Public Sub PastePivot(Editor As Word.Document, Pivot As PivotTable)
	Pivot.PivotSelect "", xlDataAndLabel, TRUE
	Selection.Copy
	With Editor.Application.Selection
		.MoveDown wdLine, 2, wdMove
		.PasteExcelTable False, False, FALSE
		AutoFitTables Editor
	End With
End Sub

Public Sub AutoFitTables(Editor As Word.Document)
	Dim Tbl As Word.Table
	TrackProgress Editor.Tables.Count
	For Each Tbl In Editor.Tables
		With Tbl
			With .Borders
				.InsideLineStyle = wdLineStyleSingle
				.OutsideLineStyle = wdLineStyleSingle
				.InsideLineWidth = wdLineWidth050pt
				.OutsideLineWidth = wdLineWidth050pt
				.InsideColor = Accent1
				.OutsideColor = Accent1
			End With
			.Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
			.Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
			.Borders.Shadow = FALSE
			.AutoFitBehavior wdAutoFitContent
			.TopPadding = InchesToPoints(0.05)
			.BottomPadding = InchesToPoints(0.05)
			.LeftPadding = InchesToPoints(0.08)
			.RightPadding = InchesToPoints(0.08)
		End With
		UpdateProgress
	Next
End Sub
