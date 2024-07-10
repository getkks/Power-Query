'
'
'	Copyright 2021 - 2021 Sutherland Global Services
'
'	Project			: xvba-app		@ => d:\Karthik K Selvan\Testing\Checker
'
'	File			: modPivot.vba		@ => d:\Karthik K Selvan\Testing\Checker\Code\VBA\modPivot.vba
'	Created			: Thursday, 21st October 2021 5:47:30 pm
'
'	Author			: IN086482 (karthikk.selvan@sutherlandglobal.com)
'
'	Change History:
'
'
Option Explicit

Public Sub ChangeCache(Book As Workbook, Pivot As PivotTable, Source As Range)
    Pivot.ChangePivotCache Book.PivotCaches.Create(xlDatabase, Source)
End Sub

Public Sub ChangeSheetPivotCache(Book As Workbook, SheetName As String, Source As ListObject)
    Dim Pivot As PivotTable
    TrackProgress Book.Sheets(SheetName).PivotTables.Count
    For Each Pivot In Book.Sheets(SheetName).PivotTables
        Pivot.ChangePivotCache Book.PivotCaches.Create(xlDatabase, Source.Name)
        'ChangeCache Book, Pivot, Source.Range
        UpdateProgress
    Next
End Sub
