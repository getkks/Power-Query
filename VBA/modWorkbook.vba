'
'
'	Copyright 2021 - 2021 Sutherland Global Services
'
'	Project			: xvba-app		@ => d:\Karthik K Selvan\Testing\Checker
'
'	File			: modWorkbook.vba		@ => d:\Karthik K Selvan\Testing\Checker\Code\VBA\modWorkbook.vba
'	Created			: Thursday, 21st October 2021 5:48:00 pm
'
'	Author			: IN086482 (karthikk.selvan@sutherlandglobal.com)
'
'	Change History:
'
'
Option Explicit

Public Sub CleanWorkbook(Book As Workbook, Optional ConditionalFormat As Boolean = False)
    Dim qry As WorkbookQuery, Con As WorkbookConnection, Shp As Excel.Shape, Sh As Worksheet
    TrackProgress Book.Queries.Count
    For Each qry In Book.Queries
        qry.Delete
        UpdateProgress
    Next
    TrackProgress Book.Connections.Count
    For Each Con In Book.Connections
        Con.Delete
        UpdateProgress
    Next
    For Each Sh In Book.Worksheets
        TrackProgress Sh.Shapes.Count
        For Each Shp In Sh.Shapes
            Shp.Delete
            UpdateProgress
        Next
        If ConditionalFormat Then Sh.Cells.FormatConditions.Delete
    Next
End Sub
