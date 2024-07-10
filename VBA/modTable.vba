'
'
'	Copyright 2021 - 2021 Sutherland Global Services
'
'	Project			: xvba-app		@ => d:\Karthik K Selvan\Testing\Checker
'
'	File			: modTable.vba		@ => d:\Karthik K Selvan\Testing\Checker\Code\VBA\modTable.vba
'	Created			: Thursday, 21st October 2021 5:45:24 pm
'
'	Author			: IN086482 (karthikk.selvan@sutherlandglobal.com)
'
'	Change History:
'
'
Option Explicit

Public Sub RefreshTables(Tables As Collection)
    Dim Tbl
    TrackProgress Tables.Count 'UBound(Tables) - LBound(Tables) + 1
    For Each Tbl In Tables
        Debug.Print "Table: " & Tbl.Name & " Refreshing..."
        Tbl.Refresh
        With Tbl
            .TableStyle = "TableStyleLight9"
            .ShowTableStyleFirstColumn = False
            .ShowTableStyleColumnStripes = False
            .ShowTableStyleLastColumn = False
            .ShowTableStyleRowStripes = False
            .ShowAutoFilter = False
        End With
        UpdateProgress
    Next
End Sub
