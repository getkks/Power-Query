'
'
'	Copyright 2021 - 2021 Sutherland Global Services
'
'	Project			: xvba-app		@ => d:\Karthik K Selvan\Testing\Checker
'
'	File			: modFormat.vba		@ => d:\Karthik K Selvan\Testing\Checker\Code\VBA\modFormat.vba
'	Created			: Thursday, 21st October 2021 5:42:07 pm
'
'	Author			: IN086482 (karthikk.selvan@sutherlandglobal.com)
'
'	Change History:
'
'
Option Explicit

Public Sub FormatColumns(Sheet As Worksheet, TableColumns, Format As String)
    FormatTableColumns Sheet, Sheet.ListObjects(1), TableColumns, Format
End Sub

Public Sub FormatTableColumns(Sheet As Worksheet, Table As ListObject, TableColumns, Format As String)
    'Sheet.Activate
    Dim Col
    TrackProgress UBound(TableColumns) - LBound(TableColumns) + 1
    For Each Col In TableColumns
        Sheet.Range(Table.Name & "[" & Col & "]").NumberFormat = Format
        UpdateProgress
    Next
End Sub
