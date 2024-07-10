'
'
'	Copyright 2021 - 2021 Sutherland Global Services
'
'	Project			: xvba-app		@ => d:\Karthik K Selvan\Testing\Checker
'
'	File			: modProgress.vba		@ => d:\Karthik K Selvan\Testing\Checker\Code\VBA\modProgress.vba
'	Created			: Thursday, 21st October 2021 5:22:07 pm
'
'	Author			: IN086482 (karthikk.selvan@sutherlandglobal.com)
'
'	Change History:
'
'
Option Explicit

Dim OpsCount As Integer, OpsCounter As Integer

Public Sub InitProgress(Count As Integer)
	OpsCount = Count
	OpsCounter = 0
	RefreshProgress
End Sub

Public Sub EndProgress()
	OpsCount = 0
	OpsCounter = 0
	With Parameter("Progress")
		.Value = 0
		.RowHeight = 0
	End With
End Sub

Public Sub UpdateProgress()
	OpsCounter = OpsCounter + 1
	RefreshProgress
End Sub

Public Sub TrackProgress(Count As Integer)
	OpsCount = OpsCount + Count
	RefreshProgress
End Sub

Public Sub RefreshProgress()
	Dim State As Boolean
	State = Application.ScreenUpdating
	If Not State Then Application.ScreenUpdating = TRUE
	With Parameter("Progress")
		.Value = IIf(OpsCount <> 0 And OpsCounter <> 0, OpsCounter / OpsCount, 0)
		.RowHeight = 20
	End With
	If Not State Then Application.ScreenUpdating = FALSE
End Sub
