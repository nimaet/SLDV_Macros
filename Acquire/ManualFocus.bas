' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' Assigns focus manually at all scan points
'#Uses "..\SwitchToAcquisitionMode.bas"
Option Explicit
Dim oScanHeadDevices As ScanHeadDevices
Dim oMeasPoints As MeasPoints
Dim pfX As Single
Dim pfY As Single
Dim posZ As Double
Dim dCoordX As Double
Dim dCoordY As Double
Dim dCoordZ As Double

Dim oFocus As Long
Dim oFocusArr(1) As Long

Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' ------------------------------------------------------------------------------
	If (Not SwitchToAcquisitionMode()) Then
		MsgBox "Switch to acquisition mode failed."
		Exit Sub
	End If

	Set oScanHeadDevices = Application.Acquisition.Infos.ScanHeadDevicesInfo.ScanHeadDevices
	Set oMeasPoints = Application.Acquisition.Infos.MeasPoints

	Dim oSensorHead As ISensorHead
	Set oSensorHead = Application.Acquisition.Infos.Vibrometers.Item(1).Controllers.Item(0).SensorHeads.Item(0)

	Dim oAlignment2D As Alignment2D
	Set oAlignment2D = Application.Acquisition.Infos.Alignments.Alignments2D(1)

	Dim oMeasPoint As MeasPoint
	Dim oScanHeadDevice As ScanHeadDevice
	Set oScanHeadDevice = oScanHeadDevices(1)


	For Each oMeasPoint In oMeasPoints
		oMeasPoint.VideoXY(pfX, pfY)
		oAlignment2D.VideoToScanner(pfX, pfY, dCoordX, dCoordY)
		oScanHeadDevice.ScanHeadControl.ScannerControl.SetBeamPosition(dCoordX, dCoordY)

		oSensorHead.StartAutoFocus()
		While oSensorHead.AutoFocusInProgress
			Wait(0.1)
		Wend
		oFocus = oSensorHead.FocusPosition
		oFocusArr(0) = oFocus
		Debug.Print oFocus
		Debug.Print oMeasPoint.FocusValues(0)

		oMeasPoint.FocusValues(0) = oFocus
		Wait(0.5)
	Next oMeasPoint



	'Acquisition.AssignFocusAutomatically(ptcAssignFocusAutomaticallyCapsSelected)



End Sub

