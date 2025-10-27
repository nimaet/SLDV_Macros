' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This Macro moves the scanner continuously around the scan field.
' The Macro works for PSV and PSV 3D systems.
'
' This macro can only be executed within the PSV software because
' it requires hardware access.
'
' References
' - Polytec PhysicalUnit Type Library
' - Polytec PolyAlignment Type Library
' - Polytec PolyDigitalFilters Type Library
' - Polytec PolyFile Type Library
' - Polytec PolyFrontEnd Type Library
' - Polytec PolyGenerators Type Library
' - Polytec PolyMath Type Library
' - Polytec PolyProperties Type Library
' - Polytec PolyScanHead Type Library
' - Polytec PolySignal Type Library
' - Polytec Vibrometer Type Library
' - Polytec PolyWaveforms Type Library
' - Polytec WindowFunction Type Library
' - Polytec SignalDescription Type Library
' - Polytec PSV Type Library
'-----------------------------------------------------------------------

'#Uses "..\SwitchToAcquisitionMode.bas"

Option Explicit

Dim oScanHeadDevices As ScanHeadDevices

Dim dMax As Double
Dim bZeroPosition As Boolean
Dim dStep As Double

Dim dX As Double, dY As Double
Dim dMovingTimeSpeedUpFactor As Double


Sub Main

	' switch to acquisition mode
	If (Not SwitchToAcquisitionMode()) Then
		MsgBox "Switch to acquisition mode failed."
		Exit Sub
	End If

	Set oScanHeadDevices = Application.Acquisition.Infos.ScanHeadDevicesInfo.ScanHeadDevices ' ScanHeadDevices gives use a clone copy of the scanheaddevices, call them only once a time

	Begin Dialog UserDialog 320,140 ' %GRID:10,7,1,1
		Text 20,42,90,14,"Max. Angle:",.Text1
		TextBox 160,42,60,21,.Angle
		OKButton 210,112,90,21
		Text 20,14,270,14,"Please enter the angle for the scan range",.Text2
		Text 230,42,60,14,"°",.Text3
		CheckBox 20,77,240,14,"Include Zero Position in Path",.ZeroPosition
	End Dialog
	Dim dlg As UserDialog

	dlg.Angle = GetSetting("ShowScanRange", "Setup", "MaxAngle", "20")
	bZeroPosition = (GetSetting("ShowScanRange", "Setup", "ZeroPosition", "True") = "True")
	If bZeroPosition Then
		dlg.ZeroPosition = 1
	Else
		dlg.ZeroPosition = 0
	End If

	Dialog dlg

	dMax  = CDbl(dlg.Angle)
	bZeroPosition = (dlg.ZeroPosition = 1)

	SaveSetting("ShowScanRange", "Setup", "MaxAngle", dMax)
	Dim strZeroPosition As String
	If bZeroPosition Then
		strZeroPosition = "True"
	Else
		strZeroPosition = "False"
	End If
	SaveSetting("ShowScanRange", "Setup", "ZeroPosition", strZeroPosition)

	Dim oFrontEnd As FrontEnd
	Set oFrontEnd = Application.Acquisition.Infos.Hardware.ActiveFrontEnd

	Dim bFrontEndDac As Boolean
	bFrontEndDac = False

	If oFrontEnd.Caps And ptcFrontEndCapsControllable Then
		bFrontEndDac = oFrontEnd.Control.Caps And ptcFrontEndControlCapsDac
	End If

	' We want to move the scanner with about 400 ms per cycle:
	If bFrontEndDac Then
		dMovingTimeSpeedUpFactor = 50 * oScanHeadDevices.Count
		dStep = 0.5 * dMax  ' We use 10° steps. ScanHeadDevice.LaserControl.SetBeamPositionEx uses a cosine for moving the mirrors.
						    ' So we can't damage the mirrors.
	Else
		dMovingTimeSpeedUpFactor = 2.5 * oScanHeadDevices.Count
		dStep = 0.25 * dMax ' We use 5° steps. ScanHeadDevice.LaserControl.SetBeamPositionEx uses a cosine for moving the mirrors.
						 	' So we can't damage the mirrors.
	End If

	dX = -dMax
	dY =  dMax

	Call SetPosition(dX, dY) ' set start position

	While True
		Call MoveDacsX(-dMax,  dMax)
		Call MoveDacsY( dMax, -dMax)
		Call MoveDacsX( dMax, -dMax)
		Call MoveDacsY(-dMax,  dMax)

		If bZeroPosition Then
			Call SetPosition(0.0, 0.0)
			Wait 0.2
		End If
	Wend

End Sub


Sub MoveDacsX(dStartX As Double, dEndX As Double)
	Dim dStepX As Double
	dStepX = dStep

	If dEndX < dStartX Then
		dStepX = -dStepX
	End If

	For dX = dStartX To dEndX Step dStepX
		Call SetPosition(dX, dY)
    Next
    dX = dEndX
End Sub


Sub MoveDacsY(dStartY As Double, dEndY As Double)
	Dim dStepY As Double
	dStepY = dStep

	If dEndY < dStartY Then
		dStepY = -dStepY
	End If

	For dY = dStartY To dEndY Step dStepY
		Call SetPosition(dX, dY)
    Next
    dY = dEndY
End Sub


Sub SetPosition(dX As Double, dY As Double)
    Dim oScanHeadDevice As ScanHeadDevice
    For Each oScanHeadDevice In oScanHeadDevices
    	oScanHeadDevice.ScanHeadControl.ScannerControl.SetBeamPositionEx(dX, dY, dMovingTimeSpeedUpFactor)
    Next oScanHeadDevice
End Sub
