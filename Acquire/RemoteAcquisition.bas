' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This Macro displays a user dialog which allows the user to remote 
' control the acquisition. 
'
' Shows how to
' - use the acquisition object
' - get the acquisition point data (default for 1D vibrometer)
' - use the Acquisition state callback procedure
' - create a user dialog
'
' References
' - Polytec PSV\VibSoft Type Library
' - Polytec PhysicalUnit Type Library
' - Polytec PolyAlignment Type Library
' - Polytec PolyDigitalFilters Type Library
' - Polytec PolyFile Type Library
' - Polytec PolyFrontEnd Type Library
' - Polytec PolyGenerators Type Library
' - Polytec PolyInplane Type Library
' - Polytec PolyMath Type Library
' - Polytec PolyProperties Type Library
' - Polytec PolyScanHead Type Library
' - Polytec PolySignal Type Library
' - Polytec PolyWaveforms Type Library
' - Polytec Vibrometer Type Library
' - Polytec WindowFunction Type Library
' - Polytec SignalDescription Type Library
' - Polytec PolyAcquisition Type Library
' ----------------------------------------------------------------------

'#Uses "..\SwitchToAcquisitionMode.bas"

Option Explicit

Dim GetPointData As Boolean
Dim SingleShot As Boolean


Sub Main
' -------------------------------------------------------------------------------
' Main procedure.
' -------------------------------------------------------------------------------
	If Not SwitchToAcquisitionMode() Then
		MsgBox("Switch to acquisition mode failed.", vbOkOnly)
		End
	End If

	Debug.Clear

	GetPointData = False
	SingleShot = False


	' Define User dialog
	Begin Dialog UserDialog 300,224,"Acquisition (remote control)",.RemoteAcquisitionDlg ' %GRID:10,7,1,1
		GroupBox 10,161,280,28,"",.GroupBox1
		PushButton 10,70,130,28,"Scan All",.ScanAll
		PushButton 160,70,130,28,"Scan Continue",.ScanContinue
		PushButton 10,105,130,28,"Scan Remeasure",.ScanRemeasure
		PushButton 10,35,130,28,"Start Single",.StartSingle
		PushButton 160,35,130,28,"Start Continuous",.StartContinuous
		PushButton 160,105,130,28,"Stop",.StopAcquisition
		Text 20,7,260,14,"Version",.Version,2
		Text 20,170,260,14,"State",.State,2
		CheckBox 40,143,220,14,"Get acquisition point data",.CheckBoxGetData
		OKButton 90,196,130,21
	End Dialog

	' Create the dialog
	Dim AcquisitionDialog As UserDialog
	' Show the dialog
	Dialog AcquisitionDialog

	' After closing the dialog, stop any pending acquisition
	Acquisition.Stop

	MsgBox("Macro has finished.", vbOkOnly)
End Sub

Private Function RemoteAcquisitionDlg(DlgItem$, Action%, SuppValue&) As Boolean
' -------------------------------------------------------------------------------
' Dialog for remote acquisition.
' -------------------------------------------------------------------------------
	With Acquisition
		Select Case Action%
		Case 1 ' Dialog box initialisation
			DlgText "Version", Name+" "+Version
			If IsVibSoft() Then
				DlgVisible "ScanAll", False
				DlgVisible "ScanContinue", False
				DlgVisible "ScanRemeasure", False
			End If

			DlgValue "CheckBoxGetData", GetPointData

			Call AcqStateChanged(.State)
		Case 2 ' Values changed or buttons clicked
			RemoteAcquisitionDlg = True
			If DlgItem$ = "StartSingle" Then
				.Start ptcAcqStartSingle
			ElseIf DlgItem$ = "StartContinuous" Then
				.Start ptcAcqStartContinuous
			ElseIf DlgItem$ = "StopAcquisition" Then
				.Stop
			ElseIf DlgItem$ = "ScanAll" Then
				.Scan(ptcScanAll)
			ElseIf DlgItem$ = "ScanContinue" Then
				On Error Resume Next
				.Scan(ptcScanContinue)
			ElseIf DlgItem$ = "ScanRemeasure" Then
				On Error Resume Next
				.Scan(ptcScanRemeasure)
			ElseIf DlgItem$ = "CheckBoxGetData" Then
				GetPointData = Not GetPointData
			Else
				RemoteAcquisitionDlg = False
			End If
		Case 3 ' Text box or combo box changed
		Case 4 ' Focus changed
		Case 5 ' Idle
			If .State <> ptcAcqStateStopped Then
				Wait 0.5
			End If
			RemoteAcquisitionDlg = True
		Case 6 ' Function key
		Case Else
			RemoteAcquisitionDlg = False
		End Select
	End With
End Function

Public Sub AcqStateChanged(ByVal AcqState As PTCAcqState)
' -------------------------------------------------------------------------------
' Acquisition state changed. Set buttons.
' This is a callback function from PSV. Must be Public, do not change it.
' -------------------------------------------------------------------------------
	Dim bStart As Boolean
	' Update ScanState string in dialog
	DlgText "State", AcqStateStr(AcqState)
	'
	bStart = False
	If AcqState = ptcAcqStateStopped Then
		bStart = True
		SingleShot = False
	End If
	' Enable/disable dialog buttons
	DlgEnable "StartSingle", bStart
	DlgEnable "StartContinuous", bStart
	DlgEnable "ScanAll", bStart
	DlgEnable "ScanContinue", bStart
	DlgEnable "ScanRemeasure", bStart
	DlgEnable "StopAcquisition", Not bStart
	DlgEnable "OK", bStart

	If AcqState = ptcAcqStateSingle Then
		SingleShot = True
	End If
End Sub

Public Sub ScanStateChanged(ByVal ScanState As PTCScanState, ByVal ScanPoint As Long)
' -------------------------------------------------------------------------------
' Scan state changed.
' Get the point data for some signals to demonstrate the access to the acquisition point data,
' if data acquisition for current scan point has finished.
' -------------------------------------------------------------------------------

	Select Case ScanState
	Case ptcScanStateEndScanPoint
		If GetPointData And SingleShot Then
			Dim ChannelString As String

			If Application.Acquisition.Infos.Hardware.ActiveFrontEnd.Caps And ptcFrontEndCaps3D Then
				ChannelString = "Vib X"
			Else
				ChannelString = "Vib"
			End If

			' In the following lines please input the data corresponding with the PSV settings
			Call GetData(ptcDomainTime, ChannelString, "Velocity", ptcDisplaySamples, ScanPoint)
			Call GetData(ptcDomainSpectrum, ChannelString, "Velocity", ptcDisplayMag, ScanPoint)
		End If
	End Select

End Sub

Private Function AcqStateStr(AcqState As PTCAcqState) As String
' -------------------------------------------------------------------------------
' Get acquisition string.
' -------------------------------------------------------------------------------
	Select Case AcqState
	Case ptcAcqStateSingle
		AcqStateStr = "Single"
	Case ptcAcqStateContinuous
		AcqStateStr = "Continuous"
	Case ptcAcqStateStopped
		AcqStateStr = "Stopped"
	Case ptcAcqStateScanAll
		AcqStateStr = "Scan All"
	Case ptcAcqStateScanContinue
		AcqStateStr = "Scan Continue"
	Case ptcAcqStateScanRemeasure
		AcqStateStr = "Scan Remeasure"
	End Select
End Function

Private Function IsVibSoft() As Boolean
' -------------------------------------------------------------------------------
' Check we have Vibsoft running.
' -------------------------------------------------------------------------------
	IsVibSoft = Application.Mode = ptcApplicationModeNormal
End Function

Private Sub GetData(DomainType As PTCDomainType, ChannelString As String, SignalString As String, DisplayType As PTCDisplayType, ScanPoint As Long)
' -------------------------------------------------------------------------------
' Get the acquisition point data for a specific domain, channel, signal and display for the current scan point index.
' Trace the data to the debug output window.
' -------------------------------------------------------------------------------

	Dim oAcquisition As Acquisition
	Set oAcquisition = Application.Acquisition

	Dim oDisplay As Display
	Set oDisplay = oAcquisition.PointDomains.type(DomainType).Channels.Item(ChannelString).Signals.Item(SignalString).Displays.type(DisplayType)

	' Get the acquisition data object used for the current data acquisition.
	' Release this acquisition data instance before new data acquisition will be started.
	Dim oAcquisitionPointData As AcquisitionPointData
	Set oAcquisitionPointData = oAcquisition.GetData(oDisplay)

	Dim oAcquisitionPointDataStream As AcquisitionPointDataStream
	Set oAcquisitionPointDataStream = oAcquisitionPointData.OpenDataStream(oAcquisitionPointData.BlockCount, 0)

	Debug.Print "ScanPoint Index: ";ScanPoint
	Debug.Print "Domain: ";oDisplay.Signal.Channel.Domain.Name
	Debug.Print "Channel: ";ChannelString
	Debug.Print "Signal: ";SignalString
	Debug.Print "Display: ";oDisplay.Name

	Dim oXAxis As AcquisitionXAxis
	Set oXAxis = oAcquisitionPointData.XAxis

	Dim oYAxes As AcquisitionYAxes
	Set oYAxes = oAcquisitionPointData.YAxes

	Dim AxesString As String
	AxesString = oXAxis.Name + " / " + oXAxis.Unit

	If (oAcquisitionPointDataStream.Stride <> oYAxes.Count) Then
		MsgBox "oAcquisitionPointDataStream.Stride should be equal to oYAxes.Count."
	End If

	Dim oYAxis As AcquisitionYAxis

	For Each oYAxis In oYAxes
		AxesString = AxesString + "    " + oYAxis.Name + " / " + oYAxis.Unit
	Next oYAxis

	Debug.Print AxesString

	Dim Stride As Long
	Stride = oAcquisitionPointDataStream.Stride

	While (oAcquisitionPointDataStream.Position < oAcquisitionPointDataStream.Length)

		Dim XIndex As Long
		XIndex = oAcquisitionPointDataStream.Position

		Dim Data() As Single
		oAcquisitionPointDataStream.Read(0, -1, Data)

		Dim LineString As String

		Dim SampleIndex As Long
		SampleIndex = 0

		Dim DataIndex, StrideIndex As Long

		For DataIndex = LBound(Data) To UBound(Data) Step Stride

			LineString = CStr(oXAxis.GetMidX(XIndex+SampleIndex))

			For StrideIndex = 0 To Stride-1
				LineString = LineString + "    " + CStr(Data(DataIndex+StrideIndex))
			Next StrideIndex

			Debug.Print LineString

			SampleIndex = SampleIndex + 1

		Next DataIndex
	Wend

	Debug.Print ""

	oAcquisitionPointDataStream.Close

End Sub
