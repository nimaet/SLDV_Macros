' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This macro shows the use of user defined bands.
'
' The macro fits the peak within the given cursor range at each scan point and generates
' a user defined band with the amplitude and phase at the respective peak frequencies.
'
' How to use this macro:
' - start PSV and open the file
' - select single scan point or average spectrum and the vibrometer channel
' - define an area within spectrum with the difference or band cursor
' - start the macro
'
' References
' - Polytec PhysicalUnit Type Library
' - Polytec PolyAlignment Type Library
' - Polytec PolyFile Type Library
' - Polytec PolyFrontEnd Type Library
' - Polytec PolyMath Type Library
' - Polytec PolyProperties Type Library
' - Polytec PolySignal Type Library
' - Polytec WindowFunction Type Library
' - Polytec SignalDescription Type Library
' ----------------------------------------------------------------------

Option Explicit


Sub Main

On Error GoTo Error_Handler

	If (Application.Documents.Count = 0) Then
		MsgBox("No document open!")
		Exit All
	End If

	Dim oDoc As VibDocument
	Set oDoc = Application.ActiveDocument
	Dim  strName As String
	strName = oDoc.Name

	' get the area window from the active document
	Dim bWndFound As Boolean
	bWndFound = False

	Dim oAreaWindow As AreaWindow
	Dim oWnd As Window
	For Each oWnd In oDoc.Windows
		If (oWnd.type = ptcWindowTypeArea) Then
			bWndFound = True
			Set oAreaWindow = oWnd
			Exit For
		End If
	Next

	If (bWndFound = False) Then
		MsgBox("No area window!")
		Exit All
	End If

	' check if view is single scan point or average spectrum
	Select Case oAreaWindow.AreaView.Settings.ViewMode
	Case ptcViewModeSinglePoint
	Case ptcViewModeAverage
	Case ptcViewModePointSinglePoint
	Case ptcViewModePointAverage
	Case Else
		MsgBox("Please set the view to single scan point or average spectrum.")
		Exit All
	End Select

	' get the channel name
	Dim oChannel3 As IChannel2
	Set oChannel3 = oAreaWindow.AreaView.Display.Signal.Channel
	Dim strChannelName As String
	strChannelName = oChannel3.Name

	Dim pointBuildFlags As PTCSignalBuildFlags
	Dim bandBuildFlags As PTCSignalBuildFlags
	pointBuildFlags = ptcBuildPointDataXYZ
	bandBuildFlags = ptcBuildBandDataXYZ

	Dim is3d As Boolean

	If (oChannel3.Caps And ptcChannelCapsVector) Then
		is3d = True
		pointBuildFlags = ptcBuildPointData3d
		bandBuildFlags = ptcBuildBandData3d
	End If

	If (oChannel3.Caps And ptcChannelCapsUser) Then
		MsgBox("This macro doesn't work for user defined signals.")
		Exit All
	End If

	' get the signal name
	Dim strSignalName As String
	strSignalName = oAreaWindow.AreaView.Display.Signal.Name

	' check if X-axis is Hz
	If (oAreaWindow.AnalyzerView.XAxis.Unit <> "Hz") Then
		MsgBox("X-axis value must be a frequency.")
		Exit All
	End If

	' get the left and right cursor positions in Hz
	Dim dStartFrequency As Double
	Dim dEndFrequency As Double

	If (oAreaWindow.AnalyzerView.Cursor.CursorType = ptcCursorTypePeak) Then
		Dim oPeakCursorSettings As PeakCursorSettings
		Set oPeakCursorSettings = oAreaWindow.AnalyzerView.Settings.CursorSettings

		dStartFrequency = oPeakCursorSettings.LeftPosition
		dEndFrequency = oPeakCursorSettings.RightPosition
	ElseIf (oAreaWindow.AnalyzerView.Cursor.CursorType = ptcCursorTypeDouble) Then
		Dim oDoubleCursorSettings As DoubleCursorSettings
		Set oDoubleCursorSettings = oAreaWindow.AnalyzerView.Settings.CursorSettings

		dStartFrequency = oDoubleCursorSettings.LeftPosition
		dEndFrequency = oDoubleCursorSettings.RightPosition
	Else
		MsgBox("Wrong cursor type. Use Differential Cursor or Band Cursor.")
		Exit All
	End If

	Dim analyzerChannelName As String
	analyzerChannelName = oAreaWindow.AnalyzerView.Display.Signal.Channel.Name

	Dim analyzerSignalName As String
	analyzerSignalName = oAreaWindow.AnalyzerView.Display.Signal.Name


	If (MsgBox("This macro will modify the file '" + strName + _
		"'. We strongly recommend to work with a backup copy of original data only. " + _
		"Do you want to continue?", vbYesNo) = vbNo) Then
		Exit Sub
	End If

	Dim strDisplayName As String
	strDisplayName = oAreaWindow.AnalyzerView.Display.Name

	Dim lFirstIndex As Long
	Dim lLastIndex As Long

	' get the start and end frequency from the analyzer window
	lFirstIndex = oAreaWindow.AnalyzerView.XAxis.GetIndex(dStartFrequency)
	lLastIndex  = oAreaWindow.AnalyzerView.XAxis.GetIndex(dEndFrequency)

	' save and close active document
	oDoc.Save()
	oDoc.Close()

	' ----------------------------------------------------------------------
	' open file with PolyFile
	' ----------------------------------------------------------------------
	Dim oFile As New PolyFile
	oFile.ReadOnly = False
	oFile.Open(strName)

	Dim oPointDomain As PointDomain
	Dim oPointSignal As Signal

	If Not oFile.GetPointDomains.Exists(ptcDomainSpectrum) Then
		MsgBox("Missing domain for point data. Can not evaluate the file.")
		Exit All
	End If

	Set oPointDomain = oFile.GetPointDomains(pointBuildFlags).type(ptcDomainSpectrum)
	Set oPointSignal = oPointDomain.Channels.Item(strChannelName).Signals(strSignalName)

	If (analyzerSignalName <> strSignalName Or analyzerChannelName <> strChannelName) Then
		If (oPointSignal.Description.XAxis.Min > dStartFrequency Or oPointSignal.Description.XAxis.Max < dEndFrequency) Then
			MsgBox("At least one of the selected frequencies is out of X-Axis range. Please select the same signal in both Area View and Analyzer View.")
			Exit All
		End If
	End If

	Set oPointDomain = oFile.GetPointDomains(pointBuildFlags).type(ptcDomainSpectrum)
	Set oPointSignal = oPointDomain.Channels.Item(strChannelName).Signals(strSignalName)

	Dim domainType As PTCDomainType
	domainType = ptcDomainNotAvail

	' get the number of frames
	Dim oPointDisplay As Display
	Set oPointDisplay = oPointSignal.Displays(strDisplayName)

	Dim lFrameCount As Long
	lFrameCount = oPointDomain.DataPoints.Item(1).GetFrames(oPointDisplay)

	Dim bFirstFrame As Boolean
	bFirstFrame = True

	Dim bMultiFrameMode As Boolean
	bMultiFrameMode = oFile.Infos.AcquisitionInfoModes.ActiveMode() = ptcAcqModeMultiFrame

	' Check for MIMO type
	If Not bMultiFrameMode Then
		If oPointSignal.Description.FunctionType = ptcFunctionPCAPrincipalInputsType Or oPointSignal.Description.FunctionType = ptcFunctionPCAVirtualCoherencesType Then
			MsgBox("This macro doesn't work for MIMO signals.")
			GoTo Error_Handler
		End If
	End If

	Dim bComplex As Boolean
	bComplex = oPointDisplay.Signal.Description.Complex

	Dim arrScanStatus() As Long
	arrScanStatus = oPointDomain.DataPoints.GetScanStatus(oPointSignal)

	Dim dPeakFrequency As Double

	Dim lComplexFactor As Long
	lComplexFactor = 1
	If (bComplex) Then
		lComplexFactor = 2
	End If

	Dim l3dFactor As Long
	l3dFactor = 1
	If (oPointSignal.Channel.Caps And ptcChannelCapsVector) Then
		l3dFactor = 3
	End If

	Dim lFactor As Long
	lFactor = lComplexFactor * l3dFactor

	Dim lBandDataLength As Long
	lBandDataLength = oPointDomain.DataPoints.Count * lComplexFactor

	Dim pointWithPeakoutsideSelection As Long
	pointWithPeakoutsideSelection = 0

	Dim lFrame As Long
	For lFrame = 1 To lFrameCount

		Dim arrBandData() As Single
		ReDim arrBandData(0 To lFactor * oPointDomain.DataPoints.Count - 1)

		Dim lIndex As Long
		lIndex = 0

		Dim lIndexIncr As Long
		If (bComplex) Then
			lIndexIncr = 2
		Else
			lIndexIncr = 1
		End If

		' ----------------------------------------------------------------------
		' loop over all scan points to get the peak frequency of each scan point
		' ----------------------------------------------------------------------
		Dim oDataPoint As DataPoint
		Dim dataPointIndex As Long

		For dataPointIndex = 1 To oPointDomain.DataPoints.Count
			Set oDataPoint = oPointDomain.DataPoints(dataPointIndex)
			Dim ptcStatus As PTCScanStatus
			ptcStatus = oDataPoint.GetScanStatus(oPointDisplay)

			' check if the data point is valid (i.e. has data)
			If (ptcStatus Or ptcScanStatusValid) Then

				Dim oDisplay As Display

				If (bComplex) Then
					Set oDisplay = oPointSignal.Displays.type(ptcDisplayRealImag)
				Else
					Set oDisplay = oPointSignal.Displays.type(ptcDisplayMag)
				End If

				Dim Data() As Single
				If (bMultiFrameMode = True) Then
					Data = oDataPoint.GetDataSection(oDisplay, lFrame, 1, lFirstIndex, lLastIndex)
				Else
					Data = oDataPoint.GetDataSection(oDisplay, 0, 1, lFirstIndex, lLastIndex)
				End If

				Dim lDataCount As Long
				lDataCount = (UBound(Data) - LBound(Data))/lFactor

				Dim lDataLength As Long
				lDataLength = lDataCount * lComplexFactor

				Dim Data3d() As Single
				ReDim Data3d(0 To lDataLength - 1)
				Dim lIndexLoop As Integer, l3dLoop As Integer, lComplexLoop As Integer

				lIndexLoop = 0
				l3dLoop = 0
				lComplexLoop = 0

				Dim oPeakFit As PeakFit
				For l3dLoop = 0 To l3dFactor - 1
					For lIndexLoop = 0 To lDataCount - 1
						For lComplexLoop = 0 To lComplexFactor - 1
							Data3d(lIndexLoop * lComplexFactor + lComplexLoop) = Data(l3dLoop * lDataLength + lIndexLoop * lComplexFactor + lComplexLoop)
						Next
					Next

					Set oPeakFit = New PeakFit
					oPeakFit.Fit(dStartFrequency, dEndFrequency, Data3d, bComplex, False)

					Dim dAmplitude As Double

					dAmplitude = oPeakFit.GetPeak(False, dPeakFrequency)

					If (dPeakFrequency < dStartFrequency Or dPeakFrequency > dEndFrequency) Then ' set band status to invalid if peak frequency is out of defined frequency range
						arrScanStatus(dataPointIndex - 1) = ptcScanStatusInvalidated
						If (pointWithPeakoutsideSelection = 0) Then
							pointWithPeakoutsideSelection = oDataPoint.MeasPoint.Label
						End If
					End If

					If (bComplex) Then
						Dim daAmplitude() As Double
						Dim daPeakFrequency(1) As Double
						daPeakFrequency(0) = dPeakFrequency

						daAmplitude = oPeakFit.GetFunctionXY(daPeakFrequency)

						arrBandData(l3dLoop*lBandDataLength + lIndex) 	= daAmplitude(0)
						arrBandData(l3dLoop*lBandDataLength + lIndex + 1) = daAmplitude(1)
					Else
						arrBandData(lIndex) = dAmplitude
					End If

				Next

				Set oPeakFit = Nothing

			End If

			lIndex = lIndex + lIndexIncr

		Next

		' ----------------------------------------------------------------------
		' Add new user band.
		' ----------------------------------------------------------------------
		If (bFirstFrame = True) Then

			Dim oSignal As Signal
			Set oSignal = oFile.GetBandDomains(bandBuildFlags).type(ptcDomainRMS).Channels.Item(strChannelName).Signals(strSignalName)

			Dim oBandDefUsr As New BandDef

			oBandDefUsr.Peak  = dPeakFrequency 		' use the peak frequency from last scan point
			oBandDefUsr.Start = dStartFrequency
			oBandDefUsr.Stop  = dEndFrequency

			Dim oSignalDescUsr As SignalDescription
			Set oSignalDescUsr = oSignal.Description.Clone()

			oSignalDescUsr.Name = "Peak Band " + oSignal.Channel.Name + " " + oSignal.Name + " " + CStr(oBandDefUsr.Peak) + " Hz"

			oSignalDescUsr.Complex = bComplex
			oSignalDescUsr.DomainType = ptcDomainSpectrum

			Dim oDataBand As DataBand
			Set oDataBand = oFile.GetBandDomains.FindBand(oSignalDescUsr, True, oBandDefUsr, oSignal)

			If (Not oDataBand Is Nothing) Then
				If (MsgBox("This user defined band already exists and will be replaced with the new band. " + _
					"Do you want to continue?", vbYesNo) = vbNo) Then
					GoTo Error_Handler
				End If

				oFile.GetBandDomains.RemoveBand(oSignal, oBandDefUsr)
			End If

			Set oDataBand = oFile.GetBandDomains.AddBand(oSignalDescUsr, oBandDefUsr, oSignal) ' add the new user defined band

			bFirstFrame = False

			strChannelName = "Usr"
			If (is3d) Then
				strChannelName = "Usr 3D"
			End If
			strSignalName = oSignalDescUsr.Name
			domainType = oSignal.Channel.Domain.type
		End If

		oDataBand.SetData(oSignal, lFrame, arrBandData)

	Next lFrame

	oDataBand.SetScanStatus(oSignal, arrScanStatus)

	oFile.Save()

Error_Handler:

	If Not (oFile Is Nothing) And oFile.IsOpen Then
		oFile.Close()
	End If

	If (Err.Number <> 0) Then
		MsgBox Err.Description
	Else
		MsgBox "The peak for at least one point (e.g. with index " + Str(pointWithPeakoutsideSelection) + ") lies outside your selected frequency range. The status for this point in the generated signal has been set to invalidated.", vbInformation
	End If

	If Len(strName) > 0 Then
		Application.Documents.Open(strName) ' reopen the file

		For Each oWnd In Application.ActiveDocument.Windows
			If (oWnd.type = ptcWindowTypeArea And Len(strChannelName) > 0 And Len(strSignalName)>0) Then
				Set oAreaWindow = oWnd
				If (domainType <> ptcDomainNotAvail) Then
					oAreaWindow.AreaView.Settings.DisplaySettings.Domain = domainType
				End If
				oAreaWindow.AreaView.Settings.DisplaySettings.Channel = strChannelName
				oAreaWindow.AreaView.Settings.DisplaySettings.Channel = strSignalName
				Exit For
			End If
		Next
	End If

End Sub
