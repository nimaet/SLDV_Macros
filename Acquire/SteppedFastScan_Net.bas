' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This macro is an example for a stepped fast scan measurement.
' you can set the start and end frequency, the frequency step,
' the bandwidth and the filename for the resulting combined file.
' Note that this allowed only equidistant frequency steps.
'
' When you start the macro a series of fast scans is done.
' The data of the fast scans are combined in a single file as user
' defined data sets. For every combination of domain/channel/signal/
' display in the original fast scan data you will find corresponding
' data in the user defined data set (channel 'Usr') of the resulting
' combined file.
' The data is stored at every measurement point as a function of the
' frequency.
'
' E.g., if you selected start = 1000 Hz, end = 2000 Hz, step = 100 Hz
' you will get 11 frequency lines at these frequencies. You can select
' to display e.g. FRF's by selecting the 'FFT Vib & Ref1 FRF Velocity /
' Voltage' signal in the 'Usr' channel. You will be able to set a cursor
' at a frequency line and to display the area data at this line for all
' scan points. The displayed area data corresponds to the area data of
' the original fast scan data.
'
' Only this combined fast scan file will be stored, the other files
' will be created only temporarily. This combined file will contain
' the original data of the first fast scan plus the combined data of
' all fast scans.
'
' To display the calculated data in PSV do the following:
' - start PSV and open the scan file
' - select Presentation/View/SingleScanPoint
' - in analyzer window toolbar select Channel/Usr
'
' References
' - Polytec PSV Type Library
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
' ----------------------------------------------------------------------
'#Language "WWB.NET"

Option Explicit

Imports System
Imports System.Collections.Generic

Const c_strFileFilter As String = "Scan File (*.svd)|*.svd|All Files (*.*)|*.*||"
Const c_strFileExt As String = "svd"

Enum PTCAveragingMode
	ptcAveragingModeOff        = 0
	ptcAveragingModeMagnitude  = 1
	ptcAveragingModeComplex	   = 2
End Enum

Dim m_sFileName As String
Dim m_sFilePath As String
Dim m_dFreqStart As Double
Dim m_dFreqEnd As Double
Dim m_dFreqStep As Double
Dim m_dBandWidth As Double
Dim m_lAveraging As Long
Dim m_ptcAveragingMode As PTCAveragingMode


Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	On Error GoTo MErrorHandler
	Dim bRe As Boolean
	Dim j As Integer
	Dim k As Integer

	MsgBox("Notice:" + vbCrLf + vbCrLf + _
		"After measurement the macro needs some time to calculate the data." + vbCrLf + _
		"During this time you can not see any changes." + vbCrLf + vbCrLf + _
		"The macro informs you after it has finished.", vbOkOnly)

	' Check we are in the acquisition mode.
	bRe = SwitchToAcquisitionMode()
	If Not bRe Then
		MsgBox("The macro exits now.", vbOkOnly)
		Exit Sub
	End If

	' Dialog for user input.
	bRe = FastScanDialog()
	If Not bRe Then
		MsgBox("The macro exits now.", vbOkOnly)
		Exit Sub
	End If

	' Make an array for base file and temporary files.
	Dim scanFileNames As New List(Of String)
	scanFileNames = FastScanFiles()
	Dim scanFileCount As Integer
	scanFileCount = scanFileNames.Count

	' Now execute FastScan.
	bRe =  FastScanMeasurement(scanFileNames)
	If Not bRe Then
		MsgBox("Error measurement! The macro exits now.", vbOkOnly)
		Exit Sub
	End If

	' Switch to the presentation mode.
	Dim oPresApp As Application
	oPresApp = GetPresentationInstance()
	If (oPresApp Is Nothing) Then
		MsgBox("The macro exits now.", vbOkOnly)
	End If

	' close all windows
	oPresApp.Windows.CloseAll()

	' Now make calculation.
	' ---------------------

	' We have to open the file for read/write, otherwise we cannot save our
	' user defined dataset to the file.
	Dim oFiles As New Dictionary(Of String, PolyFile)
	' Open the base file first.
	If Not OpenFile(oFiles, scanFileNames(0)) Then
		MsgBox("Error open file! The macro exits now.", vbOkOnly)
		Exit Sub
	End If

	Dim oAcqInfoModes As AcquisitionInfoModes
    oAcqInfoModes = oFiles(m_sFileName).Infos.AcquisitionInfoModes

    Dim oAcqProps As AcquisitionPropertiesContainer
    oAcqProps = oAcqInfoModes.ActiveProperties

	'Create empty PointDomain in the base file.
	Dim oPointDomains As PointDomains
	Dim oPointDomain As PointDomain
	oPointDomains = oFiles(m_sFileName).GetPointDomains(PTCSignalBuildFlags.ptcBuildPointData3d)

    ' Get the number of datapoints.
	Dim lngPoints As Integer
	lngPoints = oFiles(m_sFileName).Infos.MeasPoints.Count

	' Create an array holding data.
	Dim oData As System.Array = System.Array.CreateInstance(GetType(Single), 6*lngPoints, scanFileCount)	' allocate enough space for 3D data

	' Some variables for the calculating of the yaxis.
	Dim oStat As Statistics
	Dim oVector As Vector

	Dim oBandDomain As BandDomain
	oBandDomain = oFiles(m_sFileName).GetBandDomains().type(PTCDomainType.ptcDomainSpectrum)

	' Attention: necessary to make a cast !!!
	Dim oDomain As Domain
	oDomain = oBandDomain
	Dim ptcType As PTCDomainType
	ptcType = oDomain.type

	Dim oChannel As Channel

	For Each oChannel In oBandDomain.Channels

		' Check for 3D-Signal in Banddomain.
		Dim bIs3D As Boolean
		bIs3D = (oChannel.Caps And PTCChannelCapsType.ptcChannelCapsVector) <> 0

		oStat = New Statistics
		oVector = New Vector

		Dim oSignal As Signal
		For Each oSignal In oChannel.Signals

			Dim bComplex As Boolean
			bComplex = oSignal.Description.Complex

			Dim currentFileIndex = 0
			' Get data from each FastScan file.
			For Each scanFileName As String In scanFileNames

				If Not m_sFileName = scanFileName Then
					' Open next data file.
					If Not OpenFile(oFiles, scanFileName) Then
						MsgBox("Error open file! The macro exits now.", vbOkOnly)
						Exit Sub
					End If
				End If

				' Attention: necessary to make a cast !!!
				Dim oBandDomainTemp As BandDomain
				Dim oDomainTemp As Domain

				oBandDomainTemp = oFiles(scanFileName).GetBandDomains().type(ptcType)
				oDomainTemp = oBandDomainTemp

				Dim oChannelTemp As Channel
				Dim oSignalTemp As Signal
			    Dim oDisplay As Display
				oChannelTemp = oDomainTemp.Channels(oChannel.Name)
				oSignalTemp = oChannelTemp.Signals(oSignal.Name)

				If bComplex Then
			    	oDisplay = oSignalTemp.Displays(PTCDisplayType.ptcDisplayRealImag)
			    Else
			    	oDisplay = oSignalTemp.Displays(PTCDisplayType.ptcDisplayMag)
			    End If

				' Get data, make data array.
				' Normally we have only one band.
				Dim oDataBand As DataBand
				Dim oDataTemp() As Single
				For Each oDataBand In oBandDomainTemp.GetDataBands(oSignalTemp)
					' We have no MultiFrame, that's why frame = 0.
					oDataTemp = oDataBand.GetData(oDisplay,0)
				Next oDataBand

				' Replace NaN (Not a Number) values with 0.
				' NaN values are caused from MeasPoints without data e.g. scan status PTCScanStatus.ptcScanStatusDisabled or ptcScanStatusNone.
				ReplaceNaN(oDataTemp, 0, bComplex, bIs3D, oFiles(scanFileName).Infos.MeasPoints)

				' Add up the data in oStat to calculate the max of
				' all magnitudes for the max of the yaxis.
				If bComplex Then
					oStat.Add(oVector.Magnitude(oDataTemp))
				Else
					oStat.Add(oVector.Magnitude(oVector.Complex(oDataTemp)))
				End If

				' Insert the data in the real data array.
				If bComplex Then
					k=0
					For j = LBound(oDataTemp) To ((UBound(oDataTemp)+1)/2-1)
						oData.SetValue(oDataTemp(k), k, currentFileIndex)
						oData.SetValue(oDataTemp(k+1), k+1, currentFileIndex)
						k=k+2
					Next j
				Else
					For k = LBound(oDataTemp) To UBound(oDataTemp)
						oData.SetValue(oDataTemp(k), k, currentFileIndex)
					Next k
				End If

				' Close open data file.
				If Not m_sFileName = scanFileName Then
					oFiles(scanFileName).Close
				End If

				currentFileIndex = currentFileIndex + 1
    		Next

			' Get min, max for the yaxis.
	        Dim sYMax As Single
	        Dim sYMin As Single
			sYMax = CSng(oVector.Max(oStat.Max()))
			sYMin = -sYMax

			Dim oUsrSignal As Signal
			oUsrSignal = AddSignal(oPointDomains, oSignal, sYMin, sYMax, scanFileCount)

			' Add the data. First resort the data.
			Dim oDataPoint As DataPoint
			Dim iIndex As Integer
			Dim iSkip As Integer

			If bComplex Then
				If bIs3D Then
					ReDim oDataTemp(6*scanFileCount-1)
				Else
					ReDim oDataTemp(2*scanFileCount-1)
				End If
				iSkip=2

			Else
				If bIs3D Then
					ReDim oDataTemp(3*scanFileCount-1)
				Else
					ReDim oDataTemp(scanFileCount-1)
				End If
				iSkip=1
			End If

			oPointDomain = oPointDomains.type(oUsrSignal.Channel.Domain.type)
			For Each oDataPoint In oPointDomain.DataPoints

				' Exclude disabled points.
				Dim oMeasPoint As MeasPoint
				oMeasPoint = oDataPoint.MeasPoint
				If (oMeasPoint.ScanStatus And PTCScanStatus.ptcScanStatusDisabled) = 0 Then

					iIndex = oDataPoint.MeasPoint.Index-1

					If bIs3D Then
						Dim d As Long
						For d = 0 To 2
			                For k = 0 To scanFileCount-1
								'real
								oDataTemp((k+d*scanFileCount)*iSkip)=oData.GetValue((iIndex+d*lngPoints)*iSkip, k)
								If bComplex Then
									'imag
									oDataTemp((k+d*scanFileCount)*iSkip + 1)=oData.GetValue((iIndex+d*lngPoints)*iSkip + 1, k)
								End If
							Next k
						Next d
					Else
		                For k = 0 To scanFileCount-1
							oDataTemp(k*iSkip)=oData.GetValue(iIndex*iSkip, k)
							If bComplex Then
								oDataTemp(k*iSkip+1)=oData.GetValue(iIndex*iSkip+1, k)
							End If
						Next k
					End If

					' Add the data.
					oDataPoint.SetData(oUsrSignal, 1, oDataTemp)
				End If
			Next oDataPoint

		Next oSignal

		oStat = Nothing
		oVector = Nothing

	Next oChannel

	' Close the base file.
	oFiles(m_sFileName).Save()
	oFiles(m_sFileName).Close()

	' Now clean up and show the result.
	' ---------------------------------

	If MsgBox("Do you want to delete the files for the individual frequencies and keep only the combined result?", vbYesNo Or vbQuestion) = vbYes Then
		' Delete all temporary files.
		Dim i As Integer
		For i = 1 To scanFileNames.Count-1
			' if we are still in acquisition mode, we cannot delete the last file
			If (Application.Mode = PTCApplicationMode.ptcApplicationModePresentation) Then
				Kill scanFileNames(i)
			End If
		Next
	End If

	' Show the result.
	oPresApp.Windows.CloseAll()
	oPresApp.Documents.Open(scanFileNames(0))

	MsgBox("Macro has finished.", vbOkOnly)

	oPresApp.Activate()

MErrorHandler:
	Select Case Err.Number
	Case 0
	Case Else
		For Each file As KeyValuePair(Of String, PolyFile) In oFiles
			file.Value.Close()
		Next
		MsgBox("Macro has finished with error:" & Err.Description, vbOkOnly)
		Resume MErrorHandler
	End Select
End Sub

Function AddSignal(oPointDomains As PointDomains, oSignal As Signal, sYMin As Single, sYMax As Single, fileCount As Integer)
' -------------------------------------------------------------------------------
'	Add user signal.
' -------------------------------------------------------------------------------
	AddSignal = Nothing

	Dim oSigDesc As SignalDescription
	oSigDesc = oSignal.Description.Clone()

	With oSigDesc
		.Name = oSignal.Channel.Domain.Name + " " + oSignal.Channel.Name + " " + oSignal.Name
		.DataType = PTCDataType.ptcDataPoint

		With oSigDesc.XAxis
			.Name = "Frequency"
			.Unit = "Hz"
			.Min = m_dFreqStart
			.Max = m_dFreqEnd
			.MaxCount = fileCount
		End With

		With oSigDesc.YAxis
			.Min = sYMin
			.Max = sYMax
		End With
	End With

	Dim oExistingSignal As Signal
	oExistingSignal = oPointDomains.FindSignal(oSigDesc, True)

	' Check if a signal with the same name exits already. We will overwrite it.
	If (Not oExistingSignal Is Nothing) Then
		oExistingSignal.Channel.Signals.Remove(oSigDesc.Name)
	End If

	AddSignal = oPointDomains.AddSignal(oSigDesc)
End Function


Function FastScanMeasurement(ByRef scanFileNames As List(Of String)) As Boolean
' -------------------------------------------------------------------------------
'	Execute FastScan measurement.
' -------------------------------------------------------------------------------
	On Error GoTo MErrorHandler

	Dim dblFreqScan As Double
	Dim i As Integer

	i = 0
	For dblFreqScan = m_dFreqStart To m_dFreqEnd Step m_dFreqStep

		Dim Acq As Acquisition
		Acq = Application.Acquisition

        Acq.Mode = PTCAcqMode.ptcAcqModeFastScan

		Dim ActiveProp As AcquisitionProperties
		ActiveProp = Acq.ActiveProperties

        Dim FastScan As FastScanAcqProperties
        FastScan = ActiveProp.FastScansProperties(1)
        FastScan.Frequency = dblFreqScan
        FastScan.Bandwidth = m_dBandWidth

	    Dim Avarage As	AverageAcqProperties
	    Avarage = ActiveProp.AverageProperties
		Select Case m_ptcAveragingMode
		Case PTCAveragingMode.ptcAveragingModeOff
			Avarage.type = PTCAverageType.ptcAverageOff
		Case PTCAveragingMode.ptcAveragingModeMagnitude
			Avarage.type = PTCAverageType.ptcAverageMagnitude
		Case PTCAveragingMode.ptcAveragingModeComplex
			Avarage.type = PTCAverageType.ptcAverageComplex
		End Select
		Avarage.Count = m_lAveraging

        Acq.ScanFileName = scanFileNames(i)

        Acq.Scan(PTCScanMode.ptcScanAll)

        While Acq.State <> PTCAcqState.ptcAcqStateStopped
            Wait 0.01
        End While

		i = i + 1
	Next dblFreqScan

	FastScanMeasurement = True

MErrorHandler:
	Select Case Err.Number
	Case 0
	Case Else
		MsgBox("Error:" & Err.Description, vbOkOnly)
		FastScanMeasurement = False
		Resume MErrorHandler
	End Select
End Function

' *******************************************************************************
' * Helper functions and subroutines
' *******************************************************************************

Const c_OFN_HIDEREADONLY As Long = 4
Const c_DEFAULT_FILE As String = "C:\temp\SteppedFastScan\Test.svd"

Private Function FastScanDialog() As Boolean
' -------------------------------------------------------------------------------
'	Base dialog for user input.
' -------------------------------------------------------------------------------
	Begin Dialog UserDialog 480,245,"Stepped FastScans",.SetupDialogProc ' %GRID:10,7,1,1
		GroupBox 10,14,310,133,"Frequency",.GroupBoxFrequency
		Text 30,35,140,14,"Start Frequency [Hz]:",.TextStart
		TextBox 180,35,130,21,.fStart
		Text 30,63,140,14,"End Frequency [Hz]:",.TextEnd
		TextBox 180,63,130,21,.fEnd
		Text 30,91,140,14,"Step Frequency [Hz]:",.TextStop
		TextBox 180,91,130,21,.fStep
		Text 30,119,140,14,"Bandwidth [Hz]:",.TextBandwidth
		TextBox 180,119,130,21,.fBandwidth

		GroupBox 330,14,140,133,"Averaging",.GroupBoxAveraging
		OptionGroup .iAveragingMode
			OptionButton 350,35,90,14,"Off",.OptionButtonOff
			OptionButton 350,63,90,14,"Magnitude",.OptionButtonMagnitude
			OptionButton 350,91,90,14,"Complex",.OptionButtonComplex
		TextBox 350,119,90,21,.iAveraging

		GroupBox 10,154,460,56,"Filename",.GroupBoxFilename
		TextBox 30,175,280,21,.sFile
		PushButton 350,175,100,21,"Browse...",.pbFile

		OKButton 250,217,100,21
		CancelButton 370,217,100,21
	End Dialog
    Dim dlg As UserDialog

   ' Show dialog (wait for any button pressed).
    Dim iRe As Integer
    iRe = Dialog(dlg)

	Select Case iRe
		Case 0
	    	' Cancel button pressed.
			FastScanDialog = False
		Case -1
	    	' OK button pressed.
			FastScanDialog = True
	End Select
End Function

Private Function SetupDialogProc(DlgItem$, Action%, SuppValue&) As Boolean
' -------------------------------------------------------------------------------
'	Second dialog for user input.
' -------------------------------------------------------------------------------
	Select Case Action%
	Case 1 ' Dialog box initialization.
		DlgValue  "iAveragingMode", PTCAveragingMode.ptcAveragingModeOff
		DlgEnable "iAveraging"    , False
		DlgText   "iAveraging"    , "3"

	Case 2 ' Value changing or button pressed.

		Select Case DlgItem
			Case "iAveragingMode"
				If DlgValue("iAveragingMode") = PTCAveragingMode.ptcAveragingModeOff Then
					DlgEnable "iAveraging",False
				Else
					DlgEnable "iAveraging",True
				End If

			Case "pbFile"
				' Get filename and path.
				DlgText "sFile", FileSaveDialog()

				' Prevent button press from closing the dialog box.
				SetupDialogProc = True

			Case "OK"
				If Not CheckDialogInput Then
					' Prevent button press from closing the dialog box.
					SetupDialogProc = True
				End If
		End Select

	Case 3 ' TextBox or ComboBox text changed.
	Case 4 ' Focus changed.
	Case 5 ' Idle.
	Case 6 ' Function key.
	End Select
End Function

Private Function FileSaveDialog() As String
' -------------------------------------------------------------------------------
'	Base dialog to select a file for saving data.
' -------------------------------------------------------------------------------
	On Error GoTo MCreateError

	Dim fod As Object
	fod = CreateObject("MSComDlg.CommonDialog")
	fod.Filter = c_strFileFilter
	fod.Flags = c_OFN_HIDEREADONLY
	fod.CancelError = True
	On Error GoTo MCancelError
	fod.ShowSave
	FileSaveDialog = fod.FileName

	GoTo MEnd
MCancelError:
	FileSaveDialog = ""
	GoTo MEnd
MCreateError:
	FileSaveDialog = GetFilePath(, c_strFileExt, CurDir(), "Save as", 2)
MEnd:
End Function

Private Function CheckDialogInput() As Boolean
' -------------------------------------------------------------------------------
'	Check the user input.
' -------------------------------------------------------------------------------
	Dim strValue As String

	strValue = DlgText("fStart")
	If (strValue = "") Then
		MsgBox("Please insert the start frequency.", vbOkOnly)
		CheckDialogInput = False
		Exit Function
	End If

	strValue = DlgText("fEnd")
	If (strValue = "") Then
		MsgBox("Please insert the end frequency.", vbOkOnly)
		CheckDialogInput = False
		Exit Function
	End If

	strValue = DlgText("fStep")
	If (strValue = "") Then
		MsgBox("Please insert the frequency step.", vbOkOnly)
		CheckDialogInput = False
		Exit Function
	End If

	strValue = DlgText("fBandwidth")
	If (strValue = "") Then
		MsgBox("Please insert the bandwidth.", vbOkOnly)
		CheckDialogInput = False
		Exit Function
	End If

	strValue = DlgText("sFile")
	If (strValue = "") Then
		MsgBox("Please insert the file name.", vbOkOnly)
		CheckDialogInput = False
		Exit Function
	End If

	' Set the variables to work with it.
	m_dFreqStart = CDbl( DlgText("fStart"))
	m_dFreqEnd   = CDbl( DlgText("fEnd"))
	m_dFreqStep  = CDbl( DlgText("fStep"))
	m_dBandWidth = CDbl( DlgText("fBandwidth"))
	m_lAveraging = CLng( DlgText("iAveraging"))
	m_ptcAveragingMode = DlgValue("iAveragingMode")
	m_sFileName  = DlgText("sFile")

	If m_dFreqStart >= m_dFreqEnd Then
		MsgBox("Please insert a correct start and end frequency.", vbOkOnly)
		CheckDialogInput = False
		Exit Function
	End If

	If m_dFreqStep >= m_dFreqEnd - m_dFreqStart Then
		MsgBox("Please insert a correct start and end frequency and frequency step.", vbOkOnly)
		CheckDialogInput = False
		Exit Function
	End If

	If m_dFreqStep <= 0 Then
		MsgBox("Please insert a correct frequency step (>0).", vbOkOnly)
		CheckDialogInput = False
		Exit Function
	End If

	If m_dBandWidth <= 0 Then
		MsgBox("Please insert a correct bandwidth (>0).", vbOkOnly)
		CheckDialogInput = False
		Exit Function
	End If

	If m_ptcAveragingMode <> POLYPROPERTIESLib.ptcAverageOff Then
		strValue = DlgText("iAveraging")
		If (strValue = "") Then
			MsgBox("Please insert the averaging.", vbOkOnly)
			CheckDialogInput = False
			Exit Function
		End If

		If ((m_lAveraging < 3) Or (m_lAveraging > 1000000&)) Then
			MsgBox("The averaging should be between 3 and 1.000.000.", vbOkOnly)
			CheckDialogInput = False
			Exit Function
		End If

	End If

	CheckDialogInput = True
End Function

Private Function FastScanFiles() As List(Of String)
' -------------------------------------------------------------------------------
' 	Make an array for base file and temporary files.
' -------------------------------------------------------------------------------
	Dim scanFileNames As New List(Of String)
	Dim dblFreqScan As Double
	Dim i As Integer

	m_sFilePath = Left(m_sFileName, InStrRev(m_sFileName,"\"))

	i = 0
	For dblFreqScan = m_dFreqStart To m_dFreqEnd Step m_dFreqStep
		If i = 0 Then
			scanFileNames.Add(m_sFileName)
		Else
			' Round to avoid numerical errors.
			scanFileNames.Add(m_sFilePath + CStr(Round(dblFreqScan,4)) + ".svd")
		End If
		i = i + 1
	Next dblFreqScan
	Return scanFileNames
End Function

Private Function SwitchToAcquisitionMode() As Boolean
' -------------------------------------------------------------------------------
'	Check and switch to AcquisitionMode, so we can make measurement.
' -------------------------------------------------------------------------------
	If Application.Mode = PTCApplicationMode.ptcApplicationModePresentation Then
		Dim oAcquisitionInstance As New AcquisitionInstance
		If (oAcquisitionInstance.IsRunning) Then
			MsgBox "Please run this macro in the PSV acquisition instance."
			SwitchToAcquisitionMode = False
			Exit Function
		End If
		Application.Mode = PTCApplicationMode.ptcApplicationModeAcquisition
	End If
	If Application.Mode = PTCApplicationMode.ptcApplicationModePresentation Then
		Beep
		MsgBox("Cannot switch to Acquisition Mode.", vbOkOnly)
		SwitchToAcquisitionMode = False
	Else
		SwitchToAcquisitionMode = True
	End If
End Function

Private Function GetPresentationInstance() As Application
' -------------------------------------------------------------------------------
'	Gets the PSV instance that is running in presentation mode
'   If possible this instance is switched to presentation mode,
'   otherwise the application object of the running presentation instance
'   is returned.
' -------------------------------------------------------------------------------
	If (Application.Mode = PTCApplicationMode.ptcApplicationModePresentation) Then
		GetPresentationInstance = Application
		Exit Function
	End If
	Dim oPresentationInstance As New PresentationInstance
	If (oPresentationInstance.IsRunning) Then
		GetPresentationInstance = oPresentationInstance.GetApplication(False, 0)
		Exit Function
	End If
	Application.Mode = PTCApplicationMode.ptcApplicationModePresentation
	If (Application.Mode = PTCApplicationMode.ptcApplicationModeAcquisition) Then
		MsgBox("Cannot switch to Presentation Mode.", vbOkOnly)
		GetPresentationInstance = Nothing
	Else
		GetPresentationInstance = Application
	End If
End Function

Private Function OpenFile(ByRef oFileList As Dictionary(Of String, PolyFile), strFileName As String) As Boolean
' -------------------------------------------------------------------------------
'	Instantiate PolyFile object, open the file.
' -------------------------------------------------------------------------------
	On Error GoTo MErrorHandler
	Dim bRe As Boolean
	bRe = True

	Dim oFile As PolyFile
	If Not oFileList.ContainsKey(strFileName) Then
		oFile = New PolyFile
		oFileList.Add(strFileName, oFile)
	Else
		oFile = oFileList(strFileName)
	End If

	If oFile.ReadOnly Then
		oFile.ReadOnly = False
	End If

	On Error Resume Next
	If Not oFile.IsOpen Then
		oFile.Open (strFileName)
	End If

	On Error GoTo 0
	If Not oFile.IsOpen Then
		MsgBox("Can not open the file."& vbCrLf & _
		"Check the file attribute is not read only!", vbExclamation)
		bRe = False
	End If
MErrorHandler:
	Select Case Err.Number
	Case 0
		OpenFile = bRe
	Case Else
		bRe = False
		MsgBox("Error:" & Err.Description, vbOkOnly)
		Resume MErrorHandler
	End Select
End Function

Private Sub ReplaceNaN(ByRef oData() As Single, newValue As Single, complex As Boolean, is3D As Boolean, oMeasPoints As MeasPoints)
' -------------------------------------------------------------------------------
'	Replace NaN (Not a Number) band data values in oData vector with the newValue.
'	NaN values are caused from measpoints without data e.g. scan status PTCScanStatus.ptcScanStatusDisabled or ptcScanStatusNone.
' -------------------------------------------------------------------------------
	Dim oMeasPoint As MeasPoint
	For Each oMeasPoint In oMeasPoints

		If (oMeasPoint.ScanStatus And PTCScanStatus.ptcScanStatusValid) = 0 Then

			Dim indexZeroBased As Long = oMeasPoint.Index-1

			ReplaceNaNValue(oData, 0, complex, indexZeroBased, newValue)

			If is3D Then

				' replace the y-direction values
				ReplaceNaNValue(oData, oMeasPoints.Count, complex, indexZeroBased, newValue)

				' replace the z-direction values
				ReplaceNaNValue(oData, 2*oMeasPoints.Count, complex, indexZeroBased, newValue)

			End If

		End If

	Next oMeasPoint

End Sub

Private Sub ReplaceNaNValue(ByRef oData() As Single, offset3D As Long, complex As Boolean, index As Long, newValue As Single)

	' replace the values for one direction complex or non complex.
	If complex Then
		oData(2*offset3D+2*index) = newValue
		oData(2*offset3D+2*index+1) = newValue
	Else
		oData(offset3D+index) = newValue
	End If

End Sub
