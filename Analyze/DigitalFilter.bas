'#Reference {A65C100F-C1FE-4C3B-9C43-46F4FB4C3BC3}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolyFile.dll#Polytec PolyFile Type Library, $Revision: 3$ UnicodeRelease, Build on Jul 13 2005 at 21:34:50
'#Reference {58E59DE7-08B4-4975-AEBC-206321D07689}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolyDigitalFilters.dll#Polytec PolyDigitalFilters Type Library, $Revision: 3$ UnicodeRelease, Build on Jul 13 2005 at 21:23:18
'#Reference {E44752C9-2D41-48A6-9B74-66D5B7505325}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolyProperties.dll#Polytec PolyProperties Type Library, $Revision: 3$ UnicodeRelease, Build on Jul 13 2005 at 21:23:19
'#Reference {E68EA160-8AD4-11D3-8F08-00104BB924B2}#1.0#0#C:\Program Files\Common Files\Polytec\COM\SignalDescription.dll#Polytec SignalDescription Type Library, $Revision: 3$ UnicodeRelease, Build on Jul 13 2005 at 21:23:16
'#Reference {CE68D434-5052-431F-BE75-F3C23458127A}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolySignal.dll#Polytec PolySignal Type Library, $Revision: 3$ UnicodeRelease, Build on Jul 13 2005 at 21:23:17
'#Reference {F08ACE20-C7AD-46CA-8001-D5158D9B0224}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolyMath.dll#Polytec PolyMath Type Library, $Revision: 3$ UnicodeRelease, Build on Jul 13 2005 at 21:23:21
' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This Macro calculates the tranmission functions of the digital filters
' that were used during data taking and adds them into the FFT domain
' of the first single point data of the file. The transmission functions for all
' qualities of the filters (very low to very high) are put as frames into
' the user defined data sets.
'
' When running the macro you are asked to navigate to a single point files
' or scan file. This file has to meet the following conditions:
'
' - you have to have exclusive write access to the file. We strongly recommend to
'   use a backup copy of your original file with this macro. The macro
'   will fail if the file is open in PSV or VibSoft.
'
' To display the calculated data in PSV/VibSoft do the following:
' - start PSV/VibSoft and open the file
' - select Analyzer/Domain/FFT
' - select Analyzer/Channel/Usr
' - select one of the signals offered in Analyzer/Signal
'
' References
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
Option Explicit

Const c_strFileFilter As String = "Single Point File (*.pvd)|*.pvd|Scan File (*.svd)|*.svd|All Files (*.*)|*.*||"
Const c_strFileExt As String = "pvd;svd"

' size of the transmission function
Const c_lCount As Long = 4096

Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	Dim oFile As New PolyFile

	' get filename and path
	Dim strFileName As String
	strFileName = FileOpenDialog()

	If (strFileName = "") Then
		MsgBox("No filename has been specified, macro exits now.", vbOkOnly)
		Exit Sub
	End If

	If (MsgBox("This macro will modify the file '" + strFileName + _
		"'. We strongly recommend to work with a backup copy of original data only. " + _
		"Do you want to continue?", vbYesNo) = vbNo) Then
		Exit Sub
	End If

	' we have to open the file for read/write, otherwise we cannot save our
	' user defined dataset to the file
	If Not OpenFile(oFile, strFileName) Then
		Exit Sub
	End If

	Dim oPointDomains As PointDomains
    Set oPointDomains = oFile.GetPointDomains(ptcBuildPointData3d)

	Dim oPointDomain As PointDomain
	Set oPointDomain = oPointDomains.type(ptcDomainSpectrum)

	Dim oDomain As Domain
	Set oDomain = oPointDomain

	Dim oDataPoint As DataPoint
	Set oDataPoint = oPointDomain.DataPoints(1)

	Dim oChannel As Channel
	Dim oSignal As Signal

	Dim oSigPro As New SignalProcessing

	Dim oAcqInfoModes As AcquisitionInfoModes
	Set oAcqInfoModes = oFile.Infos.AcquisitionInfoModes
	Dim oAcqProps As AcquisitionPropertiesContainer
	Set oAcqProps = oAcqInfoModes.ActiveProperties

	Dim oChannelsAcqProps As ChannelsAcqPropertiesContainer
	Set oChannelsAcqProps = oAcqProps.ChannelsProperties

	Dim dSampleFrequency As Double
	dSampleFrequency = GetSampleFrequency(oAcqInfoModes.ActiveMode, oAcqProps)
	If (dSampleFrequency = 0) Then
		MsgBox("The acquisition mode of the file is not supported."& vbCrLf & _
		"The macro ends now", vbExclamation)
		oFile.Close()
		Exit Sub
	End If

	' loop over all channels
	Dim oChannelAcqProps As ChannelAcqPropertiesContainer
	For Each oChannelAcqProps In oChannelsAcqProps
		If (oChannelAcqProps.Active) Then
			' channel is active
			Dim oFilter As DigitalFilter
			Set oFilter = oChannelAcqProps.DigitalFilter

			If (Not oFilter Is Nothing) Then
				' channel has an digital filter
				Dim oUsrSignal As Signal
				Set oUsrSignal = GetUserSignal(oChannelAcqProps.Name + " " + oFilter.Name, oPointDomains, dSampleFrequency, c_lCount)

				If (Not oUsrSignal Is Nothing) Then
					' loop over the qualities and store them as frames of the user defined datasets
					Dim ptcQuality As PTCDigitalFilterQuality
					Dim lFrame As Long
					lFrame = 0
					For ptcQuality = ptcDigitalFilterQualityVeryLow To ptcDigitalFilterQualityVeryHigh
						lFrame = lFrame + 1

						' calculate the coeficients of the digital filter from the settings
						Dim Coeff() As Single
						Coeff = GetFilterCoefficients(oSigPro, oFilter, ptcQuality, dSampleFrequency)

						' calculate the transmission function of the filter
						Dim Trans() As Single
						Trans = oSigPro.Transmission(Coeff, c_lCount)

						oDataPoint.SetData(oUsrSignal, lFrame, Trans)
					Next
				End If
			End If
		End If
	Next

	oFile.Save()
	oFile.Close()

	MsgBox("Macro has finished.", vbOkOnly)
End Sub

Function GetFilterCoefficients(oSigPro As SignalProcessing, oFilter As DigitalFilter, ptcQuality As PTCDigitalFilterQuality, dSampleFreq As Double)
' -------------------------------------------------------------------------------
' Calculates the coefficients of a digital filter.
' -------------------------------------------------------------------------------
	Select Case oFilter.type
		Case ptcDigitalFilterLowpass
			Dim oFilterLowpass As DigitalFilterLowpass
			Set oFilterLowpass = oFilter
			GetFilterCoefficients = oSigPro.FilterCoefficients(oFilter.type, ptcQuality, dSampleFreq, oFilterLowpass.CutoffFreq)
		Case ptcDigitalFilterHighpass
			Dim oFilterHighpass As DigitalFilterHighpass
			Set oFilterHighpass = oFilter
			GetFilterCoefficients = oSigPro.FilterCoefficients(oFilter.type, ptcQuality, dSampleFreq, oFilterHighpass.CutoffFreq)
		Case ptcDigitalFilterBandpass
			Dim oFilterBandpass As DigitalFilterBandpass
			Set oFilterBandpass = oFilter
			GetFilterCoefficients = oSigPro.FilterCoefficients(oFilter.type, ptcQuality, dSampleFreq, oFilterBandpass.CutoffFreq1, oFilterBandpass.CutoffFreq2)
		Case ptcDigitalFilterNotch
			Dim oFilterNotch As DigitalFilterNotch
			Set oFilterNotch = oFilter
			GetFilterCoefficients = oSigPro.FilterCoefficients(oFilter.type, ptcQuality, dSampleFreq, oFilterNotch.CutoffFreq1, oFilterNotch.CutoffFreq2)
	End Select
End Function

Function GetSampleFrequency(ptcMode As PTCAcqMode, oAcqProps As AcquisitionPropertiesContainer) As Double
' -------------------------------------------------------------------------------
' Gets the sampling frequency depending of the different acquisition modes.
' Returns 0 for an unsupported acquisition mode.
' -------------------------------------------------------------------------------

	GetSampleFrequency = 0.0

	Select Case ptcMode
	Case ptcAcqModeFft
		GetSampleFrequency = oAcqProps.FftProperties.SampleFrequency
	Case ptcAcqModeTime
		GetSampleFrequency = oAcqProps.TimeProperties.SampleFrequency
	End Select
End Function


Function GetUserSignal(strName As String, oPointDomains As PointDomains, dSampleFrequency As Double, lCount As Long) As Signal
' -------------------------------------------------------------------------------
' Adds a user signal to the point domains. For the description the axes of the given display are used.
' Returns Nothing if the user does not want to overwrite an existing user signal with the same name.
' -------------------------------------------------------------------------------

	Dim oUsrSigDesc As New SignalDescription

	With oUsrSigDesc
		.Name = strName
		.DataType = ptcDataPoint
		.DomainType = ptcDomainSpectrum
		.Complex = False
		.PowerSignal = False

		.XAxis.Name = "Frequency"
		.XAxis.Unit = "Hz"
		.XAxis.MaxCount = lCount
		.XAxis.Min = 0
		.XAxis.Max = 0.5 * dSampleFrequency

		.YAxis.Name = "Transmission"
		.YAxis.Unit = ""
		.YAxis.Min = 0.0
		.YAxis.Max = 1.0

		.FunctionType = ptcFunctionFiniteImpulseResponseFilterType
		.ResponseDOFs.Assign(0, ptcScalarDir, "Usr", "Transmission", "")
	End With

	Set GetUserSignal = Nothing

	Dim oUsrSignal As Signal

	' check if a signal with the same name exits already
	Set oUsrSignal = oPointDomains.FindSignal(oUsrSigDesc, True)

	If (oUsrSignal Is Nothing) Then
		Set oUsrSignal = oPointDomains.AddSignal(oUsrSigDesc)
	Else
		If (MsgBox("A user defined signal with the name '" + oUsrSigDesc.Name + "' exists already. Do you want to replace it?", vbYesNo) = vbYes) Then
			oUsrSignal.Channel.Signals.Update(oUsrSignal.Name, oUsrSigDesc)
		End If
	End If

	Set GetUserSignal = oUsrSignal

End Function


' *******************************************************************************
' * Helper functions and subroutines
' *******************************************************************************

Const c_OFN_HIDEREADONLY As Long = 4

Private Function FileOpenDialog() As String
' -------------------------------------------------------------------------------
' Select file.
' -------------------------------------------------------------------------------
	On Error GoTo MCreateError
	Dim fod As Object
	Set fod = CreateObject("MSComDlg.CommonDialog")
	fod.Filter = c_strFileFilter
	fod.Flags = c_OFN_HIDEREADONLY
	fod.CancelError = True
	On Error GoTo MCancelError
	fod.ShowOpen
	FileOpenDialog = fod.FileName
	GoTo MEnd
MCancelError:
	FileOpenDialog = ""
	GoTo MEnd
MCreateError:
	FileOpenDialog = GetFilePath(, c_strFileExt, CurDir(), "Select a file", 2)
MEnd:
End Function

Private Function OpenFile(oFile  As PolyFile, strFileName As String) As Boolean
' -------------------------------------------------------------------------------
' Instantiate PolyFile object, open the File.
' -------------------------------------------------------------------------------
	On Error GoTo MErrorHandler
	Dim bRe As Boolean
	bRe = True

	Set oFile = New PolyFile
	If oFile.ReadOnly Then
		oFile.ReadOnly = False
	End If

	On Error Resume Next
	oFile.Open (strFileName)

	On Error GoTo 0
	If Not oFile.IsOpen Then
		MsgBox("Can not open the file."& vbCrLf & _
		"Check the file attribute is not read only!", vbExclamation)
		bRe = False
	End If
MErrorHandler:
	Select Case Err
	Case 0
		OpenFile = bRe
	Case Else
		bRe = False
		Resume MErrorHandler
	End Select
End Function
