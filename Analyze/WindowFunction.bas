'#Reference {E68EA160-8AD4-11D3-8F08-00104BB924B2}#1.0#0#C:\Program Files\Common Files\Polytec\COM\SignalDescription.dll#Polytec SignalDescription Type Library, $Revision: 4$ UnicodeRelease, Build on Jun 19 2006 at 21:12:47
'#Reference {E44752C9-2D41-48A6-9B74-66D5B7505325}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolyProperties.dll#Polytec PolyProperties Type Library, $Revision: 4$ UnicodeRelease, Build on Jun 19 2006 at 21:12:54
'#Reference {A65C100F-C1FE-4C3B-9C43-46F4FB4C3BC3}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolyFile.dll#Polytec PolyFile Type Library, $Revision: 4$ UnicodeRelease, Build on Jun 19 2006 at 21:32:47
'#Reference {CE68D434-5052-431F-BE75-F3C23458127A}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolySignal.dll#Polytec PolySignal Type Library, $Revision: 4$ UnicodeRelease, Build on Jun 19 2006 at 21:12:48
'#Reference {F08ACE20-C7AD-46CA-8001-D5158D9B0224}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolyMath.dll#Polytec PolyMath Type Library, $Revision: 4$ UnicodeRelease, Build on Jun 19 2006 at 21:13:02
' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This Macro demonstrates the influence of applying a window
' function to the time data. For that the window function
' that has been applied to calculate the FFTs is recalculated
' and added to the data as user defined data sets.
'
' The original data, the window function and the result of
' applying the window function to the time data are
' copied to frame 1, 2 and 3 of user defined datasets.
'
' The macro tries to calculate the FFTs from the raw time data before and
' after applying the window function. This might fail for several reasons:
' - the number of FFT lines is not as expected (you are using from/to
'   parameters for a diminished bandwidths)
' - you are using averaging in the FFT domain. In this case the macro will
'   not reproduce the original FFT data because only the time data of the
'   first averaging block is used.
'
' When running the macro you are asked to navigate to a single point files.
' This file has to meet the following conditions:
'
' - the file has to be acquired in FFT mode. Time domain files are not supported
'   as they do not contain the acquisition settings for the window functions.
' - scan files are not supported because they do not contain time and FFT data
'   at the same time.
' - you have to have exclusive write access to the file. We strongly recommend to
'   use a backup copy of your original file with this macro. The macro
'   will fail if the file is open in PSV or VibSoft.
'
' To display the calculated data in PSV/VibSoft do the following:
' - start PSV/VibSoft and open the file
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

Const c_strFileFilter As String = "Single Point File (*.pvd)|*.pvd|All Files (*.*)|*.*||"
Const c_strFileExt As String = "pvd"

Sub Main
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

    Dim oChannelsAcqProps As ChannelsAcqPropertiesContainer
    Set oChannelsAcqProps = oFile.Infos.AcquisitionInfoModes.ActiveProperties.ChannelsProperties

	Dim oPointDomains As PointDomains
	Set oPointDomains = oFile.GetPointDomains(ptcBuildPointData3d)

	Dim oPointDomainTime As PointDomain
	Set oPointDomainTime = oPointDomains.type(ptcDomainTime)
	Dim oPointDomainFFT As PointDomain
	If (oPointDomains.Exists(ptcDomainSpectrum)) Then
		Set oPointDomainFFT = oPointDomains.type(ptcDomainSpectrum)
	End If

	Dim oDomainTime As Domain
	Set oDomainTime = oPointDomainTime

	Dim oDataPoint As DataPoint

	Dim oVector As New Vector
	Dim oSigPro As New SignalProcessing

	Dim oChannelTime As Channel

	' loop over all channels and signals of the time domain
	For Each oChannelTime In oDomainTime.Channels
		' ignore user defined channels
		If (oChannelTime.Name <> "Usr") Then
			' get the window function of this channel
			Dim WndFctParams() As Double
			Dim WndFctType As PTCWindowFunction
			WndFctType = GetWindowFunction(oFile, oChannelTime, WndFctParams)

			Dim oSignalTime As Signal
			For Each oSignalTime In oChannelTime.Signals
				' add a new user signal to the time domain
				Dim oDisplayTime As Display
				Set oDisplayTime = oSignalTime.Displays.type(ptcDisplaySamples)
				Dim strName As String
				strName = "Window Function: " + oChannelTime.Name + " " + oSignalTime.Name
				Dim oUsrSignalTime As Signal
				Set oUsrSignalTime = GetUserSignal(strName, False, oPointDomains, oDisplayTime, oChannelsAcqProps)

				' if we have an FFT domain add a new user signal to the FFT domain - same channel and signal
				If (Not oPointDomainFFT Is Nothing) Then
					If (oPointDomainFFT.Channels.Exists(oChannelTime.Name)) Then
						Dim oDomainFFT As Domain
						Set oDomainFFT = oPointDomainFFT
						Dim oSignalFFT As Signal
						Set oSignalFFT = oDomainFFT.Channels(oChannelTime.Name).Signals(oSignalTime.Name)
						Dim oDisplayFFT As Display
						Set oDisplayFFT = oSignalFFT.Displays.type(ptcDisplayMag)
						Dim oUsrSignalFFT As Signal
						Set oUsrSignalFFT = GetUserSignal(strName, True, oPointDomains, oDisplayFFT, oChannelsAcqProps)
						Dim lFFTLines As Long
						lFFTLines = oUsrSignalFFT.Description.XAxis.MaxCount
					End If
				End If

				If (Not oUsrSignalTime Is Nothing) Then
					Dim lDataPoint As Long
					lDataPoint = 1
					For Each oDataPoint In oPointDomainTime.DataPoints
						' get the original time data
						Dim Data() As Single
						Data = oDataPoint.GetData(oDisplayTime, 0)

						' calculate the window function
						Dim dRMSCorrection As Double
						Dim WndFct() As Single
						WndFct = oSigPro.WindowFunction(WndFctType, UBound(Data) - LBound(Data) + 1, WndFctParams, dRMSCorrection)

						' set the data: original, window function, result of applying the function
						oDataPoint.SetData(oUsrSignalTime, 1, Data)
						oDataPoint.SetData(oUsrSignalTime, 2, WndFct)
						oDataPoint.SetData(oUsrSignalTime, 3, oVector.Mul(Data, WndFct))

						If (Not oUsrSignalFFT Is Nothing) Then
							Dim oDataPointFFT As DataPoint
							Set oDataPointFFT = oPointDomainFFT.DataPoints(lDataPoint)

							' set the data: FFT with no window function and FFT with window function
							oDataPointFFT.SetData(oUsrSignalFFT, 1, oSigPro.FFT(Data, lFFTLines))
							oDataPointFFT.SetData(oUsrSignalFFT, 2, oSigPro.FFT(Data, lFFTLines, WndFct))
						End If
						lDataPoint = lDataPoint + 1
					Next oDataPoint
				End If
			Next oSignalTime
		End If
	Next oChannelTime

	oFile.Save()
	oFile.Close()

	MsgBox("Macro has finished.", vbOkOnly)
End Sub

Function GetUserSignal(strName As String, bComplex As Boolean, oPointDomains As PointDomains, oDisplay As Display, oChannelsAcqProps As ChannelsAcqPropertiesContainer) As Signal
	' Adds a user signal to the user signals.
	' Returns Nothing if the user does not want to overwrite an existing user signal with the same name

	Dim oSignal As Signal
	Set oSignal = oDisplay.Signal

	' fill the properties of the user signal description

	Dim oUsrSigDesc As New SignalDescription
	Set oUsrSigDesc = oSignal.Description.Clone()

	With oUsrSigDesc
		.Name = strName
		.Complex = bComplex
		.PowerSignal = False
	End With

	Set GetUserSignal = Nothing

	Dim oSignalUser As Signal
	Set oSignalUser = oPointDomains.FindSignal(oUsrSigDesc, True)

	' check if a signal with the same name exits already
	If (oSignalUser Is Nothing) Then
		Set oSignalUser = oPointDomains.AddSignal(oUsrSigDesc)
	Else
		If (MsgBox("A user defined signal with the name '" + oUsrSigDesc.Name + "' exists already. Do you want to replace it?", vbYesNo) = vbYes) Then
			oSignalUser.Channel.Signals.Update(oSignalUser.Name, oUsrSigDesc)
		End If
	End If

	Set GetUserSignal = oSignalUser

End Function

Function GetChannelAcqPropsByName(oChannelsAcqProps As ChannelsAcqPropertiesContainer, strName As String) As ChannelAcqPropertiesContainer
' -------------------------------------------------------------------------------
' gets the channel acquisition properties by the short name of the channel
' we have to use the name and not the SourceChannel value for this, because 3D channels have a single entry
' in the ChannelsAcqProperties collection but occupy three source channel numbers
' -------------------------------------------------------------------------------
    Dim oChannelAcqProps As ChannelAcqPropertiesContainer
	For Each oChannelAcqProps In oChannelsAcqProps
		If strName = oChannelAcqProps.ShortName Then
			Set GetChannelAcqPropsByName = oChannelAcqProps
			Exit Function
		End If
	Next
	Err.Raise(1, "GetChannelAcqPropsByName", "Could not find a channel with the given name")
End Function

Function GetWindowFunction(oFile As PolyFile, oChannel As Channel, Params() As Double) As PTCWindowFunction
	' check the acquisition properties of the given channel for the window function and its parameters

	Dim oChannelProps As ChannelAcqPropertiesContainer
	Set oChannelProps = GetChannelAcqPropsByName(oFile.Infos.AcquisitionInfoModes.ActiveProperties.ChannelsProperties, oChannel.Name)

	GetWindowFunction = oChannelProps.WindowFunction
	Params = oChannelProps.WindowFunctionParams

End Function


' *******************************************************************************
' * Helper functions and subroutines
' *******************************************************************************

Const c_OFN_HIDEREADONLY As Long = 4

Private Function FileOpenDialog() As String
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
	' instantiate PolyFile object, open the File
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
