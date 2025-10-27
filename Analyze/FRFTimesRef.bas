'#Reference {E44752C9-2D41-48A6-9B74-66D5B7505325}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolyProperties.dll#Polytec PolyProperties Type Library, $Revision: 3$ UnicodeRelease, Build on Jul 13 2005 at 21:23:19
'#Reference {A65C100F-C1FE-4C3B-9C43-46F4FB4C3BC3}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolyFile.dll#Polytec PolyFile Type Library, $Revision: 3$ UnicodeRelease, Build on Jul 13 2005 at 21:34:50
'#Reference {F08ACE20-C7AD-46CA-8001-D5158D9B0224}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolyMath.dll#Polytec PolyMath Type Library, $Revision: 3$ UnicodeRelease, Build on Jul 13 2005 at 21:23:21
'#Reference {E68EA160-8AD4-11D3-8F08-00104BB924B2}#1.0#0#C:\Program Files\Common Files\Polytec\COM\SignalDescription.dll#Polytec SignalDescription Type Library, $Revision: 3$ UnicodeRelease, Build on Jul 13 2005 at 21:23:16
'#Reference {CE68D434-5052-431F-BE75-F3C23458127A}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolySignal.dll#Polytec PolySignal Type Library, $Revision: 3$ UnicodeRelease, Build on Jul 13 2005 at 21:23:17
' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
'
' This macro multiplies the H1 signals of a PSV scan file
' at all measurement points with the corresponding reference signal
' averaged over all measurement points.
'
' If you have e.g. an H1 Vib Velocity / Ref1 Force signal,
' this H1 will be multiplied by the Ref1 Force average over all
' measurement points. The unit of the new signal will be m/s compared
' to m/s / N for the original signal.
'
' This allows for displaying the data in the units of the response
' and adds the additional benefit that variations of the reference
' from scan point to scan point are compensated by multiplying
' with the averaged reference spectrum.
'
' All results will be stored as user defined datasets.
'
' When running the macro you are asked to navigate to a PSV scan file.
' This file has to meet the following conditions:
'
' - you have to have exclusive write access to the file. We strongly recommend to
'   use a backup copy of your original file with this macro. The macro
'   will fail if the file is open in PSV.
' - the measurement has to be done in FFT mode
'
' To display the calculated data in PSV do the following:
' - start PSV and open the scan file
' - select Presentation/View/Average Spectrum or Single Scanpoint
' - select Presentation/Channel/Usr
' - select Presentation/Signal/...
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

Const c_strFileFilter As String = "Scan File (*.svd)|*.svd|All Files (*.*)|*.*||"
Const c_strFileExt As String = "svd"

Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	Dim oFile  As PolyFile

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

    ' provide access to the acquisition properties
    Dim oAcqInfoModes As AcquisitionInfoModes
    Set oAcqInfoModes = oFile.Infos.AcquisitionInfoModes

    ' check the acquisition mode - should be FFT
    Dim AcqMode As PTCAcqMode
    AcqMode = oAcqInfoModes.ActiveMode
    If (AcqMode <> ptcAcqModeFft) Then
        MsgBox("Please select a file with acquisition mode FFT", vbExclamation)
    	oFile.Close()
        Exit Sub
    End If

    Dim oChannelsAcqProps As ChannelsAcqPropertiesContainer
    Set oChannelsAcqProps = oAcqInfoModes.ActiveProperties.ChannelsProperties

    ' set source for data to channel and signal
	Dim oAverageDomains As PointAverageDomains
	Set oAverageDomains = oFile.GetPointAverageDomains(ptcBuildPointData3d)

	Dim oAverageDomain As PointAverageDomain
	Set oAverageDomain = oAverageDomains.type(ptcDomainSpectrum)

	' create the new signal in the point domain
	Dim oPointDomains As PointDomains
	Set oPointDomains = oFile.GetPointDomains(ptcBuildPointData3d)

	Dim oPointDomain As PointDomain
	Set oPointDomain = oPointDomains.type(ptcDomainSpectrum)

    Dim oVector As New Vector

	Dim oChannel As Channel

	' loop over all channels and signals of the domain
	For Each oChannel In oPointDomain.Channels

		If (oChannel.Caps And ptcChannelCapsUser) = 0 Then

			Dim sChannel1 As String
			Dim sChannel2 As String
			Dim iPos As Integer

			iPos = InStr(oChannel.Name, "&")
			If iPos > 0 Then

				sChannel1 = Trim(Mid$(oChannel.Name,1,iPos - 2))
				sChannel2 = Trim(Mid$(oChannel.Name,iPos + 1))

				Dim oSignal As Signal

				For Each oSignal In oChannel.Signals

					If (oSignal.Description.FunctionType = ptcFunctionFrequencyResponseFunctionH1Type) Then

						Dim sSignal1 As String
						Dim sSignal2 As String

						iPos = InStr(oSignal.Name, "/")

						If iPos > 0 Then

							sSignal1 = Trim(Mid$(oSignal.Name,4,iPos - 4))
							sSignal2 = Trim(Mid$(oSignal.Name,iPos + 1))

							Dim oSignalAvg As Signal
							Dim oDomainAvg As Domain
							Set oDomainAvg = oAverageDomains.type(ptcDomainSpectrum)
							Dim oChannelAvg As Channel
							Set oChannelAvg = oDomainAvg.Channels(sChannel2)
							Set oSignalAvg = oChannelAvg.Signals(sSignal2)

							Dim oDomainPoint As Domain
							Set oDomainPoint = oPointDomain

							Dim oChannelPoint As Channel
							Set oChannelPoint = oDomainPoint.Channels(sChannel1)

							Dim bIs3D As Boolean
							bIs3D = (oChannelPoint.Caps And ptcChannelCapsVector) <> 0

							Dim strSignalName As String
							strSignalName = sChannel1 & " " & sSignal1 & " ( H1 * Average " & sChannel2 & " )"

							Dim oUsrSignal As Signal
							Set oUsrSignal = AddUserSignal(oPointDomains, oChannelPoint.Signals(sSignal1).Description, strSignalName, GetChannelAcqPropsByName(oChannelsAcqProps, sChannel1), bIs3D)

						    Dim asglAvg() As Single
						    asglAvg = oAverageDomain.GetData(oSignalAvg.Displays(ptcDisplayMag), 0)
						    If bIs3D Then
						    	' in case of 3D signals we have to expand the reference signal
						    	' such that X Y and Z of the response signal are multiplied with
						    	' the same reference signal
						    	Dim asglAvg3D() As Single
						    	asglAvg3D = oVector.Append(asglAvg, asglAvg)
						    	asglAvg3D = oVector.Append(asglAvg3D, asglAvg)
						    	asglAvg = asglAvg3D
						    End If

						    Dim oDisplay As Display
						    Set oDisplay = oSignal.Displays(ptcDisplayRealImag)

							If (Not oUsrSignal Is Nothing) Then
								Dim oDataPoint As DataPoint
							    For Each oDataPoint In oPointDomain.DataPoints
							        ' check if measurement point is valid
							        ' test the valid flag of the point status
									Dim bValid As Boolean
							        bValid = (oDataPoint.GetScanStatus(oDisplay) And ptcScanStatusValid) <> 0
							        If (bValid) Then
										' multiply the point data with the (real) point average
							            Dim asglData() As Single
										asglData = oDataPoint.GetData(oDisplay, 0)
										Dim asglResult() As Single
										asglResult = oVector.MulCplx(asglData, oVector.Complex(asglAvg))

										oDataPoint.SetData(oUsrSignal, 1, asglResult)
							        End If
							    Next oDataPoint
							End If
						End If
					End If
				Next oSignal
			End If
		End If
	Next oChannel

	oFile.Save()
	oFile.Close()

	MsgBox("Macro has finished.", vbOkOnly)
End Sub

Function AddUserSignal(oPointDomains As PointDomains, oSignalDescription As SignalDescription, strSignalName As String, oChannelAcqProps As ChannelAcqPropertiesContainer, bIs3D As Boolean) As Signal
' -------------------------------------------------------------------------------
' Create the new signal in the point domain.
' -------------------------------------------------------------------------------
	Set AddUserSignal = Nothing

	Dim oUsrChannel As Channel

	' set the properties of the user defined signal
	Dim oUsrSigDesc As SignalDescription
	Set oUsrSigDesc = oSignalDescription.Clone()

	With oUsrSigDesc
		.Name = strSignalName
		.Complex = True
		.PowerSignal = False
	End With

	Dim oUsrSignal As Signal
	Set oUsrSignal = Nothing

	Set oUsrSignal = oPointDomains.FindSignal(oUsrSigDesc, True)

	' check if a signal with the same name exits already
	If (oUsrSignal Is Nothing) Then
		Set oUsrSignal = oPointDomains.AddSignal(oUsrSigDesc)
	Else
		If (MsgBox("A user defined signal with the name '" + oUsrSigDesc.Name + "' exists already. Do you want to replace it?", vbYesNo) = vbNo) Then
			Exit Function
		End If

		oUsrSignal.Channel.Signals.Update(oUsrSignal.Name, oUsrSigDesc)
	End If

	Set AddUserSignal = oUsrSignal

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
	FileOpenDialog = GetFilePath(, c_strFileExt, CurDir(), "Select a file", 0)
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
