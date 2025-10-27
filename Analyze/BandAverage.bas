'#Reference {E68EA160-8AD4-11D3-8F08-00104BB924B2}#1.0#0#C:\Program Files\Common Files\Polytec\COM\SignalDescription.dll#Polytec SignalDescription Type Library, $Revision: 3$ UnicodeRelease, Build on Jul 13 2005 at 21:23:16
'#Reference {CE68D434-5052-431F-BE75-F3C23458127A}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolySignal.dll#Polytec PolySignal Type Library, $Revision: 3$ UnicodeRelease, Build on Jul 13 2005 at 21:23:17
'#Reference {F08ACE20-C7AD-46CA-8001-D5158D9B0224}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolyMath.dll#Polytec PolyMath Type Library, $Revision: 3$ UnicodeRelease, Build on Jul 13 2005 at 21:23:21
'#Reference {A65C100F-C1FE-4C3B-9C43-46F4FB4C3BC3}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolyFile.dll#Polytec PolyFile Type Library, $Revision: 3$ UnicodeRelease, Build on Jul 13 2005 at 21:34:50
' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This Macro calculates the magnitude average over all bands and
' creates a user defined dataset in the point average domain,
' that displays the minimum, maximum and average velocity magnitudes
' of the vibrometer channel of all bands over the measurement point index.
'
' This macro can be used to find a good reference point for
' feeding a driving force into the object. You can easily find
' a point that has no small magnitudes in all bands (highest minimum
' magnitude) and therefore does not lie in a vibration node in all
' bands.
'
' When running the macro you are asked to navigate to a PSV scan file.
' This file has to meet the following conditions:
'
' - bands have to be defined in the file
' - you have to have exclusive write access to the file. We strongly recommend to
'   use a backup copy of your original file with this macro. The macro
'   will fail if the file is open in PSV.
' - the measurement has to be done in FFT mode and the vibrometer
'   channel has to offer a velocity signal
'
' To display the calculated data in PSV do the following:
' - start PSV and open the scan file
' - select Presentation/View/Average Spectrum
' - select Presentation/Channel/Usr
' - select Presentation/Signal/Magnitude Average over all Bands
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

	If (Not CheckLabelsEqualIndices(oFile)) Then
		If (MsgBox("Please note that the point indices of the x-axis of the signals " + _
			"produced by this macro will not correspond to the point indices shown in the area view " + _
			"because the indices were modified. " + _
			"Do you want to continue?", vbYesNo) = vbNo) Then
			oFile.Close()
			Exit Sub
		End If
	End If

	Dim BandsMin() As Single
	Dim BandsMax() As Single
	Dim BandsAvg() As Single
	Dim BandsGeoAvg() As Single

	Dim oBandDomains As BandDomains
	Set oBandDomains = oFile.GetBandDomains(ptcBuildBandData3d)

	Dim oBandDomain As BandDomain
	Set oBandDomain = oBandDomains.type(ptcDomainSpectrum)

	Dim oDomainBand As Domain
	Set oDomainBand = oBandDomain

	Dim oSignalBand As Signal
    If (Not oDomainBand.Channels(1).Signals.Exists("Velocity")) Then
		MsgBox("Signal velocity does not exist, macro exits now.", vbOkOnly)
		oFile.Close()
		Exit Sub
    End If
	Set oSignalBand = oDomainBand.Channels(1).Signals("Velocity")

	Dim oDisplayBand As Display
	Set oDisplayBand = oSignalBand.Displays.type(ptcDisplayMag)

	' calculate the average over all defined bands
	Call CalcBandAverage(oBandDomain, oDisplayBand, BandsMin, BandsMax, BandsAvg, BandsGeoAvg)

	Dim oDataBand As DataBand
	Set oDataBand = oBandDomain.GetDataBands(oSignalBand).Item(1)

	' create the new signal in the point average domain
	Dim oAverageDomains As PointAverageDomains
	Set oAverageDomains = oFile.GetPointAverageDomains(ptcBuildPointData3d)

	Dim oAverageDomain As PointAverageDomain
	Set oAverageDomain = oAverageDomains.type(ptcDomainSpectrum)
	
	Dim oDomainAverage As Domain 
	Set oDomainAverage = oAverageDomain

	Dim oSignalAverage As Signal
	Set oSignalAverage = oDomainAverage.Channels(1).Signals("Velocity")
	
	' band domains have no YAxis objects, therefore we use the YAxis of the point average domain, which
	' is equivalent
	Dim oYAxis As YAxis
	Set oYAxis = oAverageDomain.GetYAxes(oSignalAverage.Displays.type(ptcDisplayMag)).Item(1)

	' set the properties of the user defined signal
	Dim oUsrSigDesc As SignalDescription
	Set oUsrSigDesc = oSignalAverage.Description.Clone()
	With oUsrSigDesc
		.DataType = ptcDataAverage
		.DomainType = ptcDomainSpectrum
		.Name = "Magnitude Average over all Bands. Frame 1 is Min, Frame 2 is Geometric Average, Frame 3 is Average, Frame 4 is Max"
		.Complex = False
		.PowerSignal = False
		.XAxis.MaxCount = UBound(BandsAvg) - LBound(BandsAvg) + 1
		.XAxis.Min = 1
		.XAxis.Max = .XAxis.MaxCount
		.XAxis.Name = "Measurement Point"
		.XAxis.Unit = "Index"

		.YAxis.Name = oYAxis.Name
		.YAxis.Min = oYAxis.Min
		.YAxis.Max = oYAxis.Max
		.YAxis.Unit = oYAxis.Unit

		.ResponseDOFs.Assign(0, ptcPlusZTranslation, "Usr", oYAxis.Name, oYAxis.Unit)
	End With

	Dim oUsrSignal As Signal
	Set oUsrSignal = oAverageDomains.FindSignal(oUsrSigDesc, True)
	
	' check if a signal with the same name exits already
	If (oUsrSignal Is Nothing) Then
		Set oUsrSignal = oAverageDomains.AddSignal(oUsrSigDesc)
	Else
		If (MsgBox("A user defined signal with the name '" + oUsrSigDesc.Name + "' exists already. Do you want to replace it?", vbYesNo) = vbNo) Then
			oFile.Close
			Exit Sub
		End If
		oUsrSignal.Channel.Signals.Update(oUsrSignal.Name, oUsrSigDesc)
	End If

	' now add the data
	oAverageDomain.SetData(oUsrSignal, 1, BandsMin)
	oAverageDomain.SetData(oUsrSignal, 2, BandsGeoAvg)
	oAverageDomain.SetData(oUsrSignal, 3, BandsAvg)
	oAverageDomain.SetData(oUsrSignal, 4, BandsMax)

	oFile.Save()
	oFile.Close()

	MsgBox("Macro has finished.", vbOkOnly)
End Sub

Function CheckLabelsEqualIndices(oFile As PolyFile)
' -------------------------------------------------------------------------------
' Checks if the labels of the scan points are equal to their indices
' -------------------------------------------------------------------------------
	Dim oMeasPoints As MeasPoints
	Set oMeasPoints = oFile.Infos.MeasPoints

	Dim oMeasPoint As MeasPoint
	Dim iIndex As Long
	iIndex = 1
	For Each oMeasPoint In oMeasPoints
		If (oMeasPoint.Label <> iIndex) Then
			CheckLabelsEqualIndices = False
			Exit Function
		End If
		iIndex = iIndex + 1
	Next
	CheckLabelsEqualIndices = True
End Function


Sub CalcBandAverage(oBandDomain As BandDomain, oDisplay As Display, BandsMin() As Single, BandsMax() As Single, BandsAvg() As Single, BandsGeoAvg() As Single)
' -------------------------------------------------------------------------------
' Calculates the min, max and average of the magnitudes of the vib channel over the point index.
' -------------------------------------------------------------------------------
	Dim oStat As New Statistics
	Dim oGeoStat As New Statistics
	Dim i As Long
	Dim data() As Single
	Dim logData() As Double

	Dim oChannel As Channel
	Set oChannel = oDisplay.Signal.Channel
	Dim is3D As Boolean
	is3D = ((oChannel.Caps And ptcChannelCapsVector) <> 0)

	Dim oDataBand As DataBand
	For Each oDataBand In oBandDomain.GetDataBands(oDisplay.Signal)
		data = oDataBand.GetData(oDisplay, 0)
		If is3D Then
			data = GetMagnitudeOf3DData(data)
		End If
		oStat.Add(data)
		ReDim logData(LBound(data) To UBound(data))
		For i = LBound(data) To UBound(data)
			logData(i) = Log(data(i))
		Next i
		oGeoStat.Add(logData)
	Next oDataBand

	BandsMin = oStat.Min
	BandsMax = oStat.Max
	BandsAvg = oStat.Mean

	logData = oGeoStat.Mean
	ReDim BandsGeoAvg(LBound(logData) To UBound(logData))
	For i = LBound(BandsGeoAvg) To UBound(BandsGeoAvg)
		' very small magnitudes or magnitudes of 0 could lead to an overflow error when converting the
		' double value to the single value. We handle this by ignoring the overflow and initializing
		' the result with 0. Not that Exp(Log(0)) does not give an error but -1.#IND# as result.
		BandsGeoAvg(i) = 0.0
		On Error Resume Next
		BandsGeoAvg(i) = CSng(Exp(logData(i)))
		On Error GoTo 0
	Next i
End Sub

Function GetMagnitudeOf3DData(data() As Single) As Single()
' -------------------------------------------------------------------------------
' Gets Sqrt(X^2 + Y^2 + Z^2) of 3D Data
' -------------------------------------------------------------------------------
	Dim result() As Single
	Dim count As Long
	count = (UBound(data) - LBound(data) + 1) / 3

	ReDim result(0 To count - 1)

	Dim i As Long
	For i = 0 To count - 1
		result(i) = Sqr(data(i) * data(i) + data(i + count) * data(i + count) + data(i + 2 * count) * data(i + 2 * count))
	Next i
	GetMagnitudeOf3DData = result
End Function

Function Get0DbReference(oFile As PolyFile, oSignal As Signal) As Double
' -------------------------------------------------------------------------------
' Gets the 0 dB value of the given signal.
' -------------------------------------------------------------------------------
	Dim oDbReferences As DbReferences
	If (oFile.UseGlobalDbReferences) Then
		Set oDbReferences = oFile.Preferences.DbReferences
	Else
		Set oDbReferences = oFile.Infos.DbReferences
	End If

	Get0DbReference = oSignal.Get0dB(oDbReferences)
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
