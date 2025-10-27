' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This macro shows the use of user defined bands.
'
' The macro could be used for PSV 1D systems with a 3D alignment and a 3D geometry.
' It uses user defined bands to get a animation with XYZ direction.
'
' How to use this macro:
' - start PSV and open the file
' - define a band
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
'#Language "WWB-NET"

Option Explicit


Sub Main
	If (Application.Documents.Count = 0) Then
		MsgBox("No document open!")
		Exit All
	End If

	Dim oDoc As VibDocument
	oDoc = Application.ActiveDocument
	Dim  strName As String
	strName = oDoc.Name

	' get the area window from the active document
	Dim WndFound As Boolean
	WndFound = False

	Dim oAreaWindow As AreaWindow
	Dim oWnd As Window
	For Each oWnd In oDoc.Windows
		If (oWnd.type = PTCWindowType.ptcWindowTypeArea) Then
			WndFound = True
			oAreaWindow = oWnd
			Exit For
		End If
	Next

	If (WndFound = False) Then
		MsgBox("No area window!")
		Exit All
	End If

	' get active band
	Dim strBand As String
	strBand = oAreaWindow.AreaView.Settings.ActiveBand

	If (strBand = "Cursor") Then
		MsgBox("Please select a band!")
		Exit All
	End If

	Dim lBandIdx As Integer
	lBandIdx = CInt(Left$(strBand, InStr(strBand, " ") - 1))  ' get the band index

	' get domain, channel, signal and display
	Dim oDisplaySettings As DisplaySettings
	oDisplaySettings = oAreaWindow.AreaView.Settings.DisplaySettings.Clone

	If (oDisplaySettings.Channel = "Usr 3D") Then
		MsgBox("Please select a channel other than '" + oDisplaySettings.Channel + "'.")
		Exit Sub
	End If

	If (oAreaWindow.AreaView.Display.Signal.Displays.Exists(PTCDisplayType.ptcDisplayRealImag) = False) Then
		MsgBox("Please select a channel with Real & Imag. display.")
		Exit Sub
	End If

	Dim oSignal As Signal
	oSignal = oAreaWindow.AreaView.Display.Signal
	If (oSignal.Description.ResponseDOFs.Count = 0) Then
		MsgBox("Please select a response channel.")
		Exit Sub
	End If

	If (MsgBox("This macro will modify the file '" + strName + _
		"'. We strongly recommend to work with a backup copy of original data only. " + _
		"Do you want to continue?", vbYesNo) = vbNo) Then
		Exit Sub
	End If

	' save and close active document
	oDoc.Save()
	oDoc.Close()


	' ----------------------------------------------------------------------
	' open file with PolyFile
	' ----------------------------------------------------------------------
	Dim oFile As New PolyFile
	oFile.ReadOnly = False
	oFile.Open(strName)

	Dim oBandDomain As BandDomain
	oBandDomain = oFile.GetBandDomains(PTCSignalBuildFlags.ptcBuildBandData3d).type(PTCDomainType.ptcDomainSpectrum)

	Dim oSignalOrg As Signal
	oSignalOrg = oBandDomain.Channels.Item(oDisplaySettings.Channel).Signals.Item(oDisplaySettings.Signal)

	' ----------------------------------------------------------------------
	' Get the laser beam vector for X=0° and Y=0° direction from
	' the 3D alignment and calculate the normalized vector.
	' ----------------------------------------------------------------------
	Dim oAlignment3D As Alignment3D
	oAlignment3D = oFile.Infos.Alignments.Alignments3D.Item(1)

	Dim dVectorX As Double
	Dim dVectorY As Double
	Dim dVectorZ As Double

	oAlignment3D.ScannerToVector3D(0, 0, dVectorX, dVectorY, dVectorZ)

	Dim dBeamOriginX As Double
	Dim dBeamOriginY As Double
	Dim dBeamOriginZ As Double
	oAlignment3D.GetBeamOrigin(0, 0, dBeamOriginX, dBeamOriginY, dBeamOriginZ)

	dVectorX = dBeamOriginX - dVectorX
	dVectorY = dBeamOriginY - dVectorY
	dVectorZ = dBeamOriginZ - dVectorZ

	Dim dVecLength As Double
	dVecLength = Sqr(dVectorX * dVectorX + dVectorY * dVectorY + dVectorZ * dVectorZ)

	Dim dVectorNormX As Double
	Dim dVectorNormY As Double
	Dim dVectorNormZ As Double

	dVectorNormX = dVectorX / dVecLength
	dVectorNormY = dVectorY / dVecLength
	dVectorNormZ = dVectorZ / dVecLength

	' get vibration direction
	Dim ptcVibrationDir As PTCFunctionAtNodeDirection
	ptcVibrationDir = oSignalOrg.Description.ResponseDOFs.Direction

	Select Case ptcVibrationDir
		Case PTCFunctionAtNodeDirection.ptcMinusXTranslation
			dVectorNormX = -1 * dVectorNormX
		Case PTCFunctionAtNodeDirection.ptcMinusYTranslation
			dVectorNormY = -1 * dVectorNormY
		Case PTCFunctionAtNodeDirection.ptcMinusZTranslation
			dVectorNormZ = -1 * dVectorNormZ
	End Select

	Dim oDisplay As Display
	oDisplay = oSignalOrg.Displays("Real & Imag.")

	Dim oDataBands As DataBands
	oDataBands = oBandDomain.GetDataBands(oSignalOrg)

	' select the band
	Dim oDataBand As DataBand
	oDataBand = oDataBands.Item(lBandIdx)

	' get the number of frames
	Dim lFrameCount As Integer
	lFrameCount = oDataBand.GetFrames(oSignalOrg)

	' ----------------------------------------------------------------------
	' Add new user 3D band.
	' ----------------------------------------------------------------------
	Dim oBandDefUsr3D As BandDef
	oBandDefUsr3D = oBandDomain.GetDataBands(oSignalOrg).BandDefinitions(lBandIdx).Clone()

	Dim oResponse As DegreeOfFreedomID
	oResponse = oSignalOrg.Description.ResponseDOFs.Item(1)

	Dim oSignalDescUsr3D As SignalDescription
	oSignalDescUsr3D = oSignalOrg.Description.Clone() ' make a clone from the original signal description

	oSignalDescUsr3D.ResponseDOFs.Assign(oResponse.Node, PTCFunctionAtNodeDirection.ptcVector, "Usr", oResponse.Quantity, oResponse.Unit) ' make a 3D signal description

	If (oSignalDescUsr3D.ResponseDOFs.Direction <> PTCFunctionAtNodeDirection.ptcVector) Then
		MsgBox("Wrong direction!")
		Exit All
	End If

	Dim oDataBandUsr3D As DataBand
	oDataBandUsr3D = oFile.GetBandDomains.FindBand(oSignalDescUsr3D, True, oBandDefUsr3D, oSignalOrg)

	If (Not oDataBandUsr3D Is Nothing) Then
		If (MsgBox("This user defined band already exists and will be replaced with the new band. " + _
			"Do you want to continue?", vbYesNo) = vbNo) Then
			Exit Sub
		End If

		oFile.GetBandDomains.RemoveBand(oSignalOrg, oBandDefUsr3D)
	End If

	oDataBandUsr3D = oFile.GetBandDomains.AddBand(oSignalDescUsr3D, oBandDefUsr3D, oSignalOrg)  ' add the new user defined band

	' ----------------------------------------------------------------------
	' Get original band data and calculate new 3D band data
	' ----------------------------------------------------------------------

	Dim bMultiFrameMode As Boolean
	bMultiFrameMode = oFile.Infos.AcquisitionInfoModes.ActiveMode() = PTCAcqMode.ptcAcqModeMultiFrame

	Dim oVector As New Vector

	Dim lFrame As Integer
	For lFrame = 1 To lFrameCount

		' now get the data
		Dim arrBandDataOrg() As Single

		If (bMultiFrameMode = True) Then
			arrBandDataOrg = oDataBand.GetData(oDisplay, lFrame)
		Else
			arrBandDataOrg = oDataBand.GetData(oDisplay, 0)
		End If

		Dim arrScanStatus() As Integer
		arrScanStatus = oDataBand.GetScanStatus(oDisplay.Signal)


		Dim arrBandData() As Single
		Dim arrFactor(0 To 1) As Single
		arrFactor(0) = dVectorNormX
		arrBandData = oVector.MulCplx(arrBandDataOrg, arrFactor)
		arrFactor(0) = dVectorNormY
		arrBandData = oVector.Append(arrBandData, oVector.MulCplx(arrBandDataOrg, arrFactor))
		arrFactor(0) = dVectorNormZ
		arrBandData = oVector.Append(arrBandData, oVector.MulCplx(arrBandDataOrg, arrFactor))

		oDataBandUsr3D.SetData(oSignalOrg, lFrame, arrBandData)
		oDataBandUsr3D.SetScanStatus(oSignalOrg, arrScanStatus)
	Next lFrame

	oFile.Save()
	oFile.Close()

	Application.Documents.Open(strName) ' reopen the file

	For Each oWnd In Application.ActiveDocument.Windows
		If (oWnd.type = PTCWindowType.ptcWindowTypeArea) Then
			oAreaWindow = oWnd
			oAreaWindow.AreaView.Settings.DisplaySettings.Channel = "Usr 3D" ' set the "Usr 3D" channel
			Exit For
		End If
	Next

End Sub
