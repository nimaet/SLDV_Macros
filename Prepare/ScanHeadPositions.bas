' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
'
' In acquisiton mode this macro displays the current scan head positions
' and the distances between the scanheads (for PSV-3D only).
'
' In presentation mode the macro asks for a file Name of a Polytec .svd
' scan file or settings file, opens the file and displays the positions
' of the scan heads in this file.
'
' The positions are determined in the 3D alignment procedure.
' The positions are the points where the laser beam exits the scan head
' (outer side of the front plate) when the beam is positioned at 0°/0°
'
' The file has to be a file with 3D geometry, i.e. a file with
' 3D alignment.
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
' ----------------------------------------------------------------------

Option Explicit

Const c_strFileFilter As String = "Scan File (*.svd)|*.svd|Settings File (*.set)|*.set|All Files (*.*)|*.*||"
Const c_strFileExt As String = ".svd"

Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	Dim oFile As PolyFile
	Dim oScanHeadInfo As ScanHeadDevicesInfo
	Dim oAligns As Alignments3D
	Call GetInfos(oScanHeadInfo, oAligns, oFile)

	If (oScanHeadInfo Is Nothing Or oAligns Is Nothing) Then
		If (Not oFile Is Nothing) Then
			oFile.Close()
		End If
		Exit Sub
	End If

	Dim pos As String
	Dim names(1 To 3) As String
	names(1) = "Top:   "
	names(2) = "Left:  "
	names(3) = "Right: "
	Dim head As Long
	head = 1

	Dim xTop As Double,   yTop As Double,   zTop As Double
	Dim xLeft As Double,  yLeft As Double,  zLeft As Double
	Dim xRight As Double, yRight As Double, zRight As Double

	Dim oAlign3D As Alignment3D
	For Each oAlign3D In oAligns
		If oAlign3D.Valid Then
			Dim oScanner As ScanHeadScanner
			Set oScanner = oScanHeadInfo.ScanHeadDevices(head).Scanner
			Dim xOrigin As Double
			Dim yOrigin As Double
			Dim zOrigin As Double
			oScanner.GetLaserOrigin(xOrigin, yOrigin, zOrigin)

			Dim x As Double
			Dim y As Double
			Dim z As Double
			oAlign3D.ScannerToCoord3D(0.0, 0.0, Abs(zOrigin), x, y, z)
			pos = pos + names(head) + vbTab + vbTab + Format$(x, "0.000") + " m   "
			pos = pos + vbTab + Format$(y, "0.000") + " m   "
			pos = pos + vbTab + Format$(z, "0.000") + " m" + vbCrLf
			If head = 1 Then
				xTop = x
				yTop = y
				zTop = z
			ElseIf head = 2 Then
				xLeft = x
				yLeft = y
				zLeft = z
			Else
				xRight = x
				yRight = y
				zRight = z
			End If
		Else
			pos = pos + names(head) + vbTab + vbTab + "not valid" + vbCrLf
		End If
		head = head + 1
	Next

	If Len(pos) = 0 Then
		MsgBox("The file or settings do not contain a valid 3D alignment.", vbOkOnly)
		If (Not oFile Is Nothing) Then
			oFile.Close()
		End If
		Exit Sub
	End If

	pos = "Scanning head positions (x,y,z) as defined by the 3D alignment: " + vbCrLf + vbCrLf + pos

	' for PSV-3D only (3 scan heads) calculate the distances between the scan heads
	If (head = 4) Then			' PSV-3D
		Dim topLeft As Double, topRight As Double, leftRight As Double

		topLeft   = Sqr((xTop  - xLeft )^2 + (yTop  - yLeft )^2 + (zTop  - zLeft )^2)
		topRight  = Sqr((xTop  - xRight)^2 + (yTop  - yRight)^2 + (zTop  - zRight)^2)
		leftRight = Sqr((xLeft - xRight)^2 + (yLeft - yRight)^2 + (zLeft - zRight)^2)

		pos = pos + vbCrLf + "Distances between scanning heads: " + vbCrLf + vbCrLf
		pos = pos + "Top - Left:  "   + vbTab + Format$(topLeft,    "0.000") + " m" + vbCrLf
		pos = pos + "Top - Right: "   + vbTab + Format$(topRight,   "0.000") + " m" + vbCrLf
		pos = pos + "Left - Right:"   + vbTab + Format$(leftRight,  "0.000") + " m" + vbCrLf
	End If

	' distance of scan head(s) to origin
	Dim TopOrigin As Double, LeftOrigin As Double, RightOrigin As Double
	If (head = 4) Then			' PSV-3D
		pos = pos + vbCrLf + "Distances to origin:" + vbCrLf + vbCrLf
	Else
		pos = pos + vbCrLf + "Distance to origin:" + vbCrLf + vbCrLf
	End If

	TopOrigin = Sqr(xTop^2 + yTop^2 + zTop^2)
	pos = pos + "Top - Origin:"       + vbTab + Format$(TopOrigin,  "0.000") + " m" + vbCrLf
	If (head = 4) Then			' PSV-3D
		LeftOrigin = Sqr(xLeft^2 + yLeft^2 + zLeft^2)
		RightOrigin = Sqr(xRight^2 + yRight^2 + zRight^2)
		pos = pos + "Left - Origin: " + vbTab + Format$(LeftOrigin, "0.000") + " m" + vbCrLf
		pos = pos + "Right - Origin:" + vbTab + Format$(RightOrigin,"0.000") + " m" + vbCrLf
	End If

	pos = pos + vbCrLf + "The positions are the points where the laser beams exit the scanning heads" + vbCrLf
	pos = pos + "(outer side of the front panels) when the beams are positioned at 0°/0°." + vbCrLf + vbCrLf
	pos = pos + "Note: You can copy this message to the clipboard by pressing Ctrl+C."

	MsgBox(pos, vbOkOnly)

	If (Not oFile Is Nothing) Then
		oFile.Close()
	End If
End Sub

' *******************************************************************************
' * Gets the infos from the file or the current acquisition settings
' *******************************************************************************
Private Sub GetInfos(oScanHeadInfo As ScanHeadDevicesInfo, oAligns As Alignments3D, oFile As PolyFile)
	If (Application.Mode = ptcApplicationModeAcquisition) Then
		Dim oInfosAcq As InfosAcq
		Set oInfosAcq = Acquisition.Infos
		If (Not oInfosAcq.HasAlignments) Then
			MsgBox("No alignment information found in the settings.", vbOkOnly)
			Exit Sub
		End If
		Set oScanHeadInfo = oInfosAcq.ScanHeadDevicesInfo
		Set oAligns = oInfosAcq.Alignments.Alignments3D
	Else
		' get filename and path
		Dim strFileName As String
		strFileName = FileOpenDialog()

		If (strFileName = "") Then
			MsgBox("No filename has been specified, macro exits now.", vbOkOnly)
			Exit Sub
		End If

		Set oFile = New PolyFile
		oFile.Open(strFileName)

		Dim oInfos As Infos
		Set oInfos = oFile.Infos
		If (Not oInfos.HasAlignments) Then
			MsgBox("No alignment information found in the file.", vbOkOnly)
			Exit Sub
		End If

		Set oAligns = oInfos.Alignments.Alignments3D
		Set oScanHeadInfo = oInfos.ScanHeadDevicesInfo
	End If

	If (oAligns.Count = 0) Then
		MsgBox("The file or the settings do not contain a 3D alignment.", vbOkOnly)
		Set oAligns = Nothing
		Set oScanHeadInfo = Nothing
		Exit Sub
	End If
End Sub


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
	FileOpenDialog = GetFilePath(, Right$(c_strFileExt, 3), CurDir(), "Select a file", 0)
MEnd:
End Function
