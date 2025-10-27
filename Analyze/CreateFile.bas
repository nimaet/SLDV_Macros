'#Reference #System.IO, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL
' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This macro creates a PSV .svd scan file.
' A disk geometry (points and elements) is created and
' sine curves with different phases are added as user defined
' dataset "Sine". Optionally the macro copies an image to the svd scan file.
' The macro opens the file in PSV presentation mode.
' You can display the geometry and the data.

' References
' - Polytec_IO_Vibrometer
' - Polytec PolyAcquisition Type Library
' - Polytec PolyAlignment Type Library
' - Polytec PolyScanHead Type Library
' - Polytec PolyFrontEnd Type Library
' - Polytec PolyMath Type Library
' - Polytec PolyFile Type Library
' - Polytec PolyProperties Type Library
' - Polytec PolyDataVisualizer Type Library
' - Polytec PolyScope Type Library
' - Polytec PolyGenerators Type Library
' - Polytec PolyWaveforms Type Library
' - Polytec PolyDigitalFilters Type Library
' - Polytec WindowFunction Type Library
' - Polytec PolySignal Type Library
' - Polytec SignalDescription Type Library
' - Polytec PhyscicalUnit Type Library
' ----------------------------------------------------------------------

'#Language "WWB.NET"

Option Explicit

Imports System
Imports System.IO

Const c_strSvdFileFilter As String = "Scan File (*.svd)|*.svd|All Files (*.*)|*.*||"
Const c_strSvdFileExt As String = "svd"

Const c_strImageFileFilter As String = "Bitmap File (*.svd)|*.svd|(*.bmp)|*.bmp|JPG (*.jpg)|*.jpg|TIFF (*.tif)|*.tif|PNG (*.png)|*.png|GIF (*.gif)|*.gif|All Files (*.*)|*.*||"
Const c_strImageFileExt As String = "svd;bmp;jpg;tif;png;gif"

Const c_dPi As Double = Atn(1.0) * 4.0
Const c_lDiskSectors As Long = 32
Const c_lSamples As Long = 64
Const c_lVibrationNodes As Long = 2

Sub Main
	' get filename and path
	Dim strFileName As String
	strFileName = FileOpenDialog("Select a file", c_strSvdFileFilter, c_strSvdFileExt)

	If (strFileName = "") Then
		MsgBox("No filename has been specified, macro exits now.", vbOkOnly)
		Exit Sub
	End If

	If (Dir(strFileName) <> "") Then
		If (MsgBox("The macro will overwrite the file '" + strFileName + _
			"'." + "Do you want to continue?", vbYesNo) = vbNo) Then
			Exit Sub
		End If
	End If

	Dim oFile As New PolyFile
	oFile.ReadOnly = False

	oFile.Create(strFileName, POLYFILELib.ptcCreateAlways)
	oFile.Version.FileID = POLYFILELib.ptcFileIDPSVFile

	Call CreateVideoBitmap(oFile)
	Call CreateMeasPointsAndElements(oFile)
	Call CreateData(oFile)

	oFile.Save()
	oFile.Close()

	Documents.Open(strFileName)
End Sub


Sub CreateVideoBitmap(oFile As PolyFile)
	' get filename and path for existing image
	Dim strImageFileName As String
	strImageFileName = FileOpenDialog("Select a file", c_strImageFileFilter, c_strImageFileExt)

	If (strImageFileName <> "") Then

		Dim oVideoBitmap As VideoBitmap
		oVideoBitmap = oFile.Infos.Add(POLYSIGNALLib.ptcInfoVideoBitmap)

		If (System.IO.Path.GetExtension(strImageFileName) = ".svd") Then

			Dim oImageFile As New PolyFile
			oImageFile.Open(strImageFileName)

			Dim oBitmap() As Byte
			oBitmap = oImageFile.Infos.VideoBitmap.Image(POLYFILELib.ptcGraphicFormatJPEG, 0, 0)
			oVideoBitmap.SetImage(oBitmap)

			oImageFile.Close()

		Else

			oVideoBitmap.LoadImageFile(strImageFileName)

		End If

	End If
End Sub


Sub CreateMeasPointsAndElements(oFile As PolyFile)
	Dim x As Double
	Dim y As Double

	Dim oMeasPoints As MeasPoints
	oMeasPoints = oFile.Infos.Add(POLYSIGNALLib.ptcInfoMeasPoints)

	Dim oElements As Elements
	oElements = oFile.Infos.Add(POLYSIGNALLib.ptcInfoElements)

	Dim oMeasPoint As MeasPoint
	oMeasPoint = oMeasPoints.Add()
	oMeasPoint.SetCoordXYZ(0.0, 0.0, 0.0)

	Dim offsetX As Double
	offsetX = 2.0/3.0

	Dim offsetY As Double
	offsetY = 0.5

	oMeasPoint.SetVideoXY(offsetX, offsetY)

	Dim i As Integer
	For i = 1 To c_lDiskSectors
		x = Sin(i * 2 * c_dPi / c_lDiskSectors)
		y = Cos(i * 2 * c_dPi / c_lDiskSectors)
		oMeasPoint = oMeasPoints.Add()
		oMeasPoint.SetCoordXYZ(x, y, 0.0)
		oMeasPoint.SetVideoXY(0.5 * x + offsetX, - 0.5 * y + offsetY)
	Next i

	Dim elementIndices(0 To 2) As Integer
	For i = 1 To c_lDiskSectors
		elementIndices(0) = 1
		elementIndices(1) = i + 1
		If (i < c_lDiskSectors) Then
			elementIndices(2) = i + 2
		Else
			elementIndices(2) = 2
		End If

		oElements.Add(elementIndices, False)
	Next i
End Sub


Sub CreateData(oFile As PolyFile)
	Dim oUSD As New SignalDescription
	oUSD.Name = "Sine"
	oUSD.DataType = POLYSIGNALLib.ptcDataPoint
	oUSD.DomainType = POLYSIGNALLib.ptcDomainTime
	oUSD.FunctionType = POLYSIGNALLib.ptcFunctionTimeResponseType
	oUSD.XAxis.MaxCount = c_lSamples
	oUSD.XAxis.Min = 0
	oUSD.XAxis.Max = 2 * c_dPi
	oUSD.XAxis.Unit = "s"
	oUSD.XAxis.Name = "Time"
	oUSD.YAxis.Min = -1.0
	oUSD.YAxis.Max = 1.0
	oUSD.YAxis.Name = "Voltage"
	oUSD.YAxis.Unit = "V"

	oUSD.ResponseDOFs.Assign(0, POLYSIGNALLib.ptcPlusZTranslation, "Usr", "Voltage", "V")

	Dim oPointDomains As PointDomains
	oPointDomains = oFile.GetPointDomains(POLYSIGNALLib.ptcBuildPointData3d)

	Dim oSignal As Signal
	oSignal = oPointDomains.AddSignal(oUSD)

	Dim data(0 To c_lSamples-1) As Single

	Dim oPointDomain As PointDomain
	oPointDomain = oPointDomains.type(POLYSIGNALLib.ptcDomainTime)

	Dim oDataPoint As DataPoint
	oDataPoint = oPointDomain.DataPoints(1)

	oDataPoint.SetData(oSignal, 1, data)

	Dim i As Integer
	Dim j As Integer
	For i = 1 To c_lDiskSectors
		Dim phase As Double
		phase = i * c_lVibrationNodes * 2 * c_dPi / c_lDiskSectors
		For j = 0 To c_lSamples-1
			data(j) = Sin(phase + j * 2 * c_dPi / c_lSamples)
		Next j
		oDataPoint = oPointDomain.DataPoints(i + 1)
		oDataPoint.SetData(oSignal, 1, data)
	Next
End Sub


' *******************************************************************************
' * Helper functions and subroutines
' *******************************************************************************
Const c_OFN_HIDEREADONLY As Long = 4

Private Function FileOpenDialog(message As String, fileFilter As String, fileExtension As String) As String
' -------------------------------------------------------------------------------
' Select file.
' -------------------------------------------------------------------------------
	On Error GoTo MCreateError
	Dim fod As Object
	fod = CreateObject("MSComDlg.CommonDialog")
	fod.Filter = fileFilter
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
	FileOpenDialog = GetFilePath(, fileExtension, CurDir(), message, 2)
MEnd:
End Function
