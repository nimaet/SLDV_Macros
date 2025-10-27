' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
'
' This macro demonstrates how to modify the 3D coordinates of scan points
' in a settings file. As an example a constant factor is multiplied to
' the X, Y, Z coordinates of all scan points with valid coordinates.
' Please note, that all scaled scan points will receive the geometry status
' modified.
'
' When running the macro you are asked to navigate to a PSV settings file.
' This file has to meet the following conditions:
'
' - you have to have exclusive write access to the file. We strongly recommend to
'   use a backup copy of your original file with this macro.
' - the settings file has to have a scan point definition with a 3D geometry.
' - the settings file has to have a scan point definition, that was defined in APS point mode
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

'#Uses "..\SwitchToAcquisitionMode.bas"

Option Explicit

Const c_strFileFilter As String = "Settings File (*.set)|*.set|All Files (*.*)|*.*||"
Const c_strFileExt As String = "set"

Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	If Not SwitchToAcquisitionMode() Then
		MsgBox("Switch to acquisition mode failed.", vbOkOnly)
		End
	End If

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

	' we have to open the file for read/write, otherwise we cannot save the modified coordinates
	If Not OpenFile(oFile, strFileName) Then
		Exit Sub
	End If

	If (Not oFile.Infos.HasMeasPoints) Then
		MsgBox("The settings file does not contain a scan point definition." + vbCrLf + _
			   "The macro will finish now.", vbExclamation Or vbOkOnly, "Scale Geometry")
		oFile.Close()
		Exit Sub
	End If

	Dim oMeasPoints As MeasPoints
	Set oMeasPoints = oFile.Infos.MeasPoints
	If (Not oMeasPoints.Has3DCoordinates) Then
		MsgBox("The settings file does not contain a 3D geometry." + vbCrLf + _
			   "Scaling is only possible for 3D geometries." + vbCrLf + _
			   "The macro will finish now.", vbExclamation Or vbOkOnly, "Scale Geometry")
		oFile.Close()
		Exit Sub
	End If

	Dim strFactor As String
	strFactor = InputBox("Please enter the scale factor." + vbCrLf + _
					     "This factor will by multiplied with"  + vbCrLf + _
					     "the X, Y, Z coordinates of all scan" + vbCrLf + _
					     "points with valid coordinates", "Scale Factor", "1")

	If (strFactor = "") Then
		oFile.Close()
		Exit Sub
	End If

	Dim factor As Double
	On Error Resume Next
	factor = CDbl(strFactor)
	If Err.Number <> 0 Then
		MsgBox("Please enter a valid scale factor. The macro will finish now", vbExclamation, "Scale Geometry")
		oFile.Close()
		Exit Sub
	End If
	On Error GoTo 0

	Dim measPointCount As Long
	measPointCount = ScaleCoordinates(oMeasPoints, CDbl(strFactor))
	Dim totalPointCount As Long
	totalPointCount = oMeasPoints.Count

	oFile.Save()
	oFile.Close()

	MsgBox("Macro has finished. Scaled " + CStr(measPointCount) + " of " + CStr(totalPointCount) + " scan points", vbOkOnly)
End Sub

Function ScaleCoordinates(oMeasPoints As MeasPoints, factor As Double) As Long
' -------------------------------------------------------------------------------
' Scale all scan points with valid coordinates by the given factor.
' Returns the number of points scaled.
' -------------------------------------------------------------------------------
	Dim measPointCount As Long
	Dim oMeasPoint As MeasPoint
	For Each oMeasPoint In oMeasPoints
		If ((oMeasPoint.GeometryStatus And ptcGeoStatusValid) = ptcGeoStatusValid) Then
			Dim x As Double
			Dim y As Double
			Dim z As Double
			oMeasPoint.CoordXYZ(x, y, z)
			x = x * factor
			y = y * factor
			z = z * factor
			oMeasPoint.SetCoordXYZ(x, y, z)
			measPointCount = measPointCount + 1
		End If
	Next
	ScaleCoordinates = measPointCount
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
