' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This Macro transforms the defined 3D alignment from the actual
' coordinate system to a target coordinate system.
'
' Examples for using this macro:
'   - Perform high accuracy 3D-alignment with the alignment object PSV-A-450
'   - Modify and adapt 3D-Alignment if object is moved
'   - Modify and adapt 3D-Alignment if scanning heads are moved but remain in
'     a fixed relative position (e.g. when using the tripod PSV-A-T31 or PSV-A-T34)
'
' To use this macro, please proceed as follows:
'
' 1. Perform or load a 3D alignment e.g. using a well known alignment object
'    (e.g. PSV-A-450)
' 2. Define 3-20 scan points in APS point mode. You have to know the
'    coordinates of these points in the target coordinate system. Please make sure
'    to leave the 'Define Scan Points' mode after defining the points.
' 3. Measure the 3D coordinates of these points in the actual coordinate
'    system by performing a geometry scan, a triangulation (PSV 3D only),
'    a combination of both techniques or other means.
' 4. Run this macro.
' 5. Use the previous and next buttons to navigate through the point list.
'    Then number and index of the current point and the total number of points
'    are displayed.
' 6. For every point enter the target coordinates.
' 7. When you have entered the target coordinates for at least three points,
'    you can click on "Calculate" to calculate a preliminary transformation
'    between the coordinate systems. You can review the total quality and point
'    quality before proceeding to the next point or before clicking on OK.
' 8. You can clear target coordinates by clicking on "Clear". Points without
'    target coordinates will not be taken into account when calculating
'    the transformation.
' 9. Click OK to transform the 3D alignment and the coordinates of the scan points
'    to the new coordinate system. The dialog will close.
'      - or -
'    Click Cancel to quit the macro without applying any changes.
'
' The Macro works for PSV and PSV 3D systems, it requires PSV version 8.6 or higher.
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
'-----------------------------------------------------------------------

'#Uses "..\SwitchToAcquisitionMode.bas"

Option Explicit

' data structure descibing a point with both actual and target coordinates used
' to calculate the transformation
Private Type Point
	ActualX As Double
	ActualY As Double
	ActualZ As Double
	TargetX As Double
	TargetY As Double
	TargetZ As Double
	TargetValid As Boolean
	VideoX As Single
	VideoY As Single
	Label As Long
	Index As Long
End Type

Private g_points() As Point						' global list of points
Private g_lCurrentPoint As Long					' the current point shown on the dialog
Private g_oAligns3D As Alignments3D				' the alignment used for the transformation
Private g_oScanHeadScanner As ScanHeadScanner	' the simulated MSA-E-500 scanner object

' checks the prerequisits, shows the dialog and transforms the alignments and scan points
Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	If Not SwitchToAcquisitionMode() Then
		MsgBox("Switch to acquisition mode failed.", vbOkOnly)
		End
	End If
	
	Dim oMeasPoints As MeasPoints
	Set oMeasPoints = Acquisition.Infos.MeasPoints

	If (Not oMeasPoints.Has3DCoordinates) Then
		MsgBox("The scan points must have 3D coordinates", vbExclamation, "Transform 3D Alignment")
		Exit Sub
	End If

	Dim oMeasPoint As MeasPoint
	For Each oMeasPoint In oMeasPoints
		If ((oMeasPoint.GeometryStatus And ptcGeoStatusValid) = 0) Then
			MsgBox("The scan points must have 3D coordinates", vbExclamation, "Transform 3D Alignment")
			Exit Sub
		End If
	Next

	Call ReadPoints(oMeasPoints)

	Dim oAligns As AlignmentsAcq
	Set oAligns = Acquisition.Infos.Alignments

	If (UBound(g_points) < ptcAlignment3DMinPointsObj Or UBound(g_points) > ptcAlignment3DMaxPointsObj) Then
		MsgBox("Please define between " + CStr(ptcAlignment3DMinPointsObj) + " and " + CStr(ptcAlignment3DMaxPointsObj) + " scan points.", vbExclamation, "Transform 3D Alignment")
		Exit Sub
	End If

	Call CreateAlignment()

	g_lCurrentPoint = 1
	If (ShowTransformDlg()) Then
		' PSV does not allow to modify the scan point coordinates directly via scripting
		' save the scan point definition to a temporary file
		Dim strFilePath As String
		strFilePath = Environ("temp")
		strFilePath = strFilePath + "\Transform3DAlignment.set"
		Settings.Save(strFilePath)

		Dim oFile As New PolyFile
		oFile.ReadOnly = False
		oFile.Open(strFilePath)

		Call ModifyScanPoints(oFile)
		Call Modify3DAlignments(oFile)


		oFile.Save()
		oFile.Close()

		' load settings plus alignment to avoid message boxes for recalculating point positions
		Settings.Load(strFilePath, ptcSettingsAPS Or ptcSettingsAlignment)
		Kill(strFilePath)
	End If
End Sub


' reads the currently defined scan points and stores the coordinates in g_points
Private Sub ReadPoints(oMeasPoints As MeasPoints)
	ReDim g_points(1 To oMeasPoints.Count)

	Dim oMeasPoint As MeasPoint
	For Each oMeasPoint In oMeasPoints
		Dim x As Double
		Dim y As Double
		Dim z As Double
		oMeasPoint.CoordXYZ(x, y, z)
		Dim videoX As Single
		Dim videoY As Single
		oMeasPoint.VideoXY(videoX, videoY)
		With g_points(oMeasPoint.Index)
			.ActualX = x
			.ActualY = y
			.ActualZ = z
			.VideoX = videoX
			.VideoY = videoY
			.TargetX = 0.0
			.TargetY = 0.0
			.TargetZ = 0.0
			.TargetValid = False
			.Label = oMeasPoint.Label
			.Index = oMeasPoint.Index
		End With
	Next
End Sub

' defines and shows the transform dialog, true if user clicked OK
Private Function ShowTransformDlg() As Boolean
	Begin Dialog UserDialog 740,329,"Transform 3D Alignment",.ShowTransformDlgProc ' %GRID:10,7,1,1
		Text 160,21,120,14,"Point No. (Index):"
		Text 300,21,60,14,"1 (1)",.PointNumber
		Text 380,21,40,14,"of",.OfPoints
		Text 220,56,130,14,"Actual Coordinates",.Text1
		Text 400,56,160,14,"Target Coordinates",.Text2
		Text 210,84,20,14,"X:",.XDir
		Text 240,84,70,14,"100.0",.ActualX,1
		Text 320,84,30,14,"mm"
		Text 210,112,20,14,"Y:",.YDir
		Text 240,112,70,14,"100.0",.ActualY,1
		Text 320,112,30,14,"mm"
		Text 210,140,20,14,"Z:",.ZDir
		Text 240,140,70,14,"100.0",.ActualZ,1
		Text 320,140,30,14,"mm"
		TextBox 400,82,110,16,.TargetX
		Text 530,84,50,14,"mm"
		TextBox 400,110,110,16,.TargetY
		Text 530,112,50,14,"mm"
		TextBox 400,138,110,16,.TargetZ
		Text 530,140,50,14,"mm"
		Text 160,231,100,14,"Total Quality:",.TotalQualityLabel
		Text 270,231,90,14,"1.0",.TotalQuality,1
		Text 370,231,50,14,"mm",.TotalQualityMM
		Text 160,259,100,14,"Point Quality:",.PointQualityLabel
		Text 270,259,90,14,"1.0",.PointQuality,1
		Text 370,259,50,14,"mm",.PointQualityMM
		PushButton 630,77,90,84,"Next >",.NextButton
		PushButton 30,77,90,84,"< Previous",.PrevButton
		PushButton 400,168,110,21,"Calculate",.CalcButton
		PushButton 400,196,110,21,"Clear",.ClearButton
		OKButton 520,294,90,21
		CancelButton 630,294,90,21
	End Dialog
	Dim dlg As UserDialog
	ShowTransformDlg = (Dialog(dlg) = -1)
End Function

' formats a distance as mm
Private Function FormatMM(dDistance As Double) As String
	FormatMM = Format(1000.0 * dDistance, "########0.0")
End Function

' returns the number of points that have a valid target coordinate
Private Function TargetValidPointCount() As Long
	TargetValidPointCount = 0
	Dim lPoint As Long
	For lPoint = LBound(g_points) To UBound(g_points)
		If (g_points(lPoint).TargetValid) Then
			TargetValidPointCount = TargetValidPointCount + 1
		End If
	Next
End Function


' updates the buttons and labels of the transform dialg
Private Sub UpdateDlg()
	Dim curPoint As Point
	curPoint = g_points(g_lCurrentPoint)

	DlgText "PointNumber", CStr(curPoint.Index) + "(" + CStr(curPoint.Label) + ")"
	DlgText "OfPoints", "of " + CStr(UBound(g_points))
	DlgText "ActualX", FormatMM(curPoint.ActualX)
	DlgText "ActualY", FormatMM(curPoint.ActualY)
	DlgText "ActualZ", FormatMM(curPoint.ActualZ)

	DlgVisible "PrevButton", g_lCurrentPoint > LBound(g_points)
	DlgVisible "NextButton", g_lCurrentPoint < UBound(g_points)

	Dim bAlignValid As Boolean
	Call UpdateAlignment()
	bAlignValid = g_oAligns3D(1).Valid

	DlgVisible "TotalQualityLabel", bAlignValid
	DlgVisible "TotalQuality", bAlignValid
	DlgVisible "TotalQualityMM", bAlignValid
	DlgVisible "PointQualityLabel", bAlignValid
	DlgVisible "PointQuality", bAlignValid
	DlgVisible "PointQualityMM", bAlignValid
	DlgVisible "CalcButton", TargetValidPointCount() + 1 >= ptcAlignment3DMinPointsObj

	If (bAlignValid) Then
		DlgText "TotalQuality",FormatMM(g_oAligns3D(1).CurrentQuality)
		DlgText "PointQuality",GetPointQuality(g_lCurrentPoint)
	End If

	If (curPoint.TargetValid) Then
		DlgText "TargetX", FormatMM(curPoint.TargetX)
		DlgText "TargetY", FormatMM(curPoint.TargetY)
		DlgText "TargetZ", FormatMM(curPoint.TargetZ)
	Else
		DlgText "TargetX", ""
		DlgText "TargetY", ""
		DlgText "TargetZ", ""
	End If
End Sub

' gets the quality for the point with the given index into the global point list
Private Function GetPointQuality(lPoint As Long) As String
	Dim oAlign3DPoint As Align3DPoint
	For Each oAlign3DPoint In g_oAligns3D(1).Align3DPoints
		If (g_points(lPoint).Label = oAlign3DPoint.Label) Then
			GetPointQuality = FormatMM(oAlign3DPoint.Quality)
			Exit Function
		End If
	Next
	GetPointQuality = ""
	Exit Function
End Function

Private Function ValidateDouble(strDlgItem As String, strName As String, ByRef dValue As Double) As Boolean
	ValidateDouble = False
	If Mid(CStr(1.1), 2, 1) = "." Then
		If InStr(DlgText(strDlgItem), ",") > 0 Then
			MsgBox("Please use . as decimal symbol for " + strName, vbExclamation)
			DlgFocus(strDlgItem)
			Exit Function
		End If
	Else
		If InStr(DlgText(strDlgItem), ".") > 0 Then
			MsgBox("Please use , as decimal symbol for " + strName, vbExclamation)
			DlgFocus(strDlgItem)
			Exit Function
		End If
	End If

	On Error Resume Next
	dValue = CDbl(DlgText(strDlgItem))
	If (Err.Number <> 0) Then
		On Error GoTo 0
		MsgBox("Please enter a valid number for " + strName, vbExclamation)
		DlgFocus(strDlgItem)
		Exit Function
	End If
	On Error GoTo 0

	ValidateDouble = True
End Function

' saves and validates the changes to the target coordinates of the dialog
Private Function SaveAndValidateDlg() As Boolean
	Dim x As Double
	Dim y As Double
	Dim z As Double

	SaveAndValidateDlg = True

	If (DlgText("TargetX") = "" And DlgText("TargetY") = "" And DlgText("TargetZ") = "") Then
		g_points(g_lCurrentPoint).TargetValid = False
	Else
		SaveAndValidateDlg = ValidateDouble("TargetX", "the x target coordinate", x)
		If (SaveAndValidateDlg) Then
			SaveAndValidateDlg = ValidateDouble("TargetY", "the y target coordinate", y)
		End If
		If (SaveAndValidateDlg) Then
			SaveAndValidateDlg = ValidateDouble("TargetZ", "the z target coordinate", z)
		End If
		g_points(g_lCurrentPoint).TargetValid = SaveAndValidateDlg
	End If

	If (SaveAndValidateDlg) Then
		g_points(g_lCurrentPoint).TargetX = x / 1000.0
		g_points(g_lCurrentPoint).TargetY = y / 1000.0
		g_points(g_lCurrentPoint).TargetZ = z / 1000.0
	Else
		g_points(g_lCurrentPoint).TargetX = 0.0
		g_points(g_lCurrentPoint).TargetY = 0.0
		g_points(g_lCurrentPoint).TargetY = 0.0
	End If
End Function

' calculates the target coordinates from the actual coordinates and the global transformation alignment
Private Function CalculateTargetCoordinates(lPoint As Long)
	Dim oAlign3D As Alignment3D
	Set oAlign3D = g_oAligns3D(1)

	With g_points(lPoint)
		oAlign3D.Coord3DToScanner(.ActualX, .ActualY, .ActualZ, .TargetX, .TargetY, .TargetZ)
		' transform from left hand to right hand system
		.TargetZ = -.TargetZ
		.TargetValid = True
	End With
End Function

' dialog procedure of the transform dialog
' handles all button clicks and updates the point list and transformation alignment accordingly
Private Function ShowTransformDlgProc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim curPoint As Point
	curPoint = g_points(g_lCurrentPoint)

	Select Case Action%
	Case 1 ' Dialog box initialization
		Call UpdateDlg()
	Case 2 ' Value changing or button pressed
		If DlgItem$ = "NextButton" Then
			If (g_lCurrentPoint < UBound(g_points)) Then
				If (SaveAndValidateDlg()) Then
					g_lCurrentPoint = g_lCurrentPoint + 1
					Call UpdateDlg()
					DlgFocus("TargetX")
				End If
			End If
			ShowTransformDlgProc = True
		ElseIf DlgItem$ = "PrevButton" Then
			If (g_lCurrentPoint > LBound(g_points)) Then
				If (SaveAndValidateDlg()) Then
					g_lCurrentPoint = g_lCurrentPoint - 1
					Call UpdateDlg()
					DlgFocus("TargetX")
				End If
			End If
			ShowTransformDlgProc = True
		ElseIf DlgItem$ = "CalcButton" Then
			If SaveAndValidateDlg() Then
				Call UpdateDlg()
				If (Not g_oAligns3D(1).Valid) Then
					MsgBox("Please specify target coordinates for at least " + CStr(ptcAlignment3DMinPointsObj) + " points.", vbExclamation)
					DlgFocus("TargetX")
				End If
			Else
				DlgFocus("TargetX")
			End If
			ShowTransformDlgProc = True
		ElseIf DlgItem$ = "ClearButton" Then
			g_points(g_lCurrentPoint).TargetX = 0.0
			g_points(g_lCurrentPoint).TargetY = 0.0
			g_points(g_lCurrentPoint).TargetZ = 0.0
			g_points(g_lCurrentPoint).TargetValid = False
			Call UpdateDlg()
			ShowTransformDlgProc = True
		ElseIf DlgItem$ = "OK" Then
			If SaveAndValidateDlg() Then
				Call UpdateDlg()
				If (Not g_oAligns3D(1).Valid) Then
					MsgBox("Please specify target coordinates for at least " + CStr(ptcAlignment3DMinPointsObj) + " points.", vbExclamation)
					ShowTransformDlgProc = True
					DlgFocus("TargetX")
				End If
			Else
				ShowTransformDlgProc = True
				DlgFocus("TargetX")
			End If
		End If
	End Select
End Function

' creates the global transformation alignment
' as a scan head the MSA-E-500 is used, because for this type of scanhead the
' transformation from 3D coordinates to scanner coordinates is a simple translation and
' rotation, i.e. a transformation between two cartesian coordinate systems.
Private Sub CreateAlignment()
	Set g_oAligns3D = New Alignments3D
	g_oAligns3D.Count = 1

    Dim o3DAlign As Alignment3D
    Set o3DAlign = g_oAligns3D.Item(1)

    Set g_oScanHeadScanner = CreateScanHeadScanner(ptcScanHeadTypeMSAI500)
End Sub

' updates the global transformation alignment when the list of points with valid target coordinates changed
Private Sub UpdateAlignment()
    Dim oAlign3DPoints As New Align3DPoints

	Dim lPoint As Long
	For lPoint = LBound(g_points) To UBound(g_points)
		If (g_points(lPoint).TargetValid) Then
			Call AddPoint(oAlign3DPoints, g_points(lPoint))
		End If
	Next

    Dim oAlign3D As Alignment3D
    Set oAlign3D = g_oAligns3D(1)

    If (oAlign3DPoints.Count >= ptcAlignment3DMinPointsObj) Then
	    Dim dTargetQuality As Double
	    dTargetQuality = 0.000001 ' 1 µm
	    oAlign3D.Calculate(oAlign3DPoints, ptcCoordDefModeFreePoints, False, dTargetQuality, g_oScanHeadScanner)
	Else
		oAlign3D.Invalidate()
	End If
End Sub

' create a scan head scanner object for the given front end
Private Function CreateScanHeadScanner(ScanHeadType As PTCScanHeadType) As ScanHeadScanner

	Dim oPolyScanHeads As New PolyScanHeads

	Dim oScanHeadTypes As ScanHeadTypes
	Set oScanHeadTypes = oPolyScanHeads.ScanHeadTypes

	oScanHeadTypes.Init(Nothing)

	Dim oScanHeadType As ScanHeadType
	Set oScanHeadType = oScanHeadTypes.type(ScanHeadType)

	Set CreateScanHeadScanner = oScanHeadType.ScanHeadDevices.Item(1).Scanner
End Function

' adds a point as alignment point
Private Sub AddPoint(oAlign3DPoints As Align3DPoints, alignPoint As Point)
    Dim oAlign3DPoint As Align3DPoint
    Set oAlign3DPoint = oAlign3DPoints.Add()

    oAlign3DPoint.Label = alignPoint.Label

    oAlign3DPoint.VideoX = alignPoint.VideoX
    oAlign3DPoint.VideoY = alignPoint.VideoY
    oAlign3DPoint.Caps = ptcAlign3DPointCapsVideo

    oAlign3DPoint.ScannerX =  alignPoint.TargetX
    oAlign3DPoint.ScannerY =  alignPoint.TargetY
    oAlign3DPoint.Distance = -alignPoint.TargetZ ' Attention: right hand to left hand coordinate system

    oAlign3DPoint.Caps = oAlign3DPoint.Caps Or ptcAlign3DPointCapsScanner
    oAlign3DPoint.Caps = oAlign3DPoint.Caps Or ptcAlign3DPointCapsDistance

    oAlign3DPoint.X = alignPoint.ActualX
    oAlign3DPoint.Y = alignPoint.ActualY
    oAlign3DPoint.Z = alignPoint.ActualZ
    oAlign3DPoint.Caps = oAlign3DPoint.Caps Or ptcAlign3DPointCaps3DCoord

    oAlign3DPoint.PointType = ptcAlign3DPointObjFree
End Sub

' transform the defined scan points to the new coordinate system
Private Sub ModifyScanPoints(oFile As PolyFile)
	Dim oMeasPoints As MeasPoints
	Set oMeasPoints = oFile.Infos.type(ptcInfoMeasPoints)

	Dim oMeasPoint As MeasPoint
	For Each oMeasPoint In oMeasPoints
		With oMeasPoint
			If ((.GeometryStatus And ptcGeoStatusValid) <> 0) Then
				Dim dActualX As Double
				Dim dActualY As Double
				Dim dActualZ As Double
				.CoordXYZ(dActualX, dActualY, dActualZ)
				Dim dTargetX As Double
				Dim dTargetY As Double
				Dim dTargetZ As Double
				g_oAligns3D(1).Coord3DToScanner(dActualX, dActualY, dActualZ, dTargetX, dTargetY, dTargetZ)
				' transform from left hand to right hand system
				oMeasPoint.SetCoordXYZ(dTargetX, dTargetY, -dTargetZ)
			End If
		End With
	Next
End Sub

' transforms all defined alignments to the new coordinate system by transforming the alignment points
' and recalculating the alignment
Private Sub Modify3DAlignments(oFile As PolyFile)
	Dim oAligns3D As Alignments3D
	Set oAligns3D = oFile.Infos.Alignments.Alignments3D ' creates a copy

	Dim oScanHeadDevicesInfo As ScanHeadDevicesInfoAcq
	Set oScanHeadDevicesInfo = oFile.Infos.ScanHeadDevicesInfo

	Dim lHead As Long
	For lHead = 1 To oAligns3D.Count
		Call Modify3DAlignment(oAligns3D(lHead), oScanHeadDevicesInfo.ScanHeadDevices(lHead).Scanner)
	Next lHead
End Sub

' transforms an alignment to the new coordinate system by transforming the alignment points and
' recalculating the alignment
Private Sub Modify3DAlignment(oAlign3D As Alignment3D, oScanner As ScanHeadScanner)
	Dim oAlign3DPoints As Align3DPoints
	Set oAlign3DPoints = oAlign3D.Align3DPoints

	Dim oAlign3DPoint As Align3DPoint
	For Each oAlign3DPoint In oAlign3DPoints
		With oAlign3DPoint
			Dim dTargetX As Double
			Dim dTargetY As Double
			Dim dTargetZ As Double
			g_oAligns3D(1).Coord3DToScanner(.X, .Y, .Z, dTargetX, dTargetY, dTargetZ)
			.X = dTargetX
			.Y = dTargetY
			.Z = -dTargetZ ' convert from left handed to right handed
		End With
	Next

	oAlign3D.Calculate(oAlign3DPoints, oAlign3D.CoordDefinitionMode, oAlign3D.Mirror, oAlign3D.TargetQuality, oScanner)
End Sub
