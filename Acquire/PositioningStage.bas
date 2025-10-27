'#Reference {50A7E9B0-70EF-11D1-B75A-00A0C90564FE}#1.0#0#C:\Windows\system32\SHELL32.dll#Microsoft Shell Controls And Automation
'
' This macro shows the use of the XY positioning stage for stiching.
'
' This file has to meet the following conditions:
'
' - you have a MSA system with XY positioning stage
'

Option Explicit


Dim g_strDir As String

Sub Main
	' switch to acquisition mode
	If (Not SwitchToAcquisitionMode()) Then
		Exit Sub
	End If

	Dim oAlignmentsAcq As AlignmentsAcq
	Set oAlignmentsAcq = Application.Acquisition.Infos.Alignments

	Dim oSHDevs As ScanHeadDevices
	Set oSHDevs = Application.Acquisition.Infos.ScanHeadDevicesInfo.ScanHeadDevices	' ScanHeadDevices gives use a clone copy of the scanheaddevices, call them only once a time

	If Application.Acquisition.Infos.HasPositionDevice = False Then
		Exit Sub
	End If

	Dim oPosDevice As PositionDevice
	Set oPosDevice = Application.Acquisition.Infos.PositionDevice

	If oPosDevice.HasXYAxes = False Then
		Exit Sub
	End If

	If oPosDevice.IsReferencedXY = False Then
		MsgBox("XY positioning stage is not referenced. Please run reference of XY stage and restart the macro.")
		Exit Sub
	End If

	Dim startXValue As Double
	startXValue = 0.00022
	Dim startYValue As Double
	startYValue = 0.0001

	Dim moveXValue As Double
	moveXValue = 0.0004
	Dim moveYValue As Double
	moveYValue = 0.0003

	' the scan point indexes of the files will differ by this offset
	Dim indexOffset As Long
	indexOffset = 10000

	'Set starting position
	Call MoveXYStageToPosition(oPosDevice, startXValue, startYValue)

	g_strDir = BrowseFolder(0, "Select a folder to save to")
	If g_strDir = "" Then End

	Dim targetX As Double
	Dim targetY As Double
	oPosDevice.GetPositionXY(targetX, targetY)

	Call ScanAtPosition(1)

	'Call MoveXYStageRelative(oPosDevice, -moveXValue, 0.0)
	targetX = targetX - moveXValue
	Call MoveXYStageToPosition(oPosDevice, targetX, targetY)
	Call ChangeAlignment(oAlignmentsAcq, oSHDevs, -moveXValue, 0.0)
	Call ScanAtPosition(2)

	'Call MoveXYStageRelative(oPosDevice, 0.0, -moveYValue)
	targetY = targetY - moveYValue
	Call MoveXYStageToPosition(oPosDevice, targetX, targetY)
	Call ChangeAlignment(oAlignmentsAcq, oSHDevs, 0.0, -moveYValue)
	Call ScanAtPosition(3)

	'Call MoveXYStageRelative(oPosDevice, moveXValue, 0.0)
	targetX = targetX + moveXValue
	Call MoveXYStageToPosition(oPosDevice, targetX, targetY)
	Call ChangeAlignment(oAlignmentsAcq, oSHDevs, moveXValue, 0.0)
	Call ScanAtPosition(4)

	'Call MoveXYStageRelative(oPosDevice, 0.0, moveYValue)
	targetY = targetY + moveYValue
	Call MoveXYStageToPosition(oPosDevice, targetX, targetY)
	Call ChangeAlignment(oAlignmentsAcq, oSHDevs, 0.0, moveYValue)

	If (Not SwitchToPresentationMode()) Then
		Exit Sub
	End If

	Dim oPolyFile As PolyFile
	Set oPolyFile = New PolyFile
	oPolyFile.FileName = g_strDir + "\PositionDeviceScan_Stiched.svd"
	oPolyFile.ReadOnly = False
	Dim oContainedFiles As ContainedFiles
	Set oContainedFiles = oPolyFile.ContainedFiles
	Dim FileName As String
	Dim Position As Long
	For Position = 1 To 4
		FileName = g_strDir + "\PositionDeviceScan" + Format(Position) + ".svd"
		Call AddIndexOffset(FileName, Position * indexOffset)
		oContainedFiles.Add(FileName)
	Next Position

	oPolyFile.Save()
	oPolyFile.Close()

	Application.Documents.Open(g_strDir + "\PositionDeviceScan_Stiched.svd")

End Sub


Sub ScanAtPosition(ByVal Position As Integer)
    Acquisition.ScanFileName = g_strDir + "\PositionDeviceScan" + Format(Position) + ".svd"
    Acquisition.Scan(ptcScanAll)

    While Acquisition.State <> ptcAcqStateStopped
        Wait(1)
    Wend
End Sub


Sub WaitWhileMoving(ByRef oPosDevice As PositionDevice)
	While oPosDevice.IsMovingXY = True
		Wait(0.5)
	Wend
End Sub

Sub MoveXYStageToPosition(ByRef oPosDevice As PositionDevice, ByVal PosX As Double, ByVal PosY As Double)

	If PosX >= oPosDevice.PositionXMin And PosX <= oPosDevice.PositionXMax And PosY >= oPosDevice.PositionYMin And PosY <= oPosDevice.PositionYMax Then
		oPosDevice.StartMoveXY(PosX, PosY)
		WaitWhileMoving(oPosDevice)
	End If
	Wait(3)
End Sub

Sub MoveXYStageRelative(ByRef oPosDevice As PositionDevice, ByVal RelX As Double, ByVal RelY As Double)
	Dim currentX As Double
	Dim currentY As Double
	oPosDevice.GetPositionXY(currentX, currentY)
	Dim absoluteX As Double
	Dim absoluteY As Double
	absoluteX = currentX + RelX
	absoluteY = currentY + RelY

	Call MoveXYStageToPosition(oPosDevice, absoluteX, absoluteY)
End Sub

Sub ChangeAlignment(ByRef oAlignmentsAcq As AlignmentsAcq, ByRef oSHDevices As ScanHeadDevices, ByVal TranslationX As Double, ByVal TranslationY As Double)

	Dim oAlignments3D As Alignments3D
	Set oAlignments3D = oAlignmentsAcq.Alignments3D	' get a copy of the actual 3D alignments

	Dim oAlign3D As Alignment3D
	For Each oAlign3D In oAlignments3D
		If oAlign3D.Valid Then
			Dim oAlignPoints3D As New Align3DPoints
			Set oAlignPoints3D = oAlign3D.Align3DPoints ' Align3DPoints gives use a clone copy of the existing points

			Dim oAlignPoint3D As Align3DPoint
			For Each oAlignPoint3D In oAlignPoints3D
				oAlignPoint3D.X = oAlignPoint3D.X + TranslationX
				oAlignPoint3D.Y = oAlignPoint3D.Y + TranslationY
			Next oAlignPoint3D

			' calculate new 3D alignment
			oAlign3D.Calculate(oAlignPoints3D, oAlign3D.CoordDefinitionMode, False, 0.001, oSHDevices.Item(1).Scanner)
		End If
	Next oAlign3D

	Set oAlignmentsAcq.Alignments3D = oAlignments3D

End Sub

Sub AddIndexOffset(fileName As String, offset As Long)
	Dim oFile As New PolyFile
	oFile.ReadOnly = False
	oFile.Open(fileName)

	Dim indexes() As Long
	Dim labels() As Long

	Dim oMeasPoints As MeasPoints
	Set oMeasPoints = oFile.Infos.MeasPoints

	ReDim indexes(0 To oMeasPoints.Count - 1)
	ReDim labels(0 To oMeasPoints.Count - 1)

	Dim oMeasPoint As MeasPoint
	Dim i As Long
	For Each oMeasPoint In oMeasPoints
		indexes(i) = i + 1
		labels(i) = oMeasPoint.Label + offset
		i = i + 1
	Next

	oFile.Infos.MeasPoints.SetLabels(indexes, labels)

	oFile.Save()
	oFile.Close()
End Sub


Private Function SwitchToAcquisitionMode() As Boolean
	If Application.Mode = ptcApplicationModePresentation Then
		Application.Mode = ptcApplicationModeAcquisition
	End If
	If Application.Mode = ptcApplicationModePresentation Then
		Beep()
		MsgBox("Cannot switch to Acquisition Mode")
		SwitchToAcquisitionMode = False
	Else
		SwitchToAcquisitionMode = True
	End If
End Function

Private Function SwitchToPresentationMode() As Boolean
	If Application.Mode = ptcApplicationModeAcquisition Then
		Application.Mode = ptcApplicationModePresentation
	End If
	If Application.Mode = ptcApplicationModeAcquisition Then
		Beep()
		MsgBox("Cannot switch to Presentation Mode")
		SwitchToPresentationMode = False
	Else
		SwitchToPresentationMode = True
	End If
End Function

'Constants

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Const BIF_NEWDIALOGSTYLE = &H40

'Types

Private Type BrowseInfo
   hWndOwner      As PortInt
   pIDLRoot       As PortInt
   pszDisplayName As PortInt
   lpszTitle      As String
   ulFlags        As Long
   lpfnCallback   As PortInt
   lParam         As PortInt
   iImage         As Long
End Type

Private Declare Function SHBrowseForFolder   Lib "shell32"  Alias "SHBrowseForFolderA" 	 (lpbi As BrowseInfo) As PortInt
Private Declare Function SHGetPathFromIDList Lib "shell32"  Alias "SHGetPathFromIDListA" (ByVal pidList As PortInt, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat             Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As String

Public Function BrowseFolder(ByVal WindowHandle As PortInt, ByVal BrowseTitle As String) As String

	'Opens a Treeview control that displays the directories in a computer

	Dim lpIDList As PortInt
	Dim sBuffer As String
	Dim szTitle As String
	Dim tBrowseInfo As BrowseInfo

	szTitle = BrowseTitle '"Select the folder to save to"
	With tBrowseInfo
		.hWndOwner = WindowHandle
		.lpszTitle = lstrcat(szTitle, "")
		.ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_NEWDIALOGSTYLE
	End With

	lpIDList = SHBrowseForFolder(tBrowseInfo)

	If (lpIDList) Then
		sBuffer = Space$(MAX_PATH)
		SHGetPathFromIDList(lpIDList, sBuffer)
		sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
		BrowseFolder = sBuffer
	End If

End Function
