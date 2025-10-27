' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
'
' This macro demonstrates how to disable scan points dependend on their 3D coordinates.
' As an example scanpoints are disabled by a segment of cylinder.
'

'#Language "WWB.Net"
'#Uses "..\SwitchToAcquisitionMode.bas"

Imports System		'for the math object
Imports System.IO	'for the Path object
Imports System.Globalization

Option Explicit

Private Sub Explanation()
	Dim Explanation = _
		"With this macro you can enable all scan points in a specific segment of a cylinder and disable all other." + vbCrLf + vbCrLf + _
		"The cylinder axis is defined by the 'Primary Axis'."  + vbCrLf + vbCrLf + _
		"The angle value 0 ist defined by the 'Secondary Axis'." + vbCrLf + vbCrLf + _
		"Scan points in the angle range between 'Min' and 'Max' will be enabled." + vbCrLf + _
		"The values of 'Min' and 'Max' have to be in the range of 0° .. 360°" + vbCrLf + _
		"Positive angles between 'Min' and 'Max' are counterclockwise and negative angles are clockwise."

	MsgBox(Explanation, vbInformation Or vbOkOnly, appName)
End Sub


Dim appName As String
Dim fileName As String

Enum Axis
	X
	Y
	Z
End Enum

Dim primaryAxis As Axis
Dim secondaryAxis As Axis
Dim minAngle As Double
Dim maxAngle As Double

Sub Main

	Try
		Prepare()
		LoadParameters()
		DialogStart()
		SaveParameters()

	Catch exception As exception
		Dim message = exception.Message + vbCrLf + "The macro will finish now."
		MsgBox(message, vbExclamation Or vbOkOnly, appName)

	Finally
		Try
			Kill(fileName)
		Catch ' nothing
		End Try

	End Try
End Sub

Private Sub Prepare()
	appName = GetAppName()

	If Not SwitchToAcquisitionMode() Then
		Throw New exception("Switch to acquisition mode failed.")
	End If

	BackupSettings()

	If (Not Application.Acquisition.Infos.HasMeasPoints) Then
		Throw New exception("The acquisition settings does not contain any meas point collection.")
	End If

	Dim oMeasPoints = Application.Acquisition.Infos.MeasPoints
	If (0 = oMeasPoints.Count) Then
		Throw New exception("The acquisition settings does not contain any scan points.")
	End If

	Dim validMeasPointCount = 0
	Dim oMeasPoint As MeasPoint
	For Each oMeasPoint In oMeasPoints
		If ((oMeasPoint.GeometryStatus And PTCGeometryStatus.ptcGeoStatusValid) = PTCGeometryStatus.ptcGeoStatusValid) Then
			validMeasPointCount += 1
		End If
	Next

	If (validMeasPointCount = 0) Then
		Throw New exception("The acquisition settings does not contain any valid scan points.")
	End If

	If (Not oMeasPoints.Has3DCoordinates) Then
		Throw New exception("The acquisition settings does not contain a 3D geometry." + vbCrLf + _
			   				"Segment disabling is only possible for 3D geometries.")
	End If

	' Because we can't manipulate the meas point's scan status directly in the application object
	' the acquisition settings are saved to a file.
	' Later this file will be opened using PolyFile, the meas point's scan status will be changed
	' as needed and the file will be reloaded to the application(PSV).
	fileName = Path.GetTempPath() + appName + ".set"
	Application.Settings.Save(fileName)

End Sub

Private Function GetAppName() As String
	GetAppName = CallersLine

	' at this point the value of GetAppName is something like this "[D:\AP_5296-42\PrepareCylinderScan.bas|Prepare# 66] appName$ = GetAppName$()"
	' the substring between the final "\" and "|" is used as application name

	Dim endIdx = InStr(1, GetAppName , "|")
	Dim beginIdx = InStrRev(GetAppName , "\", endIdx) + 1
	GetAppName = Mid(GetAppName, beginIdx, endIdx - beginIdx)
End Function

Private Sub BackupSettings()
	Dim backupMessage = "It is recommended to save the current acquisition settings." + vbCrLf + "Do you want to do that?"

	If( vbNo = MsgBox(backupMessage , vbExclamation Or vbYesNo, appName)) Then
		Return
	End If

	Dim backupFilename = GetFilePath("backup", "Psv Setting File|*.set", , , 3)
	If (0 = Len(backupFilename)) Then
		Return
	End If

	Application.Settings.Save(backupFilename)
End Sub


Private Sub LoadParameters()
	primaryAxis = StrToAxis(GetSetting(appName, "Primary", "Axis", AxisToStr(Axis.Y)))
	secondaryAxis = StrToAxis(GetSetting(appName, "Secondary", "Axis", AxisToStr(Axis.X)))

	Dim defaultMin = 90.0
	minAngle = StrToDbl(GetSetting(appName, "Angle", "Min", DblToStr(defaultMin)))
	If(minAngle < 0 Or minAngle > 360) Then
		minAngle = defaultMin
	End If

	Dim defaultMax = 270.0
	maxAngle = StrToDbl(GetSetting(appName, "Angle", "Max", DblToStr(defaultMax)))
	If(maxAngle < 0 Or maxAngle > 360) Then
		maxAngle = defaultMax
	End If
End Sub

Private Sub SaveParameters()
	SaveSetting(appName, "Primary", "Axis", AxisToStr(primaryAxis))
	SaveSetting(appName, "Secondary", "Axis", AxisToStr(secondaryAxis))
	SaveSetting(appName, "Angle", "Min", DblToStr(minAngle))
	SaveSetting(appName, "Angle", "Max", DblToStr(maxAngle))
End Sub

Private Sub DialogStart
	Begin Dialog UserDialog 260,238,appName,.DlgFunc ' %GRID:10,7,1,1
		GroupBox 10,7,240,42,"Select Primary Axis",.GROUPBOX_PRIM_AXIS
		OptionGroup .PRIM_AXIS
			OptionButton 20,28,40,14,"X",.PRIM_AXIS_X 'PRIM_AXIS_X
			OptionButton 110,28,40,14,"Y",.PRIM_AXIS_Y 'PRIM_AXIS_Y
			OptionButton 190,28,40,14,"Z",.PRIM_AXIS_Z 'PRIM_AXIS_Z
		GroupBox 10,56,240,42,"Select Secondary Axis (0 angle)",.GROUPBOX_SECOND_AXIS
		OptionGroup .SECOND_AXIS
			OptionButton 20,77,40,14,"X",.SECOND_AXIS_X 'SECOND_AXIS_X
			OptionButton 110,77,40,14,"Y",.SECOND_AXIS_Y 'SECOND_AXIS_Y
			OptionButton 190,77,40,14,"Z",.SECOND_AXIS_Z 'SECOND_AXIS_Z
		GroupBox 10,105,240,42,"Angle Range[°]",.ANGLE_RANGE
		Text 20,126,40,14,"Min",.TEXT_MIN
		TextBox 60,126,40,14,.MIN_ANGLE
		Text 140,126,40,14,"Max",.TEXT_MAX
		TextBox 190,126,40,14,.MAX_ANGLE
		PushButton 100,154,150,21,"Explanation",.EXPLANATION 'INFORMATION
		PushButton 100,182,150,21,"Enable Points",.ENABLE_POINTS 'ENABLE_POINTS
		OKButton 100,210,150,21
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
End Sub

Dim nextOperation = ""

Rem See DialogFunc help topic for more information.
Private Function DlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Select Case Action%
		Case 1 ' Dialog box initialization
			DialogInit()

		Case 2 ' Value changing or button pressed
			Select Case DlgItem
				Case "ENABLE_POINTS"
					DialogEnableScanPoints()
					DlgFunc = True 'continue dialog

				Case "PRIM_AXIS"
					DialogAxisChanged()
					DlgFunc = True 'continue dialog

				Case "SECOND_AXIS"
					DialogAxisChanged()
					DlgFunc = True 'continue dialog

				Case "EXPLANATION"
					If (nextOperation = "ENABLE_POINTS") Then
						DialogEnableScanPoints()
					Else
						Explanation()
					End If

					SetFocus()
					DlgFunc = True 'continue dialog

			    Case Else '"PushButtonCancel" and "OK" (OK stands for Close from system menue)
					DlgFunc = False ' dialog end

			End Select
		Case 3 ' TextBox or ComboBox text changed
		Case 4 ' Focus changed
			Select Case DlgItem
				Case "MIN_ANGLE"
					nextOperation = "ENABLE_POINTS"

				Case "MAX_ANGLE"
					nextOperation = "ENABLE_POINTS"

				Case Else
					nextOperation = ""

			End Select

		Case 5 ' Idle
			Rem Wait .1 : DlgFunc = True ' Continue getting idle actions
		Case 6 ' Function key
	End Select
End Function

Private Sub DialogInit()
	Try
		DlgValue "PRIM_AXIS" ,primaryAxis
		DlgValue "SECOND_AXIS" ,secondaryAxis
		DlgText "MIN_ANGLE" ,DblToStr(minAngle)
		DlgText "MAX_ANGLE" ,DblToStr(maxAngle)
		EnableDisableControls()

	Catch exception As exception
		MsgBox(exception.Message, vbExclamation Or vbOkOnly, appName)

	End Try
End Sub

Private Sub DialogAxisChanged()
	Try
		ReadDialogData()
		EnableDisableControls()

	Catch exception As exception
		MsgBox(exception.Message, vbExclamation Or vbOkOnly, appName)

	End Try
End Sub

Private Sub	DialogEnableScanPoints()
	Try
		ReadDialogData()
		EnableScanPoints()

	Catch exception As exception
		MsgBox(exception.Message, vbExclamation Or vbOkOnly, appName)

	End Try
End Sub

Private Sub ReadDialogData()
	primaryAxis = DlgValue("PRIM_AXIS")
	secondaryAxis = DlgValue("SECOND_AXIS")
	minAngle = CAngle(DlgText("MIN_ANGLE"))
	maxAngle = CAngle(DlgText("MAX_ANGLE"))
End Sub

Private Sub EnableDisableControls()
	DlgEnable "SECOND_AXIS_X", True
	DlgEnable "SECOND_AXIS_Y", True
	DlgEnable "SECOND_AXIS_Z", True

	Select Case primaryAxis
		Case Axis.X
			DlgEnable "SECOND_AXIS_X", False
		Case Axis.Y
			DlgEnable "SECOND_AXIS_Y", False
		Case Axis.Z
			DlgEnable "SECOND_AXIS_Z", False
	End Select

	DlgEnable "ENABLE_POINTS", (primaryAxis <> secondaryAxis)

	SetFocus()
End Sub

Private Sub SetFocus()
	If(primaryAxis <> secondaryAxis) Then
		DlgFocus "ENABLE_POINTS"
	End If
End Sub


Private Function CAngle(value As String) As Double
	Try
		CAngle = StrToDbl(value)
	Catch exception As exception
		Throw New exception("The angle have to be a real number.")
	End Try

	If( CAngle < 0 Or CAngle > 360) Then
		Throw New exception("The angle have to be in the range of 0°.. 360°.")
	End If

End Function

Private Function OpenFile(strFileName As String) As PolyFile

	OpenFile = New PolyFile
	OpenFile.ReadOnly = False
	OpenFile.Open(strFileName)
	If Not OpenFile.IsOpen Then
		Throw New exception("Can not open the file.")
	End If
End Function

Private Sub	EnableScanPoints()
	Try
		Dim oFile = OpenFile(fileName)
		Dim oMeasPoints = oFile.Infos.MeasPoints
		Dim oMeasPoint As MeasPoint
		For Each oMeasPoint In oMeasPoints
			If ((oMeasPoint.GeometryStatus And  PTCGeometryStatus.ptcGeoStatusValid) = PTCGeometryStatus.ptcGeoStatusValid) Then

				Dim Angle = CalcAngle(oMeasPoint)

				'counterclockwise angle range
				If ((Angle >= minAngle) And (Angle < maxAngle)) Then
					oMeasPoint.ScanStatus = oMeasPoint.ScanStatus And Not PTCScanStatus.ptcScanStatusDisabled 'enable scanpoint
				'clockwise angle range
				ElseIf ((minAngle > maxAngle) And ((Angle >= minAngle) Or (Angle < maxAngle))) Then
					oMeasPoint.ScanStatus = oMeasPoint.ScanStatus And Not PTCScanStatus.ptcScanStatusDisabled 'enable scanpoint
				'out of range
				Else
					oMeasPoint.ScanStatus = PTCScanStatus.ptcScanStatusDisabled ' disable Scanpoint
				End If

			End If
		Next

		oFile.Save()
		oFile.Close()
		Application.Settings.Load(fileName, PTCSettings.ptcSettingsAPS)

	Finally
		Try
			oFile.Close()
		Catch ' nothing
		End Try

	End Try
End Sub

Private Function CalcAngle(ByRef oMeasPoint As MeasPoint) As Double
	Dim x , y , z As Double
	oMeasPoint.CoordXYZ(x, y, z)

	Dim Angle As Double 'angle in rad [-pi..pi]

	Select Case primaryAxis
	Case Axis.X
		Select Case secondaryAxis
			Case Axis.Y
				Angle = math.Atan2(z,y)
			Case Axis.Z
				Angle = -math.Atan2(y,z)
			Case Else
				Throw New exception("undefined case.")
		End Select
	Case Axis.Y
		Select Case secondaryAxis
			Case Axis.Z
				Angle = math.Atan2(x,z)
			Case Axis.X
				Angle = -math.Atan2(z,x)
			Case Else
				Throw New exception("undefined case.")
		End Select
	Case Axis.Z
		Select Case secondaryAxis
			Case Axis.X
				Angle = math.Atan2(y,x)
			Case Axis.Y
				Angle = -math.Atan2(x,y)
			Case Else
				Throw New exception("undefined case.")
		End Select
	End Select

	If (math.Abs(Angle) > math.PI) Then
		Throw New exception("Unexpected Value.")
	End If

	If (Angle < 0) Then
		Angle += 2 * math.PI  'angle in rad [0..2pi]
	End If


	CalcAngle = Angle * 180 / math.PI ' angle in degree [0°.. 360°]
End Function

Private Function DblToStr(value As Double) As String
	Return value.ToString(CultureInfo.InvariantCulture)
End Function

Private Function StrToDbl(ByRef value As String) As Double
	Return Double.Parse(value, CultureInfo.InvariantCulture)
End Function

Private Function AxisToStr(value As Axis) As String
	Select Case value
		Case Axis.X
			AxisToStr = "X"
		Case Axis.Y
			AxisToStr = "Y"
		Case Axis.Z
			AxisToStr = "Z"
		Case Else
			Throw New exception("undefined case.")
	End Select
End Function

Private Function StrToAxis(ByRef value As String) As Axis
	Select Case value
		Case "X"
			StrToAxis = Axis.X
		Case "Y"
			StrToAxis = Axis.Y
		Case "Z"
			StrToAxis = Axis.Z
		Case Else
			StrToAxis = Axis.X
	End Select
End Function
