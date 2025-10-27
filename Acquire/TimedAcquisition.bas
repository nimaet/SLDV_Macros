'#Reference {5BF46737-4121-4B92-AF7A-BF07871BDAE3}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolyTimer.dll#Polytec PolyTimer Type Library
' POLYTEC MACRO
' ----------------------------------------------------------------------
' This Macro does a user selectable number of single shots or scans.
' The user selects a base-filename. The macro appends the numbers 1, 2 etc.
' for the individual measurements.
' The user can select a waiting time between the measurements and
' a start time for the complete cycle.
'
' References
' - Polytec PSV\VibSoft Type Library
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
' - Polytec PolyTimer Type Library
' ----------------------------------------------------------------------
'
'#Uses "..\SwitchToAcquisitionMode.bas"

Option Explicit

' Global constants
Const CA_FileFilter = Array("Scan Data (*.svd)|*.svd|All Files (*.*)|*.*||", "Single Point Data (*.pvd)|*.pvd|All Files (*.*)|*.*||")
Const CA_FileExt = Array(".svd", ".pvd")
Const C_Scan% = 0
Const C_SinglePoint% = 1
Const C_Progress$ = "|"
Const C_MaxProgressChar% = 130
'
' Global variables
Dim BackButtonPressed As Boolean
Dim DlgAborted As Boolean
Dim AcqType As Integer
Dim BaseFileName(2) As String
Dim AcqBaseFileName(2) As String
Dim AcqCount As Integer
Dim WaitingTime As Integer
Dim AcqStartTime As Date
Dim AcqStartDirect As Boolean
Dim IsVibSoft As Boolean
Dim oPolyTimer As POLYTIMERLib.Timer
'
Sub Main
' -------------------------------------------------------------------------------
' Main procedure.
' -------------------------------------------------------------------------------
	Call InitGlobalVariables

	If Not SwitchToAcquisitionMode() Then GoTo MEnd
	If IsVibSoft Then GoTo M2
M1:
	Call SelectAcquisitionType
	If DlgAborted Then GoTo MEnd
M2:
	Call SelectBaseFilename
	If DlgAborted Then GoTo MEnd
	If BackButtonPressed Then GoTo M1
M3:
	Call SelectCountAndTime
	If DlgAborted Then GoTo MEnd
	If BackButtonPressed Then GoTo M2

	Call StartTimedAcquisition
	If DlgAborted Then GoTo MEnd
	If BackButtonPressed Then GoTo M3

	If Not AcqStartDirect Then
		Call ShowClock
		If DlgAborted Then GoTo MEnd
	End If

	If AcqStartDirect Then
		Call DoAcquisition
	End If
MEnd:
	MsgBox("Macro has finished.", vbOkOnly)
End Sub

Private Sub SelectAcquisitionType
' -------------------------------------------------------------------------------
' Select acquisition type.
' -------------------------------------------------------------------------------
	Begin Dialog UserDialog 400,203,"Prepare Timed Acquisition",.SelectAcquisitionTypeDlgProc
		Text 20,7,360,14,"Select the acquisition type.",.Text1
		OptionGroup .AcquisitionType
			OptionButton 60,49,210,14,"Single measurement",.OptionSinglePoint
			OptionButton 60,84,160,14,"Scan",.OptionScan
		PushButton 290,175,90,21,"Next >",.NextButton
		CancelButton 190,175,90,21
	End Dialog
	Dim dlg As UserDialog
	Dim iRe As Integer
	iRe= Dialog (dlg)
	If iRe = 0 Then
		DlgAborted = True
	End If
End Sub

Private Function SelectAcquisitionTypeDlgProc(DlgItem$, Action%, SuppValue&) As Boolean
' -------------------------------------------------------------------------------
' Select acquisition type.
' -------------------------------------------------------------------------------
	Select Case Action%
	Case 1 ' Dialog box initialization
		If AcqType = C_Scan% Then
			DlgValue "AcquisitionType" , 1
		ElseIf AcqType = C_SinglePoint% Then
			DlgValue "AcquisitionType" , 0
		End If
	Case 2 ' Value changing or button pressed
		If DlgItem$ = "NextButton" Then
			If DlgValue("AcquisitionType") = 0 Then
				AcqType = C_SinglePoint%
			ElseIf DlgValue("AcquisitionType") = 1 Then
				AcqType = C_Scan%
			End If
		End If
		Rem SelectAcquisitionTypeDlgProc = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem SelectAcquisitionTypeDlgProc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Private Sub SelectBaseFilename
' -------------------------------------------------------------------------------
' Select filename.
' -------------------------------------------------------------------------------
	' Declare User Dialog
	Begin Dialog UserDialog 400,203,"Prepare Timed Acquisition",.SelectBaseFilenameDlgProc
		Text 20,7,360,14,"Select a base filename.",.Text1
		Text 20,28,350,28,"The acquisition will be saved under the base filename appended with the measurement count.",.Text2
		Text 20,70,90,14,"Filepath:",.Text3
		Text 130,70,240,14,"Static",.FilePath
		Text 20,105,100,14,"Base filename:",.Text4
		TextBox 130,98,240,21,.BaseFilename
		PushButton 290,175,90,21,"Next >",.NextButton
		PushButton 20,175,90,21,"< Back",.BackButton
		CancelButton 190,175,90,21
		PushButton 280,126,90,21,"Browse...",.BrowseButton
	End Dialog
	' Show dialog
	Dim dlg As UserDialog
	Dim iRe As Integer
	iRe = Dialog (dlg)
	If iRe = 0 Then
		DlgAborted = True
	End If
End Sub

Private Function SelectBaseFilenameDlgProc(DlgItem$, Action%, SuppValue&) As Boolean
' -------------------------------------------------------------------------------
' Select filename.
' -------------------------------------------------------------------------------
	Dim sFilePath As String
	Dim sPath As String
	Dim sFileName As String
	Dim sExt As String

	Select Case Action%
	Case 1 ' Dialog box initialization
		BackButtonPressed = False

		sFilePath = AcqBaseFileName(AcqType) + CA_FileExt(AcqType)
		Call SplitPath(sFilePath, sPath, sFileName, sExt)

		DlgText "BaseFilename", sFileName
		DlgText "FilePath", sPath
		If IsVibSoft Then
			DlgVisible "BackButton", False
		End If
	Case 2 ' Value changing or button pressed
		If DlgItem$ = "BrowseButton" Then
			sFilePath = FileOpenDialog(DlgText ("BaseFilename"), CA_FileFilter(AcqType), CA_FileExt(AcqType))
			If sFilePath <>"" Then
				Call SplitPath(sFilePath, sPath, sFileName, sExt)
				DlgText "BaseFilename", sFileName
				DlgText "FilePath", sPath
			End If
			SelectBaseFilenameDlgProc = True
		ElseIf DlgItem$ = "NextButton" Or DlgItem$ = "BackButton" Then
			sFileName = DlgText("BaseFilename")
			sPath = DlgText("FilePath")
			sFilePath = sPath + sFileName
			Call SplitPath(sFilePath, sPath, sFileName, sExt)
			AcqBaseFileName(AcqType) = sPath + sFileName
		End If
		If DlgItem$ = "BackButton" Then
			BackButtonPressed = True
		End If

		Rem PrepareTimedMeasurementDlg = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem PrepareTimedMeasurementDlg = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Private Sub SelectCountAndTime
' -------------------------------------------------------------------------------
' Select count and time.
' -------------------------------------------------------------------------------
	' Declare User Dialog
	Begin Dialog UserDialog 400,203,"Prepare Timed Acquisition",.SelectCountAndTimeDlg
		Text 20,7,360,14,"Set the count of acquisitions.",.Text1
		Text 20,91,350,14,"Set the waiting time between two acquisitions.",.Text2
		TextBox 250,28,90,21,.Count
		TextBox 250,112,90,21,.Time
		PushButton 290,175,90,21,"Next >",.NextButton
		PushButton 20,175,90,21,"< Back",.BackButton
		CancelButton 190,175,90,21
		Text 70,35,140,14,"Count of acquisitions:",.Text3
		Text 70,119,160,14,"Waiting time (int sec.):",.Text4
	End Dialog
	' Show dialog
	Dim dlg As UserDialog
	Dim iRe As Integer
	iRe = Dialog (dlg)
	If iRe = 0 Then
		DlgAborted = True
	End If
End Sub

Private Function SelectCountAndTimeDlg(DlgItem$, Action%, SuppValue&) As Boolean
' -------------------------------------------------------------------------------
' Select count and time.
' -------------------------------------------------------------------------------
	Select Case Action%
	Case 1 ' Dialog box initialization
		BackButtonPressed = False
		DlgText "Count" , CStr$(AcqCount)
		DlgText "Time" , CStr$(WaitingTime)
	Case 2 ' Value changing or button pressed
		If DlgItem$ = "NextButton" Or DlgItem$ = "BackButton" Then
			AcqCount = Val(DlgText("Count"))
			WaitingTime = Val(DlgText("Time"))
		End If
		If DlgItem$ = "BackButton" Then
			BackButtonPressed = True
		End If
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
	Case 6 ' Function key
	End Select
End Function

Private Sub StartTimedAcquisition
' -------------------------------------------------------------------------------
' Input start time.
' -------------------------------------------------------------------------------
	' Declare User Dialog
	Begin Dialog UserDialog 400,203,"Start Timed Acquisition",.StartTimedAcquisitionDlg
		Text 20,7,360,14,"Select the type of signal on which the acquisition starts.",.Text1
		OptionGroup .StartOptions
			OptionButton 40,42,320,14,"&Immediately after Finish button pressed",.OptionStartDirect
			OptionButton 40,84,160,14,"&Start Acquisition at",.OptionStartTimed
		PushButton 290,175,90,21,"Finish",.FinishButton
		PushButton 20,175,90,21,"< Back",.BackButton
		CancelButton 190,175,90,21
		TextBox 230,77,120,21,.StartTime
	End Dialog
	' Show dialog
	Dim dlg As UserDialog
	Dim iRe As Integer
	iRe = Dialog (dlg)
	If iRe = 0 Then
		DlgAborted = True
	End If
End Sub

Private Function StartTimedAcquisitionDlg(DlgItem$, Action%, SuppValue&) As Boolean
' -------------------------------------------------------------------------------
' Input start time.
' -------------------------------------------------------------------------------
	Select Case Action%
	Case 1 ' Dialog box initialization
		BackButtonPressed = False
		If AcqStartDirect Then
			DlgValue "StartOptions" , 0
		Else
			DlgValue "StartOptions" , 1
		End If
		AcqStartTime = Time
		DlgText "StartTime", CStr$(AcqStartTime)
	Case 2 ' Value changing or button pressed
		If DlgItem$ = "FinishButton" Then
			If DlgValue("StartOptions") =  0 Then
				AcqStartDirect = True
			Else
				AcqStartDirect = False
			End If
			AcqStartTime = TimeValue(DlgText("StartTime"))
		End If
		If DlgItem$ = "BackButton" Then
			BackButtonPressed = True
		End If
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
		' Update Time value if edit box looses the focus
		If SuppValue& = DlgNumber("StartTime") Then
			AcqStartTime = TimeValue(DlgText("StartTime"))
			DlgText "StartTime", CStr$(AcqStartTime)
		End If

	Case 5 ' Idle
	Case 6 ' Function key
	End Select
End Function

Private Sub ShowClock
' -------------------------------------------------------------------------------
' Show clock.
' -------------------------------------------------------------------------------
	' Declare User Dialog
	Begin Dialog UserDialog 400,84,"Acquisition starts in ...",.ShowClockDlg
		GroupBox 20,7,360,35,"",.GroupBox1
		Text 120,21,140,14,"2:22",.TimeDiff
		CancelButton 150,56,90,21
	End Dialog
	' Show dialog
	Dim dlg As UserDialog
	Dim iRe As Integer
	iRe = Dialog (dlg)
	If iRe = 0 Then
		DlgAborted = True
	ElseIf iRe = 1 Then
		AcqStartDirect = True
	End If
End Sub

Private Function ShowClockDlg(DlgItem$, Action%, SuppValue&) As Boolean
' -------------------------------------------------------------------------------
' Show clock.
' -------------------------------------------------------------------------------
	Static LastTime
	Select Case Action%
	Case 1 ' Dialog box initialization
		LastTime = Time
		DlgText "TimeDiff", TimeSpan(AcqStartTime - LastTime)
	Case 2 ' Value changing or button pressed
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		If LastTime < Time Then
			LastTime = Time
			If AcqStartTime <= LastTime Then
				DlgEnd 1
			Else
				DlgText "TimeDiff", TimeSpan(AcqStartTime - LastTime)
			End If
		End If
		ShowClockDlg = True
	Case 6 ' Function key
	End Select
End Function

Private Sub DoAcquisition
' -------------------------------------------------------------------------------
' Acquisition.
' -------------------------------------------------------------------------------
	' Declare User Dialog
	Begin Dialog UserDialog 400,203,"Timed Acquisition in progress",.DoAcquisitionDlg
		GroupBox 20,105,360,35,"",.GroupBox2
		Text 20,28,100,14,"Acquisition no.",.Text1
		Text 130,28,90,14,"1 of 100",.AcqNo
		GroupBox 20,42,360,35,"",.GroupBox1
		Text 30,56,340,14,"#####",.AcqProgress
		Text 20,91,150,14,"Waiting...",.Waiting
		Text 30,119,340,14,"#####",.WaitingProgress
		PushButton 290,175,90,21,"Stop",.StopButton
	End Dialog
	' Show dialog
	Dim dlg As UserDialog
	Dim iRe As Integer
	iRe = Dialog (dlg)
	If iRe = 0 Then
		DlgAborted = True
	End If
End Sub

Private Function DoAcquisitionDlg(DlgItem$, Action%, SuppValue&) As Boolean
' -------------------------------------------------------------------------------
' Acquisition.
' -------------------------------------------------------------------------------
	Static AcqMode As Integer	' 0 = Start Acquisition , 1 = Acquisition in progress, 2 = waiting
	Static AcqNo As Integer
	Static WaitingTimer
	Static LastTimer
	Dim sProgress As String

	Select Case Action%
	Case 1 ' Dialog box initialization
		AcqMode = 0
		AcqNo = 1
		DlgText "AcqNo" , ""
		DlgText "AcqProgress" , ""
		DlgText "WaitingProgress" , ""
	Case 2 ' Value changing or button pressed
		If Acquisition.State <> ptcAcqStateStopped Then
			Acquisition.Stop 
		End If
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		If AcqMode = 0 Then	' Start Acquisition
			DlgText "AcqProgress" , ""
			DlgText "AcqNo" , CStr(AcqNo) + " of " + CStr$(AcqCount)
			If AcqType = C_Scan% Then
				Acquisition.ScanFileName = AcqBaseFileName(AcqType) + CStr$(AcqNo) + CA_FileExt(AcqType)
				Acquisition.Scan ptcScanAll
			Else
				Acquisition.Start ptcAcqStartSingle
			End If
			AcqMode = 1
		ElseIf AcqMode = 1 Then
			If Acquisition.State = ptcAcqStateStopped Then
				If AcqType = C_SinglePoint% Then
					Acquisition.Document.SaveAs(AcqBaseFileName(AcqType) + CStr$(AcqNo) + CA_FileExt(AcqType))
				End If
				AcqNo = AcqNo + 1
				If AcqNo > AcqCount Then
					DlgEnd 1
				End If
				AcqMode = 2
				WaitingTimer = oPolyTimer.Time + WaitingTime
				LastTimer = oPolyTimer.Time + 1
			Else
				sProgress = DlgText("AcqProgress")
				sProgress = sProgress + C_Progress$
				If Len(sProgress) > C_MaxProgressChar% Then
					sProgress = C_Progress$
				End If
				DlgText "AcqProgress", sProgress
			End If
		ElseIf AcqMode = 2 Then
			Dim Timer0 As Double
			Timer0 = oPolyTimer.Time
			If WaitingTimer <= Timer0 Then
				DlgText "WaitingProgress", ""
				AcqMode = 0
			ElseIf LastTimer <= Timer0 Then
				LastTimer = oPolyTimer.Time + 1
				sProgress = DlgText("WaitingProgress")
				sProgress = sProgress + C_Progress$
				If Len(sProgress) > C_MaxProgressChar% Then
					sProgress = C_Progress$
				End If
				DlgText "WaitingProgress", sProgress
			End If
		End If
		Wait 0.5	' Give the application time to do the acquisition
		DoAcquisitionDlg = True
	Case 6 ' Function key
	End Select
End Function
'
' *******************************************************************************
' * Helper functions and subroutines
' *******************************************************************************
Const c_OFN_HIDEREADONLY As Long = 4

Private Function FileOpenDialog(sFileName As String, sFilter As String, sAcqType As String) As String
' -------------------------------------------------------------------------------
' Select file.
' -------------------------------------------------------------------------------
	On Error GoTo MCreateError
	Dim fod As Object
	Set fod = CreateObject("MSComDlg.CommonDialog")
	fod.FileName = sFileName
	fod.Filter = sFilter
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
	FileOpenDialog = GetFilePath(sFileName, Right$(sAcqType, 3), CurDir(), "Select a base file", 2)
MEnd:
End Function

Private Sub SplitPath(sFilePath As String, sPath As String, sName As String, sExt As String)
' -------------------------------------------------------------------------------
' Split file.
' -------------------------------------------------------------------------------
	Dim iBSlash As Integer
	Dim iDot As Integer

	iBSlash = InStrRev(sFilePath, "\")
	If iBSlash > 0 Then
		sPath = Left$(sFilePath, iBSlash)
	Else
		sPath = ".\"
	End If

	iBSlash = InStrRev(sFilePath, "\")
	sName = Right$(sFilePath, Len(sFilePath) - iBSlash)

	iDot = InStrRev(sName, ".")
	If iDot > 0 Then
		sExt = Right$(sName, Len(sName) - iDot)
		sName = Left$(sName, iDot - 1)
	Else
		sExt = ""
	End If
End Sub

Private Function TimeSpan(T As Date) As String
' -------------------------------------------------------------------------------
' Get time.
' -------------------------------------------------------------------------------
	If DatePart("h", T) > 0 Then
		TimeSpan = Format(T, "h \h  n \m\i\n")
	ElseIf DatePart("n", T) >= 1 Then
		TimeSpan = Format(T, "n \m\i\n s \s\e\c")
	Else
		TimeSpan = Format(T, "s \s\e\c")
	End If
End Function

Private Sub InitGlobalVariables
' -------------------------------------------------------------------------------
' Initialize variables.
' -------------------------------------------------------------------------------
	BackButtonPressed = False
	DlgAborted = False
	AcqBaseFileName(C_Scan%) = CurDir$() + "\Scan"
	AcqBaseFileName(C_SinglePoint%) = CurDir$() + "\Analyzer"
	AcqCount = 5
	WaitingTime = 0
	AcqStartTime = Time
	AcqStartDirect = True
	IsVibSoft = InStr(Name, "Scanning") = 0 And InStr(Name, "PSV") = 0
	If IsVibSoft Then
		AcqType = C_SinglePoint%
	Else
		AcqType = C_Scan%
	End If
	Set oPolyTimer = New  POLYTIMERLib.Timer
	oPolyTimer.Start
End Sub
