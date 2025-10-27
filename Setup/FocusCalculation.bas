' POLYTEC HELPER MACRO
' ----------------------------------------------------------------------
' Deutsche Version: siehe unten
'
' On PSV-400 systems this macro calculates:
'   - the focal length of the lens and
'   - the infinite focus position
' for each scanning head.
'
' On all other systems, it is not necessary to use this macro.
'
' Those parameters are required if you want to enable the PSV software to calculate the focus values over the
' full range of distances.
'
' The "focal length of the lens" and the "infinite focus position" can be displayed and modified in the page "Geometry"
' of the "Preferences" dialog. The button "Calculate" on this page only calculates the "focal length of the lens".
' The calculation of the "focal length of the lens" requires a correct specification of the "infinite focus position".
' The only way to calculate the "infinite focus position" is the usage of this macro.
'
' To use this macro, you need two pairs of distance / focus values for each scanning head.
'
' Please perform the following procedure:
' 1.	Define a scan point in a distance of ~500 mm (100 mm for MR lens)
' 2.	Measure the distance between scan point and scanning head with an accuracy of 1%
'       (by either geometry scan unit or tape measure. The distance is measured
'       from the scan point to the front plate of the scanning head)
' 3.	Perform an autofocus run (make sure the laser is well focused)
' 4.	Read the focus value (value in brackets from the tooltip of
'       the autofocus button in the scanning head control bar)
' 5.	Repeat steps 1-4 for a distance of >2100 mm (1100 mm for MR lens)
' 6.	Run this macro
' 7.	Enter the two pairs of values
' 8.	Press Calculate
' 9.	Press Save
' 10.	Press OK to leave the macro
' 11.	For PSV-3D: repeat steps 1-10 for scanning heads Left and Right
'       please note: the geometry scan unit cannot be used for step 2.
'                    Please use a tape measure instead.
'
'-----------------------------------------------------------------------
' Auf PSV-400 Systemen berechnet dieses Makro:
'   - die Brennweite der Linse und
'   - die Fokusposition für Unendlich
' für jeden Scankopf.
'
' Auf allen anderen Systemen ist es nicht notwendig, dieses Makro zu verwenden.
'
' Diese Parameter werden benötigt, damit die PSV-Software Fokuswerte über den gesamten Bereich
' der Messabstände berechnen kann.
'
' Die "Brennweite der Linse" und die "Fokusposition für Unendlich" können auf der Seite "Geometrie" im Dialog "Optionen"
' angezeigt und verändert werden. Die Schaltfläche "Berechnen" auf dieser Seite berechnet nur die "Brennweite der Linse",
' hierbei wird davon ausgegangen, dass die "Fokusposition für Unendlich" korrekt ist. Die Berechnung der "Fokusposition
' für Unendlich" kann nur über dieses Makro erfolgen.
'
' Um dieses Makro zu verwenden, benötigen Sie zwei Paare von Abständen / Fokuswerten für jeden Scankopf.
'
' Bitte führen Sie die folgenden Schritte durch:
' 1.	Definieren Sie einen Scanpunkt in einem Abstand von ~500 mm (100 mm bei MR Optik)
' 2.	Messen Sie den Abstand zwischen Scanpunkt und Scankopf mit einer Genauigkeit von 1%
'       (entweder mit der Geometrie-Scaneinheit oder mit einem Maßband. Der Abstand wird
'       zwischen dem Scanpunkt und der Frontplatte des Scankopfs gemessen)
' 3.	Fokussieren Sie den Laser automatisch (verifizieren Sie, dass der Laser gut fokussiert ist)
' 4.	Lesen Sie den Fokuswert (Wert in Klammern im Tooltip der Autofokus Schaltfläche in der
'       Scankopf Steuerung der PSV Software)
' 5.	Wiederholen Sie die Schritte 1-4 für einen Messabstand >2100 mm (1100 mm bei MR Optik)
' 6.	Führen Sie dieses Makro aus
' 7.	Geben Sie die zwei Wertepaare ein
' 8.	Drücken Sie Calculate
' 9.	Drücken Sie Save
' 10.	Drücken Sie OK um das Makro zu beenden
' 11.	Beim PSV-3D: Wiederholen Sie die Schritte 1-10 für die Scanköpfe Links und Rechts
'       bitte beachten Sie: Die Geometrie-Scaneinheit kann für Punkt 2 nicht eingesetzt werden.
'                           Bitte benutzen Sie stattdessen ein Maßband.
'
'-----------------------------------------------------------------------

'#Uses "..\SwitchToAcquisitionMode.bas"
'#Uses "..\Common.bas"

Option Explicit

' Data structure describing the current state of the input fields and the
' calculated results. Modified is true, when the corresponding results
' were calculated.
Private Type Params
	Distance1 As String
	Distance2 As String
	FocusValue1 As String
	FocusValue2 As String
	FocalLength As Double
	FocusInf As Long
	Modified As Boolean
End Type

' Parameters for all scanning heads (top, left, right)
Dim g_Params(0 To 2) As Params
' Current scanning head index (0 = top, 1 = left, 2 = right)
Dim g_lScanHeadIndex As Long

Sub Main
	If Not SwitchToAcquisitionMode() Then
        MsgBox "Switch to acquisition mode failed."
        Exit Sub
    End If

	If Not IsPSV400() Then
        If IsPSV500() Then
			MsgBox "Normally it is not necessary to run this macro for PSV-500 systems."
        else
			MsgBox "This macro is not useful for your system."
	        Exit Sub									  
        End If
	End If

    ' we need administrative rights to update the registry
	If (Not CheckRegistryAccess()) Then
		Exit Sub
	End If
	Call ShowDialog()
End Sub

' displays the main dialog
Private Sub ShowDialog()
	Begin Dialog UserDialog 780,371,"Focus Calculation",.DlgProc ' %GRID:10,7,1,1
		Text 20,7,740,14,"Use this macro if you want to enable the PSV software to calculate the focus values over the full range of distances."
		Text 20,28,730,42,"The PSV software calculates laser focus values from the distances of the scan points to the scanning head using the focal length and the focus value for infinity of the scanning head. You can update the parameters for this calculation by specifying two distance / focus value pairs.",.Text1
		Text 20,70,730,28,"For a detailed step by step description of the procedure please look at the comments at the beginning of the source code of this macro.",.Text13
		Text 20,112,140,14,"Scanning Head:",.ScanHeadLabel
		Text 190,112,160,14,"PSV-I-400 LR",.ScanHeadName
		OptionGroup .ScanHead
			OptionButton 390,112,100,14,"Top",.Top
			OptionButton 500,112,100,14,"Left",.Left
			OptionButton 610,112,110,14,"Right",.Right
		Text 120,147,140,14,"Distance 1:",.Distance1Label1
		TextBox 220,147,80,16,.Distance1
		Text 320,147,50,14,"mm",.Text2
		Text 380,147,130,14,"Focus Value 1:",.Text4
		TextBox 500,147,80,16,.FocusValue1
		Text 120,175,130,14,"Distance 2:",.Distance1Label2
		TextBox 220,175,80,16,.Distance2
		Text 320,175,50,14,"mm",.Text3
		Text 380,175,130,14,"Focus Value 2:",.Text5
		TextBox 500,175,80,16,.FocusValue2
		Text 20,210,720,14,"For optimal results, focus value 1 should be smaller than 1000 and focus value 2 should be larger than 2400.",.Text6
		PushButton 620,231,130,21,"Calculate",.Calculate
		Text 260,266,170,21,"Results:",.Text7
		Text 440,266,210,21,"Curent Settings:",.Text8
		Text 120,291,140,14,"Focal Length:",.Text9
		Text 260,291,80,14,"",.FocalLengthResult
		Text 440,291,80,14,"",.FocalLengthCurrent
		Text 340,291,40,14,"mm",.Text10
		Text 520,291,40,14,"mm",.Text11
		Text 120,316,140,14,"Focus for Infinity:",.Text12
		Text 260,316,80,14,"",.FocusInfResult
		Text 440,316,80,14,"",.FocusInfCurrent
		PushButton 480,343,130,21,"Save",.Save
		OKButton 620,343,130,21
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
End Sub

' checks if the system has multiple scan heads, i.e. is a PSV-3D system
Private Function HasMultipleScanHeads() As Boolean
	HasMultipleScanHeads = Acquisition.Infos.ScanHeadDevicesInfo.ScanHeadDevices.Count = 3
End Function

' returns the name of the scanning head
Private Function ScanHeadName() As String
	ScanHeadName = Acquisition.Infos.ScanHeadDevicesInfo.ScanHeadDevices.ScanHeadType.Name
End Function

' returns the type of the scanning head
Private Function ScanHeadType() As PTCScanHeadType
	ScanHeadType = Acquisition.Infos.ScanHeadDevicesInfo.ScanHeadDevices.ScanHeadType.type
End Function

' returns true if it is a PSV-400 scanning head
Private Function IsPSV400() As Boolean
	IsPSV400 = ((ScanHeadType() = ptcScanHeadTypePSVI400_LR ) Or (ScanHeadType() = ptcScanHeadTypePSVI400_MR))
End Function

' returns true if it is a PSV-500 scanning head
Private Function IsPSV500() As Boolean
	IsPSV500 = (ScanHeadType() = ptcScanHeadTypePSV500)
End Function

Function RegSetValueString(ByVal hKey As PortInt, ByVal ValueName As String, ByVal Data As String) As Long
	Dim length As Long
	length = Len(Data) + 1

	RegSetValueString = RegSetValueStrEx(hKey, ValueName, 0, REG_SZ, Data, length)
End Function

Function RegSetValueLong(ByVal hKey As PortInt, ByVal ValueName As String, ByRef Data As Long) As Long
	RegSetValueLong = RegSetValueLngEx(hKey, ValueName, 0, REG_DWORD, Data, 4)
End Function

Function RegSetValueBool(ByVal hKey As PortInt, ByVal ValueName As String, ByRef Data As Boolean) As Long
	RegSetValueBool = RegSetValueBoolEx(hKey, ValueName, 0, REG_BINARY, Data, 1)
End Function

Function RegQueryStringValue(ByVal hKey As PortInt, ByVal strValueName As String) As String
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    'retrieve information about the key
    lDataBufSize = 256
    'Create a buffer
    strBuf = String(lDataBufSize, Chr$(0))
    'retrieve the key's content
    lResult = RegQueryValueStrEx(hKey, strValueName, 0, 0, strBuf, lDataBufSize)
    If lResult = 0 Then
        'Remove the unnecessary chr$(0)'s
        RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
    End If
End Function

' gets the focal parameters for the given scanning head (focal length and focus position for infinity)
' from the registry
Private Sub GetFocalParameters(lScanHeadIndex As Long, ByRef dFocalLength As Double, ByRef lFocusInf As Long)
	Dim strScanHeadName As String
	strScanHeadName = ScanHeadName()

    Dim hKey As PortInt

    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\Polytec\PolyScanHead\" + strScanHeadName, 0, KEY_QUERY_VALUE, hKey) <> ERROR_SUCCESS Then
         Err.Raise(1, "GetFocalParameters", "Could not access registry")
    End If

	Dim strFocusPosInfinite As String
	strFocusPosInfinite = RegQueryStringValue(hKey, "FocusPosInfinite")

	lFocusInf = 0
	If (Len(strFocusPosInfinite) > 0) Then
		Dim strFocusInf() As String
		strFocusInf = Split(strFocusPosInfinite, ";")
		If (UBound(strFocusInf) <> Acquisition.Infos.ScanHeadDevicesInfo.ScanHeadDevices.Count - 1) Then
			RegCloseKey(hKey)
        	Err.Raise(1, "GetFocalParameters", "FocusPosInfinite has invalid format: " + strFocusPosInfinite)
		End If
		lFocusInf = CLng(strFocusInf(lScanHeadIndex))
	End If

	Dim strFocalLengthKey As String
	Select Case lScanHeadIndex
		Case 0
			strFocalLengthKey = "FocalLengthTop"
		Case 1
			strFocalLengthKey = "FocalLengthLeft"
		Case 2
			strFocalLengthKey = "FocalLengthRight"
	End Select

	Dim strFocalLength As String
	strFocalLength = RegQueryStringValue(hKey, strFocalLengthKey)


	dFocalLength = 0.0
	If (Len(strFocalLength) > 0) Then
		dFocalLength = String2Dbl(strFocalLength)
	End If

    RegCloseKey(hKey)
End Sub

' checks write access to the registry
Private Function CheckRegistryAccess() As Boolean
	Dim strScanHeadName As String
	strScanHeadName = ScanHeadName()

    Dim hKey As PortInt
    Dim dwDisposition As Long

    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\Polytec\PolyScanHead\" + strScanHeadName, 0, KEY_SET_VALUE, hKey) <> ERROR_SUCCESS Then
    	If RegCreateKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\Polytec\PolyScanHead\" + strScanHeadName, 0, 0, 0, KEY_SET_VALUE, 0, hKey, dwDisposition) <> ERROR_SUCCESS Then
        	MsgBox("Could not access registry for writing. Administrative rights are required. The macro finishes now.", vbCritical + vbOkOnly, "Focus Calculation")
         	CheckRegistryAccess = False
         	Exit Function
        End If
    End If

    RegCloseKey(hKey)

    CheckRegistryAccess = True
End Function

' saves the focal parameters (focal length and focus position for infinity) of all scanning heads
' were the settings were modified to the registry
Private Sub SaveFocalParameters()

	Dim strScanHeadName As String
	strScanHeadName = ScanHeadName()

    Dim hKey As PortInt

    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\Polytec\PolyScanHead\" + strScanHeadName, 0, KEY_SET_VALUE + KEY_QUERY_VALUE, hKey) <> ERROR_SUCCESS Then
         Err.Raise(1, "SaveFocalParameters", "Could not access registry for writing. Administrative rights are required.")
    End If

	Dim strFocusPosInfinite As String
	strFocusPosInfinite = RegQueryStringValue(hKey, "FocusPosInfinite")
	If (Len(strFocusPosInfinite) = 0) Then
		strFocusPosInfinite = "0;0;0"
	End If

	Dim scanHeadCount As Long
	scanHeadCount = Acquisition.Infos.ScanHeadDevicesInfo.ScanHeadDevices.Count

	Dim i As Long
	Dim lFocusInfs(0 To 2) As Long
	Dim strFocusInf() As String
	strFocusInf = Split(strFocusPosInfinite, ";")
	If (UBound(strFocusInf) <> scanHeadCount - 1) Then
		RegCloseKey(hKey)
    	Err.Raise(1, "SaveFocalParameters", "FocusPosInfinite has invalid format: " + strFocusPosInfinite)
	End If

	For i = 0 To scanHeadCount - 1
		If (g_Params(i).Modified) Then
			If (Not CheckInfiniteValue(g_Params(i).FocusInf)) Then
				g_Params(i).Modified = False 'reset modified flag
				RegCloseKey(hKey)
				Exit Sub
			End If
			lFocusInfs(i) = g_Params(i).FocusInf
		Else
			lFocusInfs(i) = CLng(strFocusInf(i))
		End If
	Next i

	strFocusPosInfinite = CStr(lFocusInfs(0))
	For i = 1 To scanHeadCount - 1
		strFocusPosInfinite = strFocusPosInfinite + ";" + CStr(lFocusInfs(i))
	Next i

	RegSetValueString(hKey, "FocusPosInfinite", strFocusPosInfinite)

	For i = 0 To scanHeadCount - 1
		If (g_Params(i).Modified) Then
			Dim strFocalLengthKey As String
			Select Case i
				Case 0
					strFocalLengthKey = "FocalLengthTop"
				Case 1
					strFocalLengthKey = "FocalLengthLeft"
				Case 2
					strFocalLengthKey = "FocalLengthRight"
			End Select
			Call RegSetValueString(hKey, strFocalLengthKey, Dbl2String(g_Params(i).FocalLength))
			g_Params(i).Modified = False
		End If
	Next i

    RegCloseKey(hKey)
End Sub

' formats a distance as mm
Private Function FormatMM(dDistance As Double) As String
	FormatMM = Format(1000.0 * dDistance, "########0.000")
End Function

' saves the current input values for the given scanning head
Private Sub SaveParams(lScanHeadIndex As Long)
	With g_Params(lScanHeadIndex)
		.Distance1 = DlgText("Distance1")
		.Distance2 = DlgText("Distance2")
		.FocusValue1 = DlgText("FocusValue1")
		.FocusValue2 = DlgText("FocusValue2")
	End With
End Sub

' loads the current input values for the given scanning head
' if the results are valid, those are loaded, too
Private Sub LoadParams(lScanHeadIndex As Long)
	With g_Params(lScanHeadIndex)
		DlgText "Distance1",.Distance1
		DlgText "Distance2",.Distance2
		DlgText "FocusValue1",.FocusValue1
		DlgText "FocusValue2",.FocusValue2
		If (.Modified) Then
			DlgText "FocalLengthResult", FormatMM(.FocalLength)
			DlgText "FocusInfResult", CStr(.FocusInf)
		Else
			DlgText "FocalLengthResult", ""
			DlgText "FocusInfResult", ""
		End If
	End With
End Sub

' updates the fields of the dialog
Private Sub UpdateDlg()
	Dim bHasMultipleScanHeads As Boolean
	bHasMultipleScanHeads = HasMultipleScanHeads()

	DlgVisible "Top", bHasMultipleScanHeads
	DlgVisible "Left", bHasMultipleScanHeads
	DlgVisible "Right", bHasMultipleScanHeads

	DlgText "ScanHeadName", ScanHeadName()

	LoadParams(g_lScanHeadIndex)

	Dim lScanHeadIndex As Long
	lScanHeadIndex = DlgValue("ScanHead")

	Dim dFocalLength As Double
	Dim lFocusInf As Long
	Call GetFocalParameters(lScanHeadIndex, dFocalLength, lFocusInf)

	DlgText "FocalLengthCurrent", FormatMM(dFocalLength)
	DlgText "FocusInfCurrent", CStr(lFocusInf)
End Sub

' returns delta s. this value depends On the scanning head
Private Function GetDelta_s() As Double
	If IsPSV400() Then
		GetDelta_s = 1e-5
	ElseIf IsPSV500 Then
		GetDelta_s = 2.5133e-6
	Else
		Err.Raise(1, "DlgProc", "Neither a PSV-400, nor a PSV-500: " + ScanHeadName())
	End If
End Function

' calculates the focal length and focus position for infinity from the two distance / focus value pairs
Private Sub Calculate(z1 As Double, p1 As Double, z2 As Double, p2 As Double, ByRef f As Double, ByRef p0 As Double)
	Dim a As Double
	Dim b As Double
	Dim c As Double
	Dim delta_s As Double

	delta_s = GetDelta_s()

	a = z2-z1+delta_s*(p1-p2)
	b = -delta_s*(z2+z1)*(p1-p2)
	c = delta_s*z2*z1*(p1-p2)

	If (z2 > z1) Then
		f = (-b+Sqr(b^2-4*a*c))/(2*a)
	Else
		f = (-b-Sqr(b^2-4*a*c))/(2*a)
	End If

	p0 = (f*f-delta_s*f*p2+delta_s*z2*p2)/(delta_s*z2-delta_s*f)
End Sub

' checks if the given focal length deviates by more than +-5% from the nominal value
Private Sub CheckFocalLength(dFocalLength As Double)
	Dim oScanHeadDevice As ScanHeadDevice
	Set oScanHeadDevice = Acquisition.Infos.ScanHeadDevicesInfo.ScanHeadDevices(g_lScanHeadIndex + 1)

	Dim dRatio As Double
	dRatio = dFocalLength / oScanHeadDevice.Focus.FocalLength

	If (dRatio < 0.95 Or dRatio > 1.05) Then
		MsgBox("The calculated focal length (" + FormatMM(dFocalLength) + " mm) " _
		+ "deviates more than 5% from the nominal focal length (" _
		+ FormatMM(oScanHeadDevice.Focus.FocalLength) + " mm) of the scanning head." + vbCrLf _
		+ "Please check the input values and the setting of your scanning head" _
		+ " (" + ScanHeadName() + ")." _
		, vbInformation + vbOkOnly, "Focus Calculation")
	End If
End Sub

Private Function ValidateDouble(strDlgItem As String, strName As String, ByRef dValue As Double) As Boolean
	ValidateDouble = False
	If Mid(CStr(1.1), 2, 1) = "." Then
		If InStr(DlgText(strDlgItem), ",") > 0 Then
			MsgBox("Please use . as decimal symbol")
			DlgFocus(strDlgItem)
			Exit Function
		End If
	Else
		If InStr(DlgText(strDlgItem), ".") > 0 Then
			MsgBox("Please use , as decimal symbol")
			DlgFocus(strDlgItem)
			Exit Function
		End If
	End If

	On Error Resume Next
	dValue = CDbl(DlgText(strDlgItem))
	If (Err.Number <> 0) Then
		On Error GoTo 0
		MsgBox("Please enter a valid number for " + strName)
		DlgFocus(strDlgItem)
		Exit Function
	End If
	On Error GoTo 0

	ValidateDouble = True
End Function

Private Function Dbl2String(d As Double) As String
	Dbl2String = Replace(CStr(d), ",", ".")
End Function


Private Function String2Dbl(s As String) As Double
	If Mid(CStr(1.1), 2, 1) = "," Then
		If InStr(s, ".") > 0 Then
			s = Replace(s, ".", ",")
		End If
	Else
		If InStr(s, ",") > 0 Then
			s = Replace(s, ",", ".")
		End If
	End If
	String2Dbl = CDbl(s)
End Function

' returns the z offset. this vaue depends On the scanning head
Private Function GetZOffset() As Double

	If IsPSV400() Then
		If (ScanHeadType() = ptcScanHeadTypePSVI400_LR) Then
			GetZOffset = 0.1159
		ElseIf (ScanHeadType() = ptcScanHeadTypePSVI400_MR) Then
			GetZOffset = 0.15558
		Else
			Err.Raise(1, "DlgProc", "Neither a PSV-I-400-LR, nor a PSV-I-400-MR scanning head: " + ScanHeadName())
		End If

	ElseIf IsPSV500 Then
		GetZOffset = 0.21985

	Else
		Err.Raise(1, "DlgProc", "Neither a PSV-I-400-LR, nor a PSV-I-400-MR, nor a PSV-I-500 scanning head: " + ScanHeadName())
	End If
End Function

' dialog procedure of main dialog
' handles all button clicks
Private Function DlgProc(DlgItem$, Action%, SuppValue&) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgValue "ScanHead",0
		Call UpdateDlg()
		DlgFocus("Distance1")
	Case 2 ' Value changing or button pressed
		If DlgItem$ = "ScanHead" Then
			SaveParams(g_lScanHeadIndex)
			g_lScanHeadIndex = DlgValue("ScanHead")
			Call UpdateDlg()
		ElseIf DlgItem$ = "Calculate" Then
			DlgProc = True

			Dim z1 As Double
			Dim z2 As Double
			Dim p1 As Double
			Dim p2 As Double

			If (Not ValidateDouble("Distance1", "Distance 1", z1)) Then Exit Function
			z1 = z1 * 1e-3
			If (Not ValidateDouble("Distance2", "Distance 2", z2)) Then Exit Function
			z2 = z2 * 1e-3
			If (Not ValidateDouble("FocusValue1", "Focus Value 1", p1)) Then Exit Function
			If (Not ValidateDouble("FocusValue2", "Focus Value 2", p2)) Then Exit Function

			Dim f As Double
			Dim p0 As Double

			Dim zOffset As Double
			zOffset = GetZOffset()

			On Error Resume Next
			Call Calculate(z1 + zOffset, p1, z2 + zOffset, p2, f, p0)
			If (Err.Number <> 0) Then
				On Error GoTo 0
				MsgBox("The calculation of the focal parameters did not succeeed." + vbCrLf + _
					"Please check the input values.", vbInformation + vbOkOnly, "Focus Calculation")
				Exit Function
			End If
			On Error GoTo 0


			Call CheckFocalLength(f)

			DlgText "FocalLengthResult", FormatMM(f)
			DlgText "FocusInfResult", CStr(CLng(p0))

			Dim lScanHead As Long
			lScanHead = DlgValue("ScanHead")

			g_Params(lScanHead).FocalLength = f
			g_Params(lScanHead).FocusInf = p0
			g_Params(lScanHead).Modified = True

		ElseIf DlgItem$ = "Save" Then
			Dim i As Long
			Dim bModified As Boolean
			For i = LBound(g_Params) To UBound(g_Params)
				If (g_Params(i).Modified) Then
					bModified = True
				End If
			Next i
			If (bModified) Then
				Call SaveFocalParameters()
				Call SaveParams(g_lScanHeadIndex)
				Call UpdateDlg()
			Else
				MsgBox("Nothing to save. Please click on Calculate first", vbOkOnly)
			End If
			DlgProc = True
		ElseIf DlgItem$ = "OK" Then
			For i = LBound(g_Params) To UBound(g_Params)
				If (g_Params(i).Modified) Then
					bModified = True
				End If
			Next i
			If (bModified) Then
				Select Case MsgBox("The calculated results have not been saved. Do you want to save them?", vbYesNoCancel, "Save modified results?")
					Case vbYes
						Call SaveFocalParameters()
						DlgProc = False
					Case vbNo
						DlgProc = False
					Case vbCancel
						DlgProc = True
				End Select
			End If
		End If
	End Select
End Function

' checks if the given infinite value deviates by more than +-200 from the default value
Private Function CheckInfiniteValue(lInfiniteValue As Long) As Boolean
	CheckInfiniteValue = True
	Dim oScanHeadDevice As ScanHeadDevice
	Set oScanHeadDevice = Acquisition.Infos.ScanHeadDevicesInfo.ScanHeadDevices(g_lScanHeadIndex + 1)

	Dim lFocusPosInfiniteDef As Long
	lFocusPosInfiniteDef = oScanHeadDevice.Focus.FocusInfinitePos

	If (lInfiniteValue > lFocusPosInfiniteDef + 200 Or lInfiniteValue < lFocusPosInfiniteDef - 200) Then
		MsgBox("The calculated focus for infinite (" + CStr(lInfiniteValue) + ") " _
		+ "deviates more than 200 from the default infinite value (" _
		+ CStr(lFocusPosInfiniteDef) + ") of the scanning head." + vbCrLf _
		+ "Please check the input values and the setting of your scanning head" _
		+ " (" + ScanHeadName() + ")." _
		, vbCritical  + vbOkOnly, "Focus Calculation")
		CheckInfiniteValue = False
	End If
End Function


