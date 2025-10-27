' This macro is used for the calibration of the laser offset for polytec microscopic scanning systems with positioning stage.
' This macro can only be used with UHF-120, OFV-534 and MSA-100-3D with positioning stage.

'#Uses "..\SwitchToAcquisitionMode.bas"
'#Uses "..\Common.bas"

Option Explicit


Enum Directions
	E_XDirection = 0
	E_YDirection = 1
End Enum

' the name of the current macro
Private Function GetMacroName() As String
	Dim strMacroName As String
	strMacroName = "Calibrate Laser Offset"
	GetMacroName = strMacroName
End Function

'Can the current system use a positioning stage as scanning device?
Private Function PositioningStageAsScanner As Boolean
	Dim IsMicroscopic As Boolean
	IsMicroscopic = Application.Acquisition.Infos.ScanHeadDevicesInfo.ScanHeadDevices.Item(1).Caps And ptcScanHeadCapsMicroscopic
	PositioningStageAsScanner = IsMicroscopic
End Function


Sub Main
	On Error GoTo CatchErr

	If Not SwitchToAcquisitionMode() Then
		MsgBox("Switch to acquisition mode failed.", vbOkOnly)
		End
	End If

'use only for PSV sytems with positioning stage as scanner
	If Not(PositioningStageAsScanner) Then
		MsgBox("You don't need to calibrate the laser offset for this system.", vbInformation Or vbOkOnly, GetMacroName)
		Exit Sub
	End If


'some explanations
	Dim iMsgBoxRet As Integer
	iMsgBoxRet = MsgBox( _
	"This macro is used for the calibration of the laser offset on systems with a positioning stage as scanning device." + vbCrLf _
	+ "First of all you have to define two scan points in the APS point mode:"  + vbCrLf _
	+ vbCrLf _
	+ vbTab + "Point one has to be defined at the nominal laser position." + vbCrLf _
	+ vbTab + vbTab + "Use 'Set Point'(Ctrl + L) to define it." + vbCrLf _
	+ vbCrLf _
	+ vbTab + "Point two has to be defined at the current laser position." + vbCrLf _
	+ vbTab + vbTab + "In the 'Create Point' mode Alt + left click on the laser spot." + vbCrLf _
	+ vbCrLf _
	+ "Refer to the PSV software manual for detailed description." + vbCrLf  _
	+ vbCrLf _
	+ "Continue with OK when you have defined the scan points or select CANCEL when you want to abort." _
	, vbOkCancel, GetMacroName)

	If (iMsgBoxRet = vbCancel) Then
		Exit Sub
	End If

'some helper variables
	Dim bFileAvailable As Boolean
	bFileAvailable   = False
	Dim bFileOpened    As Boolean
	bFileOpened      = False

'save the settings to a file and work on this file
	Dim strFileName As String
	strFileName = GetFileNameTemp()
	Application.Settings.Save(strFileName)
	bFileAvailable = True

'get a new polyfile instance
	Dim oFile As New PolyFile
	If oFile Is Nothing Then
		Err.Raise 1
	End If

'open the settings file
	oFile.ReadOnly = False
	oFile.Open(strFileName)
	bFileOpened = True

'get the infos collection from the file
	Dim oInfos As Infos
	Set oInfos = oFile.Infos

'look for videomapping
	If oInfos.HasVideoMappingInfo <> True Then
		MsgBox("This system has no laser offset and doesn't need a certain calibration.", vbCritical + vbOkOnly, GetMacroName)
		Err.Raise 1
	End If

'get the videomapping
	Dim oVideoMappingInfo As VideoMappingInfo
	Set oVideoMappingInfo = oInfos.VideoMappingInfo

	If oVideoMappingInfo.Calibrations.Count = 0 Then
		MsgBox("There is no calibrated lens available.", vbCritical + vbOkOnly, GetMacroName)
		Err.Raise 1
	End If

'get the current lens calibration
	Dim strActiveCalibration As String
	strActiveCalibration = oVideoMappingInfo.ActiveCalibration
	Dim oLensCalibration As LensCalibration
	Set oLensCalibration = oVideoMappingInfo.Calibrations.Item(strActiveCalibration)

'copy the magnification
	Dim dMagnification(1) As Double
	dMagnification(E_XDirection) = oLensCalibration.MagnificationFactorX
	dMagnification(E_YDirection) = oLensCalibration.MagnificationFactorY

'copy the chipdimension
	Dim dCameraChipDimension(1) As Double
	dCameraChipDimension(E_XDirection) = oVideoMappingInfo.CameraSize.ChipX
	dCameraChipDimension(E_YDirection) = oVideoMappingInfo.CameraSize.ChipY

'look for measpoints
	If oInfos.HasMeasPoints <> True Then
		MsgBox("Unexpected error, no scan points available.", vbCritical + vbOkOnly, GetMacroName)
		Err.Raise 1
	End If

'get the measurement points
	Dim oMeasPoints As MeasPoints
	Set oMeasPoints = oInfos.MeasPoints

'we expect that there are exactly 2 points defined
	If oMeasPoints.Count <> 2 Then
		MsgBox("Exactly 2 scan points are expected.", vbCritical + vbOkOnly, GetMacroName)
		Err.Raise 1
	End If


'we expect that the first point is at the center position
	Dim videoCoord(1) As Single
	oMeasPoints.Item(1).VideoXY(videoCoord(E_XDirection) ,videoCoord(E_YDirection))

	Dim dCenter(1) As Double
	oVideoMappingInfo.CoordXYFromVideo(videoCoord(E_XDirection) ,videoCoord(E_YDirection), dCenter(E_XDirection) ,dCenter(E_YDirection))

'we expect that the second point is at the laser position
	oMeasPoints.Item(2).VideoXY(videoCoord(E_XDirection) ,videoCoord(E_YDirection))

	Dim dLaser(1) As Double
	oVideoMappingInfo.CoordXYFromVideo(videoCoord(E_XDirection) ,videoCoord(E_YDirection), dLaser(E_XDirection) ,dLaser(E_YDirection))

'calculate the current normalized laser offset
	Dim dNormalizedLaserOffset(1) As Double
	Dim eDirection As Directions
	For eDirection = E_XDirection To E_YDirection ' do it for X and Y
'calculate the offset according to the  marked points
		dNormalizedLaserOffset(eDirection) = (dLaser(eDirection) - dCenter(eDirection) ) * dMagnification(eDirection)
'take the current registry value into account
		dNormalizedLaserOffset(eDirection) = dNormalizedLaserOffset(eDirection) + GetRegVal(eDirection)
'round the value
		dNormalizedLaserOffset(eDirection) = Round(dNormalizedLaserOffset(eDirection) * 1e7) * 1e-7
'check the value
		If Abs(dNormalizedLaserOffset(eDirection)) > dCameraChipDimension(eDirection)/2 Then
			MsgBox("The measured laser offset is greather than the corresponding chip dimension!", vbCritical + vbOkOnly, GetMacroName)
			Err.Raise 1
		End If
'save the offset to the registry
		SetRegVal(dNormalizedLaserOffset(eDirection),eDirection)
	Next eDirection

'clear the marked points to avoid this operation again
	oMeasPoints.Clear()

'save the file
	oFile.Save()
	oFile.Close()
	bFileOpened = False

'load the measurement point to PSV again
	Application.Settings.Load(strFileName,ptcSettingsAPS)

'delete the file
	Kill(strFileName)
	bFileAvailable = False


'Success output
	Dim StrMessage As String
	StrMessage = "The values "  + vbTab + ToString(dNormalizedLaserOffset(E_XDirection)) + "m" + vbCrLf + _
	             "and " + vbTab + vbTab + ToString(dNormalizedLaserOffset(E_YDirection)) + "m" + vbCrLf + _
	             "are stored as x and y offset."
   	MsgBox(StrMessage, vbOkOnly, GetMacroName)

'PSV will read these values on reload
	Reload

'successfull
	Exit Sub

' error
CatchErr:
	If (bFileOpened = True) Then
		oFile.Close()
		bFileOpened = False
	End If
	If (bFileAvailable = True) Then
		Kill(strFileName)
        bFileAvailable = False
	End If

	MsgBox("The laser offset calibration was not successfull.", vbCritical + vbOkOnly, GetMacroName)

End Sub

Function RegSetValueString(ByVal hKey As PortInt, ByVal ValueName As String, ByVal Data As String) As Long
	Dim length As Long
	length = Len(Data) + 1

	RegSetValueString = RegSetValueStrEx(hKey, ValueName, 0, REG_SZ, Data, length)
End Function

Private Function RegQueryStringValue(ByVal hKey As PortInt, ByVal strValueName As String) As String
   	Dim errorCode As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    'retrieve information about the key
    lDataBufSize = 256
    'Create a buffer
    strBuf = String(lDataBufSize, Chr$(0))
    'retrieve the key's content
    errorCode = RegQueryValueStrEx(hKey, strValueName, 0, 0, strBuf, lDataBufSize)
	If errorCode <> ERROR_SUCCESS Then
    	RegQueryStringValue = ""
    Else
	    'Remove the unnecessary chr$(0)'s
	    RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
    End If
End Function

Private Function GetRegPath As String
	 GetRegPath = "SOFTWARE\Polytec\PSV\Scan"
End Function

Private Function GetRegVal(ByVal eDirection As Directions) As Double
    Dim hKey As PortInt
    Dim dwDisposition As Long
	Dim errorCode As Long
	hKey = 0

	errorCode = RegCreateKeyEx(HKEY_LOCAL_MACHINE, GetRegPath, 0, 0, 0, KEY_QUERY_VALUE, 0, hKey, dwDisposition)
	If errorCode <> ERROR_SUCCESS Then
    	MsgBox(APIErrorDescription(errorCode), vbCritical + vbOkOnly, GetMacroName)
		Err.Raise 1
    End If

    Dim StrVal As String
	StrVal = RegQueryStringValue(hKey,GetValueName(eDirection))
	RegCloseKey(hKey)

	On Error GoTo CatchErr
	GetRegVal = ToDouble(StrVal)
	Exit Function

CatchErr:
	GetRegVal = 0.0
	If hKey <> 0 Then
		RegCloseKey(hKey)
	End If

End Function


Private Function SetRegVal(ByVal dValue As Double, ByVal eDirection As Directions) As Long
    Dim hKey As PortInt
    Dim dwDisposition As Long
	Dim errorCode As Long
	hKey = 0

	errorCode = RegCreateKeyEx(HKEY_LOCAL_MACHINE, GetRegPath, 0, 0, 0, KEY_SET_VALUE, 0, hKey, dwDisposition)
	If errorCode <> ERROR_SUCCESS Then
    	MsgBox(APIErrorDescription(errorCode), vbCritical + vbOkOnly, GetMacroName)
		Err.Raise 1
    End If

	errorCode = RegSetValueString(hKey,GetValueName(eDirection),ToString(dValue))
	If errorCode <> ERROR_SUCCESS Then
    	MsgBox(APIErrorDescription(errorCode), vbCritical + vbOkOnly, GetMacroName)
		Err.Raise 1
    End If

	RegCloseKey(hKey)
	On Error GoTo CatchErr
	Exit Function

CatchErr:
	If hKey <> 0 Then
		RegCloseKey(hKey)
	End If

End Function


Private Function GetValueName(ByVal eDirection As Directions) As String
	Dim strValueName As String

	Select Case eDirection
	Case E_XDirection
		GetValueName= "LaserXStartPosNormalized"
	Case E_YDirection
		GetValueName= "LaserYStartPosNormalized"
	Case Else
        MsgBox("Logic error", vbCritical + vbOkOnly, GetMacroName)
		Err.Raise 1
	End Select

End Function

Private Function GetFileNameTemp As String
	Dim ret As Long, length As Long
	Dim strTempDir As String, strTempFile As String

    'get temporary directory
    strTempDir = String$(255, 0)
    length = Len(strTempDir)
    ret = GetTempPath(length, strTempDir)

    If ret <> ERROR_SUCCESS Then
        strTempDir = Left$(strTempDir, ret)
    Else
    	MsgBox("Could not locate temporary directory.", vbCritical + vbOkOnly, GetMacroName)
		Err.Raise 1
    End If

    'create file name
    strTempFile = String$(255, 0)
    ret = GetTempFileName(strTempDir, "psv", 0, strTempFile)

    If ret <> ERROR_SUCCESS Then
        GetFileNameTemp = Left$(strTempFile, InStr(1, strTempFile,".tmp") ) + "set"
    Else
    	MsgBox("Could not get temporary file name.", vbCritical + vbOkOnly, GetMacroName)
		Err.Raise 1
    End If

End Function

'Reload the saved registry values by switching the mode
Private Sub Reload()

'there is no direct method to reload the calibration into PSV --> we will switch to presentation mode
'an then again to acquisition mode
'if this failes we will quit PSV and the user has to manually restart PSV in acquisition mode

	On Error GoTo RestartManually 'if application mode couldn't be changed jump to the flag
	Application.Mode = ptcApplicationModePresentation
	Application.Mode = ptcApplicationModeAcquisition

	Exit Sub

'PSV will be closed.
RestartManually:
   	MsgBox("PSV will be closed now. Please restart it manually.", vbOkOnly, GetMacroName)
   	Application.Quit
End Sub

Private Function GetSeperator() As String
	GetSeperator = Format$(0, ".")  '--> http://us.generation-nt.com/answer/api-function-determin-decimal-seperator-help-8753082.html
End Function

Private Function ToString(ByVal dVal As Double) As String
    Dim StrVal As String
	StrVal = CStr(dVal)
	ToString = Replace(StrVal, GetSeperator(), ".")
End Function

Private Function ToDouble(ByVal StrVal As String) As Double
	StrVal = Replace(StrVal, ".", GetSeperator())
	ToDouble = CDbl(StrVal)
End Function
