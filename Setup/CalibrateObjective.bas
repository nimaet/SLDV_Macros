' This macro is used for the calibration of the objective magnification for polytec microscopic scanning vibrometers.
' This macro can only be used with MSA, MSV, UHF-120 and OFV-534 with positioning stage.

'#Uses "..\SwitchToAcquisitionMode.bas"
'#Uses "..\Common.bas"

Option Explicit

Enum Directions
	E_XDirection = 0
	E_YDirection = 1
End Enum

' the name of the current macro
Function GetMacroName() As String
	Dim strMacroName As String
	strMacroName = "Calibrate objective"
	GetMacroName = strMacroName
End Function

'Is the current system a multilens system?
Function IsMicroscopic As Boolean
	Dim oScanHeadDeviceCap As PTCScanHeadCapsType
	oScanHeadDeviceCap = Application.Acquisition.Infos.ScanHeadDevicesInfo.ScanHeadDevices.Item(1).Caps

	IsMicroscopic = ptcScanHeadCapsMicroscopic And oScanHeadDeviceCap
End Function

Sub Main
	On Error GoTo CatchErr
	
	If Not SwitchToAcquisitionMode() Then
		MsgBox("Switch to acquisition mode failed.", vbOkOnly)
		End
	End If
	
'Use only if this is a multilens system
	If Not(IsMicroscopic) Then
		MsgBox("You don't need to calibrate objective magnification for this system.", vbInformation Or vbOkOnly, GetMacroName)
		Exit Sub
	End If

'Some explanations
	Dim iMsgBoxRet As Integer
	iMsgBoxRet = MsgBox( _
	"This macro supports the objective magnification calibration for polytec microscopic scanning vibrometers." + vbCrLf _
	+ "It will calculate the pixel distance of two scan points." + vbCrLf _
	+ "Before use, you have to define two scan points with known x and y distance, " + vbCrLf _
	+ "e.g. by using the stage micrometer supplied."  + vbCrLf _
	+ "If you have a positioning stage as scanning device press ALT while definig the scan points to avoid movement."  + vbCrLf _
	+ vbCrLf _
	+ "Refer to the PSV software manual for detailed description." + vbCrLf _
	+ vbCrLf  _
	+ "Continue with OK when you have defined the scan points or select CANCEL when you want to abort." _
	, (vbInformation Or vbOkCancel), GetMacroName())

	If (iMsgBoxRet = vbCancel) Then
		Exit Sub
	End If

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
		MsgBox("Unexpected error, no video mapping available.", vbCritical + vbOkOnly, GetMacroName())
		Err.Raise 1
	End If

'get the videomapping
	Dim oVideoMappingInfo As VideoMappingInfo
	Set oVideoMappingInfo = oInfos.VideoMappingInfo

	If oVideoMappingInfo.Calibrations.Count = 0 Then
		MsgBox("There is no calibrated objective available.", vbCritical + vbOkOnly, GetMacroName)
		Err.Raise 1
	End If

'get the current objective calibration
	Dim strActiveCalibration As String
	strActiveCalibration = oVideoMappingInfo.ActiveCalibration
	Dim oLensCalibration As LensCalibration
	Set oLensCalibration = oVideoMappingInfo.Calibrations.Item(strActiveCalibration)

'copy the current magnification
	Dim dCurrentMagnification(1) As Double
	dCurrentMagnification(E_XDirection) = oLensCalibration.MagnificationFactorX
	dCurrentMagnification(E_YDirection) = oLensCalibration.MagnificationFactorY

'copy the chipdimension
	Dim dCameraChipDimension(1) As Double
	dCameraChipDimension(E_XDirection) = oVideoMappingInfo.CameraSize.ChipX
	dCameraChipDimension(E_YDirection) = oVideoMappingInfo.CameraSize.ChipY

'copy the chip pixels
	Dim lCameraChipPixels(1) As Long
	lCameraChipPixels(E_XDirection) = oVideoMappingInfo.CameraSize.ImageX
	lCameraChipPixels(E_YDirection) = oVideoMappingInfo.CameraSize.ImageY

'look for measpoints
	If oInfos.HasMeasPoints <> True Then
		MsgBox("Unexpected error, no scan points available.", vbCritical + vbOkOnly, GetMacroName())
		Err.Raise 1
	End If

'get the measurement points
	Dim oMeasPoints As MeasPoints
	Set oMeasPoints = oInfos.MeasPoints

'we expect that there are exactly 2 points defined
	If oMeasPoints.Count <> 2 Then
		MsgBox("Exactly 2 scan points are expected.", vbCritical + vbOkOnly, GetMacroName())
		Err.Raise 1
	End If

'get the first point
	Dim videoCoord(1) As Single
	oMeasPoints.Item(1).VideoXY(videoCoord(E_XDirection) ,videoCoord(E_YDirection))

	Dim dFirstPoint(1) As Double
	oVideoMappingInfo.CoordXYFromVideo(videoCoord(E_XDirection) ,videoCoord(E_YDirection), dFirstPoint(E_XDirection) ,dFirstPoint(E_YDirection))

'get the second point
	oMeasPoints.Item(2).VideoXY(videoCoord(E_XDirection) ,videoCoord(E_YDirection))

	Dim dSecondPoint(1) As Double
	oVideoMappingInfo.CoordXYFromVideo(videoCoord(E_XDirection) ,videoCoord(E_YDirection), dSecondPoint(E_XDirection) ,dSecondPoint(E_YDirection))

'calculate the pixel distances
	Dim lPixels(1) As Long
	Dim eDirection As Directions
	For eDirection = E_XDirection To E_YDirection ' do it for X and Y
		Dim dDiff As Double
		dDiff = Abs(dFirstPoint(eDirection) - dSecondPoint(eDirection) )

		Dim dFactor As Double
		dFactor = dCurrentMagnification(eDirection) * lCameraChipPixels(eDirection)/dCameraChipDimension(eDirection)

		Dim dPixels As Double
		dPixels = dDiff * dFactor

		lPixels(eDirection) = Round(dPixels)
	Next eDirection

'save the file
	oFile.Save()
	oFile.Close()
	bFileOpened = False

'delete the file
	Kill(strFileName)
	bFileAvailable = False


'Show the result in an extra file
	Dim strFileResult As String
	strFileResult = GetTempDir() + "ObjectiveCalibrationResult.txt"

   	Open strFileResult For Output As #1
   	Print #1, "The pixel distance of the scan points are:"
    Print #1, "Horizontal     " + CStr(lPixels(E_XDirection))
    Print #1, "Vertical       " + CStr(lPixels(E_YDirection))
    Print #1, "Please insert these values in the objective calibration dialog manually."
    Print #1, "The corresponding distances (in µm) have to be measured physically."
    Close #1

    Dim strStartNotepadWithFile As String
    strStartNotepadWithFile = "notepad " + strFileResult
    Shell "notepad " + strFileResult,1

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

	MsgBox("The magnification calibration was not successfull.", vbCritical + vbOkOnly, GetMacroName)
End Sub

Function GetTempDir As String
	Dim ret As Long, length As Long
	Dim strGetTempDir As String

    strGetTempDir = String$(255, 0)
    length = Len(strGetTempDir)
    ret = GetTempPath(length, strGetTempDir)

    If ret <> ERROR_SUCCESS Then
        GetTempDir = Left$(strGetTempDir, ret)
    Else
    	MsgBox("Could not locate the temporary directory.", vbCritical + vbOkOnly, GetMacroName())
		Err.Raise 1
    End If

End Function

Function GetFileNameTemp As String
	Dim ret As Long, length As Long

    'create file name
    GetFileNameTemp = String$(255, 0)
    ret = GetTempFileName(GetTempDir(), "psv", 0, GetFileNameTemp)

    If ret <> ERROR_SUCCESS Then
        GetFileNameTemp = Left$(GetFileNameTemp, InStr(1, GetFileNameTemp,".tmp") ) + "set"
    Else
    	MsgBox("Could not get the temporary file name.", vbCritical + vbOkOnly, GetMacroName())
		Err.Raise 1
    End If

End Function
