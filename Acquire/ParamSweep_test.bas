' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This Macro starts a scan and set a pulse at the Digital IO-Port every
' time a scan point is finished.
' With the procedure ScanStateChanged the macro will be informed
' by the PSV application about a state change of the scan progress.
'
' Shows how to
' - use the callback procedure ScanStateChanged
' - use the digital ports
'
' References
' - Polytec PSV Type Library
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

Const c_strFileFilter As String = "Scan File (*.svd)|*.svd|All Files (*.*)|*.*||"
Const c_strFileExt As String = "svd"

Dim Excel As Object


Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	If Not SwitchToAcquisitionMode() Then
		MsgBox "Switch to acquisition mode failed."
		End
	End If
	Call OpenExcel
	For Count = 11 To 20
		'Acquisition.ScanFileName = "D:\Yiwei\09042020\ExcR_380Hz_0_alpha5_theta_" & CStr(Count) & "-401.svd"
		Acquisition.ScanFileName = "D:\Chris\temp\test" & CStr(Count) &".svd"
		Acquisition.Scan ptcScanAll

		While Acquisition.State <> PTCAcqState.ptcAcqStateStopped
            Wait 0.1
        Wend
        SaveData(Acquisition.ScanFileName)
		Wait 10
	Next

	MsgBox("Macro has finished.", vbOkOnly)
End Sub


'
'This runs continuously in the background and executes when scan state changes
'
Public Sub ScanStateChanged(ByVal ScanState As PTCScanState, ByVal ScanPoint As Long)
' -------------------------------------------------------------------------------
'	Set a pulse at the Digital IO-Port if scan is done
' -------------------------------------------------------------------------------
	If ScanState = ptcScanStateEndScan Then
		DigitalPorts.Item(ptcDigitalPortOut1).Value = False
		Wait 0.6
		DigitalPorts.Item(ptcDigitalPortOut1).Value = True
		Wait 0.6
		DigitalPorts.Item(ptcDigitalPortOut1).Value = False
	End If
End Sub

Public Sub SaveData(ByVal strFileName As String)
	Dim oFile As PolyFile

	If Not OpenFile(oFile, strFileName) Then
		Exit Sub
	End If


	' Select a PointDomain
	Dim oPointDomains As PointDomains
	Set oPointDomains = oFile.GetPointDomains()
	Dim oPointDomain As PointDomain
	Set oPointDomain = oPointDomains("FFT")
	' use base class Domain to access channels
	Dim oDomain As Domain
	Set oDomain = oPointDomain

	' select a Display
	Dim oDisplay As Display
	Set oDisplay = _
	oDomain.Channels("Vib & Ref1").Signals("H1 Velocity / Voltage").Displays("Magnitude")


	Dim oDomainXAxis As XAxis
	Set oDomainXAxis = oPointDomain.GetXAxis(oDisplay)

	Dim fMin As Double
	Dim fMax As Double
	Dim nFFT As Integer

	fMin = oDomainXAxis.Min
	fMax = oDomainXAxis.Max
	nFFT = oDomainXAxis.MaxCount

	Debug.Print fMin
	Debug.Print fMax
	Debug.Print nFFT

	' select a measurement point
	Dim oDataPoint As DataPoint


	Dim ptCount As Integer
	ptCount = oPointDomain.DataPoints.Count


	Dim Data() As Single
	ReDim Data(1 To nFFT, 1 To ptCount+1)

	Dim ptData() As Single
	ReDim ptData(nFFT)

	'Add frequency axis
	For Idx = 1 To nFFT
		Data(Idx, 1) = fMin + (Idx - 1)*(fMax - fMin)/(nFFT - 1)
	Next

	'Populate 2D array of data
	For Count = 1 To ptCount
		Set oDataPoint = oPointDomain.DataPoints(Count)
		' now get the data

		ptData = oDataPoint.GetData(oDisplay, 0)

		'Call ReadDataAndInsertInExcel(Data, nFFT, Count+1)
		For Idx = 1 To nFFT
			Data(Idx, Count + 1) = ptData(Idx-1)
		Next
	Next

	' Assign 2D array to the spreadsheet
	Excel.Sheets(1).Range("A1").Resize(nFFT, ptCount) = Data

	' Save spreadsheet
	Dim saveName As String
	saveName = Replace(strFileName,".svd",".xlsx")
	Excel.DisplayAlerts = False
	Excel.Workbooks(1).SaveAs fileName:=saveName
	Excel.DisplayAlerts = True
	Excel.Workbooks(1).Close
End Sub

Sub OpenExcel
' -------------------------------------------------------------------------------
' Open Excel application.
' -------------------------------------------------------------------------------
   	Set Excel = CreateObject("Excel.Application")
    Excel.Visible = True
    Excel.Workbooks.Add
    'Excel.ScreenUpdating = False
End Sub

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
