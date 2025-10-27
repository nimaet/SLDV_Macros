Const c_strFileFilter As String = "Scan File (*.svd)|*.svd|All Files (*.*)|*.*||"
Const c_strFileExt As String = "svd"

Dim Excel As Object

Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------

	' get filename and path
	Dim strFileName As String

	Call OpenExcel

	Dim Count As Integer
	For Count = 0 To 0

		Dim oFile As PolyFile
		'strFileName = "D:\Chris\Goktug Leg 07-23-21\ChrisLeftLeg_motor.svd"
		'strFileName = "D:\Chris\ABH New Beam 09-15-21\f1250_z01p_0-018_nmax3-5.svd"
		'strFileName = "D:\Yiwei\12222020\ExcR_200Hz_0_alpha25_theta_" & CStr(Count) & "-501.svd"
		'strFileName = "D:\Yiwei\10162021\ExcR_5000Hz_0-09_alpha50_theta_" & CStr(Count) & "-400.svd"
		strFileName = "D:\Mustafa\new beam11-17-2021\f_2050hz_Z_zeta_0-035.svd"
		'strFileName = "D:\Mustafa\new beam 09-15-2021\bragg_bandgap_f_25_khz_v_1-0_new.svd"
	     'strFileName = "D:\Obaidullah\04-22-21\Scan_SC_InVolt0-8.svd"

		If Not OpenFile(oFile, strFileName) Then
			Exit For
		End If


		Dim saveNameRe As String
		Dim saveNameIm As String

		saveNameRe = Replace(strFileName,".svd","real.xlsx")
		saveNameIm = Replace(strFileName,".svd","imag.xlsx")

		' Select a PointDomain
		Dim oPointDomains As PointDomains
		Set oPointDomains = oFile.GetPointDomains()
		Dim oPointDomain As PointDomain
		Set oPointDomain = oPointDomains("FFT")
		' use base class Domain to access channels
		Dim oDomain As Domain
		Set oDomain = oPointDomain

		' select a Display
		Dim oDisplayRe As Display
		Dim oDisplayIm As Display

		Set oDisplayRe = oDomain.Channels("Vib & Ref1").Signals("H1 Velocity / Voltage").Displays("Real")
		Set oDisplayIm = oDomain.Channels("Vib & Ref1").Signals("H1 Velocity / Voltage").Displays("Imaginary")

		'Set oDisplayRe = oDomain.Channels("Vib & Ref1").Signals("H1 Acceleration / Acceleration").Displays("Real")
		'Set oDisplayIm = oDomain.Channels("Vib & Ref1").Signals("H1 Acceleration / Acceleration").Displays("Imaginary")


		' get frequency vector information
		Dim oDomainXAxis As XAxis
		Set oDomainXAxis = oPointDomain.GetXAxis(oDisplayRe)

		Dim fMin As Double
		Dim fMax As Double
		Dim nFFT As Integer

		fMin = oDomainXAxis.Min
		fMax = oDomainXAxis.Max
		nFFT = oDomainXAxis.MaxCount

		'Debug.Print fMin
		'Debug.Print fMax
		'Debug.Print nFFT

		' Get number of points
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

		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		'@@@@@@@@@@@@@@@@@@@@@ REAL PART @@@@@@@@@@@@@@@@@@@@@@@@
		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		For pCount = 1 To ptCount
			Set oDataPoint = oPointDomain.DataPoints(pCount)
			' now get the data

			ptData = oDataPoint.GetData(oDisplayRe, 0)

			For Idx = 1 To nFFT
				Data(Idx, pCount + 1) = ptData(Idx-1)
			Next
		Next

		' Assign 2D array to the spreadsheet
		Excel.Sheets(1).Range("A1").Resize(nFFT, ptCount+1) = Data

		' Save spreadsheet

		Excel.DisplayAlerts = False
		Excel.Workbooks(1).SaveAs FileName:=saveNameRe
		Excel.DisplayAlerts = True

		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		'@@@@@@@@@@@@@@@@@@@@@ IMAG PART @@@@@@@@@@@@@@@@@@@@@@@@
		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		For pCount = 1 To ptCount
			Set oDataPoint = oPointDomain.DataPoints(pCount)
			' now get the data

			ptData = oDataPoint.GetData(oDisplayIm, 0)

			For Idx = 1 To nFFT
				Data(Idx, pCount + 1) = ptData(Idx-1)
			Next
		Next

		' Assign 2D array to the spreadsheet
		Excel.Sheets(1).Range("A1").Resize(nFFT, ptCount+1) = Data

		' Save spreadsheet

		Excel.DisplayAlerts = False
		Excel.Workbooks(1).SaveAs FileName:=saveNameIm
		Excel.DisplayAlerts = True

		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

	Next

	Excel.Quit

End Sub



' *******************************************************************************
' * Helper functions and subroutines
' *******************************************************************************

Public Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Sub OpenExcel
' -------------------------------------------------------------------------------
' Open Excel application.
' -------------------------------------------------------------------------------
   	Set Excel = CreateObject("Excel.Application")
    Excel.Visible = True
    Excel.Workbooks.Add
    'Excel.ScreenUpdating = False
End Sub

'Sub ReadDataAndInsertInExcel(ByRef sVal() As Single, iRow As Integer, iCol As Integer)
'' -------------------------------------------------------------------------------
'' Read data from Ascii file and insert them into Excel.
'' -------------------------------------------------------------------------------
	'Excel.Sheets(1).Range(Cells(1, iCol), Cells(iRow, iCol)) = sVal
'End Sub

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
