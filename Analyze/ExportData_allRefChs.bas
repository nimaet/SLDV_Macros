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
	For Count = 0 To 32
		
		Dim oFile As PolyFile

		'strFileName = "D:\Chris\LR Calibration 01-25-21\f150_400_interp_" & CStr(Count) &".svd"
		'strFileName = "D:\Yiwei\12222020\ExcR_200Hz_0_alpha25_theta_" & CStr(Count) & "-501.svd"
		'strFileName = "D:\Yiwei\12032020\New\ExcR_390Hz_0_alpha5_theta_" & CStr(Count) & "-401.svd"
		'strFileName = "D:\Mustafa\Graded 01-27-2020\f200_z0_p_1_" & CStr(Count) & ".svd"
		'strFileName = "D:\Mohid\01-29-21 LREH\f780_z0-0025_2log" & CStr(Count) & ".svd"
		'strFileName = "D:\Obaidullah\03-26-21 Resistance Sweep new shaker\temp" & CStr(Count) & ".svd"
		strFileName = "D:\Obaidullah\05-13-2021-2nd\sweep_14_20_f_350_z_8_8_p_2_" & CStr(Count) & ".svd"

		If Not OpenFile(oFile, strFileName) Then
			Exit For
		End If


		Dim saveNameRe As String
		Dim saveNameIm As String

		saveNameRe = Replace(strFileName,".svd","vib_real.xlsx")
		saveNameIm = Replace(strFileName,".svd","vib_imag.xlsx")

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
		For RefCh = 2 To 8
			Dim refStr As String
			Dim refStrReal As String
			Dim refStrImag As String

			refStr = "Ref" & CStr(RefCh) & " & Ref1"
			refStrReal = "Ref" & CStr(RefCh) & "_real.xlsx"
			refStrImag = "Ref" & CStr(RefCh) & "_imag.xlsx"
			
			saveNameRe = Replace(strFileName,".svd",refStrReal)
			saveNameIm = Replace(strFileName,".svd",refStrImag)


			Set oDisplayRe = oDomain.Channels(refStr).Signals("H1 Voltage / Voltage").Displays("Real")
			Set oDisplayIm = oDomain.Channels(refStr).Signals("H1 Voltage / Voltage").Displays("Imaginary")
	
	
			' get frequency vector information
			Set oDomainXAxis = oPointDomain.GetXAxis(oDisplayRe)

			fMin = oDomainXAxis.Min
			fMax = oDomainXAxis.Max
			nFFT = oDomainXAxis.MaxCount
	
			'Debug.Print fMin
			'Debug.Print fMax
			'Debug.Print nFFT
	
			' Get number of points
			ptCount = oPointDomain.DataPoints.Count

			ReDim Data(1 To nFFT, 1 To ptCount+1)
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
