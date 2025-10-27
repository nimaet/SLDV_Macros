Const c_strFileFilter As String = "Scan File (*.svd)|*.svd|All Files (*.*)|*.*||"
Const c_strFileExt As String = "svd"

Dim Excel As Object

Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------

	' get filename and path
	Dim strFileName As String
	Dim pathName As String
	'pathName = "D:\Chris\Negative Capacitance 11-16-21\FRF\"
	pathName="D:\Renan\13\"
	Call OpenExcel

	strFileName = Dir(pathName + "*.pvd")
	While strFileName <> ""
		Dim oFile As PolyFile


		Debug.Print pathName+strFileName
		If Not OpenFile(oFile, pathName + strFileName) Then
			Exit While
		End If


		Dim saveNameRe As String
		Dim saveNameIm As String
		Dim saveNameCoh As String

		saveNameRe = Replace(strFileName,".pvd","real.xlsx")
		saveNameIm = Replace(strFileName,".pvd","imag.xlsx")
		saveNameCoh = Replace(strFileName,".pvd","coh.xlsx")

		' Select a PointDomain
		Dim oPointDomains As PointDomains
		Set oPointDomains = oFile.GetPointDomains()
		Dim oPointDomain As PointDomain

		Set oPointDomain = oPointDomains("FFT")
		'Set oPointDomain = oPointDomains("FFT")



		' use base class Domain to access channels
		Dim oDomain As Domain
		Set oDomain = oPointDomain

		' select a Display
		Dim oDisplayRe As Display
		Dim oDisplayIm As Display
		Dim oDisplayCoh As Display

		Set oDisplayRe = oDomain.Channels("Vib & Ref1").Signals("H1 Velocity / Voltage").Displays("Real")
		Set oDisplayIm = oDomain.Channels("Vib & Ref1").Signals("H1 Velocity / Voltage").Displays("Imaginary")
		Set oDisplayCoh = oDomain.Channels("Vib & Ref1").Signals("Coherence").Displays("Magnitude")

		'Set oDisplayRe = oDomain.Channels("Vib & Ref1").Signals("H1 Acceleration / Acceleration").Displays("Real")
		'Set oDisplayIm = oDomain.Channels("Vib & Ref1").Signals("H1 Acceleration / Acceleration").Displays("Imaginary")

		'Set oDisplayRe = oDomain.Channels("Vib & Ref1").Signals("H2 Velocity / Voltage").Displays("Real")
		'Set oDisplayIm = oDomain.Channels("Vib & Ref1").Signals("H2 Velocity / Voltage").Displays("Imaginary")

		' get frequency vector information
		Dim oDomainXAxis As XAxis
		Set oDomainXAxis = oPointDomain.GetXAxis(oDisplayRe)

		Dim fMin As Double
		Dim fMax As Double
		Dim nFFT As Integer

		fMin = oDomainXAxis.Min
		fMax = oDomainXAxis.Max
		nFFT = oDomainXAxis.MaxCount

		Debug.Print fMin
		Debug.Print fMax
		Debug.Print nFFT

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
				'Debug.Print Idx
				Data(Idx, pCount+1) = ptData(Idx-1)
			Next
		Next

		' Assign 2D array to the spreadsheet
		Excel.Sheets(1).Range("A1").Resize(nFFT, ptCount+1) = Data

		' Save spreadsheet

		Excel.DisplayAlerts = False
		Excel.Workbooks(1).SaveAs FileName:=pathName + saveNameRe
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
		Excel.Workbooks(1).SaveAs FileName:=pathName + saveNameIm
		Excel.DisplayAlerts = True

		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		'@@@@@@@@@@@@@@@@@@@@@ Coherence @@@@@@@@@@@@@@@@@@@@@@@@
		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		For pCount = 1 To ptCount
			Set oDataPoint = oPointDomain.DataPoints(pCount)
			' now get the data

			ptData = oDataPoint.GetData(oDisplayCoh, 0)

			For Idx = 1 To nFFT
				Data(Idx, pCount + 1) = ptData(Idx-1)
			Next
		Next

		' Assign 2D array to the spreadsheet
		Excel.Sheets(1).Cells.Clear
		Excel.Sheets(1).Range("A1").Resize(nFFT, ptCount+1) = Data

		' Save spreadsheet

		Excel.DisplayAlerts = False
		Excel.Workbooks(1).SaveAs FileName:=pathName + saveNameCoh
		Excel.DisplayAlerts = True

		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


		strFileName = Dir()
	Wend

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
