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

	'pathName = "D:\Chris\Wave Packet 12-14-21\"
	'pathName =  "D:\Renan\29\Time2\"

	pathName = "D:\Mustafa\new beam 02-25-2022\7500Hz sine\fm_p750_df_0-1\"

	Call OpenExcel

	strFileName = Dir(pathName + "*.svd")
	'strFileName = Dir(pathName + "ft_2100Hz_z_0-04_duration_50ms_chipr_750-3500Hz_BurstLength_1.svd")


	Dim tMin As Double
	Dim tMax As Double
	Dim nSamp As Double
	Dim Data() As Single

	While strFileName <> ""
		Dim oFile As PolyFile


		Debug.Print pathName+strFileName
		If Not OpenFile(oFile, pathName + strFileName) Then
			Exit While
		End If


		Dim saveName As String

		'saveName = Replace(strFileName,".svd","-Vib.xlsx")
		saveName = Replace(strFileName,".svd","-Ref1.xlsx")
		'saveName = Replace(strFileName,".svd","-Ref2.xlsx")

		' Select a PointDomain
		Dim oPointDomains As PointDomains
		Set oPointDomains = oFile.GetPointDomains()
		Dim oPointDomain As PointDomain

		Set oPointDomain = oPointDomains("Time")



		' use base class Domain to access channels
		Dim oDomain As Domain
		Set oDomain = oPointDomain

		' select a Display
		Dim oDisplay As Display
'
		'Set oDisplay = oDomain.Channels("Vib").Signals("Velocity").Displays("Samples")
		Set oDisplay = oDomain.Channels("Ref1").Signals("Voltage").Displays("Samples")
		'Set oDisplay = oDomain.Channels("Ref2").Signals("Voltage").Displays("Samples")


		' get frequency vector information
		Dim oDomainXAxis As XAxis
		Set oDomainXAxis = oPointDomain.GetXAxis(oDisplay)


		tMin = oDomainXAxis.Min
		tMax = oDomainXAxis.Max
		nSamp = oDomainXAxis.MaxCount

		Debug.Print tMin
		Debug.Print tMax
		Debug.Print nSamp

		' Get number of points
		Dim oDataPoint As DataPoint

		Dim ptCount As Integer
		ptCount = oPointDomain.DataPoints.Count


		ReDim Data(1 To nSamp, 1 To ptCount+1)

		Dim ptData() As Single
		ReDim ptData(nSamp)

		'Add time axis
		For Idx = 1 To nSamp
			Data(Idx, 1) = tMin + (Idx - 1)*(tMax - tMin)/(nSamp - 1)
		Next

		'Populate 2D array of data

		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		'@@@@@@@@@@@@@@@@@@@@@ REAL PART @@@@@@@@@@@@@@@@@@@@@@@@
		'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		For pCount = 1 To ptCount
			Set oDataPoint = oPointDomain.DataPoints(pCount)
			' now get the data

			ptData = oDataPoint.GetData(oDisplay, 0)

			For Idx = 1 To nSamp
				Data(Idx, pCount + 1) = ptData(Idx-1)
			Next
		Next

		' Assign 2D array to the spreadsheet
		Excel.Sheets(1).Cells.Clear
		Excel.Sheets(1).Range("A1").Resize(nSamp, ptCount+1) = Data

		' Save spreadsheet

		Excel.DisplayAlerts = False
		Excel.Workbooks(1).SaveAs FileName:=pathName + saveName
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
