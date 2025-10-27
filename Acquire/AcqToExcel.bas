' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This Macro does some single shot acquisitions in FFT-Mode. The
' complete measurement data (all domain modes) is transferred to an Excel-sheet.
'
' E.g. for the Measurement-Mode FFT:
' Column A and B contain time data,
' Column C and D contain FFT data (Magnitude),
' Column E and F contain 1/3 Octave data.
' The data of each single shot acquisition is saved to a new table.
'
' Note: This macro only works properly if the decimal symbol is set to '.'
' and the digit grouping symbol is set to ',' in the Regional Settings
' or Regional Options (for both: numbers and currency)
'
' Shows how to
' - use the Excel application with a macro
' - get access to measurement data
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
' ----------------------------------------------------------------------

'#Uses "..\SwitchToAcquisitionMode.bas"

Option Explicit

' Temporary file, will be deleted at the End of the macro
Public P_AnalyzerFileName$

' Number of single shot acquisitions to be performed
Const C_Acquisitions% = 4

Dim Excel As Object

Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	P_AnalyzerFileName = Environ("Temp")+"\Acquisition.txt"

	If Not SwitchToAcquisitionMode() Then
		MsgBox("Switch to acquisition mode failed.", vbOkOnly)
		End
	End If

    ' Open Excel application
    Call OpenExcel

    ' Set acquisition mode to fft
	Acquisition.Mode = ptcAcqModeFft
	' Do the acquisitions
	Dim i As Integer
	For i = 1 To C_Acquisitions%
		Call SingleShot
		Call TransferDomainModesToExcel(i)
	Next i

	Call DeleteTemporaryFiles

	MsgBox("Macro has finished.", vbOkOnly)
End Sub

Sub OpenExcel
' -------------------------------------------------------------------------------
' Open Excel application.
' -------------------------------------------------------------------------------
   	Set Excel = CreateObject("Excel.Application")
    Excel.Visible = True
    Excel.Workbooks.Add
End Sub

Sub SingleShot
' -------------------------------------------------------------------------------
' Make one singleshot.
' -------------------------------------------------------------------------------
	Acquisition.Start ptcAcqStartSingle
	While Acquisition.State <> ptcAcqStateStopped
		Wait 0.5
	Wend
End Sub

Sub TransferDomainModesToExcel(iAcquisitionCount As Integer)
' -------------------------------------------------------------------------------
' Write data into file.
' -------------------------------------------------------------------------------
	Dim WndAnalyzer As AnalyzerWindow
	Set WndAnalyzer = Acquisition.Document.Windows(1)
	Dim idxDomain As Integer
	For idxDomain = 1 To WndAnalyzer.AnalyzerView.Domains.Count
		Dim oDomain As Domain
		Set oDomain = WndAnalyzer.AnalyzerView.Domains.Item(idxDomain)

		' Set active domain
		WndAnalyzer.AnalyzerView.Settings.DisplaySettings.Domain = oDomain.type
		
		' Export data to a temporary file in order to access the data
		WndAnalyzer.AnalyzerView.Export(P_AnalyzerFileName$, ptcFileFormatText)

		' Read data from Ascii file and insert them into Excel.
		Call ReadAsciiFileAndInsertInExcel(P_AnalyzerFileName$, 2*(idxDomain-1), iAcquisitionCount)

	Next idxDomain
End Sub

Sub ReadAsciiFileAndInsertInExcel(sFileName As String, iRowOffset As Integer, iSheetNumber As Integer)
' -------------------------------------------------------------------------------
' Read data from Ascii file and insert them into Excel.
' -------------------------------------------------------------------------------
    If Excel.Sheets.Count < iSheetNumber Then
	    Excel.Sheets.Add ,Excel.Sheets(Excel.Sheets.Count)
	End If
	Excel.Sheets(iSheetNumber).Activate

	Dim sRowX As String
	Dim sRowY As String
	sRowX = Chr$(Asc("A") + iRowOffset)
	sRowY = Chr$(Asc("B") + iRowOffset)

	Dim iLine As Integer
	iLine = 0
	Open sFileName For Input As #1
    While Not EOF(1)
		Dim sValue As String
        Line Input #1,sValue
        iLine = iLine + 1
        If iLine > 5 Then
        	Dim iTab As Integer
	        iTab = InStr(sValue, vbTab)
    	    If iTab <> 0 Then
    	    	Dim x As Double
    	    	Dim y As Double
	    	    x = CDec(Left$(sValue, iTab - 1))
	    	    y = CDec(Right$(sValue, Len(sValue) - iTab))
				Excel.Range(sRowX + CStr$(iLine - 5)).Select
				Excel.ActiveCell.FormulaR1C1 = CStr(x)
				Excel.Range(sRowY + CStr$(iLine - 5)).Select
				Excel.ActiveCell.FormulaR1C1 = CStr(y)
			End If
		End If
    Wend
    Close #1
End Sub

Sub DeleteTemporaryFiles
' -------------------------------------------------------------------------------
' Delete files.
' -------------------------------------------------------------------------------
	Dim sFile As String
    sFile = Dir$(P_AnalyzerFileName$)
	If sFile <> "" Then
		Kill P_AnalyzerFileName$
	End If
End Sub
