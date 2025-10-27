' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This Macro does some single shot acquisitions in FFT-Mode. 
' The maximum amplitude is searched in the spectrum .
' This maximum value and the corresponding frequency are transferred to an Excel-sheet.
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

' Temporary file, will be deleted at the end of the macro
Public P_AnalyzerFileName$

' Number of single shot acquisitions to be performed
Const C_Acquisitions% = 4

Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	P_AnalyzerFileName$ = Environ("Temp")+"\AnalyzerMax.txt"

    Dim x0 As Double	' Frequency value
    Dim y0 As Double	' Amplitude value

    ' Switch application to Acquisition mode, select desired signal to be displayed
	Call PreparePsv

    ' Open Excel and create new document
    Dim Excel As Object
    Set Excel = CreateObject("Excel.Application")
    Excel.Visible = True
    Excel.Workbooks.Add

	' Acquire data, find maximum, transfer maximum to Excel
	Dim i As Integer
	For i = 1 To C_Acquisitions%
		Call SingleShot
	    Call CalcMax(x0, y0)
		Excel.Range("A" + CStr$(i%)).Select
		Excel.ActiveCell.FormulaR1C1 = CStr(x0)
		Excel.Range("B" + CStr$(i%)).Select
		Excel.ActiveCell.FormulaR1C1 = CStr(y0)
	Next i%

	Call DeleteTemporaryFiles

	MsgBox("Macro has finished.", vbOkOnly)
End Sub

Sub SingleShot
' -------------------------------------------------------------------------------
' Start a single shot measurement.
' -------------------------------------------------------------------------------
	Acquisition.Start ptcAcqStartSingle
	While Acquisition.State <> ptcAcqStateStopped
		Wait 0.5
	Wend
End Sub

Sub CalcMax(ByRef x0 As Double, ByRef y0 As Double)
' -------------------------------------------------------------------------------
' Find maximum amplitude value.
' -------------------------------------------------------------------------------
	Dim WndAnalyzer As AnalyzerWindow
	Set WndAnalyzer = Acquisition.Document.Windows(1)
	' Save the fft data to a temporary file in order to access the data
	WndAnalyzer.AnalyzerView.Export(P_AnalyzerFileName$, ptcFileFormatText)

	' Read Data from temporary Ascii File and calculate maximum y value
	Dim x As Double
	Dim y As Double
	Dim i As Integer
	i = 0
	Open P_AnalyzerFileName$ For Input As #1
    x0 = 0
    y0 = -10000000	' a small value
    While Not EOF(1)
    	Dim sValue As String
        Line Input #1,sValue
        i = i + 1
        ' Ignore header of AnalyzerMax.txt (first 5 lines)
        If i > 5 Then
        	' Search for 'tab' (delimiter) in the line
        	Dim iTab As Integer
	        iTab = InStr(sValue, vbTab)
    	    If iTab <> 0 Then
	    	    x = CDec(Left$(sValue, iTab - 1))
	    	    y = CDec(Right$(sValue, Len(sValue) - iTab))
				If y > y0 Then
					' Actual line contains greater y value than any line before
					y0 = y
					x0 = x
				End If
			End If
		End If
    Wend
    Close #1
End Sub

' *******************************************************************************
' * Helper functions and subroutines
' *******************************************************************************

Sub PreparePsv
' -------------------------------------------------------------------------------
' Prepare the PSV program to the acquisition in FFT mode.
' Switch Analyzer Window to display FFT Signal.
' -------------------------------------------------------------------------------
	If Not SwitchToAcquisitionMode() Then
		MsgBox("Switch to acquisition mode failed.", vbOkOnly)
		End
	End If

	' Switch Measurement-Mode to FFT
	Acquisition.Mode = ptcAcqModeFft

	' Check if an analyzer window is open
	' Note: Acquisition.Document.Windows only accesses the analyzer window, not the live-video
	If Acquisition.Document.Windows.Count = 0 Then
		' Open new analyzer window
		Acquisition.Document.Windows.Add
	End If

	' Switch Domain Mode of first Acquisition window to FFT
	Dim WndAnalyzer As AnalyzerWindow
	Set WndAnalyzer = Acquisition.Document.Windows(1)
	WndAnalyzer.AnalyzerView.Settings.DisplaySettings.Domain = ptcDomainModeSpectrum
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
