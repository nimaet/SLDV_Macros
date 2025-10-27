' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This Macro displays a user dialog which shows single acquisitions as
' bitmap
'
' Shows how to
' - do multiple single shot acquisitions
' - save graphics and use the bitmap
' - create a user dialog
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
Public P_AnalyzerPicture$

Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	P_AnalyzerPicture$ = Environ("Temp") + "\SimpleAnalyzer.bmp"

	If Not SwitchToAcquisitionMode() Then
		MsgBox("Switch to acquisition mode failed.", vbOkOnly)
		End
	End If

	Begin Dialog UserDialog 1040,217,"Analyzer",.ShowBitmapDlg
		Picture 0,0,1040,175,"Analyzer",0,.AnalyzerPicture
		OKButton 420,189,90,21
	End Dialog

	' Hide the application
	On Error GoTo MShowPSV
	Visible = False

	' Show the dialog
	Dim dlg As UserDialog
	Dialog dlg

MShowPSV:
	' Show the application
	Visible = True

	' Delete temporary bitmap file
	Call DeleteTemporaryFiles

	MsgBox("Macro has finished.", vbOkOnly)
End Sub

Private Function ShowBitmapDlg(DlgItem$, Action%, SuppValue&) As Boolean
' -------------------------------------------------------------------------------
' Show dialog.
' See DialogFunc help topic for more information.
' -------------------------------------------------------------------------------
	If Acquisition.Document.Windows.Count() < 1 Then
        Acquisition.Document.Windows.Add()
    End If

	Dim WndAnalyzer As AnalyzerWindow
	Set WndAnalyzer = Acquisition.Document.Windows(1)

	Select Case Action%
	Case 1 ' Dialog box initialization
		' Be sure to have an analyzer window open
		With Acquisition.Document
			If .Windows.Count = 0 Then
				.Windows.Add
			End If
			.Windows(1).Height = 250
			.Windows(1).Width = 800
		End With

		' Save Analyzer window as bitmap
		WndAnalyzer.AnalyzerView.Export(P_AnalyzerPicture$, ptcFileFormatGraphic)
		' Set bitmap to dialog
		DlgSetPicture "AnalyzerPicture",P_AnalyzerPicture$,0

	Case 2 ' Value changing or button pressed
		If Acquisition.State <> ptcAcqStateStopped Then
			Acquisition.Stop 
		End If
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		If Acquisition.State = ptcAcqStateStopped Then

			' Save Analyzer window as bitmap
			WndAnalyzer.AnalyzerView.Export(P_AnalyzerPicture$, ptcFileFormatGraphic)
			' Set bitmap to dialog
			DlgSetPicture "AnalyzerPicture",P_AnalyzerPicture$,0

			' Start next single shot acquisition
			Acquisition.Start ptcAcqStartSingle
		End If
		Wait 0.5	' Give the application time to do the acquisition
		ShowBitmapDlg = True
	Case 6 ' Function key
	End Select
End Function

Sub DeleteTemporaryFiles
' -------------------------------------------------------------------------------
' Delete files.
' -------------------------------------------------------------------------------
	Dim sFile As String
    sFile = Dir$(P_AnalyzerPicture$)
	If sFile <> "" Then
		Kill P_AnalyzerPicture$
	End If
End Sub
