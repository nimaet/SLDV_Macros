' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This Macro shows the use of the window settings.
'
' The macro tries to copy the window settings of the active window to
' all open windows.
'
' How to use this macro:
' - start PSV/VibSoft and open the files which should have the same window settings
' - activate one window and change the window settings like they should be for all Windows
' - start macro
'
' References
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
Option Explicit

Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	Dim ActiveWnd As Window
	Set ActiveWnd = Application.ActiveWindow

	If (ActiveWnd Is Nothing) Then
		Exit Sub
	End If

	Dim Wnds As Windows
	Set Wnds = Application.Windows
	Dim Wnd As Window
	Dim WndArea As AreaWindow
	Dim ActiveWndAnalyzer As AnalyzerWindow
	For Each Wnd In Wnds
		If (Not Wnd.Active()) Then
			'Wnd.Settings = ActiveWnd.Settings  'copy size and position of active window

			If (Wnd.type = ptcWindowTypeAnalyzer And ActiveWnd.type = ptcWindowTypeAnalyzer) Then
				Set ActiveWndAnalyzer = ActiveWnd
				Dim WndAnalyzer As AnalyzerWindow
				Set WndAnalyzer = Wnd
				WndAnalyzer.AnalyzerView.Settings = ActiveWndAnalyzer.AnalyzerView.Settings

			ElseIf (Wnd.type = ptcWindowTypeArea And ActiveWnd.type = ptcWindowTypeArea) Then
				Dim ActiveWndArea As AreaWindow
				Set ActiveWndArea = ActiveWnd
				Set WndArea = Wnd
				WndArea.AreaView.Settings = ActiveWndArea.AreaView.Settings
				If ((Not WndArea.AnalyzerView Is Nothing) And (Not ActiveWndArea.AnalyzerView Is Nothing)) Then
					WndArea.AnalyzerView.Settings = ActiveWndArea.AnalyzerView.Settings
				End If

			ElseIf (Wnd.type = ptcWindowTypeArea And ActiveWnd.type = ptcWindowTypeAnalyzer) Then
				Set WndArea = Wnd
				If (Not WndArea.AnalyzerView Is Nothing) Then
					Set ActiveWndAnalyzer = ActiveWnd
					WndArea.AnalyzerView.Settings = ActiveWndAnalyzer.AnalyzerView.Settings
				End If

			ElseIf (Wnd.type = ptcWindowTypeSigPro And ActiveWnd.type = ptcWindowTypeSigPro) Then
				Dim ActiveWndSigPro As SigProWindow
				Set ActiveWndSigPro = ActiveWnd
				Dim WndSigPro As SigProWindow
				Set WndSigPro = Wnd
				WndSigPro.AnalyzerView.Settings = ActiveWndSigPro.AnalyzerView.Settings
			End If
		End If
	Next Wnd

	MsgBox("Macro has finished.", vbOkOnly)
End Sub
