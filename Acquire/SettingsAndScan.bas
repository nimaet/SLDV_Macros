' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This Macro loads different settings. For every setting a scan is done.
'
' Shows how to
' - load settings
' - do a complete scan
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

Option Explicit


Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	' Array of setting names.
	' Those settings must have been created before.
	Const C_Settings = Array("Settings1.set", "Settings2.set")
	' Path with settings files
	Const C_Path = "D:\Temp\"

	' Directory for saving the scans
	Dim directory$
	directory = Environ("Temp")

	' Array of scan file names
	Dim scanFileNames
	scanFileNames = Array(directory+"\Scan01.svd", _
	                          directory+"\Scan02.svd")

	If Not SwitchToAcquisitionMode() Then
		MsgBox "Switch to acquisition mode failed."
		End
	End If

	Dim i As Integer
	Dim sFile As String

	For i% = 1 To 2
		' Load AD-settings and point definitions and camera settings (note: use 'Or' to combine flags)
		On Error GoTo SettingsNotAvailable

		sFile = C_Path + C_Settings(i%-1)
		Settings.Load sFile, ptcSettingsAll

		On Error GoTo 0		' Stop macro on Error

		' Start scan
		Acquisition.ScanFileName = scanFileNames(i%-1)
		Acquisition.Scan ptcScanAll
		
		' Wait until scan finished
		While Acquisition.State <> ptcAcqStateStopped
			Wait 1
		Wend

		GoTo NoError

SettingsNotAvailable:
		If Err.Number <> 0 Then
			MsgBox "Settings " + C_Settings(i%-1) + " not found."
		End If

NoError:
	Next i%

	MsgBox("Macro has finished.", vbOkOnly)
End Sub
