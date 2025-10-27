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

Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	If Not SwitchToAcquisitionMode() Then
		MsgBox "Switch to acquisition mode failed."
		End
	End If
	'
	' Start Scan
	Acquisition.Scan ptcScanAll
		
	' The function ScanStateChanged is only called while the macro is running, so we wait until the scan has finished
	While Acquisition.State <> ptcAcqStateStopped
		Wait 1
	Wend

	MsgBox("Macro has finished.", vbOkOnly)
End Sub

Public Sub ScanStateChanged(ByVal ScanState As PTCScanState, ByVal ScanPoint As Long)
' -------------------------------------------------------------------------------
'	Set a pulse at the Digital IO-Port.
' -------------------------------------------------------------------------------
	If ScanState = ptcScanStateEndScanPoint Then
		DigitalPorts.Item(ptcDigitalPortOut1).Pulse(True)
	End If
End Sub
