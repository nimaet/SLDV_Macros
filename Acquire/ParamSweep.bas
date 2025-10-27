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

	For Count = 0 To 400
		Acquisition.ScanFileName = "D:\Yiwei\10162021\ExcR_5000Hz_0-09_alpha50_theta_" & CStr(Count) & "-400.svd"
		'Acquisition.ScanFileName = "D:\Chris\ABH 02-08-21\f320_350_interp_" & CStr(Count) &".svd"
		'Acquisition.ScanFileName = "D:\Mustafa\Graded 02-12-2021\f150_z0_p_05_" & CStr(Count) & ".svd"
		'Acquisition.ScanFileName = "D:\Mohid\01-29-21 LREH\f780_z0-0025_2log" & CStr(Count) & ".svd"
		'Acquisition.ScanFileName = "D:\Obaidullah\03-26-21 Resistance Sweep new shaker\temp" & CStr(Count) & ".svd"
		'Acquisition.ScanFileName = "D:\Obaidullah\05-13-2021-2nd\sweep_14_20_f_350_z_8_8_p_2_" & CStr(Count) & ".svd"
		Acquisition.Scan ptcScanAll


		While Acquisition.State <> PTCAcqState.ptcAcqStateStopped
            Wait 0.1
        Wend
		Wait 10
	Next

	MsgBox("Macro has finished.", vbOkOnly)
End Sub


'
'This runs continuously in the background and executes when scan state changes
'
Public Sub ScanStateChanged(ByVal ScanState As PTCScanState, ByVal ScanPoint As Long)
' -------------------------------------------------------------------------------
'	Set a pulse at the Digital IO-Port if scan is done
' -------------------------------------------------------------------------------
	If ScanState = ptcScanStateEndScan Then
		DigitalPorts.Item(ptcDigitalPortOut1).Value = False
		Wait 0.6
		DigitalPorts.Item(ptcDigitalPortOut1).Value = True
		Wait 0.6
		DigitalPorts.Item(ptcDigitalPortOut1).Value = False
	End If
End Sub
