' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This Macro shows in simple Basic language how to load settings and
' start an area scan
'
' Shows how to
' - load settings
' - start an area scan
' - wait for the area scan to be finished
'
' see also:
'			SettingsAndScan.bas: more advanced version, settings are
'                                changed in the macro before the scan is started
'

'#Uses "..\SwitchToAcquisitionMode.bas"

' the next line helps for (typo) error checking
Option Explicit

Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	If Not SwitchToAcquisitionMode() Then
		MsgBox("Switch to acquisition mode failed.", vbOkOnly)
		End
	End If
	
	' Load Settings
	' To do:
	'		- enter path and filename to your own .set or .svd file
	'		- specify which settings to load (
	'											ptcSettingsAll
	'											ptcSettingsAcquisition
	'											ptcSettingsAPS (scan points and grid)
	'											ptcSettingsCamera (camera zoom and focus)
	'											ptcSettingsAlignment (2D-alignment; 3D-alignment)
	'											ptcSettingsWindows (window layout)
	'										 )
	Settings.Load("C:\Documents and Settings\All Users\Application Data\Polytec\PSV\9.1\Examples\Data\Example.svd",	ptcSettingsAPS)

	' Specify filename for area scan
	' To do:
	'		- specify your own file name
	Acquisition.ScanFileName = "D:\test\test.svd"

	' Start Area Scan
	Acquisition.Scan(ptcScanAll)

	' Wait until area scan is finished
	While Acquisition.State <> ptcAcqStateStopped
		Wait 1		' wait 1 second
	Wend
	' Scan has finished

	' If you want to start another scan with different settings, you can copy and paste the commands above

End Sub
