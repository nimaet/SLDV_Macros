' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This Macro shows in simple Basic language how to change settings and
' start an area scan
'
' Shows how to
' - change settings
' - start an area scan
' - wait for the area scan to be finished
'
' Note: no error checking is done, macro might fail depending on system configuration and prior settings'
'
' In order to modify the macro, please type an object name, followed by a dot. A dropdown list will be displayed,
' showing all the properties of the object.
' E.g. in order to modify the overlap setting:
' 		1. type "oFftAcqProps."
'		2. select Overlap from the list and type "="
'		3. type the desired value for overlap (in percent)
'
' Please read the "Basic Engine Manual" in Start -> All Programs -> PSV x.y

'#Uses "..\SwitchToAcquisitionMode.bas"

' The next line helps for (typo) error checking
Option Explicit


Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	If Not SwitchToAcquisitionMode() Then
		MsgBox("Switch to acquisition mode failed.", vbOkOnly)
		End
	End If
	
	' Change general settings
	' Example: turn AutoRemeasure on
	Acquisition.ActiveProperties.GeneralProperties.AutoRemeasure = True

	' Change measurement mode (FFT, FastScan, Time, ...)
	Acquisition.Mode = ptcAcqModeFft

	' Change averaging settings
	Acquisition.ActiveProperties.AverageProperties.type = ptcAverageComplex
	Acquisition.ActiveProperties.AverageProperties.Count = 10

	' Change channel settings
	' Example: channel Reference 1 (2nd channel in list)
	Dim oChannelAcqProps As ChannelAcqProperties
	Set oChannelAcqProps = Acquisition.ActiveProperties.ChannelsProperties(2)
	' Example: set calibration factor to 0.1
	oChannelAcqProps.Calibration = 0.1
	' Example: change window function to Hanning
	oChannelAcqProps.WindowFunction = ptcWindowFctHanning
	' Example: disable SE (for reference 1)
	oChannelAcqProps.SEActive = False

	' Change FFT settings
	' Example: set bandwidth 20kHz, 400 FFT lines
	Acquisition.ActiveProperties.FftProperties.Bandwidth = 20000	' Hz
	Acquisition.ActiveProperties.FftProperties.Lines = 400

	' Window function: see above

	' Change trigger settings
	' Example: set trigger source to external
	Acquisition.ActiveProperties.TriggerProperties.Source = ptcTriggerSourceExternal

	' Change SE settings (SE settings for channels: see above)
	' Example: set SE mode to standard, turn on SpeckleTracking
	If Acquisition.ActiveProperties.HasSignalEnhancementProperties Then
		Acquisition.ActiveProperties.SignalEnhancementProperties.Mode = ptcSignalEnhancementModeFastStandard
		Acquisition.ActiveProperties.SignalEnhancementProperties.SpeckleTracking = True
	End If

	' Change vibrometer settings
	' Example: first (standard) vibrometer (if PSV-3D the settings are applied to all vibrometers)
	' Example: set vibrometer range
	Acquisition.ActiveProperties.VibrometersProperties(1).VibControllerSettings.QuantitySettingsCollection.ByKey(QuantityType_Velocity).Range = "VD-03 1000 mm/s/V"

	' Change generator settings
	' Example: set "Wait for Steady State" to 10 seconds
	Acquisition.ActiveProperties.GeneratorsProperties(1).SteadyStateTime = 10	' seconds
	' Example: switch generator to sine, 1.5 kHz
	Dim oSine As New WaveformSine
	oSine.Frequency = 1500	' Hz
	Acquisition.ActiveProperties.GeneratorsProperties(1).Waveform = oSine

	If Acquisition.ActiveProperties.TriggerProperties.Source = ptcTriggerSourceExternal Then
		Acquisition.GeneratorsOn = True
	End If

	' Perform area scan

	' Specify filename for area scan
	' To do:
	'		- specify your own file name
	Acquisition.ScanFileName = "D:\temp\test.svd"

	' Start area scan
	Acquisition.Scan(ptcScanAll)

	' Wait until area scan is finished
	While Acquisition.State <> ptcAcqStateStopped
		Wait 1		' wait 1 second
	Wend
	' Scan has finished

	' If you want to start another scan with different settings, you can copy and paste the commands above

End Sub
