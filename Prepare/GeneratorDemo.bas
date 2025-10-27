' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This Macro shows how to access and change the generator properties.
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

Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	' Switch to Acquisition mode if application is PSV
	If Not SwitchToAcquisitionMode() Then
		MsgBox "Please switch to acquisition mode to start this macro."
		End
	End If

	If Acquisition.ActiveProperties.Item(ptcAcqPropertiesTypeGenerators).Count = 0 Then
		MsgBox "No Generator available." + vbCrLf + "Macro will be terminated."
		End
	End If

	Dim GeneratorProps As GeneratorAcqProperties
	Set GeneratorProps = Acquisition.ActiveProperties.Item(ptcAcqPropertiesTypeGenerators)(1)

	If (GeneratorProps.Active = False) Then
		MsgBox "Generator not active. Generator will be activated."
		GeneratorProps.Active = True
		If (GeneratorProps.Active = False) Then
			MsgBox "Not possible to activate generator."+ vbCrLf + "Macro will be terminated."
			End
		End If
	End If

	' Those settings will be used by all waveforms
	GeneratorProps.Offset = 0
	GeneratorProps.Amplitude = 3
	GeneratorProps.SteadyStateTime = 0.2
	' Turn Generator on
	Acquisition.GeneratorsOn = True

	' Sine
	MsgBox "Generator will be switched to Sine"
	Dim Sine As New WaveformSine
	Sine.Frequency = 100	' Hz
	GeneratorProps.Waveform = Sine
	SendKeys "{F5}", True		' Open A/D settings dialog, click on the generator page to view the settings

	' Periodic Chirp
	MsgBox "Generator will be switched to Periodic Chirp"
	Dim PeriodicChirp As New WaveformPeriodicChirp
	GeneratorProps.Waveform = PeriodicChirp
	SendKeys "{F5}", True		' Open A/D settings dialog, click on the generator page to view the settings

	' Turn Generator off
	Acquisition.GeneratorsOn = False

	MsgBox("Macro has finished.", vbOkOnly)
End Sub
