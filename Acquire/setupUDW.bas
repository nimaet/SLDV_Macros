' Macro to set user defined waveform file
' ----------------------------------------------------------------------
'#Language "WWB.NET"

Option Explicit

Imports System
Imports System.Collections.Generic



Sub Main
' -------------------------------------------------------------------------------
' Main procedure.
' >Checks for Acquisition Mode
' >Calls Harmonic Scan Range Dialogue box
' >Sets the filenames by Calling Harmonic Scan Files sub and counts the # of files
' >Runs Harmonic Scan Measurement to populate data for each file
' >Shows "Acquisition Completed" as a message when successfully completed
' -------------------------------------------------------------------------------

	' ----------
    Dim GenProps As GeneratorAcqProperties
	GenProps = Acquisition.ActiveProperties.Item(PTCAcqPropertiesType.ptcAcqPropertiesTypeGenerators)(1)
	Acquisition.GeneratorsOn = True
	Dim Sine As New WaveformUserDefined
	Sine.Frequency = 10	' Hz
	Sine.Load("D:/Sai/1130/VarAmpSine42500_dur20ms_625000_T80000_RR10.txt")
	GenProps.Waveform = Sine
	Acquisition.GeneratorsOn = False

End Sub
