' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This Macro calculates the settings for peak hold measurements
'
' The macro takes the settings from the frequency page and calculates
' and changes the following settings:
'
'   - Page General: 	Measurement Mode: FFT
'						Averaging: Peak Hold
'						Number of Averages
'   - Page Frequency:	Overlap (if available)
'   - Page Window:		Flat Top
'   - Page Trigger:		Off
'   - Page Generator:	Active
'						Waveform: Sweep
'						Start Frequency
'						End Freqeuncy
'						Sweep Time
'
' After updating the settings successfully, generator and measurement
' are started automatically.

'#Uses "..\SwitchToAcquisitionMode.bas"

' The next line helps for (typo) error checking
Option Explicit

Dim dFreqRes As Double
Dim dStartFreq As Double
Dim dEndFreq As Double
Dim dSweepTime As Double
Dim dOverlap As Double
Dim dAverageCount As Long

Dim i As Long

Const iSweepTimeMultiplier As Integer = 2

Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	If Not SwitchToAcquisitionMode() Then
		MsgBox("Switch to acquisition mode failed.", vbOkOnly)
		End
	End If

	If Not Application.Acquisition.State = ptcAcqStateStopped Then
		MsgBox("Please stop the acquisition before running this macro!", vbCritical + vbOkOnly)
		Exit Sub
	End If

	' turn off generator
	Application.Acquisition.GeneratorsOn = False


	' Change measurement mode (FFT)
	Acquisition.Mode = ptcAcqModeFft

	' Get FFT settings
	Dim oFftAcqProps As FftAcqProperties
	Set oFftAcqProps = Acquisition.ActiveProperties.FftProperties
	dFreqRes = oFftAcqProps.Bandwidth / oFftAcqProps.Lines
	dStartFreq = oFftAcqProps.StartFrequency
	If dStartFreq < dFreqRes Then
		dStartFreq = dFreqRes
	End If
	dEndFreq = oFftAcqProps.EndFrequency

	' Change average mode
	Dim oAverageAcqProps As AverageAcqProperties
	Set oAverageAcqProps = Acquisition.ActiveProperties.AverageProperties
	oAverageAcqProps.type = ptcAveragePeakhold

	' Try to set Overlap to 75% (if available)
	On Error Resume Next
	oFftAcqProps.Overlap = 75
	On Error GoTo 0

	dOverlap = oFftAcqProps.Overlap / 100

	dSweepTime = iSweepTimeMultiplier * (1 - dOverlap) * (dEndFreq - dStartFreq + dFreqRes) / (dFreqRes * dFreqRes)

	' Change average count
	dAverageCount = (dSweepTime * dFreqRes) / (1 - dOverlap)
	On Error Resume Next
	oAverageAcqProps.Count = dAverageCount
	On Error GoTo 0
	If oAverageAcqProps.Count <> dAverageCount Then
		' Setting average count not succeeded
		MsgBox("Calculated average count: " + CStr(dAverageCount) + Chr$(13) + Chr$(13) + "Average count too high, please reduce the number of FFT lines!", vbCritical + vbOkOnly)
		Exit Sub
	End If

	' change window function of all channels to Flat Top
	Dim oChannels As ChannelsAcqProperties
	Set oChannels = Acquisition.ActiveProperties.ChannelsProperties
	For i = 1 To oChannels.Count
		Dim oChannelAcqProps As ChannelAcqProperties
		Set oChannelAcqProps = oChannels(i)
		oChannelAcqProps.WindowFunction = ptcWindowFctFlatTop
	Next

	' Change trigger settings
	' set trigger off
	Acquisition.ActiveProperties.TriggerProperties.Source = ptcTriggerSourceOff

	' Change generator settings
	Dim oGeneratorAcqProps As GeneratorAcqProperties
	Set oGeneratorAcqProps = Acquisition.ActiveProperties.GeneratorsProperties(1)
	oGeneratorAcqProps.Active = True
	Dim oSweep As New WaveformSweep
	oSweep.StartFrequency = dStartFreq - dFreqRes / 2
	oSweep.EndFrequency = dEndFreq + dFreqRes / 2
	oSweep.SweepTime = dSweepTime

	On Error GoTo InvalidSweep
	oGeneratorAcqProps.Waveform = oSweep
	Application.Acquisition.GeneratorsOn = True
	On Error GoTo 0

	Application.Acquisition.Start ptcAcqStartSingle

	If MsgBox("Calculated sweep time: " + CStr(Round(dSweepTime, 1)) + " s" + Chr$(13) + Chr$(13) + "Settings updated successfully. Generator and single shot measurement have been startet.", vbInformation + vbOkCancel) = vbCancel Then
		Application.Acquisition.Stop
		While Not Application.Acquisition.State = ptcAcqStateStopped
			Wait 0.1
		Wend
		Application.Acquisition.GeneratorsOn = False
	End If

	Exit Sub

InvalidSweep:
	MsgBox("Calculated sweep time: " + CStr(Round(dSweepTime, 1)) + " s" + Chr$(13) + Chr$(13) + "Sweep time too long, please reduce the number of FFT lines!", vbCritical + vbOkOnly)

End Sub
