' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This macro allows you to monitor the signal level of the vibrometer controller
' while adjusting the coupling of the optical fibres to the MSA-I-500 scanning head.
' This macro can only be used with MSA-500 systems.
'
' After starting the macro, a time mode continuous single point measurement
' will be started showing the time trace of the signal level in an analyzer window.
' You can then optimize the signal level by adjusting the fibre coupling at
' the MSA-I-500 scan head. For more information please have a look into
' your hardware manual.
'
' The macro modifies the junction box settings such, that the signal level is
' available at the reference channel of the data acquisition board. It is necessary
' the macro can reset this setting at the end of the macro. Therefore you
' have to stop the continuous acquisition when you are finished, not the macro.
' The macro will stop automatically when you stop the acquisition.

'#Uses "..\SwitchToAcquisitionMode.bas"

Option Explicit

Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	If Not SwitchToAcquisitionMode() Then
		MsgBox("Switch to acquisition mode failed.", vbOkOnly)
		End
	End If
	
	If MsgBox("This macro is only applicable for adjusting the coupling of the optical fibre to the MSA-I-500 scanning head. Do not use this macro on other systems or for other purposes. Please read the Hardware manual before proceeding. Do not stop this macro but stop the continuous acquisition when you are finished." + vbCrLf + vbCrLf + "Do you want to continue?", vbYesNo) = vbNo Then
		Exit Sub
	End If

	Dim oControl As FrontEndControl
	Set oControl = Acquisition.Infos.Hardware.ActiveFrontEnd.Control

	If Not oControl.Caps And ptcFrontEndControlCapsOperatingState Then
		MsgBox("The front end doesn't support switching the reference channel to signal level. The macro will abort.")
		Exit Sub
	End If

	Dim AcqSettingsFilename As String
	AcqSettingsFilename = Environ("TEMP") + "\OldAcquisitionSettings.set"
	Settings.Save(AcqSettingsFilename)

	Acquisition.Mode = ptcAcqModeTime

	Acquisition.ActiveProperties.AverageProperties.type = ptcAverageOff

	Dim oAcqPropChannels As ChannelsAcqProperties
	Set oAcqPropChannels = Acquisition.ActiveProperties.ChannelsProperties

    oAcqPropChannels.Item(2).Active = True ' set Ref1 active

	Dim Index As Long

    For Index = 1 To oAcqPropChannels.Count

    	Dim bActive As Boolean
		bActive = Index = 2 ' deactivate others as Ref1

		Dim oAcqPropChannel As ChannelAcqProperties
		Set oAcqPropChannel =  oAcqPropChannels.Item(Index)

        oAcqPropChannel.Active = bActive

 		If bActive Then
			oAcqPropChannel.InputCoupling = ptcInputCouplingDC

			' Try to disable ICP input and select Input Impedance 1 MOhm. If property isn't available, ignore error.
			On Error Resume Next
			oAcqPropChannel.ICPInput = False
			oAcqPropChannel.InputImpedance = 1e6 ' 1 MOhm
			On Error GoTo 0

			oAcqPropChannel.InputRange = 5 ' 5V available for Ni611x or Spectrum MI3025 acuisition board
            oAcqPropChannel.DigitalFilter = Nothing

			' Try to select voltage. If reference vibrometer is used this will fail.
			On Error Resume Next
			oAcqPropChannel.Quantity = ptcPhysicalQuantityVoltage
			On Error GoTo 0

			oAcqPropChannel.IntDiffPhysicalQuantity = oAcqPropChannel.Quantity
			oAcqPropChannel.SEActive = False
		End If

    Next Index

	Acquisition.ActiveProperties.TriggerProperties.Source = ptcTriggerSourceOff

	Acquisition.ActiveProperties.TimeProperties.SampleFrequency 	= 12800 	' 12800 kHz
	Acquisition.ActiveProperties.TimeProperties.Samples				= 2 * 12800	' 2 sec


	Dim OperState As PTCFrontEndOperatingStateType
	OperState = oControl.OperatingState ' save operating state

	oControl.OperatingState = ptcFrontEndOperatingStateAutoFocus ' switch to auto focus measurement. Signal level will be switched to Ref1 channel

	Acquisition.Start(ptcAcqStartContinuous)

	' wait until user stops measurement
	While Acquisition.State <> ptcAcqStateStopped
        Wait 0.1 ' wait 100 ms
    Wend

	oControl.OperatingState = OperState ' restore operating state

	Settings.Load(AcqSettingsFilename, ptcSettingsAll)

	Kill AcqSettingsFilename

End Sub
