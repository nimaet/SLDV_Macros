'#Uses "..\..\SwitchToAcquisitionMode.bas"
'#Uses "Visa32.bas"

Option Explicit

Private Type VISA
	DefaultRM As Long			' handle to VISA resource manager
	Handle As Long				' handle to VISA device
	GeneratorName As String			' VISA device name
End Type

Private Type MinMaxValue
	Value As Double
	Min As Double
	Max As Double
End Type

Private Type GenData
	StartFreq	  As MinMaxValue		' start frequency in MHz
	EndFreq		  As MinMaxValue		' end frequency in MHz
	Amplitude     As MinMaxValue		' generator peak amplitude in V
	StepTime 	  As MinMaxValue		' sweep step time in sec
	SampRes       As Double				' frequency resolution per sample or line in Hz!!!
End Type

Dim m_Visa As VISA
Dim m_Gen  As GenData

Private Const ErrorExit = 1						' This error type will end the program
Private Const ErrorRetry = 2					' This error type will not end the program
Private Const SweepStr = "Sweep"

Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	If Not SwitchToAcquisitionMode() Then
		MsgBox("Switch to acquisition mode failed.", vbOkOnly)
		End
	End If

	If Acquisition.ActiveProperties.Item(ptcAcqPropertiesTypeGenerators).Count <> 0 Then
		MsgBox("Generator is enabled in PSV/VibSoft. Please disable it in the preferences dialog.", vbOkOnly)
		End
	End If


	InitVISA

	'reset all
	WriteStr("*RST")								' ReSeT
	WriteStr("*CLS")								' CLear Status

	With m_Gen

		With .StartFreq
			.Value	= 100.0	' set default start frequency to 100 MHz
			.Min	= 0.1	' set minimum start frequency to 0.1 MHz (=100 kHz)
			.Max 	= 1240.0' set maximum start frequency to 1240 MHz
		End With

		With .EndFreq
			.Value	= 200.0	' set default end frequency to 200 MHz
			.Min	= 0.1	' set minimum end frequency to 0.1 MHz (=100 kHz)
			.Max 	= 1240.0' set maximum end frequency to 1240 MHz
		End With

		With .Amplitude
			.Value 	= 0.01			' set default amplitude to 0.01V

			WriteStr("UNIT:POW V")								' select voltage as amplitude unit		

			WriteStr("SOURce:POWer:LIMit:AMPLitude? Minimum")	' read and set minimum amplitude
			.Min = ToDouble(ReadStr())

			WriteStr("SOURce:POWer:LIMit:AMPLitude? Maximum")	' read and set maximum amplitude
			.Max = ToDouble(ReadStr())
		End With

		With .StepTime
			.Value = 300  ' set default steptime to 300 msec
			.Min   = 50
			.Max   = 10e3
		End With

		ReadSettingsFromRegistry

		Dim oAcq As Acquisition
		Set oAcq = Acquisition

		Dim oAcqProps As AcquisitionProperties
		Set oAcqProps = Acquisition.ActiveProperties

		If (oAcq.Mode = ptcAcqModeFft) Then
			Dim oFftProps As FftAcqProperties
			Set oFftProps = oAcqProps(ptcAcqPropertiesTypeFft)
			m_Gen.SampRes = oFftProps.SampleResolution
		Else
			MsgBox("Unsupported acquisition mode. Only FFT mode is supported", vbOkOnly)
			Exit Sub
		End If
	End With

	Dim ExplanationStr As String
	ExplanationStr = "This macro supports only sweep waveforms." + vbCrLf
	ExplanationStr = ExplanationStr + "To all other possible waveforms you have direct access via PSV/Vibsoft."

	Begin Dialog UserDialog 360,252,"Generator R&S SMBV100A",.GeneratorDlg ' %GRID:10,7,1,1
		Text 20,21,310,14,"TextGeneratorName",.TextGeneratorName

		Text 20,42,300,35,ExplanationStr,.TextStaticExplanation

		TextBox 140,84,140,21,.TextBoxAmplitude
		Text 20,90,90,14,"Amplitude",.TextAmplitude
		Text 300,90,50,14,"V",.TextStaticVoltage

		TextBox 140,119,140,21,.TextBoxStartFreq
		Text 20,123,120,14,"Start Frequency",.TextStaticSF
		Text 300,123,50,14,"MHz",.TextStaticSFMHz

		TextBox 140,154,140,21,.TextBoxEndFreq
		Text 20,158,110,14,"End Frequency",.TextStaticEF
		Text 300,158,50,14,"MHz",.TextStaticEFMHz

		TextBox 140,189,140,21,.TextBoxStepTime
		Text 20,193,90,14,"Step Time",.TextStaticST
		Text 300,193,50,14,"ms",.TextStaticSTMSec

		PushButton 10,224,70,21,"Start",.PushButtonStartGenerator
		PushButton 100,224,70,21,"Stop",.PushButtonStopGenerator
		PushButton 190,224,70,21,"Cancel",.PushButtonCancel
		PushButton 280,224,70,21,"Help",.PushButtonHelp

		OKButton 0,0,0,0 'an unvisible OK button to handle close from the system menue
	End Dialog
	Dim dlg As UserDialog

	Dialog dlg

	CloseVISA
End Sub


Sub WriteStrVI(Str As String)
	Dim status As Long
	Dim i As Long

	status = viWrite(m_Visa.Handle, Str, Len(Str), i)
	If (status < VI_SUCCESS) Then
		Err.Raise ErrorExit
	End If
End Sub


Function ReadStr() As String
	Dim status As Long
	Dim i As Long
	Dim buffer As String
	buffer = Space(1000)
	status = viRead(m_Visa.Handle, buffer, Len(buffer), i)
	If (status < VI_SUCCESS) Then
		Err.Raise ErrorExit
	End If

	ReadStr = Left(buffer, i-1)
End Function


Sub WriteStr(Str As String)
	WaitOPC
	WriteStrVI(Str)
End Sub


Sub WaitOPC
	Do
		'To avoid generator timing problems query if previous commands have been processed.
		'It is important to ensure that the timeout is long enough.
		WriteStrVI("*OPC?")

		Dim Str As String
		Str = ReadStr()

		If (Str = "1") Then
			Exit Sub
		Else
			Wait 0.1
		End If
	Loop
End Sub


Sub CalcAndSetMCParam
	' *******************************************************************************************************
	' ************************ Calculate the multi carrier CW parameters ************************************
	' *******************************************************************************************************

	With m_Gen
		' Get Center frequency of chirp signal in MHz from dialog
		Dim CenterFreq As Double
		CenterFreq = ((.StartFreq.Value + .EndFreq.Value) / 2.0) * 1e6

		' Calculate center frequency of chirp on bin of FFT
		CenterFreq = Round(CenterFreq / .SampRes,0) * .SampRes

		Dim FreqSpan As Double
		FreqSpan = Abs(.StartFreq.Value - .EndFreq.Value ) * 1e6

		' Calculate number of sweep frequencies according to desired span
		Dim Count As Long
		Count = Round(FreqSpan / .SampRes,0)

		' Check for number of sweep frequencies to be odd!!!
		If (Count Mod 2) = 0 Then
			Count = Count + 1
		End If

	End With

	WriteStr("SOUR:BB:MCCW:PRES")																	' set default settings

	SetMultiCarierCW m_Gen.StartFreq.Value *1e6  , 1 , 120.0 *1e6

	SetGates

	SetSweep CenterFreq , Count

End Sub


Function StartGenerator As Boolean
	On Error GoTo CatchErr
	Dim status As Long
	Dim s As String
	Dim i As Long
	Dim ampl_RMS As Double

	WriteStr("*RST")
	WriteStr("*CLS")

	Acquisition.Stop

	CalcAndSetMCParam
	WriteStr("UNIT:POW V")								' select voltage as amplitude unit
	ampl_RMS = Round(m_Gen.Amplitude.Value / Sqr(2.0),5)
	SetCheckDeviceDbl "SOURce:POWer:LEVEL:IMMediate:AMPLitude" , ampl_RMS  , 0 , "V" , "Amplitude"	' set amplitude
	SetCheckDeviceStr "OUTPut:STATe" , "1" , "Ouput state"

	Acquisition.Start(ptcAcqStartSingle)
	SetCheckDeviceStr "SOUR:FREQ:MODE", "SWE", "Sweep Start" 'start sweep

	StartGenerator = True

	CheckSystemErrors

	Exit Function

CatchErr:
	If (Err = ErrorRetry) Then
		StartGenerator = False
	ElseIf (Err = ErrorExit) Then
		MsgBox("Error during start of generator.", vbOkOnly)
		ExitAll
	Else
		MsgBox("Unexpected error during start of generator.", vbOkOnly)
		ExitAll
	End If
End Function


Function StopGenerator As Boolean
	On Error GoTo CatchErr
	Dim status As Long
	Dim s As String
	Dim i As Long

    SetCheckDeviceStr "SOUR:FREQ:MODE" , "CW" , "sweep stop" 	' stop sweep
	SetCheckDeviceStr "OUTPut:STATe"   , "0"  , "Ouput state"

	StopGenerator = True

	Exit Function

CatchErr:
	StopGenerator = False

	MsgBox ("Error during stop of generator.",vbOkOnly)

    ExitAll
End Function



Private Function GeneratorDlg(DlgItem$, Action%, SuppValue&) As Boolean
' -------------------------------------------------------------------------------
' Dialog for generator.
' -------------------------------------------------------------------------------
	Select Case Action%
	Case 1 ' Dialog box initialisation
		DlgText "TextGeneratorName"   , m_Visa.GeneratorName
		DlgText "TextBoxAmplitude"    , ToString(m_Gen.Amplitude.Value)
		DlgText "TextBoxStartFreq"    , ToString(m_Gen.StartFreq.Value)
		DlgText "TextBoxEndFreq"      , ToString(m_Gen.EndFreq.Value)
		DlgText "TextBoxStepTime"     , ToString(m_Gen.StepTime.Value)

	Case 2 ' Values changed or buttons clicked
		GeneratorDlg = True
		Select Case DlgItem$

		Case "PushButtonStartGenerator"
			On Error Resume Next
			If Not SetDlgValues(DlgText("TextBoxAmplitude"), DlgText("TextBoxStartFreq") , DlgText("TextBoxEndFreq"), DlgText("TextBoxStepTime")) Then
				Exit Function
			End If
			If (StartGenerator = False) Then
				Exit Function
			End If
			SaveSettingsToRegistry
			GeneratorDlg = False

		Case "PushButtonStopGenerator"
			StopGenerator
			GeneratorDlg = False

		Case "PushButtonHelp"
			Dim Msg As String
			Msg =       "After changing A/D settings recall this macro. "                            + vbCrLf + vbCrLf
			Msg = Msg +	"For gated measurements:"                                                    + vbCrLf
			Msg = Msg +	"  - In the trigger settings dialog select:"                                 + vbCrLf
			Msg = Msg + "      rising edge, external source, and pre-trigger 0%."                    + vbCrLf
			Msg = Msg + "  - Connect SMBV100A ""MARKER 1"" with vibrometer controller ""TRIG IN"". " + vbCrLf
            Msg = Msg + "  - Connect SMBV100A ""MARKER 2"" with LeCroy ""AUX IN"". "                 + vbCrLf
			MsgBox( Msg, vbOkOnly, "Help" )

	    Case Else '"PushButtonCancel" and "OK" (OK stands for Close from system menue)
			GeneratorDlg = False

	    End Select
	Case 3 ' Text box or combo box changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Wait 0.5
	Case 6 ' Function key
	Case Else
		GeneratorDlg = False
	End Select
End Function


Sub InitVISA
	On Error GoTo CatchErr
	Dim status As Long

	status = viOpenDefaultRM(m_Visa.DefaultRM)
    If (status < VI_SUCCESS) Then
    	MsgBox ("VISA could not be initialized. Please verify that the NI-VISA driver is installed.",vbOK)
    	Exit All
    End If

	FindSMBV

	status = viOpen(m_Visa.DefaultRM, m_Visa.GeneratorName, VI_NULL, 10000, m_Visa.Handle)
    If (status < VI_SUCCESS) Then
    	Err.Raise 1
    End If

	status = viSetAttribute(m_Visa.Handle, VI_ATTR_TMO_VALUE, 10000) ' set timeout to 10 s
    If (status < VI_SUCCESS) Then
    	Err.Raise 1
    End If

	Exit Sub

CatchErr:
	MsgBox ("Generator could not be initialized.",vbOkOnly)
	ExitAll
End Sub

Private Sub FindSMBV
	Dim status As Long

	With m_Visa
		'Read visa name from registry and try it
		.GeneratorName = GetSetting("GeneratorSMBV100A", "VISA", "GeneratorName" , .GeneratorName)
		If ( .GeneratorName <> "" ) Then
			If ( IsSMBV100A(.GeneratorName) ) Then
				Exit Sub 'found
			End If
		End If

		Const SMBV_VISA_IPAdress = "TCPIP0::192.168.0.3::inst0::INSTR"
		Const SMBV_VISA_Alias    = "SMBV100A"

		If      (IsSMBV100A(SMBV_VISA_IPAdress)) Then 'try the polytec standard IP address for the SMBV.'Remarke: If available the device can be opened without declaration in NI-Max
			.GeneratorName = SMBV_VISA_IPAdress
			'do not exit here, Generator name will be saved to registry at the end of the subroutine

		ElseIf  (IsSMBV100A(SMBV_VISA_Alias)) Then 'try the default Alias
			.GeneratorName = SMBV_VISA_Alias
			'do not exit here, Generator name will be saved to registry at the end of the subroutine

		Else 'loop in all available resources
			Dim Count As Long
			Dim findHandler As Long
			Dim descriptor As String * VI_FIND_BUFLEN
			Dim lFound As Long
			lFound = 0

			status = viFindRsrc(.DefaultRM, "?*INSTR", findHandler, Count, descriptor)
			If (status = VI_SUCCESS) Then
				Do
					If (IsSMBV100A(descriptor)) Then
						.GeneratorName = descriptor
						lFound = lFound + 1
					End If
				Loop While (VI_SUCCESS = viFindNext(findHandler, descriptor)) ' get next visa resource

				viClose(findHandler)
			End If

			Select Case lFound
			Case 0
				MsgBox ("Generator SMBV100A could not be found.",vbOkOnly)
				ExitAll
			Case Is > 1
				MsgBox ("More the one SMBV100A found. using: " + .GeneratorName ,vbOkOnly)
			End Select
		End If

		' save to registry
		SaveSetting "GeneratorSMBV100A", "VISA", "GeneratorName" , .GeneratorName
	End With
End Sub

Private Function IsSMBV100A(ByRef descriptor As String) As Boolean
	Dim status As Long
	IsSMBV100A = False

	With m_Visa
		status = viOpen(.DefaultRM, descriptor, VI_NULL, VI_NULL, .Handle)
		If (status = VI_SUCCESS) Then
			On Error GoTo CatchErr

		    WriteStrVI("*IDN?")
			If (InStr(ReadStr(), "Rohde&Schwarz,SMBV100A")) Then
				IsSMBV100A = True
			End If

CatchErr:
			viClose(.Handle)
		End If
	End With
End Function

Private Sub ExitAll
	viClose(m_Visa.Handle)
	viClose(m_Visa.DefaultRM)

	Exit All
End Sub

Sub CloseVISA
	On Error GoTo CatchErr
	Dim status As Long

	status = viClose(m_Visa.Handle)
	If (status < VI_SUCCESS) Then
		Err.Raise 1
	End If

	status = viClose(m_Visa.DefaultRM)
	If (status < VI_SUCCESS) Then
		Err.Raise 1
	End If

	Exit Sub

CatchErr:
	ExitAll
End Sub

Public Sub ReadSettingsFromRegistry()
    On Error Resume Next ' ignore corrupt (non-convertible) registry entries

	With m_Gen
	    .Amplitude .Value = ToDouble(GetSetting("GeneratorSMBV100A", "DialogSettings", "Amplitude"	, ToString(.Amplitude.Value)))
	    .StartFreq .Value = ToDouble(GetSetting("GeneratorSMBV100A", "DialogSettings", "StartFreq"	, ToString(.StartFreq.Value)))
	    .EndFreq   .Value = ToDouble(GetSetting("GeneratorSMBV100A", "DialogSettings", "EndFreq"	, ToString(.EndFreq  .Value)))
	    .StepTime  .Value = ToDouble(GetSetting("GeneratorSMBV100A", "DialogSettings", "StepTime"	, ToString(.StepTime .Value)))
	End With
End Sub


Public Sub SaveSettingsToRegistry()
	With m_Gen
		SaveSetting "GeneratorSMBV100A", "DialogSettings", "Amplitude"	,ToString(.Amplitude.Value)
		SaveSetting "GeneratorSMBV100A", "DialogSettings", "StartFreq"	,ToString(.StartFreq.Value)
		SaveSetting "GeneratorSMBV100A", "DialogSettings", "EndFreq"    ,ToString(.EndFreq  .Value)
		SaveSetting "GeneratorSMBV100A", "DialogSettings", "StepTime"	,ToString(.StepTime .Value)
	End With
End Sub

Private Function SetCheckVal(ByRef Value As MinMaxValue, ByVal InputValue As Double, ByVal StrName As String, ByVal StrUnit As String )
	With Value
		If (InputValue < .Min Or InputValue > .Max) Then
			MsgBox( StrName + " must be between " + ToString(.Min) + " and " + ToString(.Max) + " " + StrUnit + ".", vbOkOnly)
			SetCheckVal = False
		Else
			.Value = InputValue
			SetCheckVal = True
		End If
	End With
End Function

Private Function SetDlgValues(ByVal StrAmplitude As String, ByVal StrStartFreq As String, ByVal StrEndFreq As String, ByVal StrStepTime As String )
	SetDlgValues = False
	With m_Gen
		If Not SetCheckVal(.Amplitude, ToDouble(StrAmplitude), "Amplitude" , "V") Then
			Exit Function
		End If

		If Not SetCheckVal(.StartFreq, ToDouble(StrStartFreq), "Start frequency" , "MHz") Then
			Exit Function
		End If

		If Not SetCheckVal(.EndFreq  , ToDouble(StrEndFreq)  , "End frequency"   , "MHz") Then
			Exit Function
		End If

		If Not SetCheckVal(.StepTime, ToDouble(StrStepTime), "Step time" , "sec") Then
			Exit Function
		End If
	End With

	SetDlgValues = True
End Function


Private Sub SetCheckDeviceStr(ByVal CmdStr As String, ByVal ValueStr As String , ByVal ErrorStr As String)
	WriteStr(CmdStr + " " + ValueStr)
	WriteStr(CmdStr + "?")
	If ( ReadStr() <> ValueStr ) Then
		MsgBox("Error: " + ErrorStr + " (" + ValueStr + ") could not be set.", vbOkOnly)
		Err.Raise ErrorRetry
	End If
End Sub

Private Sub SetCheckDeviceDbl(ByVal CmdStr As String, ByVal dValue As Double, ByVal dAccuracy As Double, ByVal UnitStr As String ,ByVal ErrorStr As String)
	WriteStr(CmdStr + " " + ToString(dValue) + UnitStr)
	WriteStr(CmdStr + "?")
	If (Abs(dValue - ToDouble(ReadStr())) > dAccuracy) Then
		MsgBox("Error: " + ErrorStr + " (" + ToString(dValue) + UnitStr + ") could not be set.", vbOkOnly)
		Err.Raise ErrorRetry
	End If
End Sub


Private Sub SetMultiCarierCW(ByVal CenterFreq As Double , ByVal NCarrier As Long, ByVal CarrierSpacing As Double)
	SetCheckDeviceStr "SOUR:BB:MCCW:STAT"      , ToString(1)              , "enable multi carrier cw"	' enable multi carrier cw

	SetCheckDeviceDbl "FREQUENCY"              , CenterFreq      , 0    , "Hz" , "Frequency"		' set center frequency
	SetCheckDeviceStr "SOUR:BB:MCCW:CARR:COUN" , ToString(NCarrier)     , "Number Of Carriers"		' set number of carriers
	SetCheckDeviceDbl "SOUR:BB:MCCW:CARR:SPAC" , CarrierSpacing  , 0.01 , "Hz" , "Carrier Spacing"	' set carrier spacing

	SetCheckDeviceStr "SOUR:BB:MCCW:CFAC:MODE" , "CHIR" 	            , "Optimize Crest Factor"   ' optimize crest factor for chirp
End Sub

Private Sub SetSweep(ByVal CenterFreq As Double , ByVal Count As Long)
	If (Count < 1 Or ((Count Mod 2) <> 1) ) Then ' count should be positive odd value
        MsgBox("Error: Count value is wrong", vbOkOnly)
        Err.Raise(ErrorExit)
    End If

   	SetCheckDeviceStr "SOUR:SWE:FREQ:MODE"     , "AUTO"                     , "Sweep Freq Mode"
	SetCheckDeviceStr "TRIG:FSW:SOUR"          , "AUTO"                     , "Trigger Sweep Source"
	SetCheckDeviceStr "SOUR:SWE:FREQ:SPAC"     , "LIN"                      , "Sweep Spacing"
	SetCheckDeviceStr "SOUR:SWE:FREQ:SHAP"     , "SAWT"                     , "Sweep Shape"

	SetCheckDeviceDbl "SOUR:FREQ:SPAN"         , 0             , 0    , "Hz" , "Sweep Spawn"		' reset span to be able to set all center frequencies
	SetCheckDeviceDbl "SOUR:FREQ:CENT"         , CenterFreq    , 0    , "Hz" , "Sweep Center" 		' write sweep center freq

	With m_Gen
		SetCheckDeviceDbl "SOUR:FREQ:SPAN"         , .SampRes * (Count - 1) , 0    , "Hz" , "Sweep Span"
		SetCheckDeviceDbl "SOUR:SWE:FREQ:STEP:LIN" , .SampRes               , 0    , "Hz" , "Sweep Step"
		SetCheckDeviceDbl "SOUR:SWE:FREQ:DWEL"     , .StepTime.Value * 1e-3 , 1e-6 , "s"  , "Sweep StepTime"
	End With


	Count = 2 * Count ' measure 2 cycles
	If Count > 10000 Then
		Count = 10000 ' average count should not be greater than 10000
	End If

	With Acquisition.ActiveProperties.AverageProperties
		.Count = Count
		.type = ptcAveragePeakhold
	End With
End Sub


Private Sub SetGates
	' *******************************************************************************************************
	' ************************ Calculate the trigger/Laser Gate settings ************************************
	' *******************************************************************************************************
	Dim LaserGateRiseTime As Double' = 5.6e-6	' s
	LaserGateRiseTime = Application.Acquisition.Infos.Vibrometers.Item(1).Controllers.Item(0).SensorHeads.Item(0).SensorHeadInfo.LaserDelay

	Dim DemodTime As Double
	DemodTime = 0.3					' s During demodulation time laser can be off

	WriteStr("SOUR:BB:MCCW:CLOC?")	' ask for clock frequency
	Dim Clockrate As Double
	Clockrate = ToDouble(ReadStr())		' read clock frequency according to current Multi CW settings

    Const PreRoll = 0.6

    ' Calculate laser on time (has to be measurement time + pre and post roll (max. 100%) + rise time of laser gate (~5.6 µs))
    Dim LaserOnTime As Long                 ' In Samples
    LaserOnTime = Round(Clockrate * ((2.0 * PreRoll + 1.0) / m_Gen.SampRes + LaserGateRiseTime), 0)   ' has to be an integer

    ' Calculate oscilloscope trigger delay (has to be pre roll (max. 50%) + rise time of laser gate (~5.6 µs))
    Dim ScopeTrigDelayOnTime As Long
    ScopeTrigDelayOnTime = Round(Clockrate * (PreRoll / m_Gen.SampRes + LaserGateRiseTime), 0)   ' has to be an integer

    ' Calculate duration of measurement period in samples related to clock rate
    Dim MeasPeriodInSamples As Long
    MeasPeriodInSamples = Round(Clockrate / m_Gen.SampRes, 0)

    ' Calculate No of cycles of measurement period to fit in 0.5 s (~ demodulation duration)
    Dim NMeasPeriod As Long
    NMeasPeriod = Int(DemodTime * m_Gen.SampRes + 0.5) * MeasPeriodInSamples

    ' Check if the number of samples exceeds the generator maximum
    If NMeasPeriod > 16777215 Then
        NMeasPeriod = Int(16777215 / MeasPeriodInSamples) * MeasPeriodInSamples ' reduce the number of samples to a multiple of the period length
    End If

    ' Set Laser on off gate (marker1)
    SetGate 1, LaserOnTime , NMeasPeriod , 0

    ' Set Trigger for Scope (marker2) (trigger signal is the rising edge)
    SetGate 2, LaserOnTime - ScopeTrigDelayOnTime, NMeasPeriod , ScopeTrigDelayOnTime
End Sub

Private Sub SetGate(ByVal Marker As Integer, ByVal OnTime As Long, ByVal WholeTime As Long, ByVal OffsetTime As Long)
	If (Marker < 0 Or Marker > 2) Then
        MsgBox("Error: Marker should be one or two", vbOkOnly)
        Err.Raise(ErrorExit)
    End If

	If (OnTime >= WholeTime) Then
        MsgBox("Error: OnTime should be less WholeTime", vbOkOnly)
        Err.Raise(ErrorExit)
    End If

    Dim OffTime As Long                ' In Samples
    OffTime = WholeTime - OnTime

    Dim CommandPreFix As String
    CommandPreFix = "SOUR:BB:MCCW:TRIG:OUTP" + ToString(Marker)

   	SetCheckDeviceStr CommandPreFix + ":MODE" , "RAT"        	    , "Marker On/Off Ratio Mode " 		'The command defines the signal for the selected marker output. A regular marker signal corresponding to the Time Off / Time On specifications in the commands.
   	SetCheckDeviceStr CommandPreFix + ":ONT"  , ToString(OnTime)    , "Number of on time samples" 		'The command sets the number of samples in a period (ON time + OFF time) during which the marker signal in setting SOURce:BB:MCCW:TRIGger:OUTPut:MODE RATio on the marker outputs is ON.
   	SetCheckDeviceStr CommandPreFix + ":OFFT" , ToString(OffTime)   , "Number of off time samples" 		'The command sets the number of samples in a period (ON time + OFF time) during which the marker signal in setting SOURce:BB:MCCW:TRIGger:OUTPut:MODE RATio on the marker outputs is ON.
   	SetCheckDeviceStr CommandPreFix + ":DEL"  , ToString(OffsetTime), "Number of offset time samples" 	'The command defines the delay between the signal on the marker outputs and the start of the signals, expressed in terms of samples.
End Sub

'************************************
'**Check for SMBV100A system errors**
'************************************
Private Sub CheckSystemErrors()
	Const c_strNoError As String = "0,""No error"""

	WriteStr("SYSTEM:ERROR?")
	Dim strError As String
	strError = ReadStr()
	If ( strError = c_strNoError ) Then
		Exit Sub
	End If

	Dim strErrorAll As String
	Do
		strErrorAll = strErrorAll + strError + vbCrLf
		WriteStr("SYSTEM:ERROR:NEXT?")
		strError = ReadStr()

		Dim bLoop As Boolean
	Loop Until (strError = c_strNoError)

	MsgBox("SMBV100A Error(s): " + vbCrLf + vbCrLf + strErrorAll + vbCrLf + " Please refer to the SMBV100A manual.", vbOkOnly)
	Err.Raise ErrorRetry

End Sub


Private Function GetSeperator() As String
	GetSeperator = Format$(0, ".") ' --> http://us.generation-nt.com/answer/api-function-determin-decimal-seperator-help-8753082.html
End Function

Private Function ToString(ByVal dVal As Double) As String
    Dim StrVal As String
	StrVal = CStr(dVal)
	ToString = Replace(StrVal, GetSeperator(), ".")
End Function

Private Function ToDouble(ByVal StrVal As String) As Double
	StrVal = Replace(StrVal, ".", GetSeperator())
	ToDouble = CDbl(StrVal)
End Function
