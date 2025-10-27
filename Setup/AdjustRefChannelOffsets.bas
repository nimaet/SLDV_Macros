' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This macro allows you to adjust the reference channel offset for PSV-500-H
' and MSA-100-3D-H systems.
' This macro can only be used with PSV-500-H and MSA-100-3D-H systems.
'
' After starting the macro, FFT single point measurements will be started and 
' the offset will be calculated. When the offset is stable it will be saved in
' the registry. This offset is used for reference channels within this range
' and within this coupling.

'#Language "WWB.Net"
'#Uses "..\SwitchToAcquisitionMode.bas"

Imports System		'for the math object
Imports System.IO	'for the Path object
Imports System.Collections.Generic
Imports System.Globalization
Imports Microsoft.Win32

Option Explicit

Dim ethernetAcquisitionKey As String = "SOFTWARE\Polytec\PolyAcquisition\EthernetAcquisition\"

Public Class AnalogChannelOptions
    Implements IComparable(Of AnalogChannelOptions)

    Public Property Active As Boolean
	Public Property ShortName As String
    Public Property Reference As Boolean
    Public Property InputRange As Double
	Public Property InputCoupling As PTCInputCoupling
    Public Property ICPInput As Boolean
    Public Property CurrentOffset As Double
    Public Property LastOffset As Double
    Public Property OldOffset As Double

    Public ReadOnly Property RegistryKeyPath As String
    	Get
			Return ethernetAcquisitionKey + ShortName + "\" + InputRange.ToString + "V" + CouplingString(InputCoupling)
    	End Get
    End Property

 	Public Function CompareTo(ByVal other As AnalogChannelOptions) As Integer _
        Implements System.IComparable(Of AnalogChannelOptions).CompareTo

        ' Compare the ShortName.
        Dim Compare As Integer
        Compare = String.Compare(Me.ShortName, other.ShortName, True)
        Return Compare
    End Function

End Class

Public Class AD4BoardVersion
	Private m_ad4Board0 As Double
	Private m_ad4Board1 As Double

	Public Sub New(ByVal ad4Board0 As Double, ByVal ad4Board1 As Double)
		m_ad4Board0 = ad4Board0
		m_ad4Board1 = ad4Board1
	End Sub

	Public ReadOnly Property AD4BoardVersion0 As Double
		Get
			Return m_ad4Board0
		End Get
	End Property

	Public ReadOnly Property AD4BoardVersion1 As Double
		Get
			Return m_ad4Board1
		End Get
	End Property

End Class


Dim fileName As String
Dim registryKeyName As String = "Offset[V]"
Dim minDeltaDCLimit As Double = 0.00003  '30 µV
Dim minDeltaACLimit As Double = 0.00003  '30 µV
Dim maxOffset As Double = 0.2 '200 mV
Dim maxACOffset As Double = 0.001 '1 mV
Dim versionWithoutACOffset As Double = 1.0305 'AD 4 board version with highpass to avoid AC offset.

Sub Main
	Dim registryKeysReseted = False
	Try
		Dim newValuesValid As Boolean = False
		Dim analogChannels As New List(Of AnalogChannelOptions)
		analogChannels = ReadADSettings()

		Dim versionAD4Board As AD4BoardVersion = ReadAD4BoardVersion(analogChannels.Count > 4)

		Dim count As Integer = 0
		Dim activeAnalogChannelsToAdjust As New List(Of AnalogChannelOptions)
		Dim activeAnalogChannelsToCheck As New List(Of AnalogChannelOptions)
		For Each ac As AnalogChannelOptions In analogChannels
			If ac.Active Then
				Dim ad4Version As Double = If(count < 4, versionAD4Board.AD4BoardVersion0, versionAD4Board.AD4BoardVersion1)
				If ad4Version >= versionWithoutACOffset And ac.InputCoupling = PTCInputCoupling.ptcInputCouplingAC Then
					activeAnalogChannelsToCheck.Add(ac)
				Else
					activeAnalogChannelsToAdjust.Add(ac)
				End If
			End If
			count = count + 1
		Next

		If (activeAnalogChannelsToAdjust.Count = 0 And activeAnalogChannelsToCheck.Count = 0) Then
			Throw New exception("No reference channel is selected." + vbCrLf + _
				"Please select one or more reference channels in the acquisition settings.")
		End If

		If Explanation(activeAnalogChannelsToAdjust, activeAnalogChannelsToCheck) <> True Then
			Exit Sub
		End If

		ResetRegistryKeys(activeAnalogChannelsToAdjust)
		registryKeysReseted = True

		If versionAD4Board.AD4BoardVersion0 >= versionWithoutACOffset Or versionAD4Board.AD4BoardVersion1 >= versionWithoutACOffset Then
			SetAllACOffsetsTo0
		End If

		Prepare(activeAnalogChannelsToAdjust, activeAnalogChannelsToCheck)

		MeasureOffsets(activeAnalogChannelsToAdjust, activeAnalogChannelsToCheck)

		VerifyDataToCheck(activeAnalogChannelsToCheck)
		newValuesValid = VerifyDataToAdjust(activeAnalogChannelsToAdjust)

		SetRegistryKeys(activeAnalogChannelsToAdjust, newValuesValid)

	Catch exception As exception
		If registryKeysReseted Then
			SetRegistryKeys(activeAnalogChannelsToAdjust, False)
		End If
		Dim message = exception.Message + vbCrLf + "The macro will finish now."
		MsgBox(message, vbExclamation Or vbOkOnly)

	Finally
		Try
			If (fileName <> Nothing) Then
				Application.Settings.Load(fileName, PTCSettings.ptcSettingsAcquisition Or PTCSettings.ptcSettingsWindows)
				Kill(fileName)
			End If
		Catch ' nothing
		End Try

	End Try
End Sub

Private Function Explanation(ByRef analogChannelsToAdjust As List(Of AnalogChannelOptions), _
	ByRef analogChannelsToCheck As List(Of AnalogChannelOptions)) As Boolean

	Dim adjustReferenceChannelsText As String
	For Each acAdjust As AnalogChannelOptions In analogChannelsToAdjust
		adjustReferenceChannelsText = adjustReferenceChannelsText + vbCrLf + vbTab + acAdjust.ShortName + vbTab + _
			acAdjust.InputRange.ToString + " V" + vbTab + CouplingString(acAdjust.InputCoupling)
	Next

	Dim checkReferenceChannelsText As String
	For Each acCheck As AnalogChannelOptions In analogChannelsToCheck
		checkReferenceChannelsText = checkReferenceChannelsText + vbCrLf + vbTab + acCheck.ShortName + vbTab + _
			acCheck.InputRange.ToString + " V" + vbTab + CouplingString(acCheck.InputCoupling)
	Next

	Dim explanationText = "With this macro you can adjust the offsets for the selected reference channels, " + _
		"for their current range and coupling." + vbCrLf + "Please terminate them with a 50 Ohm resistance."

	If adjustReferenceChannelsText.Length > 0 Then
		explanationText = explanationText + vbCrLf + vbCrLf + _
			"The following channels will be adjusted:" + vbCrLf + _
			adjustReferenceChannelsText
	End If

	If checkReferenceChannelsText.Length > 0 Then
		explanationText = explanationText + vbCrLf + vbCrLf + _
			"The following channels will be checked, no adjustment will be done:" + vbCrLf + _
			checkReferenceChannelsText
	End If

	Return MsgBox(explanationText, vbInformation Or vbOkCancel) = vbOK
End Function

Public Function CouplingString(InputCoupling As PTCInputCoupling) As String
	Select Case InputCoupling
		Case PTCInputCoupling.ptcInputCouplingAC
			Return "AC"
		Case PTCInputCoupling.ptcInputCouplingDC
			Return "DC"
		Case Else
			Throw New exception("Invalid input coupling.")
	End Select
End Function

Private Function ReadADSettings() As List(Of AnalogChannelOptions)

	If Not SwitchToAcquisitionMode() Then
		Throw New exception("Switch to acquisition mode failed.")
	End If

	If Not Acquisition.Infos.HasHardware Or Acquisition.Infos.Hardware.AcqBoard <> PTCAcqBoardType.ptcAcqBoardPolyEth Then
		Throw New exception("This macro is only possible for systems with ethernet data acquisition.")
	End If

	Dim analogChannels As New List(Of AnalogChannelOptions)

	Dim oChannelsAcqProps As ChannelsAcqPropertiesContainer = Acquisition.ActiveProperties.ChannelsProperties

	' loop over all channels
	For Each oChannelAcqProps As ChannelAcqPropertiesContainer In oChannelsAcqProps
		If (oChannelAcqProps.type = PTCChannelType.ptcChannelTypeAnalog) Then
			' channel is active and analog channel
			Dim analogChannel As New AnalogChannelOptions
			analogChannel.Active = oChannelAcqProps.Active
			analogChannel.ShortName = oChannelAcqProps.ShortName
			analogChannel.Reference = oChannelAcqProps.Reference
			analogChannel.InputRange = oChannelAcqProps.InputRange
			analogChannel.InputCoupling = oChannelAcqProps.InputCoupling
			analogChannel.ICPInput = oChannelAcqProps.ICPInput

			analogChannels.Add(analogChannel)
		End If
	Next

	Return analogChannels
End Function

Private Sub Prepare(ByRef analogChannelsToAdjust As List(Of AnalogChannelOptions), _
	ByRef analogChannelsToCheck As List(Of AnalogChannelOptions))

	fileName = Path.GetTempPath() + "AdjustRefChannelOffsets.set"
	Application.Settings.Save(fileName)

	Acquisition.Mode = POLYPROPERTIESLib.PTCAcqMode.ptcAcqModeFft
	Acquisition.ActiveProperties.AverageProperties.type = POLYPROPERTIESLib.PTCAverageType.ptcAverageOff
	Acquisition.ActiveProperties.TriggerProperties.Source = PolyScopeLib.PTCTriggerSource.ptcTriggerSourceOff
    Acquisition.ActiveProperties.FftProperties.Bandwidth = 800
	Acquisition.ActiveProperties.FftProperties.Lines = 1600
	Acquisition.GeneratorsOn = False

	Dim oChannelsAcqProps As ChannelsAcqPropertiesContainer = Acquisition.ActiveProperties.ChannelsProperties

	' loop over all channels
	For Each oChannelAcqProps As ChannelAcqPropertiesContainer In oChannelsAcqProps
		If (oChannelAcqProps.type = PTCChannelType.ptcChannelTypeAnalog) Then
			' channel is analog channel

			Dim Found As Boolean = False
			Dim analogChannel As New AnalogChannelOptions
			For Each acAdjust As AnalogChannelOptions In analogChannelsToAdjust
				If (acAdjust.ShortName = oChannelAcqProps.ShortName) Then
					analogChannel = acAdjust
					Found = True
					Exit For
				End If
			Next

			If Not Found Then
				For Each acCheck As AnalogChannelOptions In analogChannelsToCheck
					If (acCheck.ShortName = oChannelAcqProps.ShortName) Then
						analogChannel = acCheck
						Found = True
						Exit For
					End If
				Next
			End If

			If Not Found Then
				oChannelAcqProps.Active = False
				Continue For
			End If

			oChannelAcqProps.Active = True
			oChannelAcqProps.Reference = analogChannel.Reference
			oChannelAcqProps.InputRange = analogChannel.InputRange
			oChannelAcqProps.InputCoupling = analogChannel.InputCoupling
			Try
				oChannelAcqProps.ICPInput = analogChannel.ICPInput
			Catch 'ignore error
			End Try
			oChannelAcqProps.DigitalFilter = Nothing
			oChannelAcqProps.Quantity = PTCPhysicalQuantity.ptcPhysicalQuantityVoltage
			oChannelAcqProps.Calibration = 1
			oChannelAcqProps.SEActive = False
		End If
	Next
End Sub

Private Sub MeasureOffsets(ByRef analogChannelsToAdjust As List(Of AnalogChannelOptions), _
	ByRef analogChannelsToCheck As List(Of AnalogChannelOptions))

	For index As Integer = 1 To 25
		'Start Measurement
		Acquisition.Start(PTCAcqStartMode.ptcAcqStartSingle)

		' wait until user stops measurement
		While Acquisition.State <> PTCAcqState.ptcAcqStateStopped
			Wait(0.1) ' wait 100 ms
		End While

		Dim oAnalyzerWindow As AnalyzerWindow = Acquisition.Document.Windows(1)
		Dim oAnalyzerView As AnalyzerView = oAnalyzerWindow.AnalyzerView
		oAnalyzerView.Settings.DisplaySettings.Domain = PTCDomainType.ptcDomainTime
		Dim dataSize As Integer = oAnalyzerView.XAxis.MaxCount

		For Each acAdjust As AnalogChannelOptions In analogChannelsToAdjust
			SetChannelOffsets(acAdjust, oAnalyzerView)
		Next

		For Each acCheck As AnalogChannelOptions In analogChannelsToCheck
			SetChannelOffsets(acCheck, oAnalyzerView)
		Next

		If index <= 2 Then
			Continue For
		End If

		Dim allChannelsValid As Boolean = True
		For Each acAdjust In analogChannelsToAdjust
			If Not IsOffsetValid(acAdjust) Then
				allChannelsValid = False
				Exit For
			End If
		Next

		If allChannelsValid Then
			For Each acCheck In analogChannelsToCheck
				If Not CheckACLimit(acCheck) Then
					allChannelsValid = False
					Exit For
				End If
			Next
		End If

		If allChannelsValid Then
			Exit For
		End If
	Next
End Sub

Private Sub SetChannelOffsets(ByRef analogChannel As AnalogChannelOptions, ByRef oAnalyzerView As AnalyzerView)

	oAnalyzerView.Settings.DisplaySettings.Channel = analogChannel.ShortName

	Dim dataSize As Integer = oAnalyzerView.XAxis.MaxCount
	Dim Data() As Single = oAnalyzerView.GetDataSection(0, 0, 0, dataSize - 1)

	Dim result As Double = 0
	For Each value As Single In Data
	    result += value
	Next
	analogChannel.LastOffset = analogChannel.CurrentOffset
	analogChannel.CurrentOffset = result / dataSize
End Sub

Private Function VerifyDataToAdjust(ByRef analogChannels As List(Of AnalogChannelOptions)) As Boolean
	If analogChannels.Count = 0 Then
		Return False
	End If

	Dim failedChannelsText As String
	Dim successfulChannelsText As String
	GetChannelsText(analogChannels, successfulChannelsText, failedChannelsText, False)

	Dim explanationText As String
	If failedChannelsText.Length > 0 And successfulChannelsText.Length = 0 Then
		explanationText = _
			"The offset measurement failed for these channels:" + vbCrLf + _
			failedChannelsText + vbCrLf + vbCrLf + _
			"This might indicate a hardware problem or incorrect cabling." + vbCrLf

		MsgBox(explanationText, vbExclamation Or vbOkOnly)
		Return False
	End If

	explanationText = _
		"The offset measurement was successful for these channels:" + vbCrLf + _
		successfulChannelsText + vbCrLf + vbCrLf + _
		"Would you like to adjust these channels with their offset values?" + vbCrLf

	Return MsgBox(explanationText, vbInformation Or vbOkCancel) = vbOK
End Function

Private Sub VerifyDataToCheck(ByRef analogChannels As List(Of AnalogChannelOptions))
	If analogChannels.Count = 0 Then
		Exit Sub
	End If

	Dim failedChannelsText As String
	Dim successfulChannelsText As String
	GetChannelsText(analogChannels, successfulChannelsText, failedChannelsText, True)

	Dim explanationText As String
	If failedChannelsText.Length > 0 Then
		explanationText = _
			"The offset check failed for these channels:" + vbCrLf + _
			failedChannelsText + vbCrLf + vbCrLf + _
			"This might indicate a hardware problem or incorrect cabling." + vbCrLf
	Else
		explanationText = _
			"The offset check was successful for these channels:" + vbCrLf + _
			successfulChannelsText + vbCrLf + vbCrLf
	End If

	Dim msgBoxStyle As VbMsgBoxStyle = If(failedChannelsText.Length > 0, vbExclamation Or vbOkOnly, vbInformation Or vbOkOnly)
	MsgBox((explanationText, msgBoxStyle)
End Sub

Private Sub GetChannelsText(ByRef analogChannels As List(Of AnalogChannelOptions), _
	ByRef successfulChannelsText As String, ByRef failedChannelsText As String, acLimit As Boolean)

	successfulChannelsText = String.Empty
	failedChannelsText = String.Empty

	For Each ac As AnalogChannelOptions In analogChannels
		Dim valid As Boolean = If(acLimit, CheckACLimit(ac), IsOffsetValid(ac))
		If Not valid Then
			failedChannelsText = failedChannelsText + vbCrLf + ac.ShortName + vbTab + _
				ac.InputRange.ToString + " V" + vbTab + CouplingString(ac.InputCoupling) + vbTab + _
				ac.CurrentOffset.ToString(CultureInfo.InvariantCulture) + " V"
		Else
			successfulChannelsText = successfulChannelsText + vbCrLf + ac.ShortName + vbTab + _
				ac.InputRange.ToString + " V" + vbTab + CouplingString(ac.InputCoupling) + vbTab + _
				ac.CurrentOffset.ToString(CultureInfo.InvariantCulture) + " V"
		End If
	Next
End Sub

Private Function IsOffsetValid(ByRef ac As AnalogChannelOptions) As Boolean
	Dim limit As Double = If(ac.InputCoupling = PTCInputCoupling.ptcInputCouplingAC, minDeltaACLimit, minDeltaDCLimit)
	If ((Abs(ac.LastOffset - ac.CurrentOffset) <= limit) And (Abs(ac.CurrentOffset) <= maxOffset)) Then
		Return True
	Else
		Return False
	End If
End Function

Private Function CheckACLimit(ByRef ac As AnalogChannelOptions) As Boolean
	If (Abs(ac.CurrentOffset) <= maxACOffset) Then
		Return True
	Else
		Return False
	End If
End Function

Private Function ReadAD4BoardVersion(twoAD4Boards As Boolean) As AD4BoardVersion
	Try
		Dim MacAddress As String
		MacAddress = Acquisition.Infos.Hardware.ActiveFrontEnd.MacAddress

		Dim factory As New Polytec.IO.FrontEnd.FrontEndFactory
		Dim deviceFactory As Polytec.IO.FrontEnd.FrontEndDeviceFactory
		deviceFactory = factory.GetById(MacAddress)

		Dim deviceInterface As Polytec.IO.FrontEnd.IFrontEndDevice
		deviceInterface = deviceFactory.CreateDevice(Polytec.IO.FrontEnd.FrontEndDeviceType.Raw, 0)

		Dim device As Polytec.IO.FrontEnd.FrontEndDevice
		device = TryCast(deviceInterface, Polytec.IO.FrontEnd.FrontEndDevice)

		Dim versionAD4Board0 As Double = ReadAD4BoardVersionString(device.SendRaw(36, 3, 16, 0, 0, 1))	'FW-Version AD4-Board0 (Ref 1-4); String
		Dim versionAD4Board1 As Double = -1.0

		If twoAD4Boards Then
			versionAD4Board1 = ReadAD4BoardVersionString(device.SendRaw(36, 3, 16, 1, 0, 1))	'FW-Version AD4-Board1 (Ref 5-8); String
		End If

    	ReadAD4BoardVersion = New AD4BoardVersion(versionAD4Board0, versionAD4Board1)

	Finally
		deviceInterface.Close()
	End Try
End Function

Private Function ReadAD4BoardVersionString(versionString As String) As Double
	Dim versionAD4Board As Double
	If (Double.TryParse(versionString, versionAD4Board)) Then
		Return versionAD4Board
	End If

	Dim errorMessage = "The AD 4 board version " + versionString + " has not the correct format!" + vbCrLf + _
		"Please contact your nearest Polytec sales representative for assistance."
	Throw New exception(errorMessage)
End Function


Private Sub ResetRegistryKeys(ByRef analogChannels As List(Of AnalogChannelOptions))
	For Each ac As AnalogChannelOptions In analogChannels
		Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey(ac.RegistryKeyPath, True)
		If IsNothing(key) Then
			Continue For
		End If
		ac.OldOffset = Double.Parse(key.GetValue(registryKeyName, "0"), CultureInfo.InvariantCulture)
		key.SetValue(registryKeyName, "0")
		key.Close()
	Next
End Sub

Private Sub SetRegistryKeys(ByRef analogChannels As List(Of AnalogChannelOptions), newValues As Boolean)
	For Each ac As AnalogChannelOptions In analogChannels
		Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey(ac.RegistryKeyPath, True)
		If Not newValues Then
			If IsNothing(key) Then
				Continue For
			End If
			key.SetValue(registryKeyName, ac.OldOffset.ToString(CultureInfo.InvariantCulture))
		Else
			If IsNothing(key) Then
				key = Registry.LocalMachine.CreateSubKey(ac.RegistryKeyPath)
			End If
			If IsOffsetValid(ac) Then
				key.SetValue(registryKeyName, ac.CurrentOffset.ToString(CultureInfo.InvariantCulture))
			End If
		End If
		key.Close()
	Next
End Sub

Private Sub SetAllACOffsetsTo0()
	Dim showMessageBox As Boolean = False

	Dim rootKey As RegistryKey = Registry.LocalMachine.OpenSubKey(ethernetAcquisitionKey)
	If IsNothing(rootKey) Then
		Exit Sub
	End If

	For Each subKeyName As String In rootKey.GetSubKeyNames()
        Dim subKey As RegistryKey = rootKey.OpenSubKey(subKeyName)

		For Each subSubKeyName As String In subKey.GetSubKeyNames()
			If subSubKeyName.Contains("AC") Then
        		Dim subsubKey As RegistryKey = subKey.OpenSubKey(subSubKeyName, True)
        		If Not subsubKey.GetValue(registryKeyName, "0") = 0 Then
					subsubKey.SetValue(registryKeyName, "0")
					showMessageBox = True
        		End If
				subsubKey.Close()
			End If
       	Next
       	subKey.Close()
	Next
	rootKey.Close()

	If showMessageBox Then
		MsgBox("The offset values for AC-coupling have been set to 0.", vbInformation Or vbOkOnly)
	End If
End Sub
