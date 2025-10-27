' Macro modified based on multiple Polytec Examples 
' Retained comments from the Stepped Fast Scan code
' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This macro is an example for a stepped fast scan measurement.
' you can set the start and end frequency, the frequency step,
' the bandwidth and the filename for the resulting combined file.
' Note that this allowed only equidistant frequency steps.
'
' When you start the macro a series of fast scans is done.
' The data of the fast scans are combined in a single file as user
' defined data sets. For every combination of domain/channel/signal/
' display in the original fast scan data you will find corresponding
' data in the user defined data set (channel 'Usr') of the resulting
' combined file.
' The data is stored at every measurement point as a function of the
' frequency.
'
' E.g., if you selected start = 1000 Hz, end = 2000 Hz, step = 100 Hz
' you will get 11 frequency lines at these frequencies. You can select
' to display e.g. FRF's by selecting the 'FFT Vib & Ref1 FRF Velocity /
' Voltage' signal in the 'Usr' channel. You will be able to set a cursor
' at a frequency line and to display the area data at this line for all
' scan points. The displayed area data corresponds to the area data of
' the original fast scan data.
'
' Only this combined fast scan file will be stored, the other files
' will be created only temporarily. This combined file will contain
' the original data of the first fast scan plus the combined data of
' all fast scans.
'
' To display the calculated data in PSV do the following:
' - start PSV and open the scan file
' - select Presentation/View/SingleScanPoint
' - in analyzer window toolbar select Channel/Usr
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
'#Language "WWB.NET"

Option Explicit

Imports System
Imports System.Collections.Generic

Const c_strFileFilter As String = "Scan File (*.svd)|*.svd|All Files (*.*)|*.*||"
Const c_strFileExt As String = "svd"

Const Samp As Double = 10000
Const SampFreq As Double = 781250
Const GenAmp As Double = 0.5
Const GenSteadyState As Double = 0.0

Dim CurrFreq As Integer

Dim hsFileName As String
Dim hsFilePath As String
Dim hsFreqStart As Double
Dim hsFreqEnd As Double
Dim hsFreqStep As Double
Dim Avgs As Integer


Sub Main
' -------------------------------------------------------------------------------
' Main procedure.
' -------------------------------------------------------------------------------
Debug.Clear

	'On Error GoTo MErrorHandler
	Dim bRe As Boolean

	If Not SwitchToAcquisitionMode() Then
		MsgBox("Switch to acquisition mode failed.", vbOkOnly)
		End
	End If
	'Call SetupGenerator(28000)

	' Dialog for user input.
	bRe = HarmonicScanRangeDialog()
	If Not bRe Then
		MsgBox("The macro exits now.", vbOkOnly)
		Exit Sub
	End If

	' Make an array for all filenames.
	Dim scanFileNames As New List(Of String)
	scanFileNames = HarmonicScanFiles()
	Dim scanFileCount As Integer
	scanFileCount = scanFileNames.Count

	' Now execute Harmonic Scan.
	bRe =  HarmonicScanMeasurement(scanFileNames)
	If Not bRe Then
		MsgBox("Error measurement! The macro exits now.", vbOkOnly)
		Exit Sub
	End If

	MsgBox("Acquisition Completed.", vbOkOnly)

End Sub


Function HarmonicScanMeasurement(ByRef scanFileNames As List(Of String)) As Boolean
' -------------------------------------------------------------------------------
'	Execute Harmonic Scan measurement.
' -------------------------------------------------------------------------------
	On Error GoTo MErrorHandler

	Dim dblFreqScan As Double
	Dim i As Integer

	i = 0
	For dblFreqScan = hsFreqStart To hsFreqEnd Step hsFreqStep


        'Acq.Mode = PTCAcqMode.ptcAcqModeTime

		'Dim ActiveProp As AcquisitionProperties
		'ActiveProp = Acq.ActiveProperties

		'Dim oBandpass As New DigitalFilterBandpass
		'oBandpass.CutoffFreq = 1000.0
		'oBandpass.Quality = ptcDigitalFilterQualityVeryHigh

		'Dim TimeScan As TimeAcqProperties

        'TimeScan = ActiveProp.TimeProperties
        'TimeScan.SampleFrequency = SampFreq
        'TimeScan.Samples = Samp

	    'Dim Avarage As	AverageAcqProperties
	    'Avarage = ActiveProp.AverageProperties
		'Avarage.Type = PTCAverageType.ptcAverageTime
		'Avarage.Count = Avgs

		Call SetupGenerator(dblFreqScan)



        Application.Acquisition.ScanFileName = scanFileNames(i)

        Application.Acquisition.Scan(PTCScanMode.ptcScanAll)

        While Application.Acquisition.State <> PTCAcqState.ptcAcqStateStopped
            Wait 0.01
        End While

		i = i + 1
	Next dblFreqScan

	HarmonicScanMeasurement = True

MErrorHandler:
	Select Case Err.Number
	Case 0
	Case Else
		MsgBox("Error:" & Err.Description, vbOkOnly)
		HarmonicScanMeasurement = False
		Resume MErrorHandler
	End Select
End Function

Private Sub SetupGenerator(UpFreq%)
	
	If Acquisition.ActiveProperties.Item(PTCAcqPropertiesType.ptcAcqPropertiesTypeGenerators).Count = 0 Then
		MsgBox "No Generator available." + vbCrLf + "Macro will be terminated."
		End
	End If


	Dim GeneratorProps As GeneratorAcqProperties
	GeneratorProps = Acquisition.ActiveProperties.Item(PTCAcqPropertiesType.ptcAcqPropertiesTypeGenerators)(1)

	If (GeneratorProps.Active = False) Then
		MsgBox "Generator not active. Generator will be activated."
		GeneratorProps.Active = True
		If (GeneratorProps.Active = False) Then
			MsgBox "Not possible to activate generator."+ vbCrLf + "Macro will be terminated."
			End
		End If
	End If

	' These settings will be used by all waveforms
	GeneratorProps.Offset = 0
	GeneratorProps.Amplitude = GenAmp
	GeneratorProps.SteadyStateTime = GenSteadyState
	Dim inputFile As String
	inputFile = "D:\Sai\0909 Transistor\ExpEnv_"+ CStr(UpFreq) +"_781250_ncyc20_T12800_RR10.txt"
	'inputFile = "D:\Sai\InputH\HannWaveForm_"+ CStr(UpFreq) +"_250k_100_1000.txt"
	Debug.Print inputFile
	' Sine
	'Generator will be switched to Sine
	Dim Sine As New WaveformUserDefined
	Sine.Frequency = 10	'Repetition Rate Hz
'	Sine.Load("D:\Sai\InputH\HannWaveForm_25000_250k_100_1000.txt")
	Sine.Load(inputFile)
	GeneratorProps.Waveform = Sine

	SendKeys "{F5}", True
	'Previous line to Open A/D settings dialog, click on the generator page to view the settings


End Sub

Private Function HarmonicScanFiles() As List(Of String)
' -------------------------------------------------------------------------------
' 	Make an array for base file and temporary files.
' -------------------------------------------------------------------------------
	Dim scanFileNames As New List(Of String)
	Dim dblFreqScan As Double
	Dim i As Integer

	hsFilePath = Left(hsFileName, InStrRev(hsFileName,".")-1)

	i = 0
	For dblFreqScan = hsFreqStart To hsFreqEnd Step hsFreqStep
		' Round to avoid numerical errors.
			scanFileNames.Add(hsFilePath + CStr(Round(dblFreqScan,4)) + ".svd")

		i = i + 1
	Next dblFreqScan
	Return scanFileNames
End Function

Const c_OFN_HIDEREADONLY As Long = 4
'Const c_DEFAULT_FILE As String = "D:\temp\SteppedFastScan\Test.svd"


Private Sub UpdateFreq(UpFreq%)
	' ----------
	' Update Generator to produce sine signal at specific frequency
	' ----------

	' Sine
	' Generator will be switched to Sine
    Dim GenProps As GeneratorAcqProperties
	GenProps = Acquisition.ActiveProperties.Item(PTCAcqPropertiesType.ptcAcqPropertiesTypeGenerators)(1)
	Dim inputFile As String

	inputFile = "D:\Sai\InputH\HannWaveForm_"+ CStr(UpFreq) +"_250k_100_1000.txt"
	Debug.Print inputFile

	Dim Sig As New WaveformUserDefined
	Debug.Print Sig._FileName

	Sig.Load("D:\Sai\InputH\HannWaveForm_25000_250k_100_1000.txt")	' Hz

	GenProps.Waveform = Sig

	'GeneratorProps.Amplitude = GenAmp
	' Turn Generator off
	Acquisition.GeneratorsOn = False

End Sub

Private Function HarmonicScanRangeDialog() As Boolean
' -------------------------------------------------------------------------------
'	Base dialog for user input.
' -------------------------------------------------------------------------------
	Begin Dialog UserDialog 480,245,"Stepped FastScans",.SetupDialogProc ' %GRID:10,7,1,1
		GroupBox 10,14,310,133,"Frequency",.GroupBoxFrequency
		Text 30,35,140,14,"Start Frequency [Hz]:",.TextStart
		TextBox 180,35,130,21,.fStart
		Text 30,63,140,14,"End Frequency [Hz]:",.TextEnd
		TextBox 180,63,130,21,.fEnd
		Text 30,91,140,14,"Step Frequency [Hz]:",.TextStop
		TextBox 180,91,130,21,.fStep
		Text 30,119,140,14,"No. of Averages:",.TextAverages
		TextBox 180,119,130,21,.tAvgs

		GroupBox 10,154,460,56,"Filename",.GroupBoxFilename
		TextBox 30,175,280,21,.sFile
		PushButton 350,175,100,21,"Browse...",.pbFile

		OKButton 250,217,100,21
		CancelButton 370,217,100,21
	End Dialog
    Dim dlg As UserDialog

   ' Show dialog (wait for any button pressed).
    Dim iRe As Integer
    iRe = Dialog(dlg)

	Select Case iRe
		Case 0
	    	' Cancel button pressed.
			HarmonicScanRangeDialog = False
		Case -1
	    	' OK button pressed.
			HarmonicScanRangeDialog = True
	End Select
End Function

Private Function SetupDialogProc(DlgItem$, Action%, SuppValue&) As Boolean
' -------------------------------------------------------------------------------
'	Second dialog for user input.
' -------------------------------------------------------------------------------
	Select Case Action%
	Case 1 ' Dialog box initialization.
		DlgText   "fStart"    , "20500"
		DlgText   "fEnd"    , "22900"
		DlgText   "fStep"    , "400"
DlgText   "tAvgs"    , "3"
DlgText   "sFile"    , "D:\Sai\0909 Transistor\ON.svd"
	Case 2 ' Value changing or button pressed.

		Select Case DlgItem
			Case "pbFile"
				' Get filename and path.
				DlgText "sFile", FileSaveDialog()

				' Prevent button press from closing the dialog box.
				SetupDialogProc = True

			Case "OK"
				If Not CheckDialogInput Then
					' Prevent button press from closing the dialog box.
					SetupDialogProc = True
				End If
		End Select

	Case 3 ' TextBox or ComboBox text changed.
	Case 4 ' Focus changed.
	Case 5 ' Idle.
	Case 6 ' Function key.
	End Select
End Function

Private Function CheckDialogInput() As Boolean
' -------------------------------------------------------------------------------
'	Check the user input.
' -------------------------------------------------------------------------------
	Dim strValue As String

	strValue = DlgText("fStart")
	If (strValue = "") Then
		MsgBox("Please insert the start frequency.", vbOkOnly)
		CheckDialogInput = False
		Exit Function
	End If

	strValue = DlgText("fEnd")
	If (strValue = "") Then
		MsgBox("Please insert the end frequency.", vbOkOnly)
		CheckDialogInput = False
		Exit Function
	End If

	strValue = DlgText("fStep")
	If (strValue = "") Then
		MsgBox("Please insert the frequency step.", vbOkOnly)
		CheckDialogInput = False
		Exit Function
	End If

	strValue = DlgText("sFile")
	If (strValue = "") Then
		MsgBox("Please insert the file name.", vbOkOnly)
		CheckDialogInput = False
		Exit Function
	End If

	strValue = DlgText("TextAverages")
	If (strValue = "") Then
		MsgBox("Please insert the averaging.", vbOkOnly)
		CheckDialogInput = False
		Exit Function
	End If

	' Set the variables to work with it.
	hsFreqStart = CDbl( DlgText("fStart"))
	hsFreqEnd   = CDbl( DlgText("fEnd"))
	hsFreqStep  = CDbl( DlgText("fStep"))
	hsFileName  = DlgText("sFile")
	Avgs = CDbl( DlgText("tAvgs"))

	If hsFreqStart >= hsFreqEnd Then
		MsgBox("Please insert a correct start and end frequency.", vbOkOnly)
		CheckDialogInput = False
		Exit Function
	End If

	If hsFreqStep > hsFreqEnd - hsFreqStart Then
		MsgBox("Please insert a correct start and end frequency and frequency step.", vbOkOnly)
		CheckDialogInput = False
		Exit Function
	End If

	If hsFreqStep <= 0 Then
		MsgBox("Please insert a correct frequency step (>0).", vbOkOnly)
		CheckDialogInput = False
		Exit Function
	End If

	If ((Avgs < 3) Or (Avgs > 1000000&)) Then
		MsgBox("The averaging should be between 3 and 1.000.000.", vbOkOnly)
		CheckDialogInput = False
		Exit Function
	End If

	CheckDialogInput = True
End Function

Private Function FileSaveDialog() As String
' -------------------------------------------------------------------------------
'	Base dialog to select a file for saving data.
' -------------------------------------------------------------------------------
	On Error GoTo MCreateError

	Dim fod As Object
	fod = CreateObject("MSComDlg.CommonDialog")
	fod.Filter = c_strFileFilter
	fod.Flags = c_OFN_HIDEREADONLY
	fod.CancelError = True
	On Error GoTo MCancelError
	fod.ShowSave
	FileSaveDialog = fod.FileName

	GoTo MEnd
MCancelError:
	FileSaveDialog = ""
	GoTo MEnd
MCreateError:
	FileSaveDialog = GetFilePath(, c_strFileExt, CurDir(), "Save as", 2)
MEnd:
End Function

Private Function SwitchToAcquisitionMode() As Boolean
' -------------------------------------------------------------------------------
'	Check and switch to AcquisitionMode, so we can make measurement.
' -------------------------------------------------------------------------------
	If Application.Mode = PTCApplicationMode.ptcApplicationModePresentation Then
		Dim oAcquisitionInstance As New AcquisitionInstance
		If (oAcquisitionInstance.IsRunning) Then
			MsgBox "Please run this macro in the PSV acquisition instance."
			SwitchToAcquisitionMode = False
			Exit Function
		End If
		Application.Mode = PTCApplicationMode.ptcApplicationModeAcquisition
	End If
	If Application.Mode = PTCApplicationMode.ptcApplicationModePresentation Then
		Beep
		MsgBox("Cannot switch to Acquisition Mode.", vbOkOnly)
		SwitchToAcquisitionMode = False
	Else
		SwitchToAcquisitionMode = True
	End If
End Function

