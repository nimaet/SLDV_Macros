'#Reference {F08ACE20-C7AD-46CA-8001-D5158D9B0224}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolyMath.dll#Polytec PolyMath Type Library, $Revision: 4$ UnicodeRelease, Build on Jun 19 2006 at 21:13:02
'#Reference {E68EA160-8AD4-11D3-8F08-00104BB924B2}#1.0#0#C:\Program Files\Common Files\Polytec\COM\SignalDescription.dll#Polytec SignalDescription Type Library, $Revision: 4$ UnicodeRelease, Build on Jun 19 2006 at 21:12:47
'#Reference {CE68D434-5052-431F-BE75-F3C23458127A}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolySignal.dll#Polytec PolySignal Type Library, $Revision: 4$ UnicodeRelease, Build on Jun 19 2006 at 21:12:48
'#Reference {E44752C9-2D41-48A6-9B74-66D5B7505325}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolyProperties.dll#Polytec PolyProperties Type Library, $Revision: 4$ UnicodeRelease, Build on Jun 19 2006 at 21:12:54
'#Reference {A65C100F-C1FE-4C3B-9C43-46F4FB4C3BC3}#1.0#0#C:\Program Files\Common Files\Polytec\COM\PolyFile.dll#Polytec PolyFile Type Library, $Revision: 4$ UnicodeRelease, Build on Jun 19 2006 at 21:32:47
' POLYTEC MACRO DEMO
' ----------------------------------------------------------------------
' This Macro corrects a drift in time signals by calculating a
' linear regression of the time signal and subtracting this
' from the original data.
'
' The original data, the regression and the subtraction are
' copied to frame 1, 2 and 3 of user defined datasets.
'
' This is especially useful when using digital filters to
' integrate a signal. A small offset on the time signal
' causes a constant drift of the integrated signal. This
' can be corrected by this macro.
'
' When running the macro you are asked to navigate to a single point
' or scan files.
' This file has to meet the following conditions:
'
' - the file contains time domain data
' - you have to have exclusive write access to the file. We strongly recommend to
'   use a backup copy of your original file with this macro. The macro
'   will fail if the file is open in PSV or VibSoft.
'
' To display the calculated data in PSV/VibSoft do the following:
' - start PSV/VibSoft and open the file
' - select Analyzer/Channel/Usr
' - select one of the signals offered in Analyzer/Signal
'
' References
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
Option Explicit

Const c_strFileFilter As String = "Scan File (*.svd)|*.svd|Single Point File (*.pvd)|*.pvd|All Files (*.*)|*.*||"
Const c_strFileExt As String = "svd;pvd"

Sub Main
' -------------------------------------------------------------------------------
'	Main procedure.
' -------------------------------------------------------------------------------
	Dim oFile As New PolyFile

	' get filename and path
	Dim strFileName As String
	strFileName = FileOpenDialog()

	If (strFileName = "") Then
		MsgBox("No filename has been specified, macro exits now.", vbOkOnly)
		Exit Sub
	End If

	If (MsgBox("This macro will modify the file '" + strFileName + _
		"'. We strongly recommend to work with a backup copy of original data only. " + _
		"Do you want to continue?", vbYesNo) = vbNo) Then
		Exit Sub
	End If

	' we have to open the file for read/write, otherwise we cannot save our
	' user defined dataset to the file
	If Not OpenFile(oFile, strFileName) Then
		Exit Sub
	End If

	Dim isScanFile As Boolean
	isScanFile = oFile.Infos.HasMeasPoints

	Dim oAcqProps As AcquisitionPropertiesContainer
    Set oAcqProps = oFile.Infos.AcquisitionInfoModes.ActiveProperties

    Dim oChannelsAcqProps As ChannelsAcqPropertiesContainer
    Set oChannelsAcqProps = oAcqProps.ChannelsProperties

	Dim oPointDomains As PointDomains
	Set oPointDomains = oFile.GetPointDomains(ptcBuildPointData3d)

	Dim oPointDomain As PointDomain
	Set oPointDomain = oPointDomains.type(ptcDomainTime)

	Dim oDomain As Domain
	Set oDomain = oPointDomain

	Dim Data() As Single
	Dim Regression() As Single

	Dim oVector As New Vector

	Dim oChannel As Channel

	' loop over all channels and signals of the time domain
	For Each oChannel In oDomain.Channels
		' ignore user defined channels
		If (oChannel.Name <> "Usr") Then
			Dim oSignal As Signal
			For Each oSignal In oChannel.Signals

				Dim oDisplay As Display
				Set oDisplay = oSignal.Displays.type(ptcDisplaySamples)

				Dim oUsrSigDesc As SignalDescription
				Set oUsrSigDesc = oSignal.Description.Clone()

				' fill the properties of the user signal description
				' for the axes we copy the properties of the original
				' data
				With oUsrSigDesc
					.Name = "No Drift: " + oChannel.Name + " " + oSignal.Name
				End With

				Dim oUsrSignal As Signal
				Set oUsrSignal = oPointDomains.FindSignal(oUsrSigDesc, True)

				' check if a signal with the same name exits already
				If (oUsrSignal Is Nothing) Then
					Set oUsrSignal = oPointDomains.AddSignal(oUsrSigDesc)
				Else
					If (MsgBox("A user defined signal with the name '" + oUsrSigDesc.Name + "' exists already. Do you want to replace it?", vbYesNo) = vbYes) Then
						oUsrSignal.Channel.Signals.Update(oUsrSignal.Name, oUsrSigDesc)
					Else
						Set oUsrSignal = Nothing
					End If
				End If

				If (Not oUsrSignal Is Nothing) Then
					Dim oDataPoint As DataPoint
					For Each oDataPoint In oPointDomain.DataPoints
						Dim isValid As Boolean
						isValid = True
						If isScanFile Then
							Dim oMeasPoint As MeasPoint
							Set oMeasPoint = oDataPoint.MeasPoint
							isValid = (oMeasPoint.ScanStatus And ptcScanStatusValid) <> 0 Or (oMeasPoint.ScanStatus And ptcScanStatusInvalidated) <> 0
						End If

						If isValid Then
							Dim is3D As Boolean
							is3D = False

							Dim oDOF As DegreeOfFreedomIDs
							Set oDOF = oSignal.Description.ResponseDOFs
							If (oDOF.Count > 0) Then
								If (oDOF.Direction = ptcVector) Then
									is3D = True
								End If
							End If

							Data = oDataPoint.GetData(oDisplay, 0)
							Call CalcRegression(1, Data, Regression, is3D)

							' set the data
							oDataPoint.SetData(oUsrSignal, 1, Data())
							oDataPoint.SetData(oUsrSignal, 2, Regression())
							oDataPoint.SetData(oUsrSignal, 3, oVector.Sub(Data, Regression))
						End If
					Next oDataPoint
				End If
			Next oSignal
		End If
	Next oChannel

	oFile.Save()
	oFile.Close()

    MsgBox("Macro has finished.", vbOkOnly)
End Sub

Function GetChannelAcqPropsByName(oChannelsAcqProps As ChannelsAcqPropertiesContainer, strName As String) As ChannelAcqPropertiesContainer
' -------------------------------------------------------------------------------
' gets the channel acquisition properties by the short name of the channel
' we have to use the name and not the SourceChannel value for this, because 3D channels have a single entry
' in the ChannelsAcqProperties collection but occupy three source channel numbers
' -------------------------------------------------------------------------------
    Dim oChannelAcqProps As ChannelAcqPropertiesContainer
    For Each oChannelAcqProps In oChannelsAcqProps
        If strName = oChannelAcqProps.ShortName Then
            Set GetChannelAcqPropsByName = oChannelAcqProps
            Exit Function
        End If
    Next
    Err.Raise(1, "GetChannelAcqPropsByName", "Could not find a channel with the given name")
End Function

Sub CalcRegression(iDegree As Integer, Data() As Single, Regression() As Single, is3D As Boolean)
' -------------------------------------------------------------------------------
' Calculates a regression of degree iDegree.
' -------------------------------------------------------------------------------
	ReDim Regression(LBound(Data) To UBound(Data))

	Dim dataSets As Integer
	dataSets = 1
	If (is3D = True) Then
		dataSets = 3
	End If

	Dim Count As Long
	Count = (UBound(Data,1) - LBound(Data,1) + 1 ) / dataSets

	Dim lb As Long
	lb = LBound(Data,1)

	Dim ub As Long
	ub = Count - lb -1

	Dim xi() As Double
	Dim y() As Double
	ReDim xi(0 To iDegree) As Double
	ReDim y(0 To iDegree) As Double
	Dim x As Long
	Dim i As Integer
	Dim xc As Double
	Dim z As Double

	Dim M As PolyMathLib.Matrix
	Set M = New PolyMathLib.Matrix
	
	xc = (lb + ub) / 2

	Dim dataSet As Integer
	For dataSet = 0 To (dataSets - 1)

		M.Init( iDegree+1, iDegree+1)
		ReDim y(0 To iDegree)

		For x = lb To ub
			xi(0) = 1.0

			For i = 1 To iDegree
				xi(i) = xi(i-1)*(x-xc)
			Next i
			M.AddOuterProduct( xi)
			For i = 0 To iDegree
				y(i) = y(i) + xi(i)*Data(x)
			Next i
		Next x

		' solve the normal equations
		' the resulting vector is 1-based
		y = M.Solve( y)

		' calculate the regression
		For x = lb To ub
			' at position x ...
			xi(0) = 1.0
			For i = 1 To iDegree
				xi(i) = xi(i-1)*(x-xc)
			Next i
			z = 0.0
			For i = 0 To iDegree
				z = z + xi(i)*y(i+1)
			Next i

			Regression(x) = z
		Next x

		ub = ub + Count
		lb = lb + Count
	Next dataSet
End Sub


' *******************************************************************************
' * Helper functions and subroutines
' *******************************************************************************

Const c_OFN_HIDEREADONLY As Long = 4

Private Function FileOpenDialog() As String
' -------------------------------------------------------------------------------
' Select file.
' -------------------------------------------------------------------------------
	On Error GoTo MCreateError
	Dim fod As Object
	Set fod = CreateObject("MSComDlg.CommonDialog")
	fod.Filter = c_strFileFilter
	fod.Flags = c_OFN_HIDEREADONLY
	fod.CancelError = True
	On Error GoTo MCancelError
	fod.ShowOpen
	FileOpenDialog = fod.FileName
	GoTo MEnd
MCancelError:
	FileOpenDialog = ""
	GoTo MEnd
MCreateError:
	FileOpenDialog = GetFilePath(, c_strFileExt, CurDir(), "Select a file", 2)
MEnd:
End Function

Private Function OpenFile(oFile  As PolyFile, strFileName As String) As Boolean
' -------------------------------------------------------------------------------
' Instantiate PolyFile object, open the File.
' -------------------------------------------------------------------------------
	On Error GoTo MErrorHandler
	Dim bRe As Boolean
	bRe = True

	Set oFile = New PolyFile
	If oFile.ReadOnly Then
		oFile.ReadOnly = False
	End If

	On Error Resume Next
	oFile.Open (strFileName)

	On Error GoTo 0
	If Not oFile.IsOpen Then
		MsgBox("Can not open the file."& vbCrLf & _
		"Check the file attribute is not read only!", vbExclamation)
		bRe = False
	End If
MErrorHandler:
	Select Case Err
	Case 0
		OpenFile = bRe
	Case Else
		bRe = False
		Resume MErrorHandler
	End Select
End Function
