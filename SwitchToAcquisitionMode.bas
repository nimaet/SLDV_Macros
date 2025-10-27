Attribute VB_Name = "SwitchToAcquisitionMode"
' POLYTEC CODE MODULE
' ----------------------------------------------------------------------
' PSV or MPV: This code module switches the application to the acquisition mode.
' VibSoft:	  This code module does nothing. It is only needed in order to use
'			  the PSV examples for VibSoft.
' This code module cannot be executed directly, it is used by other macros.
' ----------------------------------------------------------------------
'
' Return values:
' True: Application is in acquisition mode after execution
' False: Application is in presentation mode after execution
'
Public Function SwitchToAcquisitionMode() As Boolean

	If Application.Mode = ptcApplicationModeNormal Then
		' VibSoft is running, there is only one ApplicationMode
		SwitchToAcquisitionMode = True 
	Else
		' PSV or MPV is running
		If Application.Mode = ptcApplicationModePresentation Then

			Dim appNamespace As String
			appNamespace = GetApplicationNamespace(Application.Name)

			Dim oAcquisitionInstance As Object
			Set oAcquisitionInstance = CreateObject(appNamespace + ".AcquisitionInstance")

			If (oAcquisitionInstance.IsRunning) Then
				MsgBox "Please run this macro in the " + Application.Name + " acquisition instance."
				SwitchToAcquisitionMode = False
				Exit Function
			End If

			' Switch to Acquisition Mode
			Application.Mode = ptcApplicationModeAcquisition
		End If
	
		' Check if Application is switched to Acquisition Mode
		If Application.Mode = ptcApplicationModePresentation Then
			Beep
			MsgBox "Cannot switch to Acquisition Mode"
			SwitchToAcquisitionMode = False
		Else
			SwitchToAcquisitionMode = True 
		End If
	End If

End Function

Private Function GetApplicationNamespace(applicationName As String)

	Select Case applicationName
	Case "Polytec Vibrometer Software"
		GetApplicationNamespace = "VibSoft"
	Case "PSV"
		GetApplicationNamespace = applicationName
	Case "MPV"
		GetApplicationNamespace = applicationName
	Case Else
		MsgBox "Application " + Application.Name + " not supported."
		GetApplicationNamespace = ""
	End Select

End Function
