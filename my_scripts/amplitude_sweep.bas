
Sub RepeatScanAmplitude()
	Dim ampList As Variant
	Dim amp As Double
	Dim i As Integer
	Dim basePath As String
	Dim measName As String
	
	' === USER SETTINGS ===
	basePath = "D:\Nima\Metamaterial beam\dat_Oct\27\"  ' save location
	ampList = Array(0.05, 0.1, 0.15, 0.2, 0.25, 0.3, 0.35, 0.4)  ' excitation amplitudes (V)
	
	' === PREPARE ===
	App.ActiveMeasurement.Stop
	App.ActiveMeasurement.DeleteAll
	App.ActiveMeasurement.New
	
	App.ActiveMeasurement.SwitchToAcquisitionMode
	
	' === MAIN LOOP ===
	For i = LBound(ampList) To UBound(ampList)
		amp = ampList(i)
		
		' --- Set excitation amplitude ---
		App.ActiveMeasurement.Generator.Amplitude = amp
		
		' --- Start scan ---
		App.ActiveMeasurement.Start
		Do While App.ActiveMeasurement.IsRunning
			DoEvents
		Loop
		
		' --- Save data ---
		measName = "Scan_Amp" & Replace(Format(amp, "0.00"), ".", "_") & "V.vib"
		App.ActiveMeasurement.SaveAs basePath & measName
		
		MsgBox "Scan completed and saved as: " & measName
	Next i
	
	MsgBox "All scans finished successfully."
End Sub
