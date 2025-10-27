'#Language "WWB-COM"

Option Explicit

Sub Main
     Dim oAuxIn As DigitalPort
     Dim oAuxOut As DigitalPort

     Set oAuxIn = Application.DigitalPorts(ptcDigitalPortIn1)
     Set oAuxOut = Application.DigitalPorts(ptcDigitalPortOut1)

     Dim Count As Integer
     oAuxOut.Value = False
     For Count = 1 To 1000
		oAuxOut.Value = True
		Wait 0.01
		oAuxOut.Value = False
		Wait 0.09
     Next
End Sub
