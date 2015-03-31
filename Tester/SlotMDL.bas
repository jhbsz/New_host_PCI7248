Attribute VB_Name = "SlotMDL"
Public Function NBModeTestFcn(ChipString As String) As Byte

Tester.Print "NB mode test ----"

  If GetDeviceName(ChipString) = "" Then
    NBModeTestFcn = 1
    Tester.Print NBModeTestFcn; " :PASS"
  Else
    NBModeTestFcn = 0
    Tester.Print NBModeTestFcn; " :FAIL"
  End If

End Function
Public Function NormalModeTestFcn(ChipString As String) As Byte

Tester.Print "Normal mode test ----"

  If GetDeviceName(ChipString) <> "" Then
    NormalModeTestFcn = 1
    Tester.Print NormalModeTestFcn; " :PASS"
  Else
     NormalModeTestFcn = 0
    Tester.Print NormalModeTestFcn; " :FAIL"
  End If
 
End Function
Public Function InitailChip()

End Function
Public Function SpeedTestFcn()

End Function
Public Function SDSlotTestFcn(SDSlot As Byte, PreviousStatus As Byte) As Byte

Tester.Print "SD Slot test ----"

If PreviousStatus <> 1 Then
  SDSlotTestFcn = 4
  Tester.Print SDSlotTestFcn; " :Previous Slot FAIL"
  Exit Function
End If
  


Tester.Print
End Function
Public Function CFSlotTestFcn(CFSlot As Byte, PreviousStatus As Byte) As Byte

If PreviousStatus <> 1 Then
  CFSlotTestFcn = 4
  Exit Function
End If

End Function
Public Function XDSlotTestFcn(XDSlot As Byte, previosStatus As Byte) As Byte

  If PreviousStatus <> 1 Then
   XDSlotTestFcn = 4
   Exit Function
End If

End Function
Public Function SMCSlotTestFcn(SMCSlot As Byte, previosStatus As Byte) As Byte

     If PreviousStatus <> 1 Then
        SMCSlotTestFcn = 4
        Exit Function
    End If

End Function
Public Function MSProSlotTestFcn(MSproSlot As Byte) As Byte

End Function
Public Function XDNoCISSlotTestFcn(XDNoCISSlot As Byte) As Byte

End Function
Public Function LightTestFcn(StdValue As Byte) As Byte

End Function
Public Function ControlSwitch(SwitchValue As String) As Byte

 

 
     If Right(SwitchValue, 1) = 1 Then
         ControlSwitch = 1
     End If
     
    If Mid(SwitchValue, 7, 1) = 1 Then
         ControlSwitch = ControlSwitch + 2
     End If
   
     If Mid(SwitchValue, 6, 1) = 1 Then
         ControlSwitch = ControlSwitch + 2 * 2
     End If
     
     If Mid(SwitchValue, 5, 1) = 1 Then
         ControlSwitch = ControlSwitch + 2 * 2 * 2
     End If
 
     If Mid(SwitchValue, 4, 1) = 1 Then
         ControlSwitch = ControlSwitch + 2 * 2 * 2 * 2
     End If

     If Mid(SwitchValue, 3, 1) = 1 Then
         ControlSwitch = ControlSwitch + 2 * 2 * 2 * 2 * 2
     End If
     
     
      If Mid(SwitchValue, 2, 1) = 1 Then
         ControlSwitch = ControlSwitch + 2 * 2 * 2 * 2 * 2 * 2
     End If
     
     If Left(SwitchValue, 1) = 1 Then
         ControlSwitch = ControlSwitch + 2 * 2 * 2 * 2 * 2 * 2 * 2
     End If


End Function

