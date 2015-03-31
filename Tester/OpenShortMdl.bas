Attribute VB_Name = "OpenShortMdl"
Option Explicit
Global OSValue(0 To 127) As Long
Global OSStandard(0 To 127) As Long
Global OSExtraDelay(0 To 127) As Integer
Global OSFlag As Byte
Global OpenShortPinNo As Byte
Global OldOSStandFileName As String
Public Sub OpenShortTest()
'Dim OSValue As Byte
'Dim OSValue(0 To 127) As Long
Dim result

Dim i As Long
Dim tmp As Long
Dim tmp2 As Byte
Dim T1
Dim DAQTime As Integer
Dim TempLogString As String

TempLogString = ""

' setting
If Tester.OSCheck.Value = 1 Then
   DAQTime = 500  ' to record standard
Else
   DAQTime = 2
End If



If OSFlag = 0 Or (OldOSStandFileName <> ChipName) Then

  Open App.Path & "/OSStandard/" & ChipName & ".txt" For Input As #6
  Input #6, OpenShortPinNo
  For i = 0 To OpenShortPinNo
    Input #6, tmp, OSStandard(i)
    If (tmp > 128) Then
        OSExtraDelay(i) = tmp
    Else
        OSExtraDelay(i) = 0
    End If
    
    
  Next i
  Close #6
  
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
        'replace by new O/S module
        Call SetTimer_1ms
    End If
  OSFlag = 1   ' inital ok
  OldOSStandFileName = ChipName
  
  CurrentOSFileName = "D:\OSFail_Log\" & ChipName & "_OSdatalog_" & Year(Date) & Format(Month(Date), "0#") _
                      & Format(Day(Date), "0#") & ".txt"
  SaveOSCounter = 1
  
End If

If Tester.SaveOSLog.Value = 1 Then
    Open CurrentOSFileName For Append As #10
End If
    
 
'Dim OSString As Stringt
 
    T1 = Timer
    For i = 0 To OpenShortPinNo
      CardResult = DO_WritePort(card, Channel_P1A, i)  ' 1111 1110
      Call Timer_1ms(DAQTime)
      If OSExtraDelay(i) <> 0 Then
        Call Timer_1ms(CInt(OSExtraDelay(i)))
      End If
      'Call Timer_1ms(4)
      CardResult = DO_ReadPort(card, Channel_P1B, OSValue(i))
    Next i
   
   Dim T2
     T2 = Timer
  
     Tester.Print "Time cycle:"; T2 - T1
     
If Tester.OSCheck.Value = 1 Then
    
    OpenShortFrm.Show
    OpenShortFrm.Cls
  
      For i = 0 To OpenShortPinNo
           
           If OSValue(i) = &HFC Then
               OpenShortFrm.Print CStr((i + 1)) & " is open *****"
           ElseIf OSValue(i) = &HFF Then
               OpenShortFrm.Print CStr((i + 1)) & " is Short ++++"
           ElseIf OSValue(i) <> &HFD Then
               OpenShortFrm.Print (i + 1); Hex(OSValue(i))
           End If
      Next
     OpenShortFrm.Print "hint: save Rec"
  Exit Sub
End If
  
  
TestResult = "PASS"

  For i = 0 To OpenShortPinNo
    ' OSValue(i) = CAndValue(OSValue(i), &HC0)
      If ((OSValue(i) <> OSStandard(i)) And (OSStandard(i) <> 0)) Then
      
         If OSValue(i) = &HFF Then
            Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         ElseIf OSValue(i) = &HFC Then
            Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         Else
            Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         End If
         
         If Tester.SaveOSLog.Value = 1 Then
            
            If OSValue(i) = &HFF Then
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Short" & vbCrLf
                'Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            ElseIf OSValue(i) = &HFC Then
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Open" & vbCrLf
                'Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            Else
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Normal" & vbCrLf
                'Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            End If
            
         End If
         
         TestResult = "Fail"
        
      End If
  Next
    
    If TestResult = "PASS" Then
        Call LabelMenu(0, 1, 0)
    Else
        Call LabelMenu(0, 2, 0)
    End If
    
    
    If (TestResult = "Fail") And (Tester.SaveOSLog.Value = 1) Then
        Print #10, "====================" & vbCrLf _
                    & "#" & SaveOSCounter & vbCrLf & "===================="
        Print #10, TempLogString & vbCrLf
        SaveOSCounter = SaveOSCounter + 1
    End If
    
    If Tester.SaveOSLog.Value = 1 Then
        Close #10
    End If
    
   Tester.Print "OpenShort :"; TestResult

End Sub

Public Sub OpenShortTest_AssignOSFileName(ThisOSName As String)
'Dim OSValue As Byte
'Dim OSValue(0 To 127) As Long
Dim result

Dim i As Long
Dim tmp As Byte
Dim tmp2 As Byte
Dim T1
Dim DAQTime As Integer
Dim TempLogString As String


TempLogString = ""

' setting
If Tester.OSCheck.Value = 1 Then
   DAQTime = 500  ' to record standard
Else
   DAQTime = 4
End If



If OSFlag = 0 Or (OldOSStandFileName <> ChipName) Then

  Open App.Path & "/OSStandard/" & ThisOSName & ".txt" For Input As #6
  Input #6, OpenShortPinNo
  For i = 0 To OpenShortPinNo
        Input #6, tmp, OSStandard(i)
  Next i
  Close #6
  
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
        'Call PowerSet2(1, "3.2", "0.5", 1, "0.2", "0.5", 1)
        'replace by new O/S module
        Call SetTimer_1ms
    End If
  OSFlag = 1   ' inital ok
  OldOSStandFileName = ChipName
  
  CurrentOSFileName = "D:\OSFail_Log\" & ChipName & "_OSdatalog_" & Year(Date) & Format(Month(Date), "0#") _
                      & Format(Day(Date), "0#") & ".txt"
  SaveOSCounter = 1
  
End If

If Tester.SaveOSLog.Value = 1 Then
    Open CurrentOSFileName For Append As #10
End If
    
 
'Dim OSString As Stringt
 
    T1 = Timer
    For i = 0 To OpenShortPinNo
      CardResult = DO_WritePort(card, Channel_P1A, i)  ' 1111 1110
      Call Timer_1ms(DAQTime)
      'Call Timer_1ms(4)
      CardResult = DO_ReadPort(card, Channel_P1B, OSValue(i))
      
    Next i
   
   Dim T2
     T2 = Timer
  
     Tester.Print "Time cycle:"; T2 - T1
     
If Tester.OSCheck.Value = 1 Then
    
    OpenShortFrm.Show
    OpenShortFrm.Cls
  
      For i = 0 To OpenShortPinNo
           
           If OSValue(i) = &HFC Then
               OpenShortFrm.Print CStr((i + 1)) & " is open *****"
           ElseIf OSValue(i) = &HFF Then
               OpenShortFrm.Print CStr((i + 1)) & " is Short ++++"
           ElseIf OSValue(i) <> &HFD Then
               OpenShortFrm.Print (i + 1); Hex(OSValue(i))
           End If
      Next
     OpenShortFrm.Print "hint: save Rec"
  Exit Sub
End If
  
  
TestResult = "PASS"

  For i = 0 To OpenShortPinNo
    ' OSValue(i) = CAndValue(OSValue(i), &HC0)
      If (OSValue(i) <> OSStandard(i)) Then
      
         If OSValue(i) = &HFF Then
            Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         ElseIf OSValue(i) = &HFC Then
            Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         Else
            Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         End If
         
         If Tester.SaveOSLog.Value = 1 Then
            
            If OSValue(i) = &HFF Then
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Short" & vbCrLf
                'Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            ElseIf OSValue(i) = &HFC Then
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Open" & vbCrLf
                'Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            Else
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Normal" & vbCrLf
                'Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            End If
            
         End If
         
         TestResult = "Fail"
        
      End If
  Next
    
    If TestResult = "PASS" Then
        Call LabelMenu(0, 1, 0)
    Else
        Call LabelMenu(0, 2, 0)
    End If
    
    
    If (TestResult = "Fail") And (Tester.SaveOSLog.Value = 1) Then
        Print #10, "====================" & vbCrLf _
                    & "#" & SaveOSCounter & vbCrLf & "===================="
        Print #10, TempLogString & vbCrLf
        SaveOSCounter = SaveOSCounter + 1
    End If
    
    If Tester.SaveOSLog.Value = 1 Then
        Close #10
    End If
    
   Tester.Print "OpenShort :"; TestResult

End Sub

Public Sub OpenShortTest_SkipZero()
'Dim OSValue As Byte
'Dim OSValue(0 To 127) As Long
Dim result

Dim i As Long
Dim tmp As Byte
Dim tmp2 As Byte
Dim T1
Dim DAQTime As Integer
Dim TempLogString As String


TempLogString = ""

' setting
If Tester.OSCheck.Value = 1 Then
   DAQTime = 500  ' to record standard
Else
   DAQTime = 4
End If



If OSFlag = 0 Or (OldOSStandFileName <> ChipName) Then

  Open App.Path & "/OSStandard/" & ChipName & ".txt" For Input As #6
  Input #6, OpenShortPinNo
  For i = 0 To OpenShortPinNo
        Input #6, tmp, OSStandard(i)
  Next i
  Close #6
  
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
        'Call PowerSet2(1, "3.2", "0.5", 1, "0.2", "0.5", 1)
        'replace by new O/S module
        Call SetTimer_1ms
    End If
  OSFlag = 1   ' inital ok
  OldOSStandFileName = ChipName
  
  CurrentOSFileName = "D:\OSFail_Log\" & ChipName & "_OSdatalog_" & Year(Date) & Format(Month(Date), "0#") _
                      & Format(Day(Date), "0#") & ".txt"
  SaveOSCounter = 1
  
End If

If Tester.SaveOSLog.Value = 1 Then
    Open CurrentOSFileName For Append As #10
End If
    
 
'Dim OSString As Stringt
 
    T1 = Timer
    For i = 0 To OpenShortPinNo
      CardResult = DO_WritePort(card, Channel_P1A, i)  ' 1111 1110
      Call Timer_1ms(DAQTime)
      'Call Timer_1ms(4)
      CardResult = DO_ReadPort(card, Channel_P1B, OSValue(i))
      
    Next i
   
   Dim T2
     T2 = Timer
  
     Tester.Print "Time cycle:"; T2 - T1
     
If Tester.OSCheck.Value = 1 Then
    
    OpenShortFrm.Show
    OpenShortFrm.Cls
  
      For i = 0 To OpenShortPinNo
           
           If OSValue(i) = &HFC Then
               OpenShortFrm.Print CStr((i + 1)) & " is open *****"
           ElseIf OSValue(i) = &HFF Then
               OpenShortFrm.Print CStr((i + 1)) & " is Short ++++"
           ElseIf OSValue(i) <> &HFD Then
               OpenShortFrm.Print (i + 1); Hex(OSValue(i))
           End If
      Next
     OpenShortFrm.Print "hint: save Rec"
  Exit Sub
End If
  
  
TestResult = "PASS"

  For i = 0 To OpenShortPinNo
    ' OSValue(i) = CAndValue(OSValue(i), &HC0)
      If ((OSValue(i) <> OSStandard(i)) And (OSStandard(i) <> 0)) Then
      
         If OSValue(i) = &HFF Then
            Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         ElseIf OSValue(i) = &HFC Then
            Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         Else
            Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         End If
         
         If Tester.SaveOSLog.Value = 1 Then
            
            If OSValue(i) = &HFF Then
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Short" & vbCrLf
                'Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            ElseIf OSValue(i) = &HFC Then
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Open" & vbCrLf
                'Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            Else
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Normal" & vbCrLf
                'Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            End If
            
         End If
         
         TestResult = "Fail"
        
      End If
  Next
    
    If TestResult = "PASS" Then
        OS_Result = 1
        Call LabelMenu(0, 1, 0)
    Else
        OS_Result = 0
        Call LabelMenu(0, 2, 0)
    End If
    
    
    If (TestResult = "Fail") And (Tester.SaveOSLog.Value = 1) Then
        Print #10, "====================" & vbCrLf _
                    & "#" & SaveOSCounter & vbCrLf & "===================="
        Print #10, TempLogString & vbCrLf
        SaveOSCounter = SaveOSCounter + 1
    End If
    
    If Tester.SaveOSLog.Value = 1 Then
        Close #10
    End If
    
   Tester.Print "OpenShort :"; TestResult

End Sub

Public Sub OpenShortTest_SkipZero_NoGPIB()
'Dim OSValue As Byte
'Dim OSValue(0 To 127) As Long
Dim result

Dim i As Long
Dim tmp As Long
Dim tmp2 As Byte
Dim T1
Dim DAQTime As Integer
Dim TempLogString As String

TempLogString = ""

' setting
If Tester.OSCheck.Value = 1 Then
   DAQTime = 500  ' to record standard
Else
   DAQTime = 1
End If



If OSFlag = 0 Or (OldOSStandFileName <> ChipName) Then

  Open App.Path & "/OSStandard/" & ChipName & ".txt" For Input As #6
  Input #6, OpenShortPinNo
'  For i = 0 To OpenShortPinNo
'        Input #6, tmp, OSStandard(i)
'  Next i
  
    For i = 0 To OpenShortPinNo
        Input #6, tmp, OSStandard(i)
        If (tmp > 128) Then
            OSExtraDelay(i) = tmp
        Else
            OSExtraDelay(i) = 0
        End If
    Next i
    Close #6
  
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
        Call SetTimer_1ms
       
    End If
    
  OSFlag = 1   ' inital ok
  OldOSStandFileName = ChipName
  
  CurrentOSFileName = "D:\OSFail_Log\" & ChipName & "_OSdatalog_" & Year(Date) & Format(Month(Date), "0#") _
                      & Format(Day(Date), "0#") & ".txt"
  SaveOSCounter = 1
  
End If

If Tester.SaveOSLog.Value = 1 Then
    Open CurrentOSFileName For Append As #10
End If
    
 
'Dim OSString As Stringt
 
    T1 = Timer
    For i = 0 To OpenShortPinNo
      CardResult = DO_WritePort(card, Channel_P1A, i)  ' 1111 1110
      Call Timer_1ms(DAQTime)
      If OSExtraDelay(i) <> 0 Then
        Call Timer_1ms(CInt(OSExtraDelay(i)))
      End If
      'Call Timer_1ms(4)
      CardResult = DO_ReadPort(card, Channel_P1B, OSValue(i))
      
    Next i
   
   Dim T2
     T2 = Timer
  
     Tester.Print "Time cycle:"; T2 - T1
     
If Tester.OSCheck.Value = 1 Then
    
    OpenShortFrm.Show
    OpenShortFrm.Cls
  
      For i = 0 To OpenShortPinNo
           
           If OSValue(i) = &HFC Then
               OpenShortFrm.Print CStr((i + 1)) & " is open *****"
           ElseIf OSValue(i) = &HFF Then
               OpenShortFrm.Print CStr((i + 1)) & " is Short ++++"
           ElseIf OSValue(i) <> &HFD Then
               OpenShortFrm.Print (i + 1); Hex(OSValue(i))
           End If
      Next
     OpenShortFrm.Print "hint: save Rec"
  Exit Sub
End If
  
  
TestResult = "PASS"

  For i = 0 To OpenShortPinNo
    ' OSValue(i) = CAndValue(OSValue(i), &HC0)
      If ((OSValue(i) <> OSStandard(i)) And (OSStandard(i) <> 0)) Then
      
         If OSValue(i) = &HFF Then
            Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         ElseIf OSValue(i) = &HFC Then
            Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         Else
            Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         End If
         
         If Tester.SaveOSLog.Value = 1 Then
            
            If OSValue(i) = &HFF Then
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Short" & vbCrLf
                'Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            ElseIf OSValue(i) = &HFC Then
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Open" & vbCrLf
                'Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            Else
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Normal" & vbCrLf
                'Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            End If
            
         End If
         
         TestResult = "Fail"
        
      End If
  Next
    
    If TestResult = "PASS" Then
        OS_Result = 1
        Call LabelMenu(0, 1, 0)
    Else
        OS_Result = 0
        Call LabelMenu(0, 2, 0)
    End If
    
    
    If (TestResult = "Fail") And (Tester.SaveOSLog.Value = 1) Then
        Print #10, "====================" & vbCrLf _
                    & "#" & SaveOSCounter & vbCrLf & "===================="
        Print #10, TempLogString & vbCrLf
        SaveOSCounter = SaveOSCounter + 1
    End If
    
    If Tester.SaveOSLog.Value = 1 Then
        Close #10
    End If
    
   Tester.Print "OpenShort :"; TestResult

End Sub

Public Sub OpenShortTest_Pin2ShortBin3()
'Dim OSValue As Byte
'Dim OSValue(0 To 127) As Long
Dim result

Dim i As Long
Dim tmp As Byte
Dim tmp2 As Byte
Dim T1
Dim DAQTime As Integer
Dim TempLogString As String


TempLogString = ""

' setting
If Tester.OSCheck.Value = 1 Then
   DAQTime = 500  ' to record standard
Else
   DAQTime = 4
End If



If OSFlag = 0 Or (OldOSStandFileName <> ChipName) Then

  Open App.Path & "/OSStandard/" & ChipName & ".txt" For Input As #6
  Input #6, OpenShortPinNo
  For i = 0 To OpenShortPinNo
        Input #6, tmp, OSStandard(i)
  Next i
  Close #6
  
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
        'Call PowerSet2(1, "3.2", "0.5", 1, "0.2", "0.5", 1)
        'replace by new O/S module
        Call SetTimer_1ms
    End If
  OSFlag = 1   ' inital ok
  OldOSStandFileName = ChipName
  
  CurrentOSFileName = "D:\OSFail_Log\" & ChipName & "_OSdatalog_" & Year(Date) & Format(Month(Date), "0#") _
                      & Format(Day(Date), "0#") & ".txt"
  SaveOSCounter = 1
  
End If

If Tester.SaveOSLog.Value = 1 Then
    Open CurrentOSFileName For Append As #10
End If
    
 
'Dim OSString As Stringt
 
    T1 = Timer
    For i = 0 To OpenShortPinNo
      CardResult = DO_WritePort(card, Channel_P1A, i)  ' 1111 1110
      Call Timer_1ms(DAQTime)
      'Call Timer_1ms(4)
      CardResult = DO_ReadPort(card, Channel_P1B, OSValue(i))
      
    Next i
   
   Dim T2
     T2 = Timer
  
     Tester.Print "Time cycle:"; T2 - T1
     
If Tester.OSCheck.Value = 1 Then
    
    OpenShortFrm.Show
    OpenShortFrm.Cls
  
      For i = 0 To OpenShortPinNo
           
           If OSValue(i) = &HFC Then
               OpenShortFrm.Print CStr((i + 1)) & " is open *****"
           ElseIf OSValue(i) = &HFF Then
               OpenShortFrm.Print CStr((i + 1)) & " is Short ++++"
           ElseIf OSValue(i) <> &HFD Then
               OpenShortFrm.Print (i + 1); Hex(OSValue(i))
           End If
      Next
     OpenShortFrm.Print "hint: save Rec"
  Exit Sub
End If
  
  
TestResult = "PASS"

  For i = 0 To OpenShortPinNo
    ' OSValue(i) = CAndValue(OSValue(i), &HC0)
      If (OSValue(i) <> OSStandard(i)) Then
      
         If OSValue(i) = &HFF Then
            Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         ElseIf OSValue(i) = &HFC Then
            Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         Else
            Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         End If
         
         If (i = 1) And (OSValue(i) = &HFF) Then
            TestResult = "Bin3"
            Exit For
         End If
         
         If Tester.SaveOSLog.Value = 1 Then
            
            If OSValue(i) = &HFF Then
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Short" & vbCrLf
                'Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            ElseIf OSValue(i) = &HFC Then
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Open" & vbCrLf
                'Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            Else
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Normal" & vbCrLf
                'Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            End If
            
         End If
         
         TestResult = "Bin2"
        
      End If
  Next
    
    If TestResult = "PASS" Then
        Call LabelMenu(0, 1, 0)
    Else
        Call LabelMenu(0, 2, 0)
    End If
    
    
    If (TestResult = "Fail") And (Tester.SaveOSLog.Value = 1) Then
        Print #10, "====================" & vbCrLf _
                    & "#" & SaveOSCounter & vbCrLf & "===================="
        Print #10, TempLogString & vbCrLf
        SaveOSCounter = SaveOSCounter + 1
    End If
    
    If Tester.SaveOSLog.Value = 1 Then
        Close #10
    End If
    
   Tester.Print "OpenShort :"; TestResult

End Sub

Public Sub OpenShortTest_Result()
'Dim OSValue As Byte
'Dim OSValue(0 To 127) As Long
Dim result

Dim i As Long
Dim tmp As Byte
Dim tmp2 As Byte
Dim T1
Dim DAQTime As Integer
Dim TempLogString As String


TempLogString = ""

' setting
If Tester.OSCheck.Value = 1 Then
   DAQTime = 500  ' to record standard
Else
   DAQTime = 4
End If



If OSFlag = 0 Or (OldOSStandFileName <> ChipName) Then

  Open App.Path & "/OSStandard/" & ChipName & ".txt" For Input As #6
  Input #6, OpenShortPinNo
  For i = 0 To OpenShortPinNo
        Input #6, tmp, OSStandard(i)
  Next i
  Close #6
  
  OSFlag = 1   ' inital ok
  OldOSStandFileName = ChipName
  
  CurrentOSFileName = "D:\OSFail_Log\" & ChipName & "_OSdatalog_" & Year(Date) & Format(Month(Date), "0#") _
                      & Format(Day(Date), "0#") & ".txt"
  SaveOSCounter = 1
  
End If

If Tester.SaveOSLog.Value = 1 Then
    Open CurrentOSFileName For Append As #10
End If
    
 
'Dim OSString As Stringt
 
    T1 = Timer
    For i = 0 To OpenShortPinNo
      CardResult = DO_WritePort(card, Channel_P1A, i)  ' 1111 1110
      Call Timer_1ms(DAQTime)
      'Call Timer_1ms(4)
      CardResult = DO_ReadPort(card, Channel_P1B, OSValue(i))
      
    Next i
   
   Dim T2
     T2 = Timer
  
     Tester.Print "Time cycle:"; T2 - T1
     
If Tester.OSCheck.Value = 1 Then
    
    OpenShortFrm.Show
    OpenShortFrm.Cls
  
      For i = 0 To OpenShortPinNo
           
           If OSValue(i) = &HFC Then
               OpenShortFrm.Print CStr((i + 1)) & " is open *****"
           ElseIf OSValue(i) = &HFF Then
               OpenShortFrm.Print CStr((i + 1)) & " is Short ++++"
           ElseIf OSValue(i) <> &HFD Then
               OpenShortFrm.Print (i + 1); Hex(OSValue(i))
           End If
      Next
     OpenShortFrm.Print "hint: save Rec"
  Exit Sub
End If
  
  
TestResult = "PASS"

For i = 0 To OpenShortPinNo
    ' OSValue(i) = CAndValue(OSValue(i), &HC0)
    If (OSValue(i) <> OSStandard(i)) Then
    
         If OSValue(i) = &HFF Then
            Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         ElseIf OSValue(i) = &HFC Then
            Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         Else
            Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         End If
         
         If Tester.SaveOSLog.Value = 1 Then
            
            If OSValue(i) = &HFF Then
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Short" & vbCrLf
                'Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            ElseIf OSValue(i) = &HFC Then
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Open" & vbCrLf
                'Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            Else
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Normal" & vbCrLf
                'Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            End If
            
         End If
         
         TestResult = "Fail"
        
    End If
Next
    
    If TestResult = "PASS" Then
        Call LabelMenu(0, 1, 0)
    Else
        Call LabelMenu(0, 2, 0)
    End If
    
    
    If (TestResult = "Fail") And (Tester.SaveOSLog.Value = 1) Then
        Print #10, "====================" & vbCrLf _
                    & "#" & SaveOSCounter & vbCrLf & "===================="
        Print #10, TempLogString & vbCrLf
        SaveOSCounter = SaveOSCounter + 1
    End If
    
    If Tester.SaveOSLog.Value = 1 Then
        Close #10
    End If
    
    If TestResult = "PASS" Then
        OS_Result = 1
    Else
        OS_Result = 2
    End If
    
    Call LabelMenu(0, OS_Result, 0)
    
    Tester.Print "OpenShort :"; TestResult
    
    TestResult = ""
    
End Sub
Public Sub OpenShortTest_Result_ByOSFileName()
'Dim OSValue As Byte
'Dim OSValue(0 To 127) As Long
Dim result

Dim i As Long
Dim tmp As Byte
Dim tmp2 As Byte
Dim T1
Dim DAQTime As Integer
Dim TempLogString As String


TempLogString = ""

' setting
If Tester.OSCheck.Value = 1 Then
   DAQTime = 500  ' to record standard
Else
   DAQTime = 4
End If


If OSFileName = "" Then
    MsgBox ("Can't Get OpenShort File !")
End If


If OSFlag = 0 Or (OldOSStandFileName <> ChipName) Then

  Open App.Path & "/OSStandard/" & OSFileName & ".txt" For Input As #6
  Input #6, OpenShortPinNo
  For i = 0 To OpenShortPinNo
        Input #6, tmp, OSStandard(i)
  Next i
  Close #6
  
  OSFlag = 1   ' inital ok
  OldOSStandFileName = ChipName
  
  CurrentOSFileName = "D:\OSFail_Log\" & ChipName & "_OSdatalog_" & Year(Date) & Format(Month(Date), "0#") _
                      & Format(Day(Date), "0#") & ".txt"
  SaveOSCounter = 1
  
End If

If Tester.SaveOSLog.Value = 1 Then
    Open CurrentOSFileName For Append As #10
End If
    
 
'Dim OSString As Stringt
 
    T1 = Timer
    For i = 0 To OpenShortPinNo
      CardResult = DO_WritePort(card, Channel_P1A, i)  ' 1111 1110
      Call Timer_1ms(DAQTime)
      'Call Timer_1ms(4)
      CardResult = DO_ReadPort(card, Channel_P1B, OSValue(i))
      
    Next i
   
   Dim T2
     T2 = Timer
  
     Tester.Print "Time cycle:"; T2 - T1
     
If Tester.OSCheck.Value = 1 Then
    
    OpenShortFrm.Show
    OpenShortFrm.Cls
  
      For i = 0 To OpenShortPinNo
           
           If OSValue(i) = &HFC Then
               OpenShortFrm.Print CStr((i + 1)) & " is open *****"
           ElseIf OSValue(i) = &HFF Then
               OpenShortFrm.Print CStr((i + 1)) & " is Short ++++"
           ElseIf OSValue(i) <> &HFD Then
               OpenShortFrm.Print (i + 1); Hex(OSValue(i))
           End If
      Next
     OpenShortFrm.Print "hint: save Rec"
  Exit Sub
End If
  
  
TestResult = "PASS"

For i = 0 To OpenShortPinNo
    ' OSValue(i) = CAndValue(OSValue(i), &HC0)
    If (OSValue(i) <> OSStandard(i)) Then
    
         If OSValue(i) = &HFF Then
            Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         ElseIf OSValue(i) = &HFC Then
            Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         Else
            Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         End If
         
         If Tester.SaveOSLog.Value = 1 Then
            
            If OSValue(i) = &HFF Then
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Short" & vbCrLf
                'Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            ElseIf OSValue(i) = &HFC Then
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Open" & vbCrLf
                'Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            Else
                TempLogString = TempLogString & "Pin" & Format(i + 1, "0#") & vbTab & "Normal" & vbCrLf
                'Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            End If
            
         End If
         
         TestResult = "Fail"
        
    End If
Next
    
    If TestResult = "PASS" Then
        Call LabelMenu(0, 1, 0)
    Else
        Call LabelMenu(0, 2, 0)
    End If
    
    
    If (TestResult = "Fail") And (Tester.SaveOSLog.Value = 1) Then
        Print #10, "====================" & vbCrLf _
                    & "#" & SaveOSCounter & vbCrLf & "===================="
        Print #10, TempLogString & vbCrLf
        SaveOSCounter = SaveOSCounter + 1
    End If
    
    If Tester.SaveOSLog.Value = 1 Then
        Close #10
    End If
    
    If TestResult = "PASS" Then
        OS_Result = 1
    Else
        OS_Result = 2
    End If
    
    Call LabelMenu(0, OS_Result, 0)
    
    Tester.Print "OpenShort :"; TestResult
    
    TestResult = ""
    
End Sub

Public Sub OpenShortTest_Result_AU6433EFF35()
'Dim OSValue As Byte
'Dim OSValue(0 To 127) As Long

'This Code purpose for AU6433EFF35 Qual-Site(Short) & SRM(Normal) Socket Board pin7 (Ground) different

Dim result
Dim i As Long
Dim tmp As Byte
Dim tmp2 As Byte
Dim T1
Dim DAQTime As Integer

' setting
If Tester.OSCheck.Value = 1 Then
   DAQTime = 500  ' to record standard
Else
   DAQTime = 4
End If



If OSFlag = 0 Or (OldOSStandFileName <> ChipName) Then

  Open App.Path & "/OSStandard/" & ChipName & ".txt" For Input As #6
  Input #6, OpenShortPinNo
  For i = 0 To OpenShortPinNo
        Input #6, tmp, OSStandard(i)
  Next i
  Close #6
  
    'If PCI7248InitFinish = 0 Then
    '      PCI7248Exist
    '      Call PowerSet2(1, "3.2", "0.5", 1, "0.2", "0.5", 1)
    '       Call SetTimer_1ms
       
    'End If
  OSFlag = 1   ' inital ok
  OldOSStandFileName = ChipName
End If


    
 
'Dim OSString As Stringt
 
    T1 = Timer
  For i = 0 To OpenShortPinNo
      CardResult = DO_WritePort(card, Channel_P1A, i)  ' 1111 1110
      Call Timer_1ms(DAQTime)
      CardResult = DO_ReadPort(card, Channel_P1B, OSValue(i))
  Next i
   Dim T2
     T2 = Timer
  
     Tester.Print "Time cycle:"; T2 - T1
     
If Tester.OSCheck.Value = 1 Then
    
    OpenShortFrm.Show
    OpenShortFrm.Cls
  
      For i = 0 To OpenShortPinNo
           
           If OSValue(i) = &HFC Then
               OpenShortFrm.Print CStr((i + 1)) & " is open *****"
           ElseIf OSValue(i) = &HFF Then
               OpenShortFrm.Print CStr((i + 1)) & " is Short ++++"
           ElseIf OSValue(i) <> &HFD Then
               OpenShortFrm.Print (i + 1); Hex(OSValue(i))
           End If
      Next
     OpenShortFrm.Print "hint: save Rec"
  Exit Sub
End If
  
  
TestResult = "PASS"

  For i = 0 To OpenShortPinNo
    ' OSValue(i) = CAndValue(OSValue(i), &HC0)
      If (OSValue(i) <> OSStandard(i)) And (i <> 6) Then
             
         If OSValue(i) = &HFF Then
            Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         ElseIf OSValue(i) = &HFC Then
            Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         Else
            Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         End If
         TestResult = "Fail"
        
      End If
        
      If (i = 6) And (OSValue(i) = &HFC) Then
            Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            TestResult = "Fail"
      End If
  
  Next
  
    

    If TestResult = "PASS" Then
        OS_Result = 1
    Else
        OS_Result = 2
    End If
    
    Call LabelMenu(0, OS_Result, 0)
    
    Tester.Print "OpenShort :"; TestResult
    
    TestResult = ""
    
End Sub


Public Sub OpenShortTest_Result_AU6435ELF33()
'Dim OSValue As Byte
'Dim OSValue(0 To 127) As Long

'Skip Pin34(ctrl0) O/S test

Dim result
Dim i As Long
Dim tmp As Byte
Dim tmp2 As Byte
Dim T1
Dim DAQTime As Integer

' setting
If Tester.OSCheck.Value = 1 Then
   DAQTime = 500  ' to record standard
Else
   DAQTime = 4
End If



If OSFlag = 0 Or (OldOSStandFileName <> ChipName) Then

  Open App.Path & "/OSStandard/" & ChipName & ".txt" For Input As #6
  Input #6, OpenShortPinNo
  For i = 0 To OpenShortPinNo
        Input #6, tmp, OSStandard(i)
  Next i
  Close #6
  
  OSFlag = 1   ' inital ok
  OldOSStandFileName = ChipName
End If


    
 
'Dim OSString As Stringt
 
    T1 = Timer
  For i = 0 To OpenShortPinNo
      CardResult = DO_WritePort(card, Channel_P1A, i)  ' 1111 1110
      Call Timer_1ms(DAQTime)
      CardResult = DO_ReadPort(card, Channel_P1B, OSValue(i))
  Next i
   Dim T2
     T2 = Timer
  
     Tester.Print "Time cycle:"; T2 - T1
     
If Tester.OSCheck.Value = 1 Then
    
    OpenShortFrm.Show
    OpenShortFrm.Cls
  
      For i = 0 To OpenShortPinNo
           
           If OSValue(i) = &HFC Then
               OpenShortFrm.Print CStr((i + 1)) & " is open *****"
           ElseIf OSValue(i) = &HFF Then
               OpenShortFrm.Print CStr((i + 1)) & " is Short ++++"
           ElseIf OSValue(i) <> &HFD Then
               OpenShortFrm.Print (i + 1); Hex(OSValue(i))
           End If
      Next
     OpenShortFrm.Print "hint: save Rec"
  Exit Sub
End If
  
  
TestResult = "PASS"

  For i = 0 To OpenShortPinNo
    ' OSValue(i) = CAndValue(OSValue(i), &HC0)
      If (OSValue(i) <> OSStandard(i)) And (i <> 33) Then
             
         If OSValue(i) = &HFF Then
            Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         ElseIf OSValue(i) = &HFC Then
            Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         Else
            Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         End If
         TestResult = "Fail"
        
      End If
  
  Next
  
    

    If TestResult = "PASS" Then
        OS_Result = 1
    Else
        OS_Result = 2
    End If
    
    Call LabelMenu(0, OS_Result, 0)
    
    Tester.Print "OpenShort :"; TestResult
    
    TestResult = ""
    
End Sub
Public Sub OpenShortTest_Result_AU6435ELF34()
'Dim OSValue As Byte
'Dim OSValue(0 To 127) As Long

Dim result
Dim i As Long
Dim tmp As Byte
Dim tmp2 As Byte
Dim T1
Dim DAQTime As Integer

' setting
If Tester.OSCheck.Value = 1 Then
   DAQTime = 500  ' to record standard
Else
   DAQTime = 4
End If



If OSFlag = 0 Or (OldOSStandFileName <> ChipName) Then

  Open App.Path & "/OSStandard/" & ChipName & ".txt" For Input As #6
  Input #6, OpenShortPinNo
  For i = 0 To OpenShortPinNo
        Input #6, tmp, OSStandard(i)
  Next i
  Close #6
  
  OSFlag = 1   ' inital ok
  OldOSStandFileName = ChipName
End If


    
 
'Dim OSString As Stringt
 
    T1 = Timer
  For i = 0 To OpenShortPinNo
      CardResult = DO_WritePort(card, Channel_P1A, i)  ' 1111 1110
      Call Timer_1ms(DAQTime)
      If i = 46 Then
        Call Timer_1ms(5)
      End If
      CardResult = DO_ReadPort(card, Channel_P1B, OSValue(i))
  Next i
   Dim T2
     T2 = Timer
  
     Tester.Print "Time cycle:"; T2 - T1
     
If Tester.OSCheck.Value = 1 Then
    
    OpenShortFrm.Show
    OpenShortFrm.Cls
  
      For i = 0 To OpenShortPinNo
           
           If OSValue(i) = &HFC Then
               OpenShortFrm.Print CStr((i + 1)) & " is open *****"
           ElseIf OSValue(i) = &HFF Then
               OpenShortFrm.Print CStr((i + 1)) & " is Short ++++"
           ElseIf OSValue(i) <> &HFD Then
               OpenShortFrm.Print (i + 1); Hex(OSValue(i))
           End If
      Next
     OpenShortFrm.Print "hint: save Rec"
  Exit Sub
End If
  
  
TestResult = "PASS"

  For i = 0 To OpenShortPinNo
    ' OSValue(i) = CAndValue(OSValue(i), &HC0)
      If (OSValue(i) <> OSStandard(i)) And (i <> 1) Then
             
         If OSValue(i) = &HFF Then
            Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         ElseIf OSValue(i) = &HFC Then
            Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         Else
            Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         End If
         TestResult = "Fail"
        
      End If
  
  Next
  
    

    If TestResult = "PASS" Then
        OS_Result = 1
    Else
        OS_Result = 2
    End If
    
    Call LabelMenu(0, OS_Result, 0)
    
    Tester.Print "OpenShort :"; TestResult
    
    TestResult = ""
    
End Sub

Public Sub OpenShortTest_Result_SkipPin2()
'Dim OSValue As Byte
'Dim OSValue(0 To 127) As Long

'Skip Pin34(ctrl0) O/S test

Dim result
Dim i As Long
Dim tmp As Byte
Dim tmp2 As Byte
Dim T1
Dim DAQTime As Integer

' setting
If Tester.OSCheck.Value = 1 Then
   DAQTime = 500  ' to record standard
Else
   DAQTime = 4
End If



If OSFlag = 0 Or (OldOSStandFileName <> ChipName) Then

  Open App.Path & "/OSStandard/" & ChipName & ".txt" For Input As #6
  Input #6, OpenShortPinNo
  For i = 0 To OpenShortPinNo
        Input #6, tmp, OSStandard(i)
  Next i
  Close #6
  
  OSFlag = 1   ' inital ok
  OldOSStandFileName = ChipName
End If


    
 
'Dim OSString As Stringt
 
    T1 = Timer
  For i = 0 To OpenShortPinNo
      CardResult = DO_WritePort(card, Channel_P1A, i)  ' 1111 1110
      Call Timer_1ms(DAQTime)
      CardResult = DO_ReadPort(card, Channel_P1B, OSValue(i))
  Next i
   Dim T2
     T2 = Timer
  
     Tester.Print "Time cycle:"; T2 - T1
     
If Tester.OSCheck.Value = 1 Then
    
    OpenShortFrm.Show
    OpenShortFrm.Cls
  
      For i = 0 To OpenShortPinNo
           
           If OSValue(i) = &HFC Then
               OpenShortFrm.Print CStr((i + 1)) & " is open *****"
           ElseIf OSValue(i) = &HFF Then
               OpenShortFrm.Print CStr((i + 1)) & " is Short ++++"
           ElseIf OSValue(i) <> &HFD Then
               OpenShortFrm.Print (i + 1); Hex(OSValue(i))
           End If
      Next
     OpenShortFrm.Print "hint: save Rec"
  Exit Sub
End If
  
  
TestResult = "PASS"

  For i = 0 To OpenShortPinNo
    ' OSValue(i) = CAndValue(OSValue(i), &HC0)
      If (OSValue(i) <> OSStandard(i)) And (i <> 1) Then
             
         If OSValue(i) = &HFF Then
            Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         ElseIf OSValue(i) = &HFC Then
            Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         Else
            Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         End If
         TestResult = "Fail"
        
      End If
  
  Next
  
    

    If TestResult = "PASS" Then
        OS_Result = 1
    Else
        OS_Result = 2
    End If
    
    Call LabelMenu(0, OS_Result, 0)
    
    Tester.Print "OpenShort :"; TestResult
    
    TestResult = ""
    
End Sub

Public Sub OpenShortTest_Result_AU6438IFF30()
'Dim OSValue As Byte
'Dim OSValue(0 To 127) As Long

'This Code purpose for AU6433EFF35 Qual-Site(Short) & SRM(Normal) Socket Board pin7 (Ground) different

Dim result
Dim i As Long
Dim tmp As Byte
Dim tmp2 As Byte
Dim T1
Dim DAQTime As Integer

' setting
If Tester.OSCheck.Value = 1 Then
   DAQTime = 500  ' to record standard
Else
   DAQTime = 4
End If



If OSFlag = 0 Or (OldOSStandFileName <> ChipName) Then

  Open App.Path & "/OSStandard/" & ChipName & ".txt" For Input As #6
  Input #6, OpenShortPinNo
  For i = 0 To OpenShortPinNo
        Input #6, tmp, OSStandard(i)
  Next i
  Close #6
  
    'If PCI7248InitFinish = 0 Then
    '      PCI7248Exist
    '      Call PowerSet2(1, "3.2", "0.5", 1, "0.2", "0.5", 1)
    '       Call SetTimer_1ms
       
    'End If
  OSFlag = 1   ' inital ok
  OldOSStandFileName = ChipName
End If


    
 
'Dim OSString As Stringt
 
T1 = Timer
For i = 0 To OpenShortPinNo
    CardResult = DO_WritePort(card, Channel_P1A, i)  ' 1111 1110
    Call Timer_1ms(DAQTime)
    If i = 1 Then
        Call MsecDelay(0.2)
    End If
    CardResult = DO_ReadPort(card, Channel_P1B, OSValue(i))
Next i

Dim T2
T2 = Timer

Tester.Print "Time cycle:"; T2 - T1
     
If Tester.OSCheck.Value = 1 Then
    
    OpenShortFrm.Show
    OpenShortFrm.Cls
  
      For i = 0 To OpenShortPinNo
           
           If OSValue(i) = &HFC Then
               OpenShortFrm.Print CStr((i + 1)) & " is open *****"
           ElseIf OSValue(i) = &HFF Then
               OpenShortFrm.Print CStr((i + 1)) & " is Short ++++"
           ElseIf OSValue(i) <> &HFD Then
               OpenShortFrm.Print (i + 1); Hex(OSValue(i))
           End If
      Next
     OpenShortFrm.Print "hint: save Rec"
  Exit Sub
End If
  
  
TestResult = "PASS"

  For i = 0 To OpenShortPinNo
    ' OSValue(i) = CAndValue(OSValue(i), &HC0)
      If (OSValue(i) <> OSStandard(i)) And (i <> 6) Then
             
         If OSValue(i) = &HFF Then
            Tester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         ElseIf OSValue(i) = &HFC Then
            Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         Else
            Tester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
         End If
         TestResult = "Fail"
        
      End If
        
      If (i = 6) And (OSValue(i) = &HFC) Then
            Tester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
            TestResult = "Fail"
      End If
  
  Next
  
    

    If TestResult = "PASS" Then
        OS_Result = 1
    Else
        OS_Result = 2
    End If
    
    Call LabelMenu(0, OS_Result, 0)
    
    Tester.Print "OpenShort :"; TestResult
    
    TestResult = ""
    
End Sub

