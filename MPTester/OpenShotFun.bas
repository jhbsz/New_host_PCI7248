Attribute VB_Name = "OpenShotFun"
Option Explicit
Global OSValue(0 To 127) As Long
Global OSStandard(0 To 127) As Byte
Global OSFlag As Byte
Global OpenShortPinNo As Byte
Global OldOSStandFileName As String

Public Function OpenShortTest_AssignOSFileName(ThisOSName As String) As Byte
'Dim OSValue As Byte
'Dim OSValue(0 To 127) As Long
Dim result

Dim i As Long
Dim tmp As Byte
Dim tmp2 As Byte
Dim T1
Dim DAQTime As Integer
'Dim TempLogString As String
Dim OSResult As Byte


'TempLogString = ""

DAQTime = 4


If OSFlag = 0 Or (OldOSStandFileName <> ChipName) Then

    Open App.Path & "/OSTemplate/" & ThisOSName & ".txt" For Input As #6
    Input #6, OpenShortPinNo
    For i = 0 To OpenShortPinNo
        Input #6, tmp, OSStandard(i)
    Next i
    Close #6
  
    OSFlag = 1   ' inital ok
    OldOSStandFileName = ChipName

End If
  
T1 = Timer
For i = 0 To OpenShortPinNo
    cardresult = DO_WritePort(card, Channel_P1A, i)  ' 1111 1110
    Call Timer_1ms(DAQTime)
    cardresult = DO_ReadPort(card, Channel_P1B, OSValue(i))
Next i
   
Dim T2
T2 = Timer
     
OSResult = 1

For i = 0 To OpenShortPinNo

' OSValue(i) = CAndValue(OSValue(i), &HC0)
    If (OSValue(i) <> OSStandard(i)) Then
    
        If OSValue(i) = &HFF Then
           MPTester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
        ElseIf OSValue(i) = &HFC Then
           MPTester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
        Else
           MPTester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
        End If
        
        OSResult = 0
    
    End If
Next
   

OpenShortTest_AssignOSFileName = OSResult
   
End Function

Public Function OpenShortTest_AssignOSFileName_NewBoard(ThisOSName As String) As Byte
'Dim OSValue As Byte
'Dim OSValue(0 To 127) As Long
Dim result

Dim i As Long
Dim tmp As Byte
Dim tmp2 As Byte
Dim T1
Dim DAQTime As Integer
'Dim TempLogString As String
Dim OSResult As Byte


'TempLogString = ""

DAQTime = 1


If OSFlag = 0 Or (OldOSStandFileName <> ChipName) Then

    Open App.Path & "/OSTemplate/" & ThisOSName & ".txt" For Input As #6
    Input #6, OpenShortPinNo
    For i = 0 To OpenShortPinNo
        Input #6, tmp, OSStandard(i)
    Next i
    Close #6
  
    OSFlag = 1   ' inital ok
    OldOSStandFileName = ChipName

End If
  
T1 = Timer
For i = 0 To OpenShortPinNo
    cardresult = DO_WritePort(card, Channel_P1A, i)  ' 1111 1110
    Call Timer_500us(DAQTime)
    cardresult = DO_ReadPort(card, Channel_P1B, OSValue(i))
Next i
   
Dim T2
T2 = Timer
     
OSResult = 1

For i = 0 To OpenShortPinNo

' OSValue(i) = CAndValue(OSValue(i), &HC0)
    If (OSValue(i) <> OSStandard(i)) Then
    
        If OSValue(i) = &HFF Then
           MPTester.Print "Short: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
        ElseIf OSValue(i) = &HFC Then
           MPTester.Print "Open: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
        Else
           MPTester.Print "XXXX: "; (i + 1), Hex((OSValue(i))), Hex(OSStandard(i))
        End If
        
        OSResult = 0
    
    End If
Next
   

OpenShortTest_AssignOSFileName_NewBoard = OSResult
   
End Function
