Attribute VB_Name = "AU6621Reader"
Public OpenCOM2Flag As Boolean
Public AU6621ContiFail As Integer
Public AU6621TestCount As Integer

Public Sub AU6601A71CLF21Testsub()

Dim TempRes As Byte
Dim BootTime As Single
Dim RemoveTime As Single
Dim LEDVal As Long

BootTime = 0.2
RemoveTime = 0.5

'2014/04/14: Copy from AU6621E84CFF21TestSub

' 8 7 6 5   4 3 2 1
' | | | |   | | | |
' | | | |   | | | |
'
'   H O O   O M S E
'   O P P   P S D N
'   T T T   T     A
'   P 2 1   0
'   L
'   U
'   G


' 1 1 0 0   1 1 1 0    opt0 enable       0xCE
' 1 1 0 1   0 1 1 0    opt1 enable       0xD6
' 1 1 1 0   0 1 1 0    opt2 enable       0xE6
' 1 1 1 1   1 1 1 0    opt0,1,2 enable   0xFE
' 1 0 0 0   0 1 1 0    HotPlug enable(remove device) 0x86
 
MPTester.TestResultLab = ""

TempRes = 0
If PCI7248InitFinish = 0 Then
    PCI7248Exist
End If

If Not OpenCOM2Flag Then
    MPTester.MSComm2.CommPort = 2
    If MPTester.MSComm2.PortOpen Then
        MPTester.MSComm2.PortOpen = False
    End If
    MPTester.MSComm2.Settings = "9600,N,8,1"
    MPTester.MSComm2.PortOpen = True
    
    MPTester.MSComm2.InBufferCount = 0
    MPTester.MSComm2.InputLen = 0
    OpenCOM2Flag = True
End If
   
MPTester.MSComm2.Output = "AU6621SetClear"


cardresult = DO_WritePort(card, Channel_P1A, &HC6)
Call MsecDelay(BootTime)
If (AU6621COM_RWTest("VEN1AEADEV6601", 3#, False) <> 1) Then    'check VID/DID
    TempRes = 2
    MPTester.Print "Find 1AEA/6601 fail ..."
    AU6621ContiFail = AU6621ContiFail + 1
    GoTo AU6601TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 1AEA/6601 PASS"
End If
'
'If (AU6621COM_RWTest("ven1AEAdev6621", 3#, False) <> 1) Then    'check DVID/DDID
'    TempRes = 2
'    MPTester.Print "Find 1AEA/6621 fail ..."
'    GoTo AU6601TestEndLabel
'Else
'    TempRes = 1
'    MPTester.Print "Check 1AEA/6621 PASS"
'End If


'R/W SD
'flow5 (P1A write 1111 1100 & PowerON) (target boot 1AEA/6621, SUBSYS: 0001/0001)
'cardresult = DO_WritePort(card, Channel_P1A, &HFC)
cardresult = DO_WritePort(card, Channel_P1A, &HC4)
Call MsecDelay(0.02)
'Call GPIBWrite("OUT1 1")
'Call GPIBWrite("OUT2 1")
Call MsecDelay(BootTime)

If (AU6621COM_RWTest("AU6601A71SDFun", 5#, True) <> 1) Then
    TempRes = 3
    MPTester.Print "SD Test Fail ..."
    GoTo AU6601TestEndLabel
Else
    TempRes = 1
    MPTester.Print "SD Test PASS"
End If


'R/W MS
'cardresult = DO_WritePort(card, Channel_P1A, &HFE)
cardresult = DO_WritePort(card, Channel_P1A, &HC6)
Call MsecDelay(0.1)
'cardresult = DO_WritePort(card, Channel_P1A, &HFA)
cardresult = DO_WritePort(card, Channel_P1A, &HC2)

If (AU6621COM_RWTest("AU6601A71MSFun", 5#, True) <> 1) Then
    TempRes = 5
    MPTester.Print "MS Test Fail ..."
    GoTo AU6601TestEndLabel
Else
    TempRes = 1
    MPTester.Print "MS Test PASS"
End If

If TempRes = 1 Then
    cardresult = DO_ReadPort(card, Channel_P1B, LEDVal)
    If (LEDVal And &H1) <> 0 Then
        TempRes = 4
        MPTester.Print "LED Fail ..."
    Else
        MPTester.Print "LED PASS"
    End If
End If

AU6601TestEndLabel:


'    Call GPIBWrite("OUT1 0")
'    Call GPIBWrite("OUT2 0")
    
    'cardresult = DO_WritePort(card, Channel_P1A, &HFF)
'    cardresult = DO_WritePort(card, Channel_P1A, &H86)
'    Call MsecDelay(0.2)
    cardresult = DO_WritePort(card, Channel_P1A, &H81)
    
    MPTester.MSComm2.InBufferCount = 0
    COM2RecBuf = ""
    EntryTime = Timer
    Do
        COM2RecBuf = COM2RecBuf & MPTester.MSComm2.Input
        If COM2RecBuf <> "" Then
        
            If InStr(COM2RecBuf, "NoDev") > 0 Then
                Exit Do
            End If
        End If
        PassingTime = Timer - EntryTime
    Loop Until (PassingTime >= 4)
    
'    If (PassingTime >= 4) Then
'        MPTester.MSComm2.Output = "DevScanPCIeBus"
'        Call MsecDelay(3#)
'    End If
    
'    AU6621TestCount = AU6621TestCount + 1
'    If AU6621TestCount > 50 Then
'        AU6621TestCount = 0
'        MPTester.MSComm2.Output = "RestartPCIHOST"
'        Call MsecDelay(4#)
'    End If
    
'    If AU6621ContiFail >= 5 Then
'        Call MsecDelay(1.2)
'        MPTester.MSComm2.Output = "DevScanPCIeBus"
'        Call MsecDelay(8#)
'        AU6621ContiFail = 0
'    Else
'        Call MsecDelay(0.6)
'    End If
    
    Call MsecDelay(0.2)
    
    If TempRes = 2 Then
        MPTester.TestResultLab = "Bin2: Enum Fail"
        TestResult = "Bin2"
    ElseIf TempRes = 3 Then
        MPTester.TestResultLab = "Bin3: SD Fail"
        TestResult = "Bin3"
        AU6621ContiFail = 0
    ElseIf TempRes = 5 Then
        MPTester.TestResultLab = "Bin5: MS Fail"
        TestResult = "Bin5"
        AU6621ContiFail = 0
    ElseIf TempRes = 4 Then
        MPTester.TestResultLab = "Bin4: LED Fail"
        TestResult = "Bin4"
        AU6621ContiFail = 0
    ElseIf TempRes = 1 Then
        MPTester.TestResultLab = "PASS"
        TestResult = "PASS"
        AU6621ContiFail = 0
    Else
        MPTester.TestResultLab = "Undefine Fail"
        TestResult = "Bin2"
    End If

End Sub

Public Sub AU6601A71CFF30TestSub()

Dim TempRes As Byte
Dim BootTime As Single
Dim RemoveTime As Single
Dim LEDVal As Long
Dim OSRes As Byte

BootTime = 0.2
RemoveTime = 0.5

'2014/04/14: Copy from AU6621E84CFF21TestSub

' 8 7 6 5   4 3 2 1
' | | | |   | | | |
' | | | |   | | | |
'
'   H O O   O M S E
'   O P P   P S D N
'   T T T   T     A
'   P 2 1   0
'   L
'   U
'   G


' 1 1 0 0   1 1 1 0    opt0 enable       0xCE
' 1 1 0 1   0 1 1 0    opt1 enable       0xD6
' 1 1 1 0   0 1 1 0    opt2 enable       0xE6
' 1 1 1 1   1 1 1 0    opt0,1,2 enable   0xFE
' 1 0 0 0   0 1 1 0    HotPlug enable(remove device) 0x86
 

TempRes = 0
If PCI7248InitFinish = 0 Then
    PCI7248Exist
    Call SetTimer_500us
End If

If Not OpenCOM2Flag Then
    MPTester.MSComm2.CommPort = 2
    If MPTester.MSComm2.PortOpen Then
        MPTester.MSComm2.PortOpen = False
    End If
    MPTester.MSComm2.Settings = "9600,N,8,1"
    MPTester.MSComm2.PortOpen = True
    
    MPTester.MSComm2.InBufferCount = 0
    MPTester.MSComm2.InputLen = 0
    OpenCOM2Flag = True
End If

MPTester.MSComm2.Output = "AU6621SetClear"

MPTester.Cls
MPTester.TestResultLab = ""

'Switch to OpenShort site
cardresult = DO_WritePort(card, Channel_P1C, &H0)
Call MsecDelay(0.1)
OSRes = OpenShortTest_AssignOSFileName_NewBoard(ChipName)

If OSRes = 1 Then
    cardresult = DO_WritePort(card, Channel_P1C, &HFF)  'switch to FT2
    MPTester.Print "OpenShot Test ...PASS"
    Call MsecDelay(0.1)
Else
    MPTester.Print "OpenShot Test ...Fail"
    GoTo AU6601TestEndLabel
End If


cardresult = DO_WritePort(card, Channel_P1A, &HC6)
Call MsecDelay(BootTime)
If (AU6621COM_RWTest("VEN1AEADEV6601", 3#, False) <> 1) Then    'check VID/DID
    TempRes = 2
    MPTester.Print "Find 1AEA/6601 fail ..."
    AU6621ContiFail = AU6621ContiFail + 1
    GoTo AU6601TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 1AEA/6601 PASS"
End If


'R/W SD
'flow5 (P1A write 1111 1100 & PowerON) (target boot 1AEA/6621, SUBSYS: 0001/0001)
'cardresult = DO_WritePort(card, Channel_P1A, &HFC)
cardresult = DO_WritePort(card, Channel_P1A, &HC4)
Call MsecDelay(0.02)
'Call GPIBWrite("OUT1 1")
'Call GPIBWrite("OUT2 1")
Call MsecDelay(BootTime)

If (AU6621COM_RWTest("AU6601A71SDFun", 5#, True) <> 1) Then
    TempRes = 3
    MPTester.Print "SD Test Fail ..."
    GoTo AU6601TestEndLabel
Else
    TempRes = 1
    MPTester.Print "SD Test PASS"
End If


'R/W MS
'cardresult = DO_WritePort(card, Channel_P1A, &HFE)
cardresult = DO_WritePort(card, Channel_P1A, &HC6)
Call MsecDelay(0.1)
'cardresult = DO_WritePort(card, Channel_P1A, &HFA)
cardresult = DO_WritePort(card, Channel_P1A, &HC2)

If (AU6621COM_RWTest("AU6601A71MSFun", 5#, True) <> 1) Then
    TempRes = 5
    MPTester.Print "MS Test Fail ..."
    GoTo AU6601TestEndLabel
Else
    TempRes = 1
    MPTester.Print "MS Test PASS"
End If

If TempRes = 1 Then
    cardresult = DO_ReadPort(card, Channel_P1B, LEDVal)
    If (LEDVal And &H2) <> 0 Then
        TempRes = 4
        MPTester.Print "LED Fail ..."
    Else
        MPTester.Print "LED PASS"
    End If
End If

AU6601TestEndLabel:


    cardresult = DO_WritePort(card, Channel_P1A, &H81)
    
    If OSRes = 1 Then
        MPTester.MSComm2.InBufferCount = 0
        COM2RecBuf = ""
        EntryTime = Timer
        Do
            COM2RecBuf = COM2RecBuf & MPTester.MSComm2.Input
            If COM2RecBuf <> "" Then
            
                If InStr(COM2RecBuf, "NoDev") > 0 Then
                    Exit Do
                End If
            End If
            PassingTime = Timer - EntryTime
        Loop Until (PassingTime >= 4)
        
        Call MsecDelay(0.2)
    End If
    
'    If (PassingTime >= 4) Then
'        MPTester.MSComm2.Output = "DevScanPCIeBus"
'        Call MsecDelay(3#)
'    End If
    
'    AU6621TestCount = AU6621TestCount + 1
'    If AU6621TestCount > 50 Then
'        AU6621TestCount = 0
'        MPTester.MSComm2.Output = "RestartPCIHOST"
'        Call MsecDelay(4#)
'    End If
    
'    If AU6621ContiFail >= 5 Then
'        Call MsecDelay(1.2)
'        MPTester.MSComm2.Output = "DevScanPCIeBus"
'        Call MsecDelay(8#)
'        AU6621ContiFail = 0
'    Else
'        Call MsecDelay(0.6)
'    End If
    
    
    If OSRes <> 1 Then
        MPTester.TestResultLab = "Bin2: OpenShort Fail"
        TestResult = "Bin2"
    ElseIf TempRes = 2 Then
        MPTester.TestResultLab = "Bin2: Enum Fail"
        TestResult = "Bin2"
    ElseIf TempRes = 3 Then
        MPTester.TestResultLab = "Bin3: SD Fail"
        TestResult = "Bin3"
        AU6621ContiFail = 0
    ElseIf TempRes = 5 Then
        MPTester.TestResultLab = "Bin5: MS Fail"
        TestResult = "Bin5"
        AU6621ContiFail = 0
    ElseIf TempRes = 4 Then
        MPTester.TestResultLab = "Bin4: LED Fail"
        TestResult = "Bin4"
        AU6621ContiFail = 0
    ElseIf TempRes = 1 Then
        MPTester.TestResultLab = "PASS"
        TestResult = "PASS"
        AU6621ContiFail = 0
    Else
        MPTester.TestResultLab = "Undefine Fail"
        TestResult = "Bin2"
    End If

End Sub
Public Sub AU6621E84CFF21Testsub()

Dim TempRes As Byte
Dim BootTime As Single
Dim RemoveTime As Single
Dim LEDVal As Long

BootTime = 0.2
RemoveTime = 0.5

'20131205 AU6621(PCIE Card Reader) communication NBTester by COM2-RS232(USB cable)
'Transfer GPIO, GPIB test flow to single site control (purpose to reduce test time)

' 8 7 6 5   4 3 2 1
' | | | |   | | | |
' | | | |   | | | |
'
'   H O O   O M S E
'   O P P   P S D N
'   T T T   T     A
'   P 2 1   0
'   L
'   U
'   G


' 1 1 0 0   1 1 1 0    opt0 enable       0xCE
' 1 1 0 1   0 1 1 0    opt1 enable       0xD6
' 1 1 1 0   0 1 1 0    opt2 enable       0xE6
' 1 1 1 1   1 1 1 0    opt0,1,2 enable   0xFE
' 1 0 0 0   0 1 1 0    HotPlug enable(remove device) 0x86
 

TempRes = 0
If PCI7248InitFinish = 0 Then
    PCI7248Exist
End If

If Not OpenCOM2Flag Then
    MPTester.MSComm2.CommPort = 2
    If MPTester.MSComm2.PortOpen Then
        MPTester.MSComm2.PortOpen = False
    End If
    MPTester.MSComm2.Settings = "9600,N,8,1"
    MPTester.MSComm2.PortOpen = True
    
    MPTester.MSComm2.InBufferCount = 0
    MPTester.MSComm2.InputLen = 0
    OpenCOM2Flag = True
End If

MPTester.TestResultLab = ""
MPTester.MSComm2.Output = "AU6621SetClear"
'cardresult = DO_WritePort(card, Channel_P1A, &HFF)
'Call MsecDelay(0.02)
   
'MPTester.TestResultLab = ""
'MPTester.MSComm2.InBufferCount = 0
'COM2RecBuf = ""
'EntryTime = Timer

cardresult = DO_WritePort(card, Channel_P1A, &HC6)
Call MsecDelay(BootTime)
If (AU6621COM_RWTest("VEN1AEADEV6621", 3#, False) <> 1) Then    'check VID/DID
    TempRes = 2
    MPTester.Print "Find 1AEA/6621 fail ..."
    AU6621ContiFail = AU6621ContiFail + 1
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 1AEA/6621 PASS"
End If
'
'If (AU6621COM_RWTest("ven1AEAdev6621", 3#, False) <> 1) Then    'check DVID/DDID
'    TempRes = 2
'    MPTester.Print "Find 1AEA/6621 fail ..."
'    GoTo AU6621TestEndLabel
'Else
'    TempRes = 1
'    MPTester.Print "Check 1AEA/6621 PASS"
'End If


'R/W SD
'flow5 (P1A write 1111 1100 & PowerON) (target boot 1AEA/6621, SUBSYS: 0001/0001)
'cardresult = DO_WritePort(card, Channel_P1A, &HFC)
cardresult = DO_WritePort(card, Channel_P1A, &HC4)
Call MsecDelay(0.02)
'Call GPIBWrite("OUT1 1")
'Call GPIBWrite("OUT2 1")
Call MsecDelay(BootTime)

If (AU6621COM_RWTest("AU6621B82SDFun", 5#, True) <> 1) Then
    TempRes = 3
    MPTester.Print "SD Test Fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "SD Test PASS"
End If


'R/W MS
'cardresult = DO_WritePort(card, Channel_P1A, &HFE)
cardresult = DO_WritePort(card, Channel_P1A, &HC6)
Call MsecDelay(0.1)
'cardresult = DO_WritePort(card, Channel_P1A, &HFA)
cardresult = DO_WritePort(card, Channel_P1A, &HC2)

If (AU6621COM_RWTest("AU6621B82MSFun", 5#, True) <> 1) Then
    TempRes = 5
    MPTester.Print "MS Test Fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "MS Test PASS"
End If

If TempRes = 1 Then
    cardresult = DO_ReadPort(card, Channel_P1B, LEDVal)
    If (LEDVal And &H1) <> 0 Then
        TempRes = 4
        MPTester.Print "LED Fail ..."
    Else
        MPTester.Print "LED PASS"
    End If
End If

AU6621TestEndLabel:


'    Call GPIBWrite("OUT1 0")
'    Call GPIBWrite("OUT2 0")
    
    'cardresult = DO_WritePort(card, Channel_P1A, &HFF)
'    cardresult = DO_WritePort(card, Channel_P1A, &H86)
'    Call MsecDelay(0.2)
    cardresult = DO_WritePort(card, Channel_P1A, &H81)
    
    MPTester.MSComm2.InBufferCount = 0
    COM2RecBuf = ""
    EntryTime = Timer
    Do
        COM2RecBuf = COM2RecBuf & MPTester.MSComm2.Input
        If COM2RecBuf <> "" Then
        
            If InStr(COM2RecBuf, "NoDev") > 0 Then
                Exit Do
            End If
        End If
        PassingTime = Timer - EntryTime
    Loop Until (PassingTime >= 4)
    
'    If (PassingTime >= 4) Then
'        MPTester.MSComm2.Output = "DevScanPCIeBus"
'        Call MsecDelay(3#)
'    End If
    
'    AU6621TestCount = AU6621TestCount + 1
'    If AU6621TestCount > 50 Then
'        AU6621TestCount = 0
'        MPTester.MSComm2.Output = "RestartPCIHOST"
'        Call MsecDelay(4#)
'    End If
    
'    If AU6621ContiFail >= 5 Then
'        Call MsecDelay(1.2)
'        MPTester.MSComm2.Output = "DevScanPCIeBus"
'        Call MsecDelay(8#)
'        AU6621ContiFail = 0
'    Else
'        Call MsecDelay(0.6)
'    End If
    
    Call MsecDelay(0.2)
    
    If TempRes = 2 Then
        MPTester.TestResultLab = "Bin2: Enum Fail"
        TestResult = "Bin2"
    ElseIf TempRes = 3 Then
        MPTester.TestResultLab = "Bin3: SD Fail"
        TestResult = "Bin3"
        AU6621ContiFail = 0
    ElseIf TempRes = 5 Then
        MPTester.TestResultLab = "Bin5: MS Fail"
        TestResult = "Bin5"
        AU6621ContiFail = 0
    ElseIf TempRes = 4 Then
        MPTester.TestResultLab = "Bin4: LED Fail"
        TestResult = "Bin4"
        AU6621ContiFail = 0
    ElseIf TempRes = 1 Then
        MPTester.TestResultLab = "PASS"
        TestResult = "PASS"
        AU6621ContiFail = 0
    Else
        MPTester.TestResultLab = "Undefine Fail"
        TestResult = "Bin2"
    End If

End Sub

Public Sub AU6621E84CFF30Testsub()

Dim TempRes As Byte
Dim BootTime As Single
Dim RemoveTime As Single
Dim LEDVal As Long
Dim OSRes As Byte

BootTime = 0.2
RemoveTime = 0.5

'20131205 AU6621(PCIE Card Reader) communication NBTester by COM2-RS232(USB cable)
'Transfer GPIO, GPIB test flow to single site control (purpose to reduce test time)

' 8 7 6 5   4 3 2 1
' | | | |   | | | |
' | | | |   | | | |
'
'   H O O   O M S E
'   O P P   P S D N
'   T T T   T     A
'   P 2 1   0
'   L
'   U
'   G


' 1 1 0 0   1 1 1 0    opt0 enable       0xCE
' 1 1 0 1   0 1 1 0    opt1 enable       0xD6
' 1 1 1 0   0 1 1 0    opt2 enable       0xE6
' 1 1 1 1   1 1 1 0    opt0,1,2 enable   0xFE
' 1 0 0 0   0 1 1 0    HotPlug enable(remove device) 0x86
 

TempRes = 0
If PCI7248InitFinish = 0 Then
    PCI7248Exist
    Call SetTimer_500us
End If

If Not OpenCOM2Flag Then
    MPTester.MSComm2.CommPort = 2
    If MPTester.MSComm2.PortOpen Then
        MPTester.MSComm2.PortOpen = False
    End If
    MPTester.MSComm2.Settings = "9600,N,8,1"
    MPTester.MSComm2.PortOpen = True
    
    MPTester.MSComm2.InBufferCount = 0
    MPTester.MSComm2.InputLen = 0
    OpenCOM2Flag = True
End If

MPTester.MSComm2.Output = "AU6621SetClear"

MPTester.Cls
MPTester.TestResultLab = ""

'Switch to OpenShort site
cardresult = DO_WritePort(card, Channel_P1C, &H0)
Call MsecDelay(0.3)
OSRes = OpenShortTest_AssignOSFileName_NewBoard(ChipName)

If OSRes = 1 Then
    cardresult = DO_WritePort(card, Channel_P1C, &HFF)  'switch to FT2
    Call MsecDelay(0.1)
    MPTester.Print "OpenShot Test ...PASS"

Else
    MPTester.Print "OpenShot Test ...Fail"
    GoTo AU6621TestEndLabel
End If
   
MPTester.MSComm2.Output = "AU6621SetClear"
'cardresult = DO_WritePort(card, Channel_P1A, &HFF)
'Call MsecDelay(0.02)
   
'MPTester.TestResultLab = ""
'MPTester.MSComm2.InBufferCount = 0
'COM2RecBuf = ""
'EntryTime = Timer

cardresult = DO_WritePort(card, Channel_P1A, &HC6)
Call MsecDelay(BootTime)
If (AU6621COM_RWTest("VEN1AEADEV6621", 3#, False) <> 1) Then    'check VID/DID
    TempRes = 2
    MPTester.Print "Find 1AEA/6621 fail ..."
    AU6621ContiFail = AU6621ContiFail + 1
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 1AEA/6621 PASS"
End If
'
'If (AU6621COM_RWTest("ven1AEAdev6621", 3#, False) <> 1) Then    'check DVID/DDID
'    TempRes = 2
'    MPTester.Print "Find 1AEA/6621 fail ..."
'    GoTo AU6621TestEndLabel
'Else
'    TempRes = 1
'    MPTester.Print "Check 1AEA/6621 PASS"
'End If


'R/W SD
'flow5 (P1A write 1111 1100 & PowerON) (target boot 1AEA/6621, SUBSYS: 0001/0001)
'cardresult = DO_WritePort(card, Channel_P1A, &HFC)
cardresult = DO_WritePort(card, Channel_P1A, &HC4)
Call MsecDelay(0.02)
'Call GPIBWrite("OUT1 1")
'Call GPIBWrite("OUT2 1")
Call MsecDelay(BootTime)

If (AU6621COM_RWTest("AU6621B82SDFun", 5#, True) <> 1) Then
    TempRes = 3
    MPTester.Print "SD Test Fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "SD Test PASS"
End If


'R/W MS
'cardresult = DO_WritePort(card, Channel_P1A, &HFE)
cardresult = DO_WritePort(card, Channel_P1A, &HC6)
Call MsecDelay(0.1)
'cardresult = DO_WritePort(card, Channel_P1A, &HFA)
cardresult = DO_WritePort(card, Channel_P1A, &HC2)

If (AU6621COM_RWTest("AU6621B82MSFun", 5#, True) <> 1) Then
    TempRes = 5
    MPTester.Print "MS Test Fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "MS Test PASS"
End If

If TempRes = 1 Then
    cardresult = DO_ReadPort(card, Channel_P1B, LEDVal)
    If (LEDVal And &H2) <> 0 Then
        TempRes = 4
        MPTester.Print "LED Fail ..."
    Else
        MPTester.Print "LED PASS"
    End If
End If

AU6621TestEndLabel:


'    Call GPIBWrite("OUT1 0")
'    Call GPIBWrite("OUT2 0")
    
    'cardresult = DO_WritePort(card, Channel_P1A, &HFF)
'    cardresult = DO_WritePort(card, Channel_P1A, &H86)
'    Call MsecDelay(0.2)
    cardresult = DO_WritePort(card, Channel_P1A, &H81)
    
    If OSRes = 1 Then
        MPTester.MSComm2.InBufferCount = 0
        COM2RecBuf = ""
        EntryTime = Timer
        Do
            COM2RecBuf = COM2RecBuf & MPTester.MSComm2.Input
            If COM2RecBuf <> "" Then
            
                If InStr(COM2RecBuf, "NoDev") > 0 Then
                    Exit Do
                End If
            End If
            PassingTime = Timer - EntryTime
        Loop Until (PassingTime >= 4)
        
        Call MsecDelay(0.2)
    End If
    
'    If (PassingTime >= 4) Then
'        MPTester.MSComm2.Output = "DevScanPCIeBus"
'        Call MsecDelay(3#)
'    End If
    
'    AU6621TestCount = AU6621TestCount + 1
'    If AU6621TestCount > 50 Then
'        AU6621TestCount = 0
'        MPTester.MSComm2.Output = "RestartPCIHOST"
'        Call MsecDelay(4#)
'    End If
    
'    If AU6621ContiFail >= 5 Then
'        Call MsecDelay(1.2)
'        MPTester.MSComm2.Output = "DevScanPCIeBus"
'        Call MsecDelay(8#)
'        AU6621ContiFail = 0
'    Else
'        Call MsecDelay(0.6)
'    End If
    
    If OSRes <> 1 Then
        MPTester.TestResultLab = "Bin2: OpenShort Fail"
        TestResult = "Bin2"
    ElseIf TempRes = 2 Then
        MPTester.TestResultLab = "Bin2: Enum Fail"
        TestResult = "Bin2"
    ElseIf TempRes = 3 Then
        MPTester.TestResultLab = "Bin3: SD Fail"
        TestResult = "Bin3"
        AU6621ContiFail = 0
    ElseIf TempRes = 5 Then
        MPTester.TestResultLab = "Bin5: MS Fail"
        TestResult = "Bin5"
        AU6621ContiFail = 0
    ElseIf TempRes = 4 Then
        MPTester.TestResultLab = "Bin4: LED Fail"
        TestResult = "Bin4"
        AU6621ContiFail = 0
    ElseIf TempRes = 1 Then
        MPTester.TestResultLab = "PASS"
        TestResult = "PASS"
        AU6621ContiFail = 0
    Else
        MPTester.TestResultLab = "Undefine Fail"
        TestResult = "Bin2"
    End If

End Sub

Public Sub AU6621C83CFF21Testsub()

Dim TempRes As Byte
Dim BootTime As Single
Dim RemoveTime As Single
Dim LEDVal As Long

BootTime = 0.2
RemoveTime = 0.5

'20131205 AU6621(PCIE Card Reader) communication NBTester by COM2-RS232(USB cable)
'Transfer GPIO, GPIB test flow to single site control (purpose to reduce test time)

' 8 7 6 5   4 3 2 1
' | | | |   | | | |
' | | | |   | | | |
'
'   H O O   O M S E
'   O P P   P S D N
'   T T T   T     A
'   P 2 1   0
'   L
'   U
'   G


' 1 1 0 0   1 1 1 0    opt0 enable       0xCE
' 1 1 0 1   0 1 1 0    opt1 enable       0xD6
' 1 1 1 0   0 1 1 0    opt2 enable       0xE6
' 1 1 1 1   1 1 1 0    opt0,1,2 enable   0xFE
' 1 0 0 0   0 1 1 0    HotPlug enable(remove device) 0x86
 

TempRes = 0
If PCI7248InitFinish = 0 Then
    PCI7248Exist
End If

If Not OpenCOM2Flag Then
    MPTester.MSComm2.CommPort = 2
    If MPTester.MSComm2.PortOpen Then
        MPTester.MSComm2.PortOpen = False
    End If
    MPTester.MSComm2.Settings = "9600,N,8,1"
    MPTester.MSComm2.PortOpen = True
    
    MPTester.MSComm2.InBufferCount = 0
    MPTester.MSComm2.InputLen = 0
    OpenCOM2Flag = True
End If
   
MPTester.MSComm2.Output = "AU6621SetClear"
'cardresult = DO_WritePort(card, Channel_P1A, &HFF)
'Call MsecDelay(0.02)
   
MPTester.TestResultLab = ""

'flow1 (P1A write 1111 0110 & PowerON) (target boot 1AEA/6621, SUBSYS: 1179/F920)
'cardresult = DO_WritePort(card, Channel_P1A, &HF6)

cardresult = DO_WritePort(card, Channel_P1A, &HCE)
Call MsecDelay(0.02)
'Call PowerSet2(1, "3.3", "0.5", 1, "3.3", "0.5", 1)
Call MsecDelay(BootTime)
If (AU6621COM_RWTest("VEN1AEADEV6621", 3#, False) <> 1) Then    'check VID/DID
    TempRes = 2
    MPTester.Print "Find 1AEA/6621 fail ..."
    AU6621ContiFail = AU6621ContiFail + 1
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 1AEA/6621 PASS"
End If

If (AU6621COM_RWTest("ven1179devF920", 3#, False) <> 1) Then    'check DVID/DDID
    TempRes = 2
    MPTester.Print "Find 1179/F920 fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 1179/F920 PASS"
End If

cardresult = DO_WritePort(card, Channel_P1A, &HCF And &HBF)


MPTester.MSComm2.InBufferCount = 0
COM2RecBuf = ""
EntryTime = Timer
Do
    COM2RecBuf = COM2RecBuf & MPTester.MSComm2.Input
    If COM2RecBuf <> "" Then
    
        If InStr(COM2RecBuf, "NoDev") > 0 Then
            Exit Do
        End If
    End If
    PassingTime = Timer - EntryTime
Loop Until (PassingTime >= 4)

'If (PassingTime >= 4) Then
'    MPTester.MSComm2.Output = "DevScanPCIeBus"
'    Call MsecDelay(3#)
'End If

Call MsecDelay(RemoveTime)


'flow2 (P1A write 1110 1110 & PowerON) (target boot 1AEA/6601, SUBSYS: 1179/F920)
'cardresult = DO_WritePort(card, Channel_P1A, &HEE)
cardresult = DO_WritePort(card, Channel_P1A, &HD6)
'Call GPIBWrite("OUT1 1")
'Call GPIBWrite("OUT2 1")
Call MsecDelay(BootTime)


If (AU6621COM_RWTest("VEN1AEADEV6601", 3#, False) <> 1) Then    'check VID/DID
    TempRes = 2
    MPTester.Print "Find 1AEA/6601 fail ..."
    AU6621ContiFail = AU6621ContiFail + 1
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 1AEA/6601 PASS"
End If

If (AU6621COM_RWTest("ven1179devF900", 3#, False) <> 1) Then    'check DVID/DDID
    TempRes = 2
    MPTester.Print "Find 1179/F900 fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 1179/F900 PASS"
End If

cardresult = DO_WritePort(card, Channel_P1A, &HD7 And &HBF)

MPTester.MSComm2.InBufferCount = 0
COM2RecBuf = ""
EntryTime = Timer
Do
    COM2RecBuf = COM2RecBuf & MPTester.MSComm2.Input
    If COM2RecBuf <> "" Then
    
        If InStr(COM2RecBuf, "NoDev") > 0 Then
            Exit Do
        End If
    End If
    PassingTime = Timer - EntryTime
Loop Until (PassingTime >= 4)

'If (PassingTime >= 4) Then
'    MPTester.MSComm2.Output = "DevScanPCIeBus"
'    Call MsecDelay(3#)
'End If
Call MsecDelay(RemoveTime)


'flow3 (P1A write 1101 1110 & PowerON) (target boot 1AEA/6601, SUBSYS: 1179/F940)
'cardresult = DO_WritePort(card, Channel_P1A, &HDE)
cardresult = DO_WritePort(card, Channel_P1A, &HE6)
'Call GPIBWrite("OUT1 1")
'Call GPIBWrite("OUT2 1")
Call MsecDelay(BootTime)

If (AU6621COM_RWTest("VEN1AEADEV6601", 3#, False) <> 1) Then    'check VID/DID
    TempRes = 2
    MPTester.Print "Find 1AEA/6601 fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 1AEA/6601 PASS"
End If

If (AU6621COM_RWTest("ven1179devF940", 3#, False) <> 1) Then    'check DVID/DDID
    TempRes = 2
    MPTester.Print "Find 1179/F940 fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 1179/F940 PASS"
End If

cardresult = DO_WritePort(card, Channel_P1A, &HE7 And &HBF)

MPTester.MSComm2.InBufferCount = 0
COM2RecBuf = ""
EntryTime = Timer
Do
    COM2RecBuf = COM2RecBuf & MPTester.MSComm2.Input
    If COM2RecBuf <> "" Then
    
        If InStr(COM2RecBuf, "NoDev") > 0 Then
            Exit Do
        End If
    End If
    PassingTime = Timer - EntryTime
Loop Until (PassingTime >= 4)

'If (PassingTime >= 4) Then
'    MPTester.MSComm2.Output = "DevScanPCIeBus"
'    Call MsecDelay(3#)
'End If

Call MsecDelay(RemoveTime)

'flow4 (P1A write 1100 0110 & PowerON) (target boot 1AEA/6601, SUBSYS: 0001/0001)
'cardresult = DO_WritePort(card, Channel_P1A, &HC6)
cardresult = DO_WritePort(card, Channel_P1A, &HFE)
'Call GPIBWrite("OUT1 1")
'Call GPIBWrite("OUT2 1")
Call MsecDelay(BootTime)

If (AU6621COM_RWTest("VEN1AEADEV6601", 3#, False) <> 1) Then    'check VID/DID
    TempRes = 2
    MPTester.Print "Find 1AEA/6601 fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 1AEA/6601 PASS"
End If

If (AU6621COM_RWTest("ven0001dev0001", 3#, False) <> 1) Then    'check DVID/DDID
    TempRes = 2
    MPTester.Print "Find 0001/0001 fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 0001/0001 PASS"
End If

cardresult = DO_WritePort(card, Channel_P1A, &HFF And &HBF)

MPTester.MSComm2.InBufferCount = 0
COM2RecBuf = ""
EntryTime = Timer
Do
    COM2RecBuf = COM2RecBuf & MPTester.MSComm2.Input
    If COM2RecBuf <> "" Then
    
        If InStr(COM2RecBuf, "NoDev") > 0 Then
            Exit Do
        End If
    End If
    PassingTime = Timer - EntryTime
Loop Until (PassingTime >= 4)

'If (PassingTime >= 4) Then
'    MPTester.MSComm2.Output = "DevScanPCIeBus"
'    Call MsecDelay(3#)
'End If

Call MsecDelay(RemoveTime)

'R/W SD
'flow5 (P1A write 1111 1100 & PowerON) (target boot 1AEA/6621, SUBSYS: 0001/0001)
'cardresult = DO_WritePort(card, Channel_P1A, &HFC)
cardresult = DO_WritePort(card, Channel_P1A, &HC4)
Call MsecDelay(0.02)
'Call GPIBWrite("OUT1 1")
'Call GPIBWrite("OUT2 1")
Call MsecDelay(BootTime)

If (AU6621COM_RWTest("AU6621B82SDFun", 5#, True) <> 1) Then
    TempRes = 3
    MPTester.Print "SD Test Fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "SD Test PASS"
End If


'R/W MS
'cardresult = DO_WritePort(card, Channel_P1A, &HFE)
cardresult = DO_WritePort(card, Channel_P1A, &HC6)
Call MsecDelay(0.1)
'cardresult = DO_WritePort(card, Channel_P1A, &HFA)
cardresult = DO_WritePort(card, Channel_P1A, &HC2)

If (AU6621COM_RWTest("AU6621B82MSFun", 5#, True) <> 1) Then
    TempRes = 5
    MPTester.Print "MS Test Fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "MS Test PASS"
End If

If TempRes = 1 Then
    cardresult = DO_ReadPort(card, Channel_P1B, LEDVal)
    If (LEDVal And &H1) <> 0 Then
        TempRes = 4
        MPTester.Print "LED Fail ..."
    Else
        MPTester.Print "LED PASS"
    End If
End If

AU6621TestEndLabel:


'    Call GPIBWrite("OUT1 0")
'    Call GPIBWrite("OUT2 0")
    
    'cardresult = DO_WritePort(card, Channel_P1A, &HFF)
'    cardresult = DO_WritePort(card, Channel_P1A, &H86)
'    Call MsecDelay(0.2)
    cardresult = DO_WritePort(card, Channel_P1A, &H81)
    
    MPTester.MSComm2.InBufferCount = 0
    COM2RecBuf = ""
    EntryTime = Timer
    Do
        COM2RecBuf = COM2RecBuf & MPTester.MSComm2.Input
        If COM2RecBuf <> "" Then
        
            If InStr(COM2RecBuf, "NoDev") > 0 Then
                Exit Do
            End If
        End If
        PassingTime = Timer - EntryTime
    Loop Until (PassingTime >= 4)
    
'    If (PassingTime >= 4) Then
'        MPTester.MSComm2.Output = "DevScanPCIeBus"
'        Call MsecDelay(3#)
'    End If
    
'    AU6621TestCount = AU6621TestCount + 1
'    If AU6621TestCount > 50 Then
'        AU6621TestCount = 0
'        MPTester.MSComm2.Output = "RestartPCIHOST"
'        Call MsecDelay(4#)
'    End If
    
'    If AU6621ContiFail >= 5 Then
'        Call MsecDelay(1.2)
'        MPTester.MSComm2.Output = "DevScanPCIeBus"
'        Call MsecDelay(8#)
'        AU6621ContiFail = 0
'    Else
'        Call MsecDelay(0.6)
'    End If
    
    Call MsecDelay(0.2)
    
    If TempRes = 2 Then
        MPTester.TestResultLab = "Bin2: Enum Fail"
        TestResult = "Bin2"
    ElseIf TempRes = 3 Then
        MPTester.TestResultLab = "Bin3: SD Fail"
        TestResult = "Bin3"
        AU6621ContiFail = 0
    ElseIf TempRes = 5 Then
        MPTester.TestResultLab = "Bin5: MS Fail"
        TestResult = "Bin5"
        AU6621ContiFail = 0
    ElseIf TempRes = 4 Then
        MPTester.TestResultLab = "Bin4: LED Fail"
        TestResult = "Bin4"
        AU6621ContiFail = 0
    ElseIf TempRes = 1 Then
        MPTester.TestResultLab = "PASS"
        TestResult = "PASS"
        AU6621ContiFail = 0
    Else
        MPTester.TestResultLab = "Undefine Fail"
        TestResult = "Bin2"
    End If

End Sub

Public Sub AU6621C83CFF30Testsub()

Dim TempRes As Byte
Dim BootTime As Single
Dim RemoveTime As Single
Dim LEDVal As Long
Dim OSRes As Byte
Dim LEDOnFlag  As Boolean

BootTime = 0.2
RemoveTime = 0.5

'20131205 AU6621(PCIE Card Reader) communication NBTester by COM2-RS232(USB cable)
'Transfer GPIO, GPIB test flow to single site control (purpose to reduce test time)

' 8 7 6 5   4 3 2 1
' | | | |   | | | |
' | | | |   | | | |
'
'   H O O   O M S E
'   O P P   P S D N
'   T T T   T     A
'   P 2 1   0
'   L
'   U
'   G


' 1 1 0 0   1 1 1 0    opt0 enable       0xCE
' 1 1 0 1   0 1 1 0    opt1 enable       0xD6
' 1 1 1 0   0 1 1 0    opt2 enable       0xE6
' 1 1 1 1   1 1 1 0    opt0,1,2 enable   0xFE
' 1 0 0 0   0 1 1 0    HotPlug enable(remove device) 0x86
 

TempRes = 0
If PCI7248InitFinish = 0 Then
    PCI7248Exist
    Call SetTimer_500us
End If

If Not OpenCOM2Flag Then
    MPTester.MSComm2.CommPort = 2
    If MPTester.MSComm2.PortOpen Then
        MPTester.MSComm2.PortOpen = False
    End If
    MPTester.MSComm2.Settings = "9600,N,8,1"
    MPTester.MSComm2.PortOpen = True
    
    MPTester.MSComm2.InBufferCount = 0
    MPTester.MSComm2.InputLen = 0
    OpenCOM2Flag = True
End If

MPTester.MSComm2.Output = "AU6621SetClear"

MPTester.Cls
MPTester.TestResultLab = ""

'Switch to OpenShort site
cardresult = DO_WritePort(card, Channel_P1C, &H0)
Call MsecDelay(0.3)
OSRes = OpenShortTest_AssignOSFileName_NewBoard(ChipName)


cardresult = DO_WritePort(card, Channel_P1A, &HCF)

If OSRes = 1 Then
    cardresult = DO_WritePort(card, Channel_P1C, &HFF)  'switch to FT2
    MPTester.Print "OpenShot Test ...PASS"
    Call MsecDelay(0.1)
Else
    MPTester.Print "OpenShot Test ...Fail"
    GoTo AU6621TestEndLabel
End If


'flow1 (P1A write 1111 0110 & PowerON) (target boot 1AEA/6621, SUBSYS: 1179/F920)
'cardresult = DO_WritePort(card, Channel_P1A, &HF6)

cardresult = DO_WritePort(card, Channel_P1A, &HCE)
Call MsecDelay(0.02)
'Call PowerSet2(1, "3.3", "0.5", 1, "3.3", "0.5", 1)
Call MsecDelay(BootTime)
If (AU6621COM_RWTest("VEN1AEADEV6621", 3#, False) <> 1) Then    'check VID/DID
    TempRes = 2
    MPTester.Print "Find 1AEA/6621 fail ..."
    AU6621ContiFail = AU6621ContiFail + 1
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 1AEA/6621 PASS"
End If

If (AU6621COM_RWTest("ven1179devF920", 3#, False) <> 1) Then    'check DVID/DDID
    TempRes = 2
    MPTester.Print "Find 1179/F920 fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 1179/F920 PASS"
End If

cardresult = DO_WritePort(card, Channel_P1A, &HCF And &HBF)


MPTester.MSComm2.InBufferCount = 0
COM2RecBuf = ""
EntryTime = Timer
Do
    COM2RecBuf = COM2RecBuf & MPTester.MSComm2.Input
    If COM2RecBuf <> "" Then
    
        If InStr(COM2RecBuf, "NoDev") > 0 Then
            Exit Do
        End If
    End If
    PassingTime = Timer - EntryTime
Loop Until (PassingTime >= 4)

'If (PassingTime >= 4) Then
'    MPTester.MSComm2.Output = "DevScanPCIeBus"
'    Call MsecDelay(3#)
'End If

Call MsecDelay(RemoveTime)


'flow2 (P1A write 1110 1110 & PowerON) (target boot 1AEA/6601, SUBSYS: 1179/F920)
'cardresult = DO_WritePort(card, Channel_P1A, &HEE)
cardresult = DO_WritePort(card, Channel_P1A, &HD6)
'Call GPIBWrite("OUT1 1")
'Call GPIBWrite("OUT2 1")
Call MsecDelay(BootTime)


If (AU6621COM_RWTest("VEN1AEADEV6601", 3#, False) <> 1) Then    'check VID/DID
    TempRes = 2
    MPTester.Print "Find 1AEA/6601 fail ..."
    AU6621ContiFail = AU6621ContiFail + 1
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 1AEA/6601 PASS"
End If

If (AU6621COM_RWTest("ven1179devF900", 3#, False) <> 1) Then    'check DVID/DDID
    TempRes = 2
    MPTester.Print "Find 1179/F900 fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 1179/F900 PASS"
End If

cardresult = DO_WritePort(card, Channel_P1A, &HD7 And &HBF)

MPTester.MSComm2.InBufferCount = 0
COM2RecBuf = ""
EntryTime = Timer
Do
    COM2RecBuf = COM2RecBuf & MPTester.MSComm2.Input
    If COM2RecBuf <> "" Then
    
        If InStr(COM2RecBuf, "NoDev") > 0 Then
            Exit Do
        End If
    End If
    PassingTime = Timer - EntryTime
Loop Until (PassingTime >= 4)

'If (PassingTime >= 4) Then
'    MPTester.MSComm2.Output = "DevScanPCIeBus"
'    Call MsecDelay(3#)
'End If
Call MsecDelay(RemoveTime)


'flow3 (P1A write 1101 1110 & PowerON) (target boot 1AEA/6601, SUBSYS: 1179/F940)
'cardresult = DO_WritePort(card, Channel_P1A, &HDE)
cardresult = DO_WritePort(card, Channel_P1A, &HE6)
'Call GPIBWrite("OUT1 1")
'Call GPIBWrite("OUT2 1")
Call MsecDelay(BootTime)

If (AU6621COM_RWTest("VEN1AEADEV6601", 3#, False) <> 1) Then    'check VID/DID
    TempRes = 2
    MPTester.Print "Find 1AEA/6601 fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 1AEA/6601 PASS"
End If

If (AU6621COM_RWTest("ven1179devF940", 3#, False) <> 1) Then    'check DVID/DDID
    TempRes = 2
    MPTester.Print "Find 1179/F940 fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 1179/F940 PASS"
End If

cardresult = DO_WritePort(card, Channel_P1A, &HE7 And &HBF)

MPTester.MSComm2.InBufferCount = 0
COM2RecBuf = ""
EntryTime = Timer
Do
    COM2RecBuf = COM2RecBuf & MPTester.MSComm2.Input
    If COM2RecBuf <> "" Then
    
        If InStr(COM2RecBuf, "NoDev") > 0 Then
            Exit Do
        End If
    End If
    PassingTime = Timer - EntryTime
Loop Until (PassingTime >= 4)

'If (PassingTime >= 4) Then
'    MPTester.MSComm2.Output = "DevScanPCIeBus"
'    Call MsecDelay(3#)
'End If

Call MsecDelay(RemoveTime)

'flow4 (P1A write 1100 0110 & PowerON) (target boot 1AEA/6601, SUBSYS: 0001/0001)
'cardresult = DO_WritePort(card, Channel_P1A, &HC6)
cardresult = DO_WritePort(card, Channel_P1A, &HFE)
'Call GPIBWrite("OUT1 1")
'Call GPIBWrite("OUT2 1")
Call MsecDelay(BootTime)

If (AU6621COM_RWTest("VEN1AEADEV6601", 3#, False) <> 1) Then    'check VID/DID
    TempRes = 2
    MPTester.Print "Find 1AEA/6601 fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 1AEA/6601 PASS"
End If

If (AU6621COM_RWTest("ven0001dev0001", 3#, False) <> 1) Then    'check DVID/DDID
    TempRes = 2
    MPTester.Print "Find 0001/0001 fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "Check 0001/0001 PASS"
End If

cardresult = DO_WritePort(card, Channel_P1A, &HFF And &HBF)

MPTester.MSComm2.InBufferCount = 0
COM2RecBuf = ""
EntryTime = Timer
Do
    COM2RecBuf = COM2RecBuf & MPTester.MSComm2.Input
    If COM2RecBuf <> "" Then
    
        If InStr(COM2RecBuf, "NoDev") > 0 Then
            Exit Do
        End If
    End If
    PassingTime = Timer - EntryTime
Loop Until (PassingTime >= 4)

'If (PassingTime >= 4) Then
'    MPTester.MSComm2.Output = "DevScanPCIeBus"
'    Call MsecDelay(3#)
'End If

Call MsecDelay(RemoveTime)

'R/W SD
'flow5 (P1A write 1111 1100 & PowerON) (target boot 1AEA/6621, SUBSYS: 0001/0001)
'cardresult = DO_WritePort(card, Channel_P1A, &HFC)
cardresult = DO_WritePort(card, Channel_P1A, &HC4)
Call MsecDelay(0.02)
'Call GPIBWrite("OUT1 1")
'Call GPIBWrite("OUT2 1")
Call MsecDelay(BootTime)

If (AU6621COM_RWTest("AU6621B82SDFun", 5#, True) <> 1) Then
    TempRes = 3
    MPTester.Print "SD Test Fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "SD Test PASS"
End If


'R/W MS
'cardresult = DO_WritePort(card, Channel_P1A, &HFE)
cardresult = DO_WritePort(card, Channel_P1A, &HC6)
Call MsecDelay(0.1)
'cardresult = DO_WritePort(card, Channel_P1A, &HFA)
cardresult = DO_WritePort(card, Channel_P1A, &HC2)

Call MsecDelay(0.1)
cardresult = DO_ReadPort(card, Channel_P1B, LEDVal)
If (LEDVal And &H2 = 2) Then
    LEDOnFlag = True
End If

If (AU6621COM_RWTest("AU6621B82MSFun", 5#, True) <> 1) Then
    TempRes = 5
    MPTester.Print "MS Test Fail ..."
    GoTo AU6621TestEndLabel
Else
    TempRes = 1
    MPTester.Print "MS Test PASS"
End If

If (TempRes = 1) Then
    If (Not LEDOnFlag) Then
        cardresult = DO_ReadPort(card, Channel_P1B, LEDVal)
        If (LEDVal And &H2) <> 0 Then
            TempRes = 4
            MPTester.Print "LED Fail ..."
        Else
            MPTester.Print "LED PASS"
        End If
    End If
End If

AU6621TestEndLabel:


'    Call GPIBWrite("OUT1 0")
'    Call GPIBWrite("OUT2 0")
    
    'cardresult = DO_WritePort(card, Channel_P1A, &HFF)
'    cardresult = DO_WritePort(card, Channel_P1A, &H86)
'    Call MsecDelay(0.2)
    cardresult = DO_WritePort(card, Channel_P1A, &H81)
    
    If OSRes = 1 Then
        MPTester.MSComm2.InBufferCount = 0
        COM2RecBuf = ""
        EntryTime = Timer
        Do
            COM2RecBuf = COM2RecBuf & MPTester.MSComm2.Input
            If COM2RecBuf <> "" Then
            
                If InStr(COM2RecBuf, "NoDev") > 0 Then
                    Exit Do
                End If
            End If
            PassingTime = Timer - EntryTime
        Loop Until (PassingTime >= 4)
        
        Call MsecDelay(0.2)
    End If
    
    
    
'    If (PassingTime >= 4) Then
'        MPTester.MSComm2.Output = "DevScanPCIeBus"
'        Call MsecDelay(3#)
'    End If
    
'    AU6621TestCount = AU6621TestCount + 1
'    If AU6621TestCount > 50 Then
'        AU6621TestCount = 0
'        MPTester.MSComm2.Output = "RestartPCIHOST"
'        Call MsecDelay(4#)
'    End If
    
'    If AU6621ContiFail >= 5 Then
'        Call MsecDelay(1.2)
'        MPTester.MSComm2.Output = "DevScanPCIeBus"
'        Call MsecDelay(8#)
'        AU6621ContiFail = 0
'    Else
'        Call MsecDelay(0.6)
'    End If
    
    If OSRes <> 1 Then
        MPTester.TestResultLab = "Bin2: OpenShort Fail"
        TestResult = "Bin2"
    ElseIf TempRes = 2 Then
        MPTester.TestResultLab = "Bin2: Enum Fail"
        TestResult = "Bin2"
    ElseIf TempRes = 3 Then
        MPTester.TestResultLab = "Bin3: SD Fail"
        TestResult = "Bin3"
        AU6621ContiFail = 0
    ElseIf TempRes = 5 Then
        MPTester.TestResultLab = "Bin5: MS Fail"
        TestResult = "Bin5"
        AU6621ContiFail = 0
    ElseIf TempRes = 4 Then
        MPTester.TestResultLab = "Bin4: LED Fail"
        TestResult = "Bin4"
        AU6621ContiFail = 0
    ElseIf TempRes = 1 Then
        MPTester.TestResultLab = "PASS"
        TestResult = "PASS"
        AU6621ContiFail = 0
    Else
        MPTester.TestResultLab = "Undefine Fail"
        TestResult = "Bin2"
    End If

End Sub

Public Function AU6621COM_RWTest(WriteString As String, ItemTimeOut As Single, IsDiskTest As Boolean) As Byte

Dim EntryTime As Long
Dim PassingTime As Long
Dim RecTarget As String
Dim COM2RecBuf
Dim WaitTime As Long

If Not IsDiskTest Then
    RecTarget = "Ready"
    WaitTime = 15
Else
    RecTarget = "Device"
    WaitTime = 8
End If

MPTester.MSComm2.InBufferCount = 0
COM2RecBuf = ""
EntryTime = Timer
Do
    COM2RecBuf = COM2RecBuf & MPTester.MSComm2.Input
    If COM2RecBuf <> "" Then
        If (RecTarget = "Ready") And InStr(COM2RecBuf, "Device") Then
            MPTester.MSComm2.Output = "DevScanPCIeBus"
            Call MsecDelay(4#)
        End If
    
        If InStr(COM2RecBuf, RecTarget) > 0 Then
            Exit Do
        End If
    End If
    PassingTime = Timer - EntryTime
Loop Until (PassingTime >= WaitTime)

If PassingTime >= WaitTime Then
    AU6621COM_RWTest = 3
    'Exit Function
End If

Call MsecDelay(0.1)
MPTester.MSComm2.InBufferCount = 0
COM2RecBuf = ""
MPTester.MSComm2.Output = WriteString

EntryTime = Timer
Do
    COM2RecBuf = COM2RecBuf & MPTester.MSComm2.Input
    If COM2RecBuf <> "" Then
        If InStr(COM2RecBuf, "PASS") > 0 Then
            AU6621COM_RWTest = 1
            Exit Do
        End If
        
        If InStr(COM2RecBuf, "Bin") > 0 Then
            AU6621COM_RWTest = 2
            Exit Do
        End If
    End If
    PassingTime = Timer - EntryTime
Loop Until (PassingTime >= ItemTimeOut)

'If AU6621COM_RWTest <> 1 Then
'    AU6621ContiFail = AU6621ContiFail + 1
'
'    If AU6621ContiFail >= 3 Then
''        Call GPIBWrite("OUT1 0")
''        Call GPIBWrite("OUT2 0")
'        cardresult = DO_WritePort(card, Channel_P1A, &HC1)
'        Call MsecDelay(0.5)
'        MPTester.MSComm2.Output = "DevScanPCIeBus"
'        Call MsecDelay(8#)
'        AU6621ContiFail = 0
'    End If
'    Exit Function
'End If


If PassingTime >= ItemTimeOut Then
    AU6621COM_RWTest = 3
    Exit Function
End If


MPTester.MSComm2.InBufferCount = 0
COM2RecBuf = ""
EntryTime = Timer
Do
    COM2RecBuf = COM2RecBuf & MPTester.MSComm2.Input
    If COM2RecBuf <> "" Then
        If InStr(COM2RecBuf, "Ready") > 0 Then
            Exit Do
        End If
    End If
    PassingTime = Timer - EntryTime
Loop Until (PassingTime >= 5)

If PassingTime >= 5 Then
    AU6621COM_RWTest = 3
    Exit Function
End If


End Function

