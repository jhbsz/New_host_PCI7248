Attribute VB_Name = "AU9420KeyBoard"
Option Explicit

Public Const WM_USER = &H400
Public Const WM_FT_TEST_READY = WM_USER + &H800
Public Const WM_FT_TEST_START = WM_USER + &H100
Public Const WM_FT_TEST_UNKNOW_FAIL = WM_USER + &H120
Public Const WM_FT_TEST_OTP_FAIL = WM_USER + &H200
Public Const WM_FT_TEST_BIN_FAIL = WM_USER + &H250
Public Const WM_FT_TEST_GPIO_FAIL = WM_USER + &H420
Public Const WM_FT_TEST_CLEAR = WM_USER + &H450
Public Const WM_FT_TEST_PASS = WM_USER + &H400

Public Const AtheistDebug = True
 
Public Sub LoadAP_Click_AU9420()

Dim TimePass
Dim rt2

    ' find window
    winHwnd = FindWindow(vbNullString, "AU9420_FT_Tool")
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\Keyboard\" & "Package_Tool.exe", "", "", SW_SHOW)
    End If
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags

End Sub

Public Sub LoadOTPAP_Click_AU9420()

Dim TimePass
Dim rt2

    ' find window
    winHwnd = FindWindow(vbNullString, "AmBentleyMP")
 
    ' run program
    If winHwnd = 0 Then
        ChDir (App.Path & "\Keyboard\" & ChipName)
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\Keyboard\" & ChipName & "\AmBentleyMP.exe", "", "", SW_SHOW)
    End If
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags

End Sub

 Public Sub StartRWTest_Click_AU9420()
 
 Dim rt2
    
    winHwnd = FindWindow(vbNullString, "AU9420_FT_Tool")
    rt2 = PostMessage(winHwnd, WM_FT_TEST_START, 0&, 0&)
 
 End Sub
 
  Public Sub StartOTP_Click_AU9420()
 
 Dim rt2
    
    winHwnd = FindWindow(vbNullString, "AmBentleyMP")
    rt2 = PostMessage(winHwnd, WM_FT_TEST_START, 0&, 0&)
 
 End Sub
 
 Public Sub Clear_Click_AU9420()
 
 Dim rt2
    
    winHwnd = FindWindow(vbNullString, "AU9420_FT_Tool")
    rt2 = PostMessage(winHwnd, WM_FT_TEST_CLEAR, 0&, 0&)
 
 End Sub

Public Sub LoadAP_Click_AU9420_ReadEEPROM()

Dim TimePass
Dim rt2

    ' find window
    winHwnd = FindWindow(vbNullString, "AU9420_FT_Tool_EEPROM")
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\Keyboard\" & ChipName & "\Package_Tool.exe", "", "", SW_SHOW)
    End If
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags

End Sub

 Public Sub StartRWTest_Click_AU9420_ReadEEPROM()
 
 Dim rt2
    
    winHwnd = FindWindow(vbNullString, "AU9420_FT_Tool_EEPROM")
    rt2 = PostMessage(winHwnd, WM_FT_TEST_START, 0&, 0&)
 
 End Sub
 
 Public Sub Clear_Click_AU9420_ReadEEPROM()
 
 Dim rt2
    
    winHwnd = FindWindow(vbNullString, "AU9420_FT_Tool_EEPROM")
    rt2 = PostMessage(winHwnd, WM_FT_TEST_CLEAR, 0&, 0&)
 
 End Sub

Public Sub AU9420TestSub()
 
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim RT_Count As Integer
Dim TmpStr As String
Dim TmpCount As Integer

    RT_Count = 0
    MPTester.TestResultLab = ""
    
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If
     
    AlcorMPMessage = 0
    NewChipFlag = 0
    
    If OldChipName <> ChipName Then
        ' reset program
        winHwnd = FindWindow(vbNullString, "AU9420_FT_Tool")
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, "AU9420_FT_Tool")
            Loop While winHwnd <> 0
        End If
        NewChipFlag = 1
    End If
              
    OldChipName = ChipName
    
    If NewChipFlag = 1 Or FindWindow(vbNullString, "AU9420_FT_Tool") = 0 Then
        MPTester.Print "wait for FT Tool Ready"
        Call LoadAP_Click_AU9420
    End If
     
    OldTimer = Timer
    AlcorMPMessage = 0
            
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        PassTime = Timer - OldTimer
    Loop Until AlcorMPMessage = WM_FT_TEST_READY Or PassTime > 5 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                  
    Call Clear_Click_AU9420
    MPTester.Print "Ready Time="; PassTime
            
    If PassTime > 5 Then
        MPTester.Print "AU9540 AP Ready Fail"
        Exit Sub
    End If
    
    cardresult = DO_WritePort(card, Channel_P1A, &H0)  'Open ENA Power 1111_1110
    TempRes = WaitDevOn("vid_058f")
    
    If TempRes Then
        MPTester.Print "VID= " & Mid(DevicePathName, InStr(1, DevicePathName, "vid"), 8) & _
                "; PID= " & Mid(DevicePathName, InStr(1, DevicePathName, "pid"), 8)
    Else
        MPTester.TestResultLab = "Bin2: Unknow"
        TestResult = "Bin2"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power f111_1111
        Call MsecDelay(1#)
        Exit Sub
    End If
    
    MPTester.Print "AU9540 begin test........"
    
RT_Label:
    
    OldTimer = Timer
    AlcorMPMessage = 0
    Call StartRWTest_Click_AU9420
    
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
                 
        PassTime = Timer - OldTimer
               
    Loop Until AlcorMPMessage = WM_FT_TEST_UNKNOW_FAIL _
            Or AlcorMPMessage = WM_FT_TEST_GPIO_FAIL _
            Or AlcorMPMessage = WM_FT_TEST_PASS _
            Or PassTime > 5 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        
    If (AlcorMPMessage <> WM_FT_TEST_PASS) And (RT_Count < 11) Then
        Call Clear_Click_AU9420
        Call MsecDelay(0.2)
        RT_Count = RT_Count + 1
        GoTo RT_Label
    End If
        
    MPTester.Print "RW work Time="; PassTime
    
    '===========================================================
    '  RW Time Out Fail
    '===========================================================
            
    If PassTime > 5 Then
        MPTester.TestResultLab = "Bin3:Time Out Fail"
        TestResult = "Bin3"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        MsecDelay (0.5)
        Exit Sub
    End If
            
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    
    Do
        Call MsecDelay(0.2)
        TmpStr = GetDeviceName_NoReply("vid_058f")
        TmpCount = TmpCount + 1
    Loop Until (TmpCount >= 40) Or (TmpStr = "")
    
    If (AlcorMPMessage <> WM_FT_TEST_PASS) And (FailCloseAP) Then
        Close_FT_AP ("AU9420_FT_Tool")
    End If
    
    Select Case AlcorMPMessage
                
        Case WM_FT_TEST_UNKNOW_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:UnKnow Fail"
            ContFail = ContFail + 1
            
        Case WM_FT_TEST_GPIO_FAIL
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:GPIO Error "
            ContFail = ContFail + 1
        
        Case WM_FT_TEST_PASS
            MPTester.TestResultLab = "PASS "
            TestResult = "PASS"
            ContFail = 0
        
        Case Else
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:Undefine Fail"
            ContFail = ContFail + 1
            
    End Select
                            
End Sub

Public Sub AU9420OTPSub()
 
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim RT_Count As Integer
Dim TmpStr As String
Dim TmpCount As Integer

    RT_Count = 0
    MPTester.TestResultLab = ""
    
    Call PowerSet2(1, "6.7", "0.5", 1, "6.7", "0.5", 1)
    
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If
     
    AlcorMPMessage = 0
    NewChipFlag = 0
    
    If OldChipName <> ChipName Then
        ' reset program
        winHwnd = FindWindow(vbNullString, "AmBentleyMP")
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, "AmBentleyMP")
            Loop While winHwnd <> 0
        End If
        NewChipFlag = 1
    End If
              
    OldChipName = ChipName
    
    If NewChipFlag = 1 Or FindWindow(vbNullString, "AmBentleyMP") = 0 Then
        MPTester.Print "wait for OTP Tool Ready"
        Call LoadOTPAP_Click_AU9420
    End If
     
    OldTimer = Timer
    AlcorMPMessage = 0
            
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        PassTime = Timer - OldTimer
    Loop Until AlcorMPMessage = WM_FT_TEST_READY Or PassTime > 5 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                  
    MPTester.Print "Ready Time="; PassTime
            
    If PassTime > 5 Then
        MPTester.Print "AU9420 OTP Tool Ready Fail"
        Exit Sub
    End If
    
    cardresult = DO_WritePort(card, Channel_P1A, &H0)  'Open ENA Power 1111_1110
    TempRes = WaitDevOn("vid_058f")
    
    If TempRes Then
        MPTester.Print "VID= " & Mid(DevicePathName, InStr(1, DevicePathName, "vid"), 8) & _
                "; PID= " & Mid(DevicePathName, InStr(1, DevicePathName, "pid"), 8)
    Else
        MPTester.TestResultLab = "Bin2: Unknow"
        TestResult = "Bin2"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power f111_1111
        Call MsecDelay(1#)
        Exit Sub
    End If
    
    MPTester.Print "AU9420 Start Programming........"
    
    OldTimer = Timer
    AlcorMPMessage = 0
    Call StartOTP_Click_AU9420
    
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
                 
        PassTime = Timer - OldTimer
               
    Loop Until AlcorMPMessage = WM_FT_TEST_UNKNOW_FAIL _
            Or AlcorMPMessage = WM_FT_TEST_BIN_FAIL _
            Or AlcorMPMessage = WM_FT_TEST_OTP_FAIL _
            Or AlcorMPMessage = WM_FT_TEST_PASS _
            Or PassTime > 10 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        
    MPTester.Print "RW work Time="; PassTime
    
    '===========================================================
    '  RW Time Out Fail
    '===========================================================
            
    If PassTime > 10 Then
        MPTester.TestResultLab = "Bin5:Time Out Fail"
        TestResult = "Bin5"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        MsecDelay (0.5)
        Exit Sub
    End If
            
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    
    Do
        Call MsecDelay(0.2)
        TmpStr = GetDeviceName_NoReply("vid_058f")
        TmpCount = TmpCount + 1
    Loop Until (TmpCount >= 40) Or (TmpStr = "")
    
    If (AlcorMPMessage <> WM_FT_TEST_PASS) And (FailCloseAP) Then
        Close_FT_AP ("AmBentleyMP")
    End If
    
    Select Case AlcorMPMessage
                
        Case WM_FT_TEST_UNKNOW_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:UnKnow Fail"
            ContFail = ContFail + 1
            
        Case WM_FT_TEST_OTP_FAIL
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:OTP Fail"
            ContFail = ContFail + 1
        
        Case WM_FT_TEST_BIN_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = "Bin4:BinFile Fail"
            ContFail = ContFail + 1
        
        Case WM_FT_TEST_PASS
            MPTester.TestResultLab = "PASS "
            TestResult = "PASS"
            ContFail = 0
        
        Case Else
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:Undefine Fail"
            ContFail = ContFail + 1
            
    End Select
                            
End Sub

Public Sub AU9420EQCTestSub()
 
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim HV_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String
Dim RT_Count As Integer
Dim TmpStr As String
Dim TmpCount As Integer
    
    MPTester.TestResultLab = ""
    HV_Flag = False
    HV_Result = ""
    LV_Result = ""
    
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If
    
Routine_Label:
    
    RT_Count = 0
    AlcorMPMessage = 0
    NewChipFlag = 0

    If Not HV_Flag Then
        Call PowerSet2(1, "5.5", "0.5", 1, "5.5", "0.5", 1)
        MPTester.Print "Begin HV Test ..."
    Else
        Call PowerSet2(1, "4.7", "0.5", 1, "4.7", "0.5", 1)
        MPTester.Print vbCrLf & "Begin LV Test ..."
    End If

    If OldChipName <> ChipName Then
        ' reset program
        winHwnd = FindWindow(vbNullString, "AU9420_FT_Tool")
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, "AU9420_FT_Tool")
            Loop While winHwnd <> 0
        End If
        NewChipFlag = 1
    End If
          
    OldChipName = ChipName
    
    If NewChipFlag = 1 Or FindWindow(vbNullString, "AU9420_FT_Tool") = 0 Then
        MPTester.Print "wait for FT Tool Ready"
        Call LoadAP_Click_AU9420
    End If
     
    OldTimer = Timer
    AlcorMPMessage = 0
        
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        PassTime = Timer - OldTimer
    Loop Until AlcorMPMessage = WM_FT_TEST_READY Or PassTime > 5 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              

    Call Clear_Click_AU9420
    MPTester.Print "Ready Time="; PassTime
    
    If PassTime > 5 Then
        MPTester.Print "AU9540 AP Ready Fail"
        Exit Sub
    End If

    cardresult = DO_WritePort(card, Channel_P1A, &H0)  'Open ENA Power 1111_1110
    TempRes = WaitDevOn("vid_058f")
    
    If TempRes Then
        MPTester.Print "VID= " & Mid(DevicePathName, InStr(1, DevicePathName, "vid"), 8) & _
                "; PID= " & Mid(DevicePathName, InStr(1, DevicePathName, "pid"), 8)
    Else
        MPTester.TestResultLab = "Bin2: Unknow"
        TestResult = "Bin2"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power f111_1111
        Call MsecDelay(1#)
        Exit Sub
    End If

    MPTester.Print "AU9540 begin test........"
    
RT_Label:
    
    OldTimer = Timer
    AlcorMPMessage = 0
    Call StartRWTest_Click_AU9420

    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        PassTime = Timer - OldTimer
    Loop Until AlcorMPMessage = WM_FT_TEST_UNKNOW_FAIL _
            Or AlcorMPMessage = WM_FT_TEST_GPIO_FAIL _
            Or AlcorMPMessage = WM_FT_TEST_PASS _
            Or PassTime > 5 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
    
    If (AlcorMPMessage = WM_FT_TEST_UNKNOW_FAIL) And (RT_Count < 11) Then
        Call Clear_Click_AU9420
        Call MsecDelay(0.2)
        RT_Count = RT_Count + 1
        'MPTester.Print RT_Count
        GoTo RT_Label
    End If

    MPTester.Print "RW work Time="; PassTime
    '===========================================================
    '  RW Time Out Fail
    '===========================================================
            
    If PassTime > 5 Then
        MPTester.TestResultLab = "Bin3:Time Out Fail"
        TestResult = "Bin3"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        MsecDelay (0.5)
        Exit Sub
    End If
            
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
       
    Do
        Call MsecDelay(0.2)
        TmpStr = GetDeviceName_NoReply("vid_058f")
        TmpCount = TmpCount + 1
    Loop Until (TmpCount >= 40) Or (TmpStr = "")
               
    If (AlcorMPMessage <> WM_FT_TEST_PASS) And (FailCloseAP) Then
        Close_FT_AP ("AU9420_FT_Tool")
    End If

    Select Case AlcorMPMessage
        Case WM_FT_TEST_UNKNOW_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:UnKnow Fail"
            ContFail = ContFail + 1
                 
            If Not HV_Flag Then
                MPTester.Print "HV: UnKnow Fail"
                HV_Result = TestResult
            Else
                MPTester.Print "LV: UnKnow Fail"
                LV_Result = TestResult
            End If
        
        Case WM_FT_TEST_GPIO_FAIL
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:GPIO Error"
            ContFail = ContFail + 1
            
            If Not HV_Flag Then
                MPTester.Print "HV: GPIO Fail"
                HV_Result = TestResult
            Else
                MPTester.Print "LV: GPIO Fail"
                LV_Result = TestResult
            End If
        
        Case WM_FT_TEST_PASS
            MPTester.TestResultLab = "PASS"
            TestResult = "PASS"
            ContFail = 0
            
            If Not HV_Flag Then
                MPTester.Print "HV: PASS"
                HV_Result = TestResult
            Else
                MPTester.Print "LV: PASS"
                LV_Result = TestResult
            End If
        
        Case Else
            TestResult = "Bin2"
            MPTester.TestResultLab = "HV:Undefine Fail"
            ContFail = ContFail + 1
            
            If Not HV_Flag Then
                MPTester.Print "HV: Undefine Fail"
                HV_Result = TestResult
            Else
                MPTester.Print "LV: Undefine Fail"
                LV_Result = TestResult
            End If
        
    End Select
        
    If Not HV_Flag Then
        HV_Flag = True
        GoTo Routine_Label
    End If
    
    If (HV_Result = "Bin2") And (LV_Result = "Bin2") Then
        TestResult = "Bin2"
    ElseIf (HV_Result <> "PASS") And (LV_Result = "PASS") Then
        TestResult = "Bin3"
    ElseIf (HV_Result = "PASS") And (LV_Result <> "PASS") Then
        TestResult = "Bin4"
    ElseIf (HV_Result <> "PASS") And (LV_Result <> "PASS") Then
        TestResult = "Bin5"
    ElseIf (HV_Result = "PASS") And (LV_Result = "PASS") Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
                            
End Sub

Public Sub AU9420ReadEEPROMTestSub()
 
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim RT_Count As Integer
Dim TmpStr As String
Dim TmpCount As Integer

    RT_Count = 0
    MPTester.TestResultLab = ""
    
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If
     
    AlcorMPMessage = 0
    NewChipFlag = 0
    
    If OldChipName <> ChipName Then
        ' reset program
        winHwnd = FindWindow(vbNullString, "AU9420_FT_Tool_EEPROM")
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, "AU9420_FT_Tool_EEPROM")
            Loop While winHwnd <> 0
        End If
        NewChipFlag = 1
    End If
              
    OldChipName = ChipName
    
    If NewChipFlag = 1 Or FindWindow(vbNullString, "AU9420_FT_Tool_EEPROM") = 0 Then
        MPTester.Print "wait for FT Tool Ready"
        Call LoadAP_Click_AU9420_ReadEEPROM
    End If
     
    OldTimer = Timer
    AlcorMPMessage = 0
            
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        PassTime = Timer - OldTimer
    Loop Until AlcorMPMessage = WM_FT_TEST_READY Or PassTime > 5 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                  
    Call Clear_Click_AU9420_ReadEEPROM
    MPTester.Print "Ready Time="; PassTime
            
    If PassTime > 5 Then
        MPTester.Print "AU9540 AP Ready Fail"
        Exit Sub
    End If
    
    cardresult = DO_WritePort(card, Channel_P1A, &H0)  'Open ENA Power 1111_1110
    TempRes = WaitDevOn("vid_058f")
    
    If TempRes Then
        MPTester.Print "VID= " & Mid(DevicePathName, InStr(1, DevicePathName, "vid"), 8) & _
                "; PID= " & Mid(DevicePathName, InStr(1, DevicePathName, "pid"), 8)
    Else
        MPTester.TestResultLab = "Bin2: Unknow"
        TestResult = "Bin2"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power f111_1111
        Call MsecDelay(1#)
        Exit Sub
    End If
    
    MPTester.Print "AU9540 begin test........"
    
RT_Label:
    
    OldTimer = Timer
    AlcorMPMessage = 0
    Call StartRWTest_Click_AU9420_ReadEEPROM
    
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
                 
        PassTime = Timer - OldTimer
               
    Loop Until AlcorMPMessage = WM_FT_TEST_UNKNOW_FAIL _
            Or AlcorMPMessage = WM_FT_TEST_GPIO_FAIL _
            Or AlcorMPMessage = WM_FT_TEST_PASS _
            Or PassTime > 5 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        
    If (AlcorMPMessage <> WM_FT_TEST_PASS) And (RT_Count < 11) Then
        Call Clear_Click_AU9420_ReadEEPROM
        Call MsecDelay(0.2)
        RT_Count = RT_Count + 1
        GoTo RT_Label
    End If
        
    MPTester.Print "RW work Time="; PassTime
    
    '===========================================================
    '  RW Time Out Fail
    '===========================================================
            
    If PassTime > 5 Then
        MPTester.TestResultLab = "Bin3:Time Out Fail"
        TestResult = "Bin3"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        MsecDelay (0.5)
        Exit Sub
    End If
            
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    
    Do
        Call MsecDelay(0.2)
        TmpStr = GetDeviceName_NoReply("vid_058f")
        TmpCount = TmpCount + 1
    Loop Until (TmpCount >= 40) Or (TmpStr = "")
    
    If (AlcorMPMessage <> WM_FT_TEST_PASS) And (FailCloseAP) Then
        Close_FT_AP ("AU9420_FT_Tool_EEPROM")
    End If
    
    Select Case AlcorMPMessage
                
        Case WM_FT_TEST_UNKNOW_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:UnKnow Fail"
            ContFail = ContFail + 1
            
        Case WM_FT_TEST_GPIO_FAIL
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:GPIO Error "
            ContFail = ContFail + 1
        
        Case WM_FT_TEST_PASS
            MPTester.TestResultLab = "PASS "
            TestResult = "PASS"
            ContFail = 0
        
        Case Else
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:Undefine Fail"
            ContFail = ContFail + 1
            
    End Select
                            
End Sub
