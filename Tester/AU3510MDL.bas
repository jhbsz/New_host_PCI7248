Attribute VB_Name = "AU3510MDL"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As msg) As Long

Public winHwnd As Long
Public AlcorMPMessage As Long
Public Const SW_SHOW = 5
Public Const PM_NOREMOVE = &H0
Public Const PM_REMOVE = &H1
Const SWP_NOMOVE = &H2 '不更動目前視窗位置
Const SWP_NOSIZE = &H1 '不更動目前視窗大小
Public Const HWND_TOPMOST = -1 '設定為最上層
Const HWND_NOTOPMOST = -2 '取消最上層設定
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Public Const WM_CLOSE = &H10
Public Const WM_DESTROY = &H2
Public Const WM_QUIT = &H12

Public recvStatus As Long

Public Type msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Option Explicit

Public Const WM_USER = &H400
Public Const WM_FT_TEST_START = WM_USER + &H100
Public Const WM_FT_TEST_UNKNOWN = WM_USER + &H200
Public Const WM_FT_TEST_FAIL = WM_USER + &H220
Public Const WM_FT_TEST_PASS = WM_USER + &H400


Public Sub AU3522DLF20TestSub()

Dim PassTime, OldTimer
Dim rt2
Dim mMsg As msg
Dim winHwnd As Long
Dim AlcorMessage As Long

    
    startProgress = True

    Tester.Label38.Visible = True
    Tester.ProgressBar1.Visible = True
    Tester.ProgressBar1.Min = 0

    If PCI7248InitFinish = 0 Then
          PCI7248Exist
    End If
    
    ' ===== for init ctrl 1 =====
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    MsecDelay (0.01)

    ' ******** checking *********
    ' ===== (switch relay) set con 1-3 to low =====
    CardResult = DO_WritePort(card, Channel_P1A, &HFB)
    MsecDelay (0.01)
    
    ' ===== (pw on) set con 1-1 to low =====
    CardResult = DO_WritePort(card, Channel_P1A, &HFA)
    MsecDelay (0.01)
    
    ' ===== (reset) set con 1-2 low and than hi, wait 5 sec for ACT_FLAG =====
    CardResult = DO_WritePort(card, Channel_P1A, &HF8)
    MsecDelay (0.01)
    CardResult = DO_WritePort(card, Channel_P1A, &HFA)
    MsecDelay (4.5)
    
    ' ===== (read GPIO) con 2-4, if any 0 then fail (wait test 25s) =====
    ' ###### TestResult unsure ######
    OldTimer = Timer

    If InStr(1, ChipName, "20") <> 0 Then
    
        Do
            CardResult = DO_ReadPort(card, Channel_P1B, recvStatus)
            MsecDelay (0.01)
    
            If (recvStatus And &HF0) < &H80 Then
                Tester.Print "ACT_FLAG fail"
                TestResult = "Bin2"
                GoTo skipUSB
            End If
            
            PassTime = Timer - OldTimer
            
            ' 20141031 add (recvStatus And &HF0) = &H90 to reduce time
        Loop Until (recvStatus And &HF0) = &H90 Or _
                    (recvStatus And &HF0) = &HA0 Or _
                    (recvStatus And &HF0) = &HB0 Or _
                    (recvStatus And &HF0) = &HC0 Or _
                    (recvStatus And &HF0) = &HD0 Or _
                    (recvStatus And &HF0) = &HE0 Or _
                    PassTime > 42
    Else
    
        Do
            CardResult = DO_ReadPort(card, Channel_P1B, recvStatus)
            MsecDelay (0.01)
    
            If (recvStatus And &HF0) < &H80 Then
                Tester.Print "ACT_FLAG fail"
                TestResult = "Bin2"
                GoTo skipUSB
            End If
            
            PassTime = Timer - OldTimer
            
            ' 20141031 add (recvStatus And &HF0) = &H90 to reduce time
        Loop Until (recvStatus And &HF0) = &H90 Or _
                    (recvStatus And &HF0) = &HA0 Or _
                    (recvStatus And &HF0) = &HB0 Or _
                    (recvStatus And &HF0) = &HC0 Or _
                    (recvStatus And &HF0) = &HD0 Or _
                    (recvStatus And &HF0) = &HE0 Or _
                    PassTime > 52
    End If
    
    If (recvStatus And &HF0) = &H90 Then
        Tester.Print "Check done"
        TestResult = "Pass"
    ElseIf ((recvStatus And &HF0) = &HA0 Or (recvStatus And &HF0) = &HE0) Then
        Tester.Print "SD or USB fail"
        TestResult = "Bin5"
        GoTo skipUSB
    ElseIf ((recvStatus And &HF0) = &HB0 Or (recvStatus And &HF0) = &HC0) Then
        Tester.Print "Video fail"
        TestResult = "Bin3"
        GoTo skipUSB
    ElseIf (recvStatus And &HF0) = &HD0 Then
        Tester.Print "Audio fail"
        TestResult = "Bin4"
        GoTo skipUSB
    Else
        Tester.Print "Unknown or Time out fail"
        TestResult = "Bin2"
        GoTo skipUSB
    End If

    ' ******** testing **********
    ' ===== (switch relay) set con 1-3 to low =====
    CardResult = DO_WritePort(card, Channel_P1A, &HFE)
    MsecDelay (0.01)
    
    ' ===== (reset) set con 1-2 low and than hi =====
    CardResult = DO_WritePort(card, Channel_P1A, &HFC)
    MsecDelay (0.01)
    CardResult = DO_WritePort(card, Channel_P1A, &HFE)
    MsecDelay (0.01)
    
    ' ===== (USB testing) call AP, need WM =====
    ' find window
    winHwnd = FindWindow(vbNullString, "MP Sorting Tool")
     
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(Tester.hwnd, "open", App.Path & "\AU3522\MP SortingTool.exe", "", "", SW_SHOW)
        
        Call MsecDelay(0.05)
        winHwnd = FindWindow(vbNullString, "MP Sorting Tool")
    End If
    
    WaitDevOn ("pid_3522")
    WaitDevOn ("pid_3522")
    WaitDevOn ("pid_3522")
    
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_FT_TEST_START, 0&, 0&)
    End If

    ' ===== USB Testing(within 1 sec) =====
    OldTimer = Timer
        
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        PassTime = Timer - OldTimer
        
    Loop Until AlcorMessage = WM_FT_TEST_PASS _
            Or AlcorMessage = WM_FT_TEST_FAIL _
            Or AlcorMessage = WM_FT_TEST_UNKNOWN _
            Or PassTime > 8
            
    If AlcorMessage = WM_FT_TEST_PASS Then
        TestResult = "PASS"
    ElseIf AlcorMessage = WM_FT_TEST_FAIL Then
        TestResult = "Bin5"                         ' for usb fail
    Else
        TestResult = "Bin2"
    End If

skipUSB:

    ' set con 1-2 low
    CardResult = DO_WritePort(card, Channel_P1A, &HF8)
    Tester.ProgressBar1.Value = 0
    startProgress = False
    Tester.Timer2.Enabled = True

End Sub

Public Sub AU3510ELF20TestSub()

Dim PassTime, OldTimer
Dim rt2
Dim mMsg As msg
Dim winHwnd As Long
Dim AlcorMessage As Long

    
    startProgress = True

    Tester.Label38.Visible = True
    Tester.ProgressBar1.Visible = True
    Tester.ProgressBar1.Min = 0

    If PCI7248InitFinish = 0 Then
          PCI7248Exist
    End If
    
    ' ===== for init ctrl 1 =====
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    MsecDelay (0.01)

    ' ******** checking *********
    ' ===== (switch relay) set con 1-3 to low =====
    CardResult = DO_WritePort(card, Channel_P1A, &HFB)
    MsecDelay (0.01)
    
    ' ===== (pw on) set con 1-1 to low =====
    CardResult = DO_WritePort(card, Channel_P1A, &HFA)
    MsecDelay (0.01)
    
    ' ===== (reset) set con 1-2 low and than hi, wait 5 sec for ACT_FLAG =====
    CardResult = DO_WritePort(card, Channel_P1A, &HF8)
    MsecDelay (0.01)
    CardResult = DO_WritePort(card, Channel_P1A, &HFA)
    MsecDelay (4.5)
    
    ' ===== (read GPIO) con 2-4, if any 0 then fail (wait test 25s) =====
    ' ###### TestResult unsure ######
    OldTimer = Timer

    If InStr(1, ChipName, "20") <> 0 Then
    
        Do
            CardResult = DO_ReadPort(card, Channel_P1B, recvStatus)
            MsecDelay (0.01)
    
            If (recvStatus And &HF) < &H8 Then
                Tester.Print "ACT_FLAG fail"
                TestResult = "Bin2"
                GoTo skipUSB
            End If
            
            PassTime = Timer - OldTimer
            
            ' 20141031 add (recvStatus And &HF) = &H9 to reduce time
        Loop Until (recvStatus And &HF) = &H9 Or _
                    (recvStatus And &HF) = &HA Or _
                    (recvStatus And &HF) = &HB Or _
                    (recvStatus And &HF) = &HC Or _
                    (recvStatus And &HF) = &HD Or _
                    (recvStatus And &HF) = &HE Or _
                    PassTime > 42
    Else
    
        Do
            CardResult = DO_ReadPort(card, Channel_P1B, recvStatus)
            MsecDelay (0.01)
    
            If (recvStatus And &HF) < &H8 Then
                Tester.Print "ACT_FLAG fail"
                TestResult = "Bin2"
                GoTo skipUSB
            End If
            
            PassTime = Timer - OldTimer
            
            ' 20141031 add (recvStatus And &HF) = &H9 to reduce time
        Loop Until (recvStatus And &HF) = &H9 Or _
                    (recvStatus And &HF) = &HA Or _
                    (recvStatus And &HF) = &HB Or _
                    (recvStatus And &HF) = &HC Or _
                    (recvStatus And &HF) = &HD Or _
                    (recvStatus And &HF) = &HE Or _
                    PassTime > 52
    End If
    
    If (recvStatus And &HF) = &H9 Then
        Tester.Print "Check done"
        TestResult = "Pass"
    ElseIf ((recvStatus And &HF) = &HA Or (recvStatus And &HF) = &HE) Then
        Tester.Print "SD or USB fail"
        TestResult = "Bin5"
        GoTo skipUSB
    ElseIf ((recvStatus And &HF) = &HB Or (recvStatus And &HF) = &HC) Then
        Tester.Print "Video fail"
        TestResult = "Bin3"
        GoTo skipUSB
    ElseIf (recvStatus And &HF) = &HD Then
        Tester.Print "Audio fail"
        TestResult = "Bin4"
        GoTo skipUSB
    Else
        Tester.Print "Unknown or Time out fail"
        TestResult = "Bin2"
        GoTo skipUSB
    End If

    ' ******** testing **********
    ' ===== (switch relay) set con 1-3 to low =====
    CardResult = DO_WritePort(card, Channel_P1A, &HFE)
    MsecDelay (0.01)
    
    ' ===== (reset) set con 1-2 low and than hi =====
    CardResult = DO_WritePort(card, Channel_P1A, &HFC)
    MsecDelay (0.01)
    CardResult = DO_WritePort(card, Channel_P1A, &HFE)
    MsecDelay (0.01)
    
    ' ===== (USB testing) call AP, need WM =====
    ' find window
    winHwnd = FindWindow(vbNullString, "MP Sorting Tool")
     
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(Tester.hwnd, "open", App.Path & "\AU3510\MP SortingTool.exe", "", "", SW_SHOW)
        
        Call MsecDelay(0.05)
        winHwnd = FindWindow(vbNullString, "MP Sorting Tool")
    End If
    
    WaitDevOn ("pid_3510")
    WaitDevOn ("pid_3510")
    WaitDevOn ("pid_3510")
    
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_FT_TEST_START, 0&, 0&)
    End If

    ' ===== USB Testing(within 1 sec) =====
    OldTimer = Timer
        
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        PassTime = Timer - OldTimer
        
    Loop Until AlcorMessage = WM_FT_TEST_PASS _
            Or AlcorMessage = WM_FT_TEST_FAIL _
            Or AlcorMessage = WM_FT_TEST_UNKNOWN _
            Or PassTime > 8
            
    If AlcorMessage = WM_FT_TEST_PASS Then
        TestResult = "PASS"
    ElseIf AlcorMessage = WM_FT_TEST_FAIL Then
        TestResult = "Bin5"                         ' for usb fail
    Else
        TestResult = "Bin2"
    End If

skipUSB:

    ' set con 1-2 low
    CardResult = DO_WritePort(card, Channel_P1A, &HF8)
    Tester.ProgressBar1.Value = 0
    startProgress = False
    Tester.Timer2.Enabled = True

End Sub


Public Sub AU3510ENG20TestSub()

Dim PassTime, OldTimer
Dim rt2
Dim mMsg As msg
Dim winHwnd As Long
Dim AlcorMessage As Long

    
    startProgress = True

    Tester.Label38.Visible = True
    Tester.ProgressBar1.Visible = True
    Tester.ProgressBar1.Min = 0

    If PCI7248InitFinish = 0 Then
          PCI7248Exist
    End If

    ' ******** testing **********
    ' ===== (switch relay) set con 1-3 to low =====
    CardResult = DO_WritePort(card, Channel_P1A, &HFE)
    MsecDelay (0.01)
    
    ' ===== (reset) set con 1-2 low and than hi =====
    CardResult = DO_WritePort(card, Channel_P1A, &HFC)
    MsecDelay (0.01)
    CardResult = DO_WritePort(card, Channel_P1A, &HFE)
    MsecDelay (0.01)
    
    ' ===== (USB testing) call AP, need WM =====
    ' find window
    winHwnd = FindWindow(vbNullString, "MP Sorting Tool")
     
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(Tester.hwnd, "open", App.Path & "\AU3510\MP SortingTool.exe", "", "", SW_SHOW)
        
        Call MsecDelay(0.05)
        winHwnd = FindWindow(vbNullString, "MP Sorting Tool")
    End If
    
    WaitDevOn ("pid_3510")
    MsecDelay (0.2)
    
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_FT_TEST_START, 0&, 0&)
    End If

    ' ===== USB Testing(within 1 sec) =====
    OldTimer = Timer
        
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        PassTime = Timer - OldTimer
        
    Loop Until AlcorMessage = WM_FT_TEST_PASS _
            Or AlcorMessage = WM_FT_TEST_FAIL _
            Or AlcorMessage = WM_FT_TEST_UNKNOWN _
            Or PassTime > 8
            
    If AlcorMessage = WM_FT_TEST_PASS Then
        TestResult = "PASS"
    ElseIf AlcorMessage = WM_FT_TEST_FAIL Then
        TestResult = "Bin5"                         ' for usb fail
        GoTo skipUSB
    Else
        TestResult = "Bin2"
        GoTo skipUSB
    End If


    ' ===== for init ctrl 1 =====
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    MsecDelay (0.01)

    ' ******** checking *********
    ' ===== (switch relay) set con 1-3 to low =====
    CardResult = DO_WritePort(card, Channel_P1A, &HFB)
    MsecDelay (0.01)
    
    ' ===== (pw on) set con 1-1 to low =====
    CardResult = DO_WritePort(card, Channel_P1A, &HFA)
    MsecDelay (0.01)
    
    ' ===== (reset) set con 1-2 low and than hi, wait 5 sec for ACT_FLAG =====
    CardResult = DO_WritePort(card, Channel_P1A, &HF8)
    MsecDelay (0.01)
    CardResult = DO_WritePort(card, Channel_P1A, &HFA)
    MsecDelay (4.5)
    
    ' ===== (read GPIO) con 2-4, if any 0 then fail (wait test 25s) =====
    ' ###### TestResult unsure ######
    OldTimer = Timer

    Do
        CardResult = DO_ReadPort(card, Channel_P1B, recvStatus)
        MsecDelay (0.01)

        If (recvStatus And &HF) < &H8 Then
            Tester.Print "ACT_FLAG fail"
            TestResult = "Bin2"
            GoTo skipUSB
        End If
        
        PassTime = Timer - OldTimer
        
    Loop Until (recvStatus And &HF) = &HA Or _
                (recvStatus And &HF) = &HB Or _
                (recvStatus And &HF) = &HC Or _
                (recvStatus And &HF) = &HD Or _
                (recvStatus And &HF) = &HE Or _
                PassTime > 42
    
    If (recvStatus And &HF) = &H9 Then
        Tester.Print "Check done"
        TestResult = "PASS"
    ElseIf ((recvStatus And &HF) = &HA Or (recvStatus And &HF) = &HE) Then
        Tester.Print "SD or USB fail"
        TestResult = "Bin5"
        'GoTo skipUSB
    ElseIf ((recvStatus And &HF) = &HB Or (recvStatus And &HF) = &HC) Then
        Tester.Print "Video fail"
        TestResult = "Bin3"
        'GoTo skipUSB
    ElseIf (recvStatus And &HF) = &HD Then
        Tester.Print "Audio fail"
        TestResult = "Bin4"
        'GoTo skipUSB
    Else
        Tester.Print "Unknown or Time out fail"
        TestResult = "Bin2"
        'GoTo skipUSB
    End If

skipUSB:

    ' set con 1-2 low
    CardResult = DO_WritePort(card, Channel_P1A, &HF8)
    Tester.ProgressBar1.Value = 0
    startProgress = False
    Tester.Timer2.Enabled = True

End Sub









